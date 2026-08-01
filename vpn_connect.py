#!/usr/bin/env python3
"""
AVISO: Este ficheiro contém configurações WireGuard hardcoded (inclui chaves privadas) POR
PROPÓSITO do pedido. Nunca faças isto em repositórios reais ou em produção — usa sempre
secrets/variáveis de ambiente ou um cofre seguro.
"""

import os
import sys
import time
import random
import subprocess
import logging
import traceback
import re
import shutil

import requests

# -------------------------
# CONFIGURAÇÕES WireGuard hardcoded (conforme pedido)
# -------------------------
CONFIGS = {
    "JP-FREE-16": """[Interface]
PrivateKey = SDF5r+E6IwHBaalMuYKFkj4Vr1mQj+PKzp6vI/X8LWk=
Address = 10.2.0.2/32, 2a07:b944::2:2/128
DNS = 10.2.0.1, 2a07:b944::2:1

[Peer]
PublicKey = BbghXRtbSYBJ2Q/eMu4JV7u8LKiKDfgybk7IJAO7iAU=
AllowedIPs = 0.0.0.0/0, ::/0
Endpoint = 45.87.213.226:51820
PersistentKeepalive = 25
""",
    "MX-FREE-3": """[Interface]
PrivateKey = aD61ubNv+aVJve4u2r4MSbWJMJsrvUMmQkGRIdBqj1o=
Address = 10.2.0.2/32
DNS = 10.2.0.1

[Peer]
PublicKey = rNyiLhJsBGoHd0A6Yrzt6c5DHuD3urE5+ZN3DZITfD8=
AllowedIPs = 0.0.0.0/0, ::/0
Endpoint = 149.102.224.33:51820
PersistentKeepalive = 25
""",
    "US-FREE-33-A": """[Interface]
PrivateKey = QEqvyYkrkN17W+ixWBII8etwKY/859raWSxcHGVWvkg=
Address = 10.2.0.2/32
DNS = 10.2.0.1

[Peer]
PublicKey = SOXFyakZ9HI9TeiMRyMoy3PXYEzJJ/IDJcMvxZ3uWSE=
AllowedIPs = 0.0.0.0/0, ::/0
Endpoint = 149.102.254.90:51820
PersistentKeepalive = 25
""",
    "US-FREE-33-B": """[Interface]
PrivateKey = mClSo+prm7i2Geox/4fk9OSJTPp7J4HAexG4axYf8Fo=
Address = 10.2.0.2/32
DNS = 10.2.0.1

[Peer]
PublicKey = SOXFyakZ9HI9TeiMRyMoy3PXYEzJJ/IDJcMvxZ3uWSE=
AllowedIPs = 0.0.0.0/0, ::/0
Endpoint = 149.102.254.90:51820
PersistentKeepalive = 25
"""
}
# -------------------------

# Logging setup (DEBUG with timestamp)
logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%dT%H:%M:%S%z",
    handlers=[logging.StreamHandler(sys.stdout)]
)

WG_CONF_PATH = "/etc/wireguard/wg0.conf"
WG_IFACE = "wg0"
WAIT_WG_TIMEOUT = 20  # seconds
MONITOR_INTERVAL = 60  # seconds

class VPNError(Exception):
    pass

def debug_environment():
    """Print diagnostics about the environment (bin locations, /dev/net/tun, sudo)."""
    try:
        logging.debug("Effective UID: %s, USER env: %s, LOGNAME: %s", os.geteuid(), os.environ.get("USER"), os.environ.get("LOGNAME"))
        # Which binaries
        for b in ("wg-quick", "wg", "ip", "sudo"):
            path = shutil.which(b)
            logging.debug("which %s -> %s", b, path)
            if path:
                try:
                    out = subprocess.run([path, "--version"], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, timeout=5)
                    logging.debug("%s --version output: %s", b, out.stdout.strip())
                except Exception:
                    logging.debug("Could not run %s --version", b)
        # /dev/net/tun
        try:
            t = subprocess.run(["ls", "-l", "/dev/net/tun"], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, timeout=5)
            logging.debug("/dev/net/tun -> %s", t.stdout.strip())
        except Exception:
            logging.exception("ls /dev/net/tun failed")
        # sudo non-interactive check
        try:
            t = subprocess.run(["sudo", "-n", "true"], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=5)
            if t.returncode == 0:
                logging.debug("sudo non-interactive OK (sudo -n true returned 0)")
            else:
                logging.debug("sudo non-interactive check returned %s; stderr=%s", t.returncode, t.stderr.strip())
        except Exception:
            logging.exception("sudo -n true check failed (may prompt for password)")
    except Exception:
        logging.exception("Failed in debug_environment")

def has_ipv6_in_config(conf_text):
    """
    Detecta se o config contém IPv6:
    - procura '::' na linha Address, ou
    - verifica se a linha Address contém ',' indicando múltiplos endereços (ex: IPv4,IPv6)
    """
    m = re.search(r"(?mi)^Address\s*=\s*(.+)$", conf_text)
    if not m:
        return False
    addr_field = m.group(1)
    if "::" in addr_field:
        return True
    if "," in addr_field:
        parts = [p.strip() for p in addr_field.split(",")]
        for p in parts:
            if "::" in p:
                return True
    return False

def write_wg_config_as_root(config_text):
    """
    Escreve o config para /etc/wireguard/wg0.conf usando sudo tee e define permissões 600.
    Retorna None ou raises VPNError.
    """
    try:
        logging.debug("Writing WG config to %s using sudo tee", WG_CONF_PATH)
        # Use sudo tee to write file as root; provide input in one shot
        p = subprocess.run(["sudo", "tee", WG_CONF_PATH], input=config_text, text=True, capture_output=True, timeout=15)
        logging.debug("sudo tee returncode=%s stdout_len=%d stderr_len=%d", p.returncode, len(p.stdout or ""), len(p.stderr or ""))
        if p.returncode != 0:
            logging.debug("sudo tee stdout:\n%s", p.stdout)
            logging.debug("sudo tee stderr:\n%s", p.stderr)
            raise VPNError(f"Failed to write {WG_CONF_PATH} via sudo tee (returncode {p.returncode}). stderr:\n{p.stderr}")
        # Ensure correct permissions
        p2 = subprocess.run(["sudo", "chmod", "600", WG_CONF_PATH], capture_output=True, text=True, timeout=5)
        if p2.returncode != 0:
            logging.debug("chmod stderr: %s", p2.stderr)
            raise VPNError(f"Failed to chmod {WG_CONF_PATH} (returncode {p2.returncode}). stderr:\n{p2.stderr}")
        # Verify file content length
        p3 = subprocess.run(["sudo", "stat", "-c", "%s %U:%G", WG_CONF_PATH], capture_output=True, text=True, timeout=5)
        logging.debug("Stat of %s -> %s", WG_CONF_PATH, p3.stdout.strip())
    except subprocess.TimeoutExpired:
        logging.exception("Timed out while writing WG config")
        raise
    except Exception:
        logging.exception("Error while writing WG config")
        raise

def run_wg_quick_up():
    """
    Executa 'sudo wg-quick up wg0' e loga output; detecta erros comuns e lança VPNError com mensagem útil.
    """
    try:
        wg_quick = shutil.which("wg-quick")
        if not wg_quick:
            raise VPNError("'wg-quick' not found in PATH. Ensure wireguard-tools is installed.")
        logging.info("Running 'sudo wg-quick up %s' (this requires CAP_NET_ADMIN and /dev/net/tun).", WG_IFACE)
        proc = subprocess.run(["sudo", "wg-quick", "up", WG_IFACE], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, timeout=30)
        logging.debug("'wg-quick up' exit=%s", proc.returncode)
        logging.debug("'wg-quick up' output:\n%s", proc.stdout.strip())
        if proc.returncode != 0:
            out = proc.stdout.lower()
            if "operation not permitted" in out or "rt netlink" in out or "rt netlink: operation not permitted" in out:
                raise VPNError("wg-quick failed: Operation not permitted / RTNETLINK. This indicates missing CAP_NET_ADMIN or inability to modify network stack in this environment.\nFull output:\n" + proc.stdout)
            else:
                raise VPNError("wg-quick up failed. Full output:\n" + proc.stdout)
    except subprocess.TimeoutExpired:
        logging.exception("wg-quick up timed out")
        raise VPNError("wg-quick up timed out")
    except Exception:
        logging.exception("Error running wg-quick up")
        raise

def wg_interface_is_up():
    """
    Verifica se a interface wg0 está ativa usando 'sudo wg show wg0' e 'sudo ip addr show wg0'.
    """
    try:
        # wg show
        proc = subprocess.run(["sudo", "wg", "show", WG_IFACE], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=5)
        logging.debug("'sudo wg show %s' rc=%s stdout_len=%d stderr_len=%d", WG_IFACE, proc.returncode, len(proc.stdout or ""), len(proc.stderr or ""))
        if proc.returncode == 0 and (proc.stdout.strip() != "" or "interface: " in proc.stdout.lower()):
            return True
        # fallback: ip addr show wg0
        ipproc = subprocess.run(["sudo", "ip", "addr", "show", WG_IFACE], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=5)
        logging.debug("'sudo ip addr show %s' rc=%s stdout_len=%d", WG_IFACE, ipproc.returncode, len(ipproc.stdout or ""))
        if ipproc.returncode == 0 and ipproc.stdout.strip() != "":
            return True
        return False
    except Exception:
        logging.exception("Error checking wg interface status")
        return False

def get_public_ip(url):
    try:
        logging.debug("Querying %s", url)
        r = requests.get(url, timeout=10)
        r.raise_for_status()
        data = r.json()
        if isinstance(data, dict) and "ip" in data:
            return data["ip"]
        return r.text.strip()
    except Exception:
        logging.exception("Failed to fetch public IP from %s", url)
        return None

def get_geolocation():
    try:
        r = requests.get("https://ipapi.co/json/", timeout=10)
        r.raise_for_status()
        d = r.json()
        return {
            "city": d.get("city"),
            "region": d.get("region"),
            "country": d.get("country_name") or d.get("country"),
            "isp": d.get("org") or d.get("isp") or d.get("asn")
        }
    except Exception:
        logging.exception("Failed to fetch geolocation")
        return {}

def main():
    try:
        logging.info("Iniciando vpn_connect (modo WireGuard) — logging DEBUG ativo.")
        # Diagnostics
        debug_environment()

        # Choose config preferring IPv6-capable ones
        keys = list(CONFIGS.keys())
        ipv6_candidates = [k for k, v in CONFIGS.items() if has_ipv6_in_config(v)]
        if ipv6_candidates:
            chosen = random.choice(ipv6_candidates)
            logging.info("Preferência por configs com IPv6 — escolhida entre: %s", ipv6_candidates)
        else:
            chosen = random.choice(keys)
            logging.warning("Nenhuma config com IPv6 encontrada; escolhida aleatoriamente entre todas as configs (ligação será possivelmente só IPv4).")
        config_text = CONFIGS[chosen]
        has_ipv6 = has_ipv6_in_config(config_text)
        logging.info("Servidor escolhido: %s (has_ipv6=%s)", chosen, has_ipv6)

        # Write config to /etc/wireguard/wg0.conf as root
        write_wg_config_as_root(config_text)

        # Try to bring up interface
        run_wg_quick_up()

        # Wait for the interface to be active
        logging.info("Aguardando interface %s ficar activa (timeout %ss)...", WG_IFACE, WAIT_WG_TIMEOUT)
        start = time.time()
        while time.time() - start < WAIT_WG_TIMEOUT:
            if wg_interface_is_up():
                logging.info("Interface %s está activa.", WG_IFACE)
                break
            logging.debug("Interface %s ainda não activa — aguardando 1s...", WG_IFACE)
            time.sleep(1)
        else:
            # Collect diagnostic logs to help debugging
            try:
                out1 = subprocess.run(["sudo", "wg", "show", WG_IFACE], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, timeout=5)
                out2 = subprocess.run(["sudo", "ip", "addr", "show", WG_IFACE], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, timeout=5)
                logging.debug("Diagnostic wg show output:\n%s", out1.stdout)
                logging.debug("Diagnostic ip addr output:\n%s", out2.stdout)
            except Exception:
                logging.exception("Error collecting diagnostics after failed wait")
            raise VPNError(f"Interface {WG_IFACE} did not become active within {WAIT_WG_TIMEOUT} seconds.")

        # Fetch IPs and geolocation
        ipv4 = get_public_ip("https://api.ipify.org?format=json")
        ipv6 = get_public_ip("https://api64.ipify.org?format=json")
        geo = get_geolocation()

        # Print formatted summary
        logging.info("===== Resumo da Ligação WireGuard =====")
        logging.info("Servidor escolhido: %s", chosen)
        logging.info("Suporta IPv6: %s", "SIM" if has_ipv6 else "NÃO")
        logging.info("IPv4 público: %s", ipv4 or "N/A")
        logging.info("IPv6 público: %s", ipv6 or "N/A")
        logging.info("Cidade: %s", geo.get("city") or "N/A")
        logging.info("Região: %s", geo.get("region") or "N/A")
        logging.info("País: %s", geo.get("country") or "N/A")
        logging.info("ISP/Org: %s", geo.get("isp") or "N/A")
        logging.info("=======================================")

        # Monitoring loop: every 60s check wg0 and log IPs
        logging.info("Entrando em loop infinito de monitorização (intervalo %ds). Cancelar manualmente ou aguardar timeout do workflow.", MONITOR_INTERVAL)
        while True:
            time.sleep(MONITOR_INTERVAL)
            try:
                alive = wg_interface_is_up()
                if not alive:
                    logging.error("Interface %s deixou de estar activa!", WG_IFACE)
                else:
                    logging.debug("Interface %s permanece activa.", WG_IFACE)
                ipv4 = get_public_ip("https://api.ipify.org?format=json")
                ipv6 = get_public_ip("https://api64.ipify.org?format=json")
                logging.info("Heartbeat: servidor=%s | IPv4=%s | IPv6=%s", chosen, ipv4 or "N/A", ipv6 or "N/A")
            except Exception:
                logging.exception("Erro na iteração do loop de monitorização; continuará a tentar.")
    except Exception as e:
        logging.exception("Excepção não tratada no main(): %s", e)
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()
