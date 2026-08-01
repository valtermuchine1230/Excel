#!/usr/bin/env python3
"""
vpn_connect.py

Objetivo:
- Selecionar aleatoriamente uma das configs WireGuard (prefere as que têm IPv6).
- Adaptar endpoint para IPv6 se possível.
- Escrever /etc/wireguard/wg0.conf, subir interface (wg-quick up).
- Confirmar interface, obter IPv4/IPv6 públicos e geolocalização.
- Imprimir resumo legível e permanecer em loop de monitorização até cancelarem.

AVISO: Este ficheiro contém chaves privadas/configs hardcoded por pedido.
Nunca faça isto em repositórios públicos sem proteção.
"""

import os
import re
import sys
import time
import random
import shutil
import subprocess
import logging
import traceback
from datetime import datetime

import requests

# -------------------------
# Configs fornecidas (já colocadas tal como pediste)
# -------------------------
CONFIGS = {
    "MX-FREE-3": """[Interface]
# Bouncing = 2
# NAT-PMP (Port Forwarding) = off
# VPN Accelerator = on
PrivateKey = aD61ubNv+aVJve4u2r4MSbWJMJsrvUMmQkGRIdBqj1o=
Address = 10.2.0.2/32
DNS = 10.2.0.1

[Peer]
# MX-FREE#3
PublicKey = rNyiLhJsBGoHd0A6Yrzt6c5DHuD3urE5+ZN3DZITfD8=
AllowedIPs = 0.0.0.0/0, ::/0
Endpoint = 149.102.224.33:51820

PersistentKeepalive = 25
""",
    "US-FREE-33-A": """[Interface]
# Bouncing = 1
# NAT-PMP (Port Forwarding) = off
# VPN Accelerator = on
PrivateKey = mClSo+prm7i2Geox/4fk9OSJTPp7J4HAexG4axYf8Fo=
Address = 10.2.0.2/32
DNS = 10.2.0.1

[Peer]
# US-FREE#33
PublicKey = SOXFyakZ9HI9TeiMRyMoy3PXYEzJJ/IDJcMvxZ3uWSE=
AllowedIPs = 0.0.0.0/0, ::/0
Endpoint = 149.102.254.90:51820

PersistentKeepalive = 25
""",
    "JP-FREE-16": """[Interface]
# Bouncing = 0
# NAT-PMP (Port Forwarding) = off
# VPN Accelerator = on
PrivateKey = SDF5r+E6IwHBaalMuYKFkj4Vr1mQj+PKzp6vI/X8LWk=
Address = 10.2.0.2/32, 2a07:b944::2:2/128
DNS = 10.2.0.1, 2a07:b944::2:1

[Peer]
# JP-FREE#16
PublicKey = BbghXRtbSYBJ2Q/eMu4JV7u8LKiKDfgybk7IJAO7iAU=
AllowedIPs = 0.0.0.0/0, ::/0
Endpoint = 45.87.213.226:51820

# Uncomment the following line (delete the # symbol) to connect to Proton VPN using IPv6.
# Endpoint = [2001:ac8:40:26::10]:51820
PersistentKeepalive = 25
""",
    "US-FREE-33-B": """[Interface]
# Bouncing = 0
# NAT-PMP (Port Forwarding) = off
# VPN Accelerator = on
PrivateKey = QEqvyYkrkN17W+ixWBII8etwKY/859raWSxcHGVWvkg=
Address = 10.2.0.2/32
DNS = 10.2.0.1

[Peer]
# US-FREE#33
PublicKey = SOXFyakZ9HI9TeiMRyMoy3PXYEzJJ/IDJcMvxZ3uWSE=
AllowedIPs = 0.0.0.0/0, ::/0
Endpoint = 149.102.254.90:51820

PersistentKeepalive = 25
"""
}

# Paths and constants
WG_CONF_PATH = "/etc/wireguard/wg0.conf"
WG_IFACE = "wg0"
WAIT_IFACE_TIMEOUT = 20
MONITOR_INTERVAL = 60

# Logging: user-facing format (info) with DEBUG available
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[logging.StreamHandler(sys.stdout)]
)
logger = logging.getLogger("vpn_connect")
# Keep debug logs available via logger.debug()

# Helper: run subprocess with capture
def run(cmd, timeout=30, check=False):
    logger.debug("RUN: %s", " ".join(cmd))
    try:
        cp = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, timeout=timeout)
        logger.debug("Exit %s, output first lines:\n%s", cp.returncode, (cp.stdout or "").splitlines()[:20])
        if check and cp.returncode != 0:
            raise RuntimeError(f"Command {cmd} failed (exit {cp.returncode}):\n{cp.stdout}")
        return cp.returncode, (cp.stdout or "")
    except subprocess.TimeoutExpired:
        raise RuntimeError(f"Command {cmd} timed out")

# Minimal pretty table printer for summary
def print_summary_row(k, v):
    logger.info(f"{k:15}: {v}")

# Detect if config contains IPv6 in Address
def config_has_ipv6(conf_text):
    m = re.search(r"(?mi)^Address\s*=\s*(.+)$", conf_text)
    if not m:
        return False
    addr = m.group(1)
    return "::" in addr or "," in addr and any("::" in p for p in [p.strip() for p in addr.split(",")])

# Try to extract an IPv6 endpoint commented in the config like "# Endpoint = [2001:...]:51820"
def find_commented_ipv6_endpoint(conf_text):
    m = re.search(r"(?m)^#\s*Endpoint\s*=\s*\[([0-9a-fA-F:]+)\]:(\d+)", conf_text)
    if m:
        return f"[{m.group(1)}]:{m.group(2)}"
    return None

# Replace Endpoint line in config with new endpoint string
def replace_endpoint(conf_text, new_endpoint):
    out = []
    replaced = False
    for line in conf_text.splitlines():
        if re.match(r"(?i)^\s*Endpoint\s*=", line) or re.match(r"(?i)^\s*#\s*Endpoint\s*=", line):
            out.append(f"Endpoint = {new_endpoint}")
            replaced = True
        else:
            out.append(line)
    if not replaced:
        # append
        out.append(f"Endpoint = {new_endpoint}")
    return "\n".join(out)

# Write config to WG_CONF_PATH as root
def write_config_as_root(conf_text):
    try:
        # Ensure directory exists
        run(["sudo", "mkdir", "-p", os.path.dirname(WG_CONF_PATH)])
        # Write via sudo tee
        rc, out = run(["sudo", "tee", WG_CONF_PATH], timeout=20)
        # Actually provide input via Popen to prevent run() capturing empty; simpler:
    except Exception:
        # fallback manual Popen with input
        p = subprocess.Popen(["sudo", "tee", WG_CONF_PATH], stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
        stdout, _ = p.communicate(conf_text)
        if p.returncode != 0:
            raise RuntimeError(f"Failed to write {WG_CONF_PATH}, rc {p.returncode}\n{stdout}")
    else:
        # our previous run() didn't provide input; do it reliably via Popen
        p = subprocess.Popen(["sudo", "tee", WG_CONF_PATH], stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
        stdout, _ = p.communicate(conf_text)
        if p.returncode != 0:
            raise RuntimeError(f"Failed to write {WG_CONF_PATH}, rc {p.returncode}\n{stdout}")
    # chmod 600
    rc, out = run(["sudo", "chmod", "600", WG_CONF_PATH])
    logger.debug("Wrote config and set permissions")

# Collect compact diagnostics (only called on failure)
def collect_diagnostics_compact():
    parts = []
    try:
        _, out = run(["sudo", "wg", "show", WG_IFACE], timeout=5)
        parts.append(("wg show", out.strip()))
    except Exception as e:
        parts.append(("wg show", f"error: {e}"))
    try:
        _, out = run(["sudo", "ip", "addr", "show", WG_IFACE], timeout=5)
        parts.append(("ip addr show", out.strip()))
    except Exception as e:
        parts.append(("ip addr show", f"error: {e}"))
    try:
        _, out = run(["sudo", "cat", WG_CONF_PATH], timeout=5)
        parts.append(("wg0.conf", out.strip().splitlines()[:30]))
    except Exception as e:
        parts.append(("wg0.conf", f"error: {e}"))
    diag_lines = []
    for k, v in parts:
        diag_lines.append(f"--- {k} ---")
        if isinstance(v, list):
            diag_lines += v
        else:
            diag_lines.append(str(v))
    return "\n".join(diag_lines)

# Ping endpoint helper for v4/v6; uses ping/ping6 where available
def ping_endpoint(endpoint_host, family=4, timeout=3):
    # endpoint_host might include port -> strip
    host = endpoint_host.split("]")[-1] if endpoint_host.startswith("[") else endpoint_host.split(":")[0]
    # If host contains [] notation, strip brackets
    host = host.strip("[]")
    if family == 6:
        cmd = ["ping", "-6", "-c", "1", "-w", str(timeout), host]
    else:
        cmd = ["ping", "-4", "-c", "1", "-w", str(timeout), host]
    try:
        rc, out = run(cmd, timeout=timeout + 2)
        return rc == 0
    except Exception:
        return False

# Bring down interface (ignore errors)
def safe_wg_down():
    try:
        run(["sudo", "wg-quick", "down", WG_IFACE], timeout=10)
    except Exception:
        logger.debug("wg-quick down (ignored) failed or was not present")

# Bring up interface and handle errors
def bring_up_interface():
    logger.info("Bringing up interface with 'sudo wg-quick up %s' ...", WG_IFACE)
    try:
        rc, out = run(["sudo", "wg-quick", "up", WG_IFACE], timeout=30)
        if rc != 0:
            raise RuntimeError(out)
    except Exception as e:
        # Collect diagnostics compact
        diag = collect_diagnostics_compact()
        raise RuntimeError(f"wg-quick up failed: {e}\n\nDIAGNOSTICS:\n{diag}")

# Wait for interface to exist
def wait_for_interface(timeout=WAIT_IFACE_TIMEOUT):
    logger.info("Aguardando interface %s ficar activa (timeout %ss)...", WG_IFACE, timeout)
    start = time.time()
    while time.time() - start < timeout:
        try:
            rc, out = run(["sudo", "wg", "show", WG_IFACE], timeout=3)
            if rc == 0 and out.strip():
                return True
            rc2, out2 = run(["sudo", "ip", "addr", "show", WG_IFACE], timeout=3)
            if rc2 == 0 and out2.strip():
                return True
        except Exception:
            logger.debug("Interface ainda não presente")
        time.sleep(1)
    return False

# Fetch public IPs and geo
def fetch_ips_and_geo():
    ipv4 = None
    ipv6 = None
    geo = {}
    try:
        r = requests.get("https://api.ipify.org?format=json", timeout=8)
        r.raise_for_status()
        data = r.json()
        ipv4 = data.get("ip")
    except Exception:
        ipv4 = None
    try:
        r = requests.get("https://api64.ipify.org?format=json", timeout=8)
        r.raise_for_status()
        data = r.json()
        ipv6 = data.get("ip")
    except Exception:
        ipv6 = None
    try:
        r = requests.get("https://ipapi.co/json/", timeout=8)
        r.raise_for_status()
        d = r.json()
        geo = {
            "city": d.get("city"),
            "region": d.get("region"),
            "country": d.get("country_name") or d.get("country"),
            "isp": d.get("org") or d.get("isp") or d.get("asn")
        }
    except Exception:
        geo = {}
    return ipv4, ipv6, geo

# Pretty summary printing
def print_connection_summary(chosen, supports_ipv6, ipv4, ipv6, geo):
    logger.info("=== VPN Connection Summary ===")
    print_summary_row("Config", chosen)
    print_summary_row("IPv6 support", "YES" if supports_ipv6 else "NO")
    print_summary_row("Public IPv4", ipv4 or "N/A")
    print_summary_row("Public IPv6", ipv6 or "N/A")
    print_summary_row("City", geo.get("city") or "N/A")
    print_summary_row("Region", geo.get("region") or "N/A")
    print_summary_row("Country", geo.get("country") or "N/A")
    print_summary_row("ISP/Org", geo.get("isp") or "N/A")
    logger.info("==============================")

def enable_sysctl_forwarding():
    # enable forwarding for ipv4 and ipv6
    try:
        run(["sudo", "sysctl", "-w", "net.ipv4.ip_forward=1"], timeout=5)
        run(["sudo", "sysctl", "-w", "net.ipv6.conf.all.forwarding=1"], timeout=5)
    except Exception:
        logger.debug("Could not set sysctl forwarding (continuing)")

def main():
    try:
        logger.info("Iniciando vpn_connect.py — vou tentar estabelecer ligação WireGuard e mostrar IPs legíveis.")
        # Show concise environment info
        logger.info("Verificando ambiente: sudo non-interactive e /dev/net/tun ...")
        try:
            rc, out = run(["sudo", "-n", "true"], timeout=5)
            if rc == 0:
                logger.info("sudo non-interactive: OK")
            else:
                logger.warning("sudo may prompt for password or not allowed non-interactively")
        except Exception:
            logger.warning("sudo check failed (may prompt)")

        # Ensure modprobe wireguard if possible
        try:
            run(["sudo", "modprobe", "wireguard"], timeout=5)
            logger.debug("modprobe wireguard OK")
        except Exception:
            logger.debug("modprobe wireguard maybe not required or failed")

        # Ensure forwarding enabled
        enable_sysctl_forwarding()

        # pick config (prefers IPv6-capable)
        keys = list(CONFIGS.keys())
        ipv6_keys = [k for k, v in CONFIGS.items() if config_has_ipv6(v)]
        if ipv6_keys:
            chosen = random.choice(ipv6_keys)
            logger.info("Preferindo configs com IPv6: escolher entre %s", ipv6_keys)
        else:
            chosen = random.choice(keys)
            logger.info("Nenhuma config com IPv6: escolhendo aleatoriamente entre todas")

        conf_text = CONFIGS[chosen]
        supports_ipv6 = config_has_ipv6(conf_text)
        logger.info("Config escolhida: %s (IPv6=%s)", chosen, "SIM" if supports_ipv6 else "NÃO")

        # If IPv6 capable, look for commented IPv6 endpoint and try to use it if reachable
        if supports_ipv6:
            ipv6_endpoint = find_commented_ipv6_endpoint(conf_text)
            if ipv6_endpoint:
                logger.info("Encontrado endpoint IPv6 comentado no ficheiro; testando reachability (%s)...", ipv6_endpoint)
                if ping_endpoint(ipv6_endpoint, family=6):
                    logger.info("Endpoint IPv6 alcançável — atualizando config para usar endpoint IPv6.")
                    conf_text = replace_endpoint(conf_text, ipv6_endpoint)
                else:
                    logger.info("Endpoint IPv6 não alcançável; manter endpoint original (IPv4).")

        # Show the endpoint that will be used (extract from conf_text)
        m = re.search(r"(?mi)^\s*Endpoint\s*=\s*(.+)$", conf_text)
        endpoint_used = m.group(1).strip() if m else "(none)"
        logger.info("Endpoint a usar: %s", endpoint_used)

        # Write config to /etc/wireguard/wg0.conf
        write_config_as_root(conf_text)
        logger.info("Config escrita em %s (root)", WG_CONF_PATH)

        # Clean down (ignore errors) and attempt up
        safe_wg_down()
        try:
            bring_up_interface()
        except Exception as e:
            # Friendly short error, then show diagnostics block
            logger.error("Falha ao levantar a interface WireGuard: %s", str(e).splitlines()[0])
            logger.info("Recolhendo diagnósticos (apresentarei apenas um resumo).")
            diag = str(e)
            # Print compact diagnostics in readable block
            logger.info("--- DIAGNÓSTICOS (resumo) ---\n%s\n--- FIM DIAGNÓSTICOS ---", diag)
            sys.exit(1)

        # Wait for interface up
        if not wait_for_interface(WAIT_IFACE_TIMEOUT):
            logger.error("A interface %s não ficou activa dentro do timeout.", WG_IFACE)
            logger.info("A recolher diagnóstico breve...")
            diag = collect_diagnostics_compact()
            logger.info("--- DIAGNÓSTICOS BREVE ---\n%s\n--- FIM ---", diag)
            sys.exit(1)

        # On success: fetch IPs and print clear summary
        ipv4, ipv6, geo = fetch_ips_and_geo()
        print_connection_summary(chosen, supports_ipv6, ipv4, ipv6, geo)

        # Monitoring loop
        logger.info("Vou agora manter a ligação e imprimir um 'heartbeat' a cada %ds. Para parar, cancela o workflow.", MONITOR_INTERVAL)
        while True:
            time.sleep(MONITOR_INTERVAL)
            try:
                up = wait_for_interface(2)
                ipv4, ipv6, _ = fetch_ips_and_geo()
                ts = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC")
                logger.info("[%s] Heartbeat | wg0 active=%s | IPv4=%s | IPv6=%s", ts, "YES" if up else "NO", ipv4 or "N/A", ipv6 or "N/A")
            except Exception:
                logger.exception("Erro no heartbeat; continuando")

    except Exception:
        logger.exception("Excepção não tratada — ver traceback abaixo")
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()
