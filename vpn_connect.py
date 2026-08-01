#!/usr/bin/env python3
"""
vpn_connect.py

AVISO: contém configurações WireGuard hardcoded (inclui chaves privadas) POR PROPÓSITO do pedido.
Nunca faças isto em repositórios públicos ou produção — usa sempre secrets/variáveis seguras.

Funcionalidade:
- Escolhe aleatoriamente uma config (prioriza configs com IPv6 quando possível).
- Escreve /etc/wireguard/wg0.conf (sudo), chmod 600.
- Executa 'sudo wg-quick down wg0' (silencioso) e 'sudo wg-quick up wg0'.
- Aguarda interface wg0 subir (wg show / ip addr), timeout configurável.
- Obtém IPv4 e IPv6 públicos e geolocalização, imprime resumo.
- Entra num loop infinito de monitorização (60s) que reimprime IPs e verifica interface.
- Logging DEBUG, captura traceback e recolhe diagnósticos em caso de falha.
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
# CONFIGURAÇÕES WireGuard hardcoded (fornecidas pelo utilizador)
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
# -------------------------

# Constants
WG_CONF_PATH = "/etc/wireguard/wg0.conf"
WG_IFACE = "wg0"
WAIT_WG_TIMEOUT = 20  # seconds to wait for wg0 to be considered up
MONITOR_INTERVAL = 60  # seconds between heartbeat checks in the infinite loop

# Logging setup (DEBUG with timestamp)
logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%dT%H:%M:%S%z",
    handlers=[logging.StreamHandler(sys.stdout)]
)


class VPNError(Exception):
    pass


def run_subprocess(cmd, timeout=30, check=False, capture_output=True, text=True):
    """Helper to run subprocess and return CompletedProcess; raises on TimeoutExpired."""
    logging.debug("Running command: %s", " ".join(cmd))
    try:
        cp = subprocess.run(cmd, timeout=timeout, check=False, capture_output=capture_output, text=text)
        logging.debug("Command exit=%s stdout_len=%d stderr_len=%d", cp.returncode, len(cp.stdout or ""), len(cp.stderr or ""))
        if capture_output:
            logging.debug("stdout:\n%s", (cp.stdout or "").strip())
            if cp.stderr:
                logging.debug("stderr:\n%s", cp.stderr.strip())
        if check and cp.returncode != 0:
            raise VPNError(f"Command {' '.join(cmd)} failed with exit {cp.returncode}. stdout:\n{cp.stdout}\nstderr:\n{cp.stderr}")
        return cp
    except subprocess.TimeoutExpired:
        logging.exception("Command timed out: %s", " ".join(cmd))
        raise


def debug_environment():
    """Prints helpful environment diagnostics to the log."""
    try:
        logging.info("=== Debug environment ===")
        logging.debug("Effective UID: %s, EUID: %s, USER env: %s", os.getuid(), os.geteuid(), os.environ.get("USER"))
        for binary in ("wg-quick", "wg", "ip", "sudo"):
            path = shutil.which(binary)
            logging.debug("which %s -> %s", binary, path or "NOT FOUND")
            if path:
                try:
                    out = subprocess.run([path, "--version"], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, timeout=5)
                    logging.debug("%s --version -> %s", binary, out.stdout.strip())
                except Exception:
                    logging.debug("could not run %s --version", binary)
        # /dev/net/tun
        try:
            cp = subprocess.run(["ls", "-l", "/dev/net/tun"], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, timeout=5)
            logging.debug("/dev/net/tun -> %s", cp.stdout.strip())
        except Exception:
            logging.exception("ls /dev/net/tun failed")
        # sudo non-interactive
        try:
            cp = subprocess.run(["sudo", "-n", "true"], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=5)
            if cp.returncode == 0:
                logging.debug("sudo -n true -> OK (no password required for this sudo invocation)")
            else:
                logging.debug("sudo -n true -> exit %s stderr=%s", cp.returncode, cp.stderr.strip())
        except Exception:
            logging.exception("sudo -n true check failed (may prompt for password)")
        logging.info("=== End debug environment ===")
    except Exception:
        logging.exception("Failed collecting environment debug info")


def has_ipv6_in_config(conf_text):
    """Detect if the config has IPv6 in Address line (contains '::' or contains a comma with '::')."""
    m = re.search(r"(?mi)^Address\s*=\s*(.+)$", conf_text)
    if not m:
        return False
    addr_field = m.group(1)
    if "::" in addr_field:
        return True
    if "," in addr_field:
        for p in [p.strip() for p in addr_field.split(",")]:
            if "::" in p:
                return True
    return False


def write_wg_config(config_text):
    """Write the given config_text to WG_CONF_PATH as root (using sudo tee), set mode 600."""
    try:
        logging.info("Writing WireGuard config to %s (sudo tee)...", WG_CONF_PATH)
        cp = subprocess.run(["sudo", "tee", WG_CONF_PATH], input=config_text, text=True, capture_output=True, timeout=20)
        logging.debug("sudo tee exit=%s stdout_len=%d stderr_len=%d", cp.returncode, len(cp.stdout or ""), len(cp.stderr or ""))
        if cp.returncode != 0:
            raise VPNError(f"Failed to write {WG_CONF_PATH} via sudo tee. stdout:\n{cp.stdout}\nstderr:\n{cp.stderr}")
        # chmod 600
        cp2 = subprocess.run(["sudo", "chmod", "600", WG_CONF_PATH], capture_output=True, text=True, timeout=5)
        if cp2.returncode != 0:
            raise VPNError(f"Failed chmod 600 {WG_CONF_PATH}. stderr:\n{cp2.stderr}")
        # quick stat
        cp3 = subprocess.run(["sudo", "stat", "-c", "%a %U:%G %s", WG_CONF_PATH], capture_output=True, text=True, timeout=5)
        logging.debug("stat %s -> %s", WG_CONF_PATH, cp3.stdout.strip())
    except subprocess.TimeoutExpired:
        logging.exception("Timeout while writing config")
        raise
    except Exception:
        logging.exception("Exception while writing wg config")
        raise


def wg_quick_down_ignore():
    """Attempt to bring down wg0 (ignore errors) to ensure a clean start."""
    try:
        logging.info("Attempting sudo wg-quick down %s (ignore errors)...", WG_IFACE)
        cp = subprocess.run(["sudo", "wg-quick", "down", WG_IFACE], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, timeout=15)
        logging.debug("wg-quick down exit=%s output:\n%s", cp.returncode, cp.stdout.strip())
    except subprocess.TimeoutExpired:
        logging.debug("wg-quick down timed out, continuing")
    except Exception:
        logging.exception("wg-quick down failed (ignored)")


def wg_quick_up():
    """Run wg-quick up with timeout and check output; raise VPNError with diagnostics if fails."""
    try:
        wg_quick_bin = shutil.which("wg-quick")
        if not wg_quick_bin:
            raise VPNError("'wg-quick' not found in PATH; install wireguard-tools.")
        logging.info("Running 'sudo wg-quick up %s'...", WG_IFACE)
        cp = subprocess.run(["sudo", "wg-quick", "up", WG_IFACE], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, timeout=30)
        logging.debug("'wg-quick up' exit=%s output:\n%s", cp.returncode, cp.stdout.strip())
        if cp.returncode != 0:
            low = (cp.stdout or "").lower()
            # detect common capability errors
            if "operation not permitted" in low or "rt netlink" in low or "permission denied" in low:
                # collect additional diagnostics
                diag = collect_diagnostics()
                raise VPNError("wg-quick up failed with permission/RTNETLINK error. This usually means missing CAP_NET_ADMIN or TUN access.\nFull output:\n" + cp.stdout + "\nDiagnostics:\n" + diag)
            else:
                diag = collect_diagnostics()
                raise VPNError("wg-quick up failed. Full output:\n" + cp.stdout + "\nDiagnostics:\n" + diag)
    except subprocess.TimeoutExpired:
        logging.exception("wg-quick up timed out")
        diag = collect_diagnostics()
        raise VPNError("wg-quick up timed out. Diagnostics:\n" + diag)
    except Exception:
        logging.exception("Error running wg-quick up")
        raise


def wg_interface_is_up():
    """Return True if wg show wg0 or ip addr show wg0 indicates the interface exists and has addresses."""
    try:
        cp = subprocess.run(["sudo", "wg", "show", WG_IFACE], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=5)
        logging.debug("sudo wg show exit=%s stdout_len=%d stderr_len=%d", cp.returncode, len(cp.stdout or ""), len(cp.stderr or ""))
        if cp.returncode == 0 and cp.stdout.strip():
            return True
        ipcp = subprocess.run(["sudo", "ip", "addr", "show", WG_IFACE], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=5)
        logging.debug("sudo ip addr show exit=%s stdout_len=%d", ipcp.returncode, len(ipcp.stdout or ""))
        if ipcp.returncode == 0 and ipcp.stdout.strip():
            return True
        return False
    except Exception:
        logging.exception("Error checking interface status")
        return False


def collect_diagnostics():
    """Collect helpful diagnostic output (wg show, ip addr, journalctl, dmesg tail) to aid debugging."""
    parts = []
    try:
        cp1 = subprocess.run(["sudo", "wg", "show"], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, timeout=5)
        parts.append("--- sudo wg show ---\n" + (cp1.stdout or ""))
    except Exception:
        parts.append("--- wg show failed ---")
    try:
        cp2 = subprocess.run(["sudo", "ip", "addr"], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, timeout=5)
        parts.append("--- sudo ip addr ---\n" + (cp2.stdout or ""))
    except Exception:
        parts.append("--- ip addr failed ---")
    try:
        # journalctl may not be available or may require privileges; still try to fetch wg-quick unit logs
        cp3 = subprocess.run(["sudo", "journalctl", "-u", f"wg-quick@{WG_IFACE}", "-n", "200", "--no-pager"], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, timeout=8)
        parts.append("--- journalctl wg-quick ---\n" + (cp3.stdout or ""))
    except Exception:
        try:
            cp3 = subprocess.run(["sudo", "journalctl", "-n", "200", "--no-pager"], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, timeout=8)
            parts.append("--- journalctl (tail) ---\n" + (cp3.stdout or ""))
        except Exception:
            parts.append("--- journalctl failed ---")
    try:
        cp4 = subprocess.run(["dmesg", "--ctime", "-T", "|", "tail", "-n", "200"], shell=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, timeout=8)
        parts.append("--- dmesg tail ---\n" + (cp4.stdout or ""))
    except Exception:
        parts.append("--- dmesg failed ---")
    return "\n".join(parts)


def get_public_ip(url):
    """Try to obtain public IP using provided url (expects JSON with 'ip')."""
    try:
        logging.debug("Fetching public IP from %s", url)
        r = requests.get(url, timeout=10)
        r.raise_for_status()
        obj = r.json()
        if isinstance(obj, dict) and "ip" in obj:
            return obj["ip"]
        return str(obj)
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
        logging.info("Starting vpn_connect.py (WireGuard) — logging DEBUG enabled.")
        debug_environment()

        # Prefer configs with IPv6 when possible
        all_keys = list(CONFIGS.keys())
        ipv6_keys = [k for k, v in CONFIGS.items() if has_ipv6_in_config(v)]
        if ipv6_keys:
            chosen = random.choice(ipv6_keys)
            logging.info("Choosing from IPv6-capable configs: %s", ipv6_keys)
        else:
            chosen = random.choice(all_keys)
            logging.warning("No IPv6-capable config found; choosing from all configs (may be IPv4-only).")
        conf_text = CONFIGS[chosen]
        supports_ipv6 = has_ipv6_in_config(conf_text)
        logging.info("Chosen config: %s (supports_ipv6=%s)", chosen, supports_ipv6)

        # Write config
        write_wg_config(conf_text)

        # Ensure previous instance down (ignore errors)
        wg_quick_down_ignore()

        # Bring up wg0
        wg_quick_up()

        # Wait for interface up
        logging.info("Waiting up to %ss for interface %s to be active...", WAIT_WG_TIMEOUT, WG_IFACE)
        start = time.time()
        while time.time() - start < WAIT_WG_TIMEOUT:
            if wg_interface_is_up():
                logging.info("Interface %s is active.", WG_IFACE)
                break
            logging.debug("Interface %s not active yet; sleeping 1s...", WG_IFACE)
            time.sleep(1)
        else:
            diag = collect_diagnostics()
            raise VPNError(f"Interface {WG_IFACE} did not become active within {WAIT_WG_TIMEOUT}s.\nDiagnostics:\n{diag}")

        # After active, fetch IPs and geolocation
        ipv4 = get_public_ip("https://api.ipify.org?format=json")
        ipv6 = get_public_ip("https://api64.ipify.org?format=json")
        geo = get_geolocation()

        logging.info("===== VPN Connection Summary =====")
        logging.info("Config chosen: %s", chosen)
        logging.info("Supports IPv6: %s", "YES" if supports_ipv6 else "NO")
        logging.info("IPv4: %s", ipv4 or "N/A")
        logging.info("IPv6: %s", ipv6 or "N/A")
        logging.info("City: %s", geo.get("city") or "N/A")
        logging.info("Region: %s", geo.get("region") or "N/A")
        logging.info("Country: %s", geo.get("country") or "N/A")
        logging.info("ISP/Org: %s", geo.get("isp") or "N/A")
        logging.info("===================================")

        # Infinite monitoring loop
        logging.info("Entering monitoring loop (interval %ds). Cancel manually or wait for job timeout.", MONITOR_INTERVAL)
        while True:
            time.sleep(MONITOR_INTERVAL)
            try:
                up = wg_interface_is_up()
                if not up:
                    logging.error("Interface %s is no longer active!", WG_IFACE)
                else:
                    logging.debug("Interface %s still active.", WG_IFACE)
                ipv4 = get_public_ip("https://api.ipify.org?format=json")
                ipv6 = get_public_ip("https://api64.ipify.org?format=json")
                logging.info("Heartbeat: config=%s | IPv4=%s | IPv6=%s", chosen, ipv4 or "N/A", ipv6 or "N/A")
            except Exception:
                logging.exception("Exception in monitoring loop; continuing.")
    except Exception as exc:
        logging.exception("Unhandled exception in main(): %s", exc)
        # Also attempt to gather diagnostics for logs
        try:
            diag = collect_diagnostics()
            logging.error("Diagnostics at failure:\n%s", diag)
        except Exception:
            logging.exception("Failed to collect diagnostics at failure")
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
