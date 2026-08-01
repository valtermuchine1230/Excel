#!/usr/bin/env python3
# vpn_connect_debug.py — versão muito verbosa com prints imediatos e flush
# Substitui o teu vpn_connect.py com esta versão para diagnóstico e operação.

import os, sys, time, subprocess, traceback, random, re
from datetime import datetime

try:
    import requests
except Exception:
    print("ERROR: requests não disponível - instala com python3 -m pip install requests", flush=True)

WG_CONF_PATH = "/etc/wireguard/wg0.conf"
WG_IFACE = "wg0"
WAIT_IFACE_TIMEOUT = 20
MONITOR_INTERVAL = 60

CONFIGS = {
    "JP-FREE-16": """[Interface]
PrivateKey = SDF5r+E6IwHBaalMuYKFkj4Vr1mQj+PKzp6vI/X8LWk=
Address = 10.2.0.2/32, 2a07:b944::2:2/128
DNS = 10.2.0.1, 2a07:b944::2:1

[Peer]
PublicKey = BbghXRtbSYBJ2Q/eMu4JV7u8LKiKDfgybk7IJAO7iAU=
AllowedIPs = 0.0.0.0/0, ::/0
Endpoint = 45.87.213.226:51820

# Endpoint = [2001:ac8:40:26::10]:51820
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
PrivateKey = mClSo+prm7i2Geox/4fk9OSJTPp7J4HAexG4axYf8Fo=
Address = 10.2.0.2/32
DNS = 10.2.0.1

[Peer]
PublicKey = SOXFyakZ9HI9TeiMRyMoy3PXYEzJJ/IDJcMvxZ3uWSE=
AllowedIPs = 0.0.0.0/0, ::/0
Endpoint = 149.102.254.90:51820

PersistentKeepalive = 25
""",
    "US-FREE-33-B": """[Interface]
PrivateKey = QEqvyYkrkN17W+ixWBII8etwKY/859raWSxcHGVWvkg=
Address = 10.2.0.2/32
DNS = 10.2.0.1

[Peer]
PublicKey = SOXFyakZ9HI9TeiMRyMoy3PXYEzJJ/IDJcMvxZ3uWSE=
AllowedIPs = 0.0.0.0/0, ::/0
Endpoint = 149.102.254.90:51820

PersistentKeepalive = 25
"""
}

def now(): return datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC")

def log(msg):
    print(f"{now()} {msg}", flush=True)

def run_cmd(cmd, timeout=30):
    """Run and return (rc, output). Always returns quickly or raises TimeoutExpired."""
    log(f">>> RUN: {' '.join(cmd)}")
    try:
        cp = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, timeout=timeout)
        out = cp.stdout or ""
        log(f"<<< RC={cp.returncode} | output (first 2000 chars):\n{out[:2000]}")
        return cp.returncode, out
    except subprocess.TimeoutExpired:
        log(f"!!! TIMEOUT running: {' '.join(cmd)}")
        raise

def write_config(conf_text):
    log("Writing WireGuard config to " + WG_CONF_PATH + " using sudo tee")
    # write via Popen to feed input reliably
    p = subprocess.Popen(["sudo", "tee", WG_CONF_PATH], stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
    out, _ = p.communicate(conf_text)
    log(f"tee returned rc={p.returncode}; stdout first 2000 chars:\n{(out or '')[:2000]}")
    if p.returncode != 0:
        raise RuntimeError("sudo tee failed")
    rc, _ = run_cmd(["sudo","chmod","600",WG_CONF_PATH], timeout=5)
    if rc != 0:
        raise RuntimeError("chmod failed")

def collect_brief():
    parts=[]
    try:
        rc,out = run_cmd(["sudo","wg","show",WG_IFACE], timeout=5)
        parts.append(("wg show", out[:2000]))
    except Exception as e:
        parts.append(("wg show","err:"+str(e)))
    try:
        rc,out = run_cmd(["sudo","ip","addr","show",WG_IFACE], timeout=5)
        parts.append(("ip addr show", out[:2000]))
    except Exception as e:
        parts.append(("ip addr show","err:"+str(e)))
    try:
        rc,out = run_cmd(["sudo","cat",WG_CONF_PATH], timeout=5)
        parts.append(("wg0.conf", out[:2000]))
    except Exception as e:
        parts.append(("wg0.conf","err:"+str(e)))
    return parts

def main():
    try:
        log("[STEP] start script")
        log("[STEP] check sudo -n true")
        try:
            rc,_ = run_cmd(["sudo","-n","true"], timeout=5)
            if rc==0:
                log("sudo non-interactive: OK")
            else:
                log("sudo non-interactive: not OK (rc!=0)")
        except Exception:
            log("sudo -n true raised exception")

        log("[STEP] check /dev/net/tun")
        try:
            rc,out = run_cmd(["ls","-l","/dev/net/tun"], timeout=5)
        except Exception:
            log("ls /dev/net/tun failed")

        log("[STEP] modprobe wireguard")
        try:
            run_cmd(["sudo","modprobe","wireguard"], timeout=5)
        except Exception:
            log("modprobe wireguard may have failed (continuing)")

        log("[STEP] enable sysctl forwarding (ipv4 and ipv6)")
        try:
            run_cmd(["sudo","sysctl","-w","net.ipv4.ip_forward=1"], timeout=5)
            run_cmd(["sudo","sysctl","-w","net.ipv6.conf.all.forwarding=1"], timeout=5)
        except Exception:
            log("sysctl commands may have failed (continuing)")

        # choose config (prefer IPv6)
        log("[STEP] choose config")
        ipv6_candidates = [k for k,v in CONFIGS.items() if re.search(r'(::)', v) or (',' in v and '::' in v)]
        if ipv6_candidates:
            chosen = random.choice(ipv6_candidates)
            log(f"Choosing IPv6-capable config: {chosen}")
        else:
            chosen = random.choice(list(CONFIGS.keys()))
            log(f"Choosing config: {chosen}")
        conf = CONFIGS[chosen]

        # show endpoint line for visibility
        m = re.search(r'(?mi)^\s*Endpoint\s*=\s*(.+)$', conf)
        endpoint = m.group(1).strip() if m else "(none)"
        log(f"Endpoint in config: {endpoint}")

        # write config
        log("[STEP] write config to file")
        write_config(conf)
        log("[STEP] stat config")
        try:
            run_cmd(["sudo","stat","-c","%a %U:%G %s", WG_CONF_PATH], timeout=5)
        except Exception:
            log("stat failed")

        # down (ignore)
        log("[STEP] wg-quick down (ignore errors)")
        try:
            run_cmd(["sudo","wg-quick","down",WG_IFACE], timeout=10)
        except Exception:
            log("wg-quick down may have failed or not existed")

        # up
        log("[STEP] wg-quick up — this is the critical step that must succeed to get interface")
        try:
            rc,out = run_cmd(["sudo","wg-quick","up",WG_IFACE], timeout=60)
            if rc != 0:
                log("wg-quick up returned non-zero; collecting brief diagnostics")
                parts = collect_brief()
                for k,v in parts:
                    log(f"--- {k} ---\n{v[:2000]}")
                log("Exiting due to wg-quick up failure")
                sys.exit(1)
            else:
                log("wg-quick up returned 0")
        except Exception as e:
            log("Exception during wg-quick up: " + str(e))
            parts = collect_brief()
            for k,v in parts:
                log(f"--- {k} ---\n{v[:2000]}")
            log("Exiting after exception in wg-quick up")
            sys.exit(1)

        # wait for interface presence
        log("[STEP] wait for wg interface")
        start = time.time()
        up = False
        while time.time() - start < WAIT_IFACE_TIMEOUT:
            try:
                rc,out = run_cmd(["sudo","wg","show",WG_IFACE], timeout=5)
                if rc==0 and out.strip():
                    up = True
                    break
                rc2,out2 = run_cmd(["sudo","ip","addr","show",WG_IFACE], timeout=5)
                if rc2==0 and out2.strip():
                    up = True
                    break
            except Exception:
                pass
            time.sleep(1)
        if not up:
            log("Interface did not appear within timeout; collecting brief diagnostics")
            parts = collect_brief()
            for k,v in parts:
                log(f"--- {k} ---\n{v[:2000]}")
            sys.exit(1)

        log("[STEP] interface up! fetching public IPs")
        # fetch IPs
        try:
            ipv4 = requests.get("https://api.ipify.org?format=json", timeout=8).json().get("ip")
        except Exception:
            ipv4 = None
        try:
            ipv6 = requests.get("https://api64.ipify.org?format=json", timeout=8).json().get("ip")
        except Exception:
            ipv6 = None
        try:
            geo = requests.get("https://ipapi.co/json/", timeout=8).json()
        except Exception:
            geo = {}

        log("=== SUMMARY ===")
        log(f"Config chosen: {chosen}")
        log(f"Interface is up: {WG_IFACE}")
        log(f"Public IPv4: {ipv4 or 'N/A'}")
        log(f"Public IPv6: {ipv6 or 'N/A'}")
        log(f"Location: {geo.get('city')}, {geo.get('region')}, {geo.get('country_name') or geo.get('country')}")
        log("=== END SUMMARY ===")

        # monitor loop
        log("Entering monitor loop, printing heartbeat every 60s")
        while True:
            time.sleep(MONITOR_INTERVAL)
            try:
                rc,out = run_cmd(["sudo","ip","addr","show",WG_IFACE], timeout=5)
                log(f"HEARTBEAT: iface present rc={rc}; ip info head:\n{out.splitlines()[:10]}")
                try:
                    ipv4 = requests.get("https://api.ipify.org?format=json", timeout=8).json().get("ip")
                except Exception:
                    ipv4 = None
                try:
                    ipv6 = requests.get("https://api64.ipify.org?format=json", timeout=8).json().get("ip")
                except Exception:
                    ipv6 = None
                log(f"HEARTBEAT IPs: IPv4={ipv4 or 'N/A'} IPv6={ipv6 or 'N/A'}")
            except Exception:
                log("HEARTBEAT: exception in monitor loop")
    except Exception:
        log("Unhandled exception:")
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()
