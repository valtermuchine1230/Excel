#!/usr/bin/env python3
"""
AVISO: As credenciais estão hardcoded diretamente neste ficheiro POR PROPÓSITO do pedido.
Nunca faças isto em repositórios reais ou em produção — usa sempre GitHub Secrets, variáveis de ambiente seguras, ou vaults.
"""

import os
import sys
import time
import random
import subprocess
import logging
import traceback

import pexpect
import requests

# -------------------------
# CREDENCIAIS (hardcoded)
# -------------------------
USERNAME = "Valter3B2"
PASSWORD = "dshgfajshsksvv+cano@gmail.com"
# -------------------------

# Logging setup
logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%dT%H:%M:%S%z",
    handlers=[logging.StreamHandler(sys.stdout)]
)

# Constants
HIDE_BINARY = "./hide.me"
VPN_IFACE = "vpn"
INTERFACE_CHECK_TIMEOUT = 30  # seconds
INTERFACE_POLL_INTERVAL = 2    # seconds
RETRY_SLEEP = 60               # seconds between periodic checks in the infinite loop

class VPNError(Exception):
    pass

class AuthError(VPNError):
    pass

class TunDeviceError(VPNError):
    pass

def run_cmd(cmd, timeout=30, check=True):
    logging.debug("Running command: %s", " ".join(cmd))
    try:
        completed = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, timeout=timeout)
        logging.debug("Return code: %s", completed.returncode)
        logging.debug("Output:\n%s", completed.stdout)
        if check and completed.returncode != 0:
            raise VPNError(f"Command {cmd} failed with code {completed.returncode}")
        return completed.stdout
    except subprocess.TimeoutExpired as e:
        logging.error("Command timed out: %s", cmd)
        raise

def list_free_servers():
    if not os.path.exists(HIDE_BINARY):
        raise VPNError(f"{HIDE_BINARY} not found. Ensure binary was built and is in the working directory.")
    try:
        out = run_cmd([HIDE_BINARY, "list", "free"])
        # Basic parsing: assume each server on its own line; filter empties
        lines = [l.strip() for l in out.splitlines()]
        servers = [l for l in lines if l and not l.lower().startswith("usage") and not l.lower().startswith("note")]
        if not servers:
            raise VPNError("No free servers found in output of './hide.me list free'. Output:\n" + out)
        logging.debug("Parsed servers: %s", servers)
        return servers
    except Exception as e:
        logging.exception("Failed to list free servers")
        raise

def get_access_token(server):
    """
    Uses pexpect to interact with './hide.me -u {username} token {server}' and send PASSWORD.
    """
    cmd = f"{HIDE_BINARY} -u {USERNAME} token {server}"
    logging.debug("Requesting access token with command: %s", cmd)
    try:
        child = pexpect.spawn(cmd, encoding='utf-8', timeout=30)
        # Log pexpect output to stdout so it's visible in workflow logs
        child.logfile = sys.stdout

        # Expect password prompt or EOF
        i = child.expect([r"[Pp]assword", r"Password:", pexpect.EOF, pexpect.TIMEOUT], timeout=20)
        if i == 0 or i == 1:
            logging.debug("Password prompt detected, sending password (via pexpect).")
            child.sendline(PASSWORD)
            # Read until EOF
            child.expect(pexpect.EOF, timeout=20)
            full_output = child.before or ""
        elif i == 2:
            full_output = child.before or ""
            logging.debug("Command ended without explicit password prompt; captured output.")
        else:
            child.close()
            raise VPNError("Timeout or unexpected response when requesting token.")
        child.close()

        logging.debug("Token command output:\n%s", full_output)

        # Look for common failure indicators
        low = full_output.lower()
        if ("incorrect" in low) or ("authentication" in low and "failed" in low) or ("invalid" in low):
            raise AuthError("Authentication failed when requesting access token. Check username/password.")

        # Try to extract a token-looking string (very loose)
        for part in full_output.split():
            if len(part) > 10 and all(c.isalnum() or c in "-._~+=" for c in part):
                logging.debug("Heuristic token candidate: %s", part)
                return part

        # If we didn't find a token, still return full output for debugging
        logging.warning("Could not heuristically parse a token from output; returning full output for inspection.")
        return full_output.strip()
    except pexpect.exceptions.ExceptionPexpect as e:
        logging.exception("pexpect interaction failed")
        raise

def start_connection(server):
    logging.info("Starting VPN connection to server: %s", server)
    # Use sudo to start connection (user requested sudo)
    # Start as subprocess and leave it running (so the connection persists)
    try:
        proc = subprocess.Popen(["sudo", HIDE_BINARY, "connect", server],
                                stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
        logging.debug("Started 'hide.me connect' with PID %s", proc.pid)
        # Stream output lines in background logging
        def stream_output(p):
            try:
                for line in iter(p.stdout.readline, ''):
                    if not line:
                        break
                    logging.debug("[hide.me connect] %s", line.rstrip())
            except Exception:
                logging.exception("Error while streaming connect output")

        # Spawn a small thread to stream output so we can continue to check interface
        import threading
        t = threading.Thread(target=stream_output, args=(proc,), daemon=True)
        t.start()
        return proc
    except Exception:
        logging.exception("Failed to start hide.me connect")
        raise

def interface_is_up():
    try:
        out = run_cmd(["ip", "addr", "show", VPN_IFACE], check=False)
        up = ("state UP" in out) or ("UP" in out.splitlines()[0]) if out else False
        logging.debug("Interface %s check result: %s", VPN_IFACE, up)
        return up
    except Exception:
        logging.exception("Error checking interface with ip addr")
        return False

def wait_for_interface(timeout=INTERFACE_CHECK_TIMEOUT):
    logging.info("Waiting for interface '%s' to become active (timeout %ss)...", VPN_IFACE, timeout)
    start = time.time()
    attempt = 0
    while time.time() - start < timeout:
        attempt += 1
        logging.debug("Interface check attempt %d", attempt)
        if interface_is_up():
            logging.info("Interface '%s' is now active.", VPN_IFACE)
            return True
        logging.debug("Interface '%s' not active yet; sleeping %ss", VPN_IFACE, INTERFACE_POLL_INTERVAL)
        time.sleep(INTERFACE_POLL_INTERVAL)
    logging.error("Timeout waiting for interface '%s' to become active after %s seconds", VPN_IFACE, timeout)
    return False

def get_public_ip(url):
    try:
        logging.debug("Querying %s", url)
        r = requests.get(url, timeout=10)
        r.raise_for_status()
        data = r.json()
        # ipify returns {"ip":"x.x.x.x"}
        if isinstance(data, dict) and "ip" in data:
            return data["ip"]
        # fallback: entire body
        return r.text.strip()
    except requests.RequestException:
        logging.exception("Failed to fetch public IP from %s", url)
        raise

def get_geolocation():
    try:
        r = requests.get("https://ipapi.co/json/", timeout=10)
        r.raise_for_status()
        data = r.json()
        return {
            "city": data.get("city"),
            "region": data.get("region"),
            "country": data.get("country_name") or data.get("country"),
            "isp": data.get("org") or data.get("isp") or data.get("asn")
        }
    except requests.RequestException:
        logging.exception("Failed to fetch geolocation from ipapi.co")
        raise

def main():
    try:
        logging.info("VPN connect script started (logging DEBUG enabled).")
        # Pre-check TUN device
        if not os.path.exists("/dev/net/tun"):
            raise TunDeviceError("/dev/net/tun not available. TUN device is required.")

        # Obtain list of free servers and pick a random one
        servers = list_free_servers()
        server = random.choice(servers)
        logging.info("Selected server: %s", server)

        # Request access token (pexpect)
        try:
            token = get_access_token(server)
            logging.info("Token obtained (or token output logged).")
            logging.debug("Token/raw-token-output: %s", token)
        except AuthError as ae:
            logging.error("Authentication error obtaining token: %s", ae)
            raise

        # Start VPN connection
        proc = start_connection(server)

        # Wait for interface
        if not wait_for_interface():
            raise VPNError("VPN interface failed to come up within timeout.")

        # Once interface up, fetch IPv4, IPv6, geolocation
        ipv4 = None
        ipv6 = None
        try:
            ipv4 = get_public_ip("https://api.ipify.org?format=json")
        except Exception:
            logging.warning("Failed to get IPv4 address; proceeding.")

        try:
            ipv6 = get_public_ip("https://api64.ipify.org?format=json")
        except Exception:
            logging.warning("Failed to get IPv6 address; proceeding.")

        geo = {}
        try:
            geo = get_geolocation()
        except Exception:
            logging.warning("Failed to get geolocation; proceeding.")

        logging.info("===== VPN Connection Summary =====")
        logging.info("Server chosen: %s", server)
        logging.info("IPv4: %s", ipv4 or "N/A")
        logging.info("IPv6: %s", ipv6 or "N/A")
        logging.info("City: %s", geo.get("city"))
        logging.info("Region: %s", geo.get("region"))
        logging.info("Country: %s", geo.get("country"))
        logging.info("ISP/Org: %s", geo.get("isp"))
        logging.info("==================================")

        # Enter infinite monitoring loop: every 60s check interface and re-print IPs
        logging.info("Entering infinite monitoring loop; will check interface every %d seconds.", RETRY_SLEEP)
        while True:
            time.sleep(RETRY_SLEEP)
            try:
                if not interface_is_up():
                    logging.error("VPN interface '%s' is no longer up!", VPN_IFACE)
                else:
                    logging.debug("VPN interface '%s' is still up.", VPN_IFACE)
                # Re-fetch IPs to detect changes
                try:
                    ipv4 = get_public_ip("https://api.ipify.org?format=json")
                except Exception:
                    ipv4 = None
                try:
                    ipv6 = get_public_ip("https://api64.ipify.org?format=json")
                except Exception:
                    ipv6 = None
                logging.info("Heartbeat: server=%s | IPv4=%s | IPv6=%s", server, ipv4 or "N/A", ipv6 or "N/A")
            except Exception:
                logging.exception("Error during monitoring loop iteration; continuing.")
    except TunDeviceError as te:
        logging.exception("TUN device error: %s", te)
        sys.exit(2)
    except AuthError as ae:
        logging.exception("Authentication error: %s", ae)
        sys.exit(3)
    except Exception:
        logging.exception("Unhandled exception in main(); exiting.")
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()
