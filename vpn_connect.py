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
    # se há vírgula, pode ser combo IPv4,IPv6
    if "," in addr_field:
        # verificar se algum dos endereços tem '::'
        parts = [p.strip() for p in addr_field.split(",")]
        for p in parts:
            if "::" in p:
                return True
    return False

def write_wg_config_as_root(config_text):
    """
    Escreve o config para /etc/wireguard/wg0.conf usando sudo tee.
    """
    try:
        logging.debug("Escrevendo configuração para %s com sudo tee (necessário privilégio root).", WG_CONF_PATH)
        completed = subprocess.run(
            ["sudo", "tee", WG_CONF_PATH],
            input=config_text,
            text=True,
            capture_output=True,
            check=False
        )
        logging.debug("sudo tee returncode=%s", completed.returncode)
        if completed.stdout:
            logging.debug("sudo tee stdout:\n%s", completed.stdout)
        if completed.stderr:
            logging.debug("sudo tee stderr:\n%s", completed.stderr)
        if completed.returncode != 0:
            raise VPNError(f"Falha ao escrever {WG_CONF_PATH} (returncode {completed.returncode})\nStderr: {completed.stderr}")
    except Exception:
        logging.exception("Erro ao escrever ficheiro de configuração WireGuard")
        raise

def run_wg_quick_up():
    """
    Executa 'sudo wg-quick up wg0' e loga todo o output.
    """
    try:
        logging.info("Executando 'sudo wg-quick up wg0' ...")
        proc = subprocess.run(
            ["sudo", "wg-quick", "up", WG_IFACE],
            capture_output=True,
            text=True,
            check=False
        )
        logging.debug("wg-quick up returncode=%s", proc.returncode)
        logging.debug("wg-quick up stdout:\n%s", proc.stdout)
        logging.debug("wg-quick up stderr:\n%s", proc.stderr)
        if proc.returncode != 0:
            raise VPNError(f"'wg-quick up' falhou com code {proc.returncode}\nstdout:\n{proc.stdout}\nstderr:\n{proc.stderr}")
    except Exception:
        logging.exception("Erro ao executar wg-quick up")
        raise

def wg_interface_is_up():
    """
    Verifica se a interface wg0 está ativa usando 'wg show wg0'.
    Retorna True se wg show devolve código 0 e output não vazio.
    """
    try:
        proc = subprocess.run(["wg", "show", WG_IFACE], capture_output=True, text=True, check=False)
        up = (proc.returncode == 0 and (proc.stdout.strip() != "" or proc.stderr.strip() == ""))
        logging.debug("wg show returncode=%s; stdout_len=%d; up=%s", proc.returncode, len(proc.stdout or ""), up)
        return up
    except Exception:
        logging.exception("Erro ao verificar interface wg0 com 'wg show'")
        return False

def get_public_ip(url):
    try:
        logging.debug("Acedendo %s", url)
        r = requests.get(url, timeout=10)
        r.raise_for_status()
        data = r.json()
        if isinstance(data, dict) and "ip" in data:
            return data["ip"]
        return r.text.strip()
    except Exception:
        logging.exception("Falha ao obter IP de %s", url)
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
        logging.exception("Falha ao obter geolocalização")
        return {}

def main():
    try:
        logging.info("Iniciando vpn_connect (modo WireGuard) — logging DEBUG ativo.")
        # Escolha do servidor com preferência para configs que tenham IPv6
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

        # Escrever config em /etc/wireguard/wg0.conf (necessita sudo)
        write_wg_config_as_root(config_text)

        # Levantar wg0
        run_wg_quick_up()

        # Esperar interface wg0 ficar activa (usa wg show)
        logging.info("Aguardando interface %s ficar ativa (timeout %ss)...", WG_IFACE, WAIT_WG_TIMEOUT)
        start = time.time()
        while time.time() - start < WAIT_WG_TIMEOUT:
            if wg_interface_is_up():
                logging.info("Interface %s está ativa.", WG_IFACE)
                break
            logging.debug("Interface %s ainda não ativa — aguardando 1s...", WG_IFACE)
            time.sleep(1)
        else:
            raise VPNError(f"Interface {WG_IFACE} não ficou ativa após {WAIT_WG_TIMEOUT} segundos.")

        # Buscar IPv4, IPv6 e geolocalização
        ipv4 = get_public_ip("https://api.ipify.org?format=json")
        ipv6 = get_public_ip("https://api64.ipify.org?format=json")
        geo = get_geolocation()

        # Log formatado claro
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

        # Loop infinito de monitorização (verifica wg0 e reimprime IPs a cada 60s)
        logging.info("Entrando em loop infinito de monitorização (intervalo %ds). Cancelar manualmente ou aguardar timeout do workflow.", MONITOR_INTERVAL)
        while True:
            time.sleep(MONITOR_INTERVAL)
            try:
                if not wg_interface_is_up():
                    logging.error("Interface %s deixou de estar activa!", WG_IFACE)
                else:
                    logging.debug("Interface %s permanece activa.", WG_IFACE)
                # Rebuscar IPs para detectar mudanças
                ipv4 = get_public_ip("https://api.ipify.org?format=json")
                ipv6 = get_public_ip("https://api64.ipify.org?format=json")
                logging.info("Heartbeat: servidor=%s | IPv4=%s | IPv6=%s", chosen, ipv4 or "N/A", ipv6 or "N/A")
            except Exception:
                logging.exception("Erro na iteração do loop de monitorização; aitilizações continuarão.")
    except Exception as e:
        logging.exception("Excepção não tratada no main(): %s", e)
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()
