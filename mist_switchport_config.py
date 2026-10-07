#!/usr/bin/env python3
"""Export the derived configuration of every switch port on every site in a
Mist org to Excel.

Mist builds a switch port's config from several layers, each able to override
the one below it:

  1. Switch Template (org networktemplate) - port profiles, networks, rules
  2. Site settings - port profiles / networks merged by name, switch rules
  3. Switch matching rule (first rule matching the switch) - default port
     profile for every port, plus per-port assignments
  4. Switch (device) config - its own port_config, port profiles, networks
     and vars
  5. Local port config (changes made on the switch where allowed)

Layers 1+2 are taken from Mist's own /sites/{id}/setting/derived merge. The
rest is resolved here, then optionally checked against the Junos config Mist
actually generates for each switch (/devices/{id}/config_cmd).

org_id, api_token and cloud are read from mist_switchport_config.ini.
"""
import configparser
import json
import os
import re
import sys
import time
from collections import Counter, defaultdict
from datetime import datetime
from urllib.parse import urlparse

import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INI_NAME = "mist_switchport_config.ini"
config = configparser.ConfigParser()
config.read(os.path.join(SCRIPT_DIR, INI_NAME))

ORG_ID = config.get("mist", "org_id", fallback="").strip()
API_TOKEN = config.get("mist", "api_token", fallback="").strip()
# Mist cloud the org lives on, e.g. "eu", "gc3", "api.eu.mist.com" or the
# portal URL "https://manage.eu.mist.com". Required.
CLOUD = config.get("mist", "cloud", fallback="").strip()
OUTPUT_DIR = os.path.expanduser(config.get("output", "directory", fallback="~").strip() or "~")
# How many switches to check against Mist's generated Junos config:
# "all", "0" / "off", or a number (spread evenly across the org).
JUNOS_CHECK = config.get("validation", "junos_check", fallback="all").strip().lower()

REQUEST_TIMEOUT = 60
API_BASE = None
API_HOST = None
API_CALLS = Counter()

PHYSICAL_PORT = re.compile(r"^(ge|xe|et|mge)-(\d+)/(\d+)/(\d+)$")

# Port profiles Mist defines itself (not visible in templates / site settings).
SYSTEM_USAGES = {
    "default": {"mode": "access", "port_network": "default"},
    "disabled": {"mode": "access", "disabled": True},
    "ap": {"mode": "trunk", "port_network": "default", "all_networks": True, "stp_edge": True},
    "uplink": {"mode": "trunk", "port_network": "default", "all_networks": True},
    "iot": {"mode": "access", "port_network": "default", "stp_edge": True},
    "inet": {"mode": "routed (L3)"},
    "evpn_uplink": {"mode": "EVPN fabric (L3)", "mtu": 9192},
    "evpn_downlink": {"mode": "EVPN fabric (L3)", "mtu": 9192},
    "evpn_esilag_access": {"mode": "EVPN fabric (L3)", "mtu": 9192},
}
SYSTEM_NETWORKS = {"default": {"vlan_id": 1}}

# Fields a switch's port_config entry may set to override its port profile.
PORT_CONFIG_OVERRIDES = ("port_network", "networks", "speed", "duplex", "mtu", "poe_disabled",
                         "disable_autoneg", "mac_limit", "voip_network")


def build_session():
    """Shared HTTP session, wired up for a TLS-inspecting proxy (e.g. Zscaler)
    when the optional [network] section is present in the .ini. With no
    [network] section, behaviour is unchanged (certifi CA store, and the
    HTTP_PROXY / HTTPS_PROXY / NO_PROXY environment variables if set)."""
    session = requests.Session()
    session.headers.update({"Authorization": f"Token {API_TOKEN}"})
    if "network" not in config:
        return session

    section = config["network"]
    ca_bundle = section.get("ca_bundle", "").strip()
    if ca_bundle:
        ca_path = os.path.expanduser(ca_bundle)
        if not os.path.isfile(ca_path):
            sys.exit(f"[network] ca_bundle does not exist: {ca_path}\n"
                     f"This should be a PEM file containing your proxy's root CA "
                     f"certificate (e.g. exported from Zscaler).")
        session.verify = ca_path
    elif not section.getboolean("verify_ssl", fallback=True):
        session.verify = False
        print("WARNING: TLS certificate verification is DISABLED ([network] verify_ssl = false). "
              "Only use this as a last resort.")
        requests.packages.urllib3.disable_warnings(requests.packages.urllib3.exceptions.InsecureRequestWarning)

    for scheme in ("http", "https"):
        proxy = section.get(f"{scheme}_proxy", "").strip()
        if proxy:
            session.proxies[scheme] = proxy
    return session


SESSION = build_session()


def http_get(url, timeout=REQUEST_TIMEOUT):
    """GET with friendly errors for the usual proxy / Zscaler failure modes."""
    try:
        return SESSION.get(url, timeout=timeout)
    except requests.exceptions.SSLError as e:
        sys.exit(f"\nTLS certificate verification failed for {url}: {e}\n"
                 f"If you're behind a TLS-inspecting proxy (e.g. Zscaler), set [network] ca_bundle "
                 f"in {INI_NAME} to the path of its root CA certificate (PEM format).")
    except requests.exceptions.ProxyError as e:
        sys.exit(f"\nCould not reach the proxy for {url}: {e}\n"
                 f"Check [network] http_proxy / https_proxy in {INI_NAME} "
                 f"or your HTTP(S)_PROXY environment variables.")


def format_eta(seconds):
    if seconds < 60:
        return f"{int(seconds)}s"
    m, s = divmod(int(seconds), 60)
    return f"{m}m {s}s"


def api_get(path, label, max_retries=5):
    """GET {API_BASE}{path} (or a full URL) with 429 back-off; counts calls by label."""
    url = path if path.startswith("http") else f"{API_BASE}{path}"
    for attempt in range(1, max_retries + 1):
        resp = http_get(url)
        API_CALLS[label] += 1
        if resp.status_code == 429:
            retry_after = int(resp.headers.get("Retry-After", 30))
            print(f"\n  Rate limited - waiting {retry_after}s (attempt {attempt}/{max_retries})...")
            time.sleep(retry_after)
            continue
        resp.raise_for_status()
        return resp.json()
    resp.raise_for_status()


def api_get_pages(path, label):
    """Page-numbered list endpoints (limit=1000&page=N)."""
    results, page = [], 1
    sep = "&" if "?" in path else "?"
    while True:
        data = api_get(f"{path}{sep}limit=1000&page={page}", label)
        results.extend(data)
        if len(data) < 1000:
            return results
        page += 1


def api_search(path, label):
    """Search endpoints - follow the `next` cursor (never page=, total is unreliable)."""
    results = []
    url = path
    while url:
        data = api_get(url, label)
        results.extend(data.get("results", []))
        sys.stdout.write(f"\r  {len(results)} records...")
        sys.stdout.flush()
        nxt = data.get("next")
        url = f"{API_HOST}{nxt}" if nxt else None
    print()
    return results


def cloud_to_host(cloud):
    """Normalise the .ini cloud value to an API host URL like https://api.eu.mist.com."""
    value = cloud.lower().strip()
    if "://" in value:
        value = urlparse(value).netloc
    value = value.strip("/")
    if value.endswith("mist.com"):
        if value.startswith("manage."):
            value = "api." + value[len("manage."):]
        return f"https://{value}"
    if value in ("us", "global", "global01"):
        return "https://api.mist.com"
    return f"https://api.{value}.mist.com"


def fetch_org_info(host):
    """Return org_info from the configured cloud, or exit with a clear reason."""
    try:
        resp = http_get(f"{host}/api/v1/orgs/{ORG_ID}", timeout=15)
    except (requests.exceptions.ConnectionError, requests.exceptions.Timeout) as e:
        sys.exit(f"Could not connect to {host}: {e}\nCheck the cloud setting in {INI_NAME}.")
    API_CALLS["org"] += 1
    if resp.status_code == 401:
        sys.exit(f"api_token is not valid on {urlparse(host).netloc} - check the token and cloud "
                 f"in {INI_NAME} (tokens only work on the cloud they were created on).")
    if resp.status_code in (403, 404):
        sys.exit(f"api_token has no access to org {ORG_ID} on {urlparse(host).netloc} - "
                 f"check org_id in {INI_NAME}.")
    resp.raise_for_status()
    return resp.json()


# --------------------------------------------------------------------------
# Port name handling
# --------------------------------------------------------------------------

def expand_ports(spec):
    """Expand a Mist port_config key into individual interface names.

    Handles "ge-0/0/0", "ge-0/0/0-15", "ge-0/0/[0-15]" and comma-separated
    lists of any of those, e.g. "ge-0/0/0-15, ge-1/0/0-10".
    """
    ports = []
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        m = re.match(r"^(.*/)\[?(\d+)(?:-(\d+))?\]?$", part)
        if not m:
            ports.append(part)
            continue
        prefix, start, end = m.group(1), int(m.group(2)), m.group(3)
        if end is None:
            ports.append(f"{prefix}{start}")
        else:
            ports.extend(f"{prefix}{i}" for i in range(start, int(end) + 1))
    return ports


def port_sort_key(port):
    m = PHYSICAL_PORT.match(port)
    if not m:
        return (9, port, 0, 0, 0)
    order = {"ge": 0, "mge": 1, "xe": 2, "et": 3}[m.group(1)]
    return (int(m.group(2)), int(m.group(3)), int(m.group(4)), order, port)


# --------------------------------------------------------------------------
# Derivation
# --------------------------------------------------------------------------

def render_vars(value, var_map):
    """Substitute {{var}} placeholders from site/switch vars. Unresolved ones are left as-is."""
    if not isinstance(value, str) or "{{" not in value:
        return value
    return re.sub(r"\{\{\s*(\w+)\s*\}\}", lambda m: str(var_map.get(m.group(1), m.group(0))), value)


def rule_matches(rule, device):
    """A switch matching rule applies when every match_* condition holds; a rule
    with no conditions matches everything. match_name/match_model compare the
    rule value against that field starting at match_<x>_offset."""
    for key, value in rule.items():
        if not key.startswith("match_") or key.endswith("_offset") or value in (None, ""):
            continue
        field = key[len("match_"):]
        actual = str(device.get(field) or "")
        if field == "role":
            if actual.lower() != str(value).lower():
                return False
        else:
            offset = int(rule.get(f"{key}_offset") or 0)
            if actual[offset:offset + len(str(value))].lower() != str(value).lower():
                return False
    return True


def find_rule(rules, device):
    for rule in rules:
        if rule_matches(rule, device):
            return rule
    return None


class SiteContext:
    """Everything needed to derive port config for switches at one site."""

    def __init__(self, site, setting, derived, templates):
        self.site = site
        self.derived = derived
        self.setting = setting
        tmpl = templates.get(derived.get("networktemplate_id") or site.get("networktemplate_id"))
        self.template = tmpl or {}
        self.template_name = derived.get("networktemplate_name") or self.template.get("name") or ""
        self.usages = dict(derived.get("port_usages") or {})
        self.networks = dict(derived.get("networks") or {})
        self.vars = dict(derived.get("vars") or {})
        sm = derived.get("switch_matching") or {}
        self.rules = (sm.get("rules") or []) if sm.get("enable", True) else []
        self.rules_source = "Site" if (setting.get("switch_matching") or {}).get("rules") else "Template"
        self.site_usage_names = set((setting.get("port_usages") or {}).keys())
        self.template_usage_names = set((self.template.get("port_usages") or {}).keys())
        self.use_usage_description = derived.get("uses_description_from_port_usage")


def usage_source(name, ctx, device):
    if name in (device.get("port_usages") or {}):
        return "Switch"
    if name in ctx.site_usage_names:
        return "Site" if name not in ctx.template_usage_names else "Site (overrides template)"
    if name in ctx.template_usage_names:
        return "Template"
    if name in SYSTEM_USAGES:
        return "System"
    return "NOT DEFINED"


def derive_switch(ctx, device, stats_ports):
    """Return one dict per physical port with the resolved config for that port."""
    usages = {**SYSTEM_USAGES, **ctx.usages, **(device.get("port_usages") or {})}
    networks = {**SYSTEM_NETWORKS, **ctx.networks, **(device.get("networks") or {})}
    var_map = {**ctx.vars, **(device.get("vars") or {})}

    rule = find_rule(ctx.rules, device)
    default_usage = (rule or {}).get("default_port_usage") or "default"

    # Per-port assignment, highest layer last so it wins. Within a layer Mist
    # processes the port_config keys in sorted order, so when a port appears in
    # more than one key the key that sorts last wins (confirmed against the
    # generated Junos) - the losing entries are reported as a warning.
    assigned = {}  # port -> (entry, source)
    conflicts = defaultdict(list)  # port -> profiles it was also given in the same layer
    for layer, source in (((rule or {}).get("port_config"), f"Rule port config ({ctx.rules_source})"),
                          (device.get("port_config"), "Switch port config")):
        layer_seen = {}
        for spec, entry in sorted((layer or {}).items()):
            for port in expand_ports(spec):
                if port in layer_seen and layer_seen[port] != entry.get("usage"):
                    conflicts[port].append(layer_seen[port])
                layer_seen[port] = entry.get("usage")
                assigned[port] = (dict(entry), source)
    for spec, entry in (device.get("local_port_config") or {}).items():
        for port in expand_ports(spec):
            base = assigned.get(port, ({}, ""))[0]
            assigned[port] = ({**base, **entry}, "Local port config")

    ports = set(p for p in stats_ports if PHYSICAL_PORT.match(p)) | set(p for p in assigned if PHYSICAL_PORT.match(p))

    def vlan(name, warnings):
        if not name:
            return "", ""
        net = networks.get(name)
        if net is None:
            warnings.append(f"Network {name} not defined")
            return name, "NOT DEFINED"
        vid = render_vars(str(net.get("vlan_id", "")), var_map)
        if str(vid).isdigit():
            return name, int(vid)
        if "{{" in str(vid):
            warnings.append(f"Network {name}: variable {vid} not defined (VLAN not created in Junos)")
        return name, vid

    rows = []
    for port in sorted(ports, key=port_sort_key):
        entry, source = assigned.get(port, ({}, ""))
        warnings = []
        if conflicts.get(port):
            warnings.append(f"Port in several port config entries - {', '.join(conflicts[port])} overridden")
        if entry.get("usage"):
            usage_name = entry["usage"]
        else:
            usage_name = default_usage
            source = f"Rule default ({ctx.rules_source})" if rule else "System default (no rule matched)"
        if usage_name not in usages:
            warnings.append(f"Port profile {usage_name} not defined")
        usage = dict(usages.get(usage_name) or {})
        overridden = [k for k in PORT_CONFIG_OVERRIDES if entry.get(k) not in (None, "")]
        for k in overridden:
            usage[k] = entry[k]

        mode = usage.get("mode") or ""
        port_net, port_vid = vlan(usage.get("port_network"), warnings)
        voice_net, voice_vid = vlan(usage.get("voip_network"), warnings)
        trunk_names, trunk_ids = "", ""
        if mode == "trunk":
            if usage.get("all_networks"):
                names = sorted(n for n in networks if n != "default")
                trunk_names = "ALL"
            else:
                names = list(usage.get("networks") or [])
                trunk_names = ", ".join(names)
            # An undefined network in "all networks" is reported once at the site, not per port.
            ids = [vlan(n, warnings if not usage.get("all_networks") else [])[1] for n in names]
            trunk_ids = ", ".join(str(i) for i in sorted(ids, key=lambda x: (not isinstance(x, int), x)))

        description = entry.get("description") or ""
        if not description and ctx.use_usage_description:
            description = usage.get("description") or ""
        storm = usage.get("storm_control") or {}
        storm_txt = f"{storm['percentage']}%" if storm.get("percentage") else ""

        lag = ""
        if entry.get("aggregated"):
            lag = f"ae{int(entry['ae_idx'])}" if entry.get("ae_idx") is not None else "ae (auto)"

        auth = usage.get("port_auth") or ""
        if usage.get("enable_mac_auth"):
            auth = (auth + " + " if auth else "") + ("MAC auth only" if usage.get("mac_auth_only") else "MAC auth")

        rows.append({
            "port": port,
            "vc_member": int(PHYSICAL_PORT.match(port).group(2)),
            "usage": usage_name,
            "usage_source": usage_source(usage_name, ctx, device),
            "assign_source": source,
            "rule": (rule or {}).get("name", ""),
            "overrides": ", ".join(overridden),
            "dynamic_usage": entry.get("dynamic_usage") or "",
            "mode": mode,
            "disabled": bool(usage.get("disabled")),
            "port_net": port_net, "port_vid": port_vid,
            "trunk_names": trunk_names, "trunk_ids": trunk_ids,
            "voice_net": voice_net, "voice_vid": voice_vid,
            "auth": auth,
            "guest_net": usage.get("guest_network") or "",
            "fail_net": usage.get("server_fail_network") or "",
            "reject_net": usage.get("server_reject_network") or "",
            "poe_disabled": bool(usage.get("poe_disabled")),
            "stp_edge": bool(usage.get("stp_edge")),
            "stp_disable": bool(usage.get("stp_disable")),
            "stp_required": bool(usage.get("stp_required")),
            "speed": usage.get("speed") or "",
            "duplex": usage.get("duplex") or "",
            "autoneg_off": bool(usage.get("disable_autoneg")),
            "mtu": usage.get("mtu") or "",
            "mac_limit": usage.get("mac_limit") or "",
            "storm": storm_txt,
            "qos": bool(usage.get("enable_qos")),
            "lag": lag,
            "esilag": bool(entry.get("esilag")),
            "critical": bool(entry.get("critical")),
            "description": description,
            "warnings": "; ".join(dict.fromkeys(warnings)),
        })
    return rows, rule


# --------------------------------------------------------------------------
# Junos check
# --------------------------------------------------------------------------

def expand_junos_member(member):
    """'ge-0/0/[16-23]' / 'ge-0/0/5' -> list of interfaces."""
    m = re.match(r"^(.*/)\[(\d+)-(\d+)\]$", member)
    if m:
        return [f"{m.group(1)}{i}" for i in range(int(m.group(2)), int(m.group(3)) + 1)]
    return [member]


def parse_junos_usages(cli):
    """Map each physical interface to the port profile Mist applied in the generated Junos.

    Profiles are applied as interface-range <profile> (with apply-groups <profile>),
    or directly as interfaces <if> apply-groups <profile>. LAG members inherit the
    profile applied to their aeN.
    """
    port_usage, ae_usage, ae_members = {}, {}, {}
    for line in cli:
        m = re.match(r"^set interfaces interface-range (\S+) member (\S+)$", line)
        if m:
            for p in expand_junos_member(m.group(2)):
                port_usage[p] = m.group(1)
            continue
        m = re.match(r"^set interfaces (\S+) apply-groups (\S+)$", line)
        if m and m.group(1) != "interface-range":
            (ae_usage if m.group(1).startswith("ae") else port_usage)[m.group(1)] = m.group(2)
            continue
        m = re.match(r"^set interfaces (\S+) ether-options 802\.3ad (ae\d+)$", line)
        if m:
            ae_members[m.group(1)] = m.group(2)
    for port, ae in ae_members.items():
        port_usage[port] = ae_usage.get(ae, port_usage.get(port, ""))
    return port_usage, ae_members


def pick_check_set(switches):
    if JUNOS_CHECK in ("0", "off", "no", "false", "none"):
        return set()
    if JUNOS_CHECK == "all":
        return set(range(len(switches)))
    try:
        n = max(0, int(JUNOS_CHECK))
    except ValueError:
        sys.exit(f"[validation] junos_check must be 'all', 'off' or a number - got '{JUNOS_CHECK}'.")
    if n >= len(switches):
        return set(range(len(switches)))
    step = len(switches) / n if n else 0
    return {int(i * step) for i in range(n)}


# --------------------------------------------------------------------------
# Excel
# --------------------------------------------------------------------------

THIN = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
RED_FILL = PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid")
AMBER_FILL = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")


def style_header(ws, headers, fill_color):
    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
        cell.border = THIN
    ws.freeze_panes = "A2"


def write_table(ws, headers, rows, fill_color, highlight=None):
    """Write a header + rows, apply an autofilter over the whole table and size columns."""
    style_header(ws, headers, fill_color)
    for r, row in enumerate(rows, 2):
        for c, value in enumerate(row, 1):
            if isinstance(value, bool):
                value = "Yes" if value else ""
            ws.cell(row=r, column=c, value=value)
        if highlight:
            fill = highlight(row)
            if fill:
                for c in range(1, len(headers) + 1):
                    ws.cell(row=r, column=c).fill = fill
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(len(rows) + 1, 1)}"
    auto_width(ws)


def auto_width(ws):
    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            if cell.value is not None:
                max_len = max(max_len, len(str(cell.value)))
        # +2 leaves room for the autofilter drop-down arrow on the header.
        ws.column_dimensions[col_letter].width = min(max_len + 4, 60)


NOTES = [
    ("Mist Switchport Config", None),
    ("", None),
    ("What this is", True),
    ("One row per physical switch port, showing the configuration Mist derives for it after all", None),
    ("template, site and switch-level overrides are applied.", None),
    ("", None),
    ("How each port's config is derived (each layer overrides the one above it)", True),
    ("1. Switch Template - port profiles, networks (VLANs) and switch matching rules.", None),
    ("2. Site settings - port profiles and networks merged by name; site rules replace template rules.", None),
    ("   Layers 1+2 come from Mist's own merge (GET /sites/{id}/setting/derived).", None),
    ("3. Switch matching rule - the first rule whose match_* conditions fit the switch (role, name,", None),
    ("   model ...). Its default port profile applies to every port; its port config assigns specific ports.", None),
    ("4. Switch config - the switch's own port_config, plus any switch-level port profiles, networks and vars.", None),
    ("5. Local port config - changes made locally on the switch, where permitted.", None),
    ("VLAN IDs written as {{variable}} are resolved from the site (and switch) vars.", None),
    ("", None),
    ("Column guide (Switchport Config sheet)", True),
    ("Port Profile Assigned By - which layer chose the port profile for this port.", None),
    ("Port Profile Defined In - Template / Site / Switch / System (Mist built-in), or NOT DEFINED.", None),
    ("Port-Level Overrides - profile settings overridden directly on the port (e.g. speed, duplex).", None),
    ("Config Warnings - undefined port profiles / networks, {{variables}} with no value (Mist then leaves", None),
    ("   that VLAN out of the Junos config), or a port listed in more than one port config entry. In that", None),
    ("   last case Mist applies the entry whose port list sorts last alphabetically, not the one listed last.", None),
    ("Live ... columns - from the switch's port stats. 'Switch offline' = no current stats for the switch;", None),
    ("   'Port not reported by switch' usually means the port is configured but doesn't physically exist.", None),
    ("Junos Profile / Junos Check - the profile Mist actually applied in the Junos config it generated", None),
    ("   for that switch (GET /devices/{id}/config_cmd). MISMATCH rows are highlighted red.", None),
    ("Live Usage Check - compares the derived profile with the profile the switch reports in its stats.", None),
    ("", None),
    ("Other sheets", True),
    ("Switches - one row per switch: matched rule, port counts, override counts and check results.", None),
    ("Issues - every port with a Junos mismatch, a live profile mismatch or a config warning.", None),
    ("Highlighting - red = differs from generated Junos; amber = differs from live stats or has a warning.", None),
]


def write_notes(ws, meta):
    for r, (text, bold) in enumerate(NOTES, 1):
        cell = ws.cell(row=r, column=1, value=text)
        if r == 1:
            cell.font = Font(bold=True, size=14)
        elif bold:
            cell.font = Font(bold=True)
    r = len(NOTES) + 2
    ws.cell(row=r, column=1, value="Run details").font = Font(bold=True)
    for i, (k, v) in enumerate(meta, r + 1):
        ws.cell(row=i, column=1, value=f"{k}: {v}")
    ws.column_dimensions["A"].width = 110


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------

def main():
    global API_BASE, API_HOST

    if API_TOKEN in ("", "YOUR_API_TOKEN_HERE") or ORG_ID in ("", "YOUR_ORG_ID_HERE") or not CLOUD:
        print(f"Set org_id, api_token and cloud in {INI_NAME} before running.")
        return

    host = cloud_to_host(CLOUD)
    print(f"Connecting to {urlparse(host).netloc}...")
    org_info = fetch_org_info(host)
    API_HOST = host
    API_BASE = f"{host}/api/v1"
    org_name = org_info.get("name", "Unknown")
    start_time = time.time()

    print(f"\n{'='*60}")
    print(f"  Mist Switchport Config - {org_name}")
    print(f"  Org ID: {ORG_ID}   Cloud: {urlparse(API_HOST).netloc}")
    print(f"{'='*60}")

    print("Fetching sites, switch templates and switch status...")
    sites = api_get_pages(f"/orgs/{ORG_ID}/sites", "sites")
    templates = {t["id"]: t for t in api_get_pages(f"/orgs/{ORG_ID}/networktemplates", "templates")}
    status = {d.get("mac"): d for d in api_get_pages(f"/orgs/{ORG_ID}/stats/devices?type=switch", "device stats")}
    print(f"  {len(sites)} sites, {len(templates)} switch templates, {len(status)} switches with stats")

    print("Fetching switch port stats (org-wide)...")
    port_stats = defaultdict(dict)
    for p in api_search(f"/orgs/{ORG_ID}/stats/ports/search?type=switch&limit=1000", "port stats"):
        port_stats[p.get("mac")][p.get("port_id")] = p

    print("Fetching per-site settings and switch configs...")
    switches = []  # (ctx, device)
    t0 = time.time()
    for i, site in enumerate(sorted(sites, key=lambda s: s.get("name", "")), 1):
        devices = api_get(f"/sites/{site['id']}/devices?type=switch&limit=1000", "site devices")
        if devices:
            setting = api_get(f"/sites/{site['id']}/setting", "site settings")
            derived = api_get(f"/sites/{site['id']}/setting/derived", "site settings")
            ctx = SiteContext(site, setting, derived, templates)
            switches.extend((ctx, d) for d in sorted(devices, key=lambda d: d.get("name") or d.get("mac")))
        elapsed = time.time() - t0
        eta = elapsed / i * (len(sites) - i)
        sys.stdout.write(f"\r  Site {i}/{len(sites)} | {len(switches)} switches | ETA: {format_eta(eta)}   ")
        sys.stdout.flush()
    print()

    check_set = pick_check_set(switches)
    if check_set:
        print(f"Checking {len(check_set)} switches against Mist's generated Junos config...")

    port_rows, switch_rows, mismatch_rows = [], [], []
    t0 = time.time()
    checked = 0
    for idx, (ctx, dev) in enumerate(switches):
        mac = dev.get("mac", "")
        st = status.get(mac, {})
        live = port_stats.get(mac, {})
        rows, rule = derive_switch(ctx, dev, live.keys())

        junos, junos_ae = None, {}
        if idx in check_set:
            data = api_get(f"/sites/{ctx.site['id']}/devices/{dev['id']}/config_cmd", "junos config")
            junos, junos_ae = parse_junos_usages(data.get("cli") or [])
            checked += 1
            elapsed = time.time() - t0
            sys.stdout.write(f"\r  Switch {checked}/{len(check_set)} | ETA: "
                             f"{format_eta(elapsed / checked * (len(check_set) - checked))}   ")
            sys.stdout.flush()

        counts = Counter()
        for row in rows:
            lp = live.get(row["port"], {})
            live_usage = lp.get("port_usage") or ""
            if not lp:
                live_check = ("Switch offline" if st.get("status") != "connected"
                              else "Port not reported by switch")
            elif not live_usage:
                live_check = "Not reported"
            else:
                live_check = "OK" if live_usage == row["usage"] else "MISMATCH"
            if junos is None:
                junos_usage, junos_check = "", "Not checked"
            else:
                junos_usage = junos.get(row["port"], "")
                if not junos_usage:
                    junos_check = "Not in Junos"
                else:
                    junos_check = "OK" if junos_usage == row["usage"] else "MISMATCH"
                if row["lag"] and junos_ae.get(row["port"]) and row["lag"] != junos_ae[row["port"]]:
                    junos_check = "MISMATCH (LAG)"
            counts[row["assign_source"]] += 1
            counts["overrides"] += bool(row["overrides"])
            counts["junos_mismatch"] += junos_check.startswith("MISMATCH")
            counts["live_mismatch"] += live_check == "MISMATCH"
            counts["up"] += bool(lp.get("up"))

            out = [
                ctx.site.get("name", ""), ctx.template_name, dev.get("name") or "", mac, dev.get("model", ""),
                dev.get("role") or "", st.get("status", "unknown"),
                row["vc_member"], row["port"], row["usage"], row["assign_source"], row["usage_source"],
                row["rule"], row["overrides"], row["dynamic_usage"], row["mode"], row["disabled"],
                row["port_net"], row["port_vid"], row["trunk_names"], row["trunk_ids"],
                row["voice_net"], row["voice_vid"], row["auth"], row["guest_net"], row["fail_net"],
                row["reject_net"], row["poe_disabled"], row["stp_edge"], row["stp_disable"],
                row["stp_required"], row["speed"], row["duplex"], row["autoneg_off"], row["mtu"],
                row["mac_limit"], row["storm"], row["qos"], row["lag"], row["esilag"], row["critical"],
                row["description"], row["warnings"],
                ("Up" if lp.get("up") else "Down") if lp else "", lp.get("speed", "") if lp else "",
                lp.get("port_mode", "") if lp else "", live_usage,
                lp.get("neighbor_system_name", "") if lp else "", lp.get("neighbor_port_desc", "") if lp else "",
                ("Yes" if lp.get("poe_on") else "") if lp else "", lp.get("power_draw", "") if lp else "",
                live_check, junos_usage, junos_check,
            ]
            port_rows.append(out)
            counts["warnings"] += bool(row["warnings"])
            issues = []
            if junos_check.startswith("MISMATCH"):
                issues.append(f"Junos {junos_check}")
            if live_check == "MISMATCH":
                issues.append("Live profile mismatch")
            if row["warnings"]:
                issues.append("Config warning")
            if issues:
                mismatch_rows.append([ctx.site.get("name", ""), dev.get("name") or "", mac, row["port"],
                                      ", ".join(issues), row["usage"], row["assign_source"], row["warnings"],
                                      live_usage, junos_usage])

        switch_rows.append([
            ctx.site.get("name", ""), ctx.template_name, dev.get("name") or "", mac, dev.get("model", ""),
            dev.get("role") or "", st.get("status", "unknown"),
            len(dev.get("virtual_chassis", {}).get("members") or []) or 1,
            (rule or {}).get("name", "NO RULE MATCHED"), ctx.rules_source if rule else "",
            (rule or {}).get("default_port_usage") or "default",
            len(rows), counts["up"],
            sum(v for k, v in counts.items() if k.startswith("Rule default") or k.startswith("System default")),
            sum(v for k, v in counts.items() if k.startswith("Rule port config")),
            counts["Switch port config"], counts["Local port config"], counts["overrides"],
            len(dev.get("port_usages") or {}), len(dev.get("networks") or {}),
            "Yes" if junos is not None else "", counts["junos_mismatch"] if junos is not None else "",
            counts["live_mismatch"], counts["warnings"],
        ])
    if check_set:
        print()

    print("Building Excel spreadsheet...")
    wb = Workbook()
    ws_notes = wb.active
    ws_notes.title = "Notes"

    port_headers = [
        "Site", "Switch Template", "Switch Name", "Switch MAC", "Model", "Role", "Switch Status",
        "VC Member", "Port", "Port Profile", "Port Profile Assigned By", "Port Profile Defined In",
        "Matched Rule", "Port-Level Overrides", "Dynamic Profile", "Mode", "Port Disabled",
        "Port/Native VLAN", "Port/Native VLAN ID", "Trunk VLANs", "Trunk VLAN IDs",
        "Voice VLAN", "Voice VLAN ID", "Port Auth", "Guest VLAN", "Server-Fail VLAN",
        "Server-Reject VLAN", "PoE Disabled", "STP Edge", "STP Disabled", "STP Required",
        "Speed", "Duplex", "Autoneg Disabled", "MTU", "MAC Limit", "Storm Control", "QoS",
        "LAG", "ESI-LAG", "Critical", "Description", "Config Warnings",
        "Live Link", "Live Speed (Mbps)", "Live Mode", "Live Port Profile", "LLDP Neighbor",
        "LLDP Neighbor Port", "Live PoE On", "Live PoE Draw (W)", "Live Usage Check",
        "Junos Profile", "Junos Check",
    ]
    live_col, junos_col = port_headers.index("Live Usage Check"), port_headers.index("Junos Check")
    warn_col = port_headers.index("Config Warnings")
    ws = wb.create_sheet("Switchport Config")
    write_table(ws, port_headers, port_rows, "1F4E78",
                highlight=lambda r: RED_FILL if str(r[junos_col]).startswith("MISMATCH")
                else AMBER_FILL if r[live_col] == "MISMATCH" or r[warn_col] else None)

    ws = wb.create_sheet("Switches")
    write_table(ws, [
        "Site", "Switch Template", "Switch Name", "Switch MAC", "Model", "Role", "Switch Status",
        "VC Members", "Matched Rule", "Rule Source", "Rule Default Profile", "Ports", "Ports Up",
        "Ports On Rule Default", "Ports Set By Rule", "Ports Set On Switch", "Ports Set Locally",
        "Ports With Overrides", "Switch-Level Profiles", "Switch-Level Networks",
        "Junos Checked", "Junos Mismatches", "Live Mismatches", "Ports With Warnings",
    ], switch_rows, "375623",
        highlight=lambda r: RED_FILL if r[21] else AMBER_FILL if r[22] or r[23] else None)

    ws = wb.create_sheet("Issues")
    write_table(ws, ["Site", "Switch Name", "Switch MAC", "Port", "Issue", "Derived Profile", "Assigned By",
                     "Config Warnings", "Live Port Profile", "Junos Profile"],
                mismatch_rows, "C00000",
                highlight=lambda r: RED_FILL if "Junos" in r[4] else None)

    total_api = sum(API_CALLS.values())
    junos_mm = sum(1 for r in port_rows if str(r[junos_col]).startswith("MISMATCH"))
    live_mm = sum(1 for r in port_rows if r[live_col] == "MISMATCH")
    warn_n = sum(1 for r in port_rows if r[warn_col])
    write_notes(ws_notes, [
        ("Organisation", org_name), ("Org ID", ORG_ID), ("Cloud", urlparse(API_HOST).netloc),
        ("Generated", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("Switches", len(switches)), ("Ports", len(port_rows)),
        ("Switches checked against Junos", checked), ("Junos mismatches", junos_mm),
        ("Live usage mismatches", live_mm), ("Ports with config warnings", warn_n),
    ])

    timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    safe_org_name = "".join(c if c.isalnum() or c in (" ", "-", "_") else "_" for c in org_name).strip().replace(" ", "_")
    filepath = os.path.join(OUTPUT_DIR, f"Mist_Switchport_Config_{safe_org_name}_{timestamp}.xlsx")
    wb.save(filepath)

    elapsed = time.time() - start_time
    print(f"\n{'='*60}")
    print(f"  Mist Switchport Config Summary - {org_name}")
    print(f"{'='*60}")
    print(f"  Sites with switches:        {len({ctx.site['id'] for ctx, _ in switches})}")
    print(f"  Switches:                   {len(switches)}")
    print(f"  Ports:                      {len(port_rows)}")
    print(f"  Switches checked vs Junos:  {checked}")
    print(f"  Junos mismatches:           {junos_mm}")
    print(f"  Live usage mismatches:      {live_mm}")
    print(f"  Ports with config warnings: {warn_n}")
    for label, n in sorted(API_CALLS.items()):
        print(f"  API calls - {label + ':':<16}{n}")
    print(f"  Total API calls:            {total_api}")
    print(f"  Total elapsed time:         {format_eta(elapsed)}")
    print(f"{'='*60}")
    print(f"  Report saved: {filepath}")


if __name__ == "__main__":
    main()
