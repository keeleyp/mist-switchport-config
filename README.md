# Mist Switchport Config

Exports the **derived** configuration of every switch port, on every switch, at
every site in a Juniper Mist organisation to Excel. The result is what each port
actually ends up with after all template, site and switch-level overrides are
applied.

## How the config is derived

Mist builds each port's config from layers. Each layer overrides the one above
it:

| # | Layer | What it contributes |
|---|---|---|
| 1 | Switch Template | Port profiles, networks (VLANs) and switch matching rules |
| 2 | Site settings | Port profiles and networks merged by name; site rules replace template rules |
| 3 | Switch matching rule | First rule whose `match_*` conditions fit the switch (role, name, model). Its default port profile applies to every port, and its port config assigns specific ports |
| 4 | Switch config | The switch's own `port_config`, plus any switch-level port profiles, networks and vars |
| 5 | Local port config | Changes made locally on the switch, where permitted |

Layers 1 and 2 are taken from Mist's own merge (`GET /sites/{id}/setting/derived`).
The script resolves the rest itself:

- `{{variable}}` VLAN IDs are filled in from the site and switch vars.
- Port ranges such as `ge-0/0/0-15, ge-1/0/0-10` are expanded per Virtual
  Chassis member.
- Built-in profiles are understood: `default`, `disabled`, `ap`, `uplink`,
  `iot`, `inet` and `evpn_*`.
- Ports are limited to those that exist on the hardware. The script uses Mist's
  device model catalogue (`/const/device_models`) and each switch's Virtual
  Chassis members, so it works for offline switches too. Blanket template ranges
  such as `mge-0/0/0-47 ... mge-7/0/0-47` don't create phantom rows. Mist still
  pushes config for those ports, and the Switches sheet counts them.
- Dynamic port profiles are understood. Mist leaves those ports out of the
  static Junos config, and the switch assigns them a profile at runtime (e.g.
  `access-point` when LLDP sees a Mist AP).
- Switch matching rules support `match_role`, slice matches such as
  `match_name[9:17]` / `match_model[0:6]`, and the older `match_name` +
  `match_name_offset` form.
- When the same port appears in more than one `port_config` entry, Mist applies
  the entry whose key **sorts last alphabetically**, not the one listed last.
  This was confirmed against the generated Junos, and these ports are flagged.

### Validation

Each port's derived profile is checked two ways:

- **Junos check:** compared against the Junos config Mist actually generates for
  the switch (`GET /sites/{id}/devices/{id}/config_cmd`). This costs one API call
  per switch and is controlled by `[validation] junos_check` (`all`, `off`, or a
  number of switches to sample).
- **Live check:** compared against the profile the switch reports in its port
  stats.

## Setup

```bash
pip install -r requirements.txt
cp mist_switchport_config.ini.example mist_switchport_config.ini
# edit org_id, api_token and cloud
python3 mist_switchport_config.py
```

`cloud` is required. Use the API host for the portal you log in to, e.g.
`api.eu.mist.com` for `manage.eu.mist.com`. Short forms such as `eu` / `gc3`, and
the portal URL itself, also work.

## Running Behind a TLS-Inspecting Proxy (e.g. Zscaler)

If your network intercepts and re-signs HTTPS traffic, TLS verification against
the Mist API will fail with a certificate error. Fill in the `[network]`
section of `mist_switchport_config.ini`:

```ini
[network]
ca_bundle = /path/to/zscaler-root-ca.pem
verify_ssl = true
http_proxy =
https_proxy =
```

- `ca_bundle`: PEM file with your proxy's root CA certificate, or a bundle that
  includes it. Your IT team or the Zscaler client app can usually export it.
- `verify_ssl = false`: disables certificate verification entirely. Only use it
  if you truly can't get the proxy's CA certificate.
- `http_proxy` / `https_proxy`: only needed if the standard `HTTP_PROXY` /
  `HTTPS_PROXY` / `NO_PROXY` environment variables aren't already set.

With no `[network]` section (or blank values), the tool behaves exactly as it
does on a normal, non-intercepted connection.

## Output

Each sheet has a frozen header, an autofilter over the whole table, and
auto-sized columns.

- **Notes:** methodology, column guide and run details.
- **Switchport Config:** one row per physical port, with these columns:
  - Switch details: site, template, switch, model, role, status, VC member and port.
  - How the profile was chosen: port profile, which layer assigned it, where
    it's defined, matched rule and port-level overrides.
  - Resolved settings: mode, native/port VLAN name and ID, trunk VLAN names and
    IDs, voice VLAN, port auth, guest/fail/reject VLANs, PoE, STP, speed/duplex,
    MTU, MAC limit, storm control, QoS, LAG/ESI-LAG, critical and description.
  - Config warnings.
  - Live link/speed/mode/profile, LLDP neighbour and PoE.
  - Live and Junos checks.
- **Switches:** one row per switch, with the matched rule, port counts per layer,
  override counts and check results.
- **Issues:** every port with a Junos mismatch, a live-profile mismatch or a
  config warning.

Highlighting: red means the port differs from the generated Junos. Amber means it
differs from live stats or has a config warning.

### Config warnings

- **Port profile `X` not defined:** a port is assigned a profile that doesn't
  exist, so Mist leaves the port out of the Junos config.
- **Network `X`: variable `{{y}}` not defined:** the VLAN ID points at a site
  variable with no value, so Mist leaves that VLAN out of the Junos config.
- **Port in several port config entries:** the port is listed twice with
  different profiles. The overridden profile is named.
