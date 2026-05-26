# Cisco MDS 9000 Health Check Script

Automated health check tool for Cisco MDS 9000 series SAN switches. Connects via SSH, runs 34 diagnostic CLI commands, analyzes the output across 20 health check modules, and generates color-coded terminal, JSON, or Excel reports.

## Features

| Module | What it checks |
|---|---|
| NX-OS Version | Running release identification |
| Uptime | Recent reboots (< 1 day warning) |
| Module Status | Linecard / supervisor failures |
| Environment | PSU, fans, temperature (>65°C alert) |
| CPU & Memory | Utilization thresholds (>60% warn, >80% critical) |
| Interface Summary | Up/down FC port counts |
| Interface Errors | CRC, invalid CRC, invalid tx words, FEC corrected/uncorrected, non-corrected errors, link/signal/sync loss, credit loss, delimiter errors (17 counters with warn/crit thresholds) |
| Transceivers | Out-of-range optics (temp, power, voltage) |
| FLOGI | Logged-in device count |
| FCNS/FLOGI Consistency | Stale FCNS entries, unregistered FLOGI logins |
| VSAN | Suspended VSAN detection |
| Zoning | Active zoneset, merge failure detection |
| Smart Zoning | Per-VSAN smart zoning status |
| ACL TCAM | TCAM/SOBP resource utilization per linecard |
| Port-Channels | Oper status, partially bundled members, speed mismatches, consistency |
| ISL Load Balance | Per-port-channel member traffic skew, ISL utilization >80% |
| Linecard Balance | Traffic distribution across modules |
| Device-Alias | Distribution mode, pending sessions, devices without aliases |
| Clock / NTP | NTP server configuration |
| Syslog | Critical/error messages in recent logs |
| **Fabric Comparison** | Cross-switch zoneset/zone/member diff + device-alias diff |

## Installation

```bash
pip install -r requirements.txt
```

## Usage

```bash
# Single switch
python cisco_mds_health_check.py -t 10.1.1.1 -u admin

# Multiple switches from file
python cisco_mds_health_check.py -f switches.txt -u admin

# JSON output
python cisco_mds_health_check.py -t 10.1.1.1 --json -o report.json

# Excel report
python cisco_mds_health_check.py -t 10.1.1.1 --xlsx health_report.xlsx

# Compare zoning between two fabrics
python cisco_mds_health_check.py --fabric-a 10.1.1.1 --fabric-b 10.2.2.1 -u admin
```

### switches.txt format

```
# One IP per line, comments with #
10.1.1.1
10.1.1.2
10.1.1.3
```

## Output Formats

- **Terminal** — color-coded severity (CRITICAL/WARNING/INFO/OK), grouped by severity
- **JSON** (`--json`) — machine-readable, one object per switch
- **Excel** (`--xlsx`) — Summary sheet + per-switch detail sheet with color-coded cells

## Requirements

- Python 3.6+
- SSH access to MDS switches (uses `cisco_nxos` device type via Netmiko)
- Switch user needs privilege to run `show` commands

## License

See [LICENSE](LICENSE) file.
