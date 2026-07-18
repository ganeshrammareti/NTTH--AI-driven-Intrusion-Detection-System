# No Time To Hack (NTTH)

> An adaptive, AI-driven network security system that combines a honeypot, intrusion detection, and an automated firewall — all controllable from a Flutter mobile/desktop app in real time.

---

## What is NTTH?

NTTH is a final-year university project built to address a real problem: home and small-office networks are largely undefended. Most people rely on their router's default settings and hope for the best.

This project takes a different approach. NTTH sits on your Linux host, watches every packet that crosses your LAN, and responds to threats automatically — blocking attackers, luring them into a fake honeypot service, and sending you live alerts through a polished Flutter app on your phone or Windows PC.

The "AI" part isn't marketing fluff. The backend uses a scikit-learn anomaly detection model alongside a rule-based IDS engine. When a threat is detected, a pipeline of autonomous agents decides what to do: quarantine the device, tighten firewall rules, update the model's feedback, or just log and watch.

---

## Features

- **Live Packet Inspection** — Scapy-powered sniffer captures raw traffic on your LAN interface. Every packet is classified and stored.
- **Intrusion Detection System (IDS)** — Rule engine + ML anomaly model work together to flag port scans, brute-force attempts, deauthentication attacks, rogue APs, and more.
- **Automated Firewall** — nftables rules are written and flushed automatically. No manual iptables commands needed.
- **Honeypot Integration** — Cowrie SSH honeypot runs as a Docker sidecar. Attackers who probe open ports get redirected into a fake shell and logged.
- **Multi-Agent Decision Pipeline** — Five specialized agents handle threat triage, enforcement, feedback learning, reporting, and escalation independently.
- **GeoIP Threat Mapping** — Attacker IPs are geolocated using MaxMind's GeoLite2 database and displayed on an interactive map in the app.
- **Real-Time WebSocket Feed** — The Flutter app stays connected via a JWT-authenticated WebSocket and receives live threat events as they happen.
- **Cross-Platform Control App** — One Flutter codebase targets Android and Windows. Includes a dashboard, device list, firewall rule manager, honeypot session log, packet inspector, and network topology view.
- **Wi-Fi Monitoring (AR9271)** — Optional support for USB wireless adapters in monitor mode to detect rogue APs and deauth floods on 802.11 networks.

---

## Tech Stack

| Layer | Technology |
|---|---|
| Backend API | FastAPI + Uvicorn |
| Database | SQLite (dev) / PostgreSQL (prod) via SQLAlchemy + Alembic |
| Packet Capture | Scapy |
| Machine Learning | scikit-learn, NumPy |
| Firewall | nftables (via subprocess) |
| Honeypot | Cowrie (Docker container) |
| GeoIP | MaxMind GeoLite2 + geoip2 |
| Auth | JWT (python-jose) + bcrypt |
| Scheduler | APScheduler |
| Mobile/Desktop App | Flutter 3 (Dart) |
| State Management | Provider + GoRouter |
| Charts & Maps | fl_chart, flutter_map |
| Deployment | Docker + Docker Compose |

---

## Project Structure

```
NTTH-main/
├── backend/                  # FastAPI security engine
│   ├── app/
│   │   ├── agents/           # Decision, enforcement, feedback, threat, reporting agents
│   │   ├── api/              # REST API route handlers
│   │   ├── core/             # App lifecycle, startup, config
│   │   ├── firewall/         # nftables rule management
│   │   ├── honeypot/         # Cowrie integration & session parsing
│   │   ├── ids/              # Rule engine + anomaly model
│   │   ├── monitor/          # Scapy packet sniffer
│   │   ├── geoip/            # MaxMind GeoIP lookup
│   │   └── websocket/        # Live event broadcasting
│   ├── alembic/              # Database migrations
│   ├── requirements.txt
│   └── Dockerfile
├── flutter_app/              # Mobile & desktop control app
│   ├── lib/
│   │   ├── screens/          # Dashboard, devices, firewall, honeypot, threat map, etc.
│   │   ├── models/           # Data models
│   │   ├── core/             # API client, WebSocket service
│   │   ├── widgets/          # Reusable UI components
│   │   └── theme/            # App theme & colors
│   └── pubspec.yaml
├── docker-compose.yml        # One-command production deployment
└── Makefile                  # Convenience targets
```

---

## Getting Started

### Prerequisites

- **Linux host** — Ubuntu 22.04 or later is recommended (the firewall and packet capture components require a Linux kernel)
- **Docker & Docker Compose** — for production deployment
- **Python 3.11+** — only needed for running the backend without Docker
- **Flutter SDK 3.3+** — only needed if you want to build the app yourself
- **MaxMind GeoLite2** account — free, required for the threat map feature

> **Note:** The backend will not run on Windows or macOS because it uses Scapy for raw packet capture and nftables for firewall management. Run it on a Linux machine or VM.

---

### Quick Start (Development — no Docker)

```bash
# 1. Clone the repo
git clone https://github.com/your-username/NTTH.git
cd NTTH/backend

# 2. Install Python dependencies
pip install -r requirements.txt

# 3. Copy the example environment file and adjust if needed
cp .env.example .env

# 4. Start the backend (SQLite is used by default in dev mode)
uvicorn app.main:app --reload --port 8000
```

The API will be available at `http://localhost:8000`.  
Interactive API docs: `http://localhost:8000/docs`

---

### Production Deployment (Docker)

```bash
# 1. Clone the repo
git clone https://github.com/your-username/NTTH.git
cd NTTH/backend

# 2. Configure your environment
cp .env.example .env
# Open .env and set your NTTH_ADMIN_PASSWORD, NTTH_SECRET_KEY, and NETWORK_INTERFACE

# 3. (Optional) Download GeoLite2 database for the threat map
# Place GeoLite2-City.mmdb inside backend/geoip/

# 4. Start everything
docker compose up -d

# 5. Follow logs
docker compose logs -f
```

To also start the Cowrie SSH honeypot:

```bash
docker compose --profile full up -d
```

The backend API runs on port **8000**, and the HTTP honeypot listens on port **8888**. If you enable Cowrie, the SSH honeypot is on port **2222**.

---

### Running the Flutter App

```bash
cd flutter_app
flutter pub get

# Run on a connected Android device
flutter run

# Run as a Windows desktop app
flutter run -d windows
```

When the app launches, enter the IP address of the machine running the NTTH backend. Log in with the admin credentials you set in `.env` (default: `admin` / `changeme` — **please change this**).

---

## App Screens

| Screen | What it shows |
|---|---|
| Dashboard | Live threat count, recent events, system health at a glance |
| Devices | All devices discovered on the LAN with their status |
| Firewall | Active nftables rules, add/remove rules manually |
| Honeypot | Cowrie session log — commands attackers typed, credentials they tried |
| Threat Map | World map with geolocated attacker IPs |
| Packet Inspector | Raw packet stream with filtering |
| Network Topology | Visual graph of devices and connections |
| System Health | CPU, memory, uptime, service status |

---

## API Overview

The REST API is documented interactively at `/docs` when the backend is running. Key endpoints:

| Method | Endpoint | Description |
|---|---|---|
| `POST` | `/api/v1/auth/login` | Authenticate and get a JWT token pair |
| `GET` | `/api/v1/devices` | List all discovered LAN devices |
| `GET` | `/api/v1/threats` | Paginated threat event log |
| `GET` | `/api/v1/firewall/rules` | Currently active firewall rules |
| `POST` | `/api/v1/firewall/rules` | Add a new firewall rule |
| `POST` | `/api/v1/firewall/flush` | Emergency flush — remove all rules (admin only) |
| `GET` | `/api/v1/honeypot/sessions` | Honeypot session log |
| `GET` | `/api/v1/system/health` | Health check (no auth required) |
| `WS` | `/ws/live?token=<jwt>` | WebSocket stream for live threat events |

---

## How the AI Pipeline Works

When the packet sniffer flags suspicious traffic, it goes through this chain:

```
Packet Sniffer (Scapy)
        ↓
   IDS Rule Engine          ← checks for known attack patterns (port scan, brute force, etc.)
        ↓
  Anomaly Model             ← scikit-learn model scores the traffic for abnormality
        ↓
  Threat Agent              ← assigns severity, enriches with GeoIP data
        ↓
  Decision Agent            ← decides: block, quarantine, honeypot redirect, or monitor
        ↓
  Enforcement Agent         ← writes the nftables rule or updates Cowrie redirect
        ↓
  Feedback Agent            ← updates the anomaly model with labelled data
        ↓
  Reporting Agent           ← stores the event, broadcasts it over WebSocket
```

Each agent runs independently and can be tuned or replaced without touching the others.

---

## Configuration

All configuration is done through environment variables. Copy `backend/.env.example` to `backend/.env` and edit:

| Variable | Default | Description |
|---|---|---|
| `ENVIRONMENT` | `development` | `development` or `production` |
| `DATABASE_URL` | SQLite | PostgreSQL URL for production |
| `SECRET_KEY` | *(must set)* | JWT signing key |
| `ADMIN_USERNAME` | `admin` | Default admin username |
| `ADMIN_PASSWORD` | `changeme` | Default admin password — **change this** |
| `NETWORK_INTERFACE` | auto-detect | LAN interface to monitor (e.g. `eth0`) |
| `WIFI_ENABLED` | `false` | Enable Wi-Fi monitor mode |
| `WIFI_INTERFACE` | auto-detect | Wireless interface for monitor mode |
| `DEAUTH_THRESHOLD` | `10` | Deauth packets per second to trigger alert |
| `ROGUE_AP_DETECTION` | `true` | Alert on unknown APs |
| `AP_WHITELIST_SSIDS` | *(empty)* | Comma-separated trusted SSIDs |

---

## Known Limitations

- The backend requires a real Linux network interface. Running it inside WSL2 with bridged networking is possible but not officially tested.
- The anomaly detection model is pre-trained on synthetic traffic. It will need retraining on real network data for best accuracy in your environment.
- Cowrie honeypot is optional — the rest of NTTH works without it.
- Wi-Fi monitor mode has only been tested with the AR9271 USB adapter (Atheros chipset). Other adapters that support monitor mode should work but are untested.

---

## Contributing

This was built as a final-year university project, but contributions and suggestions are welcome. Open an issue if you find a bug, or submit a pull request with a clear description of what you changed and why.

---

## License

This project is released under the MIT License. See `LICENSE` for details.

---

## Acknowledgements

- [Cowrie](https://github.com/cowrie/cowrie) — the SSH/Telnet honeypot this project integrates with
- [Scapy](https://scapy.net/) — packet crafting and capture library
- [MaxMind GeoLite2](https://dev.maxmind.com/geoip/geolite2-free-geolocation-data/) — free GeoIP database
- [FastAPI](https://fastapi.tiangolo.com/) — modern Python web framework
- [Flutter](https://flutter.dev/) — cross-platform UI toolkit
