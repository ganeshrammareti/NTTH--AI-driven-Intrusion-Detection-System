No Time To Hack (NTTH)
An adaptive, AI-driven network security system that combines a honeypot, intrusion detection, and an automated firewall — all controllable from a Flutter mobile/desktop app in real time.

What is NTTH?
NTTH is a final-year university project built to address a real problem: home and small-office networks are largely undefended. Most people rely on their router's default settings and hope for the best.

This project takes a different approach. NTTH sits on your Linux host, watches every packet that crosses your LAN, and responds to threats automatically — blocking attackers, luring them into a fake honeypot service, and sending you live alerts through a polished Flutter app on your phone or Windows PC.

The "AI" part isn't marketing fluff. The backend uses a scikit-learn anomaly detection model alongside a rule-based IDS engine. When a threat is detected, a pipeline of autonomous agents decides what to do: quarantine the device, tighten firewall rules, update the model's feedback, or just log and watch.

Features
Live Packet Inspection — Scapy-powered sniffer captures raw traffic on your LAN interface. Every packet is classified and stored.
Intrusion Detection System (IDS) — Rule engine + ML anomaly model work together to flag port scans, brute-force attempts, deauthentication attacks, rogue APs, and more.
Automated Firewall — nftables rules are written and flushed automatically. No manual iptables commands needed.
Honeypot Integration — Cowrie SSH honeypot runs as a Docker sidecar. Attackers who probe open ports get redirected into a fake shell and logged.
Multi-Agent Decision Pipeline — Five specialized agents handle threat triage, enforcement, feedback learning, reporting, and escalation independently.
GeoIP Threat Mapping — Attacker IPs are geolocated using MaxMind's GeoLite2 database and displayed on an interactive map in the app.
Real-Time WebSocket Feed — The Flutter app stays connected via a JWT-authenticated WebSocket and receives live threat events as they happen.
Cross-Platform Control App — One Flutter codebase targets Android and Windows. Includes a dashboard, device list, firewall rule manager, honeypot session log, packet inspector, and network topology view.
Wi-Fi Monitoring (AR9271) — Optional support for USB wireless adapters in monitor mode to detect rogue APs and deauth floods on 802.11 networks.

Tech Stack ---------------
Layer	- Technology
Backend API -	FastAPI + Uvicorn
Database -	SQLite (dev) / PostgreSQL (prod) via SQLAlchemy + Alembic
Packet Capture -	Scapy
Machine Learning	- scikit-learn, NumPy
Firewall -	nftables (via subprocess)
Honeypot -	Cowrie (Docker container)
GeoIP	- MaxMind GeoLite2 + geoip2
Auth -	JWT (python-jose) + bcrypt
Scheduler -	APScheduler
Mobile/Desktop App	Flutter 3 (Dart)
State Management	- Provider + GoRouter
Charts & Maps	- fl_chart, flutter_map
Deployment	- Docker + Docker Compose
