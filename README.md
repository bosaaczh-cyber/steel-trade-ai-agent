# 🏗️ Steel Trade AI Agent System

**Multi-Agent AI system for B2B steel pipe export automation**

An end-to-end AI-powered sales automation platform built for the steel pipe export industry. The system uses multiple AI agents (powered by CrewAI + Claude API) to automatically find potential customers, score leads, generate personalized outreach emails in multiple languages, and send them via cloud email services — all managed through a web dashboard.

> 🔥 This is NOT a demo or tutorial project — it's a production system that has been used to send real outreach emails to real companies across the Middle East and Latin America.

---

## System Architecture

```
┌─────────────────────────────────────────────────────┐
│                   Web Dashboard (Flask)              │
│         Login / Settings / Leads / Send / MTC        │
└──────────────────────┬──────────────────────────────┘
                       │
        ┌──────────────┼──────────────┐
        ▼              ▼              ▼
  ┌──────────┐  ┌───────────┐  ┌───────────────┐
  │ Research  │  │  Scoring  │  │ Email Writer  │
  │  Agent    │  │  Agent    │  │    Agent      │
  │ (CrewAI)  │  │ (CrewAI)  │  │  (CrewAI)    │
  └─────┬─────┘  └─────┬─────┘  └──────┬────────┘
        │              │               │
        ▼              ▼               ▼
  ┌──────────┐  ┌───────────┐  ┌───────────────┐
  │  Serper   │  │   Lead    │  │  Aliyun DM    │
  │ Search API│  │ Database  │  │ (Email Send)  │
  └──────────┘  └───────────┘  └───────────────┘
```

## What It Does

### 1. 🔍 Customer Research (Multi-Agent)
- Uses Serper API to search for potential B2B customers by region and industry
- **Research Agent** extracts structured lead data (company, contact, needs) from search results
- Supports multiple target regions: Middle East, Latin America, Southeast Asia

### 2. 📊 Lead Scoring
- **Scoring Agent** evaluates each lead on multiple dimensions:
  - Company scale and procurement potential
  - Match with product capabilities
  - Market accessibility
- Classifies leads into A/B/C tiers for prioritized outreach

### 3. ✉️ AI Email Generation
- **Writer Agent** generates personalized cold emails for A/B tier leads
- Supports English and Spanish (with proper usted format)
- Each email references the prospect's specific projects and needs
- Generates complete emails with Subject line and body

### 4. 📤 Batch Email Sending
- Integrates with Aliyun DirectMail API (no SDK dependency, pure REST)
- Rate-limited sending with configurable intervals
- Full send logging with success/failure tracking
- Auto-sync sent emails to follow-up database

### 5. 📋 MTC Document Generation
- Generates Mill Test Certificates (MTC) for steel pipe orders
- Fills in mechanical properties, chemical composition, test results
- Exports to formatted documents

### 6. 🌐 Web Dashboard
- User authentication system with role-based access
- Company profile management (customizable for any steel company)
- Real-time send progress tracking
- Lead management with search and filtering
- Send history and statistics

## Tech Stack

| Component | Technology |
|-----------|-----------|
| AI Framework | CrewAI (Multi-Agent orchestration) |
| LLM | Claude API (via OpenAI-compatible endpoint) |
| Web Search | Serper API |
| Backend | Flask (Python) |
| Email Service | Aliyun DirectMail (REST API, no SDK) |
| Frontend | HTML + Vanilla JS + CSS |
| Data Storage | JSON file-based |

## Project Structure

```
├── app.py                 # Flask web server & API routes
├── steel_master.py        # Full pipeline: search → score → write → export
├── steel_bulk.py          # Bulk outreach for multiple regions/industries
├── Steel_sender.py        # Email sending module (Aliyun DirectMail)
├── Steel_stats.py         # Analytics and reporting
├── steel_email_finder.py  # Email discovery and enrichment
├── steel_followup.py      # Follow-up tracking system
├── steel_mtc.py           # Mill Test Certificate generator
├── templates/
│   ├── base.html          # Base template with navigation
│   ├── dashboard.html     # Main dashboard
│   ├── leads.html         # Lead management page
│   ├── send.html          # Email sending interface
│   ├── settings.html      # Company settings
│   ├── mtc.html           # MTC generator page
│   ├── detail.html        # Lead detail view
│   ├── login.html         # Login page
│   └── register.html      # Registration page
└── .env.example           # Environment variables template
```

## Quick Start

### 1. Clone and install dependencies

```bash
git clone https://github.com/YOUR_USERNAME/steel-trade-ai-agent.git
cd steel-trade-ai-agent
pip install flask crewai python-dotenv python-docx requests
```

### 2. Configure environment variables

```bash
cp .env.example .env
# Edit .env with your API keys
```

### 3. Run the web dashboard

```bash
python app.py
# Open http://localhost:5000
# Default login: admin / gmt2026
```

### 4. Or run the full pipeline from command line

```bash
python steel_master.py
```

## Design Decisions

### Why Multi-Agent instead of Single-Agent?

Each agent has a focused role and specialized prompt, which produces better results than a single agent trying to do everything:

- **Research Agent**: Optimized for extracting structured data from messy search results
- **Scoring Agent**: Has domain expertise prompt for evaluating steel pipe trade leads
- **Writer Agent**: Specialized in B2B cold email writing with industry-specific knowledge

### Why Aliyun DirectMail without SDK?

The official Aliyun SDK is heavy and has many dependencies. I implemented the API signature algorithm from scratch using only Python standard libraries (`hmac`, `hashlib`, `urllib`), which makes the system lighter and easier to deploy.

### Why JSON instead of a Database?

For a single-user or small-team tool processing hundreds of leads, JSON files are simpler to set up, easier to debug (just open the file), and have zero infrastructure requirements. If scaling to thousands of concurrent users, migrating to SQLite or PostgreSQL would be the next step.

## Screenshots

*(Coming soon)*

## License

MIT License — feel free to learn from and adapt this code for your own use.

## About

Built by a steel trade professional who learned AI development to solve real business problems. This project demonstrates that domain expertise + AI application skills can create genuine business value.

If you're interested in the technical details or have questions about the architecture, feel free to open an issue!
