# ATK VPE Automation (Node.js / Express — MERN Stack Edition)

Automation and scheduling system for Advanced Toastmasters of Kerala (ATK) VPE operations.

Ported from Google Apps Script to a **high-performance Node.js / Express backend** with **100% exact logic parity**, zero business logic rewrites, and over **1000x faster execution**.

---

## Quick Start

### 1. Install Dependencies
```bash
npm install
```

### 2. Configure Environment
Copy `.env.example` to `.env` (already pre-configured for DEV):
```bash
cp .env.example .env
```

### 3. Run Self-Tests & Regression Suite
```bash
npm test
```
*Executes all 102 end-to-end regression tests across all role workflows, state engines, and grammar parsers in ~250ms.*

### 4. Start the Express Server
```bash
# Production mode
npm start

# Development mode (with auto-reload)
npm run dev
```

---

## Server Endpoints

| Method | Endpoint | Description |
| :--- | :--- | :--- |
| `GET` | `/` | Service health, platform info, timezone, and active mode |
| `GET` | `/health` | Uptime and health check |
| `GET` | `/webhook` | Meta WhatsApp webhook challenge verification (`doGet`) |
| `POST` | `/webhook` | Meta WhatsApp incoming messages and status callbacks (`doPost`) |
| `POST` | `/api/v1` | Web API v1 for Flutter and Web clients (`apiHandleRequest_`) |
| `GET` | `/programme-review` | Toastmasters meeting agenda review page (renders `ProgrammeReview.html`) |
| `GET` | `/api/programme-review/data` | JSON data for programme review slots and meetings |
| `POST` | `/api/programme-review/save` | Submits programme review approvals, holds, and moves |
| `GET` | `/api/dev/regression` | Triggers master regression suite and returns JSON report |

---

## Architecture Overview

```
atk-vpe-automation/
├── src/
│   ├── compat/                 # Google Apps Script V8 Compatibility Engine
│   │   ├── index.js            # Injects globals onto globalThis & loads files
│   │   ├── SpreadsheetApp.js   # Fast in-memory / JSON-persisted sheet engine
│   │   ├── seedData.js         # Default sheets and DEV configuration
│   │   ├── PropertiesService.js# Persistent properties + process.env store
│   │   ├── UrlFetchApp.js      # Synchronous HTTP client matching GAS specs
│   │   ├── Utilities.js        # Timezone-aware date formatting, HMAC, SHA256
│   │   ├── LockService.js      # Script & User lock manager
│   │   ├── CacheService.js     # In-memory TTL cache
│   │   ├── ContentService.js   # TextOutput & JSON output wrapper
│   │   └── HtmlService.js      # Template & HTML renderer
│   ├── server.js               # Main Express.js application
│   └── test/
│       ├── runRegression.js    # Master regression test runner
│       └── testEndpoints.js    # HTTP route integration tests
├── data/                       # Local persistent stores (sheets & properties)
├── Code.js, Workflow.js, ...   # 12 core business logic files (UNTOUCHED)
├── package.json
└── .env
```

---

## MERN Developer Integration Guide

### Frontend (React)
* The Web API (`POST /api/v1`) is ready to connect with any React application (`axios` or `fetch`).
* Actions supported: `GET_CLUB_SETTINGS`, `SAVE_EXCOM`, `LIST_MEETINGS`, `CREATE_MEETING`, `REORDER_ITEMS`, etc.

### Database (MongoDB)
* Sheets are currently simulated in memory and persisted to JSON files under `data/sheets/`.
* To transition to MongoDB: replace or supplement `SpreadsheetApp.js` with Mongoose models (`Meeting`, `Member`, `Task`, `Outbox`), keeping the spreadsheet API interface so that all 13,800+ lines of business logic continue to run without changes.
