# Bloomberg Platform — Recovery Guide

## Quick facts

| Item | Value |
|------|-------|
| Platform URL | http://localhost/platform |
| Platform URL (LAN) | http://10.18.68.179/platform/#reporter |
| Served by | `jira-dashboard` PM2 process (port 80) |
| Platform files | `C:\Users\cnegoescu\bloomberg-platform\public\` |
| Main file | `C:\Users\cnegoescu\bloomberg-platform\public\index.html` |
| GitHub repo | https://github.com/cristiannegoescu-ux/Bloomberg-Platform |
| Reporter Workspace | `C:\Users\cnegoescu\jira-dashboard` — **do not touch** |
| Reporter Workspace URL | http://localhost (port 80, root path) |

---

## GitHub login (if on a new machine)

```
Username : cristiannegoescu-ux
Email    : cristian.negoescu@ubisoft.com
```

Authenticate the GitHub CLI:
```powershell
gh auth login
# Choose: GitHub.com → HTTPS → Login with a web browser
```

---

## Platform features

- **No access gate** — opens directly, no password required
- **Blue circle favicon** with Bloomberg B logo
- **Pages:** Reporter, Portal, Integrations, Subscription Plans
- **Bloom AI chat widget** — bottom-right floating assistant

### Reporter page

- **Inline report card** — shows crash/defect injected by Bloom Assistant
- **Call Stack** — displayed in report card
- **Screenshot** — game canvas screenshot attached to each report
- **History sidebar** — all received reports, each deletable with ✕
- **Sent Reports table** — persisted to localStorage, survives page reload
  - Each row is deletable with ✕
  - Click the ID link to open a full detail modal (stack + screenshot)
- **Drag-and-drop .dmp** — drop a `.bloom.dmp` file onto the report area to
  generate a report card from file (screenshot included if present)

### Bloom AI integration — crash report injection flow

```
Game (S/F key)
  → POST http://127.0.0.1:3333/defect   (JSON with screenshot, stack, metadata)
  → Bloom Assistant main.js
      → stores pendingReport
      → shows toast notification (dark card, no flash, non-clickable)
      → widget turns red with badge count
  → User clicks Bloom widget
      → dismisses toast
      → opens Platform window
      → sends 'show-report' IPC to preload.js
      → preload dispatches CustomEvent 'bloom-show-report' on window
      → page bloomShowReport() renders inline report card + history entry
```

Bloom Assistant injects via `preload.js` using a `CustomEvent` that crosses the
`contextIsolation` boundary — direct `window.bloomShowReport()` calls from the
preload do NOT work because the preload and page have separate window objects.

Payload shape sent by the game:
```json
{
  "id":         "CRASH-0001",
  "type":       "Access Violation",
  "severity":   "Critical",
  "component":  "GameEngine",
  "summary":    "Access violation — read from null pointer at 0x00000000",
  "session":    "ABC123",
  "score":      80,
  "stack":      "AccessViolation: 0x00000000\n  at MemoryManager.read() :88",
  "timestamp":  "2026-06-30 09:14:02 UTC",
  "isCrash":    true,
  "screenshot": "data:image/png;base64,..."
}
```

---

## Restore from scratch (new PC or after disk loss)

```powershell
# 1. Clone the platform
git clone https://github.com/cristiannegoescu-ux/Bloomberg-Platform.git C:\Users\cnegoescu\bloomberg-platform
cd C:\Users\cnegoescu\bloomberg-platform
npm install

# 2. Clone the Reporter Workspace (separate repo — leave it untouched)
git clone https://github.com/cristiannegoescu-ux/Bloomberg-Reporter-Next-Dashboard.git C:\Users\cnegoescu\jira-dashboard
cd C:\Users\cnegoescu\jira-dashboard
npm install
cd backend && npm install
cd ..\frontend && npm install && npm run build

# 3. Create backend .env
# File: C:\Users\cnegoescu\jira-dashboard\backend\.env
# Contents:
#   JIRA_URL=https://ne1-tomcat-jira153-dc.ubisoft.org/jira
#   JIRA_PROJECT_KEY=QF#
#   JIRA_BOARD_ID=<your board id>
#   PORT=80

# 4. Install PM2 globally if missing
npm install -g pm2

# 5. Start both processes
cd C:\Users\cnegoescu\jira-dashboard
pm2 start backend/server.js --name jira-dashboard

cd C:\Users\cnegoescu\bloomberg-platform
pm2 start server.js --name bloomberg-platform

# 6. Save PM2 so it survives reboots
pm2 save
pm2 startup
# Run the command it prints
```

---

## Day-to-day: restart after PC reboot

```powershell
pm2 resurrect
# or manually:
pm2 start jira-dashboard
pm2 start bloomberg-platform
```

---

## Redeploy platform changes (edit index.html, then):

```powershell
cd C:\Users\cnegoescu\bloomberg-platform
git add public/index.html
git commit -m "Update platform UI"
git push
pm2 restart jira-dashboard
```

---

## PM2 process map

| PM2 name | Port | Serves | Repo |
|----------|------|--------|------|
| `jira-dashboard` | 80 | Reporter Workspace (`/`) + Platform (`/platform`) | bloomberg-reporter-next-dashboard |
| `bloomberg-platform` | 3001 | Platform standalone (not used in prod) | bloomberg-platform |

> **Note:** In production the platform is served by `jira-dashboard` at `/platform/`.
> The `bloomberg-platform` PM2 process is a standalone fallback only.

---

## Regenerate the PowerPoint

```powershell
cd C:\Users\cnegoescu\bloomberg-platform
node generate-pptx.js
# Output: Bloomberg_Reporter_Platform.pptx
```
