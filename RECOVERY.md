# Bloomberg Platform — Recovery Guide

## Quick facts

| Item | Value |
|------|-------|
| Platform URL | http://localhost/platform |
| Platform URL (port) | http://localhost:80/platform |
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
- **Reporter page** includes `bloomShowReport(data)` — a JS function callable from
  Bloom Assistant via `executeJavaScript` to display a crash/defect modal
- **Crash modal** shows: ID, Type, Severity, Session, Timestamp, Score, Summary,
  Call Stack — with "Add to Sent Reports" button

### Bloom AI integration (crash report injection)

When Bloom Assistant receives a crash/defect from the Game Client, it calls:
```js
window.bloomShowReport({
  id, type, severity, component, session, score,
  summary, stack, timestamp, isCrash
})
```
This renders a full-screen modal on the Reporter page. No page reload required.

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
