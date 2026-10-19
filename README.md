# 5 Minute Dog

Waitlist landing page for **5 Minute Dog**: one 5-minute positive-reinforcement training exercise every morning, by email (WhatsApp coming soon). First 7 days free, then from €5/month.

- `index.html` — single-file page, no dependencies.
- `llms.txt`, `sitemap.xml` — AI-assistant summary and sitemap.
- `growth-log.md` — weekly growth changes and what to measure.
- `apps-script.gs` — Google Apps Script web app that appends signups to the "5 Minute Dog – Waitlist" Google Sheet.

Setup: deploy `apps-script.gs` as a web app (Execute as: Me, Access: Anyone) and paste its `/exec` URL into `WEBAPP_URL` in `index.html`. Until then, signups are stored only in the visitor's browser (localStorage).

Hosted via GitHub Pages (Settings → Pages → Deploy from branch → `main` / root).
