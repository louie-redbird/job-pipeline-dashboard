# Job Pipeline Dashboard

Personal job application tracker. Apps Script web app backed by a Google Sheet, with a single-page HTML dashboard for the frontend.

## Files

- **Code.gs** — Apps Script server. All sheet reads/writes, prompt builders for Claude/ChatGPT, batch handlers.
- **Index.html** — Single-page dashboard. Vanilla JS, no build step, all CSS inline.

## Deployment

This repo is the source of truth. Apps Script is the deployment target.

Workflow:
1. Edit Code.gs or Index.html locally
2. Commit changes (git commit -am "what changed")
3. Copy contents into the Apps Script editor (the matching file)
4. Save in Apps Script
5. Deploy → Manage Deployments → edit the active deployment → Deploy (creates a new version)

The deployed dashboard is at https://jobs.redbirdagency.com.au

## Architecture

- Storage: Google Sheet, tab "Job Pipeline v2", 38 columns
- Sheet ID: stored as SHEET_ID constant at the top of Code.gs
- Frontend → backend: google.script.run.apiRouter({method, params/body}) — bypasses CORS by going through the Apps Script sandbox
- No build step. Everything is plain JS / CSS / HTML in one file each.
