# Push guide

This document records how `wps-vba-status-tracker` was published to GitHub, so
you can reproduce or update it.

## Option A — Standard `git` push (easiest, when your network allows it)

```bash
cd wps-vba-status-tracker
git init
git add .
git commit -m "Initial release: WPS/Excel VBA sub-item status tracker"
git branch -M main
git remote add origin https://github.com/ivy92888-glitch/wps-vba-status-tracker.git
git push -u origin main
```

If your environment blocks the Git protocol, create a fine-grained or classic
personal access token (PAT) with the **repo** scope and push with it:

```bash
git remote set-url origin https://ivy92888-glitch:<TOKEN>@github.com/ivy92888-glitch/wps-vba-status-tracker.git
git push -u origin main
```

## Option B — GitHub REST API (used here; works through restrictive proxies)

When the Git protocol is blocked, the repo was created and populated via the
GitHub REST API using a classic PAT with `repo` scope:

1. **Create the (empty) repository**
   `POST https://api.github.com/user/repos` → `{ "name": "wps-vba-status-tracker", "description": "...", "private": false }`
2. **Upload each file** with the Contents API
   `PUT https://api.github.com/repos/ivy92888-glitch/wps-vba-status-tracker/contents/<path>`
   body: `{ "message": "Add <path>", "content": "<base64>", "branch": "main" }`
   (binary files such as `examples/sample-data.xlsx` are uploaded the same way,
   base64-encoded.)

The first file upload creates the `main` branch; each subsequent upload adds a
commit. No local Git remote or protocol access is required.

## Security note

- Use a **classic PAT with the `repo` scope** (or a fine-grained token limited to
  this repo). Never commit the token to the repository.
- Delete or revoke the PAT after you finish publishing. Treat it like a password.
- This repo intentionally contains **no** company data, screenshots, or
  business-specific wording — only the generic, rewritten implementation.
