# AGENTS.md

## Cursor Cloud specific instructions

### Product overview

**Frame Configurator** is a single local web app (React + Vite frontend, Express API backend) for designing window/frame assemblies. Project state persists in `server-data/project.json` (gitignored).

### Services

| Service | Port | Start command |
|---------|------|---------------|
| Express API | `8787` (default) | `npm run server` |
| Vite dev server | `5181` | `npm run dev` |

For full dev mode (API + HMR), run **both** services. Vite proxies `/api` to `http://localhost:8787` (see `vite.config.ts`).

Alternative one-process mode after build: `npm run stable:lan` (Express serves API + `dist/` on port `5181`).

### Common commands

See `README.md` and `package.json` scripts:

- **Lint:** `npm run lint` (warnings only in `App.tsx` hook deps — no errors)
- **Build:** `npm run build`
- **Tests:** No automated test suite is configured; CI runs lint + build only (`.github/workflows/ci.yml`)

### Dev server startup

Prefer two tmux sessions (reliable for long-running processes):

```bash
# Session 1 — API
npm run server

# Session 2 — frontend
npm run dev
```

`npm run dev:full` starts both in one shell (`node server.js & vite ...`); use in a tmux session if you want a single command.

### Gotchas

- **Frontend-only** (`npm run dev` without API): UI loads but persistence falls back to in-memory template; always run the Express server for save/load.
- **`server-data/project.json`** is gitignored; demo edits do not need to be committed. Auto-save also writes snapshots under `server-data/history/`.
- **Node 20+** matches CI; the VM ships with Node 22.
