# AGENTS.md

## Cursor Cloud specific instructions

### Project overview

This is a basic fullstack web application (Node.js/Express backend + static HTML/CSS/JS frontend) using npm workspaces. The main application code lives on feature branches; the `main` branch may be minimal.

### Services

| Service | Directory | Command | Port |
|---------|-----------|---------|------|
| Backend (Express) + Frontend (static) | `backend/` + `frontend/` | `npm run dev` | 3000 |

The backend serves both the API endpoints and the static frontend files.

### Key commands

- **Install deps**: `npm install` (from repo root; uses npm workspaces)
- **Dev server**: `npm run dev` (starts backend with `--watch` for hot reload)
- **Production start**: `npm run start`

### API endpoints

- `GET /api/health` — health check
- `GET /api/message` — sample business endpoint
- `GET /` — serves the static frontend

### Notes

- The dev server uses Node.js `--watch` flag for automatic restarts on file changes.
- No database or external service dependencies — the app is self-contained.
- Frontend is plain HTML/CSS/JS (no build step required).
