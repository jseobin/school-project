# school-project

Cloudflare-native full-stack starter.

## Stack

- Frontend: Cloudflare Pages static assets
- Backend: Cloudflare Pages Functions
- Database: Cloudflare D1

## Project files

- `wrangler.toml`: source of truth for Pages and D1 bindings
- `functions/api/health.js`: runtime and database health endpoint
- `functions/api/messages.js`: guestbook-style read/write API
- `migrations/0001_initial.sql`: D1 schema

## Commands

- Apply the schema remotely:
  `npx wrangler d1 execute school-project-db --remote --file migrations/0001_initial.sql`
- Start local development:
  `npx wrangler pages dev .`
- List Pages deployments:
  `npx wrangler pages deployment list --project-name school-project`
