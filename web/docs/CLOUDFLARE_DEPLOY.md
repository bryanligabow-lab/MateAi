# Cloudflare deployment

This project uses Cloudflare's current recommendation for new Next.js-style apps: vinext on Workers. No custom domain is hardcoded. Deploying the existing `mateai-site` Worker name preserves its attached apex/www routes.

1. Authenticate and run `npx wrangler whoami`.
2. Create D1: `npx wrangler d1 create mateai-production`.
3. Add its non-secret ID as binding `DB` in `wrangler.jsonc`.
4. Run `npx wrangler d1 migrations apply mateai-production --remote`.
5. Set secrets with `npx wrangler secret put`, never in Git.
6. Configure public build values per environment.
7. Run typecheck, lint, tests, build, then `npm run deploy`.
8. Verify apex/www, security headers, forms, signed webhook, and mobile checkout.

Use separate preview/production databases. Workers deployment versions provide rollback.
