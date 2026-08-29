# MateAI marketing and billing website

Production-oriented SaaS site built with TypeScript, React, vinext, Tailwind, Motion, Cloudflare Workers/D1, and Paddle Billing.

## Local development

```bash
cp .env.example .env.local
npm install
npm run dev
```

Quality gates: `npm run typecheck`, `npm run lint`, `npm test`, and `npm run build`. Paddle defaults to Sandbox. Missing tokens or IDs show a clear unavailable state. Never put Paddle server secrets in a `NEXT_PUBLIC_*` variable.

See `docs/` for architecture, Paddle, and Cloudflare deployment.
