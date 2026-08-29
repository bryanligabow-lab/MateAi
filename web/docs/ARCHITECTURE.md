# Architecture

The App Router owns routes and SEO. Components are separated into marketing, forms, checkout, animation, and UI. `config/` is the single source for pricing, industries, and public data. Verified payment event -> subscription service -> provisioning service keeps Paddle decoupled from future MateAI, EasyPanel, CRM, or n8n adapters.

Cloudflare Workers runs the app and D1 stores webhook idempotency, subscription summaries, pending provisioning jobs, and rate-limited leads. No card data is handled. `BillingExportService` is only a future VisualFAC/SRI boundary and sends nothing today. Dashboard names and values are marked demo data; simulated chat is not represented as live AI.
