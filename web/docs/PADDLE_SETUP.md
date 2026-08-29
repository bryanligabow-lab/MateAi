# Paddle Sandbox setup

1. Create Starter and Growth products in Paddle Sandbox.
2. Add monthly + one-time setup prices: Starter USD 149/month + USD 399; Growth USD 299/month + USD 799. Never invent IDs.
3. Put each Sandbox `pri_...` ID in its matching public environment variable. Pro remains sales-only.
4. Create a Sandbox client-side token. Never expose an API key.
5. Configure the approved website/default payment link if Paddle requests it.
6. Add `https://YOUR_DOMAIN/api/paddle/webhook`, subscribe to transaction completed/payment failed and subscription created/updated/canceled/paused/resumed, then set its secret with Wrangler.

Paddle supports recurring and one-time items together when recurring intervals match. Terms, Privacy, and recurring authorization are required and never preselected. Test Starter, Growth, success/failure, duplicate delivery, lifecycle events, direct/refreshed success URL, missing token, invalid ID, and mobile. Only a verified `transaction.completed` webhook queues provisioning.
