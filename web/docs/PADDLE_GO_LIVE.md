# Paddle go-live checklist

Do not perform these steps without explicit authorization.

1. Complete live verification and domain approval.
2. Configure checkout website/default payment link if required.
3. Create separate live products and recurring/setup prices.
4. Create live client token, server API key, webhook destination, and webhook secret.
5. Replace every Sandbox value; Sandbox IDs do not work in Live.
6. Set `NEXT_PUBLIC_PADDLE_ENV=production` only after approval.
7. Deploy secrets with Cloudflare, never Git.
8. Run a controlled transaction and verify webhook, subscription, exactly one provisioning job, analytics, cancellation, and policy copy before activating campaigns.
