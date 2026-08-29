export type AnalyticsEvent = "page_view" | "hero_try_ai" | "book_demo_click" | "pricing_view" | "plan_selected" | "checkout_started" | "checkout_completed" | "demo_started" | "lead_submitted" | "onboarding_started" | "onboarding_completed";

declare global { interface Window { dataLayer?: unknown[]; fbq?: (...args: unknown[]) => void; } }

export function hasAnalyticsConsent() {
  if (typeof window === "undefined") return false;
  return window.localStorage.getItem("mateai-consent")?.includes("analytics") ?? false;
}

export function trackEvent(name: AnalyticsEvent, properties: Record<string, string | number | boolean> = {}) {
  if (typeof window === "undefined" || !hasAnalyticsConsent()) return;
  const safe = Object.fromEntries(Object.entries(properties).filter(([key]) => !/email|phone|name|conversation|payment/i.test(key)));
  window.dataLayer?.push({ event: name, ...safe });
}
