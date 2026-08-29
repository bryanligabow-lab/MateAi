export type PlanKey = "starter" | "growth" | "pro";

export type PricingPlan = {
  key: PlanKey;
  name: string;
  displayPrice: string;
  setupDisplay: string;
  productId?: string;
  monthlyPriceId?: string;
  setupPriceId?: string;
  features: string[];
  cta: string;
  popular?: boolean;
  salesOnly?: boolean;
};

export const pricingPlans: PricingPlan[] = [
  {
    key: "starter",
    name: "Starter",
    displayPrice: "$149",
    setupDisplay: "Setup from $399",
    monthlyPriceId: process.env.NEXT_PUBLIC_PADDLE_STARTER_MONTHLY_PRICE_ID,
    setupPriceId: process.env.NEXT_PUBLIC_PADDLE_STARTER_SETUP_PRICE_ID,
    features: ["AI Website Assistant", "24/7 Answers", "Lead Capture", "English + Spanish", "Email Notifications"],
    cta: "Get started",
  },
  {
    key: "growth",
    name: "Growth",
    displayPrice: "$299",
    setupDisplay: "Setup from $799",
    monthlyPriceId: process.env.NEXT_PUBLIC_PADDLE_GROWTH_MONTHLY_PRICE_ID,
    setupPriceId: process.env.NEXT_PUBLIC_PADDLE_GROWTH_SETUP_PRICE_ID,
    features: ["Everything in Starter", "Appointments", "CRM Integration", "Custom Lead Qualification", "Analytics", "Automations"],
    cta: "Get started",
    popular: true,
  },
  {
    key: "pro",
    name: "Pro",
    displayPrice: "$499+",
    setupDisplay: "Setup from $1,499",
    monthlyPriceId: process.env.NEXT_PUBLIC_PADDLE_PRO_MONTHLY_PRICE_ID,
    setupPriceId: process.env.NEXT_PUBLIC_PADDLE_PRO_SETUP_PRICE_ID,
    features: ["Everything in Growth", "Custom Integrations", "Multiple Locations", "Advanced Workflows", "Priority Support"],
    cta: "Talk to sales",
    salesOnly: true,
  },
];

export function getPlan(key: string) {
  return pricingPlans.find((plan) => plan.key === key);
}
