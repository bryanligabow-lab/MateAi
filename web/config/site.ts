export const siteConfig = {
  name: "MateAI",
  description: "AI receptionists for local service businesses.",
  url: process.env.NEXT_PUBLIC_SITE_URL || "https://mateaiagents.com",
  salesEmail: process.env.NEXT_PUBLIC_SALES_EMAIL || "hello@mateaiagents.com",
  supportEmail: process.env.NEXT_PUBLIC_SUPPORT_EMAIL || "support@mateaiagents.com",
  contactEmail: process.env.NEXT_PUBLIC_CONTACT_EMAIL || "hello@mateaiagents.com",
  legalReviewRequired: true,
  cancellation: {
    billingCycle: "Monthly",
    notice: "Cancel before your next renewal to avoid the next recurring charge.",
    refundWindowDays: null as number | null,
    setupFeeRefundable: null as boolean | null,
  },
} as const;

export const navigation = [
  { label: "Services", href: "/" },
  { label: "Chatbots", href: "/services/web-chatbots" },
  { label: "Industries", href: "/industries" },
  { label: "Pricing", href: "/pricing" },
  { label: "Demo", href: "/demo" },
  { label: "About", href: "/about" },
] as const;
