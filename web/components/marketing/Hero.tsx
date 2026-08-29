"use client";
import { ButtonLink } from "@/components/ui/ButtonLink";
import { ChatDemo } from "./ChatDemo";
import { getIndustry } from "@/config/industries";
import { trackEvent } from "@/lib/analytics";

export function Hero(){const industry=getIndustry("plumbing")!;
 return <section className="hero"><div className="hero-glow"/><div className="hero-grid"><div className="hero-copy"><div className="eyebrow hero-motion motion-1"><span/>AI receptionists for local businesses</div><h1 className="hero-motion motion-2">Turn website visitors into customers — <em>24/7.</em></h1><p className="hero-sub hero-motion motion-3">MateAI gives your business an AI receptionist that answers questions, captures leads, and books appointments — even when you&apos;re closed.</p><div className="hero-actions hero-motion motion-4"><ButtonLink href="/demo" onClick={()=>trackEvent("hero_try_ai")}>Try the AI <span>↗</span></ButtonLink><ButtonLink href="/book-demo" variant="secondary">Book a Free Demo</ButtonLink></div><div className="trust-row hero-motion motion-5">{["Available 24/7","English + Spanish","Works with your website","No rebuild required"].map(item=><span key={item}>✓ {item}</span>)}</div></div><div className="hero-product hero-enter enter-product"><div className="product-orbit orbit-one">Lead captured</div><div className="product-orbit orbit-two">Appointment ready</div><ChatDemo industry={industry}/></div></div></section>}
