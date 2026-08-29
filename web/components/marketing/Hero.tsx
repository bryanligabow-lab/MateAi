"use client";
import { motion, useReducedMotion } from "motion/react";
import { ButtonLink } from "@/components/ui/ButtonLink";
import { ChatDemo } from "./ChatDemo";
import { getIndustry } from "@/config/industries";
import { trackEvent } from "@/lib/analytics";

export function Hero(){const reduced=useReducedMotion();const industry=getIndustry("plumbing")!;
 return <section className={`hero ${reduced?"reduced":""}`}><div className="hero-glow"/><div className="hero-grid"><div className="hero-copy"><motion.div className="eyebrow hero-enter enter-1"><span/>AI receptionists for local businesses</motion.div><motion.h1 className="hero-enter enter-2">Turn website visitors into customers — <em>24/7.</em></motion.h1><motion.p className="hero-sub hero-enter enter-3">MateAI gives your business an AI receptionist that answers questions, captures leads, and books appointments — even when you&apos;re closed.</motion.p><motion.div className="hero-actions hero-enter enter-4" onClick={(e)=>{const target=e.target as HTMLElement;if(target.closest("a")?.getAttribute("href")==="/demo")trackEvent("hero_try_ai")}}><ButtonLink href="/demo">Try the AI <span>↗</span></ButtonLink><ButtonLink href="/book-demo" variant="secondary">Book a Free Demo</ButtonLink></motion.div><motion.div className="trust-row hero-enter enter-5">{["Available 24/7","English + Spanish","Works with your website","No rebuild required"].map(item=><span key={item}>✓ {item}</span>)}</motion.div></div><motion.div className="hero-product hero-enter enter-product"><div className="product-orbit orbit-one">Lead captured</div><div className="product-orbit orbit-two">Appointment ready</div><ChatDemo industry={industry}/></motion.div></div></section>}
