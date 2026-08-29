import Link from "next/link";
import { Logo } from "@/components/ui/Logo";
import { industries } from "@/config/industries";
import { siteConfig } from "@/config/site";

export function Footer(){
  return <footer className="footer"><div className="footer-grid"><div><Logo/><p>AI Receptionists for Local Businesses</p><a href={`mailto:${siteConfig.contactEmail}`}>{siteConfig.contactEmail}</a></div>
    <div><strong>Product</strong><Link href="/demo">Demo</Link><Link href="/pricing">Pricing</Link><Link href="/services/web-chatbots#how-it-works">How It Works</Link></div>
    <div><strong>Industries</strong>{industries.slice(0,4).map(i=><Link key={i.key} href={`/${i.key}`}>{i.name}</Link>)}</div>
    <div><strong>Company</strong><Link href="/about">About</Link><Link href="/contact">Contact</Link><Link href="/book-demo">Book a Demo</Link></div>
    <div><strong>Legal</strong><Link href="/privacy">Privacy</Link><Link href="/terms">Terms</Link><Link href="/cancellation">Cancellation</Link><Link href="/ai-disclosure">AI Disclosure</Link></div></div>
    <div className="footer-bottom"><span>© {new Date().getFullYear()} MateAI</span><span>Built for businesses that value every lead.</span></div></footer>;
}
