import type { Metadata } from "next";
import "./globals.css";
import { Header } from "@/components/marketing/Header";
import { Footer } from "@/components/marketing/Footer";
import { CookieConsent } from "@/components/marketing/CookieConsent";
import { ReloadToHome } from "@/components/navigation/ReloadToHome";
import { siteConfig } from "@/config/site";

export const metadata: Metadata = {
  metadataBase: new URL(siteConfig.url),
  title: { default: "MateAI | AI Receptionists for Local Businesses", template: "%s | MateAI" },
  description: siteConfig.description,
  alternates: { canonical: "/" },
  openGraph: { type: "website", siteName: "MateAI", title: "Never miss another website lead.", description: siteConfig.description, url: "/" },
  twitter: { card: "summary_large_image", title: "MateAI", description: siteConfig.description },
  icons: { icon: "/mateai-logo.png", apple: "/mateai-logo.png" },
  robots: { index: true, follow: true },
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  const schema = { "@context":"https://schema.org", "@type":"Organization", name:"MateAI", url:siteConfig.url, description:siteConfig.description };
  return <html lang="es"><body><ReloadToHome/><a className="skip-link" href="#main">Ir al contenido</a><Header/><main id="main">{children}</main><Footer/><CookieConsent/><script type="application/ld+json" dangerouslySetInnerHTML={{__html:JSON.stringify(schema)}}/></body></html>;
}
