import type { Metadata } from "next";
import Link from "next/link";
import { notFound } from "next/navigation";
import { getService, type Language, serviceCatalog } from "@/config/services";

export function generateStaticParams(){return serviceCatalog.filter(service=>service.slug!=="web-chatbots").map(service=>({slug:service.slug}))}
export async function generateMetadata({params}:{params:Promise<{slug:string}>}):Promise<Metadata>{const service=getService((await params).slug);return service?{title:service.en.name,description:service.en.description}:{}}

const shared={
 es:{kicker:"Soluciones MateAI",headline:"Tecnología creada alrededor de tu negocio.",body:"Diseñamos, construimos e integramos la solución completa para que tu equipo trabaje mejor, responda más rápido y convierta más oportunidades.",discover:"Lo que podemos construir",items:["Estrategia y diseño del flujo","Implementación a la medida","Integración con tus herramientas","Medición, soporte y mejora continua"],cta:"Hablemos de tu proyecto",back:"Ver todos los servicios"},
 en:{kicker:"MateAI Solutions",headline:"Technology built around your business.",body:"We design, build, and integrate the complete solution so your team works better, responds faster, and converts more opportunities.",discover:"What we can build",items:["Workflow strategy and design","Custom implementation","Integration with your existing tools","Measurement, support, and continuous improvement"],cta:"Let’s discuss your project",back:"View all services"},
} as const;

export default async function ServicePage({params,searchParams}:{params:Promise<{slug:string}>;searchParams:Promise<{lang?:string}>}){const service=getService((await params).slug);if(!service||service.slug==="web-chatbots")notFound();const language:Language=(await searchParams).lang==="en"?"en":"es";const content=service[language];const t=shared[language];return <><section className="service-hero"><div className="service-hero-orb"/><div><Link className="service-back" href="/">← {t.back}</Link><span className="section-kicker">{t.kicker} · {service.icon}</span><h1>{content.name}</h1><p>{content.description}</p><Link className="button button-primary" href={`/book-demo?lang=${language}`}>{t.cta} <span>↗</span></Link></div></section><section className="section service-detail"><span className="section-kicker">{t.discover}</span><h2>{t.headline}</h2><p className="section-intro">{t.body}</p><div className="service-detail-grid">{t.items.map((item,index)=><div key={item}><b>0{index+1}</b><h3>{item}</h3></div>)}</div></section></>}
