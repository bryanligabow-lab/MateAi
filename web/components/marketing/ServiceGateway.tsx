"use client";
import { useEffect, useState } from "react";
import Link from "next/link";
import { type Language, serviceCatalog } from "@/config/services";

const copy={
  es:{kicker:"Soluciones digitales con IA",title:"¿Qué servicio necesitas?",intro:"Selecciona una solución para descubrir cómo MateAI puede construirla alrededor de tu negocio.",services:"Selecciona un servicio",featured:"Experiencia disponible",enter:"Ver servicio",custom:"¿No encuentras exactamente lo que necesitas?",talk:"Cuéntanos tu idea",back:"Cambiar idioma"},
  en:{kicker:"AI-powered digital solutions",title:"What service do you need?",intro:"Choose a solution to discover how MateAI can build it around your business.",services:"Choose a service",featured:"Full experience available",enter:"Explore service",custom:"Don’t see exactly what you need?",talk:"Tell us your idea",back:"Change language"},
} as const;

const welcomes=["Bienvenidos","Welcome","Bienvenue","Bem-vindos","Benvenuti","Willkommen"];
type Phase="preloader"|"language"|"services";

export function ServiceGateway(){const [phase,setPhase]=useState<Phase>("preloader");const [word,setWord]=useState(0);const [language,setLanguage]=useState<Language>("es");const t=copy[language];
 useEffect(()=>{history.scrollRestoration="manual";window.scrollTo(0,0);document.body.style.overflow="hidden";const interval=window.setInterval(()=>setWord(value=>(value+1)%welcomes.length),430);const timer=window.setTimeout(()=>setPhase("language"),2800);return()=>{window.clearInterval(interval);window.clearTimeout(timer);document.body.style.overflow=""}},[]);
 const chooseLanguage=(value:Language)=>{setLanguage(value);setPhase("services");window.scrollTo(0,0)};
 if(phase==="preloader")return <section className="welcome-screen" aria-live="polite"><div className="welcome-mark">MATEAI</div><div className="welcome-word" key={welcomes[word]}>{welcomes[word]}</div><div className="welcome-progress"><i/></div></section>;
 if(phase==="language")return <section className="choice-screen"><div className="choice-glow"/><div className="choice-shell"><span className="choice-brand">MATEAI</span><span className="section-kicker">Bienvenidos · Welcome</span><h1>Elige tu idioma.<br/><em>Choose your language.</em></h1><div className="choice-actions"><button onClick={()=>chooseLanguage("es")}><small>01</small><strong>Español</strong><span>Continuar ↗</span></button><button onClick={()=>chooseLanguage("en")}><small>02</small><strong>English</strong><span>Continue ↗</span></button></div></div></section>;
 return <section className="gateway"><div className="gateway-glow gateway-glow-one"/><div className="gateway-glow gateway-glow-two"/><div className="gateway-shell"><button className="gateway-back" onClick={()=>setPhase("language")}>← {t.back}</button><div className="gateway-copy"><span className="section-kicker">{t.kicker}</span><h1>{t.title}</h1><p>{t.intro}</p></div><div className="service-heading"><span>{t.services}</span><i/></div><div className="service-cards">{serviceCatalog.map((service,index)=>{const content=service[language];return <Link className={`service-card ${"featured" in service&&service.featured?"featured":""}`} href={`/services/${service.slug}?lang=${language}`} key={service.slug}><div className="service-card-top"><b>{service.icon}</b><span>0{index+1}</span></div><h2>{content.name}</h2><p>{content.description}</p><div className="service-card-link"><span>{t.enter}</span><i>↗</i></div>{"featured" in service&&service.featured&&<em>{t.featured}</em>}</Link>})}</div><div className="gateway-custom"><span>{t.custom}</span><Link href={`/contact?lang=${language}`}>{t.talk} <b>↗</b></Link></div></div></section>}
