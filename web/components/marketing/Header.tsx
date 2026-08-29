"use client";
import { useEffect, useState } from "react";
import Link from "next/link";
import { usePathname } from "next/navigation";
import { Logo } from "@/components/ui/Logo";
import { navigation } from "@/config/site";

export function Header() {
  const pathname=usePathname(); const gateway=pathname==="/";
  const [open,setOpen]=useState(false); const [scrolled,setScrolled]=useState(false);
  useEffect(()=>{ const onScroll=()=>setScrolled(window.scrollY>20); onScroll(); window.addEventListener("scroll",onScroll,{passive:true}); return()=>window.removeEventListener("scroll",onScroll); },[]);
  useEffect(()=>{ document.body.style.overflow=open?"hidden":""; return()=>{document.body.style.overflow=""}; },[open]);
  return <header className={`site-header ${scrolled?"is-scrolled":""}`}>
    <div className="nav-shell"><Logo/>{!gateway&&<nav className="desktop-nav" aria-label="Primary">{navigation.map(item=><Link key={item.href} href={item.href}>{item.label}</Link>)}</nav>}
      <div className="nav-actions"><Link className="button button-primary nav-cta" href={gateway?"/contact":"/book-demo"}>{gateway?"Contacto / Contact":"Book a Demo"}</Link>{!gateway&&<button className="menu-button" onClick={()=>setOpen(!open)} aria-expanded={open} aria-controls="mobile-menu" aria-label="Toggle menu"><span/><span/></button>}</div>
    </div>
    {!gateway&&<div className={`mobile-menu ${open?"is-open":""}`} id="mobile-menu">{navigation.map(item=><Link key={item.href} href={item.href} onClick={()=>setOpen(false)}>{item.label}</Link>)}<Link className="button button-primary" href="/book-demo" onClick={()=>setOpen(false)}>Book a Demo</Link></div>}
  </header>;
}
