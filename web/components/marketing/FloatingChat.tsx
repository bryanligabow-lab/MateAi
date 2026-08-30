"use client";
import { type FormEvent, useEffect, useRef, useState } from "react";
import { usePathname } from "next/navigation";

type Message = { role: "user" | "assistant"; content: string };

export function FloatingChat(){
 const pathname=usePathname();const [available,setAvailable]=useState(pathname!=="/");const [open,setOpen]=useState(false);const [loading,setLoading]=useState(false);const [input,setInput]=useState("");const [messages,setMessages]=useState<Message[]>([{role:"assistant",content:"Hi! I’m MateAI. Ask me anything — English or Español."}]);const endRef=useRef<HTMLDivElement>(null);
 useEffect(()=>{if(pathname!=="/"){setAvailable(true);return}setAvailable(false);const show=()=>setAvailable(true);window.addEventListener("mateai:language-selected",show);return()=>window.removeEventListener("mateai:language-selected",show)},[pathname]);
 useEffect(()=>endRef.current?.scrollIntoView({behavior:"smooth",block:"nearest"}),[messages,loading]);
 const submit=async(event:FormEvent)=>{event.preventDefault();const content=input.trim();if(!content||loading)return;const next=[...messages,{role:"user" as const,content}];setMessages(next);setInput("");setLoading(true);try{const response=await fetch("/api/chat",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({messages:next.slice(-12),language:"en"})});const data=await response.json() as {answer?:string;error?:string};setMessages(current=>[...current,{role:"assistant",content:data.answer??data.error??"I couldn’t answer right now."}])}catch{setMessages(current=>[...current,{role:"assistant",content:"Connection interrupted. Please try again."}])}finally{setLoading(false)}};
 if(!available)return null;
 return <aside className={`floating-chat ${open?"is-open":""}`}><button className="floating-chat-launcher" onClick={()=>setOpen(value=>!value)} aria-expanded={open} aria-controls="mateai-floating-panel"><span className="floating-chat-pulse"/><b>{open?"Close":"Chat with MateAI"}</b><i>{open?"×":"↗"}</i></button><div className="floating-chat-panel" id="mateai-floating-panel" aria-hidden={!open}><header><span className="status-dot"/><div><strong>MateAI</strong><small>Ollama-powered · online</small></div></header><div className="floating-chat-messages" aria-live="polite">{messages.map((message,index)=><div className={`message ${message.role==="user"?"visitor":"ai"}`} key={index}>{message.content}</div>)}{loading&&<div className="typing"><i/><i/><i/></div>}<div ref={endRef}/></div><form onSubmit={submit}><input value={input} onChange={event=>setInput(event.target.value)} placeholder="Ask me anything…" maxLength={2000} disabled={loading} aria-label="Message MateAI"/><button disabled={loading||!input.trim()} aria-label="Send">↑</button></form></div></aside>;
}
