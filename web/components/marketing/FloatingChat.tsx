"use client";
import { type FormEvent, useEffect, useRef, useState } from "react";
import { usePathname } from "next/navigation";
import { streamChat, type ClientChatMessage } from "@/lib/chat-client";

type Message = ClientChatMessage;

export function FloatingChat(){
 const pathname=usePathname();const [available,setAvailable]=useState(pathname!=="/");const [open,setOpen]=useState(false);const [loading,setLoading]=useState(false);const [input,setInput]=useState("");const [messages,setMessages]=useState<Message[]>([{role:"assistant",content:"Hi! I’m MateAI. Ask me anything — English or Español."}]);const endRef=useRef<HTMLDivElement>(null);
 useEffect(()=>{if(pathname!=="/"){setAvailable(true);return}setAvailable(false);const show=()=>setAvailable(true);window.addEventListener("mateai:language-selected",show);return()=>window.removeEventListener("mateai:language-selected",show)},[pathname]);
 useEffect(()=>endRef.current?.scrollIntoView({behavior:"smooth",block:"nearest"}),[messages,loading]);
 const submit=async(event:FormEvent)=>{event.preventDefault();const content=input.trim();if(!content||loading)return;const next=[...messages,{role:"user" as const,content}];setMessages([...next,{role:"assistant",content:""}]);setInput("");setLoading(true);try{await streamChat(next,token=>setMessages(current=>current.map((message,index)=>index===current.length-1?{...message,content:message.content+token}:message)))}catch(error){setMessages(current=>current.map((message,index)=>index===current.length-1?{...message,content:error instanceof Error?error.message:"Connection interrupted. Please try again."}:message))}finally{setLoading(false)}};
 if(!available)return null;
 return <aside className={`floating-chat ${open?"is-open":""}`}><button className="floating-chat-launcher" onClick={()=>setOpen(value=>!value)} aria-expanded={open} aria-controls="mateai-floating-panel"><span className="floating-chat-pulse"/><b>{open?"Close":"Chat with MateAI"}</b><i>{open?"×":"↗"}</i></button><div className="floating-chat-panel" id="mateai-floating-panel" aria-hidden={!open}><header><span className="status-dot"/><div><strong>MateAI</strong><small>Ollama-powered · online</small></div></header><div className="floating-chat-messages" aria-live="polite">{messages.map((message,index)=><div className={`message ${message.role==="user"?"visitor":"ai"}`} key={index}>{message.content}</div>)}{loading&&<div className="typing"><i/><i/><i/></div>}<div ref={endRef}/></div><form onSubmit={submit}><input value={input} onChange={event=>setInput(event.target.value)} placeholder="Ask me anything…" maxLength={2000} disabled={loading} aria-label="Message MateAI"/><button disabled={loading||!input.trim()} aria-label="Send">↑</button></form></div></aside>;
}
