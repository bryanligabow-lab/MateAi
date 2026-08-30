"use client";
import { type FormEvent, useEffect, useRef, useState } from "react";
import type { Industry } from "@/config/industries";

type ChatMessage = { role: "user" | "assistant"; content: string };

export function ChatDemo({compact=false}:{industry:Industry;compact?:boolean}){
 const [messages,setMessages]=useState<ChatMessage[]>([{role:"assistant",content:"Hi! I’m MateAI. Ask me anything — in English or Spanish."}]);
 const [input,setInput]=useState(""); const [loading,setLoading]=useState(false); const endRef=useRef<HTMLDivElement>(null);
 useEffect(()=>endRef.current?.scrollIntoView({behavior:"smooth",block:"nearest"}),[messages,loading]);
 const submit=async(event:FormEvent)=>{event.preventDefault();const content=input.trim();if(!content||loading)return;const next=[...messages,{role:"user" as const,content}];setMessages(next);setInput("");setLoading(true);try{const response=await fetch("/api/chat",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({messages:next.slice(-12),language:"en"})});const data=await response.json() as {answer?:string;error?:string};setMessages(current=>[...current,{role:"assistant",content:data.answer??data.error??"I couldn’t answer that right now."}])}catch{setMessages(current=>[...current,{role:"assistant",content:"The connection was interrupted. Please try again."}])}finally{setLoading(false)}};
 return <div className={`chat-demo live-chat ${compact?"compact":""}`} aria-label="Live MateAI assistant"><div className="chat-top"><span className="status-dot"/><div><strong>MateAI Assistant</strong><small>Ollama-powered · online</small></div><span className="demo-label">LIVE</span></div><div className="chat-body" aria-live="polite">{messages.map((message,index)=><div key={index} className={`message ${message.role==="user"?"visitor":"ai"}`}>{message.content}</div>)}{loading&&<div className="typing" aria-label="MateAI is thinking"><i/><i/><i/></div>}<div ref={endRef}/></div><form className="chat-input live-chat-input" onSubmit={submit}><input value={input} onChange={event=>setInput(event.target.value)} maxLength={2000} placeholder="Ask me anything…" aria-label="Message MateAI" disabled={loading}/><button type="submit" disabled={loading||!input.trim()} aria-label="Send message">↑</button></form></div>;
}
