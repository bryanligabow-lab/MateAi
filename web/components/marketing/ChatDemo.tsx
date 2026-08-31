"use client";
import { type FormEvent, useEffect, useRef, useState } from "react";
import type { Industry } from "@/config/industries";
import { streamChat, type ClientChatMessage } from "@/lib/chat-client";

type ChatMessage = ClientChatMessage;

export function ChatDemo({compact=false}:{industry:Industry;compact?:boolean}){
 const [messages,setMessages]=useState<ChatMessage[]>([{role:"assistant",content:"Hi! I’m MateAI. Ask me anything — in English or Spanish."}]);
 const [input,setInput]=useState(""); const [loading,setLoading]=useState(false); const endRef=useRef<HTMLDivElement>(null);
 useEffect(()=>endRef.current?.scrollIntoView({behavior:"smooth",block:"nearest"}),[messages,loading]);
 const submit=async(event:FormEvent)=>{event.preventDefault();const content=input.trim();if(!content||loading)return;const next=[...messages,{role:"user" as const,content}];setMessages([...next,{role:"assistant",content:""}]);setInput("");setLoading(true);try{await streamChat(next,token=>setMessages(current=>current.map((message,index)=>index===current.length-1?{...message,content:message.content+token}:message)))}catch(error){setMessages(current=>current.map((message,index)=>index===current.length-1?{...message,content:error instanceof Error?error.message:"The connection was interrupted. Please try again."}:message))}finally{setLoading(false)}};
 return <div className={`chat-demo live-chat ${compact?"compact":""}`} aria-label="Live MateAI assistant"><div className="chat-top"><span className="status-dot"/><div><strong>MateAI Assistant</strong><small>Ollama-powered · online</small></div><span className="demo-label">LIVE</span></div><div className="chat-body" aria-live="polite">{messages.map((message,index)=><div key={index} className={`message ${message.role==="user"?"visitor":"ai"}`}>{message.content}</div>)}{loading&&<div className="typing" aria-label="MateAI is thinking"><i/><i/><i/></div>}<div ref={endRef}/></div><form className="chat-input live-chat-input" onSubmit={submit}><input value={input} onChange={event=>setInput(event.target.value)} maxLength={2000} placeholder="Ask me anything…" aria-label="Message MateAI" disabled={loading}/><button type="submit" disabled={loading||!input.trim()} aria-label="Send message">↑</button></form></div>;
}
