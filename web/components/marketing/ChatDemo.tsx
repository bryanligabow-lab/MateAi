"use client";
import { motion, useReducedMotion } from "motion/react";
import { useEffect, useState } from "react";
import type { Industry } from "@/config/industries";

export function ChatDemo({industry,compact=false}:{industry:Industry;compact?:boolean}){
 const reduced=useReducedMotion(); const [visible,setVisible]=useState(reduced?industry.conversation.length:1);
 useEffect(()=>{setVisible(1);if(reduced){setVisible(industry.conversation.length);return;}const timer=setInterval(()=>setVisible(v=>v>=industry.conversation.length?1:v+1),1700);return()=>clearInterval(timer)},[industry,reduced]);
 return <div className={`chat-demo ${compact?"compact":""}`} aria-label={`Simulated ${industry.name} AI receptionist demo`}><div className="chat-top"><span className="status-dot"/><div><strong>{industry.name} Service</strong><small>MateAI receptionist · online</small></div><span className="demo-label">SIMULATION</span></div><div className="chat-body">{industry.conversation.slice(0,visible).map((message,index)=><motion.div key={`${industry.key}-${index}`} className={`message ${message.role}`} initial={false} animate={{opacity:1,y:0,scale:1}}>{message.text}</motion.div>)}{visible<industry.conversation.length&&<div className="typing"><i/><i/><i/></div>}</div><div className="chat-input">Type a message… <span>↑</span></div></div>;
}
