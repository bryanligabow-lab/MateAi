"use client";
import { motion, useReducedMotion } from "motion/react";
import type { ReactNode } from "react";

export function Reveal({children,className="",delay=0}:{children:ReactNode;className?:string;delay?:number}){
 const reduced=useReducedMotion(); return <motion.div className={className} initial={reduced?false:{opacity:0,y:24,filter:"blur(8px)"}} whileInView={{opacity:1,y:0,filter:"blur(0px)"}} viewport={{once:true,amount:.2}} transition={{duration:.75,delay,ease:[.22,.8,.2,1]}}>{children}</motion.div>;
}
