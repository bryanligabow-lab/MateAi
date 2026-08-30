"use client";
import { useEffect } from "react";
import { usePathname } from "next/navigation";

export function ReloadToHome(){const pathname=usePathname();
 useEffect(()=>{const navigation=performance.getEntriesByType("navigation")[0] as PerformanceNavigationTiming|undefined;if(navigation?.type==="reload"&&pathname!=="/")window.location.replace("/")},[pathname]);
 return null;
}
