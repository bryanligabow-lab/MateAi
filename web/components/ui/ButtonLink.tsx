import Link from "next/link";
import type { MouseEventHandler, ReactNode } from "react";

export function ButtonLink({ href, children, variant="primary", className="", onClick }: { href:string; children:ReactNode; variant?:"primary"|"secondary"|"ghost"; className?:string; onClick?:MouseEventHandler<HTMLAnchorElement> }) {
  return <Link className={`button button-${variant} ${className}`} href={href} onClick={onClick}>{children}</Link>;
}
