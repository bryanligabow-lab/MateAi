import Link from "next/link";

export function Logo() {
  return <Link className="logo" href="/" aria-label="MateAI home"><span className="logo-mark">M</span><span>MateAI</span></Link>;
}
