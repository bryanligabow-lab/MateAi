import Image from "next/image";
import Link from "next/link";

export function Logo() {
  return <Link className="logo" href="/" aria-label="MateAI home"><Image className="logo-image" src="/mateai-logo.png" alt="MateAI" width={52} height={52} priority unoptimized/></Link>;
}
