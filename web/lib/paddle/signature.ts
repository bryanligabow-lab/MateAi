export async function verifyPaddleSignature(rawBody: string, signature: string | null, secret: string) {
  if (!signature || !secret) return false;
  const parts = Object.fromEntries(signature.split(";").map((part) => part.split("=") as [string,string]));
  const timestamp = parts.ts;
  const expected = parts.h1;
  if (!timestamp || !expected || Math.abs(Date.now() / 1000 - Number(timestamp)) > 5) return false;
  const key = await crypto.subtle.importKey("raw", new TextEncoder().encode(secret), { name:"HMAC", hash:"SHA-256" }, false, ["sign"]);
  const digest = await crypto.subtle.sign("HMAC", key, new TextEncoder().encode(`${timestamp}:${rawBody}`));
  const actual = [...new Uint8Array(digest)].map((byte) => byte.toString(16).padStart(2,"0")).join("");
  if (actual.length !== expected.length) return false;
  let mismatch = 0;
  for (let index=0; index<actual.length; index++) mismatch |= actual.charCodeAt(index) ^ expected.charCodeAt(index);
  return mismatch === 0;
}
