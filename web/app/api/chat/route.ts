import { env } from "cloudflare:workers";
import { z } from "zod";

const messageSchema = z.object({
  role: z.enum(["user", "assistant"]),
  content: z.string().trim().min(1).max(2_000),
});
const requestSchema = z.object({
  messages: z.array(messageSchema).min(1).max(12),
  language: z.enum(["en", "es"]).default("en"),
});

const systemPrompts = {
  en: "You are MateAI, a helpful general-purpose AI assistant and guide to MateAI services. Answer in the user's language in at most 160 words. For website chatbot installation, explain MateAI's usual process: discovery, knowledge and behavior configuration, adding a small widget script, connecting requested channels or systems, testing, and launch. Do not invent facts; acknowledge uncertainty. Never claim to have performed actions you cannot perform.",
  es: "Eres MateAI, un asistente de IA de propósito general y guía de los servicios de MateAI. Responde en el idioma del usuario y usa como máximo 160 palabras. Para la instalación de chatbots web, explica el proceso habitual de MateAI: descubrimiento, configuración del conocimiento y comportamiento, instalación de un pequeño script, conexión de canales o sistemas solicitados, pruebas y lanzamiento. No inventes hechos; reconoce la incertidumbre. Nunca afirmes haber realizado acciones que no puedes realizar.",
} as const;

function jsonError(error: string, status: number) {
  return Response.json({ error }, { status });
}

export async function POST(request: Request) {
  const contentLength = Number(request.headers.get("content-length") ?? 0);
  if (contentLength > 30_000) return jsonError("Request too large", 413);

  const parsed = requestSchema.safeParse(await request.json().catch(() => null));
  if (!parsed.success) return jsonError("Invalid chat request", 400);

  const ip = request.headers.get("cf-connecting-ip") ?? "unknown";
  const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(ip));
  const ipHash = [...new Uint8Array(digest)].map((byte) => byte.toString(16).padStart(2, "0")).join("");
  const recent = await env.DB.prepare(
    "SELECT COUNT(*) count FROM chat_requests WHERE ip_hash=? AND created_at > datetime('now','-15 minutes')",
  ).bind(ipHash).first<{ count: number }>();
  if ((recent?.count ?? 0) >= 20) return jsonError("Too many requests. Please try again later.", 429);

  await env.DB.prepare(
    "INSERT INTO chat_requests (id,ip_hash,created_at) VALUES (?,?,datetime('now'))",
  ).bind(crypto.randomUUID(), ipHash).run();

  if (!env.OLLAMA_BASIC_AUTH) return jsonError("Chat is not configured", 503);
  const response = await fetch(`${env.OLLAMA_BASE_URL}/api/chat`, {
    method: "POST",
    headers: {
      Authorization: `Basic ${env.OLLAMA_BASIC_AUTH}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      model: env.OLLAMA_MODEL,
      stream: true,
      messages: [
        { role: "system", content: systemPrompts[parsed.data.language] },
        ...parsed.data.messages,
      ],
      keep_alive: "15m",
      options: { temperature: 0.4, num_predict: 220, num_ctx: 2_048 },
    }),
    signal: AbortSignal.timeout(120_000),
  }).catch(() => null);

  if (!response?.ok) return jsonError("The AI is temporarily unavailable", 502);
  if (!response.body) return jsonError("The AI returned an empty response", 502);
  return new Response(response.body, { headers: { "Cache-Control": "no-store", "Content-Type": "application/x-ndjson; charset=utf-8" } });
}
