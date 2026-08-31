export type ClientChatMessage = { role: "user" | "assistant"; content: string };

export async function streamChat(messages: ClientChatMessage[], onText: (text: string) => void) {
  const response = await fetch("/api/chat", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ messages: messages.slice(-12), language: "en" }),
  });
  if (!response.ok || !response.body) {
    const data = await response.json().catch(() => ({})) as { error?: string };
    throw new Error(data.error ?? "The AI is temporarily unavailable");
  }

  const reader = response.body.pipeThrough(new TextDecoderStream()).getReader();
  let buffer = "";
  while (true) {
    const { value, done } = await reader.read();
    if (done) break;
    buffer += value;
    const lines = buffer.split("\n");
    buffer = lines.pop() ?? "";
    for (const line of lines) {
      if (!line.trim()) continue;
      const event = JSON.parse(line) as { message?: { content?: string }; error?: string };
      if (event.error) throw new Error(event.error);
      if (event.message?.content) onText(event.message.content);
    }
  }
  if (buffer.trim()) {
    const event = JSON.parse(buffer) as { message?: { content?: string }; error?: string };
    if (event.error) throw new Error(event.error);
    if (event.message?.content) onText(event.message.content);
  }
}
