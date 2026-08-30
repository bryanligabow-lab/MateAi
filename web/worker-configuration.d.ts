declare namespace Cloudflare {
  interface Env {
    DB: D1Database;
    PADDLE_WEBHOOK_SECRET?: string;
    PADDLE_API_KEY?: string;
    MATEAI_BACKEND_URL?: string;
    MATEAI_BACKEND_API_KEY?: string;
    OLLAMA_BASE_URL: string;
    OLLAMA_MODEL: string;
    OLLAMA_BASIC_AUTH?: string;
  }
}
