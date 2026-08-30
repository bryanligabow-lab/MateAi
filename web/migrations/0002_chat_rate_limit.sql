CREATE TABLE IF NOT EXISTS chat_requests (
  id TEXT PRIMARY KEY,
  ip_hash TEXT NOT NULL,
  created_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_chat_requests_rate ON chat_requests(ip_hash, created_at);
