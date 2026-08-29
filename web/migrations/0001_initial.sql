CREATE TABLE IF NOT EXISTS webhook_events (event_id TEXT PRIMARY KEY,event_type TEXT NOT NULL,occurred_at TEXT NOT NULL,processed_at TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS subscriptions (paddle_subscription_id TEXT PRIMARY KEY,paddle_customer_id TEXT,plan TEXT NOT NULL,status TEXT NOT NULL,business_name TEXT,email TEXT,created_at TEXT DEFAULT CURRENT_TIMESTAMP,updated_at TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS provisioning_jobs (event_id TEXT PRIMARY KEY,transaction_id TEXT NOT NULL,status TEXT NOT NULL,created_at TEXT NOT NULL,completed_at TEXT);
CREATE TABLE IF NOT EXISTS lead_submissions (id TEXT PRIMARY KEY,first_name TEXT NOT NULL,last_name TEXT NOT NULL,business_name TEXT NOT NULL,website TEXT NOT NULL,email TEXT NOT NULL,phone TEXT NOT NULL,industry TEXT NOT NULL,ip_hash TEXT NOT NULL,created_at TEXT NOT NULL);
CREATE INDEX IF NOT EXISTS idx_lead_submissions_rate ON lead_submissions(ip_hash,created_at);
