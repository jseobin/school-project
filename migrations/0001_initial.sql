CREATE TABLE IF NOT EXISTS guestbook_entries (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  author TEXT NOT NULL,
  content TEXT NOT NULL,
  created_at TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_guestbook_entries_created_at
ON guestbook_entries(created_at DESC);
