CREATE TABLE IF NOT EXISTS ledger (id INTEGER PRIMARY KEY CHECK(id=1), revision INTEGER NOT NULL, data TEXT NOT NULL);
INSERT OR IGNORE INTO ledger VALUES(1,0,'{"revision":0,"records":[],"expenses":[],"purchases":[],"history":[]}');
CREATE TABLE IF NOT EXISTS imports (id TEXT PRIMARY KEY, filename TEXT NOT NULL, imported_at TEXT NOT NULL, group_count INTEGER NOT NULL, data TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS source_chunks (id TEXT NOT NULL, part INTEGER NOT NULL, data BLOB NOT NULL, PRIMARY KEY(id,part));
CREATE TABLE IF NOT EXISTS login_attempts (key TEXT PRIMARY KEY, expires INTEGER NOT NULL, attempts INTEGER NOT NULL);
