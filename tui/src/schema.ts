import { DatabaseSync } from "node:sqlite";

export const CURRENT_SCHEMA_VERSION = 3;

export type SchemaMigration = {
  version: number;
  apply: (db: DatabaseSync) => void;
};

function tableExists(db: DatabaseSync, name: string): boolean {
  return Boolean(db.prepare("SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?").get(name));
}

export function isHybridOrLegacyDatabase(db: DatabaseSync): boolean {
  return tableExists(db, "mvp_workbooks") || tableExists(db, "vocab_entries");
}

export const SCHEMA_MIGRATIONS: SchemaMigration[] = [
  {
    version: 1,
    apply(db) {
      db.exec(`
        CREATE TABLE app_settings (
          singleton_id INTEGER PRIMARY KEY CHECK (singleton_id = 1),
          current_workbook_id INTEGER NULL,
          tier_colors_enabled INTEGER NOT NULL DEFAULT 1 CHECK (tier_colors_enabled IN (0, 1)),
          FOREIGN KEY (current_workbook_id) REFERENCES workbooks(id) ON DELETE SET NULL
        );

        CREATE TABLE workbooks (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          name TEXT NOT NULL COLLATE NOCASE UNIQUE CHECK (trim(name) <> ''),
          vocabulary_kind TEXT NOT NULL CHECK (vocabulary_kind IN ('preset_language', 'other_language', 'non_language')),
          vocabulary_label TEXT NOT NULL CHECK (trim(vocabulary_label) <> ''),
          vocabulary_language_code TEXT NULL,
          preset_enabled INTEGER NOT NULL DEFAULT 0 CHECK (preset_enabled IN (0, 1)),
          pos_enabled INTEGER NOT NULL DEFAULT 0 CHECK (pos_enabled IN (0, 1)),
          created_at TEXT NOT NULL,
          updated_at TEXT NOT NULL,
          CHECK (
            (vocabulary_kind = 'preset_language' AND vocabulary_language_code IS NOT NULL)
            OR (vocabulary_kind <> 'preset_language' AND vocabulary_language_code IS NULL)
          )
        );

        CREATE TABLE workbook_fields (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          workbook_id INTEGER NOT NULL,
          field_key TEXT NOT NULL CHECK (trim(field_key) <> ''),
          role TEXT NOT NULL CHECK (role IN ('meaning', 'optional')),
          position INTEGER NOT NULL CHECK (position >= 1),
          label TEXT NOT NULL CHECK (trim(label) <> ''),
          language_code TEXT NULL,
          is_required INTEGER NOT NULL DEFAULT 0 CHECK (is_required IN (0, 1)),
          is_visible INTEGER NOT NULL DEFAULT 0 CHECK (is_visible IN (0, 1)),
          provenance TEXT NOT NULL DEFAULT 'custom' CHECK (provenance IN ('preset', 'custom')),
          UNIQUE (workbook_id, field_key),
          UNIQUE (workbook_id, role, position),
          UNIQUE (id, workbook_id),
          FOREIGN KEY (workbook_id) REFERENCES workbooks(id) ON DELETE CASCADE
        );

        CREATE TABLE entries (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          workbook_id INTEGER NOT NULL,
          vocabulary TEXT NOT NULL CHECK (trim(vocabulary) <> ''),
          created_at TEXT NOT NULL,
          updated_at TEXT NOT NULL,
          UNIQUE (id, workbook_id),
          FOREIGN KEY (workbook_id) REFERENCES workbooks(id) ON DELETE CASCADE
        );

        CREATE TABLE entry_field_values (
          entry_id INTEGER NOT NULL,
          field_id INTEGER NOT NULL,
          workbook_id INTEGER NOT NULL,
          value TEXT NOT NULL DEFAULT '',
          PRIMARY KEY (entry_id, field_id),
          FOREIGN KEY (entry_id, workbook_id) REFERENCES entries(id, workbook_id) ON DELETE CASCADE,
          FOREIGN KEY (field_id, workbook_id) REFERENCES workbook_fields(id, workbook_id) ON DELETE CASCADE
        );

        CREATE TABLE pos_tags (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          workbook_id INTEGER NOT NULL,
          name TEXT NOT NULL COLLATE NOCASE CHECK (trim(name) <> ''),
          is_predefined INTEGER NOT NULL DEFAULT 0 CHECK (is_predefined IN (0, 1)),
          UNIQUE (workbook_id, name),
          UNIQUE (id, workbook_id),
          FOREIGN KEY (workbook_id) REFERENCES workbooks(id) ON DELETE CASCADE
        );

        CREATE TABLE entry_pos_tags (
          entry_id INTEGER NOT NULL,
          tag_id INTEGER NOT NULL,
          workbook_id INTEGER NOT NULL,
          PRIMARY KEY (entry_id, tag_id),
          FOREIGN KEY (entry_id, workbook_id) REFERENCES entries(id, workbook_id) ON DELETE CASCADE,
          FOREIGN KEY (tag_id, workbook_id) REFERENCES pos_tags(id, workbook_id) ON DELETE CASCADE
        );

        CREATE TABLE entry_stats (
          entry_id INTEGER PRIMARY KEY,
          test_count INTEGER NOT NULL DEFAULT 0 CHECK (test_count >= 0),
          error_count INTEGER NOT NULL DEFAULT 0 CHECK (error_count BETWEEN 0 AND 3),
          last_tested TEXT NULL,
          next_test_deadline TEXT NULL,
          FOREIGN KEY (entry_id) REFERENCES entries(id) ON DELETE CASCADE
        );

        CREATE INDEX idx_entries_workbook_id ON entries(workbook_id, id);
        CREATE INDEX idx_fields_workbook_order ON workbook_fields(workbook_id, role, position);
        CREATE INDEX idx_entry_values_workbook ON entry_field_values(workbook_id, entry_id);
        CREATE INDEX idx_entry_pos_workbook ON entry_pos_tags(workbook_id, entry_id);

        INSERT INTO app_settings (singleton_id, current_workbook_id, tier_colors_enabled)
        VALUES (1, NULL, 1);
      `);
    },
  },
  {
    version: 2,
    apply(db) {
      db.exec(`
        CREATE TABLE tag_types (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          workbook_id INTEGER NOT NULL,
          name TEXT NOT NULL COLLATE NOCASE CHECK (trim(name) <> ''),
          position INTEGER NOT NULL CHECK (position >= 1),
          UNIQUE (workbook_id, name),
          UNIQUE (workbook_id, position),
          UNIQUE (id, workbook_id),
          FOREIGN KEY (workbook_id) REFERENCES workbooks(id) ON DELETE CASCADE
        );

        CREATE TABLE tags (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          tag_type_id INTEGER NOT NULL,
          workbook_id INTEGER NOT NULL,
          name TEXT NOT NULL COLLATE NOCASE CHECK (trim(name) <> ''),
          UNIQUE (tag_type_id, name),
          UNIQUE (id, workbook_id),
          FOREIGN KEY (tag_type_id, workbook_id) REFERENCES tag_types(id, workbook_id) ON DELETE CASCADE
        );

        CREATE TABLE entry_tags (
          entry_id INTEGER NOT NULL,
          tag_id INTEGER NOT NULL,
          workbook_id INTEGER NOT NULL,
          PRIMARY KEY (entry_id, tag_id),
          FOREIGN KEY (entry_id, workbook_id) REFERENCES entries(id, workbook_id) ON DELETE CASCADE,
          FOREIGN KEY (tag_id, workbook_id) REFERENCES tags(id, workbook_id) ON DELETE CASCADE
        );
      `);

      const workbooks = db.prepare(`SELECT w.id, w.pos_enabled,
        EXISTS(SELECT 1 FROM pos_tags p WHERE p.workbook_id = w.id) AS has_tags
        FROM workbooks w ORDER BY w.id`).all() as Array<{ id: number; pos_enabled: number; has_tags: number }>;
      const addType = db.prepare("INSERT INTO tag_types (workbook_id, name, position) VALUES (?, 'Part of Speech', 1)");
      const copyTags = db.prepare("INSERT INTO tags (id, tag_type_id, workbook_id, name) SELECT id, ?, workbook_id, name FROM pos_tags WHERE workbook_id = ? ORDER BY id");
      for (const workbook of workbooks) {
        if (!Number(workbook.pos_enabled) && !Number(workbook.has_tags)) continue;
        const result = addType.run(Number(workbook.id));
        copyTags.run(Number(result.lastInsertRowid), Number(workbook.id));
      }
      db.exec(`
        INSERT INTO entry_tags (entry_id, tag_id, workbook_id)
        SELECT entry_id, tag_id, workbook_id FROM entry_pos_tags;

        DROP TABLE entry_pos_tags;
        DROP TABLE pos_tags;
        ALTER TABLE workbooks DROP COLUMN pos_enabled;

        CREATE INDEX idx_tag_types_workbook_order ON tag_types(workbook_id, position);
        CREATE INDEX idx_tags_type_name ON tags(tag_type_id, name);
        CREATE INDEX idx_entry_tags_workbook ON entry_tags(workbook_id, entry_id);
      `);
    },
  },
  {
    version: 3,
    apply(db) {
      db.exec(`
        ALTER TABLE tag_types
        ADD COLUMN is_visible INTEGER NOT NULL DEFAULT 0 CHECK (is_visible IN (0, 1));
      `);
    },
  },
];

export function runSchemaMigrations(db: DatabaseSync, migrations: SchemaMigration[] = SCHEMA_MIGRATIONS): void {
  db.exec("PRAGMA foreign_keys = ON");
  if (isHybridOrLegacyDatabase(db)) {
    throw new Error("This database uses the legacy/hybrid schema. Run 'npm run db:convert:apply' before starting VocabHelper.");
  }

  db.exec("CREATE TABLE IF NOT EXISTS schema_migrations (version INTEGER PRIMARY KEY, applied_at TEXT NOT NULL)");
  const applied = new Set((db.prepare("SELECT version FROM schema_migrations").all() as Array<{ version: number }>).map((row) => Number(row.version)));
  const ordered = [...migrations].sort((a, b) => a.version - b.version);
  for (const migration of ordered) {
    if (applied.has(migration.version)) continue;
    db.exec("BEGIN IMMEDIATE");
    try {
      migration.apply(db);
      db.prepare("INSERT INTO schema_migrations (version, applied_at) VALUES (?, ?)").run(migration.version, new Date().toISOString());
      db.exec("COMMIT");
    } catch (error) {
      try { db.exec("ROLLBACK"); } catch { /* Preserve the migration error. */ }
      throw new Error(`Database migration ${migration.version} failed: ${error instanceof Error ? error.message : String(error)}`);
    }
  }
}

export function assertDatabaseIntegrity(db: DatabaseSync): void {
  const integrity = db.prepare("PRAGMA integrity_check").get() as { integrity_check?: string } | undefined;
  if (integrity?.integrity_check !== "ok") throw new Error(`SQLite integrity check failed: ${integrity?.integrity_check ?? "unknown error"}`);
  const foreignKeys = db.prepare("PRAGMA foreign_key_check").all();
  if (foreignKeys.length > 0) throw new Error(`SQLite foreign-key check failed with ${foreignKeys.length} violation(s).`);
}
