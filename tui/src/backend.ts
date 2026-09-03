import { defaultDbPath, EntryRow, VocabularyRepository } from "./db.js";

export class VocabularyBackend {
  private readonly repository: VocabularyRepository;

  constructor(dbPath: string = defaultDbPath()) {
    this.repository = new VocabularyRepository(dbPath);
  }

  listEntries(): EntryRow[] {
    return this.repository.listEntries();
  }

  countEntries(): number {
    return this.repository.countEntries();
  }

  getEntry(entryId: number): EntryRow | null {
    return this.repository.getEntry(entryId);
  }

  addEntry(vocabulary: string, meaning: string): EntryRow {
    return this.repository.addEntry(vocabulary, meaning);
  }

  updateEntry(entryId: number, vocabulary: string, meaning: string): EntryRow {
    return this.repository.updateEntry(entryId, vocabulary, meaning);
  }

  deleteEntry(entryId: number): void {
    this.repository.deleteEntry(entryId);
  }

  close(): void {
    this.repository.close();
  }
}
