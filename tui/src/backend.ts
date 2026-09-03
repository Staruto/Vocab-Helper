import { defaultDbPath, EntryRow, WorkbookRow, VocabularyRepository } from "./db.js";

export class VocabularyBackend {
  private readonly repository: VocabularyRepository;

  constructor(dbPath: string = defaultDbPath()) {
    this.repository = new VocabularyRepository(dbPath);
  }

  listEntries(workbookId?: number): EntryRow[] {
    return this.repository.listEntries(workbookId);
  }

  listWorkbooks(): WorkbookRow[] {
    return this.repository.listWorkbooks();
  }

  getWorkbook(workbookId: number): WorkbookRow | null {
    return this.repository.getWorkbook(workbookId);
  }

  createWorkbook(name: string): WorkbookRow {
    return this.repository.createWorkbook(name);
  }

  deleteWorkbook(workbookId: number): number | null {
    return this.repository.deleteWorkbook(workbookId);
  }

  getCurrentWorkbookId(): number | null {
    return this.repository.getCurrentWorkbookId();
  }

  setCurrentWorkbookId(workbookId: number): WorkbookRow {
    return this.repository.setCurrentWorkbookId(workbookId);
  }

  countEntries(workbookId?: number): number {
    return this.repository.countEntries(workbookId);
  }

  getEntry(entryId: number): EntryRow | null {
    return this.repository.getEntry(entryId);
  }

  addEntry(workbookId: number, vocabulary: string, meaning: string): EntryRow {
    return this.repository.addEntry(workbookId, vocabulary, meaning);
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
