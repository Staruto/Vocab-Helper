import { defaultDbPath, EntryRow, MeaningAttribute, WorkbookRow, VocabularyRepository } from "./db.js";

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

  updateWorkbookSettings(
    workbookId: number,
    name: string,
    vocabularyLabel: string,
    vocabularyLanguageCode: string | null,
    meaningAttributes: MeaningAttribute[],
  ): WorkbookRow {
    return this.repository.updateWorkbookSettings(workbookId, name, vocabularyLabel, vocabularyLanguageCode, meaningAttributes);
  }

  listMeaningAttributes(workbookId: number): MeaningAttribute[] {
    return this.repository.listMeaningAttributes(workbookId);
  }

  createWorkbook(
    name: string,
    vocabularyLabel = "Vocabulary",
    vocabularyLanguageCode: string | null = null,
    meaningAttributes: MeaningAttribute[] = [{ position: 1, label: "Meaning 1", languageCode: null }],
  ): WorkbookRow {
    return this.repository.createWorkbook(name, vocabularyLabel, vocabularyLanguageCode, meaningAttributes);
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

  addEntry(workbookId: number, vocabulary: string, meaning: string, meanings?: string[]): EntryRow {
    return this.repository.addEntry(workbookId, vocabulary, meaning, meanings);
  }

  updateEntry(entryId: number, vocabulary: string, meaning: string, meanings?: string[]): EntryRow {
    return this.repository.updateEntry(entryId, vocabulary, meaning, meanings);
  }

  deleteEntry(entryId: number): void {
    this.repository.deleteEntry(entryId);
  }

  close(): void {
    this.repository.close();
  }
}
