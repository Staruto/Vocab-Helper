import { defaultDbPath, EntryRow, MeaningAttribute, MetadataAttribute, PosTag, WorkbookRow, VocabularyRepository } from "./db.js";

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
    presetEnabled = vocabularyLanguageCode === "JP",
  ): WorkbookRow {
    return this.repository.updateWorkbookSettings(workbookId, name, vocabularyLabel, vocabularyLanguageCode, meaningAttributes, presetEnabled);
  }

  listMeaningAttributes(workbookId: number): MeaningAttribute[] {
    return this.repository.listMeaningAttributes(workbookId);
  }

  createWorkbook(
    name: string,
    vocabularyLabel = "Vocabulary",
    vocabularyLanguageCode: string | null = null,
    meaningAttributes: MeaningAttribute[] = [{ position: 1, label: "Meaning 1", languageCode: null }],
    presetEnabled = vocabularyLanguageCode === "JP",
  ): WorkbookRow {
    return this.repository.createWorkbook(name, vocabularyLabel, vocabularyLanguageCode, meaningAttributes, presetEnabled);
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

  getTierColorsEnabled(): boolean { return this.repository.getTierColorsEnabled(); }
  setTierColorsEnabled(enabled: boolean): boolean { return this.repository.setTierColorsEnabled(enabled); }
  getEntryStats(entryId: number) { return this.repository.getEntryStats(entryId); }
  recordTestResult(entryId: number, isCorrect: boolean, decreaseError = true): EntryRow { return this.repository.recordTestResult(entryId, isCorrect, decreaseError); }
  selectPracticeCandidates(workbookId: number, count: number): EntryRow[] { return this.repository.selectPracticeCandidates(workbookId, count); }
  increasePriority(entryId: number): EntryRow { return this.repository.increasePriority(entryId); }
  decreasePriority(entryId: number): EntryRow { return this.repository.decreasePriority(entryId); }

  getEntry(entryId: number): EntryRow | null {
    return this.repository.getEntry(entryId);
  }

  addEntry(workbookId: number, vocabulary: string, meaning: string, meanings?: string[], attributes?: Record<string, string>, posTagIds?: number[]): EntryRow {
    return this.repository.addEntry(workbookId, vocabulary, meaning, meanings, attributes, posTagIds);
  }

  updateEntry(entryId: number, vocabulary: string, meaning: string, meanings?: string[], attributes?: Record<string, string>, posTagIds?: number[]): EntryRow {
    return this.repository.updateEntry(entryId, vocabulary, meaning, meanings, attributes, posTagIds);
  }

  deleteEntry(entryId: number): void {
    this.repository.deleteEntry(entryId);
  }

  listMetadataAttributes(workbookId: number): MetadataAttribute[] { return this.repository.listMetadataAttributes(workbookId); }
  updateMetadataAttributes(workbookId: number, attributes: MetadataAttribute[]): WorkbookRow { return this.repository.updateMetadataAttributes(workbookId, attributes); }
  listPosTags(workbookId: number): PosTag[] { return this.repository.listPosTags(workbookId); }
  addPosTag(workbookId: number, name: string): PosTag { return this.repository.addPosTag(workbookId, name); }
  renamePosTag(tagId: number, name: string): void { this.repository.renamePosTag(tagId, name); }
  deletePosTag(tagId: number): void { this.repository.deletePosTag(tagId); }
  setEntryPosTags(entryId: number, tagIds: number[]): void { this.repository.setEntryPosTags(entryId, tagIds); }

  close(): void {
    this.repository.close();
  }
}
