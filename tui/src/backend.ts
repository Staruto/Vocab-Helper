import { CreateWorkbookInput, defaultDbPath, EntryRow, MeaningAttribute, MeaningPromotionImpact, MetadataAttribute, TagType, TagUpdateImpact, WorkbookAttributesDraft, WorkbookConfigurationInput, WorkbookRow, WorkbookTagsDraft, WorkbookUpdateImpact, VocabularyRepository } from "./db.js";

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
    meaningAttributes: MeaningAttribute[] = [{ position: 1, label: "Primary Meaning", languageCode: null }],
    presetEnabled = vocabularyLanguageCode === "JP",
  ): WorkbookRow {
    return this.repository.createWorkbook(name, vocabularyLabel, vocabularyLanguageCode, meaningAttributes, presetEnabled);
  }

  createConfiguredWorkbook(input: CreateWorkbookInput): WorkbookRow { return this.repository.createConfiguredWorkbook(input); }
  previewWorkbookUpdate(workbookId: number, input: WorkbookConfigurationInput): WorkbookUpdateImpact { return this.repository.previewWorkbookUpdate(workbookId, input); }
  updateConfiguredWorkbook(workbookId: number, input: WorkbookConfigurationInput, confirmDataLoss = false): WorkbookRow { return this.repository.updateConfiguredWorkbook(workbookId, input, confirmDataLoss); }

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

  addEntry(workbookId: number, vocabulary: string, meaning: string, meanings?: string[], attributes?: Record<string, string>, tagIds?: number[]): EntryRow {
    return this.repository.addEntry(workbookId, vocabulary, meaning, meanings, attributes, tagIds);
  }

  updateEntry(entryId: number, vocabulary: string, meaning: string, meanings?: string[], attributes?: Record<string, string>, tagIds?: number[]): EntryRow {
    return this.repository.updateEntry(entryId, vocabulary, meaning, meanings, attributes, tagIds);
  }

  deleteEntry(entryId: number): void {
    this.repository.deleteEntry(entryId);
  }

  listMetadataAttributes(workbookId: number): MetadataAttribute[] { return this.repository.listMetadataAttributes(workbookId); }
  getMeaningPromotionImpact(workbookId: number, fieldId: number): MeaningPromotionImpact { return this.repository.getMeaningPromotionImpact(workbookId, fieldId); }
  previewWorkbookAttributesUpdate(workbookId: number, draft: WorkbookAttributesDraft): WorkbookUpdateImpact { return this.repository.previewWorkbookAttributesUpdate(workbookId, draft); }
  updateWorkbookAttributes(workbookId: number, draft: WorkbookAttributesDraft, confirmDataLoss = false): WorkbookRow { return this.repository.updateWorkbookAttributes(workbookId, draft, confirmDataLoss); }
  listTagTypes(workbookId: number): TagType[] { return this.repository.listTagTypes(workbookId); }
  previewWorkbookTagsUpdate(workbookId: number, draft: WorkbookTagsDraft): TagUpdateImpact { return this.repository.previewWorkbookTagsUpdate(workbookId, draft); }
  updateWorkbookTags(workbookId: number, draft: WorkbookTagsDraft, confirmDataLoss = false): TagType[] { return this.repository.updateWorkbookTags(workbookId, draft, confirmDataLoss); }

  close(): void {
    this.repository.close();
  }
}
