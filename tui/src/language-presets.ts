export type LanguagePresetDefinition = { attributes: Array<{ key: string; label: string; languageCode: string | null }>; partOfSpeechTags: string[] };

function exampleFields(languageCode: string): LanguagePresetDefinition["attributes"] {
  return [
    ...Array.from({ length: 5 }, (_, index) => ({ key: `example_sentence_${index + 1}`, label: `Example Sentence ${index + 1}`, languageCode })),
    { key: "common_collocations", label: "Common Collocations", languageCode },
    { key: "similar_words", label: "Similar Words", languageCode },
  ];
}

export const LANGUAGE_PRESET_DEFINITIONS: Record<string, LanguagePresetDefinition> = {
  JP: {
    attributes: [{ key: "kana", label: "読み方 (kana)", languageCode: "JP" }, ...exampleFields("JP")],
    partOfSpeechTags: ["名詞", "固有名詞", "イ形容詞", "ナ形容詞", "動詞 (自動詞)", "動詞 (他動詞)", "副詞", "連体詞", "接続詞", "連語", "その他"],
  },
  EN: { attributes: exampleFields("EN"), partOfSpeechTags: ["n.", "v.", "adj.", "adv.", "pron.", "prep.", "conj.", "phrase."] },
  DE: {
    attributes: exampleFields("DE"),
    partOfSpeechTags: ["m. noun - maskulines Substantiv", "f. noun - feminines Substantiv", "n. noun - neutrales Substantiv", "art. - Artikel", "adj. - Adjektiv", "pron. - Pronomen", "num. - Numerale", "adv. - Adverb", "prep. - Präposition", "conj. - Konjunktion", "interj. - Interjektion"],
  },
  ZH: { attributes: exampleFields("ZH"), partOfSpeechTags: [] },
  KO: { attributes: exampleFields("KO"), partOfSpeechTags: [] },
  ES: { attributes: exampleFields("ES"), partOfSpeechTags: [] },
  FR: { attributes: exampleFields("FR"), partOfSpeechTags: [] },
};

