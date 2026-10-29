import type { MetadataAttribute } from "./db.js";
import { LANGUAGE_PRESET_DEFINITIONS } from "./language-presets.js";

export function createPresetAttributes(languageCode: string | null): MetadataAttribute[] {
  return (languageCode ? LANGUAGE_PRESET_DEFINITIONS[languageCode]?.attributes ?? [] : []).map((field, index) => ({
    ...field, role: "optional", presetKey: field.key, active: true,
    required: false, visible: false, displayOrder: index + 1, provenance: "preset",
  }));
}

export function isActiveAttribute(field: MetadataAttribute): boolean {
  return !field.presetKey || field.active !== false;
}

export function entryAttributes(fields: MetadataAttribute[]): MetadataAttribute[] {
  return fields.filter((field) => field.role === "optional" && isActiveAttribute(field));
}

export function isPresetActivationInput(input: string, key: { ctrl: boolean; meta?: boolean }): boolean {
  return !key.ctrl && !key.meta && input.toLowerCase() === "t";
}
