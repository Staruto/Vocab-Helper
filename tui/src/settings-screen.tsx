import React, { useEffect, useMemo, useRef, useState } from "react";
import { Box, Text, useInput, useStdout } from "ink";
import { VocabularyBackend } from "./backend.js";
import { MetadataAttribute, TagDataLossError, TagTypeDraft, WorkbookDataLossError, WorkbookRow, WorkbookSettingsChanges } from "./db.js";
import { loadLabeledTextFile } from "./import.js";
import { SETTINGS_SECTIONS, settingsTabWindow, SettingsSection } from "./settings-navigation.js";
import { CaretInputLine } from "./text-input.js";
import { parseDailyGoalInput } from "./workbook-display.js";
import truncate from "cli-truncate";
import displayWidth from "string-width";

export type SettingsPanelControls = {
  focused: boolean;
  onTop: () => void;
  setNavigationHint: React.Dispatch<React.SetStateAction<string>>;
};
export type AttributeDraftControls = {
  vocabularyLabel: string; setVocabularyLabel: React.Dispatch<React.SetStateAction<string>>;
  fields: MetadataAttribute[]; setFields: React.Dispatch<React.SetStateAction<MetadataAttribute[]>>;
};
export type TagDraftControls = { types: TagTypeDraft[]; setTypes: React.Dispatch<React.SetStateAction<TagTypeDraft[]>> };
type Props = {
  workbook: WorkbookRow; backend: VocabularyBackend; onCancel: () => void; onQuit: () => void;
  renderAttributes: (panel: SettingsPanelControls, draft: AttributeDraftControls, leave: () => void) => JSX.Element;
  renderTags: (panel: SettingsPanelControls, draft: TagDraftControls, leave: () => void) => JSX.Element;
};
const selectedColor = "#c5a8ff";
const auxiliary = "#8a8a8a";

function padLine(text: string, width: number): string {
  const clipped = truncate(text, width);
  return `${clipped}${" ".repeat(Math.max(0, width - displayWidth(clipped)))}`;
}

export function SettingsScreen({ workbook, backend, onCancel, onQuit, renderAttributes, renderTags }: Props): JSX.Element {
  const { stdout } = useStdout();
  const [width, setWidth] = useState(stdout?.columns ?? 80);
  const [sectionIndex, setSectionIndex] = useState(0);
  const section = SETTINGS_SECTIONS[sectionIndex];
  const [focus, setFocus] = useState<"sections" | "content">("sections");
  const [mode, setMode] = useState<"edit" | "exit" | "destructive">("edit");
  const [choice, setChoice] = useState(0);
  const [confirm, setConfirm] = useState("");
  const [warnings, setWarnings] = useState<string[]>([]);
  const [message, setMessage] = useState("");
  const [saving, setSaving] = useState(false);
  const [panelNavigationHint, setPanelNavigationHint] = useState("");
  const request = useRef(0);
  const initialFields = useMemo(() => workbook.metadataAttributes.filter((field) => field.role !== "vocabulary"), [workbook]);
  const initialTypes = useMemo<TagTypeDraft[]>(() => backend.listTagTypes(workbook.id).map((type) => ({ id: type.id, name: type.name, visible: type.visible, tags: type.tags.map((tag) => ({ id: tag.id, name: tag.name })) })), [backend, workbook.id]);
  const [vocabularyLabel, setVocabularyLabel] = useState(workbook.vocabularyLabel);
  const [fields, setFields] = useState(initialFields);
  const [types, setTypes] = useState(initialTypes);
  const [baseline, setBaseline] = useState(workbook);
  const [importPath, setImportPath] = useState(workbook.importFilePath ?? "");
  const [goal, setGoal] = useState(workbook.dailyGoal?.toString() ?? "");
  const [appearanceIndex, setAppearanceIndex] = useState(0);
  const [tierColors, setTierColors] = useState(() => backend.getTierColorsEnabled());
  const [progress, setProgress] = useState(() => backend.getWorkbookProgressEnabled());
  const attributeDirty = vocabularyLabel !== workbook.vocabularyLabel || JSON.stringify(fields) !== JSON.stringify(initialFields);
  const tagDirty = JSON.stringify(types) !== JSON.stringify(initialTypes);
  const pathDirty = importPath !== (baseline.importFilePath ?? "");
  const goalDirty = goal !== (baseline.dailyGoal?.toString() ?? "");
  const dirty = attributeDirty || tagDirty || pathDirty || goalDirty;
  const contentFocused = focus === "content" && mode === "edit" && !saving;
  useEffect(() => {
    if (!stdout) return;
    const resize = () => setWidth(stdout.columns ?? 80);
    stdout.on("resize", resize);
    return () => { stdout.off("resize", resize); };
  }, [stdout]);
  useEffect(() => () => { request.current += 1; }, []);

  function leave(): void {
    if (dirty) { setChoice(0); setMode("exit"); setMessage(""); }
    else onCancel();
  }
  async function save(all: boolean, confirmDataLoss = false): Promise<void> {
    const currentRequest = ++request.current;
    const changes: WorkbookSettingsChanges = {};
    let target: SettingsSection = section;
    setSaving(true); setMessage("");
    try {
      const removalWarnings: string[] = [];
      if (all && attributeDirty) {
        target = "Attributes"; changes.attributes = { vocabularyLabel, fields };
        removalWarnings.push(...backend.previewWorkbookAttributesUpdate(workbook.id, changes.attributes).populatedFields.map((field) => `${field.label}: ${field.valueCount} populated value(s)`));
      }
      if (all && tagDirty) {
        target = "Tags"; changes.tags = { types };
        removalWarnings.push(...backend.previewWorkbookTagsUpdate(workbook.id, changes.tags).removals.map((item) => `${item.typeName}${item.tagName ? ` / ${item.tagName}` : ""}: ${item.assignmentCount} assignment(s) across ${item.entryCount} entry/entries`));
      }
      if ((all && goalDirty) || (!all && section === "Daily Goal")) {
        target = "Daily Goal"; changes.dailyGoal = parseDailyGoalInput(goal);
      }
      if ((all && pathDirty) || (!all && section === "Import")) {
        target = "Import";
        changes.importFilePath = importPath.trim() ? (await loadLabeledTextFile(importPath.trim())).path : null;
      }
      if (currentRequest !== request.current) return;
      if (removalWarnings.length && !confirmDataLoss) {
        setWarnings(removalWarnings); setConfirm(""); setMode("destructive"); return;
      }
      const updated = backend.updateWorkbookSettingsSections(workbook.id, changes, confirmDataLoss);
      if (all) { onCancel(); return; }
      setBaseline(updated);
      if (section === "Import") setImportPath(updated.importFilePath ?? "");
      else setGoal(updated.dailyGoal?.toString() ?? "");
      setMessage("Saved.");
    } catch (error) {
      if (currentRequest !== request.current) return;
      if (error instanceof WorkbookDataLossError) target = "Attributes";
      if (error instanceof TagDataLossError) target = "Tags";
      setSectionIndex(SETTINGS_SECTIONS.indexOf(target)); setFocus("content"); setMode("edit");
      setMessage(error instanceof Error ? error.message : "Could not save settings.");
    } finally { if (currentRequest === request.current) setSaving(false); }
  }
  useInput((input, key) => {
    if (key.ctrl && input === "c") { request.current += 1; onQuit(); return; }
    if (saving) {
      if (key.escape) { request.current += 1; setSaving(false); setMode("edit"); }
      return;
    }
    if (mode === "exit") {
      if (key.escape) { setMode("edit"); return; }
      if (key.upArrow) setChoice((value) => (value + 2) % 3);
      else if (key.downArrow) setChoice((value) => (value + 1) % 3);
      else if (key.return) { if (choice === 0) void save(true); else if (choice === 1) onCancel(); else setMode("edit"); }
      return;
    }
    if (mode === "destructive") {
      if (key.escape) { setMode("edit"); setConfirm(""); return; }
      if (key.return) { if (confirm.trim().toLowerCase() === "yes") void save(true, true); else setMessage("Type yes to confirm data removal."); }
      return;
    }
    if (focus === "sections") {
      if (key.escape) leave();
      else if (key.leftArrow || key.rightArrow) {
        setSectionIndex((value) => (value + (key.leftArrow ? SETTINGS_SECTIONS.length - 1 : 1)) % SETTINGS_SECTIONS.length);
        setAppearanceIndex(0); setMessage("");
      } else if (key.downArrow || key.return) setFocus("content");
      return;
    }
    if (section === "Attributes" || section === "Tags") return;
    if (key.escape) { leave(); return; }
    if (key.upArrow) {
      if (section === "Appearance" && appearanceIndex > 0) setAppearanceIndex((value) => value - 1);
      else setFocus("sections");
      return;
    }
    if (section === "Appearance") {
      if (key.downArrow) setAppearanceIndex(1);
      else if (input === " ") {
        if (appearanceIndex === 0) setTierColors(backend.setTierColorsEnabled(!tierColors));
        else setProgress(backend.setWorkbookProgressEnabled(!progress));
      }
    } else {
      if (key.ctrl && input === "u") { section === "Import" ? setImportPath("") : setGoal(""); setMessage(""); }
      else if (key.return) void save(false);
    }
  });
  const panel: SettingsPanelControls = { focused: contentFocused, onTop: () => setFocus("sections"), setNavigationHint: setPanelNavigationHint };
  const tabs = settingsTabWindow(width, sectionIndex);

  const navigationHints: string[] = [];
  if (focus === "sections") {
    navigationHints.push("←/→ sections", "↓/Enter settings", "Esc leave");
  }
  else if (focus === "content" && panelNavigationHint) {
    navigationHints.push(panelNavigationHint);
  } else {
    if (section === "Import" || section === "Daily Goal") {
      navigationHints.push("Enter saves", "Ctrl+U clears");
    }
    if (section === "Appearance") {
      navigationHints.push("↑↓ select", "Space toggles");
    }
    navigationHints.push("Esc leave");
  }

  return <Box flexDirection="column" width={width}>
    <Box flexDirection="column" display={mode === "edit" ? "none" : "flex"}>
      <Box justifyContent='center'><Text color="cyan" bold>{mode === "exit" ? "Unsaved settings changes" : "Confirm data removal"}</Text></Box>
      <Text color={auxiliary}>Choose what to do with your staged changes.</Text>
      <Text>{" "}</Text>
    {mode === "exit" ? ["Save changes", "Discard changes", "Continue editing"].map((label, index) => <Text key={label} inverse={choice === index}>{padLine(`${index === choice ? ">" : " "} ${label}`, width)}</Text>) : <>
      <Text color={auxiliary}>Saving will permanently remove:</Text>
      {warnings.map((line, index) => <Text key={index} color="red" wrap="truncate-end">{line}</Text>)}
      <CaretInputLine prefix="Type yes: " value={confirm} onChange={setConfirm} focus={mode === "destructive" && !saving} width={width} inputKey="settings-confirm" />
    </>}
    <Text color="red" wrap="truncate-end">{message || " "}</Text>
    <Text color={auxiliary} wrap="truncate-end">{saving ? "Validating settings... Esc cancels" : mode === "exit" ? "↑↓ select | Enter confirm | Esc continue editing" : "Enter confirm | Esc continue editing"}</Text>
    </Box>
    <Box flexDirection="column" display={mode === "edit" ? "flex" : "none"}>
    <Box justifyContent="center"><Text color="cyan" bold wrap="truncate-end">{`Settings - ${workbook.name}`}</Text></Box>
    <Text color={selectedColor} wrap="truncate-end">{"─".repeat(width)}</Text>
    <Text wrap="truncate-end"><Text color={selectedColor} bold>Settings</Text>{"   "}{tabs.map((index, position) => <Text key={index}>{position ? "   " : ""}<Text inverse={index === sectionIndex} color={index === sectionIndex && focus === "sections" ? selectedColor : "white"} bold = {index === sectionIndex && focus === "sections"}>{index === sectionIndex ? ` ${SETTINGS_SECTIONS[index]} ` : SETTINGS_SECTIONS[index]}</Text></Text>)}</Text>
    <Text color={message === "Saved." ? "green" : "red"} wrap="truncate-end">{message || (saving ? "Validating settings..." : " ")}</Text>
    {section === "Attributes" ? renderAttributes(panel, { vocabularyLabel, setVocabularyLabel, fields, setFields }, leave) : null}
    {section === "Tags" ? renderTags(panel, { types, setTypes }, leave) : null}
      {section === "Import" ? (
        <Box flexDirection="column">
          <CaretInputLine
            prefix="  Default import file (UTF-8 .txt): "
            value={importPath}
            onChange={(value) => { setImportPath(value); setMessage(""); }}
            focus={contentFocused}
            width={width}
            inputKey="settings-import"
          />
        </Box>
      ) : null}
    {section === "Daily Goal" ? <Box flexDirection="column"><CaretInputLine prefix="  Daily goal (Words per day): " value={goal} onChange={(value) => { setGoal(value); setMessage(""); }} focus={contentFocused} width={width} inputKey="settings-goal" /></Box> : null}
    {section === "Appearance" ? <Box flexDirection="column">{[{ label: "Tier colors", enabled: tierColors }, { label: "Workbook homepage progress", enabled: progress }].map((option, index) => <Text key={option.label} wrap="truncate-end" color={contentFocused && index === appearanceIndex ? selectedColor : option.enabled ? "green" : auxiliary}>{`${contentFocused && index === appearanceIndex ? ">" : " "} [${option.enabled ? "●" : " "}] ${option.label}`}</Text>)}</Box> : null}
    <Text>{" "}</Text><Text color={auxiliary} wrap="truncate-end">{navigationHints.join(" | ")}</Text>
    </Box>
  </Box>;
}
