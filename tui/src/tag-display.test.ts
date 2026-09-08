import assert from "node:assert/strict";
import test from "node:test";
import { EntryRow, TagType } from "./db.js";
import { fitTagBadges, visibleAssignedTagGroups } from "./tag-display.js";

const entry = { tags: [{ id: 12, tagTypeId: 2, name: "Travel" }, { id: 10, tagTypeId: 1, name: "noun" }] } as EntryRow;
const types: TagType[] = [
  { id: 1, workbookId: 1, name: "Part of Speech", position: 1, visible: true, tags: [{ id: 10, tagTypeId: 1, name: "noun" }, { id: 11, tagTypeId: 1, name: "verb" }] },
  { id: 2, workbookId: 1, name: "Topic", position: 2, visible: false, tags: [{ id: 12, tagTypeId: 2, name: "Travel" }] },
];

test("visible tag groups include assigned tags only and preserve type and tag order", () => {
  assert.deepEqual(visibleAssignedTagGroups(entry, types), [{ typeId: 1, typeName: "Part of Speech", tagNames: ["noun"] }]);
  assert.deepEqual(visibleAssignedTagGroups(entry, types.map((type) => ({ ...type, visible: true }))), [
    { typeId: 1, typeName: "Part of Speech", tagNames: ["noun"] },
    { typeId: 2, typeName: "Topic", tagNames: ["Travel"] },
  ]);
});

test("badge fitting keeps complete badges and skips ones that do not fit", () => {
  assert.deepEqual(fitTagBadges(["noun", "very long tag", "v."], 11, (text) => text.length), ["[noun]", "[v.]"]);
  assert.deepEqual(fitTagBadges(["名詞", "動詞"], 9, (text) => Array.from(text).reduce((width, char) => width + (/[^\x00-\xff]/.test(char) ? 2 : 1), 0)), ["[名詞]"]);
});
