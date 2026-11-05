export const LIST_ORDER_OPTIONS = [
  { value: "add_asc", label: "Added: oldest first", parameter: "add asc" },
  { value: "add_desc", label: "Added: newest first", parameter: "add desc" },
  { value: "error_asc", label: "Error count: ascending (unpracticed last)", parameter: "error asc" },
  { value: "error_desc", label: "Error count: descending (unpracticed first)", parameter: "error desc" },
] as const;

export type ListOrder = typeof LIST_ORDER_OPTIONS[number]["value"];
export const LIST_ORDER_USAGE = "Usage: /list [add asc | add desc | error asc | error desc]";

export function parseListOrder(args: string[]): ListOrder | null {
  if (args.length !== 2) return null;
  return LIST_ORDER_OPTIONS.find((option) => option.parameter === args.join(" ").toLowerCase())?.value ?? null;
}
