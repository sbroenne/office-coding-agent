/**
 * Convert a snake_case tool name into a friendly, sentence-case label.
 *
 * Examples:
 *   get_workbook_info     → "Get workbook info"
 *   set_range_values      → "Set range values"
 *   create_table          → "Create table"
 *   add_conditional_format → "Add conditional format"
 *   list_named_ranges     → "List named ranges"
 */
export function humanizeToolName(toolName: string): string {
  const words = toolName.replace(/_/g, ' ').trim();
  if (!words) return toolName;
  return words.charAt(0).toUpperCase() + words.slice(1);
}
