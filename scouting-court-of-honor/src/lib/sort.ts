/**
 * Utility functions for sorting scout names alphabetically by last name.
 */

/**
 * Extracts a sort key string for a scout's name so that names are sorted by last name.
 * Handles "First Last", "First Middle Last", "Last, First", and common suffixes (Jr., Sr., III, etc.).
 */
export function getLastNameSortKey(name: string): string {
  if (!name) return "";
  const trimmed = name.trim();

  // If already in "Last, First" format
  if (trimmed.includes(",")) {
    const parts = trimmed.split(",");
    const lastName = parts[0].trim();
    const firstName = parts.slice(1).join(" ").trim();
    return `${lastName} ${firstName}`.toLowerCase();
  }

  // Strip common suffixes (e.g., Jr., Sr., II, III, IV, V)
  const suffixRegex = /\s+(jr\.?|sr\.?|i{2,3}|iv|v)$/i;
  const clean = trimmed.replace(suffixRegex, "");
  const parts = clean.split(/\s+/);

  if (parts.length <= 1) {
    return clean.toLowerCase();
  }

  const lastName = parts[parts.length - 1];
  const firstName = parts.slice(0, parts.length - 1).join(" ");
  return `${lastName} ${firstName}`.toLowerCase();
}

/**
 * Comparator function to sort scout names alphabetically by last name, then by first name.
 */
export function compareByLastName(a: string, b: string): number {
  const keyA = getLastNameSortKey(a);
  const keyB = getLastNameSortKey(b);
  const cmp = keyA.localeCompare(keyB);
  if (cmp !== 0) return cmp;
  return a.localeCompare(b);
}
