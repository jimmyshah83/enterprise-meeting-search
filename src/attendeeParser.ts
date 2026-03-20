/**
 * Extracts person names from natural language hints.
 *
 * Strategy:
 * 1. Look for explicit attendee phrases: "John Smith was there",
 *    "with Sarah Connor", "attended by …", "invited …", etc.
 * 2. Fall back to heuristic capitalised-word bigrams / trigrams that don't
 *    match common stop words or meeting-type keywords.
 */

/** Words that look capitalised but are not person names. */
const STOP_WORDS = new Set([
  'I', 'A', 'An', 'The', 'And', 'Or', 'But', 'In', 'On', 'At', 'To',
  'For', 'Of', 'With', 'From', 'By', 'It', 'Is', 'Was', 'Were', 'Be',
  'Had', 'Have', 'Has', 'Not', 'No', 'So', 'As', 'That', 'This', 'These',
  'Those', 'There', 'Their', 'They', 'He', 'She', 'We', 'You', 'My', 'His',
  'Her', 'Our', 'Your', 'Its', 'Me', 'Him', 'Us', 'Them',
  'Teams', 'Zoom', 'Slack', 'Meet', 'Skype', 'Meeting', 'Call', 'Channel',
  'External', 'Internal', 'Monday', 'Tuesday', 'Wednesday', 'Thursday',
  'Friday', 'Saturday', 'Sunday', 'January', 'February', 'March', 'April',
  'May', 'June', 'July', 'August', 'September', 'October', 'November',
  'December', 'Jan', 'Feb', 'Mar', 'Apr', 'Jun', 'Jul', 'Aug', 'Sep',
  'Oct', 'Nov', 'Dec',
]);

/**
 * Attendee trigger patterns. The capture group should contain the name.
 * NOTE: these patterns do NOT use the `i` flag so that `[A-Z]` reliably
 * matches only uppercase letters (case-sensitive name detection).
 * The trigger words themselves are written in lowercase to match typical
 * natural-language input; add variants as needed.
 */
const ATTENDEE_PATTERNS: RegExp[] = [
  // "John Smith was there / attended / joined / participated"
  /([A-Z][a-z]+(?:\s+[A-Z][a-z]+)+)\s+(?:was there|attended|joined|participated|was present|is present|will be there)/g,
  // "with John Smith"
  /\bwith\s+([A-Z][a-z]+(?:\s+[A-Z][a-z]+)+)/g,
  // "invited John Smith" / "include John Smith"
  /\b(?:invited|include|including|involving|alongside)\s+([A-Z][a-z]+(?:\s+[A-Z][a-z]+)+)/g,
  // "attended by John Smith"
  /\battended\s+by\s+([A-Z][a-z]+(?:\s+[A-Z][a-z]+)+)/g,
  // "John Smith and Jane Doe" – capture both sides when joined by "and"
  /([A-Z][a-z]+\s+[A-Z][a-z]+)\s+and\s+([A-Z][a-z]+\s+[A-Z][a-z]+)/g,
];

/**
 * Extracts person names from a natural language hint string.
 *
 * @param hint - The raw user-supplied hint text.
 * @returns An array of unique person name strings.
 */
export function parseAttendees(hint: string): string[] {
  const found = new Set<string>();

  for (const pattern of ATTENDEE_PATTERNS) {
    // Reset lastIndex for global regexes reused across calls
    pattern.lastIndex = 0;
    let match: RegExpExecArray | null;
    while ((match = pattern.exec(hint)) !== null) {
      // Some patterns have two capture groups (the "and" pattern)
      for (let g = 1; g < match.length; g++) {
        if (match[g]) {
          found.add(normalise(match[g]));
        }
      }
    }
  }

  // Heuristic fallback: capitalised word bigrams / trigrams not in stop list
  if (found.size === 0) {
    const tokens = hint.split(/\s+/);
    for (let i = 0; i < tokens.length; i++) {
      const t = tokens[i].replace(/[^A-Za-z'-]/g, '');
      if (!isCapitalisedName(t)) continue;

      // Try trigram first
      if (i + 2 < tokens.length) {
        const t2 = tokens[i + 1].replace(/[^A-Za-z'-]/g, '');
        const t3 = tokens[i + 2].replace(/[^A-Za-z'-]/g, '');
        if (isCapitalisedName(t2) && isCapitalisedName(t3)) {
          found.add(`${t} ${t2} ${t3}`);
          i += 2;
          continue;
        }
      }
      // Try bigram
      if (i + 1 < tokens.length) {
        const t2 = tokens[i + 1].replace(/[^A-Za-z'-]/g, '');
        if (isCapitalisedName(t2)) {
          found.add(`${t} ${t2}`);
          i += 1;
          continue;
        }
      }
    }
  }

  return Array.from(found);
}

function isCapitalisedName(word: string): boolean {
  return (
    word.length > 1 &&
    /^[A-Z][a-z]/.test(word) &&
    !STOP_WORDS.has(word)
  );
}

function normalise(name: string): string {
  return name.trim().replace(/\s+/g, ' ');
}
