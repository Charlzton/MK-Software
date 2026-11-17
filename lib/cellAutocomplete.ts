export interface AutocompleteEntry {
  value: string;
  lower: string;
  normalizedKey: string;
  count: number;
  lastRowIdx: number;
}

const EXCLUDED_EXACT_HEADERS = new Set([
  'date',
  'date_of_birth',
  'date of birth',
  'dob',
  'oct-date',
  'oct date',
  'review date',
  'date of renewal',
  'date of delivery',
  'time',
  'jan',
  'feb',
  'mar',
  'apr',
  'may',
  'jun',
  'jul',
  'aug',
  'sep',
  'sept',
  'oct',
  'nov',
  'dec',
  'phone',
  'phone_no',
  'phone no',
  'tel',
  'telephone',
  'mobile',
  'contact',
  'email',
  'e-mail',
  'tin',
  'patient no',
  'patient_no',
  'patient id',
  'invoice',
  'invoice no',
  'reference',
  'ref',
  'ref no',
  'licence',
  'license',
  'licence no',
  'license no',
  'id',
  'age',
  'qty',
  'quantity',
  'count',
  'test_fee',
  'test fee',
  'fee',
  'price',
  'cost',
  'amount',
  'total',
  'glass total',
  'glasses bal',
  'bal',
  'balance',
  'paid',
  'unpaid',
  'drops_no',
  'drops',
  'new',
  'renewal',
  'target',
  'sph',
  'cyl',
  'axis',
  'add',
  'pd',
  'va',
  'va right',
  'va left',
  'iop',
  'cdr',
  'pachym',
  'md',
  'psd',
  'vfi',
  'bp check',
  'bp',
  'signal strength',
  'f. loss',
  'f.loss',
  'questions',
  'p. seen',
  'p. missed',
  'sup',
  'nas',
  'inf',
  'temp',
  'rnfl',
  'availability',
  'sms',
  'status'
]);

const EXCLUDED_HEADER_PATTERNS: RegExp[] = [
  /\bdate\b/i,
  /\bdob\b/i,
  /\btime\b/i,
  /\bphone\b/i,
  /\btel\b/i,
  /\bemail\b/i,
  /\bfee\b/i,
  /\btotal\b/i,
  /\bbal(?:ance)?\b/i,
  /\bprice\b/i,
  /\bcost\b/i,
  /\bamount\b/i,
  /\bpaid\b/i,
  /\btin\b/i,
  /\binvoice\b/i,
  /\blicen[cs]e\b/i,
  /\bpatient\s*no\b/i,
  /\bdrops_no\b/i,
  /\b(?:sph|cyl|axis|pachym|iop|cdr|rnfl)\b/i
];

const NUMERIC_RE = /^[+-]?(?:\d+(?:,\d{3})*(?:\.\d+)?|\.\d+)$/;
const CURRENCY_RE = /^(?:ghs|gh¢|cedis?|\$|€|£)?\s*[+-]?(?:\d+(?:,\d{3})*(?:\.\d+)?|\.\d+)$/i;
const DATE_LIKE_RE =
  /^(?:\d{1,4}[\/\-.\s]\d{1,2}[\/\-.\s]\d{1,4}|\d{1,2}[\/\-.\s](?:jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*[\/\-.\s]?\d{0,4}|(?:jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\s+\d{1,2},?\s+\d{2,4})$/i;
const PHONE_LIKE_RE = /^(?:\+?\d[\d\s\-()]{6,}\d)$/;
const BOOLEAN_STATUS_SET = new Set([
  'paid',
  'unpaid',
  'yes',
  'no',
  'true',
  'false',
  'y',
  'n',
  'ok',
  'done',
  'pending',
  'sent',
  'available',
  'unavailable'
]);

function isExcludedValueShape(raw: string): boolean {
  const s = raw.trim();
  if (!s) return true;
  if (s.startsWith('=')) return true;
  if (NUMERIC_RE.test(s) || CURRENCY_RE.test(s)) return true;
  if (DATE_LIKE_RE.test(s)) return true;
  if (PHONE_LIKE_RE.test(s)) return true;
  if (BOOLEAN_STATUS_SET.has(s.toLowerCase())) return true;
  // Pure email addresses shouldn't populate shared autocomplete lists
  if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s)) return true;
  return false;
}

/**
 * Determine whether smart cell autocomplete should be active for a worksheet column.
 * Excludes dates, numeric/billing/fee/total/count columns, Paid/status/boolean toggles,
 * phones/IDs/TIN/emails, and columns whose populated data is overwhelmingly numeric/date/boolean.
 */
export function isColumnEligibleForAutocomplete(
  headers: string[],
  colIndex: number,
  rows: (string | number)[][],
  sheetName?: string
): boolean {
  if (colIndex < 0) return false;

  const rawHeader = (headers[colIndex] ?? '').trim();
  const normalizedHeader = rawHeader.toLowerCase().replace(/\s+/g, ' ');

  if (normalizedHeader) {
    if (EXCLUDED_EXACT_HEADERS.has(normalizedHeader)) {
      return false;
    }
    for (const pattern of EXCLUDED_HEADER_PATTERNS) {
      if (pattern.test(normalizedHeader)) {
        return false;
      }
    }
  }

  // On the inventory count grid ("Medications"), columns store stock counts rather than text categories
  const normalizedSheet = (sheetName ?? '').trim().toLowerCase();
  if (
    normalizedSheet === 'medications' &&
    normalizedHeader !== 'function' &&
    normalizedHeader !== 'trade name' &&
    normalizedHeader !== 'dosage'
  ) {
    return false;
  }

  // Inspect existing values in this column to auto-exclude numeric/date/status columns
  let nonEmptyCount = 0;
  let excludedShapeCount = 0;
  let validTextCount = 0;

  for (let r = 0; r < rows.length; r++) {
    const cell = rows[r]?.[colIndex];
    if (cell === null || cell === undefined) continue;
    const str = String(cell).trim();
    if (!str) continue;

    nonEmptyCount++;
    if (typeof cell === 'number' || isExcludedValueShape(str)) {
      excludedShapeCount++;
    } else {
      validTextCount++;
    }

    // Fast-exit once enough rows have been sampled
    if (nonEmptyCount >= 80) break;
  }

  // Reception "PATIENT'S DATA" has numeric Medication fee column; keep excluded if no text drugs exist
  if (normalizedHeader === 'medication' && validTextCount === 0) {
    return false;
  }

  if (nonEmptyCount === 1 && excludedShapeCount === 1) {
    return false;
  }

  if (nonEmptyCount >= 2 && excludedShapeCount / nonEmptyCount >= 0.6) {
    return false;
  }

  return true;
}

/**
 * Build a deduplicated, frequency-counted index of previously entered text values in a column.
 * Operates strictly in memory on already-loaded sheet data (zero network/DB calls).
 */
export function buildColumnSuggestionIndex(
  rows: (string | number)[][],
  colIndex: number
): AutocompleteEntry[] {
  if (colIndex < 0 || !Array.isArray(rows) || rows.length === 0) {
    return [];
  }

  const byKey = new Map<string, AutocompleteEntry>();

  for (let r = 0; r < rows.length; r++) {
    const raw = rows[r]?.[colIndex];
    if (raw === null || raw === undefined || typeof raw === 'number') continue;

    const trimmed = String(raw).trim();
    if (trimmed.length < 2) continue;
    if (isExcludedValueShape(trimmed)) continue;

    const normalizedKey = trimmed.toLowerCase().replace(/\s+/g, ' ');
    const existing = byKey.get(normalizedKey);
    if (existing) {
      existing.count += 1;
      existing.lastRowIdx = r;
      // Prefer the cleanest/most recent casing
      existing.value = trimmed;
      existing.lower = trimmed.toLowerCase();
    } else {
      byKey.set(normalizedKey, {
        value: trimmed,
        lower: trimmed.toLowerCase(),
        normalizedKey,
        count: 1,
        lastRowIdx: r
      });
    }
  }

  return Array.from(byKey.values());
}

/**
 * Filter & rank matching autocomplete suggestions for the user's current input:
 * 1. Prefix matches first
 * 2. Partial substring matches second
 * 3. Sorted within each tier by usage frequency, then recency, then alphabetically
 */
export function filterAutocompleteSuggestions(
  entries: AutocompleteEntry[],
  rawQuery: string,
  limit: number = 8
): string[] {
  if (!entries || entries.length === 0) return [];
  if (!rawQuery || rawQuery.trimStart().startsWith('=')) return [];

  const trimmedQuery = rawQuery.trim();
  const qLower = trimmedQuery.toLowerCase();
  if (!qLower) return [];

  const prefixMatches: AutocompleteEntry[] = [];
  const partialMatches: AutocompleteEntry[] = [];

  for (const entry of entries) {
    // Skip suggesting the exact identical string the user already typed
    if (entry.value === trimmedQuery) continue;

    if (entry.lower.startsWith(qLower)) {
      prefixMatches.push(entry);
    } else if (entry.lower.includes(qLower)) {
      partialMatches.push(entry);
    }
  }

  const compareEntries = (a: AutocompleteEntry, b: AutocompleteEntry) => {
    if (b.count !== a.count) return b.count - a.count;
    if (b.lastRowIdx !== a.lastRowIdx) return b.lastRowIdx - a.lastRowIdx;
    return a.value.localeCompare(b.value);
  };

  prefixMatches.sort(compareEntries);
  partialMatches.sort(compareEntries);

  const combined = [...prefixMatches, ...partialMatches];
  return combined.slice(0, limit).map((e) => e.value);
}
