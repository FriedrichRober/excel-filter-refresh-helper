/*
 * Writes a text describing the active column filters of a worksheet into a
 * status worksheet, so a formula in the workbook can show it.
 *
 * The status worksheet (usually very hidden) has a header row and then one row
 * per worksheet:
 *
 *   A  worksheet name (exact match)                          written by the workbook
 *   B  text of the active column filters ("" when none)      written by the add-in
 *   C  value of D at the moment B was written                written by the add-in
 *   D  formula counting the visible rows of the worksheet    written by the workbook
 *   E  AutoFilter columns not to name: zero-based indexes,   written by the workbook
 *      ";"-separated, e.g. "0;1"
 *
 * A formula shows B only while C equals D, so B and C are always written
 * together. The text must match what a desktop macro writes for the same
 * filters, character by character.
 */

export const STATUS_SHEET_NAME = "_FilterStatus";
export const MAX_STATUS_LENGTH = 200;

const SEPARATOR = " · ";
const QUOTE_OPEN = "„";
const QUOTE_CLOSE = "“";
const ELLIPSIS = "…";
const EMPTY_VALUE = "(leer)";
const OWN_RULE = "eigene Regel";
const MAX_NAMED_VALUES = 2;

/**
 * Reads the worksheet names in column A of the status sheet.
 *
 * @param {any[][]} values Values of column A, starting at sheet row firstRowNumber.
 * @param {number} firstRowNumber 1-based sheet row of values[0].
 * @returns {Map<string, number>} 1-based sheet row by exact worksheet name
 */
export function parseStatusRows(values, firstRowNumber = 1) {
  const rows = new Map();

  (values || []).forEach((row, index) => {
    const rowNumber = firstRowNumber + index;
    const name = row[0];
    // Row 1 holds the headers.
    if (rowNumber > 1 && typeof name === "string" && name !== "" && !rows.has(name)) {
      rows.set(name, rowNumber);
    }
  });

  return rows;
}

/** Parses column E, e.g. "0;1" or 0, into a set of zero-based column indexes. */
export function parseExcludedColumns(value) {
  const excluded = new Set();

  for (const part of String(value ?? "").split(";")) {
    const text = part.trim();
    if (/^\d+$/.test(text)) {
      excluded.add(Number(text));
    }
  }

  return excluded;
}

/**
 * Builds the status text.
 *
 * @param {any[]} headers Header cells of the AutoFilter range, left to right.
 * @param {object[]} criteria AutoFilter.criteria, one entry per column.
 * @param {Set<number>} excluded Column indexes not to name.
 * @returns {string}
 */
export function formatFilterStatus(headers, criteria, excluded = new Set()) {
  const parts = [];

  (criteria || []).forEach((criterion, columnIndex) => {
    if (excluded.has(columnIndex) || !criterion || !criterion.filterOn) {
      return;
    }
    parts.push(`${formatHeader(headers[columnIndex])} ${formatSetting(criterion)}`);
  });

  const text = parts.join(SEPARATOR);
  return text.length > MAX_STATUS_LENGTH ? text.slice(0, MAX_STATUS_LENGTH - 1) + ELLIPSIS : text;
}

export function formatHeader(value) {
  return String(value ?? "")
    .replace(/\r/g, "")
    .replace(/-\n/g, "-")
    .replace(/\n/g, " ")
    .replace(/^ +| +$/g, "");
}

function formatSetting(criterion) {
  if (criterion.filterOn === "Values") {
    return formatValues(criterion.values || []);
  }
  if (criterion.filterOn === "Custom") {
    return formatCustom(criterion);
  }
  return OWN_RULE;
}

function formatValues(values) {
  // Dates of a date tree arrive as objects, not strings.
  const texts = values.filter((value) => typeof value === "string");

  if (texts.length === 0) {
    return OWN_RULE;
  }

  let text = texts
    .slice(0, MAX_NAMED_VALUES)
    .map((value) => quote(value === "" ? EMPTY_VALUE : value))
    .join(", ");
  if (texts.length > MAX_NAMED_VALUES) {
    text += ` +${texts.length - MAX_NAMED_VALUES}`;
  }
  return text;
}

function formatCustom(criterion) {
  const first = criterion.criterion1 ?? "";
  const second = criterion.criterion2;

  if (second === undefined || second === null || second === "") {
    return quote(formatCriterion(first));
  }

  let joiner;
  if (criterion.operator === "And") {
    joiner = " und ";
  } else {
    joiner = isPlainValue(first) && isPlainValue(second) ? ", " : " oder ";
  }
  return quote(formatCriterion(first)) + joiner + quote(formatCriterion(second));
}

/** "=X" is shown as X, "=" or "" as (leer); comparisons stay as they are. */
function formatCriterion(criterion) {
  const text = String(criterion ?? "");
  if (text === "" || text === "=") {
    return EMPTY_VALUE;
  }
  if (isEquals(text)) {
    return text.slice(1);
  }
  return text;
}

function isEquals(text) {
  return text.startsWith("=") && !/^=[=<>]/.test(text);
}

function isPlainValue(criterion) {
  const text = String(criterion ?? "");
  return isEquals(text) && !/[*?]/.test(text);
}

function quote(text) {
  return QUOTE_OPEN + text + QUOTE_CLOSE;
}
