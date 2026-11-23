/*
 * A1 address parsing and intersection helpers.
 *
 * Supports single cells (B2), rectangular ranges (B2:D5), absolute references
 * ($B$2), whole columns (B:D), whole rows (2:5), optional worksheet prefixes
 * (Sheet1!B2, 'My sheet'!B2) and comma-separated lists of these areas.
 */

const MAX_ROWS = 1048576;
const MAX_COLUMNS = 16384;

/**
 * Parses a comma-separated list of A1 addresses into rectangular areas.
 * @param {string} text
 * @returns {{ top: number, left: number, bottom: number, right: number }[]}
 */
export function parseAddressList(text) {
  const areas = splitOutsideQuotes(String(text ?? ""), ",")
    .map((part) => part.trim())
    .filter((part) => part !== "")
    .map(parseArea);

  if (areas.length === 0) {
    throw new Error("No cell address was given.");
  }

  return areas;
}

/**
 * Returns true when any area of the first list overlaps any area of the second.
 */
export function areasIntersect(firstAreas, secondAreas) {
  return firstAreas.some((first) =>
    secondAreas.some(
      (second) =>
        first.top <= second.bottom &&
        second.top <= first.bottom &&
        first.left <= second.right &&
        second.left <= first.right
    )
  );
}

/**
 * Returns true when every inner area lies completely inside the outer area.
 */
export function areasInside(innerAreas, outer) {
  return innerAreas.every(
    (inner) =>
      inner.top >= outer.top &&
      inner.bottom <= outer.bottom &&
      inner.left >= outer.left &&
      inner.right <= outer.right
  );
}

/**
 * Converts a 1-based column number to its letters, e.g. 1 -> "A", 28 -> "AB".
 */
export function columnLetters(column) {
  let letters = "";

  for (let rest = column; rest > 0; rest = Math.floor((rest - 1) / 26)) {
    letters = String.fromCharCode(65 + ((rest - 1) % 26)) + letters;
  }

  return letters;
}

function parseArea(part) {
  const address = stripWorksheetPrefix(part).replace(/\$/g, "").toUpperCase().trim();
  const ends = address.split(":");

  if (ends.length > 2) {
    throw new Error(`Invalid address "${part}".`);
  }

  const start = parseReference(ends[0], part);
  const end = ends.length === 2 ? parseReference(ends[1], part) : start;

  if (start.kind !== end.kind || (ends.length === 1 && start.kind !== "cell")) {
    throw new Error(`Invalid address "${part}".`);
  }

  const area = {
    top: Math.min(start.row, end.row),
    left: Math.min(start.column, end.column),
    bottom: Math.max(start.row, end.row),
    right: Math.max(start.column, end.column),
  };

  if (start.kind === "column") {
    area.top = 1;
    area.bottom = MAX_ROWS;
  } else if (start.kind === "row") {
    area.left = 1;
    area.right = MAX_COLUMNS;
  }

  return area;
}

function parseReference(reference, originalPart) {
  let match = /^([A-Z]{1,3})(\d{1,7})$/.exec(reference);
  if (match) {
    return checkLimits(
      { kind: "cell", column: columnToNumber(match[1]), row: Number(match[2]) },
      originalPart
    );
  }

  match = /^([A-Z]{1,3})$/.exec(reference);
  if (match) {
    return checkLimits({ kind: "column", column: columnToNumber(match[1]), row: 1 }, originalPart);
  }

  match = /^(\d{1,7})$/.exec(reference);
  if (match) {
    return checkLimits({ kind: "row", column: 1, row: Number(match[1]) }, originalPart);
  }

  throw new Error(`Invalid address "${originalPart}".`);
}

function checkLimits(reference, originalPart) {
  if (
    reference.row < 1 ||
    reference.row > MAX_ROWS ||
    reference.column < 1 ||
    reference.column > MAX_COLUMNS
  ) {
    throw new Error(`Address "${originalPart}" is outside the worksheet.`);
  }

  return reference;
}

function columnToNumber(letters) {
  let column = 0;

  for (const letter of letters) {
    column = column * 26 + (letter.charCodeAt(0) - 64);
  }

  return column;
}

function stripWorksheetPrefix(part) {
  return splitWorksheetPrefix(part).address;
}

/**
 * Splits "'My sheet'!B2" into its worksheet name ("My sheet", the quotes and
 * doubled quotes undone) and its address ("B2"); the worksheet is null
 * without a prefix.
 * @param {string} part
 * @returns {{ worksheet: string | null, address: string }}
 */
export function splitWorksheetPrefix(part) {
  let inQuotes = false;
  let separatorIndex = -1;

  for (let index = 0; index < part.length; index++) {
    if (part[index] === "'") {
      inQuotes = !inQuotes;
    } else if (part[index] === "!" && !inQuotes) {
      separatorIndex = index;
    }
  }

  if (separatorIndex < 0) {
    return { worksheet: null, address: part };
  }

  let worksheet = part.slice(0, separatorIndex).trim();
  if (worksheet.length >= 2 && worksheet.startsWith("'") && worksheet.endsWith("'")) {
    worksheet = worksheet.slice(1, -1).replace(/''/g, "'");
  }
  return { worksheet, address: part.slice(separatorIndex + 1) };
}

/** Splits a comma-separated address list into its parts; commas in quoted sheet names stay. */
export function splitAddressList(text) {
  return splitOutsideQuotes(String(text ?? ""), ",")
    .map((part) => part.trim())
    .filter((part) => part !== "");
}

function splitOutsideQuotes(text, separator) {
  const parts = [];
  let inQuotes = false;
  let current = "";

  for (const character of text) {
    if (character === "'") {
      inQuotes = !inQuotes;
    }

    if (character === separator && !inQuotes) {
      parts.push(current);
      current = "";
    } else {
      current += character;
    }
  }

  parts.push(current);
  return parts;
}
