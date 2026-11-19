/*
 * Reads and validates refresh rules from a configuration worksheet.
 *
 * The first row holds the headers (column order does not matter, header names
 * are case-insensitive). Enabled, Worksheet and WatchedCells are required; the
 * other columns are only needed by the actions that use them:
 *
 *   Enabled          TRUE/1 to use the rule
 *   Worksheet        worksheet name
 *   WatchedCells     cells or ranges that trigger the rule, e.g. "B2,B3" or "B2:B4"
 *   Action           filter-value (default), sort or show-columns
 *   FilterColumnIndex  filter-value: zero-based column within the AutoFilter range
 *   FilterValue        filter-value: the value to show in that column
 *   SortColumnIndex    sort: zero-based column within the AutoFilter range
 *   SortOrder          sort: ascending (default) or descending
 *   ColumnChoices      show-columns: "choice=columns;choice=columns", e.g.
 *                      "Only A=E:E,H:H;Only B=F:F,I:I". When the watched cell holds
 *                      one of the choices, its columns are shown and the columns of
 *                      the other choices are hidden; any other value shows them all.
 *   DelayMs          debounce delay in milliseconds (default 100)
 *   Platforms        web, desktop or all (default), e.g. when desktop Excel already
 *                    has its own refresh logic
 *
 * filter-value and sort run after the worksheet has recalculated; sort keeps
 * all filters of the AutoFilter range. show-columns runs right after the change.
 */

import { parseAddressList } from "./address";

export const CONFIG_SHEET_NAME = "_FilterRefreshConfig";
export const DEFAULT_DELAY_MS = 100;

export const ACTION_FILTER_VALUE = "filter-value";
export const ACTION_SORT = "sort";
export const ACTION_SHOW_COLUMNS = "show-columns";
const SUPPORTED_ACTIONS = [ACTION_FILTER_VALUE, ACTION_SORT, ACTION_SHOW_COLUMNS];

export const PLATFORM_WEB = "web";
export const PLATFORM_DESKTOP = "desktop";
const ALL_PLATFORMS = [PLATFORM_WEB, PLATFORM_DESKTOP];

const COLUMNS = {
  enabled: { header: "Enabled", required: true },
  worksheet: { header: "Worksheet", required: true },
  watchedCells: { header: "WatchedCells", required: true },
  action: { header: "Action" },
  filterColumnIndex: { header: "FilterColumnIndex" },
  filterValue: { header: "FilterValue" },
  sortColumnIndex: { header: "SortColumnIndex" },
  sortOrder: { header: "SortOrder" },
  columnChoices: { header: "ColumnChoices" },
  delayMs: { header: "DelayMs" },
  platforms: { header: "Platforms" },
};

/**
 * @param {any[][]} values Cell values of the configuration sheet's used range.
 * @param {{ id: string, name: string }[]} worksheets Worksheets of the workbook.
 * @param {number} firstRowNumber 1-based worksheet row number of values[0].
 * @returns {{ rules: object[], errors: string[] }}
 */
export function parseConfiguration(values, worksheets, firstRowNumber = 1) {
  const rules = [];
  const errors = [];

  if (!values || values.length === 0) {
    return { rules, errors };
  }

  const headers = values[0].map((value) =>
    String(value ?? "")
      .trim()
      .toLowerCase()
  );
  const columnIndexes = {};

  for (const [key, column] of Object.entries(COLUMNS)) {
    columnIndexes[key] = headers.indexOf(column.header.toLowerCase());

    if (columnIndexes[key] < 0 && column.required) {
      errors.push(`Configuration column "${column.header}" is missing.`);
    }
  }

  if (errors.length > 0) {
    return { rules, errors };
  }

  const worksheetsByName = new Map(
    worksheets.map((worksheet) => [worksheet.name.toLowerCase(), worksheet])
  );

  for (let rowIndex = 1; rowIndex < values.length; rowIndex++) {
    const row = values[rowIndex];

    if (row.every(isBlank)) {
      continue;
    }

    const cell = (key) => (columnIndexes[key] >= 0 ? row[columnIndexes[key]] : "");

    if (!isTrue(cell("enabled"))) {
      continue;
    }

    try {
      rules.push(buildRule(cell, worksheetsByName));
    } catch (error) {
      errors.push(`Row ${firstRowNumber + rowIndex}: ${error.message}`);
    }
  }

  return { rules, errors };
}

function buildRule(cell, worksheetsByName) {
  const worksheetName = String(cell("worksheet") ?? "").trim();
  if (!worksheetName) {
    throw new Error("Worksheet is empty.");
  }

  const worksheet = worksheetsByName.get(worksheetName.toLowerCase());
  if (!worksheet) {
    throw new Error(`Worksheet "${worksheetName}" was not found.`);
  }

  const watchedCells = String(cell("watchedCells") ?? "").trim();
  if (!watchedCells) {
    throw new Error("WatchedCells is empty.");
  }
  const watchedAreas = parseAddressList(watchedCells);

  let action = ACTION_FILTER_VALUE;
  if (!isBlank(cell("action"))) {
    action = String(cell("action")).trim().toLowerCase();
    if (!SUPPORTED_ACTIONS.includes(action)) {
      throw new Error(`Action "${action}" is not supported. Use ${SUPPORTED_ACTIONS.join(", ")}.`);
    }
  }

  let delayMs = DEFAULT_DELAY_MS;
  if (!isBlank(cell("delayMs"))) {
    delayMs = toNumber(cell("delayMs"));
    if (!Number.isFinite(delayMs) || delayMs < 0) {
      throw new Error("DelayMs must be a non-negative number.");
    }
  }

  const platforms = isBlank(cell("platforms")) ? ALL_PLATFORMS : parsePlatforms(cell("platforms"));

  const rule = {
    enabled: true,
    worksheetId: worksheet.id,
    worksheetName: worksheet.name,
    watchedAddresses: watchedCells,
    watchedAreas,
    action,
    delayMs,
    platforms,
  };

  if (action === ACTION_FILTER_VALUE) {
    rule.filterColumnIndex = toColumnIndex(cell("filterColumnIndex"), "FilterColumnIndex");

    rule.filterValue = String(cell("filterValue") ?? "");
    if (rule.filterValue.trim() === "") {
      throw new Error("FilterValue is empty.");
    }
  } else if (action === ACTION_SORT) {
    rule.sortColumnIndex = toColumnIndex(cell("sortColumnIndex"), "SortColumnIndex");

    const order = isBlank(cell("sortOrder")) ? "ascending" : String(cell("sortOrder")).trim().toLowerCase();
    if (order !== "ascending" && order !== "descending") {
      throw new Error(`SortOrder "${order}" is not supported. Use ascending or descending.`);
    }
    rule.ascending = order === "ascending";
  } else if (action === ACTION_SHOW_COLUMNS) {
    rule.columnChoices = parseColumnChoices(cell("columnChoices"));
  }

  return rule;
}

/**
 * Parses "choice=columns;choice=columns" into a map from choice (trimmed,
 * case-insensitive) to sheet column numbers (1-based).
 */
function parseColumnChoices(value) {
  const choices = new Map();

  for (const entry of String(value ?? "").split(";")) {
    if (entry.trim() === "") {
      continue;
    }

    const separator = entry.lastIndexOf("=");
    if (separator <= 0) {
      throw new Error(`ColumnChoices entry "${entry.trim()}" must look like "choice=E:E,H:H".`);
    }

    const choice = entry.slice(0, separator).trim().toLowerCase();
    const columns = new Set();
    for (const area of parseAddressList(entry.slice(separator + 1))) {
      for (let column = area.left; column <= area.right; column++) {
        columns.add(column);
      }
    }
    choices.set(choice, [...columns]);
  }

  if (choices.size === 0) {
    throw new Error("ColumnChoices is empty.");
  }

  return choices;
}

function parsePlatforms(value) {
  const platforms = new Set();

  for (const name of String(value).toLowerCase().split(/[\s,;]+/)) {
    if (name === "") {
      continue;
    }
    if (name === "all") {
      ALL_PLATFORMS.forEach((platform) => platforms.add(platform));
    } else if (ALL_PLATFORMS.includes(name)) {
      platforms.add(name);
    } else {
      throw new Error(`Platforms "${name}" is not supported. Use web, desktop or all.`);
    }
  }

  return platforms.size > 0 ? [...platforms] : ALL_PLATFORMS;
}

function toColumnIndex(value, header) {
  const index = toNumber(value);
  if (!Number.isInteger(index) || index < 0) {
    throw new Error(`${header} must be a non-negative whole number.`);
  }
  return index;
}

function isBlank(value) {
  return value === null || value === undefined || String(value).trim() === "";
}

function isTrue(value) {
  if (value === true || value === 1) {
    return true;
  }

  const text = String(value ?? "")
    .trim()
    .toLowerCase();
  return text === "true" || text === "1";
}

function toNumber(value) {
  if (isBlank(value)) {
    return NaN;
  }

  return typeof value === "number" ? value : Number(String(value).trim());
}
