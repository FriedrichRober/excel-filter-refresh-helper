/*
 * Reads and validates filter refresh rules from a configuration worksheet.
 *
 * Expected header row (column order does not matter, header names are
 * case-insensitive):
 *
 *   Enabled | Worksheet | WatchedCells | FilterColumnIndex | FilterValue | DelayMs | Action | Platforms
 *
 * DelayMs, Action and Platforms are optional; blank cells use the defaults below.
 * Platforms limits a rule to Excel for the web ("web"), desktop Excel ("desktop")
 * or both ("all", the default), e.g. when desktop already has its own refresh logic.
 */

import { parseAddressList } from "./address";

export const CONFIG_SHEET_NAME = "_FilterRefreshConfig";
export const DEFAULT_DELAY_MS = 100;
export const ACTION_FILTER_VALUE = "filter-value";

const SUPPORTED_ACTIONS = new Set([ACTION_FILTER_VALUE]);

export const PLATFORM_WEB = "web";
export const PLATFORM_DESKTOP = "desktop";
const ALL_PLATFORMS = [PLATFORM_WEB, PLATFORM_DESKTOP];

const COLUMNS = {
  enabled: { header: "Enabled", required: true },
  worksheet: { header: "Worksheet", required: true },
  watchedCells: { header: "WatchedCells", required: true },
  filterColumnIndex: { header: "FilterColumnIndex", required: true },
  filterValue: { header: "FilterValue", required: true },
  delayMs: { header: "DelayMs", required: false },
  action: { header: "Action", required: false },
  platforms: { header: "Platforms", required: false },
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

  const filterColumnIndex = toNumber(cell("filterColumnIndex"));
  if (!Number.isInteger(filterColumnIndex) || filterColumnIndex < 0) {
    throw new Error("FilterColumnIndex must be a non-negative whole number.");
  }

  const filterValue = String(cell("filterValue") ?? "");
  if (filterValue.trim() === "") {
    throw new Error("FilterValue is empty.");
  }

  let delayMs = DEFAULT_DELAY_MS;
  if (!isBlank(cell("delayMs"))) {
    delayMs = toNumber(cell("delayMs"));
    if (!Number.isFinite(delayMs) || delayMs < 0) {
      throw new Error("DelayMs must be a non-negative number.");
    }
  }

  let action = ACTION_FILTER_VALUE;
  if (!isBlank(cell("action"))) {
    action = String(cell("action")).trim().toLowerCase();
    if (!SUPPORTED_ACTIONS.has(action)) {
      throw new Error(`Action "${action}" is not supported.`);
    }
  }

  const platforms = isBlank(cell("platforms")) ? ALL_PLATFORMS : parsePlatforms(cell("platforms"));

  return {
    enabled: true,
    worksheetId: worksheet.id,
    worksheetName: worksheet.name,
    watchedAddresses: watchedCells,
    watchedAreas,
    filterColumnIndex,
    filterValue,
    delayMs,
    action,
    platforms,
  };
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
