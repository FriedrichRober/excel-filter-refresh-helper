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
 *   Action           filter-value (default), sort, show-columns, filter-table,
 *                    reset-on-click, toggle-on-click, multi-select-on-click or
 *                    refresh-on-activate
 *   FilterColumnIndex  filter-value: zero-based column within the AutoFilter range
 *   FilterValue        filter-value: the value to show in that column
 *                      filter-table: the values to show, separated by "|"
 *   SortColumnIndex    sort: zero-based column within the AutoFilter range
 *   SortOrder          sort: ascending (default) or descending
 *   ColumnChoices      show-columns: "choice=columns;choice=columns", e.g.
 *                      "Only A=E:E,H:H;Only B=F:F,I:I". When the watched cell holds
 *                      one of the choices, its columns are shown and the columns of
 *                      the other choices are hidden; any other value shows them all.
 *   TargetCells        filter-table: the table column it filters, "Table[Column]"
 *                      reset-on-click, toggle-on-click: single cells to write,
 *                      e.g. "B8,F8"; toggle-on-click also on other worksheets,
 *                      e.g. "C3,'Other sheet'!D3" (the first holds the value
 *                      it switches from)
 *                      multi-select-on-click: the one cell to write
 *                      refresh-on-activate: the one cell, usually on another
 *                      (hidden) worksheet, e.g. "'Applied'!C2", that holds the
 *                      value of the watched cell the worksheet was last
 *                      refreshed for
 *   TargetValue        filter-table: the value of the watched cell that filters
 *                      reset-on-click: the value written into each target cell
 *                      toggle-on-click: two different values, "first;second"
 *                      multi-select-on-click: the value meaning "every entry",
 *                      also the first entry of the cell's validation list
 *   SelectCell         toggle-on-click, multi-select-on-click: the single cell
 *                      selected afterwards; outside WatchedCells. Optional
 *                      column, only needed by these actions.
 *   SelectionTitle     multi-select-on-click: the input title (prompt title) of
 *                      the watched cell's data validation; the popup only opens
 *                      when the cell carries it, and its heading is built from
 *                      it. Optional column, only needed by this action.
 *   DelayMs          debounce delay in milliseconds (default 100); not used by
 *                    reset-on-click, must be empty for toggle-on-click and
 *                    multi-select-on-click
 *   Platforms        web, desktop or all (default), e.g. when desktop Excel already
 *                    has its own refresh logic
 *
 * filter-value and sort run after the worksheet has recalculated; sort keeps
 * all filters of the AutoFilter range. show-columns runs right after the change.
 * reset-on-click is triggered by selecting the watched cells instead of
 * changing them, so they act as a button (e.g. a merged cell): it writes
 * TargetValue into every target cell and then selects the first target cell.
 * It needs no AutoFilter.
 *
 * toggle-on-click is triggered by selection as well, so the watched cells act
 * as an on/off switch: if the first target cell holds the first value of
 * TargetValue, every target cell gets the second value, otherwise the first.
 * Then SelectCell is selected. The other rules watching the written cells run
 * as if the value had been typed, on the rule's worksheet at once and on other
 * worksheets too, unless a refresh-on-activate rule watches the written cell
 * there: those worksheets refresh when they are next activated. It needs no
 * AutoFilter and must leave DelayMs, FilterColumnIndex, FilterValue,
 * SortColumnIndex, SortOrder and ColumnChoices empty.
 *
 * filter-table runs right after the change, like show-columns: when the watched
 * cell holds TargetValue, the table column of TargetCells shows only the
 * values of FilterValue, otherwise its filter is cleared. A slicer on that
 * column shows the same, as its filter is the table's. The table may lie on
 * another worksheet, also a hidden one; the watched cell is typically a
 * toggle-on-click.
 *
 * multi-select-on-click is triggered by selection as well and opens a popup
 * with one checkbox per entry of the watched cell's data-validation list (the
 * list's first entry, TargetValue, means "every entry" and gets no checkbox).
 * On OK, the target cell gets TargetValue when every entry is ticked,
 * otherwise the ticked entries joined by ", " in list order; the other rules
 * watching it run as if the value had been typed. After OK or Cancel,
 * SelectCell is selected. It leaves the same columns empty as toggle-on-click.
 *
 * refresh-on-activate defers the refresh of a worksheet until it is opened.
 * WatchedCells is one cell (e.g. a toggle, or the first cell of a merged
 * one), TargetCells one cell that holds the value the worksheet was last
 * refreshed for. When the worksheet is activated (or is active when the
 * configuration loads) and the two values differ, the rules of the worksheet
 * watching that cell run as if the value had been typed, and then the
 * watched value is written into the target cell. It is written as well
 * whenever those rules run for that cell otherwise. Only WatchedCells,
 * TargetCells and Platforms are used; every other column must be empty.
 */

import { areasIntersect, columnLetters, parseAddressList, splitAddressList, splitWorksheetPrefix } from "./address";

export const CONFIG_SHEET_NAME = "_FilterRefreshConfig";
export const DEFAULT_DELAY_MS = 100;

export const ACTION_FILTER_VALUE = "filter-value";
export const ACTION_SORT = "sort";
export const ACTION_SHOW_COLUMNS = "show-columns";
export const ACTION_RESET_ON_CLICK = "reset-on-click";
export const ACTION_TOGGLE_ON_CLICK = "toggle-on-click";
export const ACTION_MULTI_SELECT_ON_CLICK = "multi-select-on-click";
export const ACTION_FILTER_TABLE = "filter-table";
export const ACTION_REFRESH_ON_ACTIVATE = "refresh-on-activate";
// The separator of the values of filter-table in FilterValue.
export const TABLE_VALUE_SEPARATOR = "|";
const SUPPORTED_ACTIONS = [
  ACTION_FILTER_VALUE,
  ACTION_SORT,
  ACTION_SHOW_COLUMNS,
  ACTION_RESET_ON_CLICK,
  ACTION_TOGGLE_ON_CLICK,
  ACTION_MULTI_SELECT_ON_CLICK,
  ACTION_FILTER_TABLE,
  ACTION_REFRESH_ON_ACTIVATE,
];

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
  targetCells: { header: "TargetCells" },
  targetValue: { header: "TargetValue" },
  selectCell: { header: "SelectCell" },
  selectionTitle: { header: "SelectionTitle" },
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
    cell.hasColumn = (key) => columnIndexes[key] >= 0;

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

  if (action === ACTION_TOGGLE_ON_CLICK || action === ACTION_MULTI_SELECT_ON_CLICK) {
    // Validated before DelayMs is parsed, so a wrong DelayMs gets this message.
    for (const key of CLICK_UNUSED_COLUMNS) {
      if (!isBlank(cell(key))) {
        throw new Error(`${COLUMNS[key].header} must be empty for ${action}.`);
      }
    }
  } else if (action === ACTION_REFRESH_ON_ACTIVATE) {
    for (const key of ACTIVATE_UNUSED_COLUMNS) {
      if (!isBlank(cell(key))) {
        throw new Error(`${COLUMNS[key].header} must be empty for ${action}.`);
      }
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
  } else if (action === ACTION_RESET_ON_CLICK) {
    rule.targetAddresses = parseTargetCells(cell("targetCells"));

    if (isBlank(cell("targetValue"))) {
      throw new Error("TargetValue is empty.");
    }
    // Written as-is, without trimming.
    rule.targetValue = String(cell("targetValue"));
  } else if (action === ACTION_TOGGLE_ON_CLICK) {
    // Targets may lie on other worksheets ("'Other sheet'!C3"): one click
    // switches toggles that belong together on several sheets.
    rule.targets = parseToggleTargets(cell("targetCells"), worksheet.name);
    rule.targetAddresses = rule.targets
      .filter((target) => target.worksheetName === worksheet.name)
      .map((target) => target.address);
    rule.targetAreas = rule.targetAddresses.length > 0 ? parseAddressList(rule.targetAddresses.join(",")) : [];
    rule.toggleValues = parseToggleValues(cell("targetValue"));

    if (!cell.hasColumn("selectCell")) {
      throw new Error(`Configuration column "${COLUMNS.selectCell.header}" is missing; ${ACTION_TOGGLE_ON_CLICK} needs it.`);
    }
    rule.selectAddress = parseSelectCell(cell("selectCell"), watchedAreas);
  } else if (action === ACTION_FILTER_TABLE) {
    Object.assign(rule, parseTableColumn(cell("targetCells")));

    rule.filterChoice = String(cell("targetValue") ?? "").trim();
    if (rule.filterChoice === "") {
      throw new Error("TargetValue is empty.");
    }
    rule.tableValues = String(cell("filterValue") ?? "")
      .split(TABLE_VALUE_SEPARATOR)
      .map((part) => part.trim())
      .filter((part) => part !== "");
    if (rule.tableValues.length === 0) {
      throw new Error("FilterValue is empty.");
    }
  } else if (action === ACTION_MULTI_SELECT_ON_CLICK) {
    rule.targetAddresses = parseTargetCells(cell("targetCells"));
    if (rule.targetAddresses.length !== 1) {
      throw new Error(`TargetCells must be a single cell for ${action}.`);
    }
    rule.targetAreas = parseAddressList(rule.targetAddresses[0]);

    rule.allValue = String(cell("targetValue") ?? "").trim();
    if (rule.allValue === "") {
      throw new Error("TargetValue is empty.");
    }

    for (const key of ["selectCell", "selectionTitle"]) {
      if (!cell.hasColumn(key)) {
        throw new Error(`Configuration column "${COLUMNS[key].header}" is missing; ${action} needs it.`);
      }
    }
    rule.selectAddress = parseSelectCell(cell("selectCell"), watchedAreas);

    rule.selectionTitle = String(cell("selectionTitle") ?? "").trim();
    if (rule.selectionTitle === "") {
      throw new Error("SelectionTitle is empty.");
    }
  } else if (action === ACTION_REFRESH_ON_ACTIVATE) {
    const [area] = watchedAreas;
    if (watchedAreas.length !== 1 || area.top !== area.bottom || area.left !== area.right) {
      throw new Error(`WatchedCells must be a single cell for ${action}, e.g. "C3".`);
    }
    rule.watchedAddress = `${columnLetters(area.left)}${area.top}`;

    // May lie on another worksheet, like the targets of toggle-on-click.
    const targets = parseToggleTargets(cell("targetCells"), worksheet.name);
    if (targets.length !== 1) {
      throw new Error(`TargetCells must be a single cell for ${action}.`);
    }
    const targetSheet = worksheetsByName.get(targets[0].worksheetName.toLowerCase());
    if (!targetSheet) {
      throw new Error(`Worksheet "${targets[0].worksheetName}" of TargetCells was not found.`);
    }
    rule.appliedTarget = { worksheetName: targetSheet.name, address: targets[0].address };
  }

  return rule;
}

/**
 * The value a toggle-on-click rule writes: the second value when the current
 * value is the first one (ignoring case and surrounding spaces), otherwise the
 * first value.
 */
export function nextToggleValue(currentValue, toggleValues) {
  const current = String(currentValue ?? "").trim().toLowerCase();
  return current === toggleValues[0].toLowerCase() ? toggleValues[1] : toggleValues[0];
}

/**
 * Splits a data-validation list source ("a,b,c" or "a;b;c") into its trimmed,
 * non-empty entries.
 */
export function splitListSource(source) {
  return String(source ?? "")
    .split(/[,;]/)
    .map((entry) => entry.trim())
    .filter((entry) => entry !== "");
}

/**
 * The entries ticked when the popup opens for a cell value: the entries the
 * value names (comma-separated, case ignored), or all of them when the value
 * is the "every entry" value or names none of them.
 */
export function initiallySelected(cellValue, allValue, entries) {
  const text = String(cellValue ?? "").trim().toLowerCase();
  if (text === allValue.toLowerCase()) {
    return [...entries];
  }

  const named = new Set(splitListSource(text));
  const selected = entries.filter((entry) => named.has(entry.toLowerCase()));
  return selected.length > 0 ? selected : [...entries];
}

/**
 * The value written for the ticked entries: allValue when every entry is
 * ticked, otherwise the ticked entries joined by ", " in list order.
 */
export function multiSelectValue(selected, allValue, entries) {
  const ticked = new Set(selected);
  const inOrder = entries.filter((entry) => ticked.has(entry));
  return inOrder.length === entries.length ? allValue : inOrder.join(", ");
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

/** Parses "Table[Column]" of filter-table into its table and column names. */
function parseTableColumn(value) {
  const match = /^\s*([^[\]]+?)\s*\[\s*([^[\]]+?)\s*\]\s*$/.exec(String(value ?? ""));
  if (!match) {
    throw new Error('TargetCells must name a table column, e.g. "Table1[Region]".');
  }
  return { tableName: match[1], tableColumn: match[2] };
}

/** Columns that toggle-on-click and multi-select-on-click must leave empty. */
const CLICK_UNUSED_COLUMNS = [
  "delayMs",
  "filterColumnIndex",
  "filterValue",
  "sortColumnIndex",
  "sortOrder",
  "columnChoices",
];

/** Columns that refresh-on-activate must leave empty. */
const ACTIVATE_UNUSED_COLUMNS = [
  ...CLICK_UNUSED_COLUMNS,
  "targetValue",
  "selectCell",
  "selectionTitle",
];

/** Parses "first;second" into two trimmed values that differ (ignoring case). */
function parseToggleValues(value) {
  const values = String(value ?? "")
    .split(";")
    .map((part) => part.trim());

  if (values.length !== 2 || values.some((part) => part === "")) {
    throw new Error('TargetValue must hold exactly two values separated by ";", e.g. "On;Off".');
  }
  if (values[0].toLowerCase() === values[1].toLowerCase()) {
    throw new Error("The two values of TargetValue must be different.");
  }

  return values;
}

/** Parses SelectCell into one normalized address outside the watched cells. */
function parseSelectCell(value, watchedAreas) {
  if (isBlank(value)) {
    throw new Error("SelectCell is empty.");
  }

  const areas = parseAddressList(value);
  if (areas.length !== 1 || areas[0].top !== areas[0].bottom || areas[0].left !== areas[0].right) {
    throw new Error('SelectCell must be a single cell, e.g. "A1".');
  }
  if (areasIntersect(areas, watchedAreas)) {
    throw new Error("SelectCell must lie outside WatchedCells.");
  }

  return `${columnLetters(areas[0].left)}${areas[0].top}`;
}

/** Parses TargetCells into normalized single-cell addresses, e.g. ["B8", "F8"]. */
function parseTargetCells(value) {
  if (isBlank(value)) {
    throw new Error("TargetCells is empty.");
  }

  return parseAddressList(value).map((area) => {
    if (area.top !== area.bottom || area.left !== area.right) {
      throw new Error("TargetCells must list single cells, e.g. \"B8,F8\".");
    }
    return `${columnLetters(area.left)}${area.top}`;
  });
}

/**
 * Parses the TargetCells of toggle-on-click: single cells, each on the
 * rule's worksheet or with a worksheet prefix, e.g. "C3,'Other sheet'!D3".
 * The first target holds the value the toggle switches from.
 * @returns {{ worksheetName: string, address: string }[]}
 */
function parseToggleTargets(value, worksheetName) {
  if (isBlank(value)) {
    throw new Error("TargetCells is empty.");
  }

  return splitAddressList(value).map((part) => {
    const { worksheet, address } = splitWorksheetPrefix(part);
    if (worksheet === "") {
      throw new Error(`TargetCells has no worksheet name in "${part}".`);
    }
    return { worksheetName: worksheet ?? worksheetName, address: parseTargetCells(address)[0] };
  });
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
