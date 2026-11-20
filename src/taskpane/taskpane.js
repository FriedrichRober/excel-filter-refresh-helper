/*
 * Configurable Excel filter refresh helper.
 *
 * Reads rules from a configuration worksheet. When a watched cell changes, the
 * add-in runs the rule's action on that worksheet:
 *
 *   filter-value  re-applies a value filter on one column of the AutoFilter,
 *                 leaving other columns untouched (after recalculation)
 *   sort          sorts the AutoFilter range by one column and keeps all
 *                 filters (after recalculation)
 *   show-columns  shows the columns the watched cell's value selects and hides
 *                 the other columns of the rule (right away)
 */

/* global console, document, Excel, Office, performance, window */

import { areasIntersect, columnLetters, parseAddressList } from "./address";
import {
  ACTION_FILTER_VALUE,
  ACTION_SHOW_COLUMNS,
  ACTION_SORT,
  CONFIG_SHEET_NAME,
  PLATFORM_DESKTOP,
  PLATFORM_WEB,
  parseConfiguration,
} from "./config";

// Run pending rules anyway if no calculation event follows a watched change.
const CALCULATION_FALLBACK_MS = 2000;
const RELOAD_DEBOUNCE_MS = 300;

/** @type {Map<string, object[]>} rules grouped by worksheet ID */
let rulesByWorksheetId = new Map();
let configurationFound = false;

const registeredWorksheetIds = new Set();
let worksheetAddedHandlerRegistered = false;
let reloadTimer = null;

/** Per-worksheet runtime state, keyed by worksheet ID. */
const worksheetStates = new Map();

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Excel) {
    return;
  }

  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "block";
  document.getElementById("run").onclick = refreshActiveWorksheet;
  document.getElementById("reload").onclick = loadConfiguration;
  // eslint-disable-next-line no-undef
  document.getElementById("version").textContent =
    // eslint-disable-next-line no-undef
    `Version ${__ADDIN_VERSION__} · built ${formatBuildTime(__BUILD_TIME__)}`;

  await loadConfiguration();
});

// Formats the build timestamp as "YYYY-MM-DD hh:mm" in the viewer's local time.
function formatBuildTime(ms) {
  const d = new Date(ms);
  const pad = (n) => String(n).padStart(2, "0");
  return (
    `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ` +
    `${pad(d.getHours())}:${pad(d.getMinutes())}`
  );
}

// ---------------------------------------------------------------------------
// Configuration
// ---------------------------------------------------------------------------

async function loadConfiguration() {
  setStatus("Loading configuration...");
  setProblems([]);

  try {
    const result = await Excel.run(async (context) => {
      const worksheets = context.workbook.worksheets;
      const configSheet = worksheets.getItemOrNullObject(CONFIG_SHEET_NAME);

      worksheets.load("items/id,items/name");
      configSheet.load("isNullObject");
      await context.sync();

      if (configSheet.isNullObject) {
        return { found: false };
      }

      const usedRange = configSheet.getUsedRangeOrNullObject(true);
      usedRange.load("isNullObject,values,rowIndex");
      await context.sync();

      const sheets = worksheets.items.map((worksheet) => ({
        id: worksheet.id,
        name: worksheet.name,
      }));

      if (usedRange.isNullObject) {
        return { found: true, rules: [], errors: [] };
      }

      return {
        found: true,
        ...parseConfiguration(usedRange.values, sheets, usedRange.rowIndex + 1),
      };
    });

    // The startup setting is stored in the workbook and shared by all platforms,
    // so it depends only on the configuration sheet, not on the current platform.
    await updateStartupBehavior(result.found);

    resetPendingWork();
    configurationFound = result.found;

    const platform = getCurrentPlatform();
    const activeRules = result.found
      ? result.rules.filter((rule) => rule.platforms.includes(platform))
      : [];
    rulesByWorksheetId = groupRulesByWorksheet(activeRules);

    if (!result.found) {
      setStatus("No filter refresh configuration was found.");
      return;
    }

    result.errors.forEach((message) => console.error(`Filter refresh configuration: ${message}`));
    setProblems(result.errors);

    await registerEventHandlers([...rulesByWorksheetId.keys()]);

    const ruleCount = activeRules.length;
    const invalidText =
      result.errors.length > 0
        ? ` Configuration contains ${pluralize(result.errors.length, "problem")}.`
        : "";

    if (ruleCount === 0 && result.rules.length > 0) {
      setStatus(`Automatic filter refresh is not enabled for this version of Excel.${invalidText}`);
    } else if (ruleCount === 0) {
      setStatus(`No active filter refresh rules were found.${invalidText}`);
    } else {
      setStatus(
        `Automatic filter refresh is active (${pluralize(ruleCount, "rule")} on ` +
          `${pluralize(rulesByWorksheetId.size, "worksheet")}).${invalidText}`
      );
    }
  } catch (error) {
    console.error(error);
    setStatus(`Configuration could not be loaded: ${error.message}`);
  }
}

/**
 * With the shared runtime, the add-in keeps running after the task pane is
 * closed. Loading on document open is a per-document setting, so only enable
 * it for workbooks that contain a configuration sheet.
 */
async function updateStartupBehavior(configured) {
  if (!Office.addin || !Office.context.requirements.isSetSupported("SharedRuntime", "1.1")) {
    return;
  }

  try {
    await Office.addin.setStartupBehavior(
      configured ? Office.StartupBehavior.load : Office.StartupBehavior.none
    );
  } catch (error) {
    console.error("Could not set the add-in startup behavior.", error);
  }
}

function getCurrentPlatform() {
  return Office.context.platform === Office.PlatformType.OfficeOnline
    ? PLATFORM_WEB
    : PLATFORM_DESKTOP;
}

function groupRulesByWorksheet(rules) {
  const grouped = new Map();

  for (const rule of rules) {
    if (!grouped.has(rule.worksheetId)) {
      grouped.set(rule.worksheetId, []);
    }
    grouped.get(rule.worksheetId).push(rule);
  }

  return grouped;
}

function scheduleConfigurationReload() {
  window.clearTimeout(reloadTimer);
  reloadTimer = window.setTimeout(loadConfiguration, RELOAD_DEBOUNCE_MS);
}

// ---------------------------------------------------------------------------
// Event registration
// ---------------------------------------------------------------------------

/**
 * Registers change and calculation listeners once per worksheet. Handlers look
 * up rules at event time, so reloading the configuration never needs to
 * re-register listeners for worksheets that already have them.
 */
async function registerEventHandlers(worksheetIds) {
  const newIds = worksheetIds.filter((id) => !registeredWorksheetIds.has(id));

  if (newIds.length === 0 && worksheetAddedHandlerRegistered) {
    return;
  }

  await Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;

    for (const id of newIds) {
      const worksheet = worksheets.getItem(id);
      worksheet.onChanged.add(handleWorksheetChanged);
      worksheet.onCalculated.add(handleWorksheetCalculated);
    }

    if (!worksheetAddedHandlerRegistered) {
      worksheets.onAdded.add(handleWorksheetAdded);
    }

    await context.sync();
  });

  newIds.forEach((id) => registeredWorksheetIds.add(id));
  worksheetAddedHandlerRegistered = true;
}

async function handleWorksheetAdded() {
  // A configured worksheet may have been added; re-read the configuration.
  scheduleConfigurationReload();
}

// ---------------------------------------------------------------------------
// Change and calculation sequence
// ---------------------------------------------------------------------------

/** Actions that depend on formula results and therefore wait for recalculation. */
function waitsForCalculation(rule) {
  return rule.action !== ACTION_SHOW_COLUMNS;
}

async function handleWorksheetChanged(event) {
  const rules = rulesByWorksheetId.get(event.worksheetId);

  if (!rules) {
    return;
  }

  let changedAreas;
  try {
    changedAreas = parseAddressList(event.address);
  } catch (error) {
    console.debug(`Ignoring change with unsupported address: ${error.message}`);
    return;
  }

  const matchingRules = rules.filter((rule) => areasIntersect(rule.watchedAreas, changedAreas));

  if (matchingRules.length === 0) {
    return;
  }

  const state = getWorksheetState(event.worksheetId);
  const immediateRules = matchingRules.filter((rule) => !waitsForCalculation(rule));
  const calculatedRules = matchingRules.filter(waitsForCalculation);

  if (immediateRules.length > 0) {
    immediateRules.forEach((rule) => state.immediateRules.add(rule));
    state.immediateChangedAt = performance.now();

    window.clearTimeout(state.immediateTimer);
    state.immediateTimer = window.setTimeout(
      () => flushImmediateRules(event.worksheetId),
      maxDelay(state.immediateRules)
    );
  }

  if (calculatedRules.length > 0) {
    calculatedRules.forEach((rule) => state.pendingRules.add(rule));
    state.changedAt = performance.now();

    window.clearTimeout(state.debounceTimer);
    window.clearTimeout(state.fallbackTimer);
    state.fallbackTimer = window.setTimeout(
      () => flushPendingRules(event.worksheetId),
      CALCULATION_FALLBACK_MS
    );

    setStatus("Change detected. Waiting for calculation.");
  }
}

async function handleWorksheetCalculated(event) {
  const state = worksheetStates.get(event.worksheetId);

  if (!state || state.pendingRules.size === 0) {
    return;
  }

  window.clearTimeout(state.debounceTimer);
  state.debounceTimer = window.setTimeout(
    () => flushPendingRules(event.worksheetId),
    maxDelay(state.pendingRules)
  );
}

function flushPendingRules(worksheetId) {
  const state = getWorksheetState(worksheetId);

  window.clearTimeout(state.debounceTimer);
  window.clearTimeout(state.fallbackTimer);

  const rules = [...state.pendingRules];
  state.pendingRules.clear();

  if (rules.length > 0) {
    runRefresh(worksheetId, rules, state.changedAt);
  }
}

function flushImmediateRules(worksheetId) {
  const state = getWorksheetState(worksheetId);

  window.clearTimeout(state.immediateTimer);

  const rules = [...state.immediateRules];
  state.immediateRules.clear();

  if (rules.length > 0) {
    runRefresh(worksheetId, rules, state.immediateChangedAt);
  }
}

function maxDelay(rules) {
  return Math.max(0, ...[...rules].map((rule) => rule.delayMs));
}

// ---------------------------------------------------------------------------
// Concurrency
// ---------------------------------------------------------------------------

function getWorksheetState(worksheetId) {
  if (!worksheetStates.has(worksheetId)) {
    worksheetStates.set(worksheetId, {
      pendingRules: new Set(),
      changedAt: 0,
      debounceTimer: null,
      fallbackTimer: null,
      immediateRules: new Set(),
      immediateChangedAt: 0,
      immediateTimer: null,
      running: false,
      queuedRules: new Set(),
      queuedStartedAt: 0,
    });
  }

  return worksheetStates.get(worksheetId);
}

function clearWaitingWork(state) {
  window.clearTimeout(state.debounceTimer);
  window.clearTimeout(state.fallbackTimer);
  window.clearTimeout(state.immediateTimer);
  state.pendingRules.clear();
  state.immediateRules.clear();
}

function resetPendingWork() {
  for (const state of worksheetStates.values()) {
    clearWaitingWork(state);
    state.queuedRules.clear();
  }
}

/**
 * Runs the rules for one worksheet. While a refresh for that worksheet is
 * running, later requests are merged into a single queued refresh. Other
 * worksheets are not blocked.
 */
async function runRefresh(worksheetId, rules, startedAt) {
  const state = getWorksheetState(worksheetId);

  if (state.running) {
    rules.forEach((rule) => state.queuedRules.add(rule));
    state.queuedStartedAt = startedAt;
    return;
  }

  state.running = true;

  try {
    await applyRules(worksheetId, rules, startedAt);
  } finally {
    state.running = false;
  }

  if (state.queuedRules.size > 0) {
    const queuedRules = [...state.queuedRules];
    state.queuedRules.clear();
    await runRefresh(worksheetId, queuedRules, state.queuedStartedAt);
  }
}

// ---------------------------------------------------------------------------
// Actions
// ---------------------------------------------------------------------------

async function applyRules(worksheetId, rules, startedAt) {
  // Merge rules: one value filter per column, the last sort rule, and every
  // column rule, so each piece of work is done once.
  const valuesByColumn = new Map();
  const columnRules = [];
  let sortRule = null;

  for (const rule of rules) {
    if (rule.action === ACTION_FILTER_VALUE) {
      if (!valuesByColumn.has(rule.filterColumnIndex)) {
        valuesByColumn.set(rule.filterColumnIndex, new Set());
      }
      valuesByColumn.get(rule.filterColumnIndex).add(rule.filterValue);
    } else if (rule.action === ACTION_SORT) {
      sortRule = rule;
    } else if (rule.action === ACTION_SHOW_COLUMNS) {
      columnRules.push(rule);
    }
  }

  if (valuesByColumn.size === 0 && !sortRule && columnRules.length === 0) {
    return;
  }

  const stepStartedAt = performance.now();
  const doneText = sortRule
    ? "Sorted"
    : valuesByColumn.size > 0
      ? "Filter refreshed"
      : "Columns updated";
  setStatus(sortRule ? "Sorting..." : valuesByColumn.size > 0 ? "Refreshing filter..." : "Updating columns...");

  let criteriaLost = false;

  try {
    await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getItem(worksheetId);
      const autoFilter = worksheet.autoFilter;
      const filterRange = autoFilter.getRangeOrNullObject();
      const usesFilter = sortRule !== null || valuesByColumn.size > 0;
      // Filtering is usually allowed on protected sheets; sorting and hiding
      // columns need the protection lifted for the change.
      const needsUnprotect = sortRule !== null || columnRules.length > 0;

      filterRange.load("isNullObject,columnCount");
      if (needsUnprotect) {
        worksheet.protection.load("protected,options");
      }
      const choiceCells = columnRules.map((rule) => {
        const area = rule.watchedAreas[0];
        const cell = worksheet.getRange(`${columnLetters(area.left)}${area.top}`);
        cell.load("values");
        return cell;
      });
      await context.sync();

      if (usesFilter) {
        validateFilterColumns(filterRange, [
          ...valuesByColumn.keys(),
          ...(sortRule ? [sortRule.sortColumnIndex] : []),
        ]);
      }

      const protectionOptions =
        needsUnprotect && worksheet.protection.protected ? worksheet.protection.options : null;
      if (protectionOptions) {
        worksheet.protection.unprotect();
      }

      try {
        columnRules.forEach((rule, index) => showColumns(worksheet, rule, choiceCells[index].values[0][0]));

        if (sortRule) {
          criteriaLost = !(await sortKeepingFilters(context, autoFilter, filterRange, sortRule, valuesByColumn));
        } else {
          // apply() with a column index only replaces that column's criteria;
          // it is much faster than autoFilter.reapply().
          for (const [columnIndex, values] of valuesByColumn) {
            autoFilter.apply(filterRange, columnIndex, {
              filterOn: Excel.FilterOn.values,
              values: [...values],
            });
          }
        }

        await context.sync();
      } finally {
        if (protectionOptions) {
          try {
            worksheet.protection.protect(protectionOptions);
            await context.sync();
          } catch (error) {
            console.error("Could not protect the worksheet again.", error);
          }
        }
      }
    });

    const finishedAt = performance.now();
    const totalMs = Math.round(finishedAt - (startedAt || stepStartedAt));
    const stepMs = Math.round(finishedAt - stepStartedAt);

    console.debug(`${doneText}: total ${totalMs} ms, Excel step ${stepMs} ms.`);
    setStatus(
      `${doneText} in ${totalMs} ms; Excel step ${stepMs} ms.` +
        (criteriaLost
          ? " Excel could not read the existing filters, so only the configured filters were applied again."
          : "")
    );
  } catch (error) {
    console.error(error, error.debugInfo);
    setStatus(`Refresh failed: ${describeError(error)}`);
  }
}

/** Adds the failing Excel operation (e.g. "Range.sort.apply") to the message. */
function describeError(error) {
  const location = error && error.debugInfo && error.debugInfo.errorLocation;
  return location ? `${error.message} (${location})` : String(error && error.message);
}

function validateFilterColumns(filterRange, columnIndexes) {
  if (filterRange.isNullObject) {
    throw new Error("The worksheet has no AutoFilter range.");
  }

  for (const columnIndex of columnIndexes) {
    if (columnIndex >= filterRange.columnCount) {
      throw new Error(
        `Column index ${columnIndex} is outside the AutoFilter range ` +
          `(${pluralize(filterRange.columnCount, "column")}).`
      );
    }
  }
}

/**
 * Shows the columns of the choice in the watched cell and hides the columns of
 * the other choices. A value that is not one of the choices shows them all.
 */
function showColumns(worksheet, rule, value) {
  const shown = rule.columnChoices.get(String(value ?? "").trim().toLowerCase());
  const allColumns = [...new Set([...rule.columnChoices.values()].flat())].sort((a, b) => a - b);

  // Set adjacent columns with the same state in one range.
  let runStart = null;
  let runHidden = null;
  let previous = null;

  const flush = () => {
    if (runStart !== null) {
      const address = `${columnLetters(runStart)}:${columnLetters(previous)}`;
      worksheet.getRange(address).columnHidden = runHidden;
    }
  };

  for (const column of allColumns) {
    const hidden = shown ? !shown.includes(column) : false;

    if (runStart === null || column !== previous + 1 || hidden !== runHidden) {
      flush();
      runStart = column;
      runHidden = hidden;
    }
    previous = column;
  }
  flush();
}

/**
 * Excel sorts only the rows a filter shows, so the criteria are saved, cleared
 * for the sort and applied again. Columns with a filter-value rule get that
 * rule's value instead of the saved one. A saved criterion that cannot be
 * applied again is skipped, so the sheet shows more rows instead of stopping.
 *
 * Excel sometimes fails to return the criteria (an internal error on
 * AutoFilter.criteria). The criteria are read in their own sync so that this
 * does not stop the refresh; the sort then keeps only the filter-value rules.
 *
 * @returns {Promise<boolean>} false when the saved criteria could not be read.
 */
async function sortKeepingFilters(context, autoFilter, filterRange, sortRule, valuesByColumn) {
  let criteria = [];
  let criteriaRead = true;

  try {
    autoFilter.load("criteria");
    await context.sync();
    criteria = (autoFilter.criteria || []).map(cleanCriteria);
  } catch (error) {
    criteriaRead = false;
    console.error("Could not read the AutoFilter criteria; keeping only the filter-value rules.", error, error.debugInfo);
  }

  for (const [columnIndex, values] of valuesByColumn) {
    criteria[columnIndex] = { filterOn: Excel.FilterOn.values, values: [...values] };
  }

  autoFilter.clearCriteria();
  filterRange.sort.apply(
    [{ key: sortRule.sortColumnIndex, ascending: sortRule.ascending }],
    false,
    true
  );
  await context.sync();

  const filtered = [];
  criteria.forEach((criterion, columnIndex) => {
    if (criterion) {
      filtered.push([columnIndex, criterion]);
      autoFilter.apply(filterRange, columnIndex, criterion);
    }
  });

  try {
    await context.sync();
  } catch (error) {
    console.error("Re-applying the filters at once failed; applying them one by one.", error);

    for (const [columnIndex, criterion] of filtered) {
      try {
        autoFilter.apply(filterRange, columnIndex, criterion);
        await context.sync();
      } catch (columnError) {
        console.error(`Filter on column index ${columnIndex} could not be applied again.`, columnError);
      }
    }
  }

  return criteriaRead;
}

function cleanCriteria(criterion) {
  if (!criterion || !criterion.filterOn) {
    return null;
  }

  const cleaned = {};
  for (const [key, value] of Object.entries(criterion)) {
    if (value !== undefined && value !== null && value !== "") {
      cleaned[key] = value;
    }
  }
  return cleaned;
}

// ---------------------------------------------------------------------------
// Manual fallback
// ---------------------------------------------------------------------------

async function refreshActiveWorksheet() {
  if (!configurationFound) {
    setStatus("No filter refresh configuration was found.");
    return;
  }

  let worksheetId;

  try {
    worksheetId = await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      worksheet.load("id");
      await context.sync();
      return worksheet.id;
    });
  } catch (error) {
    console.error(error);
    setStatus(`Refresh failed: ${error.message}`);
    return;
  }

  const rules = rulesByWorksheetId.get(worksheetId);

  if (!rules) {
    setStatus("No filter refresh rule is configured for the active worksheet.");
    return;
  }

  // The manual run covers anything still waiting.
  clearWaitingWork(getWorksheetState(worksheetId));

  await runRefresh(worksheetId, rules, performance.now());
}

// ---------------------------------------------------------------------------
// UI helpers
// ---------------------------------------------------------------------------

function setStatus(message) {
  const status = document.getElementById("status");

  if (status) {
    status.textContent = message;
  }
}

function setProblems(messages) {
  const list = document.getElementById("problems");

  if (!list) {
    return;
  }

  list.replaceChildren(
    ...messages.map((message) => {
      const item = document.createElement("li");
      item.textContent = message;
      return item;
    })
  );
  list.hidden = messages.length === 0;
}

function pluralize(count, noun) {
  return `${count} ${noun}${count === 1 ? "" : "s"}`;
}
