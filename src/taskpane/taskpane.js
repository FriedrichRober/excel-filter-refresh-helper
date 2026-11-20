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
 *
 * reset-on-click rules are triggered by selection instead of change: when the
 * selection lies entirely inside the watched cells (e.g. a merged "button"
 * cell), the add-in writes the rule's value into its target cells and selects
 * the first target cell, so the next click on the button is a new selection.
 * These rules are not run by the task pane's Run button.
 */

/* global console, document, Excel, Office, performance, window */

import { areasInside, areasIntersect, columnLetters, parseAddressList } from "./address";
import {
  ACTION_FILTER_VALUE,
  ACTION_RESET_ON_CLICK,
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
// Change events from the add-in's own sort can arrive shortly after it ends.
const OWN_SORT_GRACE_MS = 1000;

/** @type {Map<string, object[]>} rules grouped by worksheet ID */
let rulesByWorksheetId = new Map();
let configurationFound = false;

const registeredWorksheetIds = new Set();
const selectionRegisteredWorksheetIds = new Set();
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

    await registerEventHandlers(
      worksheetIdsWithRules(activeRules, (rule) => !isSelectionRule(rule)),
      worksheetIdsWithRules(activeRules, isSelectionRule)
    );

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

function worksheetIdsWithRules(rules, predicate) {
  return [...new Set(rules.filter(predicate).map((rule) => rule.worksheetId))];
}

/** Rules triggered by selecting their watched cells instead of changing them. */
function isSelectionRule(rule) {
  return rule.action === ACTION_RESET_ON_CLICK;
}

function scheduleConfigurationReload() {
  window.clearTimeout(reloadTimer);
  reloadTimer = window.setTimeout(loadConfiguration, RELOAD_DEBOUNCE_MS);
}

// ---------------------------------------------------------------------------
// Event registration
// ---------------------------------------------------------------------------

/**
 * Registers change and calculation listeners once per worksheet with change
 * rules, and selection listeners once per worksheet with selection rules.
 * Handlers look up rules at event time, so reloading the configuration never
 * needs to re-register listeners for worksheets that already have them.
 */
async function registerEventHandlers(changeWorksheetIds, selectionWorksheetIds) {
  const newIds = changeWorksheetIds.filter((id) => !registeredWorksheetIds.has(id));
  let newSelectionIds = selectionWorksheetIds.filter(
    (id) => !selectionRegisteredWorksheetIds.has(id)
  );

  if (newSelectionIds.length > 0 && !Office.context.requirements.isSetSupported("ExcelApi", "1.7")) {
    console.error("Selection events (ExcelApi 1.7) are not supported; reset-on-click rules are inactive.");
    newSelectionIds = [];
  }

  if (newIds.length === 0 && newSelectionIds.length === 0 && worksheetAddedHandlerRegistered) {
    return;
  }

  await Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;

    for (const id of newIds) {
      const worksheet = worksheets.getItem(id);
      worksheet.onChanged.add(handleWorksheetChanged);
      worksheet.onCalculated.add(handleWorksheetCalculated);
    }

    for (const id of newSelectionIds) {
      worksheets.getItem(id).onSelectionChanged.add(handleWorksheetSelectionChanged);
    }

    if (!worksheetAddedHandlerRegistered) {
      worksheets.onAdded.add(handleWorksheetAdded);
    }

    await context.sync();
  });

  newIds.forEach((id) => registeredWorksheetIds.add(id));
  newSelectionIds.forEach((id) => selectionRegisteredWorksheetIds.add(id));
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
  // Selection rules never react to changes, e.g. the writes of reset-on-click.
  const rules = (rulesByWorksheetId.get(event.worksheetId) || []).filter(
    (rule) => !isSelectionRule(rule)
  );

  if (rules.length === 0) {
    return;
  }

  let changedAreas;
  try {
    changedAreas = parseAddressList(event.address);
  } catch (error) {
    console.debug(`Ignoring change with unsupported address: ${error.message}`);
    return;
  }

  const state = getWorksheetState(event.worksheetId);

  if (isOwnSortChange(state, changedAreas)) {
    console.debug(`Ignoring change from the add-in's own sort: ${event.address}`);
    return;
  }

  const matchingRules = rules.filter((rule) => areasIntersect(rule.watchedAreas, changedAreas));

  if (matchingRules.length === 0) {
    return;
  }
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

/**
 * Sorting rewrites the cells of the sorted range, which raises a change event.
 * If the watched cells lie in that range, the sort would trigger itself again.
 * While the add-in sorts (and shortly after), changes of more than one cell
 * inside the sorted range are treated as its own; single-cell edits still count.
 */
function isOwnSortChange(state, changedAreas) {
  if (!state.ownSortArea) {
    return false;
  }

  if (!state.sorting && performance.now() > state.ownSortIgnoreUntil) {
    state.ownSortArea = null;
    return false;
  }

  const singleCell = changedAreas.length === 1 &&
    changedAreas[0].top === changedAreas[0].bottom &&
    changedAreas[0].left === changedAreas[0].right;

  return !singleCell && areasInside(changedAreas, state.ownSortArea);
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
    runRefresh(worksheetId, rules, { changedAt: state.changedAt, readyAt: performance.now() });
  }
}

function flushImmediateRules(worksheetId) {
  const state = getWorksheetState(worksheetId);

  window.clearTimeout(state.immediateTimer);

  const rules = [...state.immediateRules];
  state.immediateRules.clear();

  if (rules.length > 0) {
    runRefresh(worksheetId, rules, { changedAt: state.immediateChangedAt, readyAt: performance.now() });
  }
}

function maxDelay(rules) {
  return Math.max(0, ...[...rules].map((rule) => rule.delayMs));
}

// ---------------------------------------------------------------------------
// Selection (reset-on-click)
// ---------------------------------------------------------------------------

async function handleWorksheetSelectionChanged(event) {
  const rules = (rulesByWorksheetId.get(event.worksheetId) || []).filter(isSelectionRule);

  if (rules.length === 0) {
    return;
  }

  let selectedAreas;
  try {
    selectedAreas = parseAddressList(event.address);
  } catch (error) {
    console.debug(`Ignoring selection with unsupported address: ${error.message}`);
    return;
  }

  // Containment, not intersection: selecting the whole sheet, a whole row or
  // dragging across the button must not trigger it. The add-in's own select()
  // of a target cell lies outside the button and is ignored here as well.
  const matchingRules = rules.filter((rule) =>
    selectedAreas.every((selected) =>
      rule.watchedAreas.some((watched) => areasInside([selected], watched))
    )
  );

  if (matchingRules.length === 0) {
    return;
  }

  const state = getWorksheetState(event.worksheetId);
  if (state.resetting) {
    return;
  }

  state.resetting = true;
  try {
    await resetTargetCells(event.worksheetId, matchingRules);
  } finally {
    state.resetting = false;
  }
}

/**
 * Writes each rule's value into its target cells and selects the first target
 * cell, moving the selection off the button so the next click fires again.
 */
async function resetTargetCells(worksheetId, rules) {
  const values = new Map();
  for (const rule of rules) {
    rule.targetAddresses.forEach((address) => values.set(address, rule.targetValue));
  }
  const firstAddress = rules[0].targetAddresses[0];

  try {
    await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getItem(worksheetId);

      for (const [address, value] of values) {
        worksheet.getRange(address).values = [[value]];
      }
      worksheet.getRange(firstAddress).select();

      await context.sync();
    });

    setStatus(`Reset ${pluralize(values.size, "cell")} on ${rules[0].worksheetName}.`);
  } catch (error) {
    console.error(error, error.debugInfo);
    setStatus(`Reset failed: ${describeError(error)}`);
  }
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
      queuedTiming: null,
      sorting: false,
      resetting: false,
      ownSortArea: null,
      ownSortIgnoreUntil: 0,
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
    state.queuedTiming = null;
  }
}

/**
 * Runs the rules for one worksheet. While a refresh for that worksheet is
 * running, later requests are merged into a single queued refresh. Other
 * worksheets are not blocked.
 */
async function runRefresh(worksheetId, rules, timing) {
  const state = getWorksheetState(worksheetId);

  if (state.running) {
    rules.forEach((rule) => state.queuedRules.add(rule));
    // Keep the earliest change so the reported total covers the whole wait.
    state.queuedTiming = state.queuedTiming
      ? { changedAt: Math.min(state.queuedTiming.changedAt, timing.changedAt), readyAt: state.queuedTiming.readyAt }
      : timing;
    return;
  }

  state.running = true;

  try {
    await applyRules(worksheetId, rules, timing);
  } finally {
    state.running = false;
  }

  if (state.queuedRules.size > 0) {
    const queuedRules = [...state.queuedRules];
    const queuedTiming = state.queuedTiming;
    state.queuedRules.clear();
    state.queuedTiming = null;
    await runRefresh(worksheetId, queuedRules, queuedTiming);
  }
}

// ---------------------------------------------------------------------------
// Actions
// ---------------------------------------------------------------------------

async function applyRules(worksheetId, rules, timing) {
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

  const state = getWorksheetState(worksheetId);
  const sortTiming = { sortMs: 0, filterMs: 0 };

  try {
    await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getItem(worksheetId);
      const autoFilter = worksheet.autoFilter;
      const filterRange = autoFilter.getRangeOrNullObject();
      const usesFilter = sortRule !== null || valuesByColumn.size > 0;
      // Filtering is usually allowed on protected sheets; sorting and hiding
      // columns need the protection lifted for the change.
      const needsUnprotect = sortRule !== null || columnRules.length > 0;

      filterRange.load(sortRule ? "isNullObject,columnCount,address" : "isNullObject,columnCount");
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
          state.ownSortArea = parseAddressList(filterRange.address)[0];
          state.sorting = true;
          try {
            await sortKeepingFilters(
              context,
              autoFilter,
              filterRange,
              sortRule,
              configuredFilters(worksheetId, valuesByColumn),
              sortTiming
            );
          } finally {
            state.sorting = false;
            state.ownSortIgnoreUntil = performance.now() + OWN_SORT_GRACE_MS;
          }
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
    const ms = (value) => `${Math.round(value)} ms`;
    const parts = [
      `calculation wait ${ms(timing.readyAt - timing.changedAt)}`,
      `queue ${ms(stepStartedAt - timing.readyAt)}`,
      `Excel step ${ms(finishedAt - stepStartedAt)}`,
    ];
    if (sortRule && sortTiming.filterMs > 0) {
      parts.push(`sort ${ms(sortTiming.sortMs)}`, `filters one by one ${ms(sortTiming.filterMs)}`);
    } else if (sortRule) {
      parts.push(`sort and filters ${ms(sortTiming.sortMs)}`);
    }
    const text = `${doneText} in ${ms(finishedAt - timing.changedAt)} (${parts.join(", ")}).`;

    console.debug(text);
    setStatus(text);
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
 * The filters to restore after a sort: every filter-value rule configured for
 * the worksheet, so the filters on other columns come back too. Filters set by
 * hand in Excel are not kept; reading them (AutoFilter.criteria) is slow and
 * often fails with an internal error.
 *
 * @returns {Map<number, Set<string>>} values by filter column index
 */
function configuredFilters(worksheetId, triggeredValuesByColumn) {
  const filters = new Map();

  for (const rule of rulesByWorksheetId.get(worksheetId) || []) {
    if (rule.action === ACTION_FILTER_VALUE) {
      if (!filters.has(rule.filterColumnIndex)) {
        filters.set(rule.filterColumnIndex, new Set());
      }
      filters.get(rule.filterColumnIndex).add(rule.filterValue);
    }
  }

  // Rules triggered in this run win, e.g. after a configuration reload.
  for (const [columnIndex, values] of triggeredValuesByColumn) {
    filters.set(columnIndex, values);
  }

  return filters;
}

/**
 * Excel sorts only the rows a filter shows, so all filters are cleared, the
 * range is sorted and the configured filters are applied again. All of it runs
 * in one batch; if re-applying fails, each filter is tried on its own, so a
 * failing filter shows more rows instead of stopping the refresh.
 */
async function sortKeepingFilters(context, autoFilter, filterRange, sortRule, filters, timing) {
  const startedAt = performance.now();

  autoFilter.clearCriteria();
  filterRange.sort.apply(
    [{ key: sortRule.sortColumnIndex, ascending: sortRule.ascending }],
    false,
    true
  );

  try {
    applyFilters(autoFilter, filterRange, filters);
    await context.sync();
    timing.sortMs = performance.now() - startedAt;
    return;
  } catch (error) {
    console.error("Sorting and re-applying the filters in one batch failed; retrying in steps.", error, error.debugInfo);
  }

  // Operations after the failing one were not run; repeat them step by step.
  autoFilter.clearCriteria();
  filterRange.sort.apply(
    [{ key: sortRule.sortColumnIndex, ascending: sortRule.ascending }],
    false,
    true
  );
  await context.sync();

  const filtersStartedAt = performance.now();
  timing.sortMs = filtersStartedAt - startedAt;

  for (const [columnIndex, values] of filters) {
    try {
      applyFilters(autoFilter, filterRange, new Map([[columnIndex, values]]));
      await context.sync();
    } catch (columnError) {
      console.error(`Filter on column index ${columnIndex} could not be applied again.`, columnError);
    }
  }

  timing.filterMs = performance.now() - filtersStartedAt;
}

function applyFilters(autoFilter, filterRange, filters) {
  for (const [columnIndex, values] of filters) {
    autoFilter.apply(filterRange, columnIndex, {
      filterOn: Excel.FilterOn.values,
      values: [...values],
    });
  }
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

  // reset-on-click rules only run when their button is clicked.
  const rules = (rulesByWorksheetId.get(worksheetId) || []).filter(
    (rule) => !isSelectionRule(rule)
  );

  if (rules.length === 0) {
    setStatus("No filter refresh rule is configured for the active worksheet.");
    return;
  }

  // The manual run covers anything still waiting.
  clearWaitingWork(getWorksheetState(worksheetId));

  const now = performance.now();
  await runRefresh(worksheetId, rules, { changedAt: now, readyAt: now });
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
