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
 *
 * toggle-on-click rules are triggered by selection too: the add-in switches
 * the target cells between the rule's two values (the first target cell
 * decides), selects the rule's SelectCell and then runs the other rules that
 * watch the written cells, as if the user had typed the value. It does not
 * wait for the change event of its own write and ignores it if it comes.
 *
 * Selection rules are not run by the task pane's Run button.
 *
 * In Excel for the web, the add-in also writes the active column filters of
 * the worksheets listed in a status worksheet into it (see filterStatus.js),
 * after every filter change and after its own actions.
 */

/* global console, document, Excel, Office, performance, window */

import { areasInside, areasIntersect, columnLetters, parseAddressList } from "./address";
import {
  ACTION_FILTER_VALUE,
  ACTION_RESET_ON_CLICK,
  ACTION_SHOW_COLUMNS,
  ACTION_SORT,
  ACTION_TOGGLE_ON_CLICK,
  CONFIG_SHEET_NAME,
  DEFAULT_DELAY_MS,
  PLATFORM_DESKTOP,
  PLATFORM_WEB,
  nextToggleValue,
  parseConfiguration,
} from "./config";
import { STATUS_SHEET_NAME, formatFilterStatus, parseExcludedColumns, parseStatusRows } from "./filterStatus";

// Run pending rules anyway if no calculation event follows a watched change.
const CALCULATION_FALLBACK_MS = 2000;
const RELOAD_DEBOUNCE_MS = 300;
// Change events from the add-in's own sort can arrive shortly after it ends.
const OWN_SORT_GRACE_MS = 1000;
// The change event of the add-in's own toggle write may arrive late, if at all.
const OWN_TOGGLE_GRACE_MS = 2000;
// Temporary diagnostics: run each step of a sort in its own sync and report
// the time of every step. Slower overall; set to false for normal use.
const SORT_DIAGNOSTICS = false;

/** @type {Map<string, object[]>} rules grouped by worksheet ID */
let rulesByWorksheetId = new Map();
let configurationFound = false;

/** @type {Map<string, number>} 1-based status sheet row by worksheet ID */
let statusRowsByWorksheetId = new Map();

const changeRegisteredWorksheetIds = new Set();
const calculatedRegisteredWorksheetIds = new Set();
const filteredRegisteredWorksheetIds = new Set();
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
      const statusSheet = worksheets.getItemOrNullObject(STATUS_SHEET_NAME);

      worksheets.load("items/id,items/name");
      configSheet.load("isNullObject");
      statusSheet.load("isNullObject");
      await context.sync();

      const configRange = configSheet.isNullObject ? null : configSheet.getUsedRangeOrNullObject(true);
      const statusNames = statusSheet.isNullObject
        ? null
        : statusSheet.getRange("A:A").getUsedRangeOrNullObject(true);
      configRange?.load("isNullObject,values,rowIndex");
      statusNames?.load("isNullObject,values,rowIndex");
      if (configRange || statusNames) {
        await context.sync();
      }

      const sheets = worksheets.items.map((worksheet) => ({
        id: worksheet.id,
        name: worksheet.name,
      }));

      // Worksheets not listed in the status sheet are ignored.
      const statusRows = new Map();
      if (statusNames && !statusNames.isNullObject) {
        const rowsByName = parseStatusRows(statusNames.values, statusNames.rowIndex + 1);
        for (const sheet of sheets) {
          if (rowsByName.has(sheet.name)) {
            statusRows.set(sheet.id, rowsByName.get(sheet.name));
          }
        }
      }

      if (!configRange) {
        return { found: false, statusFound: !!statusNames, statusRows };
      }
      if (configRange.isNullObject) {
        return { found: true, statusFound: !!statusNames, statusRows, rules: [], errors: [] };
      }

      return {
        found: true,
        statusFound: !!statusNames,
        statusRows,
        ...parseConfiguration(configRange.values, sheets, configRange.rowIndex + 1),
      };
    });

    // The startup setting is stored in the workbook and shared by all platforms,
    // so it depends only on the configuration and status sheets, not on the
    // current platform.
    await updateStartupBehavior(result.found || result.statusFound);

    resetPendingWork();
    configurationFound = result.found;

    const platform = getCurrentPlatform();
    const activeRules = result.found
      ? result.rules.filter((rule) => rule.platforms.includes(platform))
      : [];
    rulesByWorksheetId = groupRulesByWorksheet(activeRules);
    // On desktop, a macro writes the filter status.
    statusRowsByWorksheetId = platform === PLATFORM_WEB ? result.statusRows : new Map();

    const changeIds = worksheetIdsWithRules(activeRules, (rule) => !isSelectionRule(rule));
    await registerEventHandlers(
      changeIds,
      [...new Set([...changeIds, ...statusRowsByWorksheetId.keys()])],
      [...statusRowsByWorksheetId.keys()],
      worksheetIdsWithRules(activeRules, isSelectionRule)
    );
    // Bring the status up to date with the filters the workbook was opened with.
    statusRowsByWorksheetId.forEach((row, worksheetId) => scheduleFilterStatusUpdate(worksheetId));

    if (!result.found) {
      setStatus("No filter refresh configuration was found.");
      return;
    }

    result.errors.forEach((message) => console.error(`Filter refresh configuration: ${message}`));
    setProblems(result.errors);

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
  return rule.action === ACTION_RESET_ON_CLICK || rule.action === ACTION_TOGGLE_ON_CLICK;
}

function scheduleConfigurationReload() {
  window.clearTimeout(reloadTimer);
  reloadTimer = window.setTimeout(loadConfiguration, RELOAD_DEBOUNCE_MS);
}

// ---------------------------------------------------------------------------
// Event registration
// ---------------------------------------------------------------------------

/**
 * Registers each kind of listener once per worksheet that needs it: change
 * listeners for worksheets with change rules, calculation listeners for those
 * and for worksheets listed in the status sheet, filter listeners for listed
 * worksheets, and selection listeners for worksheets with selection rules.
 * Handlers look up rules at event time, so reloading the configuration never
 * needs to re-register listeners for worksheets that already have them.
 */
async function registerEventHandlers(
  changeWorksheetIds,
  calculatedWorksheetIds,
  filteredWorksheetIds,
  selectionWorksheetIds
) {
  const notIn = (registered) => (id) => !registered.has(id);
  const newChangeIds = changeWorksheetIds.filter(notIn(changeRegisteredWorksheetIds));
  const newCalculatedIds = calculatedWorksheetIds.filter(notIn(calculatedRegisteredWorksheetIds));
  // A filter change also recalculates the listed worksheet when its status
  // formula depends on the visible rows; onFiltered (ExcelApi 1.13) catches
  // the change even when it does not.
  const newFilteredIds = Office.context.requirements.isSetSupported("ExcelApi", "1.13")
    ? filteredWorksheetIds.filter(notIn(filteredRegisteredWorksheetIds))
    : [];
  let newSelectionIds = selectionWorksheetIds.filter(notIn(selectionRegisteredWorksheetIds));

  if (newSelectionIds.length > 0 && !Office.context.requirements.isSetSupported("ExcelApi", "1.7")) {
    console.error("Selection events (ExcelApi 1.7) are not supported; reset-on-click and toggle-on-click rules are inactive.");
    newSelectionIds = [];
  }

  if (
    newChangeIds.length === 0 &&
    newCalculatedIds.length === 0 &&
    newFilteredIds.length === 0 &&
    newSelectionIds.length === 0 &&
    worksheetAddedHandlerRegistered
  ) {
    return;
  }

  await Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;

    for (const id of newChangeIds) {
      worksheets.getItem(id).onChanged.add(handleWorksheetChanged);
    }

    for (const id of newCalculatedIds) {
      worksheets.getItem(id).onCalculated.add(handleWorksheetCalculated);
    }

    for (const id of newFilteredIds) {
      worksheets.getItem(id).onFiltered.add(handleWorksheetFiltered);
    }

    for (const id of newSelectionIds) {
      worksheets.getItem(id).onSelectionChanged.add(handleWorksheetSelectionChanged);
    }

    if (!worksheetAddedHandlerRegistered) {
      worksheets.onAdded.add(handleWorksheetAdded);
    }

    await context.sync();
  });

  newChangeIds.forEach((id) => changeRegisteredWorksheetIds.add(id));
  newCalculatedIds.forEach((id) => calculatedRegisteredWorksheetIds.add(id));
  newFilteredIds.forEach((id) => filteredRegisteredWorksheetIds.add(id));
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

  if (isOwnToggleChange(state, changedAreas, event)) {
    // The toggle already queued the rules watching these cells.
    console.debug(`Ignoring change from the add-in's own toggle: ${event.address}`);
    return;
  }

  queueChangedRules(
    event.worksheetId,
    rules.filter((rule) => areasIntersect(rule.watchedAreas, changedAreas))
  );
}

/**
 * Queues change rules as after a change of their watched cells: rules that
 * wait for calculation run after the next calculation (or the fallback), the
 * others after their delay. Rules already waiting are merged, not repeated.
 */
function queueChangedRules(worksheetId, matchingRules) {
  if (matchingRules.length === 0) {
    return;
  }

  const state = getWorksheetState(worksheetId);
  const immediateRules = matchingRules.filter((rule) => !waitsForCalculation(rule));
  const calculatedRules = matchingRules.filter(waitsForCalculation);

  if (immediateRules.length > 0) {
    immediateRules.forEach((rule) => state.immediateRules.add(rule));
    state.immediateChangedAt = performance.now();

    window.clearTimeout(state.immediateTimer);
    state.immediateTimer = window.setTimeout(
      () => flushImmediateRules(worksheetId),
      maxDelay(state.immediateRules)
    );
  }

  if (calculatedRules.length > 0) {
    calculatedRules.forEach((rule) => state.pendingRules.add(rule));
    state.changedAt = performance.now();

    window.clearTimeout(state.debounceTimer);
    window.clearTimeout(state.fallbackTimer);
    state.fallbackTimer = window.setTimeout(
      () => flushPendingRules(worksheetId),
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

/**
 * After a toggle, the add-in has already queued the rules watching the cells
 * it wrote. If Excel raises a change event for that write anyway, it is
 * ignored for a short while, unless Excel says it came from somewhere else.
 */
function isOwnToggleChange(state, changedAreas, event) {
  if (!state.ownToggleAreas) {
    return false;
  }

  if (performance.now() > state.ownToggleIgnoreUntil) {
    state.ownToggleAreas = null;
    return false;
  }

  // triggerSource (ExcelApi 1.14) is "ThisLocalAddin" for the add-in's write.
  if (event.triggerSource !== undefined && event.triggerSource !== "ThisLocalAddin") {
    return false;
  }

  return changedAreas.every((changed) =>
    state.ownToggleAreas.some((written) => areasInside([changed], written))
  );
}

async function handleWorksheetCalculated(event) {
  scheduleFilterStatusUpdate(event.worksheetId);

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
// Selection (reset-on-click, toggle-on-click)
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
  // of a target cell or SelectCell lies outside the button and is ignored here
  // as well.
  const matchingRules = rules.filter((rule) =>
    selectedAreas.every((selected) =>
      rule.watchedAreas.some((watched) => areasInside([selected], watched))
    )
  );

  if (matchingRules.length === 0) {
    return;
  }

  const state = getWorksheetState(event.worksheetId);
  if (state.clicking) {
    return;
  }

  state.clicking = true;
  try {
    await writeClickedTargets(event.worksheetId, matchingRules);
  } finally {
    state.clicking = false;
  }
}

/**
 * Writes the target cells of the clicked rules: reset-on-click writes its
 * value, toggle-on-click switches to its other value. Then the selection moves
 * off the button so the next click fires again: to the first target cell for
 * reset-on-click, to SelectCell for toggle-on-click (the first clicked rule
 * decides).
 *
 * The change rules watching the toggled cells are queued as if the value had
 * been typed. That happens before the write is sent, so the calculation event
 * that follows the write is not missed, and the change event of the write is
 * ignored if Excel raises it.
 */
async function writeClickedTargets(worksheetId, rules) {
  const toggleRules = rules.filter((rule) => rule.action === ACTION_TOGGLE_ON_CLICK);
  const values = new Map();
  const selectAddress =
    rules[0].action === ACTION_TOGGLE_ON_CLICK ? rules[0].selectAddress : rules[0].targetAddresses[0];

  try {
    await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getItem(worksheetId);

      const currentCells = toggleRules.map((rule) => {
        const range = worksheet.getRange(rule.targetAddresses[0]);
        range.load("values");
        return range;
      });
      if (currentCells.length > 0) {
        await context.sync();
      }

      for (const rule of rules) {
        if (rule.action === ACTION_RESET_ON_CLICK) {
          rule.targetAddresses.forEach((address) => values.set(address, rule.targetValue));
        }
      }
      const toggledAreas = [];
      toggleRules.forEach((rule, index) => {
        const value = nextToggleValue(currentCells[index].values[0][0], rule.toggleValues);
        rule.targetAddresses.forEach((address) => values.set(address, value));
        toggledAreas.push(...rule.targetAreas);
      });

      for (const [address, value] of values) {
        worksheet.getRange(address).values = [[value]];
      }
      worksheet.getRange(selectAddress).select();

      if (toggledAreas.length > 0) {
        queueToggledRules(worksheetId, toggledAreas);
      }

      await context.sync();
    });
  } catch (error) {
    console.error(error, error.debugInfo);
    setStatus(`${toggleRules.length > 0 ? "Toggle" : "Reset"} failed: ${describeError(error)}`);
    return;
  }

  if (toggleRules.length === 0) {
    setStatus(`Reset ${pluralize(values.size, "cell")} on ${rules[0].worksheetName}.`);
  } else if (!hasRefreshWork(getWorksheetState(worksheetId))) {
    // With dependent rules, their progress is shown instead.
    const address = toggleRules[0].targetAddresses[0];
    setStatus(`Switched ${address} on ${rules[0].worksheetName} to "${values.get(address)}".`);
  }
}

function queueToggledRules(worksheetId, toggledAreas) {
  const state = getWorksheetState(worksheetId);
  state.ownToggleAreas = toggledAreas;
  state.ownToggleIgnoreUntil = performance.now() + OWN_TOGGLE_GRACE_MS;

  queueChangedRules(
    worksheetId,
    (rulesByWorksheetId.get(worksheetId) || []).filter(
      (rule) => !isSelectionRule(rule) && areasIntersect(rule.watchedAreas, toggledAreas)
    )
  );
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
      statusTimer: null,
      statusUpdating: false,
      statusRequested: false,
      running: false,
      queuedRules: new Set(),
      queuedTiming: null,
      sorting: false,
      clicking: false,
      ownSortArea: null,
      ownSortIgnoreUntil: 0,
      ownToggleAreas: null,
      ownToggleIgnoreUntil: 0,
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

/** True while a refresh of the worksheet is running or waiting to run. */
function hasRefreshWork(state) {
  return (
    state.running ||
    state.sorting ||
    state.pendingRules.size > 0 ||
    state.immediateRules.size > 0 ||
    state.queuedRules.size > 0
  );
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
    return;
  }

  // The add-in's own filter and sort change the filters too.
  scheduleFilterStatusUpdate(worksheetId);
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
  const diagnostics = SORT_DIAGNOSTICS && sortRule ? { steps: [] } : null;
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

      const choiceCells = columnRules.map((rule) =>
        worksheet.getRange(`${columnLetters(rule.watchedAreas[0].left)}${rule.watchedAreas[0].top}`)
      );
      const queueLoads = (withCriteria) => {
        filterRange.load(
          sortRule
            ? `isNullObject,columnCount,address${diagnostics ? ",rowCount" : ""}`
            : "isNullObject,columnCount"
        );
        if (withCriteria) {
          autoFilter.load("criteria");
        }
        if (needsUnprotect) {
          worksheet.protection.load("protected,options");
        }
        if (diagnostics) {
          context.workbook.application.load("calculationMode");
        }
        choiceCells.forEach((cell) => cell.load("values"));
      };

      // A sort clears the filters, so the active ones are read to restore them
      // afterwards. They are loaded with everything else to avoid an extra
      // round trip. Excel sometimes fails on AutoFilter.criteria with an
      // internal error; then the load is repeated without them and only the
      // configured filters are restored.
      const loadStartedAt = performance.now();
      let activeCriteria = null;
      if (sortRule) {
        try {
          queueLoads(true);
          await context.sync();
          activeCriteria = autoFilter.criteria || [];
        } catch (error) {
          console.error("Could not read the active filters; restoring only the configured ones.", error, error.debugInfo);
          criteriaLost = true;
          queueLoads(false);
          await context.sync();
        }
      } else {
        queueLoads(false);
        await context.sync();
      }
      if (diagnostics) {
        diagnostics.steps.push(`load ${formatMs(performance.now() - loadStartedAt)}`);
        diagnostics.context = [
          `${filterRange.rowCount} rows x ${filterRange.columnCount} columns`,
          `calculation ${context.workbook.application.calculationMode}`,
          `protected ${worksheet.protection.protected}`,
        ];
      }

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
            const filters = activeCriteria
              ? filtersToRestore(activeCriteria, valuesByColumn)
              : configuredFilters(worksheetId, valuesByColumn);
            if (diagnostics) {
              await sortWithDiagnostics(context, autoFilter, filterRange, sortRule, filters, diagnostics);
            } else {
              await sortKeepingFilters(context, autoFilter, filterRange, sortRule, filters, sortTiming);
            }
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
            const protectStartedAt = performance.now();
            await context.sync();
            if (diagnostics) {
              diagnostics.steps.push(`protect ${formatMs(performance.now() - protectStartedAt)}`);
            }
          } catch (error) {
            console.error("Could not protect the worksheet again.", error);
          }
        }
      }
    });

    const finishedAt = performance.now();
    const ms = formatMs;
    const parts = [
      `calculation wait ${ms(timing.readyAt - timing.changedAt)}`,
      `queue ${ms(stepStartedAt - timing.readyAt)}`,
      `Excel step ${ms(finishedAt - stepStartedAt)}`,
    ];
    if (sortRule && sortTiming.filterMs > 0) {
      parts.push(`sort ${ms(sortTiming.sortMs)}`, `filters one by one ${ms(sortTiming.filterMs)}`);
    } else if (sortRule && !diagnostics) {
      parts.push(`sort and filters ${ms(sortTiming.sortMs)}`);
    }
    let text = `${doneText} in ${ms(finishedAt - timing.changedAt)} (${parts.join(", ")}).`;
    if (criteriaLost) {
      text += " Excel could not read the active filters, so only the configured filters were applied again.";
    }
    if (diagnostics) {
      text += ` Diagnostics:${diagnostics.steps.join(", ")}. [${diagnostics.context.join(", ")}]`;
    }

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
 * The filters to restore after a sort: the filters active in Excel (including
 * ones set by hand), with the filter-value rules triggered in this run
 * replacing the criteria of their columns.
 *
 * @returns {Map<number, object>} filter criteria by column index
 */
function filtersToRestore(activeCriteria, triggeredValuesByColumn) {
  const filters = new Map();

  activeCriteria.forEach((criterion, columnIndex) => {
    const cleaned = cleanCriteria(criterion);
    if (cleaned) {
      filters.set(columnIndex, cleaned);
    }
  });

  for (const [columnIndex, values] of triggeredValuesByColumn) {
    filters.set(columnIndex, valuesCriteria(values));
  }

  return filters;
}

/**
 * Fallback when the active filters cannot be read: every filter-value rule
 * configured for the worksheet, so the filters on other columns come back too.
 * Filters set by hand in Excel are lost in this case.
 *
 * @returns {Map<number, object>} filter criteria by column index
 */
function configuredFilters(worksheetId, triggeredValuesByColumn) {
  const valuesByColumn = new Map();

  for (const rule of rulesByWorksheetId.get(worksheetId) || []) {
    if (rule.action === ACTION_FILTER_VALUE) {
      if (!valuesByColumn.has(rule.filterColumnIndex)) {
        valuesByColumn.set(rule.filterColumnIndex, new Set());
      }
      valuesByColumn.get(rule.filterColumnIndex).add(rule.filterValue);
    }
  }

  // Rules triggered in this run win, e.g. after a configuration reload.
  for (const [columnIndex, values] of triggeredValuesByColumn) {
    valuesByColumn.set(columnIndex, values);
  }

  const filters = new Map();
  for (const [columnIndex, values] of valuesByColumn) {
    filters.set(columnIndex, valuesCriteria(values));
  }
  return filters;
}

function valuesCriteria(values) {
  return { filterOn: Excel.FilterOn.values, values: [...values] };
}

/**
 * Criteria read from Excel contain empty properties that apply() rejects;
 * keep only the ones that are set. Returns null for a column without a filter.
 */
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

/**
 * Excel sorts only the rows a filter shows, so all filters are cleared, the
 * range is sorted and the saved filters are applied again. All of it runs
 * in one batch; if re-applying fails, each filter is tried on its own, so a
 * failing filter shows more rows instead of stopping the refresh.
 *
 * Suspending calculation during the batch (suspendApiCalculationUntilNextSync
 * plus a recalculation afterwards) was measured slower in Excel for the web,
 * so Excel calculates as it normally would.
 */
async function sortKeepingFilters(context, autoFilter, filterRange, sortRule, filters, timing) {
  const queueClearAndSort = () => {
    autoFilter.clearCriteria();
    filterRange.sort.apply(
      [{ key: sortRule.sortColumnIndex, ascending: sortRule.ascending }],
      false,
      true
    );
  };
  const startedAt = performance.now();

  queueClearAndSort();
  let inOneBatch = true;
  try {
    applyFilters(autoFilter, filterRange, filters);
    await context.sync();
    timing.sortMs = performance.now() - startedAt;
  } catch (error) {
    console.error("Sorting and re-applying the filters in one batch failed; retrying in steps.", error, error.debugInfo);
    inOneBatch = false;
  }

  if (!inOneBatch) {
    // Operations after the failing one were not run; repeat them step by step.
    queueClearAndSort();
    await context.sync();

    const filtersStartedAt = performance.now();
    timing.sortMs = filtersStartedAt - startedAt;

    for (const [columnIndex, criterion] of filters) {
      try {
        applyFilters(autoFilter, filterRange, new Map([[columnIndex, criterion]]));
        await context.sync();
      } catch (columnError) {
        console.error(`Filter on column index ${columnIndex} could not be applied again.`, columnError);
      }
    }

    timing.filterMs = performance.now() - filtersStartedAt;
  }
}

/**
 * Same work as sortKeepingFilters, but every step runs in its own sync so its
 * time can be measured. An empty sync before and after shows the cost of one
 * round trip and any work Excel still does after the last step.
 */
async function sortWithDiagnostics(context, autoFilter, filterRange, sortRule, filters, diagnostics) {
  const step = async (name, queue) => {
    queue();
    const startedAt = performance.now();
    try {
      await context.sync();
      diagnostics.steps.push(`${name} ${formatMs(performance.now() - startedAt)}`);
    } catch (error) {
      diagnostics.steps.push(`${name} FAILED ${formatMs(performance.now() - startedAt)}`);
      throw error;
    }
  };

  await step("empty sync", () => {});
  await step("clear criteria", () => autoFilter.clearCriteria());
  await step("sort", () =>
    filterRange.sort.apply([{ key: sortRule.sortColumnIndex, ascending: sortRule.ascending }], false, true)
  );

  for (const [columnIndex, criterion] of filters) {
    const label = criterion.values
      ? pluralize(criterion.values.length, "value")
      : criterion.filterOn;
    try {
      await step(`filter col ${columnIndex} (${label})`, () =>
        applyFilters(autoFilter, filterRange, new Map([[columnIndex, criterion]]))
      );
    } catch (error) {
      console.error(`Filter on column index ${columnIndex} could not be applied again.`, error);
    }
  }

  await step("empty sync after", () => {});
}

function formatMs(value) {
  return `${Math.round(value)} ms`;
}

function applyFilters(autoFilter, filterRange, filters) {
  for (const [columnIndex, criterion] of filters) {
    autoFilter.apply(filterRange, columnIndex, criterion);
  }
}

// ---------------------------------------------------------------------------
// Filter status
// ---------------------------------------------------------------------------

async function handleWorksheetFiltered(event) {
  scheduleFilterStatusUpdate(event.worksheetId);
}

/**
 * Debounces status updates. Several calculation and filter events follow one
 * filter change, and writing the status causes another calculation; that one
 * finds the status unchanged and writes nothing, so it does not loop.
 */
function scheduleFilterStatusUpdate(worksheetId) {
  if (!statusRowsByWorksheetId.has(worksheetId)) {
    return;
  }

  const state = getWorksheetState(worksheetId);
  window.clearTimeout(state.statusTimer);
  state.statusTimer = window.setTimeout(() => runFilterStatusUpdate(worksheetId), DEFAULT_DELAY_MS);
}

async function runFilterStatusUpdate(worksheetId) {
  const state = getWorksheetState(worksheetId);

  // The refresh schedules another update when it is done.
  if (hasRefreshWork(state)) {
    return;
  }
  if (state.statusUpdating) {
    state.statusRequested = true;
    return;
  }

  state.statusUpdating = true;
  try {
    await updateFilterStatus(worksheetId);
  } finally {
    state.statusUpdating = false;
  }

  if (state.statusRequested) {
    state.statusRequested = false;
    scheduleFilterStatusUpdate(worksheetId);
  }
}

/**
 * Writes the filter text (B) and the visible row count it belongs to (C) into
 * the worksheet's status row. C is the value of D read after the filters, so
 * the workbook's formula only shows B while the filters are still the same.
 * If the filters cannot be read, nothing is written.
 */
async function updateFilterStatus(worksheetId) {
  const row = statusRowsByWorksheetId.get(worksheetId);
  if (!row) {
    return;
  }

  try {
    await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getItem(worksheetId);
      const autoFilter = worksheet.autoFilter;
      const filterRange = autoFilter.getRangeOrNullObject();
      const statusRow = context.workbook.worksheets.getItem(STATUS_SHEET_NAME).getRange(`B${row}:E${row}`);

      // Loads run in order: the visible row count (D) is read after the filters.
      filterRange.load("isNullObject");
      autoFilter.load("criteria");
      statusRow.load("values");
      await context.sync();

      let text = "";
      if (!filterRange.isNullObject) {
        const headerRow = filterRange.getRow(0);
        headerRow.load("values,text");
        await context.sync();

        // Header text as shown, unless the column is too narrow to show it.
        const headers = headerRow.values[0].map((value, index) =>
          typeof value === "string" ? value : headerRow.text[0][index]
        );
        const [, , , excludedText] = statusRow.values[0];
        text = formatFilterStatus(headers, autoFilter.criteria || [], parseExcludedColumns(excludedText));
      }

      const [writtenText, writtenCount, visibleCount] = statusRow.values[0];
      if (writtenText === text && writtenCount === visibleCount) {
        return;
      }
      // A refresh may have started meanwhile; it schedules a new update.
      if (hasRefreshWork(getWorksheetState(worksheetId))) {
        return;
      }

      context.workbook.worksheets
        .getItem(STATUS_SHEET_NAME)
        .getRange(`B${row}:C${row}`).values = [[text, visibleCount]];
      await context.sync();
    });
  } catch (error) {
    // E.g. Excel failing on AutoFilter.criteria with an internal error. The
    // workbook then shows only the row count, so this is not reported.
    console.debug("Filter status was not updated.", error, error.debugInfo);
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

  // reset-on-click and toggle-on-click rules only run when their cell is clicked.
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
