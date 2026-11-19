/*
 * Configurable Excel filter refresh helper.
 *
 * Reads rules from a configuration worksheet. When a watched cell changes, the
 * add-in waits for the worksheet calculation and then re-applies a value filter
 * on one column of the worksheet's AutoFilter, leaving other columns untouched.
 */

/* global console, document, Excel, Office, performance, window */

import { areasIntersect, parseAddressList } from "./address";
import {
  ACTION_FILTER_VALUE,
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

  await loadConfiguration();
});

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
  matchingRules.forEach((rule) => state.pendingRules.add(rule));
  state.changedAt = performance.now();

  window.clearTimeout(state.debounceTimer);
  window.clearTimeout(state.fallbackTimer);
  state.fallbackTimer = window.setTimeout(
    () => flushPendingRules(event.worksheetId),
    CALCULATION_FALLBACK_MS
  );

  setStatus("Change detected. Waiting for calculation.");
}

async function handleWorksheetCalculated(event) {
  const state = worksheetStates.get(event.worksheetId);

  if (!state || state.pendingRules.size === 0) {
    return;
  }

  const delayMs = Math.max(...[...state.pendingRules].map((rule) => rule.delayMs));

  window.clearTimeout(state.debounceTimer);
  state.debounceTimer = window.setTimeout(() => flushPendingRules(event.worksheetId), delayMs);
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
      running: false,
      queuedRules: new Set(),
      queuedStartedAt: 0,
    });
  }

  return worksheetStates.get(worksheetId);
}

function resetPendingWork() {
  for (const state of worksheetStates.values()) {
    window.clearTimeout(state.debounceTimer);
    window.clearTimeout(state.fallbackTimer);
    state.pendingRules.clear();
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
// Filter action
// ---------------------------------------------------------------------------

async function applyRules(worksheetId, rules, startedAt) {
  // Merge rules into one value filter per column so each column is applied once.
  const valuesByColumn = new Map();

  for (const rule of rules) {
    if (rule.action !== ACTION_FILTER_VALUE) {
      continue;
    }
    if (!valuesByColumn.has(rule.filterColumnIndex)) {
      valuesByColumn.set(rule.filterColumnIndex, new Set());
    }
    valuesByColumn.get(rule.filterColumnIndex).add(rule.filterValue);
  }

  if (valuesByColumn.size === 0) {
    return;
  }

  const filterStartedAt = performance.now();
  setStatus("Refreshing filter...");

  try {
    await Excel.run(async (context) => {
      const autoFilter = context.workbook.worksheets.getItem(worksheetId).autoFilter;
      const filterRange = autoFilter.getRangeOrNullObject();

      filterRange.load("isNullObject,columnCount");
      await context.sync();

      if (filterRange.isNullObject) {
        throw new Error("The worksheet has no AutoFilter range.");
      }

      for (const columnIndex of valuesByColumn.keys()) {
        if (columnIndex >= filterRange.columnCount) {
          throw new Error(
            `Filter column index ${columnIndex} is outside the AutoFilter range ` +
              `(${pluralize(filterRange.columnCount, "column")}).`
          );
        }
      }

      // apply() with a column index only replaces that column's criteria;
      // it is much faster than autoFilter.reapply().
      for (const [columnIndex, values] of valuesByColumn) {
        autoFilter.apply(filterRange, columnIndex, {
          filterOn: Excel.FilterOn.values,
          values: [...values],
        });
      }

      await context.sync();
    });

    const finishedAt = performance.now();
    const totalMs = Math.round(finishedAt - (startedAt || filterStartedAt));
    const filterMs = Math.round(finishedAt - filterStartedAt);

    console.debug(`Filter refresh: total ${totalMs} ms, filter step ${filterMs} ms.`);
    setStatus(`Filter refreshed in ${totalMs} ms; filter step ${filterMs} ms.`);
  } catch (error) {
    console.error(error);
    setStatus(`Refresh failed: ${error.message}`);
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

  const rules = rulesByWorksheetId.get(worksheetId);

  if (!rules) {
    setStatus("No filter refresh rule is configured for the active worksheet.");
    return;
  }

  // The manual run covers anything still waiting for calculation.
  const state = getWorksheetState(worksheetId);
  window.clearTimeout(state.debounceTimer);
  window.clearTimeout(state.fallbackTimer);
  state.pendingRules.clear();

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
