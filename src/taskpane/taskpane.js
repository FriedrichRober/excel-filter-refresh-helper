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
 * multi-select-on-click rules are triggered by selection too: the add-in reads
 * the entries of the watched cell's data-validation list and opens a popup
 * (dialog.html) with one checkbox per entry. On OK it writes the ticked entries
 * (or the rule's "every entry" value) into the target cell and runs the rules
 * watching it, like toggle-on-click. After OK or Cancel it selects SelectCell.
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
  ACTION_FILTER_TABLE,
  ACTION_MULTI_SELECT_ON_CLICK,
  ACTION_RESET_ON_CLICK,
  ACTION_SHOW_COLUMNS,
  ACTION_SORT,
  ACTION_TOGGLE_ON_CLICK,
  CONFIG_SHEET_NAME,
  DEFAULT_DELAY_MS,
  PLATFORM_DESKTOP,
  PLATFORM_WEB,
  initiallySelected,
  multiSelectValue,
  nextToggleValue,
  parseConfiguration,
  splitListSource,
} from "./config";
import {
  STATUS_SHEET_NAME,
  formatFilterStatus,
  isActiveCriterion,
  parseExcludedColumns,
  parseStatusRows,
} from "./filterStatus";

// Run pending rules anyway if no calculation event follows a watched change.
const CALCULATION_FALLBACK_MS = 2000;
const RELOAD_DEBOUNCE_MS = 300;
// Change events from the add-in's own sort can arrive shortly after it ends.
const OWN_SORT_GRACE_MS = 1000;
// The change event of the add-in's own toggle or popup write may arrive late,
// if at all.
const OWN_WRITE_GRACE_MS = 2000;
// Size of the multi-select popup, in percent of the Excel window.
const POPUP_SIZE = { width: 30, height: 65 };
// Temporary diagnostics: run each step of a refresh (sort, filters, columns)
// in its own sync and report the time of every step. Slower overall; set to
// false for normal use.
const DIAGNOSTICS = false;
// How long after a filter change calculations still update the filter status.
const STATUS_WATCH_MS = 3000;

/** @type {Map<string, object[]>} rules grouped by worksheet ID */
let rulesByWorksheetId = new Map();
let configurationFound = false;

/** @type {Map<string, number>} 1-based status sheet row by worksheet ID */
let statusRowsByWorksheetId = new Map();

const changeRegisteredWorksheetIds = new Set();
const calculatedRegisteredWorksheetIds = new Set();
const filteredRegisteredWorksheetIds = new Set();
// Listed worksheets without a filter event (see followFilterStatus).
const statusFollowsEveryCalculation = new Set();
const selectionRegisteredWorksheetIds = new Set();
let worksheetAddedHandlerRegistered = false;
let reloadTimer = null;
// Office shows one dialog at a time; clicks while it is open are ignored.
let popupOpen = false;

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
  return (
    rule.action === ACTION_RESET_ON_CLICK ||
    rule.action === ACTION_TOGGLE_ON_CLICK ||
    rule.action === ACTION_MULTI_SELECT_ON_CLICK
  );
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
  // formula depends on the visible rows; a filter event catches the change
  // even when it does not. Worksheets have no onFiltered in every Excel
  // version, so the event is looked up per worksheet (see filterEventOf).
  const newFilteredIds = filteredWorksheetIds.filter(notIn(filteredRegisteredWorksheetIds));
  let newSelectionIds = selectionWorksheetIds.filter(notIn(selectionRegisteredWorksheetIds));

  if (newSelectionIds.length > 0 && !Office.context.requirements.isSetSupported("ExcelApi", "1.7")) {
    console.error("Selection events (ExcelApi 1.7) are not supported; click rules are inactive.");
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
      const filterEvent = filterEventOf(worksheets.getItem(id));
      if (filterEvent) {
        filterEvent.add(handleWorksheetFiltered);
      } else {
        // Only calculations can tell that its filters changed.
        statusFollowsEveryCalculation.add(id);
      }
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

/**
 * The worksheet event raised when its filter changes: onFiltered where Excel
 * has it, otherwise onRowHiddenChanged (ExcelApi 1.11), as filtering hides
 * rows; null when neither exists. Missing events must not stop the other
 * listeners from being registered.
 */
function filterEventOf(worksheet) {
  if (worksheet.onFiltered) {
    return worksheet.onFiltered;
  }
  if (Office.context.requirements.isSetSupported("ExcelApi", "1.11") && worksheet.onRowHiddenChanged) {
    return worksheet.onRowHiddenChanged;
  }
  return null;
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
  return rule.action !== ACTION_SHOW_COLUMNS && rule.action !== ACTION_FILTER_TABLE;
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

  if (isOwnWriteChange(state, changedAreas, event)) {
    // The toggle or popup already queued the rules watching these cells.
    console.debug(`Ignoring change from the add-in's own write: ${event.address}`);
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
 * After a toggle or popup write, the add-in has already queued the rules
 * watching the cells it wrote. If Excel raises a change event for that write
 * anyway, it is ignored for a short while, unless Excel says it came from
 * somewhere else.
 */
function isOwnWriteChange(state, changedAreas, event) {
  if (!state.ownWriteAreas) {
    return false;
  }

  if (performance.now() > state.ownWriteIgnoreUntil) {
    state.ownWriteAreas = null;
    return false;
  }

  // triggerSource (ExcelApi 1.14) is "ThisLocalAddin" for the add-in's write.
  if (event.triggerSource !== undefined && event.triggerSource !== "ThisLocalAddin") {
    return false;
  }

  return changedAreas.every((changed) =>
    state.ownWriteAreas.some((written) => areasInside([changed], written))
  );
}

async function handleWorksheetCalculated(event) {
  followFilterStatus(event.worksheetId);

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
// Selection (reset-on-click, toggle-on-click, multi-select-on-click)
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

  if (matchingRules.length === 0 || popupOpen) {
    return;
  }

  // A popup rule is not combined with other click rules on the same cell.
  const popupRule = matchingRules.find((rule) => rule.action === ACTION_MULTI_SELECT_ON_CLICK);
  if (popupRule) {
    await runMultiSelect(event.worksheetId, popupRule);
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
  // Toggle targets on other worksheets: name -> (address -> value).
  const otherValues = new Map();
  let switchedTo = null;
  const ownName = rules[0].worksheetName;
  const selectAddress =
    rules[0].action === ACTION_TOGGLE_ON_CLICK ? rules[0].selectAddress : rules[0].targetAddresses[0];

  try {
    await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getItem(worksheetId);
      const sheetNamed = (name) => (name === ownName ? worksheet : context.workbook.worksheets.getItem(name));

      // The first target holds the value the toggle switches from.
      const currentCells = toggleRules.map((rule) => {
        const range = sheetNamed(rule.targets[0].worksheetName).getRange(rule.targets[0].address);
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
        switchedTo = switchedTo ?? value;
        for (const target of rule.targets) {
          if (target.worksheetName === ownName) {
            values.set(target.address, value);
            continue;
          }
          if (!otherValues.has(target.worksheetName)) {
            otherValues.set(target.worksheetName, new Map());
          }
          otherValues.get(target.worksheetName).set(target.address, value);
        }
        toggledAreas.push(...rule.targetAreas);
      });

      for (const [address, value] of values) {
        worksheet.getRange(address).values = [[value]];
      }
      for (const [name, cells] of otherValues) {
        const sheet = sheetNamed(name);
        for (const [address, value] of cells) {
          sheet.getRange(address).values = [[value]];
        }
      }
      worksheet.getRange(selectAddress).select();

      if (toggledAreas.length > 0) {
        queueWrittenRules(worksheetId, toggledAreas);
      }
      // The rules of the other worksheets follow their written cells as well.
      for (const [name, cells] of otherValues) {
        const otherId = worksheetIdByName(name);
        if (otherId !== null) {
          queueWrittenRules(otherId, parseAddressList([...cells.keys()].join(",")));
        }
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
    const others = otherValues.size > 0 ? ` and on ${pluralize(otherValues.size, "other worksheet")}` : "";
    setStatus(`Switched the toggle on ${ownName}${others} to "${switchedTo}".`);
  }
}

/** The id of the worksheet named name among those with rules; null when it has none. */
function worksheetIdByName(name) {
  for (const [id, rules] of rulesByWorksheetId) {
    if (rules.some((rule) => rule.worksheetName === name)) {
      return id;
    }
  }
  return null;
}

/**
 * Queues the change rules watching cells the add-in wrote, as if the value
 * had been typed, and marks the write so its own change event is ignored.
 */
function queueWrittenRules(worksheetId, writtenAreas) {
  const state = getWorksheetState(worksheetId);
  state.ownWriteAreas = writtenAreas;
  state.ownWriteIgnoreUntil = performance.now() + OWN_WRITE_GRACE_MS;

  queueChangedRules(
    worksheetId,
    (rulesByWorksheetId.get(worksheetId) || []).filter(
      (rule) => !isSelectionRule(rule) && areasIntersect(rule.watchedAreas, writtenAreas)
    )
  );
}

// ---------------------------------------------------------------------------
// Multi-select popup (multi-select-on-click)
// ---------------------------------------------------------------------------

/**
 * Reads the entries of the watched cell's validation list, opens the popup
 * and writes the result. If the cell is not set up as the rule expects (e.g.
 * a list with another input title), nothing opens and the selection stays,
 * so the cell's own dropdown can still be used.
 */
async function runMultiSelect(worksheetId, rule) {
  if (
    !Office.context.requirements.isSetSupported("DialogApi", "1.2") ||
    !Office.context.requirements.isSetSupported("ExcelApi", "1.8")
  ) {
    setStatus("This version of Excel cannot show the selection popup (DialogApi 1.2, ExcelApi 1.8).");
    return;
  }

  popupOpen = true;
  try {
    let list;
    try {
      list = await readSelectionList(worksheetId, rule);
    } catch (error) {
      console.error(error, error.debugInfo);
      setStatus(`Selection popup not opened: ${describeError(error)}`);
      return;
    }

    let selected;
    try {
      selected = await showMultiSelectPopup({
        title: `${rule.selectionTitle} wählen`,
        entries: list.entries,
        selected: initiallySelected(list.currentValue, rule.allValue, list.entries),
      });
    } catch (error) {
      console.error(error);
      setStatus(error.message);
      selected = null;
    }

    const value = selected ? multiSelectValue(selected, rule.allValue, list.entries) : null;
    await writeMultiSelectResult(worksheetId, rule, value);
  } finally {
    popupOpen = false;
  }
}

/**
 * Loads the validation list of the first watched cell and the current value
 * of the target cell. The list must carry the rule's input title, and its
 * first entry must be the rule's "every entry" value; the other entries are
 * the choices.
 *
 * @returns {Promise<{ entries: string[], currentValue: any }>}
 */
async function readSelectionList(worksheetId, rule) {
  return Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getItem(worksheetId);
    const first = rule.watchedAreas[0];
    const listAddress = `${columnLetters(first.left)}${first.top}`;
    const validation = worksheet.getRange(listAddress).dataValidation;
    const target = worksheet.getRange(rule.targetAddresses[0]);

    validation.load("type,rule,prompt");
    target.load("values");
    await context.sync();

    const source = validation.type === Excel.DataValidationType.list ? validation.rule?.list?.source : null;
    if (source === null || source === undefined) {
      throw new Error(`${listAddress} on ${rule.worksheetName} has no validation list.`);
    }

    const title = String(validation.prompt?.title ?? "").trim();
    if (title.toLowerCase() !== rule.selectionTitle.toLowerCase()) {
      throw new Error(
        `The validation list of ${listAddress} on ${rule.worksheetName} has the input title ` +
          `"${title}", not "${rule.selectionTitle}".`
      );
    }

    const listed = await readListSource(context, worksheet, String(source));
    if (listed.length < 2 || listed[0].toLowerCase() !== rule.allValue.toLowerCase()) {
      throw new Error(
        `The validation list of ${listAddress} on ${rule.worksheetName} must start with ` +
          `"${rule.allValue}" and list at least one more entry.`
      );
    }

    return { entries: [...new Set(listed.slice(1))], currentValue: target.values[0][0] };
  });
}

/**
 * The entries of a list source: either the entries themselves ("a,b" or
 * "a;b") or, starting with "=", a reference to a range or a named range
 * whose cells hold them (Excel limits typed lists to 255 characters).
 */
async function readListSource(context, worksheet, source) {
  if (!source.trim().startsWith("=")) {
    return splitListSource(source);
  }

  const reference = source.trim().slice(1).trim();
  const separator = reference.lastIndexOf("!");
  let range;
  if (separator >= 0) {
    const sheetName = reference
      .slice(0, separator)
      .replace(/^'(.*)'$/, "$1")
      .replace(/''/g, "'");
    range = context.workbook.worksheets.getItem(sheetName).getRange(reference.slice(separator + 1));
  } else if (/^\$?[A-Z]{1,3}\$?\d+(:\$?[A-Z]{1,3}\$?\d+)?$/i.test(reference)) {
    range = worksheet.getRange(reference);
  } else {
    range = context.workbook.names.getItem(reference).getRange();
  }

  range.load("values");
  try {
    await context.sync();
  } catch (error) {
    throw new Error(`The list source "${source}" could not be read (${error.message}).`);
  }

  return range.values
    .flat()
    .map((value) => String(value ?? "").trim())
    .filter((value) => value !== "");
}

/**
 * Opens the popup and resolves with the ticked entries, or null for Cancel,
 * Esc or closing the popup. displayInIframe keeps Excel for the web from
 * treating it as a browser popup window, which could be blocked because the
 * popup does not open from a click inside the add-in.
 */
function showMultiSelectPopup(init) {
  return new Promise((resolve, reject) => {
    const url = new URL("dialog.html", window.location.href).href;

    Office.context.ui.displayDialogAsync(url, { ...POPUP_SIZE, displayInIframe: true }, (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        reject(new Error(`The selection popup could not be opened: ${result.error.message} (${result.error.code}).`));
        return;
      }

      const dialog = result.value;
      let settled = false;
      const settle = (callback) => {
        if (settled) {
          return;
        }
        settled = true;
        try {
          dialog.close();
        } catch (error) {
          // Already closed by the user.
        }
        callback();
      };

      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        let message;
        try {
          message = JSON.parse(arg.message);
        } catch (error) {
          return;
        }

        if (message.type === "ready") {
          dialog.messageChild(JSON.stringify({ type: "init", ...init }));
        } else if (message.type === "ok" && Array.isArray(message.selected)) {
          settle(() => resolve(message.selected.map(String)));
        } else if (message.type === "cancel") {
          settle(() => resolve(null));
        }
      });

      dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
        // 12006: closed by the user, which counts as Cancel.
        if (arg.error === 12006) {
          settle(() => resolve(null));
        } else {
          settle(() => reject(new Error(`The selection popup failed (${arg.error}).`)));
        }
      });
    });
  });
}

/**
 * Writes the popup's value into the target cell, unless it holds that value
 * already or the popup was cancelled (value null), and selects SelectCell so
 * the next click on the watched cell opens the popup again. The rules
 * watching the target cell are queued before the write is sent, as for
 * toggle-on-click. The target cell is expected to be unlocked, so the sheet
 * stays protected.
 */
async function writeMultiSelectResult(worksheetId, rule, value) {
  const address = rule.targetAddresses[0];
  let written = false;

  try {
    await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getItem(worksheetId);
      const target = worksheet.getRange(address);

      if (value !== null) {
        target.load("values");
        await context.sync();

        if (String(target.values[0][0] ?? "").trim() !== value) {
          target.values = [[value]];
          written = true;
        }
      }

      worksheet.getRange(rule.selectAddress).select();
      if (written) {
        queueWrittenRules(worksheetId, rule.targetAreas);
      }

      await context.sync();
    });
  } catch (error) {
    console.error(error, error.debugInfo);
    setStatus(`Selection could not be written: ${describeError(error)}`);
    return;
  }

  if (written && !hasRefreshWork(getWorksheetState(worksheetId))) {
    setStatus(`Set ${address} on ${rule.worksheetName} to "${value}".`);
  } else if (!written) {
    setStatus(value === null ? "Selection cancelled." : `${address} on ${rule.worksheetName} is unchanged.`);
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
      statusTimer: null,
      statusWatchUntil: 0,
      statusUpdating: false,
      statusRequested: false,
      running: false,
      queuedRules: new Set(),
      queuedTiming: null,
      sorting: false,
      clicking: false,
      ownSortArea: null,
      ownSortIgnoreUntil: 0,
      ownWriteAreas: null,
      ownWriteIgnoreUntil: 0,
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

/**
 * filter-table: for each rule, its table column shows only its values while
 * the watched cell holds its TargetValue; otherwise the column's filter is
 * cleared. A slicer on the column follows, as its filter is the table's.
 */
async function applyTableFilters(worksheetId, rules) {
  setStatus("Filtering tables...");
  try {
    await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getItem(worksheetId);
      const watched = rules.map((rule) => {
        const range = worksheet.getRange(rule.watchedAddresses).getCell(0, 0);
        range.load("values");
        return range;
      });
      await context.sync();

      rules.forEach((rule, index) => {
        const value = String(watched[index].values[0][0] ?? "").trim().toLowerCase();
        const filter = context.workbook.tables.getItem(rule.tableName).columns.getItem(rule.tableColumn).filter;
        if (value === rule.filterChoice.toLowerCase()) {
          filter.applyValuesFilter(rule.tableValues);
        } else {
          filter.clear();
        }
      });
      await context.sync();
    });
  } catch (error) {
    console.error(error, error.debugInfo);
    setStatus(`Table filter failed: ${describeError(error)}`);
    return;
  }
  setStatus(`Filtered ${pluralize(rules.length, "table column")}.`);
}

async function applyRules(worksheetId, rules, timing) {
  // Merge rules: one value filter per column, the last sort rule, and every
  // column rule, so each piece of work is done once.
  const valuesByColumn = new Map();
  const columnRules = [];
  let sortRule = null;

  // Table filters first: they need no AutoFilter on this worksheet.
  const tableRules = rules.filter((rule) => rule.action === ACTION_FILTER_TABLE);
  const tableStartedAt = performance.now();
  if (tableRules.length > 0) {
    await applyTableFilters(worksheetId, tableRules);
  }
  const tableMs = performance.now() - tableStartedAt;

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
  const diagnostics = DIAGNOSTICS ? { steps: [] } : null;
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
          `isNullObject,columnCount${sortRule ? ",address" : ""}${diagnostics ? ",rowCount" : ""}`
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
          filterRange.isNullObject
            ? "no AutoFilter"
            : `${filterRange.rowCount} rows x ${filterRange.columnCount} columns`,
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
        if (diagnostics) {
          const unprotectStartedAt = performance.now();
          await context.sync();
          diagnostics.steps.push(`unprotect ${formatMs(performance.now() - unprotectStartedAt)}`);
        }
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

        const finalStartedAt = performance.now();
        await context.sync();
        if (diagnostics) {
          const label = sortRule ? "final sync" : valuesByColumn.size > 0 ? "filters and columns" : "columns";
          diagnostics.steps.push(`${label} ${formatMs(performance.now() - finalStartedAt)}`);
        }
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
    // "queue" is the wait behind an earlier refresh of the same worksheet.
    const parts = [
      `calculation wait ${ms(timing.readyAt - timing.changedAt)}`,
      `queue ${ms(tableStartedAt - timing.readyAt)}`,
    ];
    if (tableRules.length > 0) {
      parts.push(`table filters ${ms(tableMs)}`);
    }
    parts.push(`Excel step ${ms(finishedAt - stepStartedAt)}`);
    if (sortRule && sortTiming.batchError) {
      parts.push(`sort ${ms(sortTiming.sortMs)}`, `filters retried ${ms(sortTiming.filterMs)}`);
    } else if (sortRule && !diagnostics) {
      parts.push(`sort and filters ${ms(sortTiming.sortMs)}`);
    }
    let text = `${doneText} in ${ms(finishedAt - timing.changedAt)} (${parts.join(", ")}).`;
    if (sortTiming.batchError) {
      text += ` Re-applying the filters in one batch failed: ${sortTiming.batchError}.`;
    }
    if (sortTiming.failedColumns?.length > 0) {
      text +=
        ` The filter on column index ${sortTiming.failedColumns.join(", ")} could not be applied again` +
        " and was left off.";
    }
    if (criteriaLost) {
      text += " Excel could not read the active filters, so only the configured filters were applied again.";
    }
    if (diagnostics) {
      text += ` Diagnostics: ${diagnostics.steps.join(", ")}. [${diagnostics.context.join(", ")}]`;
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
  if (!isActiveCriterion(criterion)) {
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
 * in one batch; if re-applying fails, the filters are retried in shrinking
 * groups (see applyFiltersBisecting), so a failing filter shows more rows
 * instead of stopping the refresh.
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
  let batchError = null;
  try {
    applyFilters(autoFilter, filterRange, filters);
    await context.sync();
    timing.sortMs = performance.now() - startedAt;
  } catch (error) {
    console.error("Sorting and re-applying the filters in one batch failed; retrying in steps.", error, error.debugInfo);
    batchError = error;
  }

  if (batchError) {
    timing.batchError = describeError(batchError);

    // A batch is not transactional: the operations before the failing one
    // stay applied. When a filter failed, the range is already sorted, so
    // only the filters are repeated; otherwise the sort is repeated too.
    const location = String(batchError.debugInfo?.errorLocation ?? "");
    if (!/AutoFilter\.apply/i.test(location)) {
      queueClearAndSort();
      await context.sync();
    }

    const filtersStartedAt = performance.now();
    timing.sortMs = filtersStartedAt - startedAt;
    timing.failedColumns = await applyFiltersBisecting(context, autoFilter, filterRange, [...filters]);
    timing.filterMs = performance.now() - filtersStartedAt;
  }
}

/**
 * Applies the filters in as few syncs as possible: a group that fails is
 * split in halves until the failing filters are found. Every sync costs a
 * round trip and a recalculation, so one sync per filter is very slow with
 * many filters. Applying a filter twice is harmless.
 *
 * @param {Array<[number, object]>} entries filter criteria by column index
 * @returns {Promise<number[]>} column indexes whose filter could not be applied
 */
async function applyFiltersBisecting(context, autoFilter, filterRange, entries) {
  if (entries.length === 0) {
    return [];
  }

  try {
    applyFilters(autoFilter, filterRange, new Map(entries));
    await context.sync();
    return [];
  } catch (error) {
    if (entries.length === 1) {
      console.error(`Filter on column index ${entries[0][0]} could not be applied again.`, entries[0][1], error, error.debugInfo);
      return [entries[0][0]];
    }
  }

  const middle = Math.ceil(entries.length / 2);
  return [
    ...(await applyFiltersBisecting(context, autoFilter, filterRange, entries.slice(0, middle))),
    ...(await applyFiltersBisecting(context, autoFilter, filterRange, entries.slice(middle))),
  ];
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
 * Updates the status after a filter change, the add-in's own refresh or a
 * configuration load, and follows the worksheet's calculations for a short
 * while after it (see followFilterStatus).
 */
function scheduleFilterStatusUpdate(worksheetId) {
  if (!statusRowsByWorksheetId.has(worksheetId)) {
    return;
  }

  getWorksheetState(worksheetId).statusWatchUntil = performance.now() + STATUS_WATCH_MS;
  queueFilterStatusUpdate(worksheetId);
}

/**
 * A calculation updates the status only shortly after a filter change: the
 * visible row count may be calculated after the filter event, and writing
 * the status calculates again (that update finds it unchanged and writes
 * nothing, so it does not loop). Calculations at other times cannot change
 * which rows a filter hides. Volatile formulas calculate every worksheet on
 * every change, so following all of them would read the filters of every
 * listed worksheet after every change. Worksheets without a filter event
 * follow every calculation, as nothing else tells that their filters changed.
 */
function followFilterStatus(worksheetId) {
  if (!statusRowsByWorksheetId.has(worksheetId)) {
    return;
  }
  const state = worksheetStates.get(worksheetId);
  if (
    statusFollowsEveryCalculation.has(worksheetId) ||
    (state && performance.now() < state.statusWatchUntil)
  ) {
    queueFilterStatusUpdate(worksheetId);
  }
}

/** Debounces status updates: several events follow one filter change. */
function queueFilterStatusUpdate(worksheetId) {
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
    queueFilterStatusUpdate(worksheetId);
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

  const startedAt = performance.now();
  let outcome = "unchanged";
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

      // The header row is only needed to name active filters.
      let text = "";
      if (!filterRange.isNullObject && (autoFilter.criteria || []).some(isActiveCriterion)) {
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
        outcome = "skipped";
        return;
      }

      context.workbook.worksheets
        .getItem(STATUS_SHEET_NAME)
        .getRange(`B${row}:C${row}`).values = [[text, visibleCount]];
      await context.sync();
      outcome = "written";
    });
  } catch (error) {
    // E.g. Excel failing on AutoFilter.criteria with an internal error. The
    // workbook then shows only the row count, so this is not reported.
    outcome = "failed";
    console.debug("Filter status was not updated.", error, error.debugInfo);
  }
  if (DIAGNOSTICS) {
    console.log(`Diagnostics: filter info update ${outcome} in ${formatMs(performance.now() - startedAt)}.`);
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

  // Click rules only run when their cell is clicked.
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
