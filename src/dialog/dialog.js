/*
 * Popup of the multi-select-on-click action: one checkbox per entry, a search
 * box, a select-all checkbox and a count. It knows nothing about the workbook;
 * the task pane sends what to show and decides what to write.
 *
 * Messages, as JSON strings:
 *   popup -> task pane  { type: "ready" }           once the popup can receive
 *                       { type: "ok", selected }    ticked entries, list order
 *                       { type: "cancel" }          Abbrechen or Esc
 *   task pane -> popup  { type: "init", title, entries, selected }
 */

/* global document, Office */

const elements = {};
let entries = [];
/** @type {HTMLInputElement[]} one checkbox per entry, same order */
let checkboxes = [];

Office.onReady(() => {
  for (const id of [
    "loading",
    "picker",
    "title",
    "search",
    "select-all",
    "count",
    "entries",
    "none-hint",
    "ok",
    "cancel",
  ]) {
    elements[id] = document.getElementById(id);
  }

  elements.search.addEventListener("input", applySearch);
  elements["select-all"].addEventListener("change", toggleAll);
  elements.ok.addEventListener("click", submit);
  elements.cancel.addEventListener("click", cancel);
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") {
      event.preventDefault();
      cancel();
    }
  });

  Office.context.ui.addHandlerAsync(
    Office.EventType.DialogParentMessageReceived,
    (arg) => receive(arg.message),
    () => send({ type: "ready" })
  );
});

function send(message) {
  Office.context.ui.messageParent(JSON.stringify(message));
}

function receive(text) {
  let message;
  try {
    message = JSON.parse(text);
  } catch {
    return;
  }

  if (message && message.type === "init") {
    show(message);
  }
}

function show({ title, entries: shownEntries, selected }) {
  entries = Array.isArray(shownEntries) ? shownEntries.map(String) : [];
  const ticked = new Set(Array.isArray(selected) ? selected : []);

  elements.title.textContent = title || "";
  document.title = title || document.title;

  checkboxes = [];
  elements.entries.replaceChildren(
    ...entries.map((entry) => {
      const item = document.createElement("li");
      const label = document.createElement("label");
      const checkbox = document.createElement("input");
      const text = document.createElement("span");

      label.className = "entry";
      checkbox.type = "checkbox";
      checkbox.checked = ticked.has(entry);
      checkbox.addEventListener("change", updateSummary);
      text.textContent = entry;

      label.append(checkbox, text);
      item.append(label);
      checkboxes.push(checkbox);
      return item;
    })
  );

  elements.loading.hidden = true;
  elements.picker.hidden = false;
  updateSummary();
  elements.search.focus();
}

/** Shows the entries whose name contains the search text, case ignored. */
function applySearch() {
  const query = elements.search.value.trim().toLowerCase();

  entries.forEach((entry, index) => {
    checkboxes[index].closest("li").hidden = query !== "" && !entry.toLowerCase().includes(query);
  });
}

/** Ticks every entry, or none when all were ticked. */
function toggleAll() {
  const allTicked = tickedCount() === entries.length;
  checkboxes.forEach((checkbox) => {
    checkbox.checked = !allTicked;
  });
  updateSummary();
}

function tickedCount() {
  return checkboxes.filter((checkbox) => checkbox.checked).length;
}

function updateSummary() {
  const count = tickedCount();
  const selectAll = elements["select-all"];

  selectAll.checked = entries.length > 0 && count === entries.length;
  selectAll.indeterminate = count > 0 && count < entries.length;
  elements.count.textContent = `${count} von ${entries.length}`;
  elements["none-hint"].hidden = count > 0;
  elements.ok.disabled = count === 0;
}

function submit() {
  if (tickedCount() === 0) {
    return;
  }
  send({ type: "ok", selected: entries.filter((entry, index) => checkboxes[index].checked) });
}

function cancel() {
  send({ type: "cancel" });
}
