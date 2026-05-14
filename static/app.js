(() => {
  "use strict";

  const FIELD_KINDS = JSON.parse(document.getElementById("field-kinds").textContent);

  const FIELDS = [
    { value: "subject", label: "Subject" },
    { value: "body", label: "Body" },
    { value: "any_text", label: "Subject or body" },
    { value: "from_name", label: "From — name" },
    { value: "from_email", label: "From — email" },
    { value: "from_any", label: "From — name or email" },
    { value: "to_line", label: "To line" },
    { value: "recipient_email", label: "Any recipient email" },
    { value: "recipient_any", label: "Any recipient (name + email)" },
    { value: "folder", label: "Folder path" },
    { value: "is_mailing_list", label: "Looks like mailing list" },
    { value: "has_list_header", label: "Has List-* header" },
    { value: "unread", label: "Unread" },
    { value: "has_attachments", label: "Has attachments" },
    { value: "received", label: "Received date" },
  ];

  const OPERATORS_BY_KIND = {
    text: [
      { value: "contains", label: "contains" },
      { value: "not_contains", label: "does not contain" },
      { value: "equals", label: "equals" },
      { value: "starts_with", label: "starts with" },
      { value: "ends_with", label: "ends with" },
      { value: "regex", label: "matches regex" },
    ],
    bool: [
      { value: "is_true", label: "is true" },
      { value: "is_false", label: "is false" },
    ],
    date: [
      { value: "before", label: "is before" },
      { value: "after", label: "is after" },
      { value: "on", label: "is on" },
    ],
  };

  const state = {
    folders: [],
    foldersById: new Map(),
    selectedFolders: new Set(),
    targetFolder: "",
    query: makeGroup(),
    results: [],
    selectedResults: new Set(),
    savedSearches: [],
    activeSavedId: "",
  };

  // ---------- helpers ----------

  function makeGroup() {
    return { type: "group", operator: "AND", negate: false, children: [makeRule()] };
  }
  function makeRule() {
    return { type: "rule", field: "subject", operator: "contains", value: "", negate: false };
  }

  function el(tag, attrs = {}, children = []) {
    const node = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs)) {
      if (k === "class") node.className = v;
      else if (k === "dataset") Object.assign(node.dataset, v);
      else if (k.startsWith("on") && typeof v === "function") {
        node.addEventListener(k.slice(2).toLowerCase(), v);
      } else if (v === true) node.setAttribute(k, "");
      else if (v === false || v == null) {/* skip */}
      else node.setAttribute(k, v);
    }
    for (const c of [].concat(children)) {
      if (c == null) continue;
      node.appendChild(typeof c === "string" ? document.createTextNode(c) : c);
    }
    return node;
  }

  function fieldKind(field) {
    return FIELD_KINDS[field] || "text";
  }

  async function api(path, options = {}) {
    const res = await fetch(path, {
      headers: { "Content-Type": "application/json" },
      ...options,
    });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) {
      const err = new Error(data.error || `Request failed (${res.status})`);
      err.data = data;
      throw err;
    }
    return data;
  }

  // ---------- status ----------

  async function checkStatus() {
    const pill = document.getElementById("status-pill");
    try {
      const data = await api("/api/status");
      if (data.outlook_available) {
        pill.textContent = "Outlook connected";
        pill.classList.add("ok");
      } else {
        pill.textContent = "Outlook not available";
        pill.classList.add("bad");
      }
    } catch (err) {
      pill.textContent = "Status error";
      pill.classList.add("bad");
    }
  }

  // ---------- folders ----------

  async function loadFolders() {
    const list = document.getElementById("folder-list");
    list.textContent = "Loading…";
    try {
      const data = await api("/api/folders");
      state.folders = data.folders;
      state.foldersById = new Map(state.folders.map((f) => [f.id, f]));
      renderFolderList();
      renderTargetFolder();
    } catch (err) {
      list.textContent = `Failed: ${err.message}`;
    }
  }

  function renderFolderList() {
    const list = document.getElementById("folder-list");
    const filter = document.getElementById("folder-filter").value.toLowerCase();
    list.innerHTML = "";
    let lastStore = null;
    for (const folder of state.folders) {
      if (filter && !folder.path.toLowerCase().includes(filter)) continue;
      if (folder.store !== lastStore) {
        list.appendChild(el("div", { class: "folder-store" }, folder.store));
        lastStore = folder.store;
      }
      const row = el("label", {
        class: "folder-row",
        style: `padding-left:${folder.depth * 12}px`,
      }, [
        el("input", {
          type: "checkbox",
          checked: state.selectedFolders.has(folder.id),
          onchange: (e) => {
            if (e.target.checked) state.selectedFolders.add(folder.id);
            else state.selectedFolders.delete(folder.id);
          },
          "data-folder-id": folder.id,
        }),
        el("span", {}, folder.name),
        el("span", { class: "folder-count" }, String(folder.count)),
      ]);
      list.appendChild(row);
    }
    if (!list.children.length) {
      list.appendChild(el("div", { class: "muted" }, "No folders match."));
    }
  }

  function renderTargetFolder() {
    const select = document.getElementById("target-folder");
    const previous = state.targetFolder;
    select.innerHTML = "";
    select.appendChild(el("option", { value: "" }, "— pick a folder —"));
    for (const folder of state.folders) {
      select.appendChild(
        el("option", { value: folder.id }, `${folder.store}: ${folder.path}`)
      );
    }
    if (previous) select.value = previous;
  }

  // ---------- query builder ----------

  function renderQuery() {
    const root = document.getElementById("query-root");
    root.innerHTML = "";
    root.appendChild(renderGroup(state.query, null, 0));
  }

  function renderGroup(group, parent, depth) {
    const node = el("div", { class: "group" });
    const header = el("div", { class: "group-header" });
    const opSelect = el("select", {
      onchange: (e) => { group.operator = e.target.value; },
    });
    for (const op of ["AND", "OR"]) {
      const o = el("option", { value: op }, op);
      if (group.operator === op) o.selected = true;
      opSelect.appendChild(o);
    }
    header.appendChild(el("span", {}, "Match"));
    header.appendChild(opSelect);
    header.appendChild(el("span", {}, "of:"));

    const negateLabel = el("label", { class: "muted" }, [
      el("input", {
        type: "checkbox",
        checked: !!group.negate,
        onchange: (e) => { group.negate = e.target.checked; },
      }),
      " NOT",
    ]);
    header.appendChild(negateLabel);

    header.appendChild(el("button", {
      class: "btn-mini",
      type: "button",
      onclick: () => { group.children.push(makeRule()); renderQuery(); },
    }, "+ Rule"));
    header.appendChild(el("button", {
      class: "btn-mini",
      type: "button",
      onclick: () => { group.children.push(makeGroup()); renderQuery(); },
    }, "+ Group"));

    if (parent) {
      header.appendChild(el("button", {
        class: "btn-mini danger",
        type: "button",
        onclick: () => {
          parent.children = parent.children.filter((c) => c !== group);
          renderQuery();
        },
      }, "Remove group"));
    }

    node.appendChild(header);

    for (const child of group.children) {
      if (child.type === "group") node.appendChild(renderGroup(child, group, depth + 1));
      else node.appendChild(renderRule(child, group));
    }

    if (!group.children.length) {
      node.appendChild(el("div", { class: "muted" }, "(empty group matches all)"));
    }

    return node;
  }

  function renderRule(rule, parent) {
    const wrap = el("div", { class: "rule" });

    const fieldSelect = el("select", {
      onchange: (e) => {
        rule.field = e.target.value;
        const kind = fieldKind(rule.field);
        const ops = OPERATORS_BY_KIND[kind];
        if (!ops.some((o) => o.value === rule.operator)) {
          rule.operator = ops[0].value;
        }
        if (kind === "bool") rule.value = "";
        renderQuery();
      },
    });
    for (const f of FIELDS) {
      const o = el("option", { value: f.value }, f.label);
      if (rule.field === f.value) o.selected = true;
      fieldSelect.appendChild(o);
    }

    const kind = fieldKind(rule.field);
    const opSelect = el("select", {
      onchange: (e) => { rule.operator = e.target.value; },
    });
    for (const op of OPERATORS_BY_KIND[kind] || []) {
      const o = el("option", { value: op.value }, op.label);
      if (rule.operator === op.value) o.selected = true;
      opSelect.appendChild(o);
    }

    let valueInput;
    if (kind === "bool") {
      valueInput = el("span", { class: "muted" }, "");
    } else if (kind === "date") {
      valueInput = el("input", {
        type: "date",
        class: "value-input",
        value: rule.value || "",
        oninput: (e) => { rule.value = e.target.value; },
      });
    } else {
      valueInput = el("input", {
        type: "text",
        class: "value-input",
        value: rule.value || "",
        placeholder: "value",
        oninput: (e) => { rule.value = e.target.value; },
      });
    }

    const negateLabel = el("label", { class: "muted" }, [
      el("input", {
        type: "checkbox",
        checked: !!rule.negate,
        onchange: (e) => { rule.negate = e.target.checked; },
      }),
      " NOT",
    ]);

    const removeBtn = el("button", {
      class: "btn-mini danger",
      type: "button",
      onclick: () => {
        parent.children = parent.children.filter((c) => c !== rule);
        renderQuery();
      },
    }, "×");

    wrap.append(negateLabel, fieldSelect, opSelect, valueInput, removeBtn);
    return wrap;
  }

  // ---------- search ----------

  async function runSearch() {
    if (state.selectedFolders.size === 0) {
      alert("Pick at least one folder to search.");
      return;
    }
    const status = document.getElementById("results-status");
    status.textContent = "Searching…";
    document.getElementById("results-body").innerHTML = "";
    document.getElementById("results-count").textContent = "";

    const payload = {
      folders: Array.from(state.selectedFolders),
      query: state.query,
      since: document.getElementById("since-input").value || null,
      per_folder_limit: Number(document.getElementById("per-folder-limit").value) || 500,
    };

    try {
      const data = await api("/api/search", {
        method: "POST",
        body: JSON.stringify(payload),
      });
      state.results = data.results;
      state.selectedResults = new Set();
      renderResults();
      const trunc = data.truncated ? " (truncated)" : "";
      status.textContent =
        `Scanned ${data.scanned} • matched ${data.matched}${trunc}`;
      document.getElementById("results-count").textContent =
        `(${data.matched} match${data.matched === 1 ? "" : "es"})`;
    } catch (err) {
      status.textContent = `Search failed: ${err.message}`;
    }
  }

  function renderResults() {
    const tbody = document.getElementById("results-body");
    tbody.innerHTML = "";
    for (const m of state.results) {
      const flags = [];
      if (m.unread) flags.push(el("span", { class: "flag unread" }, "unread"));
      if (m.is_mailing_list) flags.push(el("span", { class: "flag list" }, "list"));
      if (m.has_attachments) flags.push(el("span", { class: "flag attach" }, "📎"));

      const tr = el("tr", {}, [
        el("td", {}, el("input", {
          type: "checkbox",
          checked: state.selectedResults.has(m.id),
          onchange: (e) => {
            if (e.target.checked) state.selectedResults.add(m.id);
            else state.selectedResults.delete(m.id);
            updateSelectAll();
          },
        })),
        el("td", {}, [
          el("div", {}, m.sender_name || m.sender_email || "(unknown)"),
          el("div", { class: "muted" }, m.sender_email || ""),
        ]),
        el("td", { class: "subject" + (m.unread ? " unread" : "") }, m.subject || "(no subject)"),
        el("td", { class: "muted" }, m.folder_path),
        el("td", { class: "muted" }, formatDate(m.received)),
        el("td", {}, flags),
        el("td", {}, el("button", {
          class: "btn-mini",
          type: "button",
          onclick: () => moveOne(m.id),
        }, "Move")),
      ]);
      tbody.appendChild(tr);
    }
    updateSelectAll();
  }

  function updateSelectAll() {
    const all = document.getElementById("select-all");
    if (!state.results.length) { all.checked = false; all.indeterminate = false; return; }
    const sel = state.selectedResults.size;
    all.checked = sel === state.results.length;
    all.indeterminate = sel > 0 && sel < state.results.length;
  }

  function formatDate(iso) {
    if (!iso) return "";
    const d = new Date(iso);
    if (Number.isNaN(d.getTime())) return iso;
    return d.toLocaleString();
  }

  // ---------- move ----------

  async function moveIds(ids) {
    const target = document.getElementById("target-folder").value;
    if (!target) { alert("Pick a target folder first."); return; }
    if (!ids.length) { alert("No messages selected."); return; }
    if (!confirm(`Move ${ids.length} message(s)?`)) return;

    try {
      const data = await api("/api/move", {
        method: "POST",
        body: JSON.stringify({ message_ids: ids, target_folder_id: target }),
      });
      const moved = new Set(data.moved);
      state.results = state.results.filter((m) => !moved.has(m.id));
      for (const id of moved) state.selectedResults.delete(id);
      renderResults();
      const status = document.getElementById("results-status");
      let msg = `Moved ${data.moved.length}`;
      if (data.errors && data.errors.length) {
        msg += ` • ${data.errors.length} failed`;
        console.warn("Move errors:", data.errors);
      }
      status.textContent = msg;
      document.getElementById("results-count").textContent =
        `(${state.results.length} remaining)`;
    } catch (err) {
      alert(`Move failed: ${err.message}`);
    }
  }

  function moveOne(id) { moveIds([id]); }
  function moveSelected() { moveIds(Array.from(state.selectedResults)); }
  function moveAll() { moveIds(state.results.map((m) => m.id)); }

  // ---------- saved searches ----------

  async function loadSavedSearches() {
    try {
      const data = await api("/api/saved-searches");
      state.savedSearches = data.items || [];
      renderSavedSearches();
    } catch (err) {
      console.error("Failed to load saved searches", err);
    }
  }

  function renderSavedSearches() {
    const select = document.getElementById("saved-search-select");
    const previous = state.activeSavedId;
    select.innerHTML = "";
    select.appendChild(el("option", { value: "" }, "— Saved searches —"));
    for (const s of state.savedSearches) {
      select.appendChild(el("option", { value: s.id }, s.name));
    }
    if (previous) select.value = previous;
    updateSavedButtons();
  }

  function updateSavedButtons() {
    const id = document.getElementById("saved-search-select").value;
    state.activeSavedId = id;
    document.getElementById("update-saved").disabled = !id;
    document.getElementById("delete-saved").disabled = !id;
  }

  function applySavedSearch(entry) {
    state.selectedFolders = new Set(entry.folders || []);
    state.query = entry.query || makeGroup();
    state.targetFolder = (entry.action && entry.action.target_folder_id) || "";
    document.getElementById("since-input").value = entry.since || "";
    document.getElementById("per-folder-limit").value = entry.per_folder_limit || 500;
    document.getElementById("target-folder").value = state.targetFolder;
    renderFolderList();
    renderQuery();
  }

  function currentPayload(name) {
    return {
      name: name,
      folders: Array.from(state.selectedFolders),
      query: state.query,
      since: document.getElementById("since-input").value || null,
      per_folder_limit: Number(document.getElementById("per-folder-limit").value) || 500,
      action: {
        type: "move",
        target_folder_id: document.getElementById("target-folder").value || null,
      },
    };
  }

  function openSaveDialog() {
    const dialog = document.getElementById("save-dialog");
    document.getElementById("save-name").value = "";
    document.getElementById("save-title").textContent = "Save search";
    dialog.returnValue = "";
    dialog.showModal();
    dialog.onclose = async () => {
      if (dialog.returnValue !== "save") return;
      const name = document.getElementById("save-name").value.trim();
      if (!name) return;
      const payload = currentPayload(name);
      try {
        const entry = await api("/api/saved-searches", {
          method: "POST",
          body: JSON.stringify(payload),
        });
        await loadSavedSearches();
        document.getElementById("saved-search-select").value = entry.id;
        updateSavedButtons();
      } catch (err) {
        alert(`Save failed: ${err.message}`);
      }
    };
  }

  async function updateActiveSaved() {
    const id = state.activeSavedId;
    if (!id) return;
    const current = state.savedSearches.find((s) => s.id === id);
    const payload = currentPayload(current ? current.name : "Saved search");
    try {
      await api(`/api/saved-searches/${id}`, {
        method: "PUT",
        body: JSON.stringify(payload),
      });
      await loadSavedSearches();
    } catch (err) {
      alert(`Update failed: ${err.message}`);
    }
  }

  async function deleteActiveSaved() {
    const id = state.activeSavedId;
    if (!id) return;
    const current = state.savedSearches.find((s) => s.id === id);
    if (!confirm(`Delete saved search "${current ? current.name : id}"?`)) return;
    try {
      await api(`/api/saved-searches/${id}`, { method: "DELETE" });
      await loadSavedSearches();
      document.getElementById("saved-search-select").value = "";
      updateSavedButtons();
    } catch (err) {
      alert(`Delete failed: ${err.message}`);
    }
  }

  function loadSelectedSaved() {
    const id = state.activeSavedId;
    if (!id) return;
    const entry = state.savedSearches.find((s) => s.id === id);
    if (entry) applySavedSearch(entry);
  }

  // ---------- wiring ----------

  function init() {
    document.getElementById("folder-filter").addEventListener("input", renderFolderList);
    document.getElementById("folders-clear").addEventListener("click", () => {
      state.selectedFolders.clear();
      renderFolderList();
    });
    document.getElementById("folders-refresh").addEventListener("click", loadFolders);
    document.getElementById("run-search").addEventListener("click", runSearch);
    document.getElementById("move-selected").addEventListener("click", moveSelected);
    document.getElementById("move-all").addEventListener("click", moveAll);
    document.getElementById("target-folder").addEventListener("change", (e) => {
      state.targetFolder = e.target.value;
    });
    document.getElementById("select-all").addEventListener("change", (e) => {
      if (e.target.checked) {
        state.selectedResults = new Set(state.results.map((m) => m.id));
      } else {
        state.selectedResults.clear();
      }
      renderResults();
    });

    document.getElementById("saved-search-select").addEventListener("change", updateSavedButtons);
    document.getElementById("load-saved").addEventListener("click", loadSelectedSaved);
    document.getElementById("save-as").addEventListener("click", openSaveDialog);
    document.getElementById("update-saved").addEventListener("click", updateActiveSaved);
    document.getElementById("delete-saved").addEventListener("click", deleteActiveSaved);

    renderQuery();
    checkStatus();
    loadFolders();
    loadSavedSearches();
  }

  document.addEventListener("DOMContentLoaded", init);
})();
