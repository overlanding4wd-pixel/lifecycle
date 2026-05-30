const lifecycle = {
    options: {},
    records: [],
    sort: "updatedAt",
    direction: "desc",
};

document.addEventListener("DOMContentLoaded", async () => {
    await loadOptions();
    populateOptionSelects();

    const page = document.querySelector("[data-page]")?.dataset.page;
    if (page === "dashboard") initDashboard();
    if (page === "tracker") initTracker();
    if (page === "settings") initSettings();
    if (page === "import-export") initImportExport();
});

async function api(url, options = {}) {
    const response = await fetch(url, {
        headers: options.body instanceof FormData ? {} : { "Content-Type": "application/json" },
        ...options,
    });
    const contentType = response.headers.get("content-type") || "";
    const payload = contentType.includes("application/json") ? await response.json() : await response.text();
    if (!response.ok) {
        const message = payload.errors ? payload.errors.join(" ") : payload.error || "Request failed.";
        throw new Error(message);
    }
    return payload;
}

async function loadOptions() {
    lifecycle.options = await api("/api/options");
}

function populateOptionSelects(root = document) {
    root.querySelectorAll("select[data-options]").forEach((select) => {
        const category = select.dataset.options;
        const configured = lifecycle.options[category] || [];
        const selectedValue = select.value;
        const existingEmpty = Array.from(select.options).find((option) => option.value === "");
        select.innerHTML = "";
        if (existingEmpty || category === "priority") {
            const option = document.createElement("option");
            option.value = "";
            option.textContent = existingEmpty?.textContent || (category === "priority" ? "No priority" : "All");
            select.appendChild(option);
        }
        configured.forEach((item) => {
            const option = document.createElement("option");
            option.value = item.value;
            option.textContent = item.label || item.value;
            select.appendChild(option);
        });
        if (selectedValue) select.value = selectedValue;
    });
}

function initDashboard() {
    loadDashboard();
}

async function loadDashboard() {
    const data = await api("/api/dashboard");
    setMetric("totalI20", data.totalI20);
    setMetric("totalD20", data.totalD20);
    setMetric("overdue", data.overdue.length);
    setMetric("dueThisWeek", data.dueThisWeek.length);
    setMetric("completed", data.completed);
    setMetric("onHold", data.onHold);
    setMetric("inProgress", data.inProgress);
    setMetric("notStarted", data.notStarted);
    renderStatusSummary(data.byStatus || {});
    renderStageBreakdown(data.byStage || {});
    renderDashboardTable("[data-dashboard-overdue]", data.overdue);
    renderDashboardTable("[data-dashboard-due-week]", data.dueThisWeek);
    renderRecentTable(data.recentlyUpdated || []);
}

function setMetric(name, value) {
    const element = document.querySelector(`[data-metric="${name}"]`);
    if (element) element.textContent = value;
}

function renderStatusSummary(counts) {
    const target = document.querySelector("[data-dashboard-status]");
    if (!target) return;
    target.innerHTML = "";
    (lifecycle.options.status || []).forEach((status) => {
        const row = document.createElement("a");
        row.className = "summary-row";
        row.href = `/trackers/I20?status=${encodeURIComponent(status.value)}`;
        row.innerHTML = `<span><span class="color-dot" style="background:${status.color || "#e5e7eb"}"></span>${escapeHtml(status.label)}</span><strong>${counts[status.value] || 0}</strong>`;
        target.appendChild(row);
    });
}

function renderStageBreakdown(counts) {
    const target = document.querySelector("[data-dashboard-stage]");
    if (!target) return;
    target.innerHTML = "";
    const max = Math.max(1, ...Object.values(counts));
    (lifecycle.options.stage || []).forEach((stage) => {
        const count = counts[stage.value] || 0;
        const row = document.createElement("div");
        row.className = "bar-row";
        row.innerHTML = `
            <span>${escapeHtml(stage.label)}</span>
            <div class="bar-track"><div class="bar-fill" style="width:${(count / max) * 100}%"></div></div>
            <strong>${count}</strong>
        `;
        target.appendChild(row);
    });
}

function renderDashboardTable(selector, records) {
    const target = document.querySelector(selector);
    if (!target) return;
    target.innerHTML = records.length ? "" : `<tr><td colspan="5">No records to show.</td></tr>`;
    records.slice(0, 8).forEach((record) => {
        const row = document.createElement("tr");
        row.className = record.isOverdue ? "overdue" : "";
        row.innerHTML = `
            <td>${escapeHtml(record.trackerType)}</td>
            <td><a href="/trackers/${record.trackerType}">${escapeHtml(record.title)}</a></td>
            <td>${statusBadge(record.status)}</td>
            <td>${escapeHtml(record.owner)}</td>
            <td>${escapeHtml(record.dueDate || "-")}</td>
        `;
        target.appendChild(row);
    });
}

function renderRecentTable(records) {
    const target = document.querySelector("[data-dashboard-recent]");
    if (!target) return;
    target.innerHTML = records.length ? "" : `<tr><td colspan="6">No recent activity.</td></tr>`;
    records.forEach((record) => {
        const row = document.createElement("tr");
        row.innerHTML = `
            <td>${escapeHtml(record.trackerType)}</td>
            <td><a href="/trackers/${record.trackerType}">${escapeHtml(record.title)}</a></td>
            <td>${escapeHtml(record.stage)}</td>
            <td>${statusBadge(record.status)}</td>
            <td>${escapeHtml(record.owner)}</td>
            <td>${formatDateTime(record.updatedAt)}</td>
        `;
        target.appendChild(row);
    });
}

function initTracker() {
    const page = document.querySelector("[data-page='tracker']");
    lifecycle.trackerType = page.dataset.trackerType;
    const filterForm = document.querySelector("[data-record-filters]");
    applyUrlFilters(filterForm);
    filterForm.addEventListener("submit", (event) => {
        event.preventDefault();
        loadRecords();
    });
    document.querySelector("[data-reset-filters]").addEventListener("click", () => {
        filterForm.reset();
        loadRecords();
    });
    document.querySelectorAll("[data-sort]").forEach((button) => {
        button.addEventListener("click", () => {
            const nextSort = button.dataset.sort;
            lifecycle.direction = lifecycle.sort === nextSort && lifecycle.direction === "asc" ? "desc" : "asc";
            lifecycle.sort = nextSort;
            loadRecords();
        });
    });
    document.querySelector("[data-open-new-record]")?.addEventListener("click", openNewRecord);
    document.querySelector("[data-record-form]").addEventListener("submit", saveRecord);
    document.querySelectorAll("[data-close-drawer]").forEach((button) => button.addEventListener("click", closeDrawer));
    document.querySelector("[data-delete-record]")?.addEventListener("click", deleteCurrentRecord);
    document.querySelectorAll("[data-export]").forEach((button) => {
        button.addEventListener("click", () => downloadCurrentExport(button.dataset.export));
    });
    loadRecords();
}

function applyUrlFilters(form) {
    const params = new URLSearchParams(window.location.search);
    params.forEach((value, key) => {
        const input = form.elements[key];
        if (!input) return;
        if (input.type === "checkbox") {
            input.checked = value === "true";
        } else {
            input.value = value;
        }
    });
}

function currentFilterParams() {
    const form = document.querySelector("[data-record-filters]");
    const params = new URLSearchParams({ trackerType: lifecycle.trackerType, sort: lifecycle.sort, direction: lifecycle.direction });
    new FormData(form).forEach((value, key) => {
        if (value) params.set(key, value);
    });
    return params;
}

async function loadRecords() {
    const params = currentFilterParams();
    const data = await api(`/api/records?${params.toString()}`);
    lifecycle.records = data.records;
    renderRecordTable();
}

function renderRecordTable() {
    const target = document.querySelector("[data-record-table]");
    const empty = document.querySelector("[data-empty-records]");
    target.innerHTML = "";
    empty.hidden = lifecycle.records.length > 0;
    lifecycle.records.forEach((record) => {
        const row = document.createElement("tr");
        row.className = record.isOverdue ? "overdue" : "";
        row.innerHTML = `
            <td><button class="link-button" data-open-record="${record.id}">${escapeHtml(record.activity || record.title)}</button></td>
            <td>${escapeHtml(record.stage)}</td>
            <td>${statusBadge(record.status)}</td>
            <td>${escapeHtml(record.owner)}</td>
            <td>${escapeHtml(record.whoAtCortave || "-")}</td>
            <td>${escapeHtml(record.cortaveOwner || "-")}</td>
            <td>${escapeHtml(record.innovatorOwner || "-")}</td>
            <td>${escapeHtml(record.customer || "-")}</td>
            <td>${record.isOverdue ? "<strong>Overdue</strong><br>" : ""}${escapeHtml(record.dueDate || "-")}</td>
            <td>${record.link ? `<a href="${escapeHtml(record.link)}" target="_blank" rel="noreferrer">Open</a>` : "-"}</td>
            <td>${formatDateTime(record.updatedAt)}</td>
            <td class="right"><button class="button ghost small" data-open-record="${record.id}">Open</button></td>
        `;
        target.appendChild(row);
    });
    target.querySelectorAll("[data-open-record]").forEach((button) => {
        button.addEventListener("click", () => openExistingRecord(button.dataset.openRecord));
    });
}

function openNewRecord() {
    const form = document.querySelector("[data-record-form]");
    form.reset();
    form.elements.id.value = "";
    form.elements.trackerType.value = lifecycle.trackerType;
    form.elements.stage.value = lifecycle.trackerType;
    form.elements.status.value = "Not Started";
    form.elements.owner.value = (lifecycle.options.owner || [])[0]?.value || "";
    form.elements.priority.value = "Medium";
    form.elements.additionalFields.value = "{}";
    document.querySelector("[data-drawer-title]").textContent = "Add record";
    document.querySelector("[data-record-history]").innerHTML = "<p>No changes yet.</p>";
    document.querySelector("[data-delete-record]")?.setAttribute("hidden", "hidden");
    openDrawer();
}

async function openExistingRecord(id) {
    const data = await api(`/api/records/${id}`);
    const record = data.record;
    const form = document.querySelector("[data-record-form]");
    form.reset();
    Object.entries(record).forEach(([key, value]) => {
        const field = form.elements[key];
        if (!field) return;
        field.value = key === "additionalFields" ? JSON.stringify(value || {}, null, 2) : value || "";
    });
    document.querySelector("[data-drawer-title]").textContent = record.title;
    renderHistory(data.history || []);
    document.querySelector("[data-delete-record]")?.removeAttribute("hidden");
    openDrawer();
}

function openDrawer() {
    document.querySelector("[data-form-errors]").textContent = "";
    const drawer = document.querySelector("[data-record-drawer]");
    drawer.classList.add("open");
    drawer.setAttribute("aria-hidden", "false");
}

function closeDrawer() {
    const drawer = document.querySelector("[data-record-drawer]");
    drawer.classList.remove("open");
    drawer.setAttribute("aria-hidden", "true");
}

async function saveRecord(event) {
    event.preventDefault();
    const form = event.currentTarget;
    const errors = document.querySelector("[data-form-errors]");
    errors.textContent = "";
    let additionalFields = {};
    try {
        additionalFields = form.elements.additionalFields.value ? JSON.parse(form.elements.additionalFields.value) : {};
    } catch (error) {
        errors.textContent = "Additional imported fields must be valid JSON.";
        return;
    }
    const payload = Object.fromEntries(new FormData(form).entries());
    payload.additionalFields = additionalFields;
    const id = payload.id;
    delete payload.id;
    try {
        if (id) {
            await api(`/api/records/${id}`, { method: "PUT", body: JSON.stringify(payload) });
        } else {
            await api("/api/records", { method: "POST", body: JSON.stringify(payload) });
        }
        closeDrawer();
        loadRecords();
    } catch (error) {
        errors.textContent = error.message;
    }
}

async function deleteCurrentRecord() {
    const form = document.querySelector("[data-record-form]");
    const id = form.elements.id.value;
    if (!id || !confirm("Delete this tracker record?")) return;
    await api(`/api/records/${id}`, { method: "DELETE" });
    closeDrawer();
    loadRecords();
}

function renderHistory(history) {
    const target = document.querySelector("[data-record-history]");
    target.innerHTML = history.length ? "" : "<p>No change history recorded.</p>";
    history.forEach((item) => {
        let summary = {};
        try {
            summary = JSON.parse(item.change_summary || "{}");
        } catch (error) {
            summary = {};
        }
        const row = document.createElement("div");
        row.className = "history-item";
        row.innerHTML = `
            <strong>${escapeHtml(item.changed_by || "System")}</strong>
            <span>${formatDateTime(item.changed_at)}</span>
            <p>Status: ${escapeHtml(item.previous_status || "-")} → ${escapeHtml(item.new_status || "-")}</p>
            <p>Stage: ${escapeHtml(item.previous_stage || "-")} → ${escapeHtml(item.new_stage || "-")}</p>
            <small>${escapeHtml(Object.keys(summary).join(", ") || "Updated")}</small>
        `;
        target.appendChild(row);
    });
}

function downloadCurrentExport(format) {
    const params = currentFilterParams();
    params.set("format", format);
    window.location.href = `/api/export?${params.toString()}`;
}

function initSettings() {
    renderSettings();
    document.querySelectorAll("[data-add-option]").forEach((button) => {
        button.addEventListener("click", () => openOptionDrawer({ category: button.dataset.addOption }));
    });
    document.querySelector("[data-option-form]").addEventListener("submit", saveOption);
    document.querySelectorAll("[data-close-option-drawer]").forEach((button) => button.addEventListener("click", closeOptionDrawer));
    document.querySelector("[data-delete-option]").addEventListener("click", deleteOption);
}

function renderSettings() {
    ["status", "stage", "owner", "priority"].forEach((category) => {
        const target = document.querySelector(`[data-option-list="${category}"]`);
        if (!target) return;
        target.innerHTML = "";
        (lifecycle.options[category] || []).forEach((option) => {
            const row = document.createElement("div");
            row.className = "option-row";
            row.innerHTML = `
                <span>${option.color ? `<span class="color-dot" style="background:${option.color}"></span>` : ""}${escapeHtml(option.label)}</span>
                <button class="button ghost small" data-edit-option="${option.id}">Edit</button>
            `;
            row.querySelector("button").addEventListener("click", () => openOptionDrawer(option));
            target.appendChild(row);
        });
    });
}

function openOptionDrawer(option) {
    const form = document.querySelector("[data-option-form]");
    form.reset();
    form.elements.id.value = option.id || "";
    form.elements.category.value = option.category;
    form.elements.value.value = option.value || "";
    form.elements.label.value = option.label || "";
    form.elements.color.value = option.color || "#1f4e79";
    form.elements.sortOrder.value = option.sort_order || 100;
    document.querySelector("[data-option-title]").textContent = `${option.id ? "Edit" : "Add"} ${option.category} option`;
    document.querySelector("[data-option-errors]").textContent = "";
    if (option.id) {
        document.querySelector("[data-delete-option]").removeAttribute("hidden");
    } else {
        document.querySelector("[data-delete-option]").setAttribute("hidden", "hidden");
    }
    const drawer = document.querySelector("[data-option-drawer]");
    drawer.classList.add("open");
    drawer.setAttribute("aria-hidden", "false");
}

function closeOptionDrawer() {
    const drawer = document.querySelector("[data-option-drawer]");
    drawer.classList.remove("open");
    drawer.setAttribute("aria-hidden", "true");
}

async function saveOption(event) {
    event.preventDefault();
    const form = event.currentTarget;
    const errors = document.querySelector("[data-option-errors]");
    errors.textContent = "";
    const payload = Object.fromEntries(new FormData(form).entries());
    const id = payload.id;
    delete payload.id;
    try {
        if (id) {
            await api(`/api/options/${id}`, { method: "PUT", body: JSON.stringify(payload) });
        } else {
            await api("/api/options", { method: "POST", body: JSON.stringify(payload) });
        }
        await loadOptions();
        populateOptionSelects();
        renderSettings();
        closeOptionDrawer();
    } catch (error) {
        errors.textContent = error.message;
    }
}

async function deleteOption() {
    const id = document.querySelector("[data-option-form]").elements.id.value;
    if (!id || !confirm("Remove this option from active dropdowns?")) return;
    await api(`/api/options/${id}`, { method: "DELETE" });
    await loadOptions();
    renderSettings();
    closeOptionDrawer();
}

function initImportExport() {
    document.querySelector("[data-import-form]").addEventListener("submit", importRecords);
    document.querySelectorAll("[data-download-export]").forEach((button) => {
        button.addEventListener("click", () => {
            const form = document.querySelector("[data-export-form]");
            const params = new URLSearchParams({ format: button.dataset.downloadExport });
            new FormData(form).forEach((value, key) => {
                if (value) params.set(key, value);
            });
            window.location.href = `/api/export?${params.toString()}`;
        });
    });
}

async function importRecords(event) {
    event.preventDefault();
    const result = document.querySelector("[data-import-result]");
    result.textContent = "";
    try {
        const payload = await api("/api/import", { method: "POST", body: new FormData(event.currentTarget) });
        result.innerHTML = `<strong>${payload.imported} records imported.</strong>`;
        if (payload.errors?.length) {
            const list = payload.errors.map((item) => `<li>Row ${item.row}: ${escapeHtml(item.errors.join(", "))}</li>`).join("");
            result.innerHTML += `<ul>${list}</ul>`;
        }
    } catch (error) {
        result.textContent = error.message;
    }
}

function statusBadge(status) {
    const option = (lifecycle.options.status || []).find((item) => item.value === status);
    const background = option?.color || "";
    const style = background ? ` style="background:${escapeHtml(background)}; color:${readableTextColor(background)}"` : "";
    return `<span class="badge badge-status-${slug(status)}"${style}>${escapeHtml(status || "Unknown")}</span>`;
}

function readableTextColor(hex) {
    const value = String(hex || "").replace("#", "");
    if (!/^[0-9a-fA-F]{6}$/.test(value)) return "#111827";
    const r = parseInt(value.slice(0, 2), 16);
    const g = parseInt(value.slice(2, 4), 16);
    const b = parseInt(value.slice(4, 6), 16);
    const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
    return luminance > 0.62 ? "#111827" : "#ffffff";
}

function slug(value) {
    return String(value || "").toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/(^-|-$)/g, "");
}

function formatDateTime(value) {
    if (!value) return "-";
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return value;
    return date.toLocaleDateString(undefined, { year: "numeric", month: "short", day: "numeric" });
}

function escapeHtml(value) {
    return String(value ?? "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
}
