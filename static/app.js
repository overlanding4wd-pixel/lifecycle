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
    if (page === "partners") initPartnerDashboard();
    if (page === "account-plan") initAccountPlanDetail();
    if (document.querySelector("[data-linked-partner-panel]")) initLinkedPartnerPanel();
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
    document.querySelector("[data-account-plan-form]")?.addEventListener("submit", createAccountPlan);
    document.querySelectorAll("[data-create-plan-type]").forEach((button) => {
        button.addEventListener("click", () => openCreatePlanPanel(button.dataset.createPlanType));
    });
    document.querySelectorAll("[data-cancel-create-plan]").forEach((button) => {
        button.addEventListener("click", closeCreatePlanPanel);
    });
    document.querySelector("[data-account-plan-filters]")?.addEventListener("submit", (event) => {
        event.preventDefault();
        loadAccountPlans();
    });
    document.querySelector("[data-reset-plan-filters]")?.addEventListener("click", () => {
        document.querySelector("[data-account-plan-filters]").reset();
        loadAccountPlans();
    });
    populatePlanStageFilter();
    loadDashboard();
    loadAccountPlans();
}

function openCreatePlanPanel(accountType) {
    const panel = document.querySelector("[data-create-plan-panel]");
    const form = document.querySelector("[data-account-plan-form]");
    form.reset();
    form.elements.accountType.value = accountType;
    panel.hidden = false;
    panel.classList.add("open");
    panel.setAttribute("aria-hidden", "false");
    document.querySelector("[data-create-plan-title]").textContent = accountType === "Innovator" ? "Create Innovator Plan" : "Create Direct Customer Plan";
    document.querySelector("[data-account-name-label]").textContent = accountType === "Innovator" ? "Innovator name" : "Customer name";
    form.elements.accountName.placeholder = accountType === "Innovator" ? "Innovator / partner name" : "Direct customer name";
    form.elements.accountName.focus();
}

function closeCreatePlanPanel() {
    const panel = document.querySelector("[data-create-plan-panel]");
    panel.classList.remove("open");
    panel.setAttribute("aria-hidden", "true");
    panel.hidden = true;
}

function populatePlanStageFilter() {
    const select = document.querySelector("[data-plan-stage-filter]");
    if (!select) return;
    ["I0", "I20", "I50", "D0", "D20", "Live", "Qualified Out"].forEach((stage) => {
        const option = document.createElement("option");
        option.value = stage;
        option.textContent = stage;
        select.appendChild(option);
    });
}



async function initAccountPlanDetail() {
    const page = document.querySelector("[data-page='account-plan']");
    if (!page) return;
    lifecycle.planId = page.dataset.planId;
    document.querySelector("[data-account-plan-detail-form]").addEventListener("submit", saveAccountPlanDetail);
    document.querySelector("[data-add-lifecycle-item-form]")?.addEventListener("submit", addLifecycleItem);
    await loadAccountPlanDetail();
}

async function loadAccountPlanDetail() {
    const data = await api(`/api/account-plans/${lifecycle.planId}`);
    const plan = data.plan;
    document.querySelector("[data-plan-title]").textContent = plan.accountName;
    document.querySelector("[data-plan-subtitle]").textContent = `${plan.accountType} · ${plan.templateName} · ${plan.completedItems}/${plan.totalItems} complete`;
    const form = document.querySelector("[data-account-plan-detail-form]");
    form.elements.accountName.value = plan.accountName || "";
    form.elements.accountType.value = plan.accountType || "";
    form.elements.planOwner.value = plan.planOwner || "";
    form.elements.kickOffDate.value = plan.kickOffDate || "";
    form.elements.targetGoLiveDate.value = plan.targetGoLiveDate || "";
    form.elements.territory.value = plan.territory || "";
    form.elements.currentStageOverride.value = plan.currentStageOverride || "";
    form.elements.notes.value = plan.notes || "";
    renderPlanWorkspaceSummary(plan);
    renderLifecycleItems(plan.items || []);
}

async function saveAccountPlanDetail(event) {
    event.preventDefault();
    const errors = document.querySelector("[data-plan-errors]");
    const result = document.querySelector("[data-plan-save-result]");
    errors.textContent = "";
    result.textContent = "";
    const payload = Object.fromEntries(new FormData(event.currentTarget).entries());
    try {
        await api(`/api/account-plans/${lifecycle.planId}`, { method: "PUT", body: JSON.stringify(payload) });
        result.textContent = "Saved.";
        await loadAccountPlanDetail();
    } catch (error) {
        errors.textContent = error.message;
    }
}

function renderLifecycleItems(items) {
    const target = document.querySelector("[data-lifecycle-item-table]");
    target.innerHTML = items.map((item) => `
        <tr class="${item.isOverdue ? "overdue" : ""}" data-lifecycle-item-row="${item.id}">
            <td>${item.sortOrder}</td>
            <td><input name="stage" value="${escapeHtml(item.stage)}"></td>
            <td><textarea name="activity" rows="2">${escapeHtml(item.activity)}</textarea></td>
            <td><div class="status-edit"><button class="badge ${getStatusBadgeClass(item.status)}" type="button" data-status-preview>${escapeHtml(item.status)}</button><select name="status" data-status-value="${escapeHtml(item.status)}" hidden></select></div></td>
            <td><input name="responsibleParty" value="${escapeHtml(item.responsibleParty)}"></td>
            <td>${item.startDayOffset ?? ""}</td>
            <td><input type="date" name="actualStartDate" value="${escapeHtml(item.actualStartDate)}"></td>
            <td>${item.targetDueDayOffset ?? ""}</td>
            <td><input type="date" name="dueDate" value="${escapeHtml(item.dueDate)}"></td>
            <td><input type="date" name="completedDate" value="${escapeHtml(item.completedDate)}"></td>
            <td><input name="cortaveOwner" value="${escapeHtml(item.cortaveOwner)}"></td>
            <td><input name="accountOwner" value="${escapeHtml(item.accountOwner)}"></td>
            <td><input name="nextAction" value="${escapeHtml(item.nextAction)}"></td>
            <td><input name="link" value="${escapeHtml(item.link)}"></td>
            <td><textarea name="notes" rows="2">${escapeHtml(item.notes)}</textarea></td>
            <td><button class="button secondary small" data-save-lifecycle-item="${item.id}">Save</button><button class="button danger small" data-delete-lifecycle-item="${item.id}">Delete</button></td>
        </tr>
    `).join("") || `<tr><td colspan="16">No Lifecycle Items found for this Account Plan.</td></tr>`;
    target.querySelectorAll("select[name='status']").forEach((select) => {
        (lifecycle.options.status || []).forEach((status) => {
            const option = document.createElement("option");
            option.value = status.value;
            option.textContent = status.label || status.value;
            select.appendChild(option);
        });
        select.value = select.dataset.statusValue;
        select.addEventListener("change", async () => {
            const row = select.closest("[data-lifecycle-item-row]") || select.closest("tr");
            const preview = row?.querySelector("[data-status-preview]");
            if (preview) {
                preview.className = `badge ${getStatusBadgeClass(select.value)}`;
                preview.textContent = select.value;
            }
            const button = row?.querySelector("[data-save-lifecycle-item], [data-save-plan-item]");
            if (row && button) await saveLifecycleItemRow(row, button);
            select.hidden = true;
            if (preview) preview.hidden = false;
        });
    });
    target.querySelectorAll("[data-status-preview]").forEach((badge) => {
        badge.addEventListener("click", () => {
            const wrapper = badge.closest(".status-edit");
            const select = wrapper.querySelector("select[name='status']");
            badge.hidden = true;
            select.hidden = false;
            select.focus();
        });
    });
    target.querySelectorAll("[data-save-lifecycle-item]").forEach((button) => {
        button.addEventListener("click", async () => {
            await saveLifecycleItemRow(button.closest("[data-lifecycle-item-row]"), button);
        });
    });
    target.querySelectorAll("[data-delete-lifecycle-item]").forEach((button) => {
        button.addEventListener("click", async () => {
            if (!confirm("Delete this Lifecycle Item?")) return;
            await api(`/api/lifecycle-items/${button.dataset.deleteLifecycleItem}`, { method: "DELETE" });
            await loadAccountPlanDetail();
        });
    });
}



async function saveLifecycleItemRow(row, button) {
    if (!row || !button) return;
    const itemId = button.dataset.saveLifecycleItem || button.dataset.savePlanItem;
    const payload = {};
    row.querySelectorAll("input, select, textarea").forEach((field) => {
        payload[field.name] = field.value;
    });
    button.disabled = true;
    button.textContent = "Saving...";
    try {
        const response = await api(`/api/lifecycle-items/${itemId}`, { method: "PUT", body: JSON.stringify(payload) });
        const preview = row.querySelector("[data-status-preview]");
        if (preview && response.item?.status) {
            preview.className = `badge ${getStatusBadgeClass(response.item.status)}`;
            preview.textContent = response.item.status;
        }
        button.textContent = "Saved";
        if (document.querySelector("[data-page='account-plan']")) {
            await loadAccountPlanDetail();
        } else if (lifecycle.selectedPlanId) {
            await loadPlanLifecycleItems();
        }
    } catch (error) {
        alert(error.message);
        button.textContent = "Save";
        button.disabled = false;
    }
}

function renderPlanWorkspaceSummary(plan) {
    const target = document.querySelector("[data-plan-workspace-summary]");
    if (!target) return;
    target.innerHTML = `
        <div class="metric-card"><span>Current Stage</span><strong>${escapeHtml(plan.currentStage || "-")}</strong></div>
        <div class="metric-card"><span>Plan Status</span><strong>${escapeHtml(plan.planStatus || "-")}</strong></div>
        <div class="metric-card"><span>Plan Health</span><strong>${healthBadge(plan.healthStatus || "On Track")}</strong></div>
        <div class="metric-card"><span>Completed</span><strong>${plan.completedItems}/${plan.totalItems}</strong></div>
        <div class="metric-card"><span>Open Items</span><strong>${plan.openItems ?? Math.max((plan.totalItems || 0) - (plan.completedItems || 0), 0)}</strong></div>
        <div class="metric-card"><span>Overdue Items</span><strong>${plan.overdueItems || 0}</strong></div>
        <div class="metric-card"><span>Next Step</span><strong>${escapeHtml(plan.nextStep || "-")}</strong></div>
    `;
}

async function addLifecycleItem(event) {
    event.preventDefault();
    const form = event.currentTarget;
    const payload = Object.fromEntries(new FormData(form).entries());
    await api(`/api/account-plans/${lifecycle.planId}/items`, { method: "POST", body: JSON.stringify(payload) });
    form.reset();
    await loadAccountPlanDetail();
}

async function createAccountPlan(event) {
    event.preventDefault();
    const form = event.currentTarget;
    const errors = document.querySelector("[data-account-plan-errors]");
    const result = document.querySelector("[data-account-plan-result]");
    const submitButton = form.querySelector("button[type='submit']");
    errors.textContent = "";
    if (result) result.textContent = "";
    const payload = Object.fromEntries(new FormData(form).entries());
    try {
        if (submitButton) {
            submitButton.disabled = true;
            submitButton.textContent = "Creating...";
        }
        const response = await api("/api/account-plans", { method: "POST", body: JSON.stringify(payload) });
        form.reset();
        closeCreatePlanPanel();
        await loadAccountPlans();
        await loadDashboard();
        if (result) {
            const itemCount = response.plan?.items?.length || 0;
            result.innerHTML = `<strong>Created ${escapeHtml(response.plan.accountName)}.</strong> ${itemCount} lifecycle actions were generated from the ${escapeHtml(response.plan.templateName)} template. <a class="button secondary small" href="/account-plans/${response.plan.id}">Open Plan</a>`;
        }
    } catch (error) {
        errors.textContent = error.message;
    } finally {
        if (submitButton) {
            submitButton.disabled = false;
            submitButton.textContent = "Create Plan";
        }
    }
}

async function loadAccountPlans() {
    const table = document.querySelector("[data-account-plan-table]");
    if (!table) return;
    const params = new URLSearchParams();
    const filterForm = document.querySelector("[data-account-plan-filters]");
    if (filterForm) {
        new FormData(filterForm).forEach((value, key) => {
            if (value) params.set(key, value);
        });
    }
    const data = await api(`/api/account-plans?${params.toString()}`);
    renderAccountPlanMetrics(data.summary || {});
    renderPlansNeedingAttention(data.plans || []);
    table.innerHTML = data.plans.length ? data.plans.map((plan) => `
        <tr>
            <td><a class="link-button" href="/account-plans/${plan.id}">${escapeHtml(plan.accountName)}</a></td>
            <td>${accountTypeBadge(plan.accountType)}</td>
            <td>${escapeHtml(plan.planOwner)}</td>
            <td><span class="badge stage-badge">${escapeHtml(plan.currentStage || "-")}</span></td>
            <td>${statusBadge(plan.planStatus)}</td>
            <td>${healthBadge(plan.healthStatus)}</td>
            <td>${escapeHtml(plan.nextStep || "-")}</td>
            <td>${escapeHtml(plan.nextStepOwner || "-")}</td>
            <td>${escapeHtml(plan.nextDueDate || "-")}</td>
            <td>${plan.daysOverdue || 0}</td>
            <td>${escapeHtml(plan.targetGoLiveDate || "-")}</td>
            <td>${formatDateTime(plan.updatedAt)}</td>
            <td><a class="button secondary small" href="/account-plans/${plan.id}">Open Plan</a></td>
        </tr>
    `).join("") : emptyPlansRow();
}

function emptyPlansRow() {
    return `<tr><td colspan="13"><div class="empty-dashboard-state"><p>No plans created yet. Create your first Innovator Plan or Direct Customer Plan to start tracking lifecycle progress.</p><button class="button primary" data-empty-create="Innovator">Create Innovator Plan</button><button class="button secondary" data-empty-create="Direct Customer">Create Direct Customer Plan</button></div></td></tr>`;
}

function renderPlansNeedingAttention(plans) {
    const target = document.querySelector("[data-attention-plans]");
    if (!target) return;
    const soon = new Date();
    soon.setDate(soon.getDate() + 14);
    const attention = plans.filter((plan) => {
        const targetLive = plan.targetGoLiveDate ? new Date(`${plan.targetGoLiveDate}T00:00:00`) : null;
        return plan.daysOverdue > 0 || plan.planStatus === "On Hold" || plan.healthStatus === "Blocked" || !plan.nextStepOwner || (targetLive && targetLive <= soon && !plan.isLive);
    }).slice(0, 8);
    target.innerHTML = attention.length ? attention.map((plan) => `
        <a class="attention-card" href="/account-plans/${plan.id}">
            <strong>${escapeHtml(plan.accountName)}</strong>
            <span>${accountTypeBadge(plan.accountType)} ${healthBadge(plan.healthStatus)}</span>
            <small>${escapeHtml(plan.nextStep || "No next step")} · ${escapeHtml(plan.nextStepOwner || "Missing owner")}</small>
        </a>
    `).join("") : `<p class="empty-state">No plans need attention right now.</p>`;
}

document.addEventListener("click", (event) => {
    const button = event.target.closest("[data-empty-create]");
    if (button) openCreatePlanPanel(button.dataset.emptyCreate);
});

function renderAccountPlanMetrics(summary) {
    ["totalActivePlans", "totalInnovators", "totalDirectCustomers", "nextActionsDue", "overduePlans", "live", "qualifiedOut", "onHold"].forEach((key) => {
        const element = document.querySelector(`[data-plan-metric="${key}"]`);
        if (element) element.textContent = summary[key] || 0;
    });
}

function accountTypeBadge(accountType) {
    const label = accountType === "Innovator" ? "Innovator / Partner" : accountType;
    return `<span class="badge stage-badge">${escapeHtml(label)}</span>`;
}

function healthBadge(status) {
    return `<span class="badge ${getStatusBadgeClass(status)}">${escapeHtml(status || "Not Started")}</span>`;
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
    lifecycle.selectedPlanId = new URLSearchParams(window.location.search).get("planId") || "";
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
    if (lifecycle.selectedPlanId) {
        await loadPlanLifecycleItems();
        return;
    }
    lifecycle.records = [];
    renderPlanSelectionEmptyState();
}

async function loadPlanLifecycleItems() {
    const planData = await api(`/api/account-plans/${lifecycle.selectedPlanId}`);
    const expectedType = lifecycle.trackerType === "I20" ? "Innovator" : "Direct Customer";
    if (planData.plan.accountType !== expectedType) {
        lifecycle.records = [];
        renderPlanSelectionEmptyState(`This page shows ${expectedType} plans. The selected plan is ${planData.plan.accountType}.`);
        return;
    }
    lifecycle.selectedPlan = planData.plan;
    lifecycle.records = applyLifecycleItemFilters(planData.plan.items || []);
    renderLifecycleItemTrackerTable();
}

function applyLifecycleItemFilters(items) {
    const form = document.querySelector("[data-record-filters]");
    const data = Object.fromEntries(new FormData(form).entries());
    return items.filter((item) => {
        const search = (data.search || "").toLowerCase();
        if (search && !`${item.activity} ${item.responsibleParty} ${item.notes}`.toLowerCase().includes(search)) return false;
        if (data.status && item.status !== data.status) return false;
        if (data.stage && item.stage !== data.stage) return false;
        if (data.owner && item.responsibleParty !== data.owner && item.cortaveOwner !== data.owner && item.accountOwner !== data.owner) return false;
        if (data.dueAfter && item.dueDate < data.dueAfter) return false;
        if (data.dueBefore && item.dueDate > data.dueBefore) return false;
        if (data.overdue && !item.isOverdue) return false;
        return true;
    });
}

function renderPlanSelectionEmptyState(message) {
    const target = document.querySelector("[data-record-table]");
    const empty = document.querySelector("[data-empty-records]");
    target.innerHTML = "";
    empty.hidden = false;
    empty.textContent = message || (lifecycle.trackerType === "I20"
        ? "Select or create an Innovator Partner Plan from the Master Partner Dashboard to view I20 tracker records."
        : "Select or create a Direct Customer Account Plan from the Master Partner Dashboard to view D20 tracker records.");
}


function renderLifecycleItemTrackerTable() {
    const target = document.querySelector("[data-record-table]");
    const empty = document.querySelector("[data-empty-records]");
    target.innerHTML = "";
    empty.hidden = lifecycle.records.length > 0;
    empty.textContent = lifecycle.records.length ? "" : "No lifecycle items match the current filters for this Account Plan.";
    lifecycle.records.forEach((item) => {
        const row = document.createElement("tr");
        row.className = item.isOverdue ? "overdue" : "";
        row.innerHTML = `
            <td><textarea name="activity" rows="2">${escapeHtml(item.activity)}</textarea></td>
            <td><input name="stage" value="${escapeHtml(item.stage)}"></td>
            <td><div class="status-edit"><button class="badge ${getStatusBadgeClass(item.status)}" type="button" data-status-preview>${escapeHtml(item.status)}</button><select name="status" data-status-value="${escapeHtml(item.status)}" hidden></select></div></td>
            <td><input name="responsibleParty" value="${escapeHtml(item.responsibleParty)}"></td>
            <td><input name="cortaveOwner" value="${escapeHtml(item.cortaveOwner)}"></td>
            <td><input name="accountOwner" value="${escapeHtml(item.accountOwner)}"></td>
            <td>${escapeHtml(lifecycle.selectedPlan?.accountName || "-")}</td>
            <td><input type="date" name="dueDate" value="${escapeHtml(item.dueDate)}"></td>
            <td><input name="link" value="${escapeHtml(item.link)}" placeholder="https://..."></td>
            <td>${formatDateTime(item.updatedAt)}</td>
            <td class="right"><button class="button secondary small" data-save-plan-item="${item.id}">Save</button></td>
        `;
        row.querySelectorAll("select[name='status']").forEach((select) => {
            (lifecycle.options.status || []).forEach((status) => {
                const option = document.createElement("option");
                option.value = status.value;
                option.textContent = status.label || status.value;
                select.appendChild(option);
            });
            select.value = select.dataset.statusValue;
            select.addEventListener("change", async () => {
                const preview = row.querySelector("[data-status-preview]");
                if (preview) {
                    preview.className = `badge ${getStatusBadgeClass(select.value)}`;
                    preview.textContent = select.value;
                }
                await saveLifecycleItemRow(row, row.querySelector("[data-save-plan-item]"));
            });
        });
        target.appendChild(row);
    });
    target.querySelectorAll("[data-status-preview]").forEach((badge) => {
        badge.addEventListener("click", () => {
            const wrapper = badge.closest(".status-edit");
            const select = wrapper.querySelector("select[name='status']");
            badge.hidden = true;
            select.hidden = false;
            select.focus();
        });
    });
    target.querySelectorAll("[data-save-plan-item]").forEach((button) => {
        button.addEventListener("click", async () => {
            await saveLifecycleItemRow(button.closest("tr"), button);
        });
    });
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
    loadLifecycleTemplateSettings();
    document.querySelectorAll("[data-add-option]").forEach((button) => {
        button.addEventListener("click", () => openOptionDrawer({ category: button.dataset.addOption }));
    });
    document.querySelector("[data-option-form]").addEventListener("submit", saveOption);
    document.querySelectorAll("[data-close-option-drawer]").forEach((button) => button.addEventListener("click", closeOptionDrawer));
    document.querySelector("[data-delete-option]").addEventListener("click", deleteOption);
}


async function loadLifecycleTemplateSettings() {
    const target = document.querySelector("[data-template-list]");
    if (!target) return;
    const data = await api("/api/lifecycle-templates");
    target.innerHTML = "";
    for (const template of data.templates) {
        const items = await api(`/api/lifecycle-templates/${template.id}/items`);
        const section = document.createElement("section");
        section.className = "settings-card template-card";
        section.innerHTML = `
            <div class="settings-card-header">
                <div>
                    <h3>${escapeHtml(template.templateName)}</h3>
                    <p>${escapeHtml(template.accountType)} · ${items.items.length} items</p>
                </div>
            </div>
            <div class="table-wrap">
                <table class="data-table compact template-table">
                    <thead><tr><th>#</th><th>Stage</th><th>Activity</th><th>Owner</th><th>Offsets</th><th></th></tr></thead>
                    <tbody>${items.items.map((item) => `
                        <tr>
                            <td>${item.sortOrder}</td>
                            <td>${escapeHtml(item.stage)}</td>
                            <td>${escapeHtml(item.activity)}</td>
                            <td>${escapeHtml(item.defaultResponsibleParty)}</td>
                            <td>${item.startDayOffset} / ${item.targetDueDayOffset}</td>
                            <td><button class="button ghost small" data-edit-template-item="${item.id}" data-activity="${escapeHtml(item.activity)}" data-owner="${escapeHtml(item.defaultResponsibleParty)}" data-start="${item.startDayOffset}" data-due="${item.targetDueDayOffset}">Edit</button></td>
                        </tr>
                    `).join("")}</tbody>
                </table>
            </div>
        `;
        target.appendChild(section);
    }
    target.querySelectorAll("[data-edit-template-item]").forEach((button) => {
        button.addEventListener("click", async () => {
            const activity = prompt("Activity", button.dataset.activity || "");
            if (activity === null) return;
            const owner = prompt("Default responsible party", button.dataset.owner || "");
            if (owner === null) return;
            const start = prompt("Start day offset", button.dataset.start || "0");
            if (start === null) return;
            const due = prompt("Target due day offset", button.dataset.due || "0");
            if (due === null) return;
            await api(`/api/lifecycle-template-items/${button.dataset.editTemplateItem}`, {
                method: "PUT",
                body: JSON.stringify({ activity, defaultResponsibleParty: owner, startDayOffset: start, targetDueDayOffset: due }),
            });
            await loadLifecycleTemplateSettings();
        });
    });
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
    document.querySelector("[data-partner-import-form]")?.addEventListener("submit", importPartners);
    loadPartnerImportAdminSummary();
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


const partnerState = {
    partners: [],
    sort: "partnerName",
    direction: "asc",
    options: {},
};

async function initLinkedPartnerPanel() {
    const panel = document.querySelector("[data-linked-partner-panel]");
    if (!panel) return;
    const search = panel.querySelector("[data-partner-search]");
    search.addEventListener("input", debounce(() => searchPartnersForLink(search.value), 250));
    await renderLinkedPartnerPanel();
}

async function renderLinkedPartnerPanel() {
    const panel = document.querySelector("[data-linked-partner-panel]");
    if (!panel) return;
    const data = await api("/api/lifecycle-workbook");
    const target = panel.querySelector("[data-linked-partner-summary]");
    const partner = data.workbook.partner;
    if (!partner) {
        target.innerHTML = `<p class="empty-state">No master partner linked yet. Search and select a partner to connect it.</p>`;
        return;
    }
    target.innerHTML = `
        <div class="partner-card-header">
            <div>
                <p class="eyebrow">Master Partner</p>
                <h3>${escapeHtml(partner.partnerName)}</h3>
            </div>
            <button class="button ghost small" data-unlink-partner>Unlink</button>
        </div>
        <div class="partner-detail-grid">
            ${partnerDetail("Owner", partner.owner)}
            ${partnerDetail("Territory", partner.territory)}
            ${partnerDetail("Master stage", stageBadge(partner.masterStage), true)}
            ${partnerDetail("Date of stage change", partner.dateOfStageChange)}
            ${partnerDetail("Age of stage", partner.ageOfStage)}
            ${partnerDetail("Days overdue", partner.daysOverdue ? `<strong>${partner.daysOverdue}</strong>` : "0", true)}
            ${partnerDetail("SF Account", partner.sfAccount)}
            ${partnerDetail("Partner type", partner.partnerType)}
            ${partnerDetail("CSM involved", partner.csmInvolved)}
            ${partnerDetail("Workbook link", workbookLink(partner.workbookLink, partner.isWorkbookLinkUrl), true)}
        </div>
        <div class="partner-notes"><strong>Next steps / notes</strong><p>${escapeHtml(partner.nextStepsNotes || "No notes imported.")}</p></div>
    `;
    target.querySelector("[data-unlink-partner]").addEventListener("click", async () => {
        await api("/api/lifecycle-workbook/unlink", { method: "POST", body: JSON.stringify({}) });
        await renderLinkedPartnerPanel();
    });
}

async function searchPartnersForLink(term) {
    const panel = document.querySelector("[data-linked-partner-panel]");
    const target = panel.querySelector("[data-partner-search-results]");
    const query = term.trim();
    if (query.length < 2) {
        target.innerHTML = "";
        return;
    }
    const data = await api(`/api/partners?search=${encodeURIComponent(query)}&sort=partnerName&direction=asc`);
    target.innerHTML = data.partners.slice(0, 8).map((partner) => `
        <button class="partner-result" data-link-partner="${partner.id}">
            <strong>${escapeHtml(partner.partnerName)}</strong>
            <span>${escapeHtml([partner.owner, partner.territory, partner.masterStage].filter(Boolean).join(" · "))}</span>
        </button>
    `).join("") || `<p class="empty-state">No partner matches found.</p>`;
    target.querySelectorAll("[data-link-partner]").forEach((button) => {
        button.addEventListener("click", async () => {
            await api("/api/lifecycle-workbook/link", { method: "POST", body: JSON.stringify({ partnerId: button.dataset.linkPartner }) });
            panel.querySelector("[data-partner-search]").value = "";
            target.innerHTML = "";
            await renderLinkedPartnerPanel();
        });
    });
}

async function initPartnerDashboard() {
    await loadPartnerFilterOptions();
    const form = document.querySelector("[data-partner-filters]");
    form.addEventListener("submit", (event) => {
        event.preventDefault();
        loadPartnerDashboard();
    });
    document.querySelector("[data-reset-partner-filters]").addEventListener("click", () => {
        form.reset();
        loadPartnerDashboard();
    });
    document.querySelectorAll("[data-partner-sort]").forEach((button) => {
        button.addEventListener("click", () => {
            const nextSort = button.dataset.partnerSort;
            partnerState.direction = partnerState.sort === nextSort && partnerState.direction === "asc" ? "desc" : "asc";
            partnerState.sort = nextSort;
            loadPartnerDashboard();
        });
    });
    await loadPartnerDashboard();
    await loadOrgImpact();
}


async function loadOrgImpact() {
    const panel = document.querySelector("[data-org-summary-table]");
    if (!panel) return;
    const data = await api("/api/org-impact");
    renderOrgImpactMetrics(data.totals || {});
    renderOrgImpactSummary(data.summary || []);
    renderOrgImpactParameters(data.parameters || []);
    renderOrgImpactAssumptions(data.assumptions || []);
    renderOrgImpactAverages(data.averages || []);
    const imported = document.querySelector("[data-org-impact-imported]");
    if (imported) imported.textContent = data.importedAt ? `Imported ${formatDateTime(data.importedAt)}` : "Import Partner Dashboard to populate Org Impact";
}

function renderOrgImpactMetrics(totals) {
    ["countInStage", "totalHours", "totalDays", "totalYears"].forEach((key) => {
        const element = document.querySelector(`[data-org-metric="${key}"]`);
        if (element) element.textContent = formatNumber(totals[key] || 0);
    });
}

function renderOrgImpactSummary(rows) {
    const target = document.querySelector("[data-org-summary-table]");
    target.innerHTML = rows.length ? rows.map((row) => `
        <tr>
            <td><strong>${escapeHtml(row.label)}</strong></td>
            <td>${formatNumber(row.countInStage)}</td>
            <td>${formatNumber(row.totalHours)}</td>
            <td>${formatNumber(row.totalDays)}</td>
            <td>${formatNumber(row.totalYears)}</td>
        </tr>
    `).join("") : `<tr><td colspan="5">Import Partner Dashboard to populate Org Impact.</td></tr>`;
}

function renderOrgImpactParameters(rows) {
    const target = document.querySelector("[data-org-parameters]");
    target.innerHTML = rows.length ? rows.map((row) => `
        <div class="summary-row"><span>${escapeHtml(row.label)}</span><strong>${formatNumber(row.value)}</strong></div>
    `).join("") : `<p class="empty-state">No Org Impact parameters imported yet.</p>`;
}

function renderOrgImpactAssumptions(rows) {
    const target = document.querySelector("[data-org-assumptions-table]");
    target.innerHTML = rows.length ? rows.map((row) => `
        <tr>
            <td>${escapeHtml(row.label)}</td>
            <td>${formatNumber(row.recruitmentDays)}</td>
            <td>${formatNumber(row.onboardingDays)}</td>
            <td>${formatNumber(row.activeDays)}</td>
        </tr>
    `).join("") : `<tr><td colspan="4">No assumptions imported yet.</td></tr>`;
}

function renderOrgImpactAverages(rows) {
    const target = document.querySelector("[data-org-averages-table]");
    target.innerHTML = rows.length ? rows.map((row) => `
        <tr>
            <td>${escapeHtml(row.label)}</td>
            <td>${formatNumber(row.recruitmentHours)}</td>
            <td>${formatNumber(row.onboardingHours)}</td>
            <td>${formatNumber(row.activeHours)}</td>
        </tr>
    `).join("") : `<tr><td colspan="4">No estimated averages imported yet.</td></tr>`;
}

function formatNumber(value) {
    if (value === "" || value === null || value === undefined) return "-";
    if (typeof value === "number") return new Intl.NumberFormat(undefined, { maximumFractionDigits: 2 }).format(value);
    const number = Number(value);
    if (!Number.isNaN(number) && String(value).trim() !== "") {
        return new Intl.NumberFormat(undefined, { maximumFractionDigits: 2 }).format(number);
    }
    return escapeHtml(value);
}

async function loadPartnerFilterOptions() {
    partnerState.options = await api("/api/partners/options");
    document.querySelectorAll("[data-partner-filter-options]").forEach((select) => {
        const key = select.dataset.partnerFilterOptions;
        const first = select.querySelector("option")?.cloneNode(true);
        select.innerHTML = "";
        if (first) select.appendChild(first);
        (partnerState.options[key] || []).forEach((value) => {
            const option = document.createElement("option");
            option.value = value;
            option.textContent = value;
            select.appendChild(option);
        });
    });
}

function currentPartnerParams() {
    const form = document.querySelector("[data-partner-filters]");
    const params = new URLSearchParams({ sort: partnerState.sort, direction: partnerState.direction });
    new FormData(form).forEach((value, key) => {
        if (value) params.set(key, value);
    });
    return params;
}

async function loadPartnerDashboard() {
    const data = await api(`/api/partners?${currentPartnerParams().toString()}`);
    partnerState.partners = data.partners;
    renderPartnerMetrics(data.summary || {});
    renderPartnerSummaryList("[data-partner-stage-summary]", data.summary?.byStage || {});
    renderPartnerSummaryList("[data-partner-owner-summary]", data.summary?.byOwner || {});
    renderPartnerSummaryList("[data-partner-territory-summary]", data.summary?.byTerritory || {});
    renderPartnerTable();
}

function renderPartnerMetrics(summary) {
    ["totalPartners", "overduePartners", "withLifecycleWorkbook", "withoutLifecycleWorkbook", "qualifiedOutPartners", "activePartners"].forEach((key) => {
        const element = document.querySelector(`[data-partner-metric="${key}"]`);
        if (element) element.textContent = summary[key] || 0;
    });
}

function renderPartnerSummaryList(selector, counts) {
    const target = document.querySelector(selector);
    if (!target) return;
    const entries = Object.entries(counts).sort((a, b) => b[1] - a[1]).slice(0, 8);
    target.innerHTML = entries.length ? entries.map(([label, count]) => `
        <div class="summary-row"><span>${escapeHtml(label)}</span><strong>${count}</strong></div>
    `).join("") : `<p class="empty-state">No data imported yet.</p>`;
}

function renderPartnerTable() {
    const target = document.querySelector("[data-partner-table]");
    const empty = document.querySelector("[data-empty-partners]");
    target.innerHTML = "";
    empty.hidden = partnerState.partners.length > 0;
    partnerState.partners.forEach((partner) => {
        const row = document.createElement("tr");
        if (partner.daysOverdue > 0) row.className = "overdue";
        row.innerHTML = `
            <td><strong>${escapeHtml(partner.partnerName)}</strong></td>
            <td>${escapeHtml(partner.owner || "-")}</td>
            <td>${escapeHtml(partner.territory || "-")}</td>
            <td>${stageBadge(partner.masterStage)}</td>
            <td>${escapeHtml(partner.dateOfStageChange || "-")}</td>
            <td>${escapeHtml(partner.ageOfStage || "-")}</td>
            <td>${partner.daysOverdue ? `<strong>${partner.daysOverdue}</strong>` : "0"}</td>
            <td>${escapeHtml(partner.sfAccount || "-")}</td>
            <td>${escapeHtml(partner.partnerType || "-")}</td>
            <td>${escapeHtml(partner.csmInvolved || "-")}</td>
            <td>${workbookLink(partner.workbookLink, partner.isWorkbookLinkUrl)}</td>
            <td>${partner.linkedLifecycleCount ? `<a class="button ghost small" href="/">Open lifecycle</a>` : `<span class="muted">Not linked</span>`}</td>
        `;
        target.appendChild(row);
    });
}

async function importPartners(event) {
    event.preventDefault();
    const result = document.querySelector("[data-partner-import-result]");
    result.textContent = "";
    try {
        const payload = await api("/api/partners/import", { method: "POST", body: new FormData(event.currentTarget) });
        result.innerHTML = `<strong>${payload.created} partners created, ${payload.updated} partners updated.</strong>`;
        if (payload.errors?.length) {
            const list = payload.errors.slice(0, 10).map((item) => `<li>Row ${item.row}: ${escapeHtml(item.errors.join(", "))}</li>`).join("");
            result.innerHTML += `<ul>${list}</ul>`;
        }
        await loadPartnerImportAdminSummary();
    } catch (error) {
        result.textContent = error.message;
    }
}

async function loadPartnerImportAdminSummary() {
    if (!document.querySelector("[data-partner-import-history]")) return;
    const historyData = await api("/api/partner-import-history");
    const history = document.querySelector("[data-partner-import-history]");
    history.innerHTML = historyData.history.length ? historyData.history.map((item) => `
        <div class="history-item">
            <strong>${escapeHtml(item.filename || "Partner import")}</strong>
            <span>${formatDateTime(item.importedAt)} · ${escapeHtml(item.importedBy || "System")}</span>
            <p>${item.createdCount} created · ${item.updatedCount} updated · ${item.errorCount} errors</p>
        </div>
    `).join("") : `<p class="empty-state">No partner imports yet.</p>`;
    const partners = await api("/api/partners");
    document.querySelector("[data-partners-without-link]").textContent = partners.summary?.withoutLifecycleWorkbook || 0;
    const workbook = await api("/api/lifecycle-workbook");
    document.querySelector("[data-unmatched-workbook]").textContent = workbook.workbook.partnerId ? 0 : 1;
}

function partnerDetail(label, value, isHtml = false) {
    const displayValue = isHtml ? (value || "-") : escapeHtml(value || "-");
    return `<div><span>${escapeHtml(label)}</span><strong>${displayValue}</strong></div>`;
}

function workbookLink(value, isUrl) {
    if (!value) return "-";
    return isUrl ? `<a href="${escapeHtml(value)}" target="_blank" rel="noreferrer">Open workbook</a>` : escapeHtml(value);
}

function stageBadge(stage) {
    return stage ? `<span class="badge stage-badge">${escapeHtml(stage)}</span>` : "-";
}

function debounce(fn, wait) {
    let timeout;
    return (...args) => {
        clearTimeout(timeout);
        timeout = setTimeout(() => fn(...args), wait);
    };
}

function statusBadge(status) {
    return `<span class="badge ${getStatusBadgeClass(status)}">${escapeHtml(status || "Unknown")}</span>`;
}

function getStatusBadgeClass(status) {
    const normalized = String(status || "Not Started").toLowerCase();
    if (["completed", "live", "strong"].includes(normalized)) return "badge-status-completed";
    if (["in progress", "good"].includes(normalized)) return "badge-status-good";
    if (["on hold", "at risk"].includes(normalized)) return "badge-status-in-progress";
    if (["bad"].includes(normalized)) return "badge-status-off-track";
    if (["qualified out", "n/a"].includes(normalized)) return "badge-status-not-started";
    return "badge-status-not-started";
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
