const STORAGE_KEY = "gyosei-app-data-v2";
const LEGACY_CASES_KEY = "gyosei-cases";
const STATUS_ORDER = ["未着手", "進行中", "完了"];

const state = loadState();

const tabs = Array.from(document.querySelectorAll(".tab-btn"));
const panels = {
  cases: document.getElementById("tab-cases"),
  sales: document.getElementById("tab-sales"),
  expenses: document.getElementById("tab-expenses"),
};

const summaryGrid = document.getElementById("summary-grid");

const caseForm = document.getElementById("case-form");
const caseList = document.getElementById("case-list");
const caseEmpty = document.getElementById("case-empty");
const clearBtn = document.getElementById("clear-btn");

const saleForm = document.getElementById("sale-form");
const saleCaseSelect = document.getElementById("sale-case-id");
const salesList = document.getElementById("sales-list");
const salesEmpty = document.getElementById("sales-empty");

const expenseForm = document.getElementById("expense-form");
const expenseCaseSelect = document.getElementById("expense-case-id");
const expensesList = document.getElementById("expenses-list");
const expensesEmpty = document.getElementById("expenses-empty");

const caseItemTemplate = document.getElementById("case-item-template");
const saleItemTemplate = document.getElementById("sale-item-template");
const expenseItemTemplate = document.getElementById("expense-item-template");

bindEvents();
renderAll();

function bindEvents() {
  tabs.forEach((btn) => {
    btn.addEventListener("click", () => activateTab(btn.dataset.tab));
  });

  caseForm.addEventListener("submit", (event) => {
    event.preventDefault();
    const customerName = caseForm.elements.customerName.value.trim();
    const caseName = caseForm.elements.caseName.value.trim();
    if (!customerName || !caseName) return;

    state.cases.push({
      id: crypto.randomUUID(),
      customerName,
      caseName,
      estimateAmount: normalizeAmount(caseForm.elements.amount.value),
      receivedDate: caseForm.elements.receivedDate.value || "",
      dueDate: caseForm.elements.dueDate.value || "",
      status: normalizeStatus(caseForm.elements.status.value),
      createdAt: Date.now(),
      updatedAt: Date.now(),
    });

    persist();
    caseForm.reset();
    caseForm.elements.status.value = "未着手";
    renderAll();
  });

  clearBtn.addEventListener("click", () => {
    state.cases = [];
    state.sales = [];
    state.expenses = [];
    persist();
    renderAll();
  });

  caseList.addEventListener("click", (event) => {
    const btn = event.target;
    if (!(btn instanceof HTMLButtonElement) || !btn.classList.contains("delete-btn")) return;
    const item = btn.closest(".item");
    if (!item) return;

    const id = item.dataset.id;
    state.cases = state.cases.filter((c) => c.id !== id);
    state.sales = state.sales.filter((s) => s.caseId !== id);
    state.expenses = state.expenses.filter((e) => e.caseId !== id);
    persist();
    renderAll();
  });

  saleForm.addEventListener("submit", (event) => {
    event.preventDefault();
    if (!state.cases.length) return;

    const invoiceAmount = normalizeAmount(document.getElementById("invoice-amount").value);
    if (invoiceAmount === null) return;

    const isUnpaidChecked = document.getElementById("is-unpaid").checked;
    const paidAmount = normalizeAmount(document.getElementById("paid-amount").value);

    state.sales.push({
      id: crypto.randomUUID(),
      caseId: saleCaseSelect.value,
      invoiceAmount,
      paidAmount,
      paidDate: document.getElementById("paid-date").value || "",
      isUnpaid: isUnpaidChecked || (paidAmount ?? 0) < invoiceAmount,
      createdAt: Date.now(),
    });

    persist();
    saleForm.reset();
    renderAll();
  });

  salesList.addEventListener("click", (event) => {
    const btn = event.target;
    if (!(btn instanceof HTMLButtonElement) || !btn.classList.contains("delete-btn")) return;
    const item = btn.closest(".item");
    if (!item) return;

    const id = item.dataset.id;
    state.sales = state.sales.filter((s) => s.id !== id);
    persist();
    renderAll();
  });

  expenseForm.addEventListener("submit", (event) => {
    event.preventDefault();

    const date = document.getElementById("expense-date").value;
    const content = document.getElementById("expense-content").value.trim();
    const amount = normalizeAmount(document.getElementById("expense-amount").value);

    if (!date || !content || amount === null) return;

    state.expenses.push({
      id: crypto.randomUUID(),
      date,
      content,
      amount,
      caseId: expenseCaseSelect.value || null,
      createdAt: Date.now(),
    });

    persist();
    expenseForm.reset();
    renderAll();
  });

  expensesList.addEventListener("click", (event) => {
    const btn = event.target;
    if (!(btn instanceof HTMLButtonElement) || !btn.classList.contains("delete-btn")) return;
    const item = btn.closest(".item");
    if (!item) return;

    const id = item.dataset.id;
    state.expenses = state.expenses.filter((e) => e.id !== id);
    persist();
    renderAll();
  });
}

function renderAll() {
  renderDashboard();
  renderCaseOptions();
  renderCases();
  renderSales();
  renderExpenses();
}

function activateTab(tabKey) {
  tabs.forEach((btn) => btn.classList.toggle("active", btn.dataset.tab === tabKey));
  Object.entries(panels).forEach(([key, panel]) => {
    panel.classList.toggle("active", key === tabKey);
  });
}

function renderDashboard() {
  const salesByMonth = aggregateByMonth(state.sales, (s) => s.createdAt, (s) => s.invoiceAmount);
  const expenseByMonth = aggregateByMonth(state.expenses, (e) => e.date, (e) => e.amount);
  const monthKeys = [...new Set([...Object.keys(salesByMonth), ...Object.keys(expenseByMonth)])].sort().reverse();

  const currentMonth = monthKeys[0] || toMonthKey(new Date());
  const sales = salesByMonth[currentMonth] || 0;
  const expenses = expenseByMonth[currentMonth] || 0;
  const profit = sales - expenses;

  summaryGrid.innerHTML = "";
  [
    { label: `対象月`, value: monthLabel(currentMonth), cls: "" },
    { label: "月別売上合計", value: formatCurrency(sales), cls: "" },
    { label: "月別経費合計", value: formatCurrency(expenses), cls: "" },
    { label: "利益（売上−経費）", value: formatCurrency(profit), cls: "profit" },
  ].forEach((card) => {
    const div = document.createElement("div");
    div.className = `summary-card ${card.cls}`.trim();
    div.innerHTML = `<p class="label">${card.label}</p><p class="value">${card.value}</p>`;
    summaryGrid.appendChild(div);
  });
}

function renderCaseOptions() {
  const options = state.cases
    .slice()
    .sort((a, b) => a.createdAt - b.createdAt)
    .map((c) => `<option value="${c.id}">${escapeHtml(c.customerName)}｜${escapeHtml(c.caseName)}</option>`)
    .join("");

  saleCaseSelect.innerHTML = state.cases.length ? options : '<option value="">案件を先に登録してください</option>';
  saleCaseSelect.disabled = !state.cases.length;

  expenseCaseSelect.innerHTML = `<option value="">案件に紐付けしない</option>${options}`;
}

function renderCases() {
  caseList.innerHTML = "";
  const sorted = state.cases.slice().sort(sortCases);

  sorted.forEach((entry) => {
    const node = caseItemTemplate.content.cloneNode(true);
    const item = node.querySelector(".item");
    const title = node.querySelector(".title");
    const meta = node.querySelector(".meta");

    item.dataset.id = entry.id;
    title.textContent = `${entry.customerName}｜${entry.caseName}`;
    meta.textContent = `見積: ${formatCurrency(entry.estimateAmount)} / 受付日: ${formatDate(entry.receivedDate)} / 期限日: ${formatDate(entry.dueDate)} / ${entry.status}`;
    caseList.appendChild(node);
  });

  caseEmpty.hidden = sorted.length > 0;
}

function renderSales() {
  salesList.innerHTML = "";
  const sorted = state.sales.slice().sort((a, b) => b.createdAt - a.createdAt);

  sorted.forEach((sale) => {
    const node = saleItemTemplate.content.cloneNode(true);
    const item = node.querySelector(".item");
    const title = node.querySelector(".title");
    const meta = node.querySelector(".meta");

    const caseName = resolveCaseName(sale.caseId);
    item.dataset.id = sale.id;
    title.textContent = `${caseName}｜請求: ${formatCurrency(sale.invoiceAmount)}`;
    meta.textContent = `入金: ${formatCurrency(sale.paidAmount)} / 入金日: ${formatDate(sale.paidDate)} / 未入金: ${sale.isUnpaid ? "はい" : "いいえ"}`;

    salesList.appendChild(node);
  });

  salesEmpty.hidden = sorted.length > 0;
}

function renderExpenses() {
  expensesList.innerHTML = "";
  const sorted = state.expenses
    .slice()
    .sort((a, b) => toSortTimestamp(b.date) - toSortTimestamp(a.date));

  sorted.forEach((expense) => {
    const node = expenseItemTemplate.content.cloneNode(true);
    const item = node.querySelector(".item");
    const title = node.querySelector(".title");
    const meta = node.querySelector(".meta");

    item.dataset.id = expense.id;
    title.textContent = `${formatDate(expense.date)}｜${expense.content}`;
    meta.textContent = `金額: ${formatCurrency(expense.amount)} / 紐付け: ${expense.caseId ? resolveCaseName(expense.caseId) : "なし"}`;

    expensesList.appendChild(node);
  });

  expensesEmpty.hidden = sorted.length > 0;
}

function resolveCaseName(caseId) {
  const found = state.cases.find((c) => c.id === caseId);
  if (!found) return "（削除済み案件）";
  return `${found.customerName}｜${found.caseName}`;
}

function sortCases(a, b) {
  const dueA = toSortTimestamp(a.dueDate);
  const dueB = toSortTimestamp(b.dueDate);
  if (dueA !== dueB) return dueA - dueB;

  const statusA = STATUS_ORDER.indexOf(normalizeStatus(a.status));
  const statusB = STATUS_ORDER.indexOf(normalizeStatus(b.status));
  if (statusA !== statusB) return statusA - statusB;
  return a.createdAt - b.createdAt;
}

function toSortTimestamp(value) {
  if (typeof value === "number") return value;
  if (!value) return Number.MAX_SAFE_INTEGER;
  const t = Date.parse(value);
  return Number.isNaN(t) ? Number.MAX_SAFE_INTEGER : t;
}

function aggregateByMonth(list, dateSelector, amountSelector) {
  return list.reduce((acc, item) => {
    const key = toMonthKey(dateSelector(item));
    if (!key) return acc;
    acc[key] = (acc[key] || 0) + (amountSelector(item) || 0);
    return acc;
  }, {});
}

function toMonthKey(source) {
  const date = source instanceof Date ? source : new Date(source);
  if (Number.isNaN(date.getTime())) return "";
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
}

function monthLabel(key) {
  const [y, m] = key.split("-");
  if (!y || !m) return "-";
  return `${y}年${Number(m)}月`;
}

function formatCurrency(value) {
  if (typeof value !== "number" || Number.isNaN(value)) return "未入力";
  return `${new Intl.NumberFormat("ja-JP").format(value)} 円`;
}

function formatDate(dateText) {
  if (!dateText) return "未設定";
  const date = new Date(dateText);
  if (Number.isNaN(date.getTime())) return "未設定";
  return new Intl.DateTimeFormat("ja-JP").format(date);
}

function normalizeAmount(raw) {
  if (raw === "" || raw === null || raw === undefined) return null;
  const value = Number(raw);
  if (!Number.isFinite(value) || value < 0) return null;
  return Math.floor(value);
}

function normalizeStatus(status) {
  return STATUS_ORDER.includes(status) ? status : "未着手";
}

function escapeHtml(text) {
  return String(text)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function persist() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) {
      const parsed = JSON.parse(raw);
      return sanitizeState(parsed);
    }

    const legacyRaw = localStorage.getItem(LEGACY_CASES_KEY);
    const legacyCases = legacyRaw ? JSON.parse(legacyRaw) : [];
    return sanitizeState({ cases: legacyCases, sales: [], expenses: [] });
  } catch {
    return { cases: [], sales: [], expenses: [] };
  }
}

function sanitizeState(input) {
  const safe = input && typeof input === "object" ? input : {};
  const cases = Array.isArray(safe.cases) ? safe.cases : [];
  const sales = Array.isArray(safe.sales) ? safe.sales : [];
  const expenses = Array.isArray(safe.expenses) ? safe.expenses : [];

  return {
    cases: cases
      .filter((c) => c && typeof c === "object")
      .map((c) => ({
        id: c.id || crypto.randomUUID(),
        customerName: String(c.customerName || "").trim(),
        caseName: String(c.caseName || "").trim(),
        estimateAmount: normalizeAmount(c.estimateAmount ?? c.amount),
        receivedDate: String(c.receivedDate || ""),
        dueDate: String(c.dueDate || ""),
        status: normalizeStatus(c.status),
        createdAt: Number.isFinite(c.createdAt) ? c.createdAt : Date.now(),
        updatedAt: Number.isFinite(c.updatedAt) ? c.updatedAt : Date.now(),
      }))
      .filter((c) => c.customerName && c.caseName),
    sales: sales
      .filter((s) => s && typeof s === "object")
      .map((s) => {
        const invoiceAmount = normalizeAmount(s.invoiceAmount);
        const paidAmount = normalizeAmount(s.paidAmount);
        return {
          id: s.id || crypto.randomUUID(),
          caseId: String(s.caseId || ""),
          invoiceAmount: invoiceAmount ?? 0,
          paidAmount,
          paidDate: String(s.paidDate || ""),
          isUnpaid: Boolean(s.isUnpaid) || (paidAmount ?? 0) < (invoiceAmount ?? 0),
          createdAt: Number.isFinite(s.createdAt) ? s.createdAt : Date.now(),
        };
      })
      .filter((s) => s.caseId),
    expenses: expenses
      .filter((e) => e && typeof e === "object")
      .map((e) => ({
        id: e.id || crypto.randomUUID(),
        date: String(e.date || ""),
        content: String(e.content || "").trim(),
        amount: normalizeAmount(e.amount) ?? 0,
        caseId: e.caseId ? String(e.caseId) : null,
        createdAt: Number.isFinite(e.createdAt) ? e.createdAt : Date.now(),
      }))
      .filter((e) => e.date && e.content),
  };
}
