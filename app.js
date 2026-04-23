const STORAGE_KEY = "gyosei-app-data-v3";
const LEGACY_STORAGE_KEYS = ["gyosei-app-data-v2", "gyosei-cases"];
const STATUS_ORDER = ["未着手", "進行中", "完了"];

const state = {
  cases: [],
  sales: [],
  expenses: [],
};

const editState = {
  caseId: null,
  saleId: null,
  expenseId: null,
};

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
const caseSubmitBtn = caseForm.querySelector('button[type="submit"]');

const saleForm = document.getElementById("sale-form");
const saleCaseSelect = document.getElementById("sale-case-id");
const salesList = document.getElementById("sales-list");
const salesEmpty = document.getElementById("sales-empty");
const saleSubmitBtn = saleForm.querySelector('button[type="submit"]');

const expenseForm = document.getElementById("expense-form");
const expenseCaseSelect = document.getElementById("expense-case-id");
const expensesList = document.getElementById("expenses-list");
const expensesEmpty = document.getElementById("expenses-empty");
const expenseSubmitBtn = expenseForm.querySelector('button[type="submit"]');

const caseItemTemplate = document.getElementById("case-item-template");
const saleItemTemplate = document.getElementById("sale-item-template");
const expenseItemTemplate = document.getElementById("expense-item-template");

initialize();

function initialize() {
  const loaded = loadData();
  state.cases = loaded.cases;
  state.sales = loaded.sales;
  state.expenses = loaded.expenses;

  bindEvents();
  renderAll();
}

function bindEvents() {
  tabs.forEach((btn) => {
    btn.addEventListener("click", () => activateTab(btn.dataset.tab));
  });

  caseForm.addEventListener("submit", handleCaseSubmit);
  saleForm.addEventListener("submit", handleSaleSubmit);
  expenseForm.addEventListener("submit", handleExpenseSubmit);

  clearBtn.addEventListener("click", () => {
    if (!window.confirm("案件・売上・経費の全データを削除します。よろしいですか？")) return;
    state.cases = [];
    state.sales = [];
    state.expenses = [];
    resetEditMode("case");
    resetEditMode("sale");
    resetEditMode("expense");
    saveData();
    renderAll();
  });

  caseList.addEventListener("click", handleCaseListAction);
  salesList.addEventListener("click", handleSalesListAction);
  expensesList.addEventListener("click", handleExpensesListAction);
}

function handleCaseSubmit(event) {
  event.preventDefault();

  const customerName = caseForm.elements.customerName.value.trim();
  const caseName = caseForm.elements.caseName.value.trim();
  if (!customerName || !caseName) return;

  const payload = {
    customerName,
    caseName,
    estimateAmount: normalizeAmount(caseForm.elements.amount.value),
    receivedDate: caseForm.elements.receivedDate.value || "",
    dueDate: caseForm.elements.dueDate.value || "",
    status: normalizeStatus(caseForm.elements.status.value),
    updatedAt: Date.now(),
  };

  if (editState.caseId) {
    const index = state.cases.findIndex((item) => item.id === editState.caseId);
    if (index !== -1) {
      state.cases[index] = { ...state.cases[index], ...payload };
    }
  } else {
    state.cases.push({
      id: crypto.randomUUID(),
      ...payload,
      createdAt: Date.now(),
    });
  }

  resetCaseForm();
  saveData();
  renderAll();
}

function handleSaleSubmit(event) {
  event.preventDefault();
  if (!state.cases.length) return;

  const invoiceAmount = normalizeAmount(saleForm.elements.invoiceAmount.value);
  if (invoiceAmount === null) return;

  const paidAmount = normalizeAmount(saleForm.elements.paidAmount.value);

  const payload = {
    caseId: saleCaseSelect.value,
    invoiceAmount,
    paidAmount,
    paidDate: saleForm.elements.paidDate.value || "",
    isUnpaid: saleForm.elements.isUnpaid.checked || (paidAmount ?? 0) < invoiceAmount,
  };

  if (!payload.caseId) return;

  if (editState.saleId) {
    const index = state.sales.findIndex((item) => item.id === editState.saleId);
    if (index !== -1) {
      state.sales[index] = {
        ...state.sales[index],
        ...payload,
        updatedAt: Date.now(),
      };
    }
  } else {
    state.sales.push({
      id: crypto.randomUUID(),
      ...payload,
      createdAt: Date.now(),
    });
  }

  resetSaleForm();
  saveData();
  renderAll();
}

function handleExpenseSubmit(event) {
  event.preventDefault();

  const date = expenseForm.elements.expenseDate.value;
  const content = expenseForm.elements.expenseContent.value.trim();
  const amount = normalizeAmount(expenseForm.elements.expenseAmount.value);

  if (!date || !content || amount === null) return;

  const payload = {
    date,
    content,
    amount,
    caseId: expenseCaseSelect.value || null,
  };

  if (editState.expenseId) {
    const index = state.expenses.findIndex((item) => item.id === editState.expenseId);
    if (index !== -1) {
      state.expenses[index] = {
        ...state.expenses[index],
        ...payload,
        updatedAt: Date.now(),
      };
    }
  } else {
    state.expenses.push({
      id: crypto.randomUUID(),
      ...payload,
      createdAt: Date.now(),
    });
  }

  resetExpenseForm();
  saveData();
  renderAll();
}

function handleCaseListAction(event) {
  const btn = event.target;
  if (!(btn instanceof HTMLButtonElement)) return;
  const item = btn.closest(".item");
  if (!item) return;
  const id = item.dataset.id;

  if (btn.classList.contains("edit-btn")) {
    const target = state.cases.find((entry) => entry.id === id);
    if (!target) return;
    editState.caseId = target.id;
    caseForm.elements.customerName.value = target.customerName;
    caseForm.elements.caseName.value = target.caseName;
    caseForm.elements.amount.value = target.estimateAmount ?? "";
    caseForm.elements.receivedDate.value = target.receivedDate || "";
    caseForm.elements.dueDate.value = target.dueDate || "";
    caseForm.elements.status.value = normalizeStatus(target.status);
    caseSubmitBtn.textContent = "案件を更新";
    return;
  }

  if (btn.classList.contains("delete-btn")) {
    if (!window.confirm("この案件を削除しますか？関連する売上・経費も削除されます。")) return;
    state.cases = state.cases.filter((entry) => entry.id !== id);
    state.sales = state.sales.filter((entry) => entry.caseId !== id);
    state.expenses = state.expenses.filter((entry) => entry.caseId !== id);
    if (editState.caseId === id) resetCaseForm();
    saveData();
    renderAll();
  }
}

function handleSalesListAction(event) {
  const btn = event.target;
  if (!(btn instanceof HTMLButtonElement)) return;
  const item = btn.closest(".item");
  if (!item) return;
  const id = item.dataset.id;

  if (btn.classList.contains("edit-btn")) {
    const target = state.sales.find((entry) => entry.id === id);
    if (!target) return;
    editState.saleId = target.id;
    saleCaseSelect.value = target.caseId;
    saleForm.elements.invoiceAmount.value = target.invoiceAmount;
    saleForm.elements.paidAmount.value = target.paidAmount ?? "";
    saleForm.elements.paidDate.value = target.paidDate || "";
    saleForm.elements.isUnpaid.checked = Boolean(target.isUnpaid);
    saleSubmitBtn.textContent = "売上を更新";
    return;
  }

  if (btn.classList.contains("delete-btn")) {
    if (!window.confirm("この売上を削除しますか？")) return;
    state.sales = state.sales.filter((entry) => entry.id !== id);
    if (editState.saleId === id) resetSaleForm();
    saveData();
    renderAll();
  }
}

function handleExpensesListAction(event) {
  const btn = event.target;
  if (!(btn instanceof HTMLButtonElement)) return;
  const item = btn.closest(".item");
  if (!item) return;
  const id = item.dataset.id;

  if (btn.classList.contains("edit-btn")) {
    const target = state.expenses.find((entry) => entry.id === id);
    if (!target) return;
    editState.expenseId = target.id;
    expenseForm.elements.expenseDate.value = target.date;
    expenseForm.elements.expenseContent.value = target.content;
    expenseForm.elements.expenseAmount.value = target.amount;
    expenseCaseSelect.value = target.caseId || "";
    expenseSubmitBtn.textContent = "経費を更新";
    return;
  }

  if (btn.classList.contains("delete-btn")) {
    if (!window.confirm("この経費を削除しますか？")) return;
    state.expenses = state.expenses.filter((entry) => entry.id !== id);
    if (editState.expenseId === id) resetExpenseForm();
    saveData();
    renderAll();
  }
}

function renderAll() {
  renderCaseOptions();
  renderCases();
  renderSales();
  renderExpenses();
  renderDashboard();
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
    { label: "対象月", value: monthLabel(currentMonth), cls: "" },
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
  const sorted = state.sales.slice().sort((a, b) => (b.updatedAt ?? b.createdAt) - (a.updatedAt ?? a.createdAt));

  sorted.forEach((sale) => {
    const node = saleItemTemplate.content.cloneNode(true);
    const item = node.querySelector(".item");
    const title = node.querySelector(".title");
    const meta = node.querySelector(".meta");

    item.dataset.id = sale.id;
    title.textContent = `${resolveCaseName(sale.caseId)}｜請求: ${formatCurrency(sale.invoiceAmount)}`;
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

function resetCaseForm() {
  resetEditMode("case");
  caseForm.reset();
  caseForm.elements.status.value = "未着手";
  caseSubmitBtn.textContent = "案件を追加";
}

function resetSaleForm() {
  resetEditMode("sale");
  saleForm.reset();
  saleSubmitBtn.textContent = "売上を登録";
}

function resetExpenseForm() {
  resetEditMode("expense");
  expenseForm.reset();
  expenseSubmitBtn.textContent = "経費を登録";
}

function resetEditMode(target) {
  if (target === "case") editState.caseId = null;
  if (target === "sale") editState.saleId = null;
  if (target === "expense") editState.expenseId = null;
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

function saveData() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function loadData() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) return sanitizeData(JSON.parse(raw));

    const oldV2Raw = localStorage.getItem(LEGACY_STORAGE_KEYS[0]);
    if (oldV2Raw) return sanitizeData(JSON.parse(oldV2Raw));

    const legacyCasesRaw = localStorage.getItem(LEGACY_STORAGE_KEYS[1]);
    if (legacyCasesRaw) {
      return sanitizeData({ cases: JSON.parse(legacyCasesRaw), sales: [], expenses: [] });
    }
  } catch {
    return { cases: [], sales: [], expenses: [] };
  }

  return { cases: [], sales: [], expenses: [] };
}

function sanitizeData(input) {
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
          updatedAt: Number.isFinite(s.updatedAt) ? s.updatedAt : Date.now(),
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
        updatedAt: Number.isFinite(e.updatedAt) ? e.updatedAt : Date.now(),
      }))
      .filter((e) => e.date && e.content),
  };
}
