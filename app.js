const SUPABASE_URL = "https://ueelzyftlbnvjvpsmpyt.supabase.co";
const SUPABASE_KEY = "sb_publishable_0DrKsieUcCyEZN_HRg8LhQ_QqFTPMtp";
const STATUS_ORDER = ["未着手", "進行中", "完了"];

const sbClient = window.supabase.createClient(SUPABASE_URL, SUPABASE_KEY);

const state = { cases: [], sales: [], expenses: [] };
const editState = { caseId: null, saleId: null, expenseId: null };

const authView = document.getElementById("auth-view");
const appView = document.getElementById("app-view");
const loadingOverlay = document.getElementById("loading-overlay");
const authForm = document.getElementById("auth-form");
const authMessage = document.getElementById("auth-message");
const signupBtn = document.getElementById("signup-btn");
const logoutBtn = document.getElementById("logout-btn");
const userLabel = document.getElementById("user-label");
const appMessage = document.getElementById("app-message");

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

let currentUser = null;
let loadingCount = 0;

initialize();

async function initialize() {
  setLoading(true);
  setAuthControlsDisabled(true);
  try {
    bindEvents();

    const {
      data: { session },
    } = await sbClient.auth.getSession();

    currentUser = session?.user ?? null;
    await applyAuthState();

    sbClient.auth.onAuthStateChange(async (_event, sessionState) => {
      setLoading(true);
      try {
        currentUser = sessionState?.user ?? null;
        await applyAuthState();
      } catch (error) {
        console.error("認証状態の更新に失敗しました。", error);
        showAuthMessage("認証状態の確認に失敗しました。ページを再読み込みしてください。", true);
      } finally {
        setLoading(false);
      }
    });
  } catch (error) {
    console.error("初期化処理に失敗しました。", error);
    showAuthMessage("初期化に失敗しました。ページを再読み込みしてください。", true);
    authView.hidden = false;
    appView.hidden = true;
  } finally {
    setLoading(false);
    setAuthControlsDisabled(false);
  }
}

function bindEvents() {
  tabs.forEach((btn) => btn.addEventListener("click", () => activateTab(btn.dataset.tab)));
  authForm.addEventListener("submit", handleLogin);
  signupBtn.addEventListener("click", handleSignup);
  logoutBtn.addEventListener("click", handleLogout);

  caseForm.addEventListener("submit", handleCaseSubmit);
  saleForm.addEventListener("submit", handleSaleSubmit);
  expenseForm.addEventListener("submit", handleExpenseSubmit);

  clearBtn.addEventListener("click", handleClearAll);
  caseList.addEventListener("click", handleCaseListAction);
  salesList.addEventListener("click", handleSalesListAction);
  expensesList.addEventListener("click", handleExpensesListAction);
  document.addEventListener("visibilitychange", handleVisibilityChange);
}

async function handleVisibilityChange() {
  if (document.hidden || !currentUser) return;

  setLoading(true);
  try {
    await loadAllData();
    renderAfterDataChanged();
  } catch (error) {
    console.error("画面復帰時の再読み込みに失敗しました。", error);
    showAppMessage("画面復帰時のデータ更新に失敗しました。", true);
  } finally {
    setLoading(false);
    clearLoadingState();
  }
}

async function applyAuthState() {
  if (!currentUser) {
    authView.hidden = false;
    appView.hidden = true;
    userLabel.textContent = "";
    clearAppMessage();
    return;
  }

  authView.hidden = true;
  appView.hidden = false;
  userLabel.textContent = currentUser.email || "ログイン中";

  await withLoading(async () => {
    await loadAllData();
    resetCaseForm();
    resetSaleForm();
    resetExpenseForm();
    renderAfterDataChanged();
  }, "データの読み込みに失敗しました。テーブル設定を確認してください。");
}

async function handleLogin(event) {
  event.preventDefault();
  const email = authForm.elements.email.value.trim();
  const password = authForm.elements.password.value;
  if (!email || !password) return;

  await withLoading(async () => {
    const { error } = await sbClient.auth.signInWithPassword({ email, password });
    if (error) throw error;
    showAuthMessage("ログインしました。", false);
  }, "ログインに失敗しました。メールアドレスまたはパスワードを確認してください。", true);
}

async function handleSignup() {
  const email = authForm.elements.email.value.trim();
  const password = authForm.elements.password.value;
  if (!email || !password) {
    showAuthMessage("メールアドレスとパスワードを入力してください。", true);
    return;
  }

  await withLoading(async () => {
    const { error } = await sbClient.auth.signUp({ email, password });
    if (error) throw error;
    showAuthMessage("新規登録が完了しました。確認メールを確認してください。", false);
  }, "新規登録に失敗しました。入力内容を確認してください。", true);
}

async function handleLogout() {
  await withLoading(async () => {
    const { error } = await sbClient.auth.signOut();
    if (error) throw error;
    showAuthMessage("ログアウトしました。", false);
    authForm.reset();
  }, "ログアウトに失敗しました。", true);
}

async function loadAllData() {
  if (!currentUser) return;

  const [casesRes, salesRes, expensesRes] = await Promise.all([
    sbClient.from("cases").select("*").eq("user_id", currentUser.id),
    sbClient.from("sales").select("*").eq("user_id", currentUser.id),
    sbClient.from("expenses").select("*").eq("user_id", currentUser.id),
  ]);

  if (casesRes.error) throw casesRes.error;
  if (salesRes.error) throw salesRes.error;
  if (expensesRes.error) throw expensesRes.error;

  state.cases = (casesRes.data || []).map(mapCaseFromDb);
  state.sales = (salesRes.data || []).map(mapSaleFromDb);
  state.expenses = (expensesRes.data || []).map(mapExpenseFromDb);
}

async function refreshAfterMutation(successMessage) {
  await loadAllData();
  renderAfterDataChanged();
  if (successMessage) showAppMessage(successMessage, false);
}

async function handleCaseSubmit(event) {
  event.preventDefault();
  if (!currentUser) return;

  const customerName = caseForm.elements.customerName.value.trim();
  const caseName = caseForm.elements.caseName.value.trim();
  if (!customerName || !caseName) return;

  const payload = {
    user_id: currentUser.id,
    customer_name: customerName,
    case_name: caseName,
    estimate_amount: normalizeAmount(caseForm.elements.amount.value),
    received_date: caseForm.elements.receivedDate.value || null,
    due_date: caseForm.elements.dueDate.value || null,
    status: normalizeStatus(caseForm.elements.status.value),
  };

  await withLoading(async () => {
    if (editState.caseId) {
      const { error } = await sbClient.from("cases").update(payload).eq("id", editState.caseId).eq("user_id", currentUser.id);
      if (error) throw error;
      resetCaseForm();
      await refreshAfterMutation("案件を更新しました。");
      return;
    }

    const { error } = await sbClient.from("cases").insert(payload);
    if (error) throw error;
    resetCaseForm();
    await refreshAfterMutation("案件を登録しました。");
  }, "案件の保存に失敗しました。入力内容を確認してください。");
}

async function handleSaleSubmit(event) {
  event.preventDefault();
  if (!currentUser || !state.cases.length) return;

  const invoiceAmount = normalizeAmount(saleForm.elements.invoiceAmount.value);
  if (invoiceAmount === null) return;

  const paidAmount = normalizeAmount(saleForm.elements.paidAmount.value);
  const caseId = saleCaseSelect.value;
  if (!caseId) return;

  const payload = {
    user_id: currentUser.id,
    case_id: caseId,
    invoice_amount: invoiceAmount,
    paid_amount: paidAmount,
    paid_date: saleForm.elements.paidDate.value || null,
    is_unpaid: saleForm.elements.isUnpaid.checked || (paidAmount ?? 0) < invoiceAmount,
  };

  await withLoading(async () => {
    if (editState.saleId) {
      const { error } = await sbClient.from("sales").update(payload).eq("id", editState.saleId).eq("user_id", currentUser.id);
      if (error) throw error;
      resetSaleForm();
      await refreshAfterMutation("売上を更新しました。");
      return;
    }

    const { error } = await sbClient.from("sales").insert(payload);
    if (error) throw error;
    resetSaleForm();
    await refreshAfterMutation("売上を登録しました。");
  }, "売上の保存に失敗しました。入力内容を確認してください。");
}

async function handleExpenseSubmit(event) {
  event.preventDefault();
  if (!currentUser) return;

  const date = expenseForm.elements.expenseDate.value;
  const content = expenseForm.elements.expenseContent.value.trim();
  const amount = normalizeAmount(expenseForm.elements.expenseAmount.value);
  if (!date || !content || amount === null) return;

  const payload = {
    user_id: currentUser.id,
    date,
    content,
    amount,
    case_id: expenseCaseSelect.value || null,
  };

  await withLoading(async () => {
    if (editState.expenseId) {
      const { error } = await sbClient.from("expenses").update(payload).eq("id", editState.expenseId).eq("user_id", currentUser.id);
      if (error) throw error;
      resetExpenseForm();
      await refreshAfterMutation("経費を更新しました。");
      return;
    }

    const { error } = await sbClient.from("expenses").insert(payload);
    if (error) throw error;
    resetExpenseForm();
    await refreshAfterMutation("経費を登録しました。");
  }, "経費の保存に失敗しました。入力内容を確認してください。");
}

async function handleCaseListAction(event) {
  const btn = event.target;
  if (!(btn instanceof HTMLButtonElement)) return;
  const item = btn.closest(".item");
  if (!item || !currentUser) return;
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

    await withLoading(async () => {
      const { error } = await sbClient.from("cases").delete().eq("id", id).eq("user_id", currentUser.id);
      if (error) throw error;
      if (editState.caseId === id) resetCaseForm();
      await refreshAfterMutation("案件を削除しました。");
    }, "案件の削除に失敗しました。");
  }
}

async function handleSalesListAction(event) {
  const btn = event.target;
  if (!(btn instanceof HTMLButtonElement)) return;
  const item = btn.closest(".item");
  if (!item || !currentUser) return;
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
    await withLoading(async () => {
      const { error } = await sbClient.from("sales").delete().eq("id", id).eq("user_id", currentUser.id);
      if (error) throw error;
      if (editState.saleId === id) resetSaleForm();
      await refreshAfterMutation("売上を削除しました。");
    }, "売上の削除に失敗しました。");
  }
}

async function handleExpensesListAction(event) {
  const btn = event.target;
  if (!(btn instanceof HTMLButtonElement)) return;
  const item = btn.closest(".item");
  if (!item || !currentUser) return;
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
    await withLoading(async () => {
      const { error } = await sbClient.from("expenses").delete().eq("id", id).eq("user_id", currentUser.id);
      if (error) throw error;
      if (editState.expenseId === id) resetExpenseForm();
      await refreshAfterMutation("経費を削除しました。");
    }, "経費の削除に失敗しました。");
  }
}

async function handleClearAll() {
  if (!currentUser) return;
  if (!window.confirm("案件・売上・経費の全データを削除します。よろしいですか？")) return;

  await withLoading(async () => {
    const [salesRes, expensesRes, casesRes] = await Promise.all([
      sbClient.from("sales").delete().eq("user_id", currentUser.id),
      sbClient.from("expenses").delete().eq("user_id", currentUser.id),
      sbClient.from("cases").delete().eq("user_id", currentUser.id),
    ]);

    if (salesRes.error) throw salesRes.error;
    if (expensesRes.error) throw expensesRes.error;
    if (casesRes.error) throw casesRes.error;

    resetCaseForm();
    resetSaleForm();
    resetExpenseForm();
    await refreshAfterMutation("全データを削除しました。");
  }, "全件削除に失敗しました。");
}

function renderAfterDataChanged() {
  renderCaseOptions();
  renderCases();
  renderSales();
  renderExpenses();
  renderDashboard();
}

function activateTab(tabKey) {
  tabs.forEach((btn) => btn.classList.toggle("active", btn.dataset.tab === tabKey));
  Object.entries(panels).forEach(([key, panel]) => panel.classList.toggle("active", key === tabKey));
}

function renderDashboard() {
  const salesByMonth = aggregateByMonth(state.sales, (s) => s.paidDate || s.createdAt, (s) => s.invoiceAmount);
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
  const sorted = state.expenses.slice().sort((a, b) => toSortTimestamp(b.date) - toSortTimestamp(a.date));

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
  return found ? `${found.customerName}｜${found.caseName}` : "（削除済み案件）";
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

function mapCaseFromDb(row) {
  return {
    id: row.id,
    customerName: row.customer_name || "",
    caseName: row.case_name || "",
    estimateAmount: normalizeAmount(row.estimate_amount),
    receivedDate: row.received_date || "",
    dueDate: row.due_date || "",
    status: normalizeStatus(row.status),
    createdAt: Date.parse(row.created_at) || Date.now(),
    updatedAt: Date.parse(row.updated_at) || Date.now(),
  };
}

function mapSaleFromDb(row) {
  const invoiceAmount = normalizeAmount(row.invoice_amount) ?? 0;
  const paidAmount = normalizeAmount(row.paid_amount);
  return {
    id: row.id,
    caseId: row.case_id,
    invoiceAmount,
    paidAmount,
    paidDate: row.paid_date || "",
    isUnpaid: Boolean(row.is_unpaid) || (paidAmount ?? 0) < invoiceAmount,
    createdAt: Date.parse(row.created_at) || Date.now(),
    updatedAt: Date.parse(row.updated_at) || Date.now(),
  };
}

function mapExpenseFromDb(row) {
  return {
    id: row.id,
    date: row.date || "",
    content: row.content || "",
    amount: normalizeAmount(row.amount) ?? 0,
    caseId: row.case_id || null,
    createdAt: Date.parse(row.created_at) || Date.now(),
    updatedAt: Date.parse(row.updated_at) || Date.now(),
  };
}

async function withLoading(task, fallbackMessage, authArea = false) {
  setLoading(true);
  try {
    clearAppMessage();
    if (authArea) showAuthMessage("", false);
    await task();
  } catch (error) {
    console.error("処理に失敗しました。", error);
    const message = error?.message ? String(error.message) : fallbackMessage;
    if (authArea) {
      showAuthMessage(`${fallbackMessage}\n詳細: ${message}`, true);
    } else {
      showAppMessage(`${fallbackMessage}\n詳細: ${message}`, true);
    }
  } finally {
    setLoading(false);
  }
}

function setLoading(isLoading) {
  if (!loadingOverlay) return;

  if (isLoading) {
    loadingCount += 1;
    loadingOverlay.hidden = false;
    return;
  }
  loadingCount = Math.max(0, loadingCount - 1);
  if (loadingCount === 0) loadingOverlay.hidden = true;
}

function clearLoadingState() {
  loadingCount = 0;
  if (!loadingOverlay) return;
  loadingOverlay.hidden = true;
}

function setAuthControlsDisabled(disabled) {
  const loginBtn = authForm?.querySelector('button[type="submit"]');
  if (loginBtn) loginBtn.disabled = disabled;
  if (signupBtn) signupBtn.disabled = disabled;
}

function showAuthMessage(text, isError) {
  authMessage.textContent = text;
  authMessage.classList.toggle("error", isError);
}

function showAppMessage(text, isError) {
  appMessage.textContent = text;
  appMessage.classList.toggle("error", isError);
}

function clearAppMessage() {
  showAppMessage("", false);
}
