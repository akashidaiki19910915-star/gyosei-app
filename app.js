const SUPABASE_URL = "https://ueelzyftlbnvjvpsmpyt.supabase.co";
const SUPABASE_KEY = "sb_publishable_0DrKsieUcCyEZN_HRg8LhQ_QqFTPMtp";
const STATUS_ORDER = ["未着手", "進行中", "完了"];
const STATUS_FILTER_KEYS = [...STATUS_ORDER, "その他"];
const DEADLINE_FILTER_KEYS = ["all", "overdue", "within7", "within30"];

const sbClient = window.supabase.createClient(SUPABASE_URL, SUPABASE_KEY, {
  auth: {
    persistSession: true,
    autoRefreshToken: true,
    detectSessionInUrl: true,
    storageKey: "gyosei-app-auth",
  },
});

const state = {
  cases: [],
  sales: [],
  expenses: [],
  fixedExpenses: [],
  selectedAggregation: "month",
  selectedMonth: toMonthKey(new Date()),
  selectedYear: new Date().getFullYear(),
  caseStatusFilter: "all",
  caseSearchQuery: "",
  caseDeadlineFilter: "all",
  salesSearchQuery: "",
  expensesSearchQuery: "",
};
const editState = { caseId: null, saleId: null, expenseId: null, fixedExpenseId: null };

const authView = document.getElementById("auth-view");
const appView = document.getElementById("app-view");
const loadingOverlay = document.getElementById("loading-overlay");
const authForm = document.getElementById("auth-form");
const authMessage = document.getElementById("auth-message");
const signupBtn = document.getElementById("signup-btn");
const logoutBtn = document.getElementById("logout-btn");
const userLabel = document.getElementById("user-label");
const appMessage = document.getElementById("app-message");
const exportCasesCsvBtn = document.getElementById("export-cases-csv-btn");
const exportSalesCsvBtn = document.getElementById("export-sales-csv-btn");
const exportExpensesCsvBtn = document.getElementById("export-expenses-csv-btn");
const exportFixedExpensesCsvBtn = document.getElementById("export-fixed-expenses-csv-btn");
const exportAllCsvBtn = document.getElementById("export-all-csv-btn");
const csvImportForm = document.getElementById("csv-import-form");
const exportExcelBtn = document.getElementById("export-excel-btn");
const excelImportForm = document.getElementById("excel-import-form");

const tabs = Array.from(document.querySelectorAll(".tab-btn"));
const panels = {
  cases: document.getElementById("tab-cases"),
  sales: document.getElementById("tab-sales"),
  expenses: document.getElementById("tab-expenses"),
};

const summaryGrid = document.getElementById("summary-grid");
const caseProfitList = document.getElementById("case-profit-list");
const caseProfitEmpty = document.getElementById("case-profit-empty");
const targetMonthInput = document.getElementById("target-month");
const targetYearInput = document.getElementById("target-year");
const monthFilter = document.getElementById("month-filter");
const yearFilter = document.getElementById("year-filter");
const aggregationRadios = Array.from(document.querySelectorAll('input[name="aggregation-unit"]'));
const yearlyBreakdown = document.getElementById("yearly-breakdown");
const yearlyBreakdownBody = document.getElementById("yearly-breakdown-body");
const unpaidAlertCard = document.getElementById("unpaid-alert-card");
const unpaidAlertSummary = document.getElementById("unpaid-alert-summary");
const unpaidAlertPeriod = document.getElementById("unpaid-alert-period");
const unpaidAlertBody = document.getElementById("unpaid-alert-body");
const unpaidAlertEmpty = document.getElementById("unpaid-alert-empty");
const unpaidAlertListWrap = document.getElementById("unpaid-alert-list-wrap");
const billingLeakAlertCard = document.getElementById("billing-leak-alert-card");
const billingLeakAlertSummary = document.getElementById("billing-leak-alert-summary");
const billingLeakAlertBody = document.getElementById("billing-leak-alert-body");
const billingLeakAlertEmpty = document.getElementById("billing-leak-alert-empty");
const billingLeakAlertListWrap = document.getElementById("billing-leak-alert-list-wrap");
const unpaidListBody = document.getElementById("unpaid-list-body");
const unpaidListEmpty = document.getElementById("unpaid-list-empty");
const unpaidListWrap = document.getElementById("unpaid-list-wrap");
const statusSummaryList = document.getElementById("status-summary-list");
const statusSummaryTotal = document.getElementById("status-summary-total");
const statusFilterClearBtn = document.getElementById("status-filter-clear-btn");
const deadlineAlertCard = document.getElementById("deadline-alert-card");
const deadlineAlertSummary = document.getElementById("deadline-alert-summary");
const deadlineAlertBody = document.getElementById("deadline-alert-body");
const deadlineAlertEmpty = document.getElementById("deadline-alert-empty");
const deadlineAlertListWrap = document.getElementById("deadline-alert-list-wrap");

const caseForm = document.getElementById("case-form");
const caseList = document.getElementById("case-list");
const caseEmpty = document.getElementById("case-empty");
const clearBtn = document.getElementById("clear-btn");
const caseSubmitBtn = caseForm.querySelector('button[type="submit"]');
const caseSearchInput = document.getElementById("case-search-input");
const caseStatusFilterSelect = document.getElementById("case-status-filter");
const caseDeadlineFilterSelect = document.getElementById("case-deadline-filter");
const caseFilterClearBtn = document.getElementById("case-filter-clear-btn");
const caseFilterCount = document.getElementById("case-filter-count");

const saleForm = document.getElementById("sale-form");
const saleCaseSelect = document.getElementById("sale-case-id");
const salesList = document.getElementById("sales-list");
const salesEmpty = document.getElementById("sales-empty");
const saleSubmitBtn = saleForm.querySelector('button[type="submit"]');
const salesSearchInput = document.getElementById("sales-search-input");
const salesFilterClearBtn = document.getElementById("sales-filter-clear-btn");
const salesFilterCount = document.getElementById("sales-filter-count");

const expenseForm = document.getElementById("expense-form");
const expenseCaseSelect = document.getElementById("expense-case-id");
const expensesList = document.getElementById("expenses-list");
const expensesEmpty = document.getElementById("expenses-empty");
const expenseSubmitBtn = expenseForm.querySelector('button[type="submit"]');
const expensesSearchInput = document.getElementById("expenses-search-input");
const expensesFilterClearBtn = document.getElementById("expenses-filter-clear-btn");
const expensesFilterCount = document.getElementById("expenses-filter-count");
const fixedExpenseForm = document.getElementById("fixed-expense-form");
const fixedExpensesList = document.getElementById("fixed-expenses-list");
const fixedExpensesEmpty = document.getElementById("fixed-expenses-empty");
const fixedExpenseSubmitBtn = fixedExpenseForm.querySelector('button[type="submit"]');

const caseItemTemplate = document.getElementById("case-item-template");
const saleItemTemplate = document.getElementById("sale-item-template");
const expenseItemTemplate = document.getElementById("expense-item-template");
const fixedExpenseItemTemplate = document.getElementById("fixed-expense-item-template");

let currentUser = null;
let loadingCount = 0;
let isResuming = false;

initialize();

if (window.location.protocol === "file:") {
  document.body.innerHTML = `
    <main class="auth-container">
      <section class="panel auth-panel">
        <h1>このアプリは https:// で開いてください</h1>
        <p class="auth-caption">GitHub Pages などの HTTPS 環境でのみ、Supabase Auth のログイン状態を正しく保持できます。</p>
      </section>
    </main>
  `;
  throw new Error("file:// では Supabase Auth を利用できません。HTTPS で配信してください。");
}

async function initialize() {
  showLoading(true);
  setAuthControlsDisabled(true);
  try {
    bindEvents();

    const {
      data: { session },
    } = await sbClient.auth.getSession();

    currentUser = session?.user ?? null;
    await applyAuthState();

    sbClient.auth.onAuthStateChange(async (_event, sessionState) => {
      showLoading(true);
      try {
        currentUser = sessionState?.user ?? null;
        await applyAuthState();
      } catch (error) {
        console.error("認証状態の更新に失敗しました。", error);
        showAuthMessage("認証状態の確認に失敗しました。ページを再読み込みしてください。", true);
      } finally {
        showLoading(false);
      }
    });
  } catch (error) {
    console.error("初期化処理に失敗しました。", error);
    showAuthMessage("初期化に失敗しました。ページを再読み込みしてください。", true);
    authView.hidden = false;
    appView.hidden = true;
  } finally {
    showLoading(false);
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
  fixedExpenseForm.addEventListener("submit", handleFixedExpenseSubmit);

  clearBtn.addEventListener("click", handleClearAll);
  caseList.addEventListener("click", handleCaseListAction);
  salesList.addEventListener("click", handleSalesListAction);
  expensesList.addEventListener("click", handleExpensesListAction);
  fixedExpensesList.addEventListener("click", handleFixedExpensesListAction);
  unpaidListBody?.addEventListener("click", handleUnpaidListAction);
  billingLeakAlertBody?.addEventListener("click", handleBillingLeakAlertAction);
  targetMonthInput?.addEventListener("input", handleTargetMonthChange);
  targetYearInput?.addEventListener("input", handleTargetYearChange);
  aggregationRadios.forEach((radio) => radio.addEventListener("change", handleAggregationChange));
  statusSummaryList?.addEventListener("click", handleStatusSummaryClick);
  statusFilterClearBtn?.addEventListener("click", () => applyCaseStatusFilter("all"));
  deadlineAlertSummary?.addEventListener("click", handleDeadlineAlertClick);
  deadlineAlertBody?.addEventListener("click", handleDeadlineAlertClick);
  caseSearchInput?.addEventListener("input", handleCaseSearchInput);
  caseStatusFilterSelect?.addEventListener("change", handleCaseStatusFilterChange);
  caseDeadlineFilterSelect?.addEventListener("change", handleCaseDeadlineFilterChange);
  caseFilterClearBtn?.addEventListener("click", clearCaseFilters);
  salesSearchInput?.addEventListener("input", handleSalesSearchInput);
  salesFilterClearBtn?.addEventListener("click", clearSalesSearch);
  expensesSearchInput?.addEventListener("input", handleExpensesSearchInput);
  expensesFilterClearBtn?.addEventListener("click", clearExpensesSearch);
  document.addEventListener("visibilitychange", handleVisibilityChange);
  window.addEventListener("focus", handleWindowFocus);
  window.addEventListener("pageshow", handlePageShow);
  exportCasesCsvBtn?.addEventListener("click", handleExportCasesCsv);
  exportSalesCsvBtn?.addEventListener("click", handleExportSalesCsv);
  exportExpensesCsvBtn?.addEventListener("click", handleExportExpensesCsv);
  exportFixedExpensesCsvBtn?.addEventListener("click", handleExportFixedExpensesCsv);
  exportAllCsvBtn?.addEventListener("click", handleExportAllCsv);
  csvImportForm?.addEventListener("submit", handleCsvImportSubmit);
  exportExcelBtn?.addEventListener("click", handleExportExcel);
  excelImportForm?.addEventListener("submit", handleExcelImportSubmit);
}

function handleAggregationChange(event) {
  const next = event?.target?.value;
  state.selectedAggregation = next === "year" ? "year" : "month";
  renderDashboard();
}

function handleTargetMonthChange(event) {
  const nextValue = normalizeMonthKey(event?.target?.value);
  state.selectedMonth = nextValue || toMonthKey(new Date());
  renderDashboard();
}

function handleTargetYearChange(event) {
  const nextValue = normalizeYear(event?.target?.value);
  state.selectedYear = nextValue || new Date().getFullYear();
  renderDashboard();
}

function handleStatusSummaryClick(event) {
  const btn = event.target.closest(".status-filter-btn");
  if (!btn) return;
  applyCaseStatusFilter(btn.dataset.statusFilter || "all");
}

function applyCaseStatusFilter(nextFilter) {
  if (nextFilter !== "all" && !STATUS_FILTER_KEYS.includes(nextFilter)) return;
  state.caseStatusFilter = nextFilter;
  if (caseStatusFilterSelect) caseStatusFilterSelect.value = nextFilter;
  activateTab("cases");
  renderDashboard();
  renderCases();
}

function handleCaseSearchInput(event) {
  state.caseSearchQuery = String(event?.target?.value ?? "").trim().toLowerCase();
  renderCases();
}

function handleCaseStatusFilterChange(event) {
  applyCaseStatusFilter(event?.target?.value || "all");
}

function handleCaseDeadlineFilterChange(event) {
  applyCaseDeadlineFilter(event?.target?.value || "all");
}

function clearCaseFilters() {
  state.caseSearchQuery = "";
  state.caseDeadlineFilter = "all";
  applyCaseStatusFilter("all");
  if (caseSearchInput) caseSearchInput.value = "";
  if (caseDeadlineFilterSelect) caseDeadlineFilterSelect.value = "all";
  renderCases();
}

function handleDeadlineAlertClick(event) {
  const button = event.target.closest("button");
  if (!(button instanceof HTMLButtonElement)) return;
  const filter = button.dataset.deadlineFilter;
  if (filter) {
    applyCaseDeadlineFilter(filter, { activateCases: true });
    return;
  }
  const caseId = button.dataset.caseId;
  if (caseId) startCaseEdit(caseId);
}

function applyCaseDeadlineFilter(nextFilter, options = {}) {
  const normalizedFilter = DEADLINE_FILTER_KEYS.includes(nextFilter) ? nextFilter : "all";
  state.caseDeadlineFilter = normalizedFilter;
  if (caseDeadlineFilterSelect) caseDeadlineFilterSelect.value = normalizedFilter;
  if (options.activateCases) activateTab("cases");
  renderCases();
}

function handleSalesSearchInput(event) {
  state.salesSearchQuery = String(event?.target?.value ?? "").trim().toLowerCase();
  renderSales();
}

function clearSalesSearch() {
  state.salesSearchQuery = "";
  if (salesSearchInput) salesSearchInput.value = "";
  renderSales();
}

function handleExpensesSearchInput(event) {
  state.expensesSearchQuery = String(event?.target?.value ?? "").trim().toLowerCase();
  renderExpenses();
}

function clearExpensesSearch() {
  state.expensesSearchQuery = "";
  if (expensesSearchInput) expensesSearchInput.value = "";
  renderExpenses();
}


async function handleVisibilityChange() {
  if (document.hidden) return;
  showLoading(false);
  await handleResumeRefresh("visibilitychange");
}

async function handleWindowFocus() {
  showLoading(false);
  await handleResumeRefresh("focus");
}

async function handlePageShow() {
  showLoading(false);
  await handleResumeRefresh("pageshow");
}

async function handleResumeRefresh(trigger) {
  if (!currentUser || isResuming) return;

  isResuming = true;
  showLoading(true);
  try {
    await loadAllData();
    renderAfterDataChanged();
  } catch (error) {
    console.error(`画面復帰時の再読み込みに失敗しました。(${trigger})`, error);
    showAppMessage("画面復帰時のデータ更新に失敗しました。", true);
  } finally {
    isResuming = false;
    showLoading(false);
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
    resetFixedExpenseForm();
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

  const [casesRes, salesRes, expensesRes, fixedExpensesRes] = await Promise.all([
    sbClient.from("cases").select("*").eq("user_id", currentUser.id),
    sbClient.from("sales").select("*").eq("user_id", currentUser.id),
    sbClient.from("expenses").select("*").eq("user_id", currentUser.id),
    sbClient.from("fixed_expenses").select("*").eq("user_id", currentUser.id),
  ]);

  if (casesRes.error) throw casesRes.error;
  if (salesRes.error) throw salesRes.error;
  if (expensesRes.error) throw expensesRes.error;
  if (fixedExpensesRes.error) throw fixedExpensesRes.error;

  state.cases = (casesRes.data || []).map(mapCaseFromDb);
  state.sales = (salesRes.data || []).map(mapSaleFromDb);
  state.expenses = (expensesRes.data || []).map(mapExpenseFromDb);
  state.fixedExpenses = (fixedExpensesRes.data || []).map(mapFixedExpenseFromDb);

  const createdCount = await ensureMonthlyFixedExpenses();
  if (createdCount > 0) {
    const refreshExpensesRes = await sbClient.from("expenses").select("*").eq("user_id", currentUser.id);
    if (refreshExpensesRes.error) throw refreshExpensesRes.error;
    state.expenses = (refreshExpensesRes.data || []).map(mapExpenseFromDb);
    showAppMessage(`${createdCount}件の固定費を当月分として自動計上しました。`, false);
  }
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
      const currentExpense = state.expenses.find((entry) => entry.id === editState.expenseId);
      const updatePayload = { ...payload };
      if (currentExpense?.fixedExpenseId) updatePayload.fixed_expense_id = currentExpense.fixedExpenseId;
      const { error } = await sbClient.from("expenses").update(updatePayload).eq("id", editState.expenseId).eq("user_id", currentUser.id);
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

async function handleFixedExpenseSubmit(event) {
  event.preventDefault();
  if (!currentUser) return;

  const content = fixedExpenseForm.elements.fixedExpenseContent.value.trim();
  const amount = normalizeAmount(fixedExpenseForm.elements.fixedExpenseAmount.value);
  const dayOfMonth = normalizeDayOfMonth(fixedExpenseForm.elements.fixedExpenseDayOfMonth.value);
  const startDate = fixedExpenseForm.elements.fixedExpenseStartDate.value;
  if (!content || amount === null || !dayOfMonth || !startDate) return;

  const payload = {
    user_id: currentUser.id,
    content,
    amount,
    day_of_month: dayOfMonth,
    start_date: startDate,
    active: Boolean(fixedExpenseForm.elements.fixedExpenseActive.checked),
  };

  await withLoading(async () => {
    if (editState.fixedExpenseId) {
      const { error } = await sbClient
        .from("fixed_expenses")
        .update(payload)
        .eq("id", editState.fixedExpenseId)
        .eq("user_id", currentUser.id);
      if (error) throw error;
      resetFixedExpenseForm();
      await refreshAfterMutation("固定費を更新しました。");
      return;
    }

    const { error } = await sbClient.from("fixed_expenses").insert(payload);
    if (error) throw error;
    resetFixedExpenseForm();
    await refreshAfterMutation("固定費を登録しました。");
  }, "固定費の保存に失敗しました。入力内容を確認してください。");
}

async function handleCaseListAction(event) {
  const btn = event.target;
  if (!(btn instanceof HTMLButtonElement)) return;
  const item = btn.closest(".item");
  if (!item || !currentUser) return;
  const id = item.dataset.id;

  if (btn.classList.contains("edit-btn")) {
    startCaseEdit(id);
    return;
  }

  if (btn.classList.contains("delete-btn")) {
    if (!window.confirm("この案件を削除しますか？関連する売上・経費も削除されます。")) return;

    await withLoading(async () => {
      const salesDeleteRes = await sbClient
        .from("sales")
        .delete()
        .eq("case_id", id)
        .eq("user_id", currentUser.id);
      if (salesDeleteRes.error) throw salesDeleteRes.error;

      const expensesDeleteRes = await sbClient
        .from("expenses")
        .delete()
        .eq("case_id", id)
        .eq("user_id", currentUser.id);
      if (expensesDeleteRes.error) throw expensesDeleteRes.error;

      const caseDeleteRes = await sbClient
        .from("cases")
        .delete()
        .eq("id", id)
        .eq("user_id", currentUser.id);
      if (caseDeleteRes.error) throw caseDeleteRes.error;

      if (editState.caseId === id) resetCaseForm();
      await refreshAfterMutation("案件を削除しました。");
    }, "案件と関連データの削除に失敗しました。");
  }
}

function startCaseEdit(caseId) {
  const target = state.cases.find((entry) => entry.id === caseId);
  if (!target) return;
  activateTab("cases");
  editState.caseId = target.id;
  caseForm.elements.customerName.value = target.customerName;
  caseForm.elements.caseName.value = target.caseName;
  caseForm.elements.amount.value = target.estimateAmount ?? "";
  caseForm.elements.receivedDate.value = target.receivedDate || "";
  caseForm.elements.dueDate.value = target.dueDate || "";
  caseForm.elements.status.value = normalizeStatus(target.status);
  caseSubmitBtn.textContent = "案件を更新";
  caseForm.scrollIntoView({ behavior: "smooth", block: "start" });
}

async function handleSalesListAction(event) {
  const btn = event.target;
  if (!(btn instanceof HTMLButtonElement)) return;
  const item = btn.closest(".item");
  if (!item || !currentUser) return;
  const id = item.dataset.id;

  if (btn.classList.contains("edit-btn")) {
    startSaleEdit(id);
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

function handleUnpaidListAction(event) {
  const btn = event.target;
  if (!(btn instanceof HTMLButtonElement)) return;
  const saleId = btn.dataset.saleId;
  if (!saleId) return;
  if (btn.classList.contains("edit-sale-btn")) {
    startSaleEdit(saleId);
  }
}

function handleBillingLeakAlertAction(event) {
  const btn = event.target;
  if (!(btn instanceof HTMLButtonElement)) return;
  const caseId = btn.dataset.caseId;
  if (!caseId) return;
  if (btn.classList.contains("register-sale-btn")) {
    openSaleFormForCase(caseId);
  }
}

function startSaleEdit(saleId) {
  const target = state.sales.find((entry) => entry.id === saleId);
  if (!target) return;
  activateTab("sales");
  editState.saleId = target.id;
  saleCaseSelect.value = target.caseId;
  saleForm.elements.invoiceAmount.value = target.invoiceAmount;
  saleForm.elements.paidAmount.value = target.paidAmount ?? "";
  saleForm.elements.paidDate.value = target.paidDate || "";
  saleForm.elements.isUnpaid.checked = Boolean(target.isUnpaid);
  saleSubmitBtn.textContent = "売上を更新";
  saleForm.scrollIntoView({ behavior: "smooth", block: "start" });
}

function openSaleFormForCase(caseId) {
  const targetCase = state.cases.find((entry) => entry.id === caseId);
  if (!targetCase || !saleCaseSelect) return;
  activateTab("sales");
  resetSaleForm();
  saleCaseSelect.value = targetCase.id;
  saleForm.elements.invoiceAmount.value = targetCase.estimateAmount ?? "";
  saleForm.elements.paidAmount.value = "";
  saleForm.elements.paidDate.value = "";
  saleForm.elements.isUnpaid.checked = true;
  saleForm.scrollIntoView({ behavior: "smooth", block: "start" });
}

async function handleFixedExpensesListAction(event) {
  const btn = event.target;
  if (!(btn instanceof HTMLButtonElement)) return;
  const item = btn.closest(".item");
  if (!item || !currentUser) return;
  const id = item.dataset.id;

  if (btn.classList.contains("edit-btn")) {
    const target = state.fixedExpenses.find((entry) => entry.id === id);
    if (!target) return;
    editState.fixedExpenseId = target.id;
    fixedExpenseForm.elements.fixedExpenseContent.value = target.content;
    fixedExpenseForm.elements.fixedExpenseAmount.value = target.amount;
    fixedExpenseForm.elements.fixedExpenseDayOfMonth.value = target.dayOfMonth;
    fixedExpenseForm.elements.fixedExpenseStartDate.value = target.startDate || "";
    fixedExpenseForm.elements.fixedExpenseActive.checked = Boolean(target.active);
    fixedExpenseSubmitBtn.textContent = "固定費を更新";
    return;
  }

  if (btn.classList.contains("toggle-btn")) {
    const target = state.fixedExpenses.find((entry) => entry.id === id);
    if (!target) return;

    await withLoading(async () => {
      const { error } = await sbClient
        .from("fixed_expenses")
        .update({ active: !target.active })
        .eq("id", id)
        .eq("user_id", currentUser.id);
      if (error) throw error;
      await refreshAfterMutation("固定費の有効/無効を更新しました。");
    }, "固定費の有効/無効更新に失敗しました。");
    return;
  }

  if (btn.classList.contains("delete-btn")) {
    if (!window.confirm("この固定費を削除しますか？")) return;
    await withLoading(async () => {
      const { error } = await sbClient.from("fixed_expenses").delete().eq("id", id).eq("user_id", currentUser.id);
      if (error) throw error;
      if (editState.fixedExpenseId === id) resetFixedExpenseForm();
      await refreshAfterMutation("固定費を削除しました。");
    }, "固定費の削除に失敗しました。");
  }
}

async function handleClearAll() {
  if (!currentUser) return;
  if (!window.confirm("案件・売上・経費・固定費の全データを削除します。よろしいですか？")) return;

  await withLoading(async () => {
    const [salesRes, expensesRes, fixedExpensesRes, casesRes] = await Promise.all([
      sbClient.from("sales").delete().eq("user_id", currentUser.id),
      sbClient.from("expenses").delete().eq("user_id", currentUser.id),
      sbClient.from("fixed_expenses").delete().eq("user_id", currentUser.id),
      sbClient.from("cases").delete().eq("user_id", currentUser.id),
    ]);

    if (salesRes.error) throw salesRes.error;
    if (expensesRes.error) throw expensesRes.error;
    if (fixedExpensesRes.error) throw fixedExpensesRes.error;
    if (casesRes.error) throw casesRes.error;

    resetCaseForm();
    resetSaleForm();
    resetExpenseForm();
    resetFixedExpenseForm();
    await refreshAfterMutation("全データを削除しました。");
  }, "全件削除に失敗しました。");
}

function handleExportCasesCsv() {
  const rows = state.cases.map((entry) => ({
    customer_name: entry.customerName,
    case_name: entry.caseName,
    estimate_amount: entry.estimateAmount ?? "",
    received_date: entry.receivedDate || "",
    due_date: entry.dueDate || "",
    status: normalizeStoredStatus(entry.status),
  }));
  downloadCsvFile("cases.csv", ["customer_name", "case_name", "estimate_amount", "received_date", "due_date", "status"], rows);
}

function handleExportSalesCsv() {
  const rows = state.sales.map((entry) => {
    const foundCase = state.cases.find((c) => c.id === entry.caseId);
    return {
      case_name: foundCase?.caseName || "",
      invoice_amount: entry.invoiceAmount ?? "",
      paid_amount: entry.paidAmount ?? "",
      paid_date: entry.paidDate || "",
      is_unpaid: entry.isUnpaid ? "true" : "false",
    };
  });
  downloadCsvFile("sales.csv", ["case_name", "invoice_amount", "paid_amount", "paid_date", "is_unpaid"], rows);
}

function handleExportExpensesCsv() {
  const rows = state.expenses.map((entry) => {
    const foundCase = state.cases.find((c) => c.id === entry.caseId);
    return {
      case_name: foundCase?.caseName || "",
      date: entry.date || "",
      content: entry.content || "",
      amount: entry.amount ?? "",
    };
  });
  downloadCsvFile("expenses.csv", ["case_name", "date", "content", "amount"], rows);
}

function handleExportFixedExpensesCsv() {
  const rows = state.fixedExpenses.map((entry) => ({
    content: entry.content,
    amount: entry.amount ?? "",
    day_of_month: entry.dayOfMonth ?? "",
    start_date: entry.startDate || "",
    active: entry.active ? "true" : "false",
  }));
  downloadCsvFile("fixed_expenses.csv", ["content", "amount", "day_of_month", "start_date", "active"], rows);
}

function handleExportAllCsv() {
  const headers = [
    "data_type",
    "customer_name", "case_name", "estimate_amount", "received_date", "due_date", "status",
    "invoice_amount", "paid_amount", "paid_date", "is_unpaid",
    "date", "content", "amount",
    "day_of_month", "start_date", "active",
  ];
  const rows = [];

  state.cases.forEach((entry) => {
    rows.push({
      data_type: "case",
      customer_name: entry.customerName,
      case_name: entry.caseName,
      estimate_amount: entry.estimateAmount ?? "",
      received_date: entry.receivedDate || "",
      due_date: entry.dueDate || "",
      status: normalizeStoredStatus(entry.status),
    });
  });
  state.sales.forEach((entry) => {
    const foundCase = state.cases.find((c) => c.id === entry.caseId);
    rows.push({
      data_type: "sale",
      case_name: foundCase?.caseName || "",
      invoice_amount: entry.invoiceAmount ?? "",
      paid_amount: entry.paidAmount ?? "",
      paid_date: entry.paidDate || "",
      is_unpaid: entry.isUnpaid ? "true" : "false",
    });
  });
  state.expenses.forEach((entry) => {
    const foundCase = state.cases.find((c) => c.id === entry.caseId);
    rows.push({
      data_type: "expense",
      case_name: foundCase?.caseName || "",
      date: entry.date || "",
      content: entry.content || "",
      amount: entry.amount ?? "",
    });
  });
  state.fixedExpenses.forEach((entry) => {
    rows.push({
      data_type: "fixed_expense",
      content: entry.content,
      amount: entry.amount ?? "",
      day_of_month: entry.dayOfMonth ?? "",
      start_date: entry.startDate || "",
      active: entry.active ? "true" : "false",
    });
  });

  downloadCsvFile("all_data.csv", headers, rows);
}

async function handleCsvImportSubmit(event) {
  event.preventDefault();
  if (!currentUser || !csvImportForm) return;
  const importType = csvImportForm.elements.csvImportType.value;
  const file = csvImportForm.elements.csvImportFile.files?.[0];
  if (!file) return;
  if (!window.confirm(`「${file.name}」を${importTypeToLabel(importType)}として取り込みます。よろしいですか？`)) return;

  await withLoading(async () => {
    const text = await readCsvFileTextWithEncodingFallback(file);
    const result = await importCsvByType(importType, text);
    await loadAllData();
    renderAfterDataChanged();
    showAppMessage(`CSV取込完了: 登録件数 ${result.insertedCount}件 / スキップ件数 ${result.skippedCount}件 / エラー件数 ${result.errorCount}件`, result.errorCount > 0);
    csvImportForm.reset();
  }, "CSV取込に失敗しました。");
}

function handleExportExcel() {
  exportExcel();
}

function exportExcel() {
  if (!window.XLSX) {
    showAppMessage("Excel出力ライブラリの読み込みに失敗しました", true);
    return;
  }

  try {
    const workbook = XLSX.utils.book_new();
    const caseHeaders = ["customer_name", "case_name", "estimate_amount", "received_date", "due_date", "status"];
    const saleHeaders = ["case_name", "invoice_amount", "paid_amount", "paid_date", "is_unpaid"];
    const expenseHeaders = ["case_name", "date", "content", "amount"];
    const fixedExpenseHeaders = ["content", "amount", "day_of_month", "start_date", "active"];
    const caseRows = state.cases.map((entry) => ({
      customer_name: entry.customerName,
      case_name: entry.caseName,
      estimate_amount: entry.estimateAmount ?? "",
      received_date: entry.receivedDate || "",
      due_date: entry.dueDate || "",
      status: normalizeStoredStatus(entry.status),
    }));
    const saleRows = state.sales.map((entry) => {
      const foundCase = state.cases.find((c) => c.id === entry.caseId);
      return {
        case_name: foundCase?.caseName || "",
        invoice_amount: entry.invoiceAmount ?? "",
        paid_amount: entry.paidAmount ?? "",
        paid_date: entry.paidDate || "",
        is_unpaid: entry.isUnpaid ? "true" : "false",
      };
    });
    const expenseRows = state.expenses.map((entry) => {
      const foundCase = state.cases.find((c) => c.id === entry.caseId);
      return {
        case_name: foundCase?.caseName || "",
        date: entry.date || "",
        content: entry.content || "",
        amount: entry.amount ?? "",
      };
    });
    const fixedExpenseRows = state.fixedExpenses.map((entry) => ({
      content: entry.content,
      amount: entry.amount ?? "",
      day_of_month: entry.dayOfMonth ?? "",
      start_date: entry.startDate || "",
      active: entry.active ? "true" : "false",
    }));

    XLSX.utils.book_append_sheet(workbook, createExcelSheet(caseRows, caseHeaders), "案件");
    XLSX.utils.book_append_sheet(workbook, createExcelSheet(saleRows, saleHeaders), "売上");
    XLSX.utils.book_append_sheet(workbook, createExcelSheet(expenseRows, expenseHeaders), "経費");
    XLSX.utils.book_append_sheet(workbook, createExcelSheet(fixedExpenseRows, fixedExpenseHeaders), "固定費");
    XLSX.writeFile(workbook, "gyosei-app-export.xlsx");
    showAppMessage("Excelファイルを出力しました。", false);
  } catch (error) {
    console.error("Excel出力に失敗しました。", error);
    showAppMessage("Excel出力に失敗しました。", true);
  }
}

function createExcelSheet(rows, headers) {
  const sheet = XLSX.utils.aoa_to_sheet([headers]);
  if (rows.length > 0) {
    XLSX.utils.sheet_add_json(sheet, rows, {
      header: headers,
      origin: "A2",
      skipHeader: true,
    });
  }
  return sheet;
}

async function handleExcelImportSubmit(event) {
  event.preventDefault();
  if (!currentUser || !excelImportForm) return;
  if (!window.XLSX) {
    showAppMessage("Excel取込ライブラリ（SheetJS）の読み込みに失敗しました。", true);
    return;
  }
  const file = excelImportForm.elements.excelImportFile.files?.[0];
  if (!file) return;
  if (!window.confirm(`「${file.name}」をExcelとして取り込みます。よろしいですか？`)) return;

  await withLoading(async () => {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array", cellDates: false });
    const result = await importWorkbookBySheet(workbook);
    await loadAllData();
    renderAfterDataChanged();
    showAppMessage(`Excel取込完了: 登録件数 ${result.insertedCount}件 / スキップ件数 ${result.skippedCount}件 / エラー件数 ${result.errorCount}件`, result.errorCount > 0);
    excelImportForm.reset();
  }, "Excel取込に失敗しました。");
}

function renderAfterDataChanged() {
  renderCaseOptions();
  renderCases();
  renderSales();
  renderExpenses();
  renderFixedExpenses();
  renderDashboard();
}

function activateTab(tabKey) {
  tabs.forEach((btn) => btn.classList.toggle("active", btn.dataset.tab === tabKey));
  Object.entries(panels).forEach(([key, panel]) => panel.classList.toggle("active", key === tabKey));
}

function renderDashboard() {
  state.selectedMonth = normalizeMonthKey(state.selectedMonth) || toMonthKey(new Date());
  state.selectedYear = normalizeYear(state.selectedYear) || new Date().getFullYear();
  state.selectedAggregation = state.selectedAggregation === "year" ? "year" : "month";

  aggregationRadios.forEach((radio) => {
    radio.checked = radio.value === state.selectedAggregation;
  });
  if (monthFilter) monthFilter.hidden = state.selectedAggregation !== "month";
  if (yearFilter) yearFilter.hidden = state.selectedAggregation !== "year";
  if (targetMonthInput) targetMonthInput.value = state.selectedMonth;
  if (targetYearInput) targetYearInput.value = String(state.selectedYear);

  const salesByMonth = aggregateByMonth(state.sales, (s) => s.paidDate || s.createdAt, (s) => s.invoiceAmount);
  const expenseByMonth = aggregateByMonth(state.expenses, (e) => e.date, (e) => e.amount);
  const isYearMode = state.selectedAggregation === "year";

  const sales = isYearMode ? sumYearFromMonthlyMap(salesByMonth, state.selectedYear) : salesByMonth[state.selectedMonth] || 0;
  const expenses = isYearMode ? sumYearFromMonthlyMap(expenseByMonth, state.selectedYear) : expenseByMonth[state.selectedMonth] || 0;
  const profit = sales - expenses;

  summaryGrid.innerHTML = "";
  const labelPrefix = isYearMode ? "年別" : "月別";
  const targetLabel = isYearMode ? `${state.selectedYear}年` : monthLabel(state.selectedMonth);
  [
    { label: isYearMode ? "対象年" : "対象月", value: targetLabel, cls: "" },
    { label: `${labelPrefix}売上合計`, value: formatCurrency(sales), cls: "" },
    { label: `${labelPrefix}経費合計`, value: formatCurrency(expenses), cls: "" },
    { label: "利益（売上−経費）", value: formatCurrency(profit), cls: `profit ${profit < 0 ? "loss" : ""}`.trim() },
  ].forEach((card) => {
    const div = document.createElement("div");
    div.className = `summary-card ${card.cls}`.trim();
    div.innerHTML = `<p class="label">${card.label}</p><p class="value">${card.value}</p>`;
    summaryGrid.appendChild(div);
  });
  renderDeadlineAlertCard();
  renderStatusSummaryCard();

  renderUnpaidAlert({
    mode: state.selectedAggregation,
    monthKey: state.selectedMonth,
    year: state.selectedYear,
  });
  renderBillingLeakAlert();
  renderYearlyBreakdown(salesByMonth, expenseByMonth);
  renderCaseProfitList({
    mode: state.selectedAggregation,
    monthKey: state.selectedMonth,
    year: state.selectedYear,
  });
  renderUnpaidList();
}

function renderDeadlineAlertCard() {
  if (!deadlineAlertCard || !deadlineAlertSummary || !deadlineAlertBody || !deadlineAlertEmpty || !deadlineAlertListWrap) return;
  const targets = getDeadlineAlertTargets();
  const counts = { overdue: 0, within7: 0, within30: 0 };
  targets.forEach((item) => {
    counts[item.deadlineStatus] += 1;
  });

  deadlineAlertCard.classList.toggle("has-alert", targets.length > 0);
  deadlineAlertSummary.innerHTML = "";

  [
    { key: "overdue", label: "期限切れ", count: counts.overdue },
    { key: "within7", label: "7日以内", count: counts.within7 },
    { key: "within30", label: "30日以内", count: counts.within30 },
  ].forEach((entry) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = `deadline-alert-chip ${entry.key}`;
    button.dataset.deadlineFilter = entry.key;
    button.textContent = `${entry.label}: ${entry.count}件`;
    button.disabled = entry.count === 0;
    deadlineAlertSummary.appendChild(button);
  });

  deadlineAlertBody.innerHTML = "";
  deadlineAlertEmpty.hidden = targets.length > 0;
  deadlineAlertListWrap.hidden = targets.length === 0;
  if (!targets.length) return;

  targets
    .slice()
    .sort((a, b) => {
      if (a.remainingDays !== b.remainingDays) return a.remainingDays - b.remainingDays;
      return sortCases(a, b);
    })
    .forEach((entry) => {
      const tr = document.createElement("tr");
      tr.className =
        entry.deadlineStatus === "overdue" ? "deadline-overdue" :
        entry.deadlineStatus === "within7" ? "deadline-within7" :
        "deadline-within30";
      tr.innerHTML = `
        <td>${escapeHtml(entry.customerName)}</td>
        <td>${escapeHtml(entry.caseName)}</td>
        <td>${formatDate(entry.dueDate)}</td>
        <td>${escapeHtml(entry.status)}</td>
        <td>${formatRemainingDays(entry.remainingDays)}</td>
        <td><button type="button" class="secondary-btn" data-case-id="${entry.id}">編集</button></td>
      `;
      deadlineAlertBody.appendChild(tr);
    });
}

function getDeadlineAlertTargets() {
  return state.cases
    .map((entry) => {
      const info = getCaseDeadlineInfo(entry);
      if (!info) return null;
      return { ...entry, ...info };
    })
    .filter(Boolean);
}

function getCaseDeadlineInfo(entry) {
  if (!entry?.dueDate) return null;
  if (getStatusCategory(entry.status) === "完了") return null;
  const dueTimestamp = toDateOnlyTimestamp(entry.dueDate);
  const todayTimestamp = getTodayTimestamp();
  if (!Number.isFinite(dueTimestamp) || !Number.isFinite(todayTimestamp)) return null;
  const remainingDays = Math.floor((dueTimestamp - todayTimestamp) / 86400000);
  if (remainingDays < 0) return { deadlineStatus: "overdue", remainingDays };
  if (remainingDays <= 7) return { deadlineStatus: "within7", remainingDays };
  if (remainingDays <= 30) return { deadlineStatus: "within30", remainingDays };
  return null;
}

function renderStatusSummaryCard() {
  if (!statusSummaryList || !statusSummaryTotal) return;
  const counts = countCasesByStatus();
  statusSummaryList.innerHTML = "";
  if (statusFilterClearBtn) statusFilterClearBtn.disabled = state.caseStatusFilter === "all";

  STATUS_FILTER_KEYS.forEach((statusKey) => {
    const button = document.createElement("button");
    const count = counts[statusKey] || 0;
    button.type = "button";
    button.className = `status-filter-btn status-${toStatusClassKey(statusKey)}`.trim();
    button.dataset.statusFilter = statusKey;
    button.setAttribute("aria-pressed", String(state.caseStatusFilter === statusKey));
    button.innerHTML = `<span>${statusKey}</span><strong>${count}件</strong>`;
    statusSummaryList.appendChild(button);
  });

  statusSummaryTotal.textContent = `全案件数：${state.cases.length}件`;
}

function renderUnpaidAlert(filter = {}) {
  const unpaidSales = getUnpaidSales();
  const count = unpaidSales.length;
  const totalUnpaidAmount = unpaidSales.reduce((sum, sale) => sum + getUnpaidAmount(sale), 0);
  const filteredUnpaid = unpaidSales.filter((sale) => isWithinFilterDate(sale.paidDate || sale.createdAt, filter));
  const periodCount = filteredUnpaid.length;
  const periodAmount = filteredUnpaid.reduce((sum, sale) => sum + getUnpaidAmount(sale), 0);

  if (unpaidAlertCard) unpaidAlertCard.classList.toggle("has-unpaid", count > 0);
  if (unpaidAlertSummary) {
    unpaidAlertSummary.textContent = count > 0
      ? `未入金 ${count}件 / 未入金合計 ${formatCurrency(totalUnpaidAmount)}`
      : "未入金はありません。";
  }
  if (unpaidAlertPeriod) {
    unpaidAlertPeriod.textContent = filter.mode === "year"
      ? `対象年(${filter.year}年): ${periodCount}件 / ${formatCurrency(periodAmount)}`
      : `対象月(${monthLabel(filter.monthKey || state.selectedMonth)}): ${periodCount}件 / ${formatCurrency(periodAmount)}`;
  }
  if (unpaidAlertEmpty) unpaidAlertEmpty.hidden = count > 0;
  if (unpaidAlertListWrap) unpaidAlertListWrap.hidden = count === 0;
  if (!unpaidAlertBody) return;
  unpaidAlertBody.innerHTML = "";

  unpaidSales
    .slice()
    .sort((a, b) => (b.updatedAt ?? b.createdAt) - (a.updatedAt ?? a.createdAt))
    .forEach((sale) => {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${escapeHtml(resolveCaseName(sale.caseId))}</td>
        <td>${formatCurrency(sale.invoiceAmount)}</td>
        <td>${formatCurrency(sale.paidAmount)}</td>
        <td>${formatCurrency(getUnpaidAmount(sale))}</td>
        <td>${formatDate(sale.paidDate || toDateString(sale.createdAt))}</td>
      `;
      unpaidAlertBody.appendChild(tr);
    });
}

function renderUnpaidList() {
  const unpaidSales = getUnpaidSales().slice().sort((a, b) => (b.updatedAt ?? b.createdAt) - (a.updatedAt ?? a.createdAt));
  if (!unpaidListBody || !unpaidListEmpty || !unpaidListWrap) return;
  unpaidListBody.innerHTML = "";
  unpaidListEmpty.hidden = unpaidSales.length > 0;
  unpaidListWrap.hidden = unpaidSales.length === 0;

  unpaidSales.forEach((sale) => {
    const linkedCase = state.cases.find((entry) => entry.id === sale.caseId);
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(linkedCase?.customerName || "（削除済み顧客）")}</td>
      <td>${escapeHtml(linkedCase?.caseName || "（削除済み案件）")}</td>
      <td>${formatCurrency(sale.invoiceAmount)}</td>
      <td>${formatCurrency(sale.paidAmount)}</td>
      <td>${formatCurrency(getUnpaidAmount(sale))}</td>
      <td>${formatDate(sale.paidDate)}</td>
      <td><button type="button" class="secondary-btn edit-sale-btn" data-sale-id="${sale.id}">編集</button></td>
    `;
    unpaidListBody.appendChild(tr);
  });
}

function renderBillingLeakAlert() {
  const billingLeakCases = getBillingLeakCandidates()
    .slice()
    .sort(sortCases);
  const count = billingLeakCases.length;

  if (billingLeakAlertCard) billingLeakAlertCard.classList.toggle("has-leak", count > 0);
  if (billingLeakAlertSummary) {
    billingLeakAlertSummary.textContent = count > 0
      ? `請求漏れ候補 ${count}件`
      : "請求漏れ候補はありません。";
  }
  if (billingLeakAlertEmpty) billingLeakAlertEmpty.hidden = count > 0;
  if (billingLeakAlertListWrap) billingLeakAlertListWrap.hidden = count === 0;
  if (!billingLeakAlertBody) return;
  billingLeakAlertBody.innerHTML = "";

  billingLeakCases.forEach((entry) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(entry.customerName)}</td>
      <td>${escapeHtml(entry.caseName)}</td>
      <td>${formatCurrency(entry.estimateAmount)}</td>
      <td>${escapeHtml(entry.status || "完了")}</td>
      <td><button type="button" class="secondary-btn register-sale-btn" data-case-id="${entry.id}">売上登録</button></td>
    `;
    billingLeakAlertBody.appendChild(tr);
  });
}

function getBillingLeakCandidates() {
  return state.cases.filter((entry) => {
    if (getStatusCategory(entry.status) !== "完了") return false;
    return !state.sales.some((sale) => sale.caseId === entry.id);
  });
}

function getUnpaidSales() {
  return state.sales.filter((sale) => {
    const invoiceAmount = sale.invoiceAmount ?? 0;
    const paidAmount = sale.paidAmount ?? 0;
    return Boolean(sale.isUnpaid) || paidAmount < invoiceAmount || !sale.paidDate;
  });
}

function getUnpaidAmount(sale) {
  return Math.max((sale.invoiceAmount ?? 0) - (sale.paidAmount ?? 0), 0);
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
  const statusFilteredCases = state.caseStatusFilter === "all"
    ? sorted
    : sorted.filter((entry) => getStatusCategory(entry.status) === state.caseStatusFilter);
  const filteredCases = statusFilteredCases.filter((entry) => {
    if (state.caseSearchQuery && !matchesCaseSearch(entry, state.caseSearchQuery)) return false;
    if (!matchesDeadlineFilter(entry, state.caseDeadlineFilter)) return false;
    return true;
  });
  const profitsByCaseId = buildCaseProfitMap();

  filteredCases.forEach((entry) => {
    const node = caseItemTemplate.content.cloneNode(true);
    const item = node.querySelector(".item");
    const title = node.querySelector(".title");
    const meta = node.querySelector(".meta");
    const profitMeta = node.querySelector(".profit-meta");
    const totals = profitsByCaseId[entry.id] || { sales: 0, expenses: 0, profit: 0 };

    item.dataset.id = entry.id;
    title.textContent = `${entry.customerName}｜${entry.caseName}`;
    meta.textContent = `見積: ${formatCurrency(entry.estimateAmount)} / ステータス: ${entry.status} / 受付日: ${formatDate(entry.receivedDate)} / 期限日: ${formatDate(entry.dueDate)}`;
    profitMeta.textContent = `売上合計: ${formatCurrency(totals.sales)} / 経費合計: ${formatCurrency(totals.expenses)} / 利益: ${formatCurrency(totals.profit)}`;
    profitMeta.classList.toggle("loss-text", totals.profit < 0);
    caseList.appendChild(node);
  });

  if (caseFilterCount) {
    caseFilterCount.textContent = `表示中 ${filteredCases.length}件 / 全${sorted.length}件`;
  }
  caseEmpty.hidden = filteredCases.length > 0;
  if (!filteredCases.length && (state.caseStatusFilter !== "all" || state.caseSearchQuery || state.caseDeadlineFilter !== "all")) {
    caseEmpty.textContent = "条件に一致する案件はありません。";
  } else {
    caseEmpty.textContent = "案件はまだ登録されていません。";
  }
}

function renderCaseProfitList(filter = {}) {
  if (!caseProfitList || !caseProfitEmpty) return;
  const profitsByCaseId = buildCaseProfitMap(filter);
  const sortedCases = state.cases
    .filter((entry) => {
      if (!filter?.mode) return true;
      const totals = profitsByCaseId[entry.id];
      return Boolean(totals) && (totals.sales !== 0 || totals.expenses !== 0);
    })
    .sort(sortCases);
  caseProfitList.innerHTML = "";

  sortedCases.forEach((entry) => {
    const totals = profitsByCaseId[entry.id] || { sales: 0, expenses: 0, profit: 0 };
    const li = document.createElement("li");
    li.className = "item";
    li.innerHTML = `
      <div>
        <p class="title">${escapeHtml(entry.customerName)}｜${escapeHtml(entry.caseName)}</p>
        <p class="meta">売上合計: ${formatCurrency(totals.sales)} / 経費合計: ${formatCurrency(totals.expenses)} / <span class="${totals.profit < 0 ? "loss-text" : ""}">利益: ${formatCurrency(totals.profit)}</span></p>
      </div>
    `;
    caseProfitList.appendChild(li);
  });

  caseProfitEmpty.hidden = sortedCases.length > 0;
  if (!sortedCases.length) {
    caseProfitEmpty.textContent =
      filter?.mode === "year" ? "選択した対象年の案件データはありません。" :
      filter?.mode === "month" ? "選択した対象月の案件データはありません。" :
      "案件データがありません。";
  } else {
    caseProfitEmpty.textContent = "案件データがありません。";
  }
}

function buildCaseProfitMap(filter = {}) {
  const map = {};
  state.cases.forEach((entry) => {
    map[entry.id] = { sales: 0, expenses: 0, profit: 0 };
  });

  state.sales.forEach((sale) => {
    if (!sale.caseId || !map[sale.caseId]) return;
    if (!isWithinFilterDate(sale.paidDate || sale.createdAt, filter)) return;
    map[sale.caseId].sales += sale.invoiceAmount || 0;
  });

  state.expenses.forEach((expense) => {
    if (!expense.caseId || !map[expense.caseId]) return;
    if (!isWithinFilterDate(expense.date, filter)) return;
    map[expense.caseId].expenses += expense.amount || 0;
  });

  Object.values(map).forEach((item) => {
    item.profit = item.sales - item.expenses;
  });

  return map;
}

function renderYearlyBreakdown(salesByMonth, expenseByMonth) {
  const isYearMode = state.selectedAggregation === "year";
  if (yearlyBreakdown) yearlyBreakdown.hidden = !isYearMode;
  if (!yearlyBreakdownBody) return;
  yearlyBreakdownBody.innerHTML = "";
  if (!isYearMode) return;

  for (let month = 1; month <= 12; month += 1) {
    const monthKey = `${state.selectedYear}-${String(month).padStart(2, "0")}`;
    const sales = salesByMonth[monthKey] || 0;
    const expenses = expenseByMonth[monthKey] || 0;
    const profit = sales - expenses;
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${month}月</td>
      <td>${formatCurrency(sales)}</td>
      <td>${formatCurrency(expenses)}</td>
      <td class="${profit < 0 ? "loss-text" : ""}">${formatCurrency(profit)}</td>
    `;
    yearlyBreakdownBody.appendChild(tr);
  }
}

function renderSales() {
  salesList.innerHTML = "";
  const sorted = state.sales.slice().sort((a, b) => (b.updatedAt ?? b.createdAt) - (a.updatedAt ?? a.createdAt));
  const filteredSales = sorted.filter((sale) => {
    if (!state.salesSearchQuery) return true;
    return matchesSalesSearch(sale, state.salesSearchQuery);
  });

  filteredSales.forEach((sale) => {
    const node = saleItemTemplate.content.cloneNode(true);
    const item = node.querySelector(".item");
    const title = node.querySelector(".title");
    const meta = node.querySelector(".meta");

    item.dataset.id = sale.id;
    title.textContent = `${resolveCaseName(sale.caseId)}｜請求: ${formatCurrency(sale.invoiceAmount)}`;
    meta.textContent = `入金: ${formatCurrency(sale.paidAmount)} / 入金日: ${formatDate(sale.paidDate)} / 未入金: ${sale.isUnpaid ? "はい" : "いいえ"}`;
    salesList.appendChild(node);
  });

  if (salesFilterCount) {
    salesFilterCount.textContent = `表示中 ${filteredSales.length}件 / 全${sorted.length}件`;
  }
  salesEmpty.hidden = filteredSales.length > 0;
  salesEmpty.textContent = filteredSales.length || !state.salesSearchQuery
    ? "売上データはまだありません。"
    : "条件に一致する売上データはありません。";
}

function renderExpenses() {
  expensesList.innerHTML = "";
  const sorted = state.expenses.slice().sort((a, b) => toSortTimestamp(b.date) - toSortTimestamp(a.date));
  const filteredExpenses = sorted.filter((expense) => {
    if (!state.expensesSearchQuery) return true;
    return matchesExpensesSearch(expense, state.expensesSearchQuery);
  });

  filteredExpenses.forEach((expense) => {
    const node = expenseItemTemplate.content.cloneNode(true);
    const item = node.querySelector(".item");
    const title = node.querySelector(".title");
    const meta = node.querySelector(".meta");

    item.dataset.id = expense.id;
    title.textContent = `${formatDate(expense.date)}｜${expense.content}`;
    meta.textContent = `金額: ${formatCurrency(expense.amount)} / 紐付け: ${expense.caseId ? resolveCaseName(expense.caseId) : "なし"}`;
    expensesList.appendChild(node);
  });

  if (expensesFilterCount) {
    expensesFilterCount.textContent = `表示中 ${filteredExpenses.length}件 / 全${sorted.length}件`;
  }
  expensesEmpty.hidden = filteredExpenses.length > 0;
  expensesEmpty.textContent = filteredExpenses.length || !state.expensesSearchQuery
    ? "経費データはまだありません。"
    : "条件に一致する経費データはありません。";
}

function renderFixedExpenses() {
  fixedExpensesList.innerHTML = "";
  const sorted = state.fixedExpenses.slice().sort((a, b) => (b.updatedAt ?? b.createdAt) - (a.updatedAt ?? a.createdAt));

  sorted.forEach((entry) => {
    const node = fixedExpenseItemTemplate.content.cloneNode(true);
    const item = node.querySelector(".item");
    const title = node.querySelector(".title");
    const meta = node.querySelector(".meta");
    const toggleBtn = node.querySelector(".toggle-btn");

    item.dataset.id = entry.id;
    title.textContent = `${entry.content}｜${formatCurrency(entry.amount)}`;
    meta.textContent = `毎月${entry.dayOfMonth}日 / 開始日: ${formatDate(entry.startDate)} / 状態: ${entry.active ? "有効" : "無効"}`;
    toggleBtn.textContent = entry.active ? "無効にする" : "有効にする";
    fixedExpensesList.appendChild(node);
  });

  fixedExpensesEmpty.hidden = sorted.length > 0;
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

function resetFixedExpenseForm() {
  resetEditMode("fixedExpense");
  fixedExpenseForm.reset();
  fixedExpenseForm.elements.fixedExpenseActive.checked = true;
  fixedExpenseSubmitBtn.textContent = "固定費を登録";
}

function resetEditMode(target) {
  if (target === "case") editState.caseId = null;
  if (target === "sale") editState.saleId = null;
  if (target === "expense") editState.expenseId = null;
  if (target === "fixedExpense") editState.fixedExpenseId = null;
}

function sortCases(a, b) {
  const dueA = toSortTimestamp(a.dueDate);
  const dueB = toSortTimestamp(b.dueDate);
  if (dueA !== dueB) return dueA - dueB;

  const statusA = STATUS_FILTER_KEYS.indexOf(getStatusCategory(a.status));
  const statusB = STATUS_FILTER_KEYS.indexOf(getStatusCategory(b.status));
  if (statusA !== statusB) return statusA - statusB;
  return a.createdAt - b.createdAt;
}

function countCasesByStatus() {
  const counts = { 未着手: 0, 進行中: 0, 完了: 0, その他: 0 };
  state.cases.forEach((entry) => {
    const category = getStatusCategory(entry.status);
    counts[category] += 1;
  });
  return counts;
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

function normalizeMonthKey(value) {
  if (typeof value !== "string") return "";
  if (!/^\d{4}-\d{2}$/.test(value)) return "";
  const month = Number(value.slice(5, 7));
  return month >= 1 && month <= 12 ? value : "";
}

function isSameMonthKey(dateSource, monthKey) {
  return toMonthKey(dateSource) === monthKey;
}

function toYearNumber(dateSource) {
  const date = dateSource instanceof Date ? dateSource : new Date(dateSource);
  if (Number.isNaN(date.getTime())) return null;
  return date.getFullYear();
}

function normalizeYear(value) {
  const year = Number(value);
  if (!Number.isInteger(year) || year < 1900 || year > 9999) return null;
  return year;
}

function isWithinFilterDate(dateSource, filter = {}) {
  if (!filter?.mode) return true;
  if (filter.mode === "month") return isSameMonthKey(dateSource, filter.monthKey);
  if (filter.mode === "year") return toYearNumber(dateSource) === normalizeYear(filter.year);
  return true;
}

function sumYearFromMonthlyMap(monthlyMap, year) {
  const normalizedYear = normalizeYear(year);
  if (!normalizedYear) return 0;
  let total = 0;
  for (let month = 1; month <= 12; month += 1) {
    const key = `${normalizedYear}-${String(month).padStart(2, "0")}`;
    total += monthlyMap[key] || 0;
  }
  return total;
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

function toDateString(dateSource) {
  const date = dateSource instanceof Date ? dateSource : new Date(dateSource);
  if (Number.isNaN(date.getTime())) return "";
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
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

function normalizeStoredStatus(status) {
  if (status === null || status === undefined) return "未着手";
  const normalized = String(status).trim();
  return normalized || "未着手";
}

function getStatusCategory(status) {
  const normalized = normalizeStoredStatus(status);
  return STATUS_ORDER.includes(normalized) ? normalized : "その他";
}

function matchesCaseSearch(entry, query) {
  const needle = query.trim().toLowerCase();
  if (!needle) return true;
  const haystack = [
    entry.customerName,
    entry.caseName,
    entry.status,
  ]
    .map((value) => String(value || "").toLowerCase())
    .join(" ");
  return haystack.includes(needle);
}

function matchesDeadlineFilter(entry, filterKey) {
  if (filterKey === "all") return true;
  const info = getCaseDeadlineInfo(entry);
  if (!info) return false;
  if (filterKey === "overdue") return info.deadlineStatus === "overdue";
  if (filterKey === "within7") return info.deadlineStatus === "within7";
  if (filterKey === "within30") return info.deadlineStatus === "within30" || info.deadlineStatus === "within7";
  return true;
}

function getTodayTimestamp() {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  return today.getTime();
}

function toDateOnlyTimestamp(value) {
  if (!value) return Number.NaN;
  if (typeof value === "string" && /^\d{4}-\d{2}-\d{2}$/.test(value)) {
    const [yearText, monthText, dayText] = value.split("-");
    const year = Number(yearText);
    const monthIndex = Number(monthText) - 1;
    const day = Number(dayText);
    return new Date(year, monthIndex, day).getTime();
  }
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return Number.NaN;
  date.setHours(0, 0, 0, 0);
  return date.getTime();
}

function formatRemainingDays(remainingDays) {
  if (!Number.isFinite(remainingDays)) return "-";
  if (remainingDays < 0) return `${Math.abs(remainingDays)}日超過`;
  if (remainingDays === 0) return "今日まで";
  return `残り${remainingDays}日`;
}

function getSalesPaymentLabel(sale) {
  if (sale.isUnpaid) return "未入金";
  const invoiceAmount = sale.invoiceAmount ?? 0;
  const paidAmount = sale.paidAmount ?? 0;
  if (paidAmount >= invoiceAmount && invoiceAmount > 0) return "入金済み";
  if (paidAmount > 0) return "一部入金";
  return "未入金";
}

function matchesSalesSearch(sale, query) {
  const needle = query.trim().toLowerCase();
  if (!needle) return true;
  const relatedCase = state.cases.find((entry) => entry.id === sale.caseId);
  const haystack = [
    relatedCase?.customerName,
    relatedCase?.caseName,
    sale.invoiceAmount,
    getSalesPaymentLabel(sale),
  ]
    .map((value) => String(value ?? "").toLowerCase())
    .join(" ");
  return haystack.includes(needle);
}

function matchesExpensesSearch(expense, query) {
  const needle = query.trim().toLowerCase();
  if (!needle) return true;
  const relatedCase = state.cases.find((entry) => entry.id === expense.caseId);
  const haystack = [
    expense.content,
    relatedCase?.caseName,
    expense.amount,
  ]
    .map((value) => String(value ?? "").toLowerCase())
    .join(" ");
  return haystack.includes(needle);
}

function toStatusClassKey(status) {
  if (status === "未着手") return "not-started";
  if (status === "進行中") return "in-progress";
  if (status === "完了") return "completed";
  return "other";
}

function normalizeDayOfMonth(raw) {
  const value = Number(raw);
  if (!Number.isInteger(value) || value < 1 || value > 31) return null;
  return value;
}

function importTypeToLabel(type) {
  if (type === "cases") return "案件CSV";
  if (type === "sales") return "売上CSV";
  if (type === "expenses") return "経費CSV";
  if (type === "fixed_expenses") return "固定費CSV";
  return "CSV";
}

function downloadCsvFile(filename, headers, rows) {
  const csv = buildCsvString(headers, rows);
  const bom = "\uFEFF";
  const blob = new Blob([bom + csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

function buildCsvString(headers, rows) {
  const headerLine = headers.map((h) => escapeCsvField(h)).join(",");
  const lines = rows.map((row) => headers.map((header) => escapeCsvField(row[header] ?? "")).join(","));
  return [headerLine, ...lines].join("\r\n");
}

function escapeCsvField(value) {
  const normalized = String(value ?? "");
  const escaped = normalized.replaceAll('"', '""');
  return `"${escaped}"`;
}

function parseCsvText(text) {
  const rows = [];
  let row = [];
  let field = "";
  let index = 0;
  let insideQuotes = false;
  const normalizedText = String(text || "").replace(/^\uFEFF/, "");

  while (index < normalizedText.length) {
    const ch = normalizedText[index];
    const next = normalizedText[index + 1];

    if (ch === '"') {
      if (insideQuotes && next === '"') {
        field += '"';
        index += 2;
        continue;
      }
      insideQuotes = !insideQuotes;
      index += 1;
      continue;
    }

    if (!insideQuotes && ch === ",") {
      row.push(field);
      field = "";
      index += 1;
      continue;
    }

    if (!insideQuotes && (ch === "\n" || ch === "\r")) {
      if (ch === "\r" && next === "\n") index += 1;
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
      index += 1;
      continue;
    }

    field += ch;
    index += 1;
  }

  row.push(field);
  rows.push(row);
  return rows;
}

async function readCsvFileTextWithEncodingFallback(file) {
  const buffer = await file.arrayBuffer();
  let utf8Text = "";
  try {
    utf8Text = new TextDecoder("utf-8", { fatal: true }).decode(buffer);
  } catch (_error) {
    utf8Text = "";
  }

  if (utf8Text && !isLikelyMojibake(utf8Text)) return utf8Text;

  try {
    const shiftJisText = new TextDecoder("shift_jis").decode(buffer);
    if (shiftJisText) return shiftJisText;
  } catch (_error) {
    // shift_jis decode failed: fall through to utf-8 tolerant decode
  }

  return utf8Text || new TextDecoder("utf-8").decode(buffer);
}

function isLikelyMojibake(text) {
  if (!text) return false;
  if (text.includes("\uFFFD")) return true;
  const suspiciousPattern = /(?:Ã.|Â.|ã.|�)/g;
  const suspiciousMatches = text.match(suspiciousPattern);
  return Boolean(suspiciousMatches && suspiciousMatches.length >= 3);
}

function sanitizeHeader(value) {
  return String(value || "").trim().toLowerCase();
}

function parseSheetToObjects(sheet) {
  if (!window.XLSX || !sheet) return { headers: [], rows: [] };
  const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: true });
  const nonEmptyRows = rawRows.filter((row) => Array.isArray(row) && !row.every((cell) => String(cell ?? "").trim() === ""));
  if (!nonEmptyRows.length) return { headers: [], rows: [] };
  const headers = nonEmptyRows[0].map((h) => sanitizeHeader(h));
  const rows = nonEmptyRows.slice(1).map((cells) => {
    const obj = {};
    headers.forEach((header, idx) => {
      obj[header] = cells[idx] ?? "";
    });
    return obj;
  });
  return { headers, rows };
}

function parseCsvToObjects(csvText) {
  const rawRows = parseCsvText(csvText);
  const nonEmptyRows = rawRows.filter((row) => !row.every((cell) => String(cell || "").trim() === ""));
  if (!nonEmptyRows.length) return { headers: [], rows: [] };
  const headers = nonEmptyRows[0].map((h) => sanitizeHeader(h));
  const rows = nonEmptyRows.slice(1).map((cells) => {
    const obj = {};
    headers.forEach((header, idx) => {
      obj[header] = String(cells[idx] ?? "").trim();
    });
    return obj;
  });
  return { headers, rows };
}

function validateRequiredHeaders(headers, requiredHeaders) {
  const missing = requiredHeaders.filter((header) => !headers.includes(header));
  if (missing.length) {
    throw new Error(`ヘッダー不足: ${missing.join(", ")}`);
  }
}

function parseFlexibleDate(value) {
  if (typeof value === "number" && Number.isFinite(value)) {
    const date = excelSerialToDateString(value);
    if (date) return date;
  }
  const trimmed = String(value || "").trim();
  if (!trimmed) return null;
  if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) return trimmed;
  if (/^\d{4}\/\d{2}\/\d{2}$/.test(trimmed)) return trimmed.replaceAll("/", "-");
  if (/^\d+(?:\.\d+)?$/.test(trimmed)) {
    const serial = Number(trimmed);
    const date = excelSerialToDateString(serial);
    if (date) return date;
  }
  return null;
}

function parseFlexibleAmount(value) {
  const normalized = String(value ?? "").trim().replaceAll(",", "");
  return normalizeAmount(normalized);
}

function parseBooleanLike(value) {
  const normalized = String(value || "").trim().toLowerCase();
  return ["1", "true", "yes", "y", "on"].includes(normalized);
}

function asTrimmedText(value) {
  return String(value ?? "").trim();
}

function excelSerialToDateString(serial) {
  if (!Number.isFinite(serial) || serial <= 0) return null;
  const baseUtcMs = Date.UTC(1899, 11, 30);
  const utcMs = baseUtcMs + Math.floor(serial) * 86400000;
  const date = new Date(utcMs);
  if (Number.isNaN(date.getTime())) return null;
  return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}-${String(date.getUTCDate()).padStart(2, "0")}`;
}

async function importWorkbookBySheet(workbook) {
  const total = { insertedCount: 0, skippedCount: 0, errorCount: 0 };
  if (!workbook || !Array.isArray(workbook.SheetNames)) return total;
  const mergeResult = (result) => {
    total.insertedCount += result.insertedCount || 0;
    total.skippedCount += result.skippedCount || 0;
    total.errorCount += result.errorCount || 0;
  };
  const getSheetRows = (sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    return parseSheetToObjects(sheet);
  };

  const caseData = getSheetRows("案件");
  if (caseData.rows.length) {
    mergeResult(await importRowsByType("cases", caseData));
    await loadCasesOnly();
  }
  const salesData = getSheetRows("売上");
  if (salesData.rows.length) {
    mergeResult(await importRowsByType("sales", salesData));
  }
  const expensesData = getSheetRows("経費");
  if (expensesData.rows.length) {
    mergeResult(await importRowsByType("expenses", expensesData));
  }
  const fixedExpensesData = getSheetRows("固定費");
  if (fixedExpensesData.rows.length) {
    mergeResult(await importRowsByType("fixed_expenses", fixedExpensesData));
  }
  return total;
}

async function loadCasesOnly() {
  if (!currentUser) return;
  const casesRes = await sbClient.from("cases").select("*").eq("user_id", currentUser.id);
  if (casesRes.error) throw casesRes.error;
  state.cases = (casesRes.data || []).map(mapCaseFromDb);
}

async function importCsvByType(importType, csvText) {
  const { headers, rows } = parseCsvToObjects(csvText);
  const normalizedRows = rows.map((row) => {
    const copied = {};
    Object.entries(row).forEach(([key, value]) => {
      copied[key] = String(value ?? "").trim();
    });
    return copied;
  });
  return importRowsByType(importType, { headers, rows: normalizedRows });
}

async function importRowsByType(importType, tableData) {
  const result = { insertedCount: 0, skippedCount: 0, errorCount: 0 };
  const headers = (tableData?.headers || []).map((h) => sanitizeHeader(h));
  const rows = (tableData?.rows || []).filter((row) => row && typeof row === "object");
  if (!rows.length) return result;

  if (importType === "cases") {
    validateRequiredHeaders(headers, ["customer_name", "case_name", "estimate_amount", "received_date", "due_date", "status"]);
    const payloads = [];
    rows.forEach((row) => {
      try {
        const customerName = asTrimmedText(row.customer_name);
        const caseName = asTrimmedText(row.case_name);
        if (!customerName || !caseName) {
          result.skippedCount += 1;
          return;
        }
        payloads.push({
          user_id: currentUser.id,
          customer_name: customerName,
          case_name: caseName,
          estimate_amount: parseFlexibleAmount(row.estimate_amount),
          received_date: parseFlexibleDate(row.received_date),
          due_date: parseFlexibleDate(row.due_date),
          status: normalizeStoredStatus(row.status),
        });
      } catch (_error) {
        result.errorCount += 1;
      }
    });
    if (payloads.length) {
      const { error } = await sbClient.from("cases").insert(payloads);
      if (error) throw error;
      result.insertedCount += payloads.length;
    }
    return result;
  }

  if (importType === "sales") {
    validateRequiredHeaders(headers, ["case_name", "invoice_amount", "paid_amount", "paid_date", "is_unpaid"]);
    const caseMap = new Map(state.cases.map((c) => [c.caseName, c.id]));
    const payloads = [];
    rows.forEach((row) => {
      try {
        const caseName = asTrimmedText(row.case_name);
        const caseId = caseMap.get(caseName);
        if (!caseId) {
          result.skippedCount += 1;
          return;
        }
        const invoiceAmount = parseFlexibleAmount(row.invoice_amount);
        if (invoiceAmount === null) {
          result.errorCount += 1;
          return;
        }
        const paidAmount = parseFlexibleAmount(row.paid_amount);
        payloads.push({
          user_id: currentUser.id,
          case_id: caseId,
          invoice_amount: invoiceAmount,
          paid_amount: paidAmount,
          paid_date: parseFlexibleDate(row.paid_date),
          is_unpaid: parseBooleanLike(row.is_unpaid) || (paidAmount ?? 0) < invoiceAmount,
        });
      } catch (_error) {
        result.errorCount += 1;
      }
    });
    if (payloads.length) {
      const { error } = await sbClient.from("sales").insert(payloads);
      if (error) throw error;
      result.insertedCount += payloads.length;
    }
    return result;
  }

  if (importType === "expenses") {
    validateRequiredHeaders(headers, ["case_name", "date", "content", "amount"]);
    const caseMap = new Map(state.cases.map((c) => [c.caseName, c.id]));
    const payloads = [];
    rows.forEach((row) => {
      try {
        const date = parseFlexibleDate(row.date);
        const content = asTrimmedText(row.content);
        const amount = parseFlexibleAmount(row.amount);
        if (!date || !content || amount === null) {
          result.errorCount += 1;
          return;
        }
        const caseName = asTrimmedText(row.case_name);
        const caseId = caseName ? (caseMap.get(caseName) || null) : null;
        payloads.push({
          user_id: currentUser.id,
          case_id: caseId,
          date,
          content,
          amount,
        });
      } catch (_error) {
        result.errorCount += 1;
      }
    });
    if (payloads.length) {
      const { error } = await sbClient.from("expenses").insert(payloads);
      if (error) throw error;
      result.insertedCount += payloads.length;
    }
    return result;
  }

  if (importType === "fixed_expenses") {
    validateRequiredHeaders(headers, ["content", "amount", "day_of_month", "start_date", "active"]);
    const existing = new Set(
      state.fixedExpenses.map((entry) => `${entry.content}__${entry.amount}__${entry.dayOfMonth}`)
    );
    const payloads = [];
    rows.forEach((row) => {
      try {
        const content = asTrimmedText(row.content);
        const amount = parseFlexibleAmount(row.amount);
        const dayOfMonth = normalizeDayOfMonth(row.day_of_month);
        const startDate = parseFlexibleDate(row.start_date);
        if (!content || amount === null || !dayOfMonth || !startDate) {
          result.errorCount += 1;
          return;
        }
        const dupKey = `${content}__${amount}__${dayOfMonth}`;
        if (existing.has(dupKey)) {
          result.skippedCount += 1;
          return;
        }
        existing.add(dupKey);
        payloads.push({
          user_id: currentUser.id,
          content,
          amount,
          day_of_month: dayOfMonth,
          start_date: startDate,
          active: parseBooleanLike(row.active),
        });
      } catch (_error) {
        result.errorCount += 1;
      }
    });
    if (payloads.length) {
      const { error } = await sbClient.from("fixed_expenses").insert(payloads);
      if (error) throw error;
      result.insertedCount += payloads.length;
    }
    return result;
  }

  throw new Error("未対応のCSV取込種別です。");
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
    status: normalizeStoredStatus(row.status),
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
    fixedExpenseId: row.fixed_expense_id || null,
    createdAt: Date.parse(row.created_at) || Date.now(),
    updatedAt: Date.parse(row.updated_at) || Date.now(),
  };
}

function mapFixedExpenseFromDb(row) {
  return {
    id: row.id,
    content: row.content || "",
    amount: normalizeAmount(row.amount) ?? 0,
    dayOfMonth: normalizeDayOfMonth(row.day_of_month) ?? 1,
    startDate: row.start_date || "",
    active: Boolean(row.active),
    createdAt: Date.parse(row.created_at) || Date.now(),
    updatedAt: Date.parse(row.updated_at) || Date.now(),
  };
}

async function ensureMonthlyFixedExpenses() {
  if (!currentUser || !state.fixedExpenses.length) return 0;

  const now = new Date();
  const year = now.getFullYear();
  const monthIndex = now.getMonth();
  const monthKey = `${year}-${String(monthIndex + 1).padStart(2, "0")}`;
  const lastDay = new Date(year, monthIndex + 1, 0).getDate();
  const monthEnd = `${monthKey}-${String(lastDay).padStart(2, "0")}`;

  const existingByFixedId = new Set(
    state.expenses.filter((expense) => isSameMonthKey(expense.date, monthKey) && expense.fixedExpenseId).map((expense) => expense.fixedExpenseId)
  );
  const toInsert = [];

  state.fixedExpenses.forEach((fixedExpense) => {
    if (!fixedExpense.active || !fixedExpense.startDate) return;
    if (fixedExpense.startDate > monthEnd) return;
    if (existingByFixedId.has(fixedExpense.id)) return;

    const targetDay = Math.min(fixedExpense.dayOfMonth, lastDay);
    const date = `${monthKey}-${String(targetDay).padStart(2, "0")}`;
    if (date < fixedExpense.startDate) return;

    toInsert.push({
      user_id: currentUser.id,
      date,
      content: fixedExpense.content,
      amount: fixedExpense.amount,
      case_id: null,
      fixed_expense_id: fixedExpense.id,
    });
  });

  if (!toInsert.length) return 0;

  const { error } = await sbClient.from("expenses").insert(toInsert);
  if (error) {
    if (String(error.message || "").includes("fixed_expense_id")) {
      throw new Error(
        "expenses.fixed_expense_id が未作成です。README記載の SQL を実行してください: alter table expenses add column if not exists fixed_expense_id uuid;"
      );
    }
    throw error;
  }

  return toInsert.length;
}

async function withLoading(task, fallbackMessage, authArea = false) {
  showLoading(true);
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
    showLoading(false);
  }
}

function showLoading(isLoading) {
  if (!loadingOverlay) return;

  if (isLoading) {
    loadingCount += 1;
    loadingOverlay.hidden = false;
    return;
  }

  loadingCount = Math.max(0, loadingCount - 1);
  loadingOverlay.hidden = true;
}

function setLoading(isLoading) {
  showLoading(isLoading);
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
