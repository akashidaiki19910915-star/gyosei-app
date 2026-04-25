const SUPABASE_URL = "https://ueelzyftlbnvjvpsmpyt.supabase.co";
const SUPABASE_KEY = "sb_publishable_0DrKsieUcCyEZN_HRg8LhQ_QqFTPMtp";
const STATUS_ORDER = ["未着手", "進行中", "完了"];
const STATUS_FILTER_KEYS = [...STATUS_ORDER, "その他"];
const DEADLINE_FILTER_KEYS = ["all", "overdue", "within7", "within30"];
const ESTIMATE_STATUS_ORDER = ["作成中", "提出済", "受注", "失注"];
const EXPENSE_PAYMENT_METHODS = ["現金", "クレジットカード", "銀行振込", "電子マネー", "口座振替", "その他"];
const OFFICE_INFO = {
  name: "あかし行政書士事務所",
  zip: "574-0044",
  address: "大阪府大東市諸福2-2-14",
  tel: "090-8193-4812",
  email: "akashigyousei@gmail.com",
  registrationNumber: "",
  transferInfo: "",
  invoiceNote: "振込手数料はご負担をお願いいたします。",
  estimateNote: "本見積の有効期限内にご発注をお願いいたします。",
};

const sbClient = window.supabase.createClient(SUPABASE_URL, SUPABASE_KEY, {
  auth: {
    persistSession: true,
    autoRefreshToken: true,
    detectSessionInUrl: true,
    storageKey: "gyosei-app-auth",
  },
});

const state = {
  clients: [],
  cases: [],
  estimates: [],
  estimateItems: [],
  sales: [],
  expenses: [],
  fixedExpenses: [],
  dailyReports: [],
  selectedAggregation: "month",
  selectedMonth: toMonthKey(new Date()),
  selectedYear: new Date().getFullYear(),
  caseStatusFilter: "all",
  caseSearchQuery: "",
  caseDeadlineFilter: "all",
  salesSearchQuery: "",
  expensesSearchQuery: "",
  dailyReportSearchQuery: "",
  dailyReportDateFilter: "all",
  estimateCustomerQuery: "",
  estimateTitleQuery: "",
  estimateStatusFilter: "all",
  estimateExpiredFilter: "all",
};
const editState = { clientId: null, caseId: null, saleId: null, expenseId: null, fixedExpenseId: null, dailyReportId: null, estimateId: null };

const authView = document.getElementById("auth-view");
const appView = document.getElementById("app-view");
const loadingOverlay = document.getElementById("loading-overlay");
const loadingForceCloseBtn = document.getElementById("loading-force-close-btn");
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
const subtabButtons = Array.from(document.querySelectorAll(".subtab-btn"));
const subtabPanels = Array.from(document.querySelectorAll(".subtab-panel"));
const panels = {
  clients: document.getElementById("tab-clients"),
  cases: document.getElementById("tab-cases"),
  estimates: document.getElementById("tab-estimates"),
  sales: document.getElementById("tab-sales"),
  expenses: document.getElementById("tab-expenses"),
  "daily-reports": document.getElementById("tab-daily-reports"),
};
const dashboardSection = document.querySelector(".dashboard");
const subtabState = {
  cases: "dashboard",
  estimates: "create",
  sales: "entry",
  expenses: "entry",
  "daily-reports": "entry",
};

const summaryGrid = document.getElementById("summary-grid");
const worktimeSummaryGrid = document.getElementById("worktime-summary-grid");
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
const nextActionAlertCard = document.getElementById("next-action-alert-card");
const nextActionAlertSummary = document.getElementById("next-action-alert-summary");
const nextActionAlertBody = document.getElementById("next-action-alert-body");
const nextActionAlertEmpty = document.getElementById("next-action-alert-empty");
const nextActionAlertListWrap = document.getElementById("next-action-alert-list-wrap");
const todayTaskCard = document.getElementById("today-task-card");
const todayTaskSummary = document.getElementById("today-task-summary");
const todayTaskBody = document.getElementById("today-task-body");
const todayTaskEmpty = document.getElementById("today-task-empty");
const todayTaskListWrap = document.getElementById("today-task-list-wrap");

const clientForm = document.getElementById("client-form");
const clientsList = document.getElementById("clients-list");
const clientsEmpty = document.getElementById("clients-empty");
const clientSubmitBtn = document.getElementById("client-submit-btn");
const caseClientSelect = document.getElementById("case-client-id");
const estimateClientSelect = document.getElementById("estimate-client-id");

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
const saleInvoiceMemoInput = document.getElementById("sale-invoice-memo");
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
const dailyReportForm = document.getElementById("daily-report-form");
const reportCaseSelect = document.getElementById("report-case-id");
const dailyReportSubmitBtn = document.getElementById("daily-report-submit-btn");
const dailyReportsBody = document.getElementById("daily-reports-body");
const dailyReportsEmpty = document.getElementById("daily-reports-empty");
const dailyReportsListWrap = document.getElementById("daily-reports-list-wrap");
const dailyReportSummaryList = document.getElementById("daily-report-summary-list");
const dailyReportSearchInput = document.getElementById("daily-report-search-input");
const dailyReportDateFilterSelect = document.getElementById("daily-report-date-filter");
const dailyReportFilterClearBtn = document.getElementById("daily-report-filter-clear-btn");
const dailyReportFilterCount = document.getElementById("daily-report-filter-count");

const caseItemTemplate = document.getElementById("case-item-template");
const saleItemTemplate = document.getElementById("sale-item-template");
const expenseItemTemplate = document.getElementById("expense-item-template");
const fixedExpenseItemTemplate = document.getElementById("fixed-expense-item-template");
const estimateForm = document.getElementById("estimate-form");
const estimateItemsWrap = document.getElementById("estimate-items");
const estimateAddItemBtn = document.getElementById("estimate-add-item-btn");
const estimateSubmitBtn = document.getElementById("estimate-submit-btn");
const estimateSubtotal = document.getElementById("estimate-subtotal");
const estimateTax = document.getElementById("estimate-tax");
const estimateTotal = document.getElementById("estimate-total");
const estimateList = document.getElementById("estimate-list");
const estimateEmpty = document.getElementById("estimate-empty");
const estimateCustomerSearch = document.getElementById("estimate-customer-search");
const estimateTitleSearch = document.getElementById("estimate-title-search");
const estimateStatusFilter = document.getElementById("estimate-status-filter");
const estimateExpiredFilter = document.getElementById("estimate-expired-filter");

let currentUser = null;
let isResuming = false;
let isLoggingOut = false;
let loadingTimerId = null;

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
  subtabButtons.forEach((btn) => btn.addEventListener("click", () => activateSubtab(btn.dataset.parentTab, btn.dataset.subtab)));
  authForm.addEventListener("submit", handleLogin);
  signupBtn.addEventListener("click", handleSignup);
  logoutBtn.addEventListener("click", handleLogout);

  clientForm?.addEventListener("submit", handleClientSubmit);
  caseForm.addEventListener("submit", handleCaseSubmit);
  saleForm.addEventListener("submit", handleSaleSubmit);
  expenseForm.addEventListener("submit", handleExpenseSubmit);
  fixedExpenseForm.addEventListener("submit", handleFixedExpenseSubmit);
  dailyReportForm.addEventListener("submit", handleDailyReportSubmit);
  estimateForm?.addEventListener("submit", handleEstimateSubmit);

  clearBtn.addEventListener("click", handleClearAll);
  clientsList?.addEventListener("click", handleClientListAction);
  caseList.addEventListener("click", handleCaseListAction);
  salesList.addEventListener("click", handleSalesListAction);
  expensesList.addEventListener("click", handleExpensesListAction);
  fixedExpensesList.addEventListener("click", handleFixedExpensesListAction);
  dailyReportsBody?.addEventListener("click", handleDailyReportsListAction);
  unpaidListBody?.addEventListener("click", handleUnpaidListAction);
  billingLeakAlertBody?.addEventListener("click", handleBillingLeakAlertAction);
  targetMonthInput?.addEventListener("input", handleTargetMonthChange);
  targetYearInput?.addEventListener("input", handleTargetYearChange);
  aggregationRadios.forEach((radio) => radio.addEventListener("change", handleAggregationChange));
  statusSummaryList?.addEventListener("click", handleStatusSummaryClick);
  statusFilterClearBtn?.addEventListener("click", () => applyCaseStatusFilter("all"));
  deadlineAlertSummary?.addEventListener("click", handleDeadlineAlertClick);
  deadlineAlertBody?.addEventListener("click", handleDeadlineAlertClick);
  nextActionAlertBody?.addEventListener("click", handleNextActionAlertClick);
  todayTaskBody?.addEventListener("click", handleTodayTaskAction);
  caseSearchInput?.addEventListener("input", handleCaseSearchInput);
  caseStatusFilterSelect?.addEventListener("change", handleCaseStatusFilterChange);
  caseDeadlineFilterSelect?.addEventListener("change", handleCaseDeadlineFilterChange);
  caseFilterClearBtn?.addEventListener("click", clearCaseFilters);
  salesSearchInput?.addEventListener("input", handleSalesSearchInput);
  salesFilterClearBtn?.addEventListener("click", clearSalesSearch);
  expensesSearchInput?.addEventListener("input", handleExpensesSearchInput);
  expensesFilterClearBtn?.addEventListener("click", clearExpensesSearch);
  dailyReportSearchInput?.addEventListener("input", handleDailyReportSearchInput);
  dailyReportDateFilterSelect?.addEventListener("change", handleDailyReportDateFilterChange);
  dailyReportFilterClearBtn?.addEventListener("click", clearDailyReportFilters);
  document.addEventListener("visibilitychange", handleVisibilityChange);
  window.addEventListener("focus", handleWindowFocus);
  window.addEventListener("pageshow", handlePageShow);
  loadingForceCloseBtn?.addEventListener("click", () => {
    showLoading(false);
    showAppMessage("読み込みを強制解除しました。必要なら再操作してください。", true);
  });
  exportCasesCsvBtn?.addEventListener("click", handleExportCasesCsv);
  exportSalesCsvBtn?.addEventListener("click", handleExportSalesCsv);
  exportExpensesCsvBtn?.addEventListener("click", handleExportExpensesCsv);
  exportFixedExpensesCsvBtn?.addEventListener("click", handleExportFixedExpensesCsv);
  exportAllCsvBtn?.addEventListener("click", handleExportAllCsv);
  csvImportForm?.addEventListener("submit", handleCsvImportSubmit);
  exportExcelBtn?.addEventListener("click", handleExportExcel);
  excelImportForm?.addEventListener("submit", handleExcelImportSubmit);
  estimateAddItemBtn?.addEventListener("click", () => addEstimateItemRow());
  estimateItemsWrap?.addEventListener("input", handleEstimateItemsInput);
  estimateItemsWrap?.addEventListener("click", handleEstimateItemsClick);
  estimateList?.addEventListener("click", handleEstimateListAction);
  caseClientSelect?.addEventListener("change", syncCaseCustomerFromClient);
  estimateClientSelect?.addEventListener("change", syncEstimateCustomerFromClient);
  estimateCustomerSearch?.addEventListener("input", (event) => {
    state.estimateCustomerQuery = String(event.target.value || "").trim().toLowerCase();
    renderEstimates();
  });
  estimateTitleSearch?.addEventListener("input", (event) => {
    state.estimateTitleQuery = String(event.target.value || "").trim().toLowerCase();
    renderEstimates();
  });
  estimateStatusFilter?.addEventListener("change", (event) => {
    state.estimateStatusFilter = event.target.value || "all";
    renderEstimates();
  });
  estimateExpiredFilter?.addEventListener("change", (event) => {
    state.estimateExpiredFilter = event.target.value || "all";
    renderEstimates();
  });
  activateTab("cases");
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

function handleNextActionAlertClick(event) {
  const button = event.target.closest("button");
  if (!(button instanceof HTMLButtonElement)) return;
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

function handleDailyReportSearchInput(event) {
  state.dailyReportSearchQuery = String(event?.target?.value ?? "").trim().toLowerCase();
  renderDailyReports();
}

function handleDailyReportDateFilterChange(event) {
  applyDailyReportDateFilter(event?.target?.value || "all");
}

function applyDailyReportDateFilter(nextFilter) {
  const normalized = ["all", "today", "month"].includes(nextFilter) ? nextFilter : "all";
  state.dailyReportDateFilter = normalized;
  if (dailyReportDateFilterSelect) dailyReportDateFilterSelect.value = normalized;
  renderDailyReports();
}

function clearDailyReportFilters() {
  state.dailyReportSearchQuery = "";
  applyDailyReportDateFilter("all");
  if (dailyReportSearchInput) dailyReportSearchInput.value = "";
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
  if (!currentUser || isResuming || isLoggingOut) return;

  isResuming = true;
  try {
    await withLoading("画面復帰時の再読み込み", async () => {
      await loadAllData();
      renderAfterDataChanged();
    });
  } catch (error) {
    console.error(`画面復帰時の再読み込みに失敗しました。(${trigger})`, error);
  } finally {
    isResuming = false;
  }
}

async function applyAuthState() {
  if (!currentUser) {
    clearAppState();
    return;
  }

  authView.hidden = true;
  appView.hidden = false;
  userLabel.textContent = currentUser.email || "ログイン中";

  await withLoading("データの読み込み", async () => {
    await loadAllData();
    resetCaseForm();
    resetEstimateForm();
    resetSaleForm();
    resetExpenseForm();
    resetFixedExpenseForm();
    resetDailyReportForm();
    renderAfterDataChanged();
  });
}

async function handleLogin(event) {
  event.preventDefault();
  const email = authForm.elements.email.value.trim();
  const password = authForm.elements.password.value;
  if (!email || !password) return;

  await withLoading("ログイン", async () => {
    const { error } = await sbClient.auth.signInWithPassword({ email, password });
    if (error) throw error;
    showAuthMessage("ログインしました。", false);
  }, { messageTarget: "auth", triggerButton: event.submitter });
}

async function handleSignup() {
  const email = authForm.elements.email.value.trim();
  const password = authForm.elements.password.value;
  if (!email || !password) {
    showAuthMessage("メールアドレスとパスワードを入力してください。", true);
    return;
  }

  await withLoading("新規登録", async () => {
    const { error } = await sbClient.auth.signUp({ email, password });
    if (error) throw error;
    showAuthMessage("新規登録が完了しました。確認メールを確認してください。", false);
  }, { messageTarget: "auth", triggerButton: signupBtn });
}

async function handleLogout() {
  isLoggingOut = true;
  try {
    await withLoading("ログアウト", async () => {
      currentUser = null;
      const { error } = await sbClient.auth.signOut();
      if (error) throw error;
      clearAppState();
      showAuthMessage("ログアウトしました。", false);
      authForm.reset();
    }, { messageTarget: "auth", triggerButton: logoutBtn });
  } finally {
    isLoggingOut = false;
  }
}

async function loadAllData() {
  if (!currentUser || isLoggingOut) return;

  const [clientsRes, casesRes, estimatesRes, estimateItemsRes, salesRes, expensesRes, fixedExpensesRes, dailyReportsRes] = await Promise.all([
    sbClient.from("clients").select("*").eq("user_id", currentUser.id),
    sbClient.from("cases").select("*").eq("user_id", currentUser.id),
    sbClient.from("estimates").select("*").eq("user_id", currentUser.id),
    sbClient.from("estimate_items").select("*").eq("user_id", currentUser.id),
    sbClient.from("sales").select("*").eq("user_id", currentUser.id),
    sbClient.from("expenses").select("*").eq("user_id", currentUser.id),
    sbClient.from("fixed_expenses").select("*").eq("user_id", currentUser.id),
    sbClient.from("daily_reports").select("*").eq("user_id", currentUser.id),
  ]);

  if (clientsRes.error) throw clientsRes.error;
  if (casesRes.error) throw casesRes.error;
  if (estimatesRes.error) throw estimatesRes.error;
  if (estimateItemsRes.error) throw estimateItemsRes.error;
  if (salesRes.error) throw salesRes.error;
  if (expensesRes.error) throw expensesRes.error;
  if (fixedExpensesRes.error) throw fixedExpensesRes.error;
  if (dailyReportsRes.error) throw dailyReportsRes.error;

  state.clients = (clientsRes.data || []).map(mapClientFromDb);
  state.cases = (casesRes.data || []).map(mapCaseFromDb);
  state.estimates = (estimatesRes.data || []).map(mapEstimateFromDb);
  state.estimateItems = (estimateItemsRes.data || []).map(mapEstimateItemFromDb);
  state.sales = (salesRes.data || []).map(mapSaleFromDb);
  state.expenses = (expensesRes.data || []).map(mapExpenseFromDb);
  state.fixedExpenses = (fixedExpensesRes.data || []).map(mapFixedExpenseFromDb);
  state.dailyReports = (dailyReportsRes.data || []).map(mapDailyReportFromDb);
  await cleanupLegacyEstimateMemoMarkers();

  const createdCount = await ensureMonthlyFixedExpenses();
  if (createdCount > 0) {
    const refreshExpensesRes = await sbClient.from("expenses").select("*").eq("user_id", currentUser.id);
    if (refreshExpensesRes.error) throw refreshExpensesRes.error;
    state.expenses = (refreshExpensesRes.data || []).map(mapExpenseFromDb);
    showAppMessage(`${createdCount}件の固定費を当月分として自動計上しました。`, false);
  }
}

async function refreshAfterMutation(successMessage, taskName = "") {
  await loadAllData();
  console.log("loadAllData done", taskName);
  renderAfterDataChanged();
  console.log("render done", taskName);
  if (successMessage) showAppMessage(successMessage, false);
}

async function handleClientSubmit(event) {
  event.preventDefault();
  if (!currentUser || !clientForm) return;
  const name = asTrimmedText(clientForm.elements.clientName.value);
  if (!name) return;
  const payload = {
    user_id: currentUser.id,
    name,
    client_type: asTrimmedText(clientForm.elements.clientType.value) || null,
    address: asTrimmedText(clientForm.elements.clientAddress.value) || null,
    tel: asTrimmedText(clientForm.elements.clientTel.value) || null,
    email: asTrimmedText(clientForm.elements.clientEmail.value) || null,
    referral_source: asTrimmedText(clientForm.elements.clientReferralSource.value) || null,
    memo: asTrimmedText(clientForm.elements.clientMemo.value) || null,
  };
  const taskName = editState.clientId ? "顧客更新" : "顧客登録";
  await withLoading(taskName, async () => {
    if (editState.clientId) {
      const { data, error } = await sbClient.from("clients").update(payload).eq("id", editState.clientId).eq("user_id", currentUser.id).select().single();
      if (error) throw error;
      if (!data) throw new Error("更新結果を取得できませんでした。");
      console.log("DB success", taskName);
      await refreshAfterMutation("顧客を更新しました。", taskName);
      resetClientForm();
      return;
    }
    const { data, error } = await sbClient.from("clients").insert(payload).select().single();
    if (error) throw error;
    if (!data) throw new Error("登録結果を取得できませんでした。");
    console.log("DB success", taskName);
    await refreshAfterMutation("顧客を登録しました。", taskName);
    resetClientForm();
  }, { triggerButton: event.submitter });
}

async function handleCaseSubmit(event) {
  event.preventDefault();
  if (!currentUser) return;

  const customerName = caseForm.elements.customerName.value.trim();
  const caseName = caseForm.elements.caseName.value.trim();
  if (!customerName || !caseName) return;

  const selectedClientId = caseForm.elements.caseClientId.value || null;
  const selectedClient = selectedClientId ? state.clients.find((entry) => entry.id === selectedClientId) : null;
  const payload = {
    user_id: currentUser.id,
    client_id: selectedClientId,
    customer_name: selectedClient?.name || customerName,
    case_name: caseName,
    estimate_amount: normalizeAmount(caseForm.elements.amount.value),
    received_date: caseForm.elements.receivedDate.value || null,
    due_date: caseForm.elements.dueDate.value || null,
    work_memo: asTrimmedText(caseForm.elements.workMemo.value) || null,
    next_action_date: caseForm.elements.nextActionDate.value || null,
    next_action: asTrimmedText(caseForm.elements.nextAction.value) || null,
    document_url: asTrimmedText(caseForm.elements.documentUrl.value) || null,
    invoice_url: asTrimmedText(caseForm.elements.invoiceUrl.value) || null,
    receipt_url: asTrimmedText(caseForm.elements.receiptUrl.value) || null,
    status: normalizeStatus(caseForm.elements.status.value),
  };

  const taskName = editState.caseId ? "案件更新" : "案件登録";
  await withLoading(taskName, async () => {
    if (editState.caseId) {
      const { data, error } = await sbClient.from("cases").update(payload).eq("id", editState.caseId).eq("user_id", currentUser.id).select().single();
      if (error) throw error;
      if (!data) throw new Error("更新結果を取得できませんでした。");
      console.log("DB success", taskName);
      await refreshAfterMutation("案件を更新しました。", taskName);
      subtabState.cases = "list";
      resetCaseForm();
      return;
    }

    const { data, error } = await sbClient.from("cases").insert(payload).select().single();
    if (error) throw error;
    if (!data) throw new Error("登録結果を取得できませんでした。");
    console.log("DB success", taskName);
    await refreshAfterMutation("案件を登録しました。", taskName);
    subtabState.cases = "list";
    resetCaseForm();
  }, { triggerButton: event.submitter });
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

  const taskName = editState.saleId ? "売上更新" : "売上登録";
  await withLoading(taskName, async () => {
    if (editState.saleId) {
      const { data, error } = await sbClient.from("sales").update(payload).eq("id", editState.saleId).eq("user_id", currentUser.id).select().single();
      if (error) throw error;
      if (!data) throw new Error("更新結果を取得できませんでした。");
      console.log("DB success", taskName);
      await refreshAfterMutation("売上を更新しました。", taskName);
      resetSaleForm();
      return;
    }

    payload.invoice_number = await getNextMonthlyNumber("sales", "invoice_number", "S");
    const { data, error } = await sbClient.from("sales").insert(payload).select().single();
    if (error) throw error;
    if (!data) throw new Error("登録結果を取得できませんでした。");
    console.log("DB success", taskName);
    await refreshAfterMutation("売上を登録しました。", taskName);
    resetSaleForm();
  }, { triggerButton: event.submitter });
}

async function handleExpenseSubmit(event) {
  event.preventDefault();
  if (!currentUser) return;

  const expenseDate = expenseForm.elements.expenseDate.value;
  const content = expenseForm.elements.expenseContent.value.trim();
  const payee = asTrimmedText(expenseForm.elements.expensePayee.value) || null;
  const paymentMethod = normalizeExpensePaymentMethod(expenseForm.elements.expensePaymentMethod.value);
  const amount = normalizeAmount(expenseForm.elements.expenseAmount.value);
  if (!expenseDate || !content || amount === null) return;

  const payload = {
    user_id: currentUser.id,
    date: expenseDate,
    content,
    amount,
    payee,
    payment_method: paymentMethod,
    case_id: expenseCaseSelect.value || null,
    receipt_url: asTrimmedText(expenseForm.elements.expenseReceiptUrl.value) || null,
  };

  const taskName = editState.expenseId ? "経費更新" : "経費登録";
  await withLoading(taskName, async () => {
    if (editState.expenseId) {
      const currentExpense = state.expenses.find((entry) => entry.id === editState.expenseId);
      const updatePayload = { ...payload };
      if (currentExpense?.fixedExpenseId) updatePayload.fixed_expense_id = currentExpense.fixedExpenseId;
      const { data, error } = await sbClient.from("expenses").update(updatePayload).eq("id", editState.expenseId).eq("user_id", currentUser.id).select().single();
      if (error) {
        showAppMessage(`経費更新に失敗しました。\n詳細: ${error.message || error}`, true);
        throw error;
      }
      if (!data) throw new Error("更新結果を取得できませんでした。");
      console.log("DB success", taskName);
      await refreshAfterMutation("経費を更新しました。", taskName);
      resetExpenseForm();
      return;
    }

    const { data, error } = await sbClient.from("expenses").insert(payload).select().single();
    if (error) {
      showAppMessage(`経費登録に失敗しました。\n詳細: ${error.message || error}`, true);
      throw error;
    }
    if (!data) throw new Error("登録結果を取得できませんでした。");
    console.log("DB success", taskName);
    await refreshAfterMutation("経費を登録しました。", taskName);
    resetExpenseForm();
  }, { triggerButton: event.submitter });
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

  const taskName = editState.fixedExpenseId ? "固定費更新" : "固定費登録";
  await withLoading(taskName, async () => {
    if (editState.fixedExpenseId) {
      const { data, error } = await sbClient
        .from("fixed_expenses")
        .update(payload)
        .eq("id", editState.fixedExpenseId)
        .eq("user_id", currentUser.id)
        .select()
        .single();
      if (error) throw error;
      if (!data) throw new Error("更新結果を取得できませんでした。");
      console.log("DB success", taskName);
      await refreshAfterMutation("固定費を更新しました。", taskName);
      resetFixedExpenseForm();
      return;
    }

    const { data, error } = await sbClient.from("fixed_expenses").insert(payload).select().single();
    if (error) throw error;
    if (!data) throw new Error("登録結果を取得できませんでした。");
    console.log("DB success", taskName);
    await refreshAfterMutation("固定費を登録しました。", taskName);
    resetFixedExpenseForm();
  }, { triggerButton: event.submitter });
}

async function handleDailyReportSubmit(event) {
  event.preventDefault();
  if (!currentUser) return;

  const reportDate = dailyReportForm.elements.reportDate.value;
  const workContent = asTrimmedText(dailyReportForm.elements.reportWorkContent.value);
  if (!reportDate || !workContent) return;

  const payload = {
    user_id: currentUser.id,
    report_date: reportDate,
    case_id: reportCaseSelect.value || null,
    work_content: workContent,
    work_minutes: normalizeAmount(dailyReportForm.elements.reportWorkMinutes.value) ?? 0,
    next_action: asTrimmedText(dailyReportForm.elements.reportNextAction.value) || null,
    memo: asTrimmedText(dailyReportForm.elements.reportMemo.value) || null,
  };

  const taskName = editState.dailyReportId ? "日報更新" : "日報登録";
  await withLoading(taskName, async () => {
    if (editState.dailyReportId) {
      const { data, error } = await sbClient.from("daily_reports").update(payload).eq("id", editState.dailyReportId).eq("user_id", currentUser.id).select().single();
      if (error) throw error;
      if (!data) throw new Error("更新結果を取得できませんでした。");
      console.log("DB success", taskName);
      await refreshAfterMutation("日報を更新しました。", taskName);
      resetDailyReportForm();
      return;
    }
    const { data, error } = await sbClient.from("daily_reports").insert(payload).select().single();
    if (error) throw error;
    if (!data) throw new Error("登録結果を取得できませんでした。");
    console.log("DB success", taskName);
    await refreshAfterMutation("日報を登録しました。", taskName);
    resetDailyReportForm();
  }, { triggerButton: event.submitter });
}

async function handleClientListAction(event) {
  const btn = event.target.closest("button");
  if (!(btn instanceof HTMLButtonElement) || !currentUser) return;
  const item = btn.closest(".item");
  const id = item?.dataset.id;
  if (!id) return;
  if (btn.classList.contains("edit-client-btn")) {
    const target = state.clients.find((entry) => entry.id === id);
    if (!target || !clientForm) return;
    activateTab("clients");
    editState.clientId = target.id;
    clientForm.elements.clientName.value = target.name || "";
    clientForm.elements.clientType.value = target.clientType || "";
    clientForm.elements.clientAddress.value = target.address || "";
    clientForm.elements.clientTel.value = target.tel || "";
    clientForm.elements.clientEmail.value = target.email || "";
    clientForm.elements.clientReferralSource.value = target.referralSource || "";
    clientForm.elements.clientMemo.value = target.memo || "";
    if (clientSubmitBtn) clientSubmitBtn.textContent = "顧客を更新";
    clientForm.scrollIntoView({ behavior: "smooth", block: "start" });
    return;
  }
  if (btn.classList.contains("delete-client-btn")) {
    if (!window.confirm("この顧客を削除しますか？案件・見積は削除されません。")) return;
    await withLoading("顧客削除", async () => {
      const { error } = await sbClient.from("clients").delete().eq("id", id).eq("user_id", currentUser.id);
      if (error) throw error;
      console.log("DB success", "顧客削除");
      if (editState.clientId === id) resetClientForm();
      await refreshAfterMutation("顧客を削除しました。", "顧客削除");
    });
  }
}

async function handleCaseListAction(event) {
  const btn = event.target;
  if (!(btn instanceof HTMLButtonElement)) return;
  const item = btn.closest(".item");
  if (!item || !currentUser) return;
  const id = item.dataset.id;

  if (btn.classList.contains("edit-btn")) {
    await startCaseEdit(id);
    return;
  }
  if (btn.classList.contains("case-print-btn")) {
    await withLoading("帳票出力", async () => {
      openInvoicePrintPreviewFromCase(id);
    });
    return;
  }
  if (btn.classList.contains("case-xlsx-btn")) {
    await withLoading("帳票出力", async () => {
      exportInvoiceDataForCase(id);
    });
    return;
  }

  if (btn.classList.contains("delete-btn")) {
    if (!window.confirm("この案件を削除しますか？関連する売上・経費も削除されます。")) return;

    await withLoading("案件削除", async () => {
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

      console.log("DB success", "案件削除");
      if (editState.caseId === id) resetCaseForm();
      await refreshAfterMutation("案件を削除しました。", "案件削除");
    });
  }
}

async function startCaseEdit(caseId) {
  await withLoading("編集ボタン処理", async () => {
    const target = state.cases.find((entry) => entry.id === caseId);
    if (!target) return;
    subtabState.cases = "entry";
    activateTab("cases");
    editState.caseId = target.id;
    caseForm.elements.caseClientId.value = target.clientId || "";
    caseForm.elements.customerName.value = target.customerName;
    caseForm.elements.caseName.value = target.caseName;
    caseForm.elements.amount.value = target.estimateAmount ?? "";
    caseForm.elements.receivedDate.value = target.receivedDate || "";
    caseForm.elements.dueDate.value = target.dueDate || "";
    caseForm.elements.workMemo.value = sanitizeLegacyEstimateMemo(target.workMemo) || "";
    caseForm.elements.nextActionDate.value = target.nextActionDate || "";
    caseForm.elements.nextAction.value = target.nextAction || "";
    caseForm.elements.documentUrl.value = target.documentUrl || "";
    caseForm.elements.invoiceUrl.value = target.invoiceUrl || "";
    caseForm.elements.receiptUrl.value = target.receiptUrl || "";
    caseForm.elements.status.value = normalizeStatus(target.status);
    caseSubmitBtn.textContent = "案件を更新";
    caseForm.scrollIntoView({ behavior: "smooth", block: "start" });
  });
}

async function handleSalesListAction(event) {
  const btn = event.target.closest("button");
  if (!(btn instanceof HTMLButtonElement)) return;
  const item = btn.closest(".item");
  if (!item || !currentUser) return;
  const id = item.dataset.id;

  if (btn.classList.contains("edit-btn")) {
    await editSale(id);
    return;
  }

  if (btn.classList.contains("delete-btn")) {
    if (!window.confirm("この売上を削除しますか？")) return;
    await withLoading("売上削除", async () => {
      const { error } = await sbClient.from("sales").delete().eq("id", id).eq("user_id", currentUser.id);
      if (error) throw error;
      console.log("DB success", "売上削除");
      if (editState.saleId === id) resetSaleForm();
      await refreshAfterMutation("売上を削除しました。", "売上削除");
    });
  }
}

async function handleExpensesListAction(event) {
  const btn = event.target.closest("button");
  if (!(btn instanceof HTMLButtonElement)) return;
  const item = btn.closest(".item");
  if (!item || !currentUser) return;
  const id = item.dataset.id;

  if (btn.classList.contains("edit-btn")) {
    await editExpense(id);
    return;
  }

  if (btn.classList.contains("delete-btn")) {
    if (!window.confirm("この経費を削除しますか？")) return;
    await withLoading("経費削除", async () => {
      const { error } = await sbClient.from("expenses").delete().eq("id", id).eq("user_id", currentUser.id);
      if (error) throw error;
      console.log("DB success", "経費削除");
      if (editState.expenseId === id) resetExpenseForm();
      await refreshAfterMutation("経費を削除しました。", "経費削除");
    });
  }
}

function handleUnpaidListAction(event) {
  const btn = event.target.closest("button");
  if (!(btn instanceof HTMLButtonElement)) return;
  const saleId = btn.dataset.saleId;
  if (!saleId) return;
  if (btn.classList.contains("edit-sale-btn")) {
    editSale(saleId).catch(() => {});
  }
}

function handleTodayTaskAction(event) {
  const btn = event.target.closest("button");
  if (!(btn instanceof HTMLButtonElement)) return;
  const taskId = btn.dataset.taskId;
  const target = btn.dataset.taskTarget;
  if (!taskId || !target) return;
  if (target === "case") {
    startCaseEdit(taskId).catch(() => {});
    return;
  }
  if (target === "sale") {
    editSale(taskId).catch(() => {});
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

async function startSaleEdit(saleId) {
  await withLoading("編集ボタン処理", async () => {
    const target = state.sales.find((entry) => entry.id === saleId);
    if (!target) return;
    subtabState.sales = "entry";
    activateTab("sales");
    editState.saleId = target.id;
    saleCaseSelect.value = target.caseId;
    saleForm.elements.invoiceAmount.value = target.invoiceAmount;
    saleForm.elements.paidAmount.value = target.paidAmount ?? "";
    saleForm.elements.paidDate.value = target.paidDate || "";
    saleForm.elements.isUnpaid.checked = Boolean(target.isUnpaid);
    saleSubmitBtn.textContent = "売上を更新";
    saleForm.scrollIntoView({ behavior: "smooth", block: "start" });
  });
}

async function startExpenseEdit(expenseId) {
  await withLoading("編集ボタン処理", async () => {
    const target = state.expenses.find((entry) => entry.id === expenseId);
    if (!target) return;
    subtabState.expenses = "entry";
    activateTab("expenses");
    editState.expenseId = target.id;
    expenseForm.elements.expenseDate.value = target.date;
    expenseForm.elements.expenseContent.value = target.content;
    expenseForm.elements.expensePayee.value = target.payee || "";
    expenseForm.elements.expensePaymentMethod.value = target.paymentMethod || "";
    expenseForm.elements.expenseAmount.value = target.amount;
    expenseForm.elements.expenseReceiptUrl.value = target.receiptUrl || "";
    expenseCaseSelect.value = target.caseId || "";
    expenseSubmitBtn.textContent = "経費を更新";
    expenseForm.scrollIntoView({ behavior: "smooth", block: "start" });
  });
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

async function startFixedExpenseEdit(fixedExpenseId) {
  await withLoading("編集ボタン処理", async () => {
    const target = state.fixedExpenses.find((entry) => entry.id === fixedExpenseId);
    if (!target) return;
    activateTab("expenses");
    editState.fixedExpenseId = target.id;
    fixedExpenseForm.elements.fixedExpenseContent.value = target.content;
    fixedExpenseForm.elements.fixedExpenseAmount.value = target.amount;
    fixedExpenseForm.elements.fixedExpenseDayOfMonth.value = target.dayOfMonth;
    fixedExpenseForm.elements.fixedExpenseStartDate.value = target.startDate || "";
    fixedExpenseForm.elements.fixedExpenseActive.checked = Boolean(target.active);
    fixedExpenseSubmitBtn.textContent = "固定費を更新";
    fixedExpenseForm.scrollIntoView({ behavior: "smooth", block: "start" });
  });
}

async function handleFixedExpensesListAction(event) {
  const btn = event.target;
  if (!(btn instanceof HTMLButtonElement)) return;
  const item = btn.closest(".item");
  if (!item || !currentUser) return;
  const id = item.dataset.id;

  if (btn.classList.contains("edit-btn")) {
    await startFixedExpenseEdit(id);
    return;
  }

  if (btn.classList.contains("toggle-btn")) {
    const target = state.fixedExpenses.find((entry) => entry.id === id);
    if (!target) return;

    await withLoading("固定費更新", async () => {
      const { data, error } = await sbClient
        .from("fixed_expenses")
        .update({ active: !target.active })
        .eq("id", id)
        .eq("user_id", currentUser.id)
        .select()
        .single();
      if (error) throw error;
      if (!data) throw new Error("更新結果を取得できませんでした。");
      console.log("DB success", "固定費更新");
      await refreshAfterMutation("固定費の有効/無効を更新しました。", "固定費更新");
    });
    return;
  }

  if (btn.classList.contains("delete-btn")) {
    if (!window.confirm("この固定費を削除しますか？")) return;
    await withLoading("固定費削除", async () => {
      const { error } = await sbClient.from("fixed_expenses").delete().eq("id", id).eq("user_id", currentUser.id);
      if (error) throw error;
      console.log("DB success", "固定費削除");
      if (editState.fixedExpenseId === id) resetFixedExpenseForm();
      await refreshAfterMutation("固定費を削除しました。", "固定費削除");
    });
  }
}

async function handleDailyReportsListAction(event) {
  const btn = event.target.closest("button");
  if (!(btn instanceof HTMLButtonElement) || !currentUser) return;
  const id = btn.dataset.dailyReportId;
  if (!id) return;

  if (btn.classList.contains("edit-daily-report-btn")) {
    await editDailyReport(id);
    return;
  }

  if (btn.classList.contains("delete-daily-report-btn")) {
    if (!window.confirm("この日報を削除しますか？")) return;
    await withLoading("日報削除", async () => {
      const { error } = await sbClient.from("daily_reports").delete().eq("id", id).eq("user_id", currentUser.id);
      if (error) throw error;
      console.log("DB success", "日報削除");
      if (editState.dailyReportId === id) resetDailyReportForm();
      await refreshAfterMutation("日報を削除しました。", "日報削除");
    });
  }
}

async function startDailyReportEdit(dailyReportId) {
  await withLoading("編集ボタン処理", async () => {
    const target = state.dailyReports.find((entry) => entry.id === dailyReportId);
    if (!target) return;
    subtabState["daily-reports"] = "entry";
    activateTab("daily-reports");
    editState.dailyReportId = target.id;
    dailyReportForm.elements.reportDate.value = target.reportDate || "";
    reportCaseSelect.value = target.caseId || "";
    dailyReportForm.elements.reportWorkContent.value = target.workContent || "";
    dailyReportForm.elements.reportWorkMinutes.value = target.workMinutes ?? 0;
    dailyReportForm.elements.reportNextAction.value = target.nextAction || "";
    dailyReportForm.elements.reportMemo.value = target.memo || "";
    dailyReportSubmitBtn.textContent = "日報を更新";
    dailyReportForm.scrollIntoView({ behavior: "smooth", block: "start" });
  });
}

async function editSale(saleId) {
  await startSaleEdit(saleId);
}

async function editExpense(expenseId) {
  await startExpenseEdit(expenseId);
}

async function editDailyReport(dailyReportId) {
  await startDailyReportEdit(dailyReportId);
}

window.editSale = editSale;
window.editExpense = editExpense;
window.editDailyReport = editDailyReport;

async function handleClearAll() {
  if (!currentUser) return;
  if (!window.confirm("顧客・見積・案件・売上・経費・固定費・日報の全データを削除します。よろしいですか？")) return;

  await withLoading("全件削除", async () => {
    const [estimateItemsRes, estimatesRes, salesRes, expensesRes, fixedExpensesRes, dailyReportsRes, casesRes, clientsRes] = await Promise.all([
      sbClient.from("estimate_items").delete().eq("user_id", currentUser.id),
      sbClient.from("estimates").delete().eq("user_id", currentUser.id),
      sbClient.from("sales").delete().eq("user_id", currentUser.id),
      sbClient.from("expenses").delete().eq("user_id", currentUser.id),
      sbClient.from("fixed_expenses").delete().eq("user_id", currentUser.id),
      sbClient.from("daily_reports").delete().eq("user_id", currentUser.id),
      sbClient.from("cases").delete().eq("user_id", currentUser.id),
      sbClient.from("clients").delete().eq("user_id", currentUser.id),
    ]);

    if (salesRes.error) throw salesRes.error;
    if (expensesRes.error) throw expensesRes.error;
    if (fixedExpensesRes.error) throw fixedExpensesRes.error;
    if (dailyReportsRes.error) throw dailyReportsRes.error;
    if (casesRes.error) throw casesRes.error;
    if (clientsRes.error) throw clientsRes.error;

    console.log("DB success", "全件削除");
    resetClientForm();
    resetCaseForm();
    resetEstimateForm();
    resetSaleForm();
    resetExpenseForm();
    resetFixedExpenseForm();
    resetDailyReportForm();
    await refreshAfterMutation("全データを削除しました。", "全件削除");
  });
}

function handleExportCasesCsv() {
  const rows = state.cases.map((entry) => ({
    client_id: entry.clientId || "",
    customer_name: entry.customerName,
    case_name: entry.caseName,
    estimate_amount: entry.estimateAmount ?? "",
    received_date: entry.receivedDate || "",
    due_date: entry.dueDate || "",
    status: normalizeStoredStatus(entry.status),
    work_memo: sanitizeLegacyEstimateMemo(entry.workMemo) || "",
    next_action_date: entry.nextActionDate || "",
    next_action: entry.nextAction || "",
    document_url: entry.documentUrl || "",
    invoice_url: entry.invoiceUrl || "",
    receipt_url: entry.receiptUrl || "",
  }));
  downloadCsvFile("cases.csv", ["client_id", "customer_name", "case_name", "estimate_amount", "received_date", "due_date", "status", "work_memo", "next_action_date", "next_action", "document_url", "invoice_url", "receipt_url"], rows);
}

function handleExportSalesCsv() {
  const rows = state.sales.map((entry) => {
    const foundCase = state.cases.find((c) => c.id === entry.caseId);
    return {
      case_name: foundCase?.caseName || "",
      invoice_number: entry.invoiceNumber || "",
      invoice_amount: entry.invoiceAmount ?? "",
      paid_amount: entry.paidAmount ?? "",
      paid_date: entry.paidDate || "",
      is_unpaid: entry.isUnpaid ? "true" : "false",
    };
  });
  downloadCsvFile("sales.csv", ["case_name", "invoice_number", "invoice_amount", "paid_amount", "paid_date", "is_unpaid"], rows);
}

function handleExportExpensesCsv() {
  const rows = state.expenses.map((entry) => {
    const foundCase = state.cases.find((c) => c.id === entry.caseId);
    return {
      case_name: foundCase?.caseName || "",
      date: entry.date || "",
      content: entry.content || "",
      amount: entry.amount ?? "",
      payee: entry.payee || "",
      payment_method: entry.paymentMethod || "",
      receipt_url: entry.receiptUrl || "",
    };
  });
  downloadCsvFile("expenses.csv", ["case_name", "date", "content", "amount", "payee", "payment_method", "receipt_url"], rows);
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
    "client_name", "client_type", "address", "tel", "email", "referral_source", "client_memo",
    "client_id", "customer_name", "case_name", "estimate_amount", "received_date", "due_date", "status", "work_memo", "next_action_date", "next_action", "document_url", "invoice_url", "receipt_url",
    "estimate_number", "invoice_number", "invoice_amount", "paid_amount", "paid_date", "is_unpaid",
    "date", "content", "amount", "payee", "payment_method",
    "day_of_month", "start_date", "active",
  ];
  const rows = [];

  state.clients.forEach((entry) => {
    rows.push({
      data_type: "client",
      client_name: entry.name || "",
      client_type: entry.clientType || "",
      address: entry.address || "",
      tel: entry.tel || "",
      email: entry.email || "",
      referral_source: entry.referralSource || "",
      client_memo: entry.memo || "",
    });
  });

  state.cases.forEach((entry) => {
    rows.push({
      data_type: "case",
      client_id: entry.clientId || "",
      customer_name: entry.customerName,
      case_name: entry.caseName,
      estimate_amount: entry.estimateAmount ?? "",
      received_date: entry.receivedDate || "",
      due_date: entry.dueDate || "",
      status: normalizeStoredStatus(entry.status),
      work_memo: sanitizeLegacyEstimateMemo(entry.workMemo) || "",
      next_action_date: entry.nextActionDate || "",
      next_action: entry.nextAction || "",
      document_url: entry.documentUrl || "",
      invoice_url: entry.invoiceUrl || "",
      receipt_url: entry.receiptUrl || "",
    });
  });
  state.estimates.forEach((entry) => {
    rows.push({
      data_type: "estimate",
      client_id: entry.clientId || "",
      customer_name: entry.customerName || "",
      case_name: entry.estimateTitle || "",
      estimate_number: entry.estimateNumber || "",
      date: entry.estimateDate || "",
      status: entry.status || "",
      amount: entry.total ?? "",
    });
  });
  state.sales.forEach((entry) => {
    const foundCase = state.cases.find((c) => c.id === entry.caseId);
    rows.push({
      data_type: "sale",
      case_name: foundCase?.caseName || "",
      invoice_number: entry.invoiceNumber || "",
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
      payee: entry.payee || "",
      payment_method: entry.paymentMethod || "",
      receipt_url: entry.receiptUrl || "",
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

  await withLoading("CSV取込", async () => {
    const text = await readCsvFileTextWithEncodingFallback(file);
    const result = await importCsvByType(importType, text);
    console.log("DB success", "CSV取込");
    await loadAllData();
    console.log("loadAllData done", "CSV取込");
    renderAfterDataChanged();
    console.log("render done", "CSV取込");
    showAppMessage(`CSV取込完了: 登録件数 ${result.insertedCount}件 / スキップ件数 ${result.skippedCount}件 / エラー件数 ${result.errorCount}件`, result.errorCount > 0);
    csvImportForm.reset();
  }, { triggerButton: event.submitter });
}

function handleExportExcel() {
  withLoading("帳票出力", async () => {
    exportExcel();
  }).catch(() => {});
}

function exportExcel() {
  if (!window.XLSX) {
    showAppMessage("Excel出力ライブラリの読み込みに失敗しました", true);
    return;
  }

  try {
    const workbook = XLSX.utils.book_new();
    const clientHeaders = ["name", "client_type", "address", "tel", "email", "referral_source", "memo"];
    const caseHeaders = ["client_id", "customer_name", "case_name", "estimate_amount", "received_date", "due_date", "status", "work_memo", "next_action_date", "next_action", "document_url", "invoice_url", "receipt_url"];
    const saleHeaders = ["case_name", "invoice_number", "invoice_amount", "paid_amount", "paid_date", "is_unpaid"];
    const expenseHeaders = ["case_name", "date", "content", "amount", "payee", "payment_method", "receipt_url"];
    const fixedExpenseHeaders = ["content", "amount", "day_of_month", "start_date", "active"];
    const clientRows = state.clients.map((entry) => ({
      name: entry.name || "",
      client_type: entry.clientType || "",
      address: entry.address || "",
      tel: entry.tel || "",
      email: entry.email || "",
      referral_source: entry.referralSource || "",
      memo: entry.memo || "",
    }));
    const caseRows = state.cases.map((entry) => ({
      client_id: entry.clientId || "",
      customer_name: entry.customerName,
      case_name: entry.caseName,
      estimate_amount: entry.estimateAmount ?? "",
      received_date: entry.receivedDate || "",
      due_date: entry.dueDate || "",
      status: normalizeStoredStatus(entry.status),
      work_memo: sanitizeLegacyEstimateMemo(entry.workMemo) || "",
      next_action_date: entry.nextActionDate || "",
      next_action: entry.nextAction || "",
      document_url: entry.documentUrl || "",
      invoice_url: entry.invoiceUrl || "",
      receipt_url: entry.receiptUrl || "",
    }));
    const saleRows = state.sales.map((entry) => {
      const foundCase = state.cases.find((c) => c.id === entry.caseId);
      return {
        case_name: foundCase?.caseName || "",
        invoice_number: entry.invoiceNumber || "",
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
        payee: entry.payee || "",
        payment_method: entry.paymentMethod || "",
        receipt_url: entry.receiptUrl || "",
      };
    });
    const fixedExpenseRows = state.fixedExpenses.map((entry) => ({
      content: entry.content,
      amount: entry.amount ?? "",
      day_of_month: entry.dayOfMonth ?? "",
      start_date: entry.startDate || "",
      active: entry.active ? "true" : "false",
    }));

    XLSX.utils.book_append_sheet(workbook, createExcelSheet(clientRows, clientHeaders), "顧客");
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

  await withLoading("Excel取込", async () => {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array", cellDates: false });
    const result = await importWorkbookBySheet(workbook);
    console.log("DB success", "Excel取込");
    await loadAllData();
    console.log("loadAllData done", "Excel取込");
    renderAfterDataChanged();
    console.log("render done", "Excel取込");
    showAppMessage(`Excel取込完了: 登録件数 ${result.insertedCount}件 / スキップ件数 ${result.skippedCount}件 / エラー件数 ${result.errorCount}件`, result.errorCount > 0);
    excelImportForm.reset();
  }, { triggerButton: event.submitter });
}

function renderAfterDataChanged() {
  renderClients();
  renderClientOptions();
  renderCaseOptions();
  renderEstimates();
  renderCases();
  renderSales();
  renderExpenses();
  renderFixedExpenses();
  renderDailyReports();
  renderDashboard();
  renderTodayTaskCard();
}

function activateTab(tabKey) {
  const normalizedTabKey = normalizeTabKey(tabKey);

  tabs.forEach((btn) => btn.classList.toggle("active", normalizeTabKey(btn.dataset.tab) === normalizedTabKey));
  Object.entries(panels).forEach(([key, panel]) => panel.classList.toggle("active", key === normalizedTabKey));

  if (dashboardSection) dashboardSection.hidden = normalizedTabKey !== "cases";
  applySubtabVisibility(normalizedTabKey);
}

function activateSubtab(parentTab, subtab) {
  const normalizedTab = normalizeTabKey(parentTab);
  if (!subtab) return;
  subtabState[normalizedTab] = subtab;
  applySubtabVisibility(normalizedTab);
}

function applySubtabVisibility(activeMainTab) {
  const normalizedTab = normalizeTabKey(activeMainTab);
  Object.keys(subtabState).forEach((tabKey) => {
    const activeSubtab = subtabState[tabKey];
    subtabButtons
      .filter((btn) => normalizeTabKey(btn.dataset.parentTab) === tabKey)
      .forEach((btn) => {
        btn.classList.toggle("active", tabKey === normalizedTab && btn.dataset.subtab === activeSubtab);
      });
    subtabPanels
      .filter((panel) => normalizeTabKey(panel.dataset.parentTab) === tabKey)
      .forEach((panel) => {
        const isCasesDashboardPanel = tabKey === "cases" && panel.id === "cases-dashboard-panel";
        const shouldShow = tabKey === normalizedTab && (panel.dataset.subtab === activeSubtab || (isCasesDashboardPanel && activeSubtab === "alerts"));
        panel.hidden = !shouldShow;
      });
  });
  if (dashboardSection && normalizedTab === "cases") {
    const isAlertMode = subtabState.cases === "alerts";
    dashboardSection.classList.toggle("alerts-only", isAlertMode);
  }
}

function normalizeTabKey(tabKey) {
  const key = String(tabKey || "").trim();
  if (!key) return "cases";
  if (["daily-reports", "dailyReports", "report", "reports", "日報"].includes(key)) return "daily-reports";
  if (["clients", "client", "顧客"].includes(key)) return "clients";
  if (["estimates", "estimate", "見積", "見積もり"].includes(key)) return "estimates";
  if (["cases", "案件"].includes(key)) return "cases";
  if (["sales", "売上"].includes(key)) return "sales";
  if (["expenses", "経費"].includes(key)) return "expenses";
  return "cases";
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
  const reportByMonth = aggregateByMonth(state.dailyReports, (r) => r.reportDate, (r) => r.workMinutes);
  const isYearMode = state.selectedAggregation === "year";

  const sales = isYearMode ? sumYearFromMonthlyMap(salesByMonth, state.selectedYear) : salesByMonth[state.selectedMonth] || 0;
  const expenses = isYearMode ? sumYearFromMonthlyMap(expenseByMonth, state.selectedYear) : expenseByMonth[state.selectedMonth] || 0;
  const workMinutes = isYearMode ? sumYearFromMonthlyMap(reportByMonth, state.selectedYear) : reportByMonth[state.selectedMonth] || 0;
  const profit = sales - expenses;

  summaryGrid.innerHTML = "";
  const labelPrefix = isYearMode ? "年別" : "月別";
  const targetLabel = isYearMode ? `${state.selectedYear}年` : monthLabel(state.selectedMonth);
  [
    { label: isYearMode ? "対象年" : "対象月", value: targetLabel, cls: "" },
    { label: `${labelPrefix}売上合計`, value: formatCurrency(sales), cls: "" },
    { label: `${labelPrefix}経費合計`, value: formatCurrency(expenses), cls: "" },
    { label: `${labelPrefix}作業時間`, value: formatMinutes(workMinutes), cls: "" },
    { label: "利益（売上−経費）", value: formatCurrency(profit), cls: `profit ${profit < 0 ? "loss" : ""}`.trim() },
  ].forEach((card) => {
    const div = document.createElement("div");
    div.className = `summary-card ${card.cls}`.trim();
    div.innerHTML = `<p class="label">${card.label}</p><p class="value">${card.value}</p>`;
    summaryGrid.appendChild(div);
  });
  renderTodayTaskCard();
  renderNextActionAlertCard();
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
  renderDashboardWorktimeSummary();
  renderUnpaidList();
}

function renderTodayTaskCard() {
  if (!todayTaskCard || !todayTaskSummary || !todayTaskBody || !todayTaskEmpty || !todayTaskListWrap) return;
  const rows = buildTodayTasks();
  todayTaskCard.classList.toggle("has-alert", rows.length > 0);
  todayTaskSummary.textContent = `対象 ${rows.length}件`;
  todayTaskBody.innerHTML = "";
  todayTaskEmpty.hidden = rows.length > 0;
  todayTaskListWrap.hidden = rows.length === 0;
  if (!rows.length) return;
  rows
    .sort((a, b) => toSortTimestamp(a.date) - toSortTimestamp(b.date))
    .forEach((row) => {
      const tr = document.createElement("tr");
      tr.className = row.urgencyClass;
      tr.innerHTML = `
        <td>${escapeHtml(row.type)}</td>
        <td>${escapeHtml(row.customerName)}</td>
        <td>${escapeHtml(row.caseName)}</td>
        <td>${escapeHtml(row.action)}</td>
        <td>${formatDate(row.date)}</td>
        <td><button type="button" class="secondary-btn" data-task-target="${row.target}" data-task-id="${row.id}">編集</button></td>
      `;
      todayTaskBody.appendChild(tr);
    });
}

function buildTodayTasks() {
  const tasks = [];
  const deadlineTargets = getDeadlineAlertTargets();
  const nextTargets = getNextActionAlertTargets();
  const todayLimit = getTodayTimestamp();
  const weekLimit = todayLimit + (7 * 86400000);
  nextTargets.forEach((entry) => {
    tasks.push({
      id: entry.id, target: "case", type: "次回対応", customerName: entry.customerName, caseName: entry.caseName,
      action: entry.nextAction || "次回対応", date: entry.nextActionDate, urgencyClass: "task-overdue",
    });
  });
  deadlineTargets.forEach((entry) => {
    const dueTs = toDateOnlyTimestamp(entry.dueDate);
    if (!(dueTs <= weekLimit)) return;
    tasks.push({
      id: entry.id, target: "case", type: "期限", customerName: entry.customerName, caseName: entry.caseName,
      action: `期限対応（${entry.status}）`, date: entry.dueDate, urgencyClass: dueTs <= todayLimit ? "task-overdue" : "task-soon",
    });
  });
  getUnpaidSales().forEach((sale) => {
    const linked = state.cases.find((entry) => entry.id === sale.caseId);
    tasks.push({
      id: sale.id, target: "sale", type: "未入金", customerName: linked?.customerName || "（削除済み顧客）", caseName: linked?.caseName || "（削除済み案件）",
      action: `未入金 ${formatCurrency(getUnpaidAmount(sale))}`, date: sale.paidDate || toDateString(sale.createdAt), urgencyClass: "task-overdue",
    });
  });
  getBillingLeakCandidates().forEach((entry) => {
    tasks.push({
      id: entry.id, target: "case", type: "請求漏れ", customerName: entry.customerName, caseName: entry.caseName,
      action: "売上登録が未実施", date: entry.updatedAt ? toDateString(entry.updatedAt) : toDateString(new Date()), urgencyClass: "task-soon",
    });
  });
  return tasks;
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


function renderNextActionAlertCard() {
  if (!nextActionAlertCard || !nextActionAlertSummary || !nextActionAlertBody || !nextActionAlertEmpty || !nextActionAlertListWrap) return;
  const targets = getNextActionAlertTargets();
  nextActionAlertCard.classList.toggle("has-alert", targets.length > 0);
  nextActionAlertSummary.textContent = `対象 ${targets.length}件`;
  nextActionAlertBody.innerHTML = "";
  nextActionAlertEmpty.hidden = targets.length > 0;
  nextActionAlertListWrap.hidden = targets.length === 0;
  if (!targets.length) return;

  targets
    .slice()
    .sort((a, b) => {
      if (a.remainingDays !== b.remainingDays) return a.remainingDays - b.remainingDays;
      return sortCases(a, b);
    })
    .forEach((entry) => {
      const tr = document.createElement("tr");
      tr.className = entry.urgencyClass;
      tr.innerHTML = `
        <td>${escapeHtml(entry.customerName)}</td>
        <td>${escapeHtml(entry.caseName)}</td>
        <td>${formatDate(entry.nextActionDate)}</td>
        <td>${escapeHtml(entry.nextAction || "未設定")}</td>
        <td><button type="button" class="secondary-btn" data-case-id="${entry.id}">編集</button></td>
      `;
      nextActionAlertBody.appendChild(tr);
    });
}

function getNextActionAlertTargets() {
  return state.cases
    .map((entry) => {
      const info = getNextActionInfo(entry);
      if (!info) return null;
      return { ...entry, ...info };
    })
    .filter((entry) => entry && entry.remainingDays <= 0 && getStatusCategory(entry.status) !== "完了");
}

function getNextActionInfo(entry) {
  if (!entry?.nextActionDate) return null;
  const nextActionTimestamp = toDateOnlyTimestamp(entry.nextActionDate);
  const todayTimestamp = getTodayTimestamp();
  if (!Number.isFinite(nextActionTimestamp) || !Number.isFinite(todayTimestamp)) return null;
  const remainingDays = Math.floor((nextActionTimestamp - todayTimestamp) / 86400000);
  const urgencyClass = remainingDays <= 0 ? "next-action-overdue" : remainingDays <= 3 ? "next-action-within3" : remainingDays <= 7 ? "next-action-within7" : "";
  return { remainingDays, urgencyClass };
}

function truncateText(value, maxLength = 80) {
  const text = String(value || "").trim();
  if (text.length <= maxLength) return text;
  return `${text.slice(0, maxLength)}…`;
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
  reportCaseSelect.innerHTML = `<option value="">案件なし</option>${options}`;
}

function renderClientOptions() {
  const options = state.clients
    .slice()
    .sort((a, b) => (b.updatedAt ?? b.createdAt) - (a.updatedAt ?? a.createdAt))
    .map((client) => `<option value="${client.id}">${escapeHtml(client.name)}${client.clientType ? `（${escapeHtml(client.clientType)}）` : ""}</option>`)
    .join("");
  if (caseClientSelect) caseClientSelect.innerHTML = `<option value="">選択しない</option>${options}`;
  if (estimateClientSelect) estimateClientSelect.innerHTML = `<option value="">選択しない</option>${options}`;
}

function syncCaseCustomerFromClient() {
  const clientId = caseClientSelect?.value;
  if (!clientId || !caseForm?.elements?.customerName) return;
  const found = state.clients.find((entry) => entry.id === clientId);
  if (found) caseForm.elements.customerName.value = found.name;
}

function syncEstimateCustomerFromClient() {
  const clientId = estimateClientSelect?.value;
  if (!clientId || !estimateForm?.elements?.customerName) return;
  const found = state.clients.find((entry) => entry.id === clientId);
  if (found) estimateForm.elements.customerName.value = found.name;
}

function renderClients() {
  if (!clientsList || !clientsEmpty) return;
  clientsList.innerHTML = "";
  const sorted = state.clients.slice().sort((a, b) => (b.updatedAt ?? b.createdAt) - (a.updatedAt ?? a.createdAt));
  sorted.forEach((client) => {
    const li = document.createElement("li");
    li.className = "item";
    li.dataset.id = client.id;
    li.innerHTML = `
      <div>
        <p class="title">${escapeHtml(client.name)}${client.clientType ? `（${escapeHtml(client.clientType)}）` : ""}</p>
        <p class="meta">住所: ${escapeHtml(client.address || "未設定")} / 電話: ${escapeHtml(client.tel || "未設定")} / メール: ${escapeHtml(client.email || "未設定")}</p>
        <p class="meta">紹介元: ${escapeHtml(client.referralSource || "未設定")} / メモ: ${escapeHtml(truncateText(client.memo || "未設定", 60))}</p>
      </div>
      <div class="row-actions"><button type="button" class="secondary-btn edit-client-btn">編集</button><button type="button" class="danger-btn delete-client-btn">削除</button></div>
    `;
    clientsList.appendChild(li);
  });
  clientsEmpty.hidden = sorted.length > 0;
}

function renderDailyReports() {
  if (!dailyReportsBody || !dailyReportsEmpty || !dailyReportsListWrap || !dailyReportSummaryList) return;
  const sorted = state.dailyReports.slice().sort((a, b) => toSortTimestamp(b.reportDate) - toSortTimestamp(a.reportDate));
  const filtered = sorted.filter((entry) => {
    if (!matchesDailyReportDateFilter(entry, state.dailyReportDateFilter)) return false;
    if (!state.dailyReportSearchQuery) return true;
    return matchesDailyReportSearch(entry, state.dailyReportSearchQuery);
  });

  dailyReportsBody.innerHTML = "";

  filtered.forEach((entry) => {
    const card = document.createElement("article");
    card.className = "daily-report-card";
    card.innerHTML = `
      <header class="daily-report-card-header">
        <div>
          <p class="daily-report-card-date">${formatDate(entry.reportDate)}</p>
          <p class="daily-report-card-case">${escapeHtml(resolveDailyReportCaseName(entry.caseId))}</p>
        </div>
        <div class="daily-report-card-actions">
          <button type="button" class="secondary-btn edit-daily-report-btn" data-daily-report-id="${entry.id}">編集</button>
          <button type="button" class="danger-btn delete-daily-report-btn" data-daily-report-id="${entry.id}">削除</button>
        </div>
      </header>
      <dl class="daily-report-card-meta">
        <div class="daily-report-field-inline">
          <dt>作業時間</dt>
          <dd>${formatMinutes(entry.workMinutes)}</dd>
        </div>
        <div class="daily-report-field">
          <dt>作業内容</dt>
          <dd>${escapeHtml(entry.workContent || "未設定")}</dd>
        </div>
        <div class="daily-report-field">
          <dt>次回対応</dt>
          <dd>${escapeHtml(entry.nextAction || "未設定")}</dd>
        </div>
        <div class="daily-report-field">
          <dt>メモ</dt>
          <dd>${escapeHtml(entry.memo || "未設定")}</dd>
        </div>
      </dl>
    `;
    dailyReportsBody.appendChild(card);
  });

  dailyReportsEmpty.hidden = filtered.length > 0;
  dailyReportsListWrap.hidden = filtered.length === 0;
  dailyReportsEmpty.textContent = filtered.length || (!state.dailyReportSearchQuery && state.dailyReportDateFilter === "all")
    ? "日報データはまだありません。"
    : "条件に一致する日報はありません。";

  if (dailyReportFilterCount) {
    dailyReportFilterCount.textContent = `表示中 ${filtered.length}件 / 全${sorted.length}件`;
  }

  const summary = buildDailyReportSummary();
  dailyReportSummaryList.innerHTML = "";
  [
    { title: "今日の作業時間合計", value: formatMinutes(summary.todayMinutes) },
    { title: "今月の作業時間合計", value: formatMinutes(summary.monthMinutes) },
    { title: "案件別作業時間合計", value: summary.caseRows || "データなし" },
  ].forEach((entry) => {
    const li = document.createElement("li");
    li.className = "item";
    li.innerHTML = `<div><p class="title">${entry.title}</p><p class="meta">${escapeHtml(entry.value)}</p></div>`;
    dailyReportSummaryList.appendChild(li);
  });
}

function matchesDailyReportSearch(entry, query) {
  const haystacks = [
    resolveDailyReportCaseName(entry.caseId),
    entry.workContent,
    entry.nextAction,
    entry.memo,
  ]
    .map((value) => String(value || "").toLowerCase());
  return haystacks.some((value) => value.includes(query));
}

function matchesDailyReportDateFilter(entry, filter) {
  if (filter === "all") return true;
  const reportDate = entry.reportDate;
  if (!reportDate) return false;
  const today = toDateString(new Date());
  if (filter === "today") return reportDate === today;
  if (filter === "month") return toMonthKey(reportDate) === toMonthKey(today);
  return true;
}

function addEstimateItemRow(defaultItem = {}) {
  if (!estimateItemsWrap) return;
  const row = document.createElement("div");
  row.className = "estimate-item-row";
  row.innerHTML = `
    <input type="text" data-key="itemName" placeholder="項目名" value="${escapeHtml(defaultItem.itemName || "")}" />
    <input type="number" data-key="quantity" min="0" step="0.01" placeholder="数量" value="${defaultItem.quantity ?? 1}" />
    <input type="number" data-key="unitPrice" min="0" step="1" placeholder="単価" value="${defaultItem.unitPrice ?? 0}" />
    <p class="meta item-amount">${formatCurrency(defaultItem.amount ?? 0)}</p>
    <button type="button" class="danger-btn estimate-item-remove-btn">削除</button>
  `;
  estimateItemsWrap.appendChild(row);
  recalcEstimateTotals();
}

function handleEstimateItemsInput() {
  recalcEstimateTotals();
}

function handleEstimateItemsClick(event) {
  const btn = event.target.closest(".estimate-item-remove-btn");
  if (!btn) return;
  btn.closest(".estimate-item-row")?.remove();
  if (!estimateItemsWrap.children.length) addEstimateItemRow();
  recalcEstimateTotals();
}

function getEstimateItemsFromForm() {
  if (!estimateItemsWrap) return [];
  return Array.from(estimateItemsWrap.querySelectorAll(".estimate-item-row"))
    .map((row, idx) => {
      const itemName = asTrimmedText(row.querySelector('[data-key="itemName"]')?.value);
      const quantity = Number(row.querySelector('[data-key="quantity"]')?.value || 0);
      const unitPrice = normalizeAmount(row.querySelector('[data-key="unitPrice"]')?.value) ?? 0;
      const amount = Math.floor((Number.isFinite(quantity) ? quantity : 0) * unitPrice);
      return { itemName, quantity, unitPrice, amount, sortOrder: idx };
    })
    .filter((item) => item.itemName);
}

function recalcEstimateTotals() {
  let subtotal = 0;
  Array.from(estimateItemsWrap?.querySelectorAll(".estimate-item-row") || []).forEach((row) => {
    const quantity = Number(row.querySelector('[data-key="quantity"]')?.value || 0);
    const unitPrice = normalizeAmount(row.querySelector('[data-key="unitPrice"]')?.value) ?? 0;
    const amount = Math.floor((Number.isFinite(quantity) ? quantity : 0) * unitPrice);
    subtotal += amount;
    const amountEl = row.querySelector(".item-amount");
    if (amountEl) amountEl.textContent = formatCurrency(amount);
  });
  const tax = Math.floor(subtotal * 0.1);
  const total = subtotal + tax;
  if (estimateSubtotal) estimateSubtotal.textContent = formatCurrency(subtotal);
  if (estimateTax) estimateTax.textContent = formatCurrency(tax);
  if (estimateTotal) estimateTotal.textContent = formatCurrency(total);
  return { subtotal, tax, total };
}

async function handleEstimateSubmit(event) {
  event.preventDefault();
  if (!currentUser) return;
  const customerName = asTrimmedText(estimateForm.elements.customerName.value);
  const estimateTitle = asTrimmedText(estimateForm.elements.estimateTitle.value);
  const estimateDate = estimateForm.elements.estimateDate.value;
  if (!customerName || !estimateTitle || !estimateDate) return;
  const totals = recalcEstimateTotals();
  const items = getEstimateItemsFromForm();
  const payload = {
    user_id: currentUser.id,
    client_id: estimateForm.elements.clientId.value || null,
    customer_name: customerName,
    estimate_title: estimateTitle,
    estimate_date: estimateDate,
    valid_until: estimateForm.elements.validUntil.value || null,
    status: normalizeEstimateStatus(estimateForm.elements.status.value),
    memo: asTrimmedText(estimateForm.elements.memo.value) || null,
    subtotal: totals.subtotal,
    tax: totals.tax,
    total: totals.total,
  };

  const taskName = editState.estimateId ? "見積更新" : "見積登録";
  await withLoading(taskName, async () => {
    clearAppMessage();
    let estimateId = editState.estimateId;
    const isUpdate = Boolean(estimateId);
    if (estimateId) {
      const { data, error } = await sbClient.from("estimates").update(payload).eq("id", estimateId).eq("user_id", currentUser.id).select().single();
      if (error) throw error;
      if (!data) throw new Error("更新結果を取得できませんでした。");
      const oldItemsDeleteRes = await sbClient.from("estimate_items").delete().eq("estimate_id", estimateId).eq("user_id", currentUser.id);
      if (oldItemsDeleteRes.error) throw oldItemsDeleteRes.error;
    } else {
      payload.estimate_number = await getNextMonthlyNumber("estimates", "estimate_number", "M", estimateDate);
      const res = await sbClient.from("estimates").insert(payload).select("id").single();
      if (res.error) throw res.error;
      if (!res.data?.id) throw new Error("登録結果を取得できませんでした。");
      estimateId = res.data.id;
    }
    if (items.length) {
      const { data, error } = await sbClient.from("estimate_items").insert(items.map((item) => ({
        user_id: currentUser.id,
        estimate_id: estimateId,
        item_name: item.itemName,
        quantity: item.quantity,
        unit_price: item.unitPrice,
        amount: item.amount,
        sort_order: item.sortOrder,
      }))).select("id");
      if (error) throw error;
      if (!data) throw new Error("明細登録結果を取得できませんでした。");
    }
    if (payload.status === "受注") await ensureCaseFromEstimate(estimateId);
    console.log("DB success", taskName);
    await refreshAfterMutation(isUpdate ? "見積を更新しました。" : "見積を登録しました。", taskName);
    resetEstimateForm();
    editState.estimateId = null;
    subtabState.estimates = "list";
    activateTab("estimates");
  }, { triggerButton: event.submitter });
}

function renderEstimates() {
  if (!estimateList || !estimateEmpty) return;
  estimateList.innerHTML = "";
  const filtered = state.estimates
    .slice()
    .sort((a, b) => toSortTimestamp(b.estimateDate) - toSortTimestamp(a.estimateDate))
    .filter((entry) => {
      if (state.estimateCustomerQuery && !String(entry.customerName || "").toLowerCase().includes(state.estimateCustomerQuery)) return false;
      if (state.estimateTitleQuery && !String(entry.estimateTitle || "").toLowerCase().includes(state.estimateTitleQuery)) return false;
      if (state.estimateStatusFilter !== "all" && entry.status !== state.estimateStatusFilter) return false;
      if (state.estimateExpiredFilter === "expired" && !isEstimateExpired(entry)) return false;
      return true;
    });
  filtered.forEach((entry) => {
    const li = document.createElement("li");
    li.className = "item estimate-card";
    li.dataset.id = entry.id;
    const itemCount = state.estimateItems.filter((row) => row.estimateId === entry.id).length;
    const casedLabel = entry.caseId ? "案件化済み" : "未案件化";
    li.innerHTML = `
      <div class="estimate-card-body">
        <p class="title estimate-card-title">${escapeHtml(entry.estimateTitle)}</p>
        <div class="estimate-card-grid">
          <p class="meta"><span>顧客名:</span> ${escapeHtml(entry.customerName)}</p>
          <p class="meta"><span>見積番号:</span> ${escapeHtml(entry.estimateNumber || "未採番")}</p>
          <p class="meta"><span>見積日:</span> ${formatDate(entry.estimateDate)}</p>
          <p class="meta"><span>有効期限:</span> ${formatDate(entry.validUntil)}</p>
          <p class="meta"><span>ステータス:</span> ${escapeHtml(entry.status)}</p>
          <p class="meta"><span>合計金額:</span> ${formatCurrency(entry.total)}</p>
          <p class="meta"><span>明細件数:</span> ${itemCount}件</p>
          <p class="meta"><span>案件化:</span> ${casedLabel}</p>
        </div>
      </div>
      <div class="row-actions estimate-card-actions">
        <button type="button" class="secondary-btn estimate-edit-btn">編集</button>
        <button type="button" class="danger-btn estimate-delete-btn">削除</button>
        <button type="button" class="secondary-btn estimate-case-btn" ${entry.caseId ? "disabled" : ""}>案件化</button>
        <button type="button" class="secondary-btn estimate-estimate-print-btn">見積書出力</button>
        <button type="button" class="secondary-btn estimate-print-btn">請求書出力</button>
        <button type="button" class="secondary-btn estimate-estimate-xlsx-btn">見積Excel</button>
        <button type="button" class="secondary-btn estimate-xlsx-btn">請求Excel</button>
        <button type="button" class="secondary-btn estimate-sale-btn">売上登録</button>
      </div>
    `;
    estimateList.appendChild(li);
  });
  estimateEmpty.hidden = filtered.length > 0;
}

function isEstimateExpired(entry) {
  if (!entry.validUntil) return false;
  return toDateOnlyTimestamp(entry.validUntil) < getTodayTimestamp();
}

async function handleEstimateListAction(event) {
  const btn = event.target.closest("button");
  if (!(btn instanceof HTMLButtonElement) || !currentUser) return;
  const item = btn.closest(".item");
  const estimateId = item?.dataset.id;
  if (!estimateId) return;
  if (btn.classList.contains("estimate-edit-btn")) return startEstimateEdit(estimateId);
  if (btn.classList.contains("estimate-delete-btn")) {
    if (!window.confirm("この見積を削除しますか？")) return;
    await withLoading("見積削除", async () => {
      await sbClient.from("estimate_items").delete().eq("estimate_id", estimateId).eq("user_id", currentUser.id);
      const { error } = await sbClient.from("estimates").delete().eq("id", estimateId).eq("user_id", currentUser.id);
      if (error) throw error;
      console.log("DB success", "見積削除");
      if (editState.estimateId === estimateId) resetEstimateForm();
      await refreshAfterMutation("見積を削除しました。", "見積削除");
    });
    return;
  }
  if (btn.classList.contains("estimate-case-btn")) return withLoading("見積案件化", async () => {
    const estimate = state.estimates.find((entry) => entry.id === estimateId);
    if (estimate?.caseId) {
      showAppMessage("この見積はすでに案件化済みです。", true);
      return;
    }
    await ensureCaseFromEstimate(estimateId, true);
    console.log("DB success", "見積案件化");
    await refreshAfterMutation("案件化しました。", "見積案件化");
  });
  if (btn.classList.contains("estimate-estimate-print-btn")) return withLoading("帳票出力", async () => openEstimatePrintPreview(estimateId));
  if (btn.classList.contains("estimate-print-btn")) return withLoading("帳票出力", async () => openInvoicePrintPreviewFromEstimate(estimateId));
  if (btn.classList.contains("estimate-estimate-xlsx-btn")) return withLoading("帳票出力", async () => exportEstimateDataForEstimate(estimateId));
  if (btn.classList.contains("estimate-xlsx-btn")) return withLoading("帳票出力", async () => exportInvoiceDataForEstimate(estimateId));
  if (btn.classList.contains("estimate-sale-btn")) return withLoading("見積から売上登録", async () => {
    await registerSaleFromEstimate(estimateId);
    console.log("DB success", "見積から売上登録");
    await refreshAfterMutation("売上を登録しました。", "見積から売上登録");
  });
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
    const caseWorkMeta = node.querySelector(".case-work-meta");
    const profitMeta = node.querySelector(".profit-meta");
    const rowActions = node.querySelector(".row-actions");
    const totals = profitsByCaseId[entry.id] || { sales: 0, expenses: 0, profit: 0 };
    const nextActionInfo = getNextActionInfo(entry);

    item.dataset.id = entry.id;
    title.textContent = `${entry.customerName}｜${entry.caseName}`;
    const urlLinks = [
      entry.documentUrl ? `<a href="${escapeHtml(entry.documentUrl)}" target="_blank" rel="noopener noreferrer">関連書類を開く</a>` : "",
      entry.invoiceUrl ? `<a href="${escapeHtml(entry.invoiceUrl)}" target="_blank" rel="noopener noreferrer">請求書を開く</a>` : "",
      entry.receiptUrl ? `<a href="${escapeHtml(entry.receiptUrl)}" target="_blank" rel="noopener noreferrer">領収書を開く</a>` : "",
    ].filter(Boolean).join(" / ");
    meta.innerHTML = `見積: ${formatCurrency(entry.estimateAmount)} / ステータス: ${escapeHtml(entry.status)} / 受付日: ${formatDate(entry.receivedDate)} / 期限日: ${formatDate(entry.dueDate)} / 次回対応日: ${formatDate(entry.nextActionDate)} / 次回対応内容: ${escapeHtml(entry.nextAction || "未設定")}${urlLinks ? ` / ${urlLinks}` : ""}`;
    if (caseWorkMeta) {
      caseWorkMeta.textContent = `作業メモ: ${truncateText(sanitizeLegacyEstimateMemo(entry.workMemo) || "未設定", 60)}`;
      caseWorkMeta.classList.remove("next-action-overdue", "next-action-within3", "next-action-within7");
      if (nextActionInfo) caseWorkMeta.classList.add(nextActionInfo.urgencyClass);
    }
    profitMeta.textContent = `売上合計: ${formatCurrency(totals.sales)} / 経費合計: ${formatCurrency(totals.expenses)} / 利益: ${formatCurrency(totals.profit)}`;
    profitMeta.classList.toggle("loss-text", totals.profit < 0);
    if (rowActions && !rowActions.querySelector(".case-print-btn")) {
      const printBtn = document.createElement("button");
      printBtn.type = "button";
      printBtn.className = "secondary-btn case-print-btn";
      printBtn.textContent = "請求書出力";
      rowActions.appendChild(printBtn);
    }
    if (rowActions && !rowActions.querySelector(".case-xlsx-btn")) {
      const xlsxBtn = document.createElement("button");
      xlsxBtn.type = "button";
      xlsxBtn.className = "secondary-btn case-xlsx-btn";
      xlsxBtn.textContent = "請求Excel";
      rowActions.appendChild(xlsxBtn);
    }
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
    const totals = profitsByCaseId[entry.id] || { sales: 0, expenses: 0, profit: 0, workMinutes: 0 };
    const li = document.createElement("li");
    li.className = "item";
    li.innerHTML = `
      <div>
        <p class="title">${escapeHtml(entry.customerName)}｜${escapeHtml(entry.caseName)}</p>
        <p class="meta">売上合計: ${formatCurrency(totals.sales)} / 経費合計: ${formatCurrency(totals.expenses)} / 作業時間: ${formatMinutes(totals.workMinutes)} / <span class="${totals.profit < 0 ? "loss-text" : ""}">利益: ${formatCurrency(totals.profit)}</span></p>
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
    map[entry.id] = { sales: 0, expenses: 0, profit: 0, workMinutes: 0 };
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

  state.dailyReports.forEach((report) => {
    if (!report.caseId || !map[report.caseId]) return;
    if (!isWithinFilterDate(report.reportDate, filter)) return;
    map[report.caseId].workMinutes += report.workMinutes || 0;
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
    title.textContent = `${resolveCaseName(sale.caseId)}｜請求: ${formatCurrency(sale.invoiceAmount)}｜請求番号: ${sale.invoiceNumber || "未採番"}`;
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
    const payeeLabel = expense.payee || "未設定";
    const paymentMethodLabel = expense.paymentMethod || "未設定";

    item.dataset.id = expense.id;
    title.textContent = `日付: ${formatDate(expense.date)}｜内容: ${expense.content}｜金額: ${formatCurrency(expense.amount)}`;
    meta.innerHTML = `支払先: ${escapeHtml(payeeLabel)} / 支払方法: ${escapeHtml(paymentMethodLabel)} / 紐付け案件: ${escapeHtml(expense.caseId ? resolveCaseName(expense.caseId) : "なし")}${expense.receiptUrl ? ` / <a href="${escapeHtml(expense.receiptUrl)}" target="_blank" rel="noopener noreferrer">領収書を開く</a>` : ""}`;
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

function resolveDailyReportCaseName(caseId) {
  if (!caseId) return "案件なし";
  const found = state.cases.find((c) => c.id === caseId);
  return found ? `${found.customerName}｜${found.caseName}` : "案件なし";
}

async function startEstimateEdit(estimateId) {
  await withLoading("編集ボタン処理", async () => {
    const target = state.estimates.find((entry) => entry.id === estimateId);
    if (!target || !estimateForm) return;
    subtabState.estimates = "create";
    activateTab("estimates");
    editState.estimateId = target.id;
    if (estimateForm.elements.clientId) estimateForm.elements.clientId.value = target.clientId || "";
    estimateForm.elements.customerName.value = target.customerName;
    estimateForm.elements.estimateTitle.value = target.estimateTitle;
    estimateForm.elements.estimateDate.value = target.estimateDate || "";
    estimateForm.elements.validUntil.value = target.validUntil || "";
    estimateForm.elements.status.value = normalizeEstimateStatus(target.status);
    estimateForm.elements.memo.value = target.memo || "";
    estimateItemsWrap.innerHTML = "";
    const rows = state.estimateItems.filter((item) => item.estimateId === target.id).sort((a, b) => a.sortOrder - b.sortOrder);
    if (!rows.length) addEstimateItemRow();
    rows.forEach((row) => addEstimateItemRow(row));
    recalcEstimateTotals();
    estimateSubmitBtn.textContent = "見積を更新";
    estimateForm.scrollIntoView({ behavior: "smooth", block: "start" });
  });
}

function resetEstimateForm() {
  resetEditMode("estimate");
  estimateForm?.reset();
  if (estimateForm?.elements?.clientId) estimateForm.elements.clientId.value = "";
  if (estimateForm?.elements?.estimateDate) estimateForm.elements.estimateDate.value = toDateString(new Date());
  if (estimateForm?.elements?.status) estimateForm.elements.status.value = "作成中";
  if (estimateItemsWrap) {
    estimateItemsWrap.innerHTML = "";
    addEstimateItemRow();
  }
  if (estimateSubmitBtn) estimateSubmitBtn.textContent = "見積を登録";
  recalcEstimateTotals();
}

function resetClientForm() {
  resetEditMode("client");
  clientForm?.reset();
  if (clientSubmitBtn) clientSubmitBtn.textContent = "顧客を登録";
}

function resetCaseForm() {
  resetEditMode("case");
  caseForm.reset();
  if (caseForm?.elements?.caseClientId) caseForm.elements.caseClientId.value = "";
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

function resetDailyReportForm() {
  resetEditMode("dailyReport");
  dailyReportForm.reset();
  dailyReportForm.elements.reportDate.value = toDateString(new Date());
  dailyReportSubmitBtn.textContent = "日報を登録";
}

function resetEditMode(target) {
  if (target === "client") editState.clientId = null;
  if (target === "case") editState.caseId = null;
  if (target === "sale") editState.saleId = null;
  if (target === "expense") editState.expenseId = null;
  if (target === "fixedExpense") editState.fixedExpenseId = null;
  if (target === "dailyReport") editState.dailyReportId = null;
  if (target === "estimate") editState.estimateId = null;
}

function normalizeEstimateStatus(status) {
  return ESTIMATE_STATUS_ORDER.includes(status) ? status : "作成中";
}

async function ensureCaseFromEstimate(estimateId, force = false) {
  const estimate = state.estimates.find((entry) => entry.id === estimateId);
  if (!estimate || (!force && estimate.status !== "受注")) return null;
  if (estimate.caseId) return estimate.caseId;
  const payload = {
    user_id: currentUser.id,
    client_id: estimate.clientId || null,
    customer_name: estimate.customerName,
    case_name: estimate.estimateTitle,
    estimate_amount: estimate.total,
    received_date: toDateString(new Date()),
    due_date: null,
    status: "未着手",
    work_memo: asTrimmedText(estimate.memo) || "",
    next_action_date: null,
    next_action: "",
  };
  const res = await sbClient.from("cases").insert(payload).select("id").single();
  if (res.error) throw res.error;
  if (!res.data?.id) throw new Error("登録結果を取得できませんでした。");
  const updateRes = await sbClient.from("estimates").update({ case_id: res.data.id }).eq("id", estimateId).eq("user_id", currentUser.id);
  if (updateRes.error) throw updateRes.error;
  return res.data.id;
}

function buildInvoiceRowsFromEstimate(estimate) {
  const items = state.estimateItems
    .filter((row) => row.estimateId === estimate.id)
    .sort((a, b) => a.sortOrder - b.sortOrder);
  if (!items.length) {
    return [{
      customer_name: estimate.customerName,
      subject: estimate.estimateTitle,
      invoice_date: toDateString(new Date()),
      item_name: estimate.estimateTitle,
      quantity: 1,
      unit_price: estimate.subtotal,
      amount: estimate.subtotal,
      subtotal: estimate.subtotal,
      tax: estimate.tax,
      total: estimate.total,
    }];
  }
  return items.map((item) => ({
    customer_name: estimate.customerName,
    subject: estimate.estimateTitle,
    invoice_date: toDateString(new Date()),
    item_name: item.itemName,
    quantity: item.quantity,
    unit_price: item.unitPrice,
    amount: item.amount,
    subtotal: estimate.subtotal,
    tax: estimate.tax,
    total: estimate.total,
  }));
}

function buildInvoiceDocumentFromEstimate(estimate, noteOverride = null) {
  const rows = buildInvoiceRowsFromEstimate(estimate);
  const invoiceDate = toDateString(new Date());
  const linkedSale = estimate.caseId
    ? state.sales.find((sale) => sale.caseId === estimate.caseId)
    : null;
  return {
    customerName: estimate.customerName || "顧客名未設定",
    subject: estimate.estimateTitle || "請求内容",
    invoiceDate,
    invoiceNumber: linkedSale?.invoiceNumber || "未採番",
    companyName: OFFICE_INFO.name,
    companyZip: OFFICE_INFO.zip,
    companyAddress: OFFICE_INFO.address,
    companyPhone: OFFICE_INFO.tel,
    companyEmail: OFFICE_INFO.email,
    registrationNumber: OFFICE_INFO.registrationNumber,
    details: rows.map((row, index) => ({
      no: index + 1,
      itemName: row.item_name,
      quantity: row.quantity,
      unitPrice: row.unit_price,
      amount: row.amount,
    })),
    subtotal: estimate.subtotal ?? 0,
    tax: estimate.tax ?? 0,
    total: estimate.total ?? 0,
    transferInfo: OFFICE_INFO.transferInfo,
    note: asTrimmedText(noteOverride ?? "") || estimate.memo || OFFICE_INFO.invoiceNote || "",
  };
}

function exportInvoiceDataForEstimate(estimateId) {
  const estimate = state.estimates.find((entry) => entry.id === estimateId);
  if (!estimate) return;
  const note = requestInvoiceMemo(estimate.memo || "");
  if (note === null) return;
  const invoiceData = buildInvoiceDocumentFromEstimate(estimate, note);
  downloadInvoiceWorkbook(invoiceData);
}

function buildInvoiceDocumentFromCase(foundCase, noteOverride = null) {
  const subtotal = foundCase.estimateAmount ?? 0;
  const tax = Math.floor(subtotal * 0.1);
  const total = subtotal + tax;
  const invoiceDate = toDateString(new Date());
  const linkedSale = state.sales.find((sale) => sale.caseId === foundCase.id);
  return {
    customerName: foundCase.customerName || "顧客名未設定",
    subject: foundCase.caseName || "請求内容",
    invoiceDate,
    invoiceNumber: linkedSale?.invoiceNumber || "未採番",
    companyName: OFFICE_INFO.name,
    companyZip: OFFICE_INFO.zip,
    companyAddress: OFFICE_INFO.address,
    companyPhone: OFFICE_INFO.tel,
    companyEmail: OFFICE_INFO.email,
    registrationNumber: OFFICE_INFO.registrationNumber,
    details: [{
      no: 1,
      itemName: foundCase.caseName,
      quantity: 1,
      unitPrice: subtotal,
      amount: subtotal,
    }],
    subtotal,
    tax,
    total,
    transferInfo: OFFICE_INFO.transferInfo,
    note: asTrimmedText(noteOverride ?? "") || OFFICE_INFO.invoiceNote || "",
  };
}

function exportInvoiceDataForCase(caseId) {
  const foundCase = state.cases.find((entry) => entry.id === caseId);
  if (!foundCase) return;
  const note = requestInvoiceMemo(asTrimmedText(saleInvoiceMemoInput?.value || ""));
  if (note === null) return;
  const invoiceData = buildInvoiceDocumentFromCase(foundCase, note);
  downloadInvoiceWorkbook(invoiceData);
}

function openEstimatePrintPreview(estimateId) {
  const estimate = state.estimates.find((entry) => entry.id === estimateId);
  if (!estimate) {
    showAppMessage("出力対象データが見つかりません", true);
    return;
  }

  try {
    const documentData = buildEstimateDocumentFromEstimate(estimate);
    openBusinessDocumentPrintWindow(documentData, { type: "estimate" });
  } catch (error) {
    console.error("見積書の出力に失敗しました。", error);
    showAppMessage("見積書の出力に失敗しました。再度お試しください。", true);
  }
}

function openInvoicePrintPreviewFromEstimate(estimateId) {
  const estimate = state.estimates.find((entry) => entry.id === estimateId);
  if (!estimate) {
    showAppMessage("出力対象データが見つかりません", true);
    return;
  }

  try {
    const note = requestInvoiceMemo(estimate.memo || "");
    if (note === null) return;
    const documentData = buildInvoiceDocumentFromEstimate(estimate, note);
    openBusinessDocumentPrintWindow(documentData, { type: "invoice" });
  } catch (error) {
    console.error("請求書の出力に失敗しました。", error);
    showAppMessage("請求書の出力に失敗しました。再度お試しください。", true);
  }
}

function openInvoicePrintPreviewFromCase(caseId) {
  const foundCase = state.cases.find((entry) => entry.id === caseId);
  if (!foundCase) {
    showAppMessage("出力対象データが見つかりません", true);
    return;
  }

  try {
    const note = requestInvoiceMemo(asTrimmedText(saleInvoiceMemoInput?.value || ""));
    if (note === null) return;
    const documentData = buildInvoiceDocumentFromCase(foundCase, note);
    openBusinessDocumentPrintWindow(documentData, { type: "invoice" });
  } catch (error) {
    console.error("請求書の出力に失敗しました。", error);
    showAppMessage("請求書の出力に失敗しました。再度お試しください。", true);
  }
}

function openBusinessDocumentPrintWindow(documentData, options = { type: "invoice" }) {
  let html = "";
  try {
    html = buildBusinessDocumentHtml(documentData, options);
  } catch (error) {
    console.error("帳票HTMLの生成に失敗しました。", error);
    showAppMessage("帳票の生成に失敗しました。再度お試しください。", true);
    return;
  }

  if (typeof html !== "string" || !html.trim()) {
    console.error("帳票HTMLの生成結果が不正です。", { documentData, options, html });
    showAppMessage("帳票の生成に失敗しました。再度お試しください。", true);
    return;
  }

  const printWindow = window.open("", "_blank");
  if (!printWindow) {
    showAppMessage("ポップアップがブロックされています。ブラウザ設定を確認してください。", true);
    return;
  }

  printWindow.document.open();
  printWindow.document.write(html);
  printWindow.document.close();
}

function downloadInvoiceWorkbook(invoiceData) {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, createBusinessDocumentSheet(invoiceData, { type: "invoice" }), "請求書");
  const invoiceDate = String(invoiceData.invoiceDate || toDateString(new Date()));
  const safeCustomerName = sanitizeFileNamePart(invoiceData.customerName || "customer");
  XLSX.writeFile(wb, `invoice_${safeCustomerName}_${invoiceDate}.xlsx`);
}

function buildEstimateDocumentFromEstimate(estimate) {
  const rows = state.estimateItems
    .filter((row) => row.estimateId === estimate.id)
    .sort((a, b) => a.sortOrder - b.sortOrder)
    .map((row, index) => ({
      no: index + 1,
      itemName: row.itemName,
      quantity: row.quantity,
      unitPrice: row.unitPrice,
      amount: row.amount,
    }));
  if (!rows.length) {
    rows.push({
      no: 1,
      itemName: estimate.estimateTitle || "見積内容",
      quantity: 1,
      unitPrice: estimate.subtotal ?? 0,
      amount: estimate.subtotal ?? 0,
    });
  }

  const estimateDate = estimate.estimateDate || toDateString(new Date());
  return {
    customerName: estimate.customerName || "顧客名未設定",
    subject: estimate.estimateTitle || "見積内容",
    estimateDate,
    estimateNumber: estimate.estimateNumber || "未採番",
    validUntil: estimate.validUntil || "",
    companyName: OFFICE_INFO.name,
    companyZip: OFFICE_INFO.zip,
    companyAddress: OFFICE_INFO.address,
    companyPhone: OFFICE_INFO.tel,
    companyEmail: OFFICE_INFO.email,
    details: rows,
    subtotal: estimate.subtotal ?? 0,
    tax: estimate.tax ?? 0,
    total: estimate.total ?? 0,
    paymentTerms: "お支払い条件：請求書受領後7日以内に銀行振込",
    note: estimate.memo || OFFICE_INFO.estimateNote,
  };
}

function downloadEstimateWorkbook(estimateData) {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, createBusinessDocumentSheet(estimateData, { type: "estimate" }), "見積書");
  const estimateDate = String(estimateData.estimateDate || toDateString(new Date()));
  const safeCustomerName = sanitizeFileNamePart(estimateData.customerName || "customer");
  XLSX.writeFile(wb, `estimate_${safeCustomerName}_${estimateDate}.xlsx`);
}

function exportEstimateDataForEstimate(estimateId) {
  const estimate = state.estimates.find((entry) => entry.id === estimateId);
  if (!estimate) return;
  const estimateData = buildEstimateDocumentFromEstimate(estimate);
  downloadEstimateWorkbook(estimateData);
}

function createBusinessDocumentSheet(documentData, options = { type: "invoice" }) {
  const isInvoice = options.type === "invoice";
  const title = isInvoice ? "請求書" : "見積書";
  const amountLabel = isInvoice ? "ご請求金額（税込）" : "お見積金額（税込）";
  const sentence = isInvoice ? "下記の通りご請求申し上げます。" : "下記の通りお見積り申し上げます。";
  const leftLabel = isInvoice ? "請求先" : "宛先";
  const dateLabel = isInvoice ? "発行日" : "見積日";
  const numberLabel = isInvoice ? "請求書番号" : "見積番号";
  const numberValue = isInvoice ? documentData.invoiceNumber : documentData.estimateNumber;
  const dateValue = isInvoice ? documentData.invoiceDate : documentData.estimateDate;
  const hasCompanyZip = Boolean(asTrimmedText(documentData.companyZip || ""));
  const hasCompanyAddress = Boolean(asTrimmedText(documentData.companyAddress || ""));
  const hasCompanyPhone = Boolean(asTrimmedText(documentData.companyPhone || ""));
  const hasRegistrationNumber = Boolean(asTrimmedText(documentData.registrationNumber || ""));
  const hasTransferInfo = Boolean(asTrimmedText(documentData.transferInfo || ""));

  const detailRows = (documentData.details || []).slice(0, 10);
  while (detailRows.length < 10) detailRows.push({ no: "", itemName: "", quantity: "", unitPrice: "", amount: "" });
  const rows = [];
  for (let i = 0; i < 35; i += 1) rows.push(["", "", "", "", "", "", "", ""]);
  rows[0][0] = title;
  rows[2][0] = leftLabel;
  rows[3][0] = `${documentData.customerName || "顧客名未設定"} 御中`;
  rows[2][5] = dateLabel;
  rows[2][6] = dateValue || "";
  rows[3][5] = numberLabel;
  rows[3][6] = numberValue || "";
  rows[4][5] = "事務所名";
  rows[4][6] = documentData.companyName || "";
  rows[5][5] = hasCompanyZip ? "郵便番号" : "";
  rows[5][6] = hasCompanyZip ? `〒${documentData.companyZip || ""}` : "";
  rows[6][5] = hasCompanyAddress ? "住所" : "";
  rows[6][6] = hasCompanyAddress ? (documentData.companyAddress || "") : "";
  rows[7][5] = hasCompanyPhone ? "電話" : "";
  rows[7][6] = hasCompanyPhone ? `TEL: ${documentData.companyPhone || ""}` : "";
  rows[8][5] = "メール";
  rows[8][6] = documentData.companyEmail ? `Email: ${documentData.companyEmail}` : "";
  rows[9][5] = isInvoice ? (hasRegistrationNumber ? "登録番号" : "") : "有効期限";
  rows[9][6] = isInvoice ? (hasRegistrationNumber ? (documentData.registrationNumber || "") : "") : (documentData.validUntil || "");

  rows[10][0] = sentence;
  rows[11][0] = amountLabel;
  rows[11][5] = documentData.total ?? 0;
  rows[13] = ["No", "品目", "", "", "数量", "単価", "金額", ""];

  detailRows.forEach((row, index) => {
    const r = 14 + index;
    rows[r] = [row.no, row.itemName, "", "", row.quantity, row.unitPrice, row.amount, ""];
  });

  rows[26][5] = "小計";
  rows[26][6] = documentData.subtotal ?? 0;
  rows[27][5] = "消費税10%";
  rows[27][6] = documentData.tax ?? 0;
  rows[28][5] = "合計";
  rows[28][6] = documentData.total ?? 0;
  rows[30][0] = hasTransferInfo ? "振込先" : "";
  rows[31][0] = hasTransferInfo ? (documentData.transferInfo || "") : "";
  rows[32][0] = "備考";
  rows[33][0] = documentData.note || "";
  if (!isInvoice) {
    rows[30][5] = "有効期限";
    rows[30][6] = documentData.validUntil || "";
    rows[31][5] = "支払条件";
    rows[31][6] = documentData.paymentTerms || "";
  }

  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws["!ref"] = "A1:H35";
  ws["!cols"] = [{ wch: 5 }, { wch: 18 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 12 }, { wch: 14 }, { wch: 4 }];
  ws["!rows"] = Array.from({ length: 35 }, (_, i) => ({ hpt: i === 0 ? 30 : 20 }));
  ws["!margins"] = { left: 0.3, right: 0.3, top: 0.4, bottom: 0.4, header: 0.2, footer: 0.2 };
  ws["!pageSetup"] = { paperSize: 9, orientation: "portrait", fitToWidth: 1, fitToHeight: 1, horizontalDpi: 300, verticalDpi: 300 };
  ws["!printArea"] = "A1:H35";
  ws["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 7 } },
    { s: { r: 2, c: 0 }, e: { r: 2, c: 3 } },
    { s: { r: 3, c: 0 }, e: { r: 3, c: 3 } },
    { s: { r: 10, c: 0 }, e: { r: 10, c: 7 } },
    { s: { r: 11, c: 0 }, e: { r: 11, c: 4 } },
    { s: { r: 13, c: 1 }, e: { r: 13, c: 3 } },
    ...Array.from({ length: 10 }, (_, i) => ({ s: { r: 14 + i, c: 1 }, e: { r: 14 + i, c: 3 } })),
    { s: { r: 30, c: 0 }, e: { r: 30, c: 4 } },
    { s: { r: 31, c: 0 }, e: { r: 31, c: 4 } },
    { s: { r: 32, c: 0 }, e: { r: 32, c: 4 } },
    { s: { r: 33, c: 0 }, e: { r: 33, c: 4 } },
  ];

  applyBusinessCellStyle(ws, 0, 0, { bold: true, fontSize: 24, align: "center" });
  applyBusinessCellStyle(ws, 3, 0, { bold: true, fontSize: 14 });
  applyBusinessCellStyle(ws, 11, 0, { bold: true, fontSize: 14 });
  applyBusinessCellStyle(ws, 11, 5, { bold: true, fontSize: 16, align: "right", numFmt: "¥#,##0" });

  for (let c = 0; c <= 7; c += 1) {
    applyBusinessCellStyle(ws, 13, c, { bold: true, align: "center", border: true, fillGray: true });
  }
  for (let r = 14; r <= 23; r += 1) {
    for (let c = 0; c <= 7; c += 1) {
      const style = { border: true };
      if (c === 4) style.align = "center";
      if (c === 5 || c === 6) style.numFmt = "¥#,##0";
      applyBusinessCellStyle(ws, r, c, style);
    }
  }
  [26, 27, 28].forEach((r) => {
    applyBusinessCellStyle(ws, r, 5, { border: true, bold: r === 28, align: "center" });
    applyBusinessCellStyle(ws, r, 6, { border: true, bold: r === 28, numFmt: "¥#,##0", align: "right" });
  });
  for (let c = 5; c <= 6; c += 1) {
    applyBusinessCellStyle(ws, 28, c, { border: true, bold: true, numFmt: c === 6 ? "¥#,##0" : undefined, align: c === 6 ? "right" : "center", fillGray: true });
  }
  for (let r = 2; r <= 9; r += 1) {
    applyBusinessCellStyle(ws, r, 5, { bold: true });
    applyBusinessCellStyle(ws, r, 6, {});
  }
  return ws;
}

function buildBusinessDocumentHtml(documentData, options = { type: "invoice" }) {
  const isInvoice = options.type === "invoice";
  const title = isInvoice ? "請求書" : "見積書";
  const description = isInvoice ? "下記のとおりご請求申し上げます。" : "下記のとおり、お見積もり申し上げます。";
  const emphasizedLabel = isInvoice ? "ご請求金額（税込）" : "お見積金額（税込）";
  const numberLabel = isInvoice ? "請求書番号" : "見積番号";
  const numberValue = isInvoice ? documentData.invoiceNumber : documentData.estimateNumber;
  const dateLabel = isInvoice ? "日付" : "日付";
  const dateValue = isInvoice ? documentData.invoiceDate : documentData.estimateDate;
  const hasCompanyZip = Boolean(asTrimmedText(documentData.companyZip || ""));
  const hasCompanyAddress = Boolean(asTrimmedText(documentData.companyAddress || ""));
  const hasCompanyPhone = Boolean(asTrimmedText(documentData.companyPhone || ""));
  const hasRegistrationNumber = Boolean(asTrimmedText(documentData.registrationNumber || ""));
  const hasTransferInfo = Boolean(asTrimmedText(documentData.transferInfo || ""));
  const extraLabel = isInvoice ? "登録番号" : "有効期限";
  const extraValue = isInvoice ? documentData.registrationNumber : documentData.validUntil;
  const extraRowHtml = !isInvoice || hasRegistrationNumber
    ? `<div class="kv-row"><strong>${extraLabel}</strong><span>${escapeHtml(extraValue || "-")}</span></div>`
    : "";
  const officeLines = [
    `<strong>${escapeHtml(documentData.companyName || "")}</strong>`,
    hasCompanyZip ? `〒${escapeHtml(documentData.companyZip)}` : "",
    hasCompanyAddress ? escapeHtml(documentData.companyAddress) : "",
    hasCompanyPhone ? `TEL: ${escapeHtml(documentData.companyPhone)}` : "",
    documentData.companyEmail ? `Email: ${escapeHtml(documentData.companyEmail)}` : "",
  ].filter(Boolean).join("<br />");
  const transferSectionHtml = hasTransferInfo
    ? `<section class="info-section transfer-section"><h3>振込先</h3><div class="note-box">${escapeHtml(documentData.transferInfo || "")}</div></section>`
    : "";
  const tableHeader = isInvoice
    ? "<tr><th>品名</th><th>単価</th><th>数量</th><th>金額</th></tr>"
    : "<tr><th>摘要</th><th>数量</th><th>単位</th><th>単価</th><th>金額</th></tr>";
  const detailRows = (documentData.details || []).map((row) => {
    if (isInvoice) {
      return `<tr><td class="item-name">${escapeHtml(row.itemName || "")}</td><td class="align-right">${formatCurrencyCompact(row.unitPrice)}</td><td class="align-center">${escapeHtml(String(row.quantity || ""))}</td><td class="align-right">${formatCurrencyCompact(row.amount)}</td></tr>`;
    }
    return `<tr><td class="item-name">${escapeHtml(row.itemName || "")}</td><td class="align-center">${escapeHtml(String(row.quantity || ""))}</td><td class="align-center">${escapeHtml(row.unit || "式")}</td><td class="align-right">${formatCurrencyCompact(row.unitPrice)}</td><td class="align-right">${formatCurrencyCompact(row.amount)}</td></tr>`;
  }).join("");

  return `<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<title>${title}</title>
<style>
@page {
  size: A4 portrait;
  margin: 12mm;
}
body {
  margin: 0;
  background: #fff;
  color: #111111;
  font-family: "Hiragino Kaku Gothic ProN", "Yu Gothic", "Meiryo", sans-serif;
}
.print-toolbar {
  max-width: 180mm;
  margin: 10px auto 0;
  display: flex;
  justify-content: flex-end;
}
.print-btn {
  border: 0;
  background: #1f3a5f;
  color: #fff;
  border-radius: 7px;
  padding: 8px 16px;
  cursor: pointer;
}
.sheet {
  width: 100%;
  max-width: 180mm;
  min-height: 273mm;
  margin: 8px auto 12px;
  padding: 0;
  background: #fff;
}
.title {
  margin: 4px 0 8px;
  text-align: center;
  font-size: 36px;
  letter-spacing: 0.32em;
  color: #1f3a5f;
  font-weight: 700;
}
.title-divider {
  height: 1.5px;
  background: #1f3a5f;
  margin-bottom: 18px;
}
.header-grid {
  display: grid;
  grid-template-columns: 1fr 320px;
  gap: 20px;
  margin-bottom: 6px;
}
.left-block { min-width: 0; }
.left-label {
  margin: 0 0 6px;
  font-size: 12px;
  color: #334155;
}
.recipient {
  display: inline-block;
  font-size: 21px;
  margin: 0;
  padding-bottom: 5px;
  border-bottom: 1px solid #1f3a5f;
}
.right-block {
  font-size: 13px;
  background: #f7f9fc;
  border: 1px solid #9aa8b8;
  padding: 10px;
}
.kv-row {
  display: grid;
  grid-template-columns: 96px 1fr;
  gap: 8px;
  padding: 2px 0;
}
.kv-row strong { color: #0f172a; }
.office-block {
  margin: 0 0 16px auto;
  width: 320px;
  border: 1px solid #9aa8b8;
  background: #f7f9fc;
  padding: 9px 10px;
  font-size: 12px;
  line-height: 1.65;
}
.intro {
  margin: 10px 0 12px;
  font-size: 14px;
}
.amount-box {
  border: 1px solid #9aa8b8;
  margin: 0 0 14px;
}
.amount-header {
  margin: 0;
  padding: 5px 10px;
  font-size: 13px;
  color: #fff;
  background: #1f3a5f;
}
.amount-value {
  margin: 0;
  padding: 10px 12px;
  font-size: 34px;
  font-weight: 800;
  color: #0f172a;
  text-align: right;
  letter-spacing: 0.03em;
}
.details-table {
  width: 100%;
  border-collapse: collapse;
  table-layout: fixed;
  font-size: 13px;
}
.details-table th, .details-table td {
  border: 1px solid #9aa8b8;
  padding: 8px;
}
.details-table th {
  background: #1f3a5f;
  color: #fff;
  text-align: center;
  font-weight: 700;
}
.details-table td {
  vertical-align: top;
  background: #fff;
}
.details-table .item-name {
  white-space: pre-wrap;
  word-break: break-word;
  overflow-wrap: anywhere;
}
.align-right { text-align: right; }
.align-center { text-align: center; }
.totals {
  width: 280px;
  margin: 12px 0 0 auto;
  border-collapse: collapse;
  font-size: 13px;
}
.totals th, .totals td {
  border: 1px solid #9aa8b8;
  padding: 7px 8px;
}
.totals th {
  text-align: left;
  background: #e8eef6;
  color: #0f172a;
}
.totals td {
  text-align: right;
  background: #fff;
}
.totals .total-row th,
.totals .total-row td {
  font-weight: 700;
  font-size: 15px;
}
.totals .total-row th { background: #1f3a5f; color: #fff; }
.document-footer {
  margin-top: 16px;
  display: grid;
  gap: 12px;
}
.info-section h3 {
  margin: 0 0 6px;
  font-size: 13px;
  color: #0f172a;
}
.note-box {
  border: 1px solid #9aa8b8;
  background: #f7f9fc;
  min-height: 62px;
  padding: 8px;
  white-space: pre-wrap;
  line-height: 1.6;
}
@media print {
  .no-print {
    display: none !important;
  }
  body {
    -webkit-print-color-adjust: exact;
    print-color-adjust: exact;
  }
  .sheet {
    margin: 0 auto;
    min-height: auto;
  }
}
</style>
</head>
<body>
  <div class="print-toolbar no-print"><button class="print-btn" onclick="window.print()">印刷 / PDF保存</button></div>
  <main class="sheet">
    <h1 class="title">${title}</h1>
    <div class="title-divider" aria-hidden="true"></div>
    <section class="header-grid">
      <div class="left-block">
        <p class="left-label">${isInvoice ? "請求先" : "宛先"}</p>
        <p class="recipient">${escapeHtml(documentData.customerName || "顧客名未設定")} 御中</p>
      </div>
      <div class="right-block">
        <div class="kv-row"><strong>${numberLabel}</strong><span>${escapeHtml(numberValue || "-")}</span></div>
        <div class="kv-row"><strong>${dateLabel}</strong><span>${escapeHtml(formatDate(dateValue))}</span></div>
        ${extraRowHtml}
      </div>
    </section>
    <section class="office-block">
      ${officeLines}
    </section>
    <p class="intro">${description}</p>
    <section class="amount-box">
      <h2 class="amount-header">${emphasizedLabel}</h2>
      <p class="amount-value">${formatCurrencyCompact(documentData.total)}</p>
    </section>
    <table class="details-table">
      <thead>${tableHeader}</thead>
      <tbody>${detailRows || (isInvoice ? "<tr><td colspan='4'>明細なし</td></tr>" : "<tr><td colspan='5'>明細なし</td></tr>")}</tbody>
    </table>
    <table class="totals">
      <tr><th>小計</th><td class="align-right">${formatCurrencyCompact(documentData.subtotal)}</td></tr>
      <tr><th>消費税10%</th><td class="align-right">${formatCurrencyCompact(documentData.tax)}</td></tr>
      <tr class="total-row"><th>合計</th><td class="align-right">${formatCurrencyCompact(documentData.total)}</td></tr>
    </table>
    <section class="document-footer">
      ${isInvoice ? transferSectionHtml : ""}
      <section class="info-section">
        <h3>${isInvoice ? "備考" : "備考"}</h3>
        <div class="note-box">${escapeHtml(isInvoice ? (documentData.note || "") : `${documentData.note || ""}\n${documentData.paymentTerms || "支払条件: 記載なし"}\n有効期限: ${documentData.validUntil || "記載なし"}`.trim())}</div>
      </section>
    </section>
  </main>
</body>
</html>`;
}

function requestInvoiceMemo(defaultValue = "") {
  const initialValue = asTrimmedText(defaultValue || "");
  const input = window.prompt("請求書備考を入力してください（未入力の場合は既定文を使用）", initialValue);
  if (input === null) return null;
  return asTrimmedText(input);
}

function formatCurrencyCompact(value) {
  const amount = Number(value);
  if (!Number.isFinite(amount)) return "¥0";
  return `¥${new Intl.NumberFormat("ja-JP").format(Math.floor(amount))}`;
}

function applyBusinessCellStyle(ws, row, col, options = {}) {
  const address = XLSX.utils.encode_cell({ r: row, c: col });
  if (!ws[address]) ws[address] = { t: "s", v: "" };
  if (options.numFmt) ws[address].z = options.numFmt;
  ws[address].s = {
    font: { name: "Meiryo UI", sz: options.fontSize || 11, bold: Boolean(options.bold) },
    alignment: { horizontal: options.align || "left", vertical: "center", wrapText: Boolean(options.wrapText) },
    border: options.border
      ? {
        top: { style: "thin", color: { rgb: "000000" } },
        bottom: { style: "thin", color: { rgb: "000000" } },
        left: { style: "thin", color: { rgb: "000000" } },
        right: { style: "thin", color: { rgb: "000000" } },
      }
      : undefined,
    fill: options.fillGray ? { patternType: "solid", fgColor: { rgb: "EDEDED" } } : undefined,
  };
}

function sanitizeFileNamePart(value) {
  return String(value || "")
    .trim()
    .replace(/[\\/:*?"<>|]/g, "_")
    .replace(/\s+/g, "_")
    .slice(0, 40) || "customer";
}

async function getNextMonthlyNumber(tableName, columnName, prefix, issueDate = toDateString(new Date())) {
  const monthKey = toMonthKey(issueDate) || toMonthKey(new Date());
  const yyyymm = monthKey.replace("-", "");
  const searchPrefix = `${prefix}-${yyyymm}-`;
  const { data, error } = await sbClient
    .from(tableName)
    .select(columnName)
    .eq("user_id", currentUser.id)
    .like(columnName, `${searchPrefix}%`);
  if (error) throw error;
  const maxSeq = (data || []).reduce((max, row) => {
    const value = String(row?.[columnName] || "");
    const matched = value.match(new RegExp(`^${prefix}-${yyyymm}-(\\d{3})$`));
    if (!matched) return max;
    return Math.max(max, Number(matched[1]));
  }, 0);
  return `${prefix}-${yyyymm}-${String(maxSeq + 1).padStart(3, "0")}`;
}

async function registerSaleFromEstimate(estimateId) {
  const estimate = state.estimates.find((entry) => entry.id === estimateId);
  if (!estimate) return;
  const caseId = estimate.caseId || await ensureCaseFromEstimate(estimateId, true);
  const payload = {
    user_id: currentUser.id,
    case_id: caseId || null,
    invoice_number: await getNextMonthlyNumber("sales", "invoice_number", "S"),
    invoice_amount: estimate.total,
    paid_amount: 0,
    paid_date: null,
    is_unpaid: true,
  };
  const { data, error } = await sbClient.from("sales").insert(payload).select().single();
  if (error) throw error;
  if (!data) throw new Error("登録結果を取得できませんでした。");
}

function buildDailyReportSummary() {
  const today = toDateString(new Date());
  const currentMonthKey = toMonthKey(new Date());
  const caseMinutes = {};
  let todayMinutes = 0;
  let monthMinutes = 0;

  state.dailyReports.forEach((entry) => {
    const minutes = entry.workMinutes || 0;
    if (entry.reportDate === today) todayMinutes += minutes;
    if (toMonthKey(entry.reportDate) === currentMonthKey) monthMinutes += minutes;
    const key = entry.caseId || "none";
    caseMinutes[key] = (caseMinutes[key] || 0) + minutes;
  });

  const caseRows = Object.entries(caseMinutes)
    .sort((a, b) => b[1] - a[1])
    .map(([caseId, minutes]) => `${resolveDailyReportCaseName(caseId === "none" ? null : caseId)}: ${formatMinutes(minutes)}`)
    .join(" / ");
  return { todayMinutes, monthMinutes, caseRows };
}

function renderDashboardWorktimeSummary() {
  if (!worktimeSummaryGrid) return;
  const summary = buildDailyReportSummary();
  worktimeSummaryGrid.innerHTML = "";
  [
    { label: "今日の作業時間", value: formatMinutes(summary.todayMinutes) },
    { label: "今月の作業時間", value: formatMinutes(summary.monthMinutes) },
  ].forEach((entry) => {
    const div = document.createElement("div");
    div.className = "summary-card";
    div.innerHTML = `<p class="label">${entry.label}</p><p class="value">${entry.value}</p>`;
    worktimeSummaryGrid.appendChild(div);
  });
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

function formatMinutes(value) {
  const minutes = Number(value);
  if (!Number.isFinite(minutes) || minutes < 0) return "0 分";
  return `${new Intl.NumberFormat("ja-JP").format(Math.floor(minutes))} 分`;
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
    entry.nextAction,
    entry.nextActionDate,
    entry.workMemo,
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
    expense.payee,
    expense.paymentMethod,
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

function normalizeExpensePaymentMethod(raw) {
  const value = asTrimmedText(raw);
  if (!value) return null;
  return EXPENSE_PAYMENT_METHODS.includes(value) ? value : "その他";
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
          client_id: asTrimmedText(row.client_id) || null,
          customer_name: customerName,
          case_name: caseName,
          estimate_amount: parseFlexibleAmount(row.estimate_amount),
          received_date: parseFlexibleDate(row.received_date),
          due_date: parseFlexibleDate(row.due_date),
          status: normalizeStoredStatus(row.status),
          work_memo: asTrimmedText(row.work_memo) || null,
          next_action_date: parseFlexibleDate(row.next_action_date),
          next_action: asTrimmedText(row.next_action) || null,
          document_url: asTrimmedText(row.document_url) || null,
          invoice_url: asTrimmedText(row.invoice_url) || null,
          receipt_url: asTrimmedText(row.receipt_url) || null,
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
          invoice_number: asTrimmedText(row.invoice_number) || null,
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
          payee: asTrimmedText(row.payee) || null,
          payment_method: normalizeExpensePaymentMethod(row.payment_method),
          receipt_url: asTrimmedText(row.receipt_url) || null,
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
    clientId: row.client_id || null,
    customerName: row.customer_name || "",
    caseName: row.case_name || "",
    estimateAmount: normalizeAmount(row.estimate_amount),
    receivedDate: row.received_date || "",
    dueDate: row.due_date || "",
    status: normalizeStoredStatus(row.status),
    workMemo: row.work_memo || "",
    nextActionDate: row.next_action_date || "",
    nextAction: row.next_action || "",
    documentUrl: row.document_url || "",
    invoiceUrl: row.invoice_url || "",
    receiptUrl: row.receipt_url || "",
    createdAt: Date.parse(row.created_at) || Date.now(),
    updatedAt: Date.parse(row.updated_at) || Date.now(),
  };
}

function mapClientFromDb(row) {
  return {
    id: row.id,
    name: row.name || "",
    clientType: row.client_type || "",
    address: row.address || "",
    tel: row.tel || "",
    email: row.email || "",
    referralSource: row.referral_source || "",
    memo: row.memo || "",
    createdAt: Date.parse(row.created_at) || Date.now(),
    updatedAt: Date.parse(row.updated_at) || Date.now(),
  };
}

function sanitizeLegacyEstimateMemo(workMemo) {
  return String(workMemo || "")
    .replace(/\[estimate:[^\]]+\]/gi, "")
    .replace(/\s{2,}/g, " ")
    .trim();
}

async function cleanupLegacyEstimateMemoMarkers() {
  if (!currentUser || !state.cases.length) return;
  const targets = state.cases.filter((entry) => String(entry.workMemo || "").includes("[estimate:"));
  if (!targets.length) return;
  await Promise.all(targets.map(async (entry) => {
    const cleanedMemo = sanitizeLegacyEstimateMemo(entry.workMemo);
    if (cleanedMemo === entry.workMemo) return;
    const { error } = await sbClient.from("cases").update({ work_memo: cleanedMemo }).eq("id", entry.id).eq("user_id", currentUser.id);
    if (error) throw error;
    entry.workMemo = cleanedMemo;
  }));
}

function mapEstimateFromDb(row) {
  return {
    id: row.id,
    clientId: row.client_id || null,
    customerName: row.customer_name || "",
    estimateNumber: row.estimate_number || "",
    estimateTitle: row.estimate_title || "",
    estimateDate: row.estimate_date || "",
    validUntil: row.valid_until || "",
    status: normalizeEstimateStatus(row.status),
    memo: row.memo || "",
    subtotal: normalizeAmount(row.subtotal) ?? 0,
    tax: normalizeAmount(row.tax) ?? 0,
    total: normalizeAmount(row.total) ?? 0,
    caseId: row.case_id || null,
    createdAt: Date.parse(row.created_at) || Date.now(),
    updatedAt: Date.parse(row.updated_at) || Date.now(),
  };
}

function mapEstimateItemFromDb(row) {
  return {
    id: row.id,
    estimateId: row.estimate_id,
    itemName: row.item_name || "",
    quantity: Number(row.quantity ?? 1) || 1,
    unitPrice: normalizeAmount(row.unit_price) ?? 0,
    amount: normalizeAmount(row.amount) ?? 0,
    sortOrder: Number(row.sort_order ?? 0) || 0,
    createdAt: Date.parse(row.created_at) || Date.now(),
  };
}

function mapSaleFromDb(row) {
  const invoiceAmount = normalizeAmount(row.invoice_amount) ?? 0;
  const paidAmount = normalizeAmount(row.paid_amount);
  return {
    id: row.id,
    caseId: row.case_id,
    invoiceNumber: row.invoice_number || "",
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
    payee: row.payee || "",
    paymentMethod: row.payment_method || "",
    receiptUrl: row.receipt_url || "",
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

function mapDailyReportFromDb(row) {
  return {
    id: row.id,
    reportDate: row.report_date || "",
    caseId: row.case_id || null,
    workContent: row.work_content || "",
    workMinutes: normalizeAmount(row.work_minutes) ?? 0,
    nextAction: row.next_action || "",
    memo: row.memo || "",
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

function setSubmitButtonsDisabled(disabled) {
  document.querySelectorAll('button[type="submit"]').forEach((button) => {
    button.disabled = disabled;
  });
}

async function withLoading(taskName, asyncFn, options = {}) {
  const { messageTarget = "app", triggerButton = null } = options;
  console.log("START", taskName);
  showLoading(true);
  setSubmitButtonsDisabled(true);
  if (triggerButton instanceof HTMLButtonElement) triggerButton.disabled = true;
  try {
    clearAppMessage();
    if (messageTarget === "auth") showAuthMessage("", false);
    return await asyncFn();
  } catch (error) {
    console.error("ERROR", taskName, error);
    const detail = error?.message ? String(error.message) : "";
    const message = `${taskName}に失敗しました。${detail}`;
    if (messageTarget === "auth") {
      showAuthMessage(message, true);
    } else {
      showAppMessage(message, true);
    }
    throw error;
  } finally {
    setSubmitButtonsDisabled(false);
    if (triggerButton instanceof HTMLButtonElement) triggerButton.disabled = false;
    showLoading(false);
    console.log("END", taskName);
  }
}

function showLoading(show) {
  if (!loadingOverlay) return;
  if (loadingTimerId) {
    window.clearTimeout(loadingTimerId);
    loadingTimerId = null;
  }

  loadingOverlay.hidden = !show;
  loadingOverlay.style.display = show ? "grid" : "none";

  if (show) {
    loadingTimerId = window.setTimeout(() => {
      showLoading(false);
      showAppMessage("読み込みが長時間継続したため自動解除しました。", true);
    }, 10000);
  }
}

function setLoading(isLoading) {
  showLoading(isLoading);
}

function clearLoadingState() {
  showLoading(false);
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

function clearAppState() {
  currentUser = null;
  state.clients = [];
  state.cases = [];
  state.estimates = [];
  state.estimateItems = [];
  state.sales = [];
  state.expenses = [];
  state.fixedExpenses = [];
  state.dailyReports = [];
  resetClientForm();
  resetCaseForm();
  resetEstimateForm();
  resetSaleForm();
  resetExpenseForm();
  resetFixedExpenseForm();
  resetDailyReportForm();
  clearAppMessage();
  clearLoadingState();
  authView.hidden = false;
  appView.hidden = true;
  userLabel.textContent = "";
}
