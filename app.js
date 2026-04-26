const SUPABASE_URL = "https://ueelzyftlbnvjvpsmpyt.supabase.co";
const SUPABASE_KEY = "sb_publishable_0DrKsieUcCyEZN_HRg8LhQ_QqFTPMtp";
const STATUS_ORDER = ["未着手", "進行中", "完了"];
const STATUS_FILTER_KEYS = [...STATUS_ORDER, "その他"];
const DEADLINE_FILTER_KEYS = ["all", "overdue", "within7", "within30"];
const SALES_PAYMENT_STATUSES = ["未入金", "一部入金", "入金済"];
const ESTIMATE_STATUS_ORDER = ["作成中", "提出済", "未回答", "受注", "失注"];
const EXPENSE_PAYMENT_METHODS = ["現金", "クレジットカード", "銀行振込", "電子マネー", "口座振替", "その他"];
const DAILY_REPORT_INTERACTION_TYPES = ["作業", "電話", "メール", "面談", "訪問", "LINE", "その他"];
const REMINDER_METHODS = ["電話", "メール", "LINE", "郵送", "訪問", "その他"];
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
  workTemplates: [],
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
  isInitialDataReady: false,
};
const editState = { clientId: null, caseId: null, workTemplateId: null, saleId: null, expenseId: null, fixedExpenseId: null, dailyReportId: null, estimateId: null };

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
const exportAnalysisExcelBtn = document.getElementById("export-analysis-excel-btn");
const exportClientAnalysisCsvBtn = document.getElementById("export-client-analysis-csv-btn");
const exportReferralAnalysisCsvBtn = document.getElementById("export-referral-analysis-csv-btn");
const excelImportForm = document.getElementById("excel-import-form");
const exportBackupJsonBtn = document.getElementById("export-backup-json-btn");
const backupRestoreForm = document.getElementById("backup-restore-form");

const tabs = Array.from(document.querySelectorAll(".tab-btn"));
const subtabButtons = Array.from(document.querySelectorAll(".subtab-btn"));
const subtabPanels = Array.from(document.querySelectorAll(".subtab-panel"));
const panels = {
  clients: document.getElementById("tab-clients"),
  cases: document.getElementById("tab-cases"),
  "work-templates": document.getElementById("tab-work-templates"),
  estimates: document.getElementById("tab-estimates"),
  sales: document.getElementById("tab-sales"),
  expenses: document.getElementById("tab-expenses"),
  "daily-reports": document.getElementById("tab-daily-reports"),
};
const dashboardSection = document.querySelector(".dashboard");
const subtabState = {
  clients: "entry",
  cases: "dashboard",
  "work-templates": "entry",
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
const analysisPeriodLabel = document.getElementById("analysis-period-label");
const clientAnalysisBody = document.getElementById("client-analysis-body");
const clientAnalysisEmpty = document.getElementById("client-analysis-empty");
const clientAnalysisWrap = document.getElementById("client-analysis-wrap");
const importantClientsBody = document.getElementById("important-clients-body");
const importantClientsEmpty = document.getElementById("important-clients-empty");
const importantClientsWrap = document.getElementById("important-clients-wrap");
const referralAnalysisBody = document.getElementById("referral-analysis-body");
const referralAnalysisEmpty = document.getElementById("referral-analysis-empty");
const referralAnalysisWrap = document.getElementById("referral-analysis-wrap");
const analyticsRankingGrid = document.getElementById("analytics-ranking-grid");
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
const pendingEstimatesCard = document.getElementById("pending-estimates-card");
const pendingEstimatesSummary = document.getElementById("pending-estimates-summary");
const pendingEstimatesBody = document.getElementById("pending-estimates-body");
const pendingEstimatesEmpty = document.getElementById("pending-estimates-empty");
const pendingEstimatesListWrap = document.getElementById("pending-estimates-list-wrap");

const clientForm = document.getElementById("client-form");
const clientsList = document.getElementById("clients-list");
const clientsEmpty = document.getElementById("clients-empty");
const clientSubmitBtn = document.getElementById("client-submit-btn");
const caseClientSelect = document.getElementById("case-client-id");
const estimateClientSelect = document.getElementById("estimate-client-id");
const clientHistoryClientSelect = document.getElementById("client-history-client-id");
const clientHistoryBody = document.getElementById("client-history-body");
const clientHistoryWrap = document.getElementById("client-history-wrap");
const clientHistoryEmpty = document.getElementById("client-history-empty");

const caseForm = document.getElementById("case-form");
const caseTemplateSelect = document.getElementById("case-template-id");
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
const saleInvoiceNumberInput = document.getElementById("sale-invoice-number");
const salesListBody = document.getElementById("sales-list-body");
const salesListWrap = document.getElementById("sales-list-wrap");
const salesEmpty = document.getElementById("sales-empty");
const saleSubmitBtn = document.getElementById("sale-submit-btn") || saleForm.querySelector('button[type="submit"]');
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
const reportClientSelect = document.getElementById("report-client-id");
const reportInteractionTypeSelect = document.getElementById("report-interaction-type");
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
const workTemplateForm = document.getElementById("work-template-form");
const workTemplateSubmitBtn = document.getElementById("work-template-submit-btn");
const workTemplatesList = document.getElementById("work-templates-list");
const workTemplatesEmpty = document.getElementById("work-templates-empty");
const workTemplateItemTemplate = document.getElementById("work-template-item-template");
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
let eventsBound = false;
let loadingCount = 0;
let loadingTimeoutId = null;
const BACKUP_TABLE_KEYS = ["clients", "work_templates", "cases", "estimates", "estimate_items", "sales", "expenses", "fixed_expenses", "daily_reports"];
const RESTORE_INSERT_ORDER = ["clients", "work_templates", "cases", "estimates", "estimate_items", "sales", "expenses", "fixed_expenses", "daily_reports"];
const RESTORE_DELETE_ORDER = ["estimate_items", "sales", "expenses", "fixed_expenses", "daily_reports", "estimates", "cases", "work_templates", "clients"];
const CASE_MUTATION_COLUMNS = [
  "user_id",
  "client_id",
  "estimate_id",
  "customer_name",
  "case_name",
  "estimate_amount",
  "received_date",
  "due_date",
  "status",
  "work_memo",
  "next_action_date",
  "next_action",
  "template_id",
  "required_documents",
  "task_list",
  "document_url",
  "invoice_url",
  "receipt_url",
];
const SALES_MUTATION_COLUMNS = [
  "user_id",
  "estimate_id",
  "case_id",
  "invoice_amount",
  "paid_amount",
  "paid_date",
  "due_date",
  "payment_status",
  "is_unpaid",
  "invoice_number",
];

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
  setAuthControlsDisabled(true);
  try {
    bindEvents();

    const {
      data: { session },
    } = await sbClient.auth.getSession();

    currentUser = session?.user ?? null;
    await withLoading("初期化", async () => {
      await applyAuthState();
    }, { messageTarget: "auth" });

    sbClient.auth.onAuthStateChange(async (_event, sessionState) => {
      if (isLoggingOut) return;
      try {
        currentUser = sessionState?.user ?? null;
        await withLoading("認証状態変更", async () => {
          await applyAuthState();
        }, { messageTarget: "auth" });
      } catch (error) {
        console.error("認証状態の更新に失敗しました。", error);
        showAuthMessage("認証状態の確認に失敗しました。ページを再読み込みしてください。", true);
      }
    });
  } catch (error) {
    console.error("初期化処理に失敗しました。", error);
    showAuthMessage("初期化に失敗しました。ページを再読み込みしてください。", true);
    authView.hidden = false;
    appView.hidden = true;
  } finally {
    setAuthControlsDisabled(false);
  }
}

function bindEvents() {
  if (eventsBound) return;
  bindCommaInputFields();
  tabs.forEach((btn) => btn.addEventListener("click", () => activateTab(btn.dataset.tab)));
  subtabButtons.forEach((btn) => btn.addEventListener("click", () => activateSubtab(btn.dataset.parentTab, btn.dataset.subtab)));
  authForm.addEventListener("submit", handleLogin);
  signupBtn.addEventListener("click", handleSignup);
  logoutBtn.addEventListener("click", handleLogout);

  clientForm?.addEventListener("submit", handleClientSubmit);
  caseForm.addEventListener("submit", handleCaseSubmit);
  workTemplateForm?.addEventListener("submit", handleWorkTemplateSubmit);
  console.log("SALE FORM FOUND", !!saleForm);
  console.log("EXPENSE FORM FOUND", !!expenseForm);
  console.log("DAILY REPORT FORM FOUND", !!dailyReportForm);
  saleForm?.addEventListener("submit", handleSaleSubmit);
  expenseForm?.addEventListener("submit", handleExpenseSubmit);
  fixedExpenseForm.addEventListener("submit", handleFixedExpenseSubmit);
  dailyReportForm?.addEventListener("submit", handleDailyReportSubmit);
  estimateForm?.addEventListener("submit", handleEstimateSubmit);

  clearBtn.addEventListener("click", handleClearAll);
  clientsList?.addEventListener("click", handleClientListAction);
  reportCaseSelect?.addEventListener("change", syncDailyReportClientFromCase);
  reportClientSelect?.addEventListener("change", syncDailyReportClientLabel);
  clientHistoryClientSelect?.addEventListener("change", renderClientHistory);
  caseList.addEventListener("click", handleCaseListAction);
  workTemplatesList?.addEventListener("click", handleWorkTemplateListAction);
  salesListBody?.addEventListener("click", handleSalesListAction);
  expensesList.addEventListener("click", handleExpensesListAction);
  fixedExpensesList.addEventListener("click", handleFixedExpensesListAction);
  dailyReportsBody?.addEventListener("click", handleDailyReportsListAction);
  unpaidListBody?.addEventListener("click", handleUnpaidListAction);
  unpaidAlertBody?.addEventListener("click", handleUnpaidListAction);
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
  pendingEstimatesBody?.addEventListener("click", handlePendingEstimateAction);
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
    forceHideLoading();
    showAppMessage("読み込みを強制解除しました。必要なら再操作してください。", true);
  });
  exportCasesCsvBtn?.addEventListener("click", handleExportCasesCsv);
  exportSalesCsvBtn?.addEventListener("click", handleExportSalesCsv);
  exportExpensesCsvBtn?.addEventListener("click", handleExportExpensesCsv);
  exportFixedExpensesCsvBtn?.addEventListener("click", handleExportFixedExpensesCsv);
  exportAllCsvBtn?.addEventListener("click", handleExportAllCsv);
  exportClientAnalysisCsvBtn?.addEventListener("click", handleExportClientAnalysisCsv);
  exportReferralAnalysisCsvBtn?.addEventListener("click", handleExportReferralAnalysisCsv);
  csvImportForm?.addEventListener("submit", handleCsvImportSubmit);
  exportExcelBtn?.addEventListener("click", handleExportExcel);
  exportAnalysisExcelBtn?.addEventListener("click", handleExportAnalysisExcel);
  excelImportForm?.addEventListener("submit", handleExcelImportSubmit);
  exportBackupJsonBtn?.addEventListener("click", handleExportBackupJson);
  backupRestoreForm?.addEventListener("submit", handleBackupRestoreSubmit);
  document.addEventListener("wheel", handleNumberInputWheel, { passive: true });
  estimateAddItemBtn?.addEventListener("click", () => addEstimateItemRow());
  estimateItemsWrap?.addEventListener("input", handleEstimateItemsInput);
  estimateItemsWrap?.addEventListener("click", handleEstimateItemsClick);
  estimateList?.addEventListener("click", handleEstimateListAction);
  caseClientSelect?.addEventListener("change", syncCaseCustomerFromClient);
  caseTemplateSelect?.addEventListener("change", handleCaseTemplateChange);
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
  eventsBound = true;
  console.log("EVENTS BOUND");
  activateTab("cases");
}

function handleAggregationChange(event) {
  const next = event?.target?.value;
  if (next === "all") state.selectedAggregation = "all";
  else state.selectedAggregation = next === "year" ? "year" : "month";
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
  forceHideLoading();
  await handleResumeRefresh("visibilitychange");
}

async function handleWindowFocus() {
  forceHideLoading();
  await handleResumeRefresh("focus");
}

async function handlePageShow() {
  forceHideLoading();
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

  setDataMutationControlsEnabled(false);
  state.isInitialDataReady = false;
  authView.hidden = true;
  appView.hidden = false;
  userLabel.textContent = currentUser.email || "ログイン中";

  await loadAllData();
  resetCaseForm();
  resetEstimateForm();
  resetSaleForm();
  resetExpenseForm();
  resetFixedExpenseForm();
  resetDailyReportForm();
  renderAfterDataChanged();
  state.isInitialDataReady = true;
  setDataMutationControlsEnabled(true);
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
      clearAppState();
      authView.hidden = false;
      appView.hidden = true;
      const { error } = await sbClient.auth.signOut();
      if (error) throw error;
      showAuthMessage("ログアウトしました。", false);
      authForm.reset();
    }, { messageTarget: "auth", triggerButton: logoutBtn });
  } finally {
    forceHideLoading();
    isLoggingOut = false;
  }
}

async function loadAllData() {
  if (!currentUser || isLoggingOut) return;

  const [clientsRes, workTemplatesRes, casesRes, estimatesRes, estimateItemsRes, salesRes, expensesRes, fixedExpensesRes, dailyReportsRes] = await Promise.all([
    sbClient.from("clients").select("*").eq("user_id", currentUser.id),
    sbClient.from("work_templates").select("*").eq("user_id", currentUser.id),
    sbClient.from("cases").select("*").eq("user_id", currentUser.id),
    sbClient.from("estimates").select("*").eq("user_id", currentUser.id),
    sbClient.from("estimate_items").select("*").eq("user_id", currentUser.id),
    sbClient.from("sales").select("*").eq("user_id", currentUser.id),
    sbClient.from("expenses").select("*").eq("user_id", currentUser.id),
    sbClient.from("fixed_expenses").select("*").eq("user_id", currentUser.id),
    sbClient.from("daily_reports").select("*").eq("user_id", currentUser.id),
  ]);

  if (clientsRes.error) throw new Error(`clients: ${formatSupabaseError(clientsRes.error)}`);
  if (workTemplatesRes.error) throw new Error(`work_templates: ${formatSupabaseError(workTemplatesRes.error)}`);
  if (casesRes.error) throw new Error(`cases: ${formatSupabaseError(casesRes.error)}`);
  if (estimatesRes.error) throw new Error(`estimates: ${formatSupabaseError(estimatesRes.error)}`);
  if (estimateItemsRes.error) throw new Error(`estimate_items: ${formatSupabaseError(estimateItemsRes.error)}`);
  if (salesRes.error) {
    console.error("LOAD SALES ERROR", salesRes.error);
    throw new Error(`sales: ${formatSupabaseError(salesRes.error)}`);
  }
  if (expensesRes.error) throw new Error(`expenses: ${formatSupabaseError(expensesRes.error)}`);
  if (fixedExpensesRes.error) throw new Error(`fixed_expenses: ${formatSupabaseError(fixedExpensesRes.error)}`);
  if (dailyReportsRes.error) throw new Error(`daily_reports: ${formatSupabaseError(dailyReportsRes.error)}`);

  state.clients = (clientsRes.data || []).map(mapClientFromDb);
  state.workTemplates = (workTemplatesRes.data || []).map(mapWorkTemplateFromDb);
  if (!state.workTemplates.length) {
    const seeded = await seedDefaultWorkTemplates();
    if (seeded.length) state.workTemplates = seeded;
  }
  state.cases = (casesRes.data || []).map(mapCaseFromDb);
  state.estimates = (estimatesRes.data || []).map(mapEstimateFromDb);
  state.estimateItems = (estimateItemsRes.data || []).map(mapEstimateItemFromDb);
  console.log("LOAD SALES RAW", salesRes.data);
  state.sales = (salesRes.data || []).map(mapSaleFromDb);
  console.log("LOAD SALES MAPPED", state.sales);
  state.expenses = (expensesRes.data || []).map(mapExpenseFromDb);
  state.fixedExpenses = (fixedExpensesRes.data || []).map(mapFixedExpenseFromDb);
  state.dailyReports = (dailyReportsRes.data || []).map(mapDailyReportFromDb);
  console.log("LOAD ALL DATA DONE");
  await cleanupLegacyEstimateMemoMarkers();

  const createdCount = await ensureMonthlyFixedExpenses();
  if (createdCount > 0) {
    const refreshExpensesRes = await sbClient.from("expenses").select("*").eq("user_id", currentUser.id);
    if (refreshExpensesRes.error) throw refreshExpensesRes.error;
    state.expenses = (refreshExpensesRes.data || []).map(mapExpenseFromDb);
    showAppMessage(`${createdCount}件の固定費を当月分として自動計上しました。`, false);
  }
}

async function refreshAfterMutation(actionName, message) {
  await loadAllData();
  console.log("LOAD DONE", actionName);
  renderAfterDataChanged();
  console.log("RENDER DONE", actionName);
  if (message) showAppMessage(message, false);
}

function ensureInitialDataReady(actionName = "操作") {
  if (state.isInitialDataReady) return true;
  showAppMessage(`${actionName}は初期データの読み込み完了後に実行できます。`, true);
  return false;
}

function pickObjectKeys(source, keys) {
  return keys.reduce((acc, key) => {
    if (Object.prototype.hasOwnProperty.call(source, key)) acc[key] = source[key];
    return acc;
  }, {});
}

function buildCasePayloadFromForm() {
  const customerName = asTrimmedText(caseForm.elements.customerName.value);
  const caseName = asTrimmedText(caseForm.elements.caseName.value);
  const selectedClientId = caseForm.elements.caseClientId.value || null;
  const selectedClient = selectedClientId ? state.clients.find((entry) => entry.id === selectedClientId) : null;
  const rawPayload = {
    user_id: currentUser.id,
    client_id: selectedClientId,
    customer_name: selectedClient?.name || customerName,
    case_name: caseName,
    estimate_amount: normalizeAmount(caseForm.elements.amount.value),
    received_date: caseForm.elements.receivedDate.value || null,
    due_date: caseForm.elements.dueDate.value || null,
    status: normalizeStatus(caseForm.elements.status.value),
    work_memo: asTrimmedText(caseForm.elements.workMemo.value) || null,
    next_action_date: caseForm.elements.nextActionDate.value || null,
    next_action: asTrimmedText(caseForm.elements.nextAction.value) || null,
    template_id: caseForm.elements.caseTemplateId.value || null,
    required_documents: asTrimmedText(caseForm.elements.requiredDocuments.value) || null,
    task_list: asTrimmedText(caseForm.elements.taskList.value) || null,
    document_url: asTrimmedText(caseForm.elements.documentUrl.value) || null,
    invoice_url: asTrimmedText(caseForm.elements.invoiceUrl.value) || null,
    receipt_url: asTrimmedText(caseForm.elements.receiptUrl.value) || null,
  };
  return pickObjectKeys(rawPayload, CASE_MUTATION_COLUMNS);
}

function formatSupabaseError(error) {
  if (!error) return "";
  const chunks = [error.message, error.code, error.details, error.hint]
    .filter((part) => typeof part === "string" && part.trim())
    .map((part) => part.trim());
  return chunks.length ? chunks.join(" / ") : "不明なエラーです。";
}

async function handleClientSubmit(event) {
  event.preventDefault();
  console.log("CLIENT SUBMIT FIRED");
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
  const actionName = taskName;
  console.log("EDIT STATE", editState);
  console.log("ACTION START", actionName, editState.clientId || "new");
  console.log("PAYLOAD", payload);
  try {
    await withLoading(taskName, async () => {
      if (editState.clientId) {
        const { data, error } = await sbClient.from("clients").update(payload).eq("id", editState.clientId).eq("user_id", currentUser.id).select().single();
        if (error) throw error;
        if (!data) throw new Error("更新結果を取得できませんでした。");
        console.log("DB DONE", actionName, editState.clientId);
        await refreshAfterMutation(actionName, "顧客を更新しました。");
        resetClientForm();
        return;
      }
      const { data, error } = await sbClient.from("clients").insert(payload).select().single();
      if (error) throw error;
      if (!data) throw new Error("登録結果を取得できませんでした。");
      console.log("DB DONE", actionName, data.id || "new");
      await refreshAfterMutation(actionName, "顧客を登録しました。");
      resetClientForm();
    }, { triggerButton: event.submitter });
  } catch (error) {
    showAppMessage(`顧客保存に失敗しました。${formatSupabaseError(error)}`, true);
  } finally {
    console.log("ACTION FINALLY", actionName);
    forceHideLoading();
  }
}

async function handleCaseSubmit(event) {
  event.preventDefault();
  console.log("CASE SUBMIT FIRED");
  if (!currentUser || !ensureInitialDataReady("案件登録")) return;

  const taskName = editState.caseId ? "案件更新" : "案件登録";
  const isEdit = Boolean(editState.caseId);
  try {
    await withLoading(taskName, async () => {
      const payload = buildCasePayloadFromForm();
      console.log("EDIT STATE", editState);
      const templateId = payload.template_id || "";
      console.log("CASE TEMPLATE VALUE", templateId);
      console.log("PAYLOAD", payload);
      if (!payload.customer_name || !payload.case_name) return;

      if (isEdit) {
        const { data, error } = await sbClient.from("cases").update(payload).eq("id", editState.caseId).eq("user_id", currentUser.id).select().single();
        if (error) throw error;
        if (!data) throw new Error("更新結果を取得できませんでした。");
        console.log("CASE INSERT SUCCESS", data);
      } else {
        const { data, error } = await sbClient.from("cases").insert(payload).select().single();
        if (error) throw error;
        if (!data) throw new Error("案件登録結果を取得できませんでした。");
        console.log("CASE INSERT SUCCESS", data);
      }
      await loadAllData();
      console.log("LOAD ALL DATA DONE");
      renderAfterDataChanged();
      console.log("RENDER DONE");
      subtabState.cases = "list";
      resetCaseForm();
      activateTab("cases");
      showAppMessage(isEdit ? "案件を更新しました。" : "案件を登録しました。", false);
    }, { triggerButton: event.submitter });
  } catch (error) {
    console.error("案件登録に失敗しました", error);
    showAppMessage(`案件保存に失敗しました。${formatSupabaseError(error)}`, true);
  } finally {
    console.log("CASE SUBMIT FINALLY");
    forceHideLoading();
  }
}

async function handleWorkTemplateSubmit(event) {
  event.preventDefault();
  console.log("WORK TEMPLATE SUBMIT FIRED");
  if (!currentUser || !workTemplateForm) return;
  const name = asTrimmedText(workTemplateForm.elements.templateName.value);
  if (!name) return;
  const payload = {
    user_id: currentUser.id,
    name,
    default_due_days: parseNumberInput(workTemplateForm.elements.templateDefaultDueDays.value),
    required_documents: asTrimmedText(workTemplateForm.elements.templateRequiredDocuments.value) || null,
    default_tasks: asTrimmedText(workTemplateForm.elements.templateDefaultTasks.value) || null,
    memo: asTrimmedText(workTemplateForm.elements.templateMemo.value) || null,
  };
  const taskName = editState.workTemplateId ? "業務テンプレート更新" : "業務テンプレート登録";
  const actionName = taskName;
  console.log("EDIT STATE", editState);
  try {
    await withLoading(taskName, async () => {
      if (editState.workTemplateId) {
        const { data, error } = await sbClient.from("work_templates").update(payload).eq("id", editState.workTemplateId).eq("user_id", currentUser.id).select().single();
        if (error) throw error;
        if (!data) throw new Error("更新結果を取得できませんでした。");
        await refreshAfterMutation(actionName, "業務テンプレートを更新しました。");
        resetWorkTemplateForm();
        return;
      }
      const { data, error } = await sbClient.from("work_templates").insert(payload).select().single();
      if (error) throw error;
      if (!data) throw new Error("登録結果を取得できませんでした。");
      await refreshAfterMutation(actionName, "業務テンプレートを登録しました。");
      resetWorkTemplateForm();
    }, { triggerButton: event.submitter });
  } catch (error) {
    showAppMessage(`業務テンプレート保存に失敗しました。${formatSupabaseError(error)}`, true);
  } finally {
    forceHideLoading();
  }
}

async function handleSaleSubmit(event) {
  event.preventDefault();
  if (!currentUser) {
    showAppMessage("ログイン状態を確認できません。", true);
    return;
  }
  if (!saleForm) return;

  startLoading("売上登録");
  try {
    clearAppMessage();
    console.log("SALE SUBMIT FIRED");

    const isEdit = Boolean(editState.saleId);
    const caseId = saleCaseSelect?.value || null;
    const invoiceAmount = parseNumberInput(document.getElementById("invoice-amount")?.value);
    const paidAmount = parseNumberInput(document.getElementById("paid-amount")?.value);
    const paidDate = document.getElementById("paid-date")?.value || null;
    const dueDate = document.getElementById("sale-due-date")?.value || document.getElementById("due-date")?.value || null;

    if (!invoiceAmount || invoiceAmount <= 0) {
      throw new Error("請求額を入力してください。");
    }

    const paymentStatus = calculatePaymentStatus(invoiceAmount, paidAmount);
    const invoiceNumber = isEdit
      ? (state.sales.find((entry) => entry.id === editState.saleId)?.invoiceNumber || null)
      : await getNextMonthlyNumber("sales", "invoice_number", "S");
    const payload = {
      user_id: currentUser.id,
      case_id: caseId || null,
      invoice_amount: invoiceAmount,
      paid_amount: paidAmount || 0,
      paid_date: paidDate || null,
      due_date: dueDate || null,
      payment_status: paymentStatus,
      is_unpaid: paymentStatus !== "入金済",
      invoice_number: invoiceNumber || null,
    };

    console.log("SALE PAYLOAD", payload);

    let query;
    if (isEdit) {
      query = sbClient
        .from("sales")
        .update(payload)
        .eq("id", editState.saleId)
        .eq("user_id", currentUser.id)
        .select()
        .single();
    } else {
      query = sbClient
        .from("sales")
        .insert(payload)
        .select()
        .single();
    }

    const { data, error } = await query;
    if (error) {
      console.error("SALE SUPABASE ERROR", error);
      throw error;
    }
    if (!data) throw new Error("売上登録結果を取得できませんでした。");

    console.log("SALE INSERT SUCCESS", data);
    await loadAllData();
    console.log("SALES COUNT AFTER LOAD", state.sales?.length);
    renderAfterDataChanged();
    resetSaleForm();
    editState.saleId = null;
    subtabState.sales = "list";
    activateSubtab("sales", "list");
    showAppMessage(isEdit ? "売上を更新しました。" : "売上を登録しました。", false);
  } catch (error) {
    console.error("売上登録に失敗しました。", error);
    showAppMessage(`売上登録に失敗しました。${error?.message || ""} ${error?.details || ""} ${error?.hint || ""} ${error?.code || ""}`, true);
  } finally {
    forceHideLoading();
  }
}

function setSaleInvoiceNumberDisplay(value = "") {
  if (!saleInvoiceNumberInput) return;
  saleInvoiceNumberInput.value = value || "（登録時に自動採番）";
}

async function handleExpenseSubmit(event) {
  event.preventDefault();
  console.log("EXPENSE SUBMIT FIRED");
  if (!currentUser || !expenseForm) return;

  const expenseDate = expenseForm.elements.expenseDate.value;
  const content = expenseForm.elements.expenseContent.value.trim();
  const payee = asTrimmedText(expenseForm.elements.expensePayee.value) || null;
  const paymentMethod = normalizeExpensePaymentMethod(expenseForm.elements.expensePaymentMethod.value);
  const amount = parseNumberInput(expenseForm.elements.expenseAmount.value);
  if (!expenseDate || !content || amount <= 0) return;

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
  const actionName = taskName;
  console.log("EDIT STATE", editState);
  console.log("ACTION START", actionName, editState.expenseId || "new");
  console.log("PAYLOAD", payload);
  try {
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
        console.log("DB DONE", actionName, editState.expenseId);
        await loadAllData();
        renderAfterDataChanged();
        resetExpenseForm();
        subtabState.expenses = "list";
        activateSubtab("expenses", "list");
        showAppMessage("経費を更新しました。", false);
        return;
      }

      const { data, error } = await sbClient.from("expenses").insert(payload).select().single();
      if (error) {
        showAppMessage(`経費登録に失敗しました。\n詳細: ${error.message || error}`, true);
        throw error;
      }
      if (!data) throw new Error("登録結果を取得できませんでした。");
      console.log("DB DONE", actionName, data.id || "new");
      await loadAllData();
      renderAfterDataChanged();
      resetExpenseForm();
      subtabState.expenses = "list";
      activateSubtab("expenses", "list");
      showAppMessage("経費を登録しました。", false);
    }, { triggerButton: event.submitter });
  } catch (error) {
    showAppMessage(`経費保存に失敗しました。${formatSupabaseError(error)}`, true);
  } finally {
    console.log("ACTION FINALLY", actionName);
    forceHideLoading();
  }
}

async function handleFixedExpenseSubmit(event) {
  event.preventDefault();
  console.log("FIXED EXPENSE SUBMIT FIRED");
  if (!currentUser) return;

  const content = fixedExpenseForm.elements.fixedExpenseContent.value.trim();
  const amount = parseNumberInput(fixedExpenseForm.elements.fixedExpenseAmount.value);
  const dayOfMonth = normalizeDayOfMonth(fixedExpenseForm.elements.fixedExpenseDayOfMonth.value);
  const startDate = fixedExpenseForm.elements.fixedExpenseStartDate.value;
  if (!content || amount <= 0 || !dayOfMonth || !startDate) return;

  const payload = {
    user_id: currentUser.id,
    content,
    amount,
    day_of_month: dayOfMonth,
    start_date: startDate,
    active: Boolean(fixedExpenseForm.elements.fixedExpenseActive.checked),
  };

  const taskName = editState.fixedExpenseId ? "固定費更新" : "固定費登録";
  const actionName = taskName;
  console.log("EDIT STATE", editState);
  console.log("ACTION START", actionName, editState.fixedExpenseId || "new");
  console.log("PAYLOAD", payload);
  try {
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
      console.log("DB DONE", actionName, editState.fixedExpenseId);
      await refreshAfterMutation(actionName, "固定費を更新しました。");
      resetFixedExpenseForm();
      return;
    }

    const { data, error } = await sbClient.from("fixed_expenses").insert(payload).select().single();
    if (error) throw error;
    if (!data) throw new Error("登録結果を取得できませんでした。");
    console.log("DB DONE", actionName, data.id || "new");
    await refreshAfterMutation(actionName, "固定費を登録しました。");
    resetFixedExpenseForm();
    }, { triggerButton: event.submitter });
  } catch (error) {
    showAppMessage(`固定費保存に失敗しました。${formatSupabaseError(error)}`, true);
  } finally {
    console.log("ACTION FINALLY", actionName);
    forceHideLoading();
  }
}

async function handleDailyReportSubmit(event) {
  event.preventDefault();
  console.log("DAILY REPORT SUBMIT FIRED");
  if (!currentUser || !dailyReportForm) return;

  const reportDate = dailyReportForm.elements.reportDate?.value || "";
  const clientId = dailyReportForm.elements.reportClientId?.value || "";
  const caseId = dailyReportForm.elements.reportCaseId?.value || "";
  const interactionType = normalizeDailyReportInteractionType(dailyReportForm.elements.reportInteractionType?.value);
  const workContent = asTrimmedText(dailyReportForm.elements.reportWorkContent?.value);
  const workMinutes = parseNumberInput(dailyReportForm.elements.reportWorkMinutes?.value);
  const nextAction = asTrimmedText(dailyReportForm.elements.reportNextAction?.value);
  const nextActionDate = dailyReportForm.elements.reportNextActionDate?.value || "";
  const memo = asTrimmedText(dailyReportForm.elements.reportMemo?.value);
  console.log("DAILY REPORT VALUES", {
    reportDate,
    clientId,
    caseId,
    interactionType,
    workContent,
    workMinutes,
    nextAction,
    nextActionDate,
    memo,
  });

  if (!reportDate) {
    showAppMessage("日付を入力してください。", true);
    return;
  }
  if (!workContent) {
    showAppMessage("作業内容を入力してください。", true);
    return;
  }

  const payload = {
    user_id: currentUser.id,
    report_date: reportDate,
    client_id: clientId || null,
    case_id: caseId || null,
    interaction_type: interactionType || "作業",
    work_content: workContent,
    work_minutes: workMinutes,
    next_action: nextAction || null,
    next_action_date: nextActionDate || null,
    memo: memo || null,
  };

  const isEdit = Boolean(editState.dailyReportId);
  const taskName = isEdit ? "日報更新" : "日報登録";
  console.log("EDIT STATE", editState);
  console.log("PAYLOAD", payload);
  try {
    await withLoading(taskName, async () => {
      if (isEdit) {
        const { data, error } = await sbClient.from("daily_reports").update(payload).eq("id", editState.dailyReportId).eq("user_id", currentUser.id).select().single();
        if (error) {
          console.error("日報登録Supabaseエラー", error);
          throw error;
        }
        if (!data) throw new Error("日報更新結果を取得できませんでした。");
        console.log("DAILY REPORT INSERT SUCCESS", data);
      } else {
        const { data, error } = await sbClient.from("daily_reports").insert(payload).select().single();
        if (error) {
          console.error("日報登録Supabaseエラー", error);
          throw error;
        }
        if (!data) throw new Error("日報登録結果を取得できませんでした。");
        console.log("DAILY REPORT INSERT SUCCESS", data);
      }
      await loadAllData();
      console.log("DAILY REPORT STATE COUNT", state.dailyReports.length);
      renderAfterDataChanged();
      resetDailyReportForm();
      activateSubtab("daily-reports", "list");
      showAppMessage(isEdit ? "日報を更新しました。" : "日報を登録しました。", false);
    }, { triggerButton: event.submitter });
  } catch (error) {
    showAppMessage(`日報登録に失敗しました。${error?.message || ""} ${error?.details || ""} ${error?.hint || ""} ${error?.code || ""}`, true);
  } finally {
    forceHideLoading();
  }
}

async function handleClientListAction(event) {
  const btn = event.target.closest("button");
  if (!(btn instanceof HTMLButtonElement) || !currentUser) return;
  const item = btn.closest("[data-id]");
  const id = item?.dataset.id;
  if (!id) {
    showAppMessage("対象データIDを取得できませんでした。", true);
    return;
  }
  if (btn.classList.contains("edit-client-btn")) {
    const target = state.clients.find((entry) => entry.id === id);
    if (!target || !clientForm) return;
    subtabState.clients = "entry";
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
    await deleteClient(id);
  }
}

async function handleCaseListAction(event) {
  const btn = event.target.closest("button");
  if (!(btn instanceof HTMLButtonElement)) return;
  const item = btn.closest("[data-id]");
  if (!item || !currentUser) return;
  const id = item.dataset.id;
  if (!id) {
    showAppMessage("対象データIDを取得できませんでした。", true);
    return;
  }

  if (btn.classList.contains("edit-btn")) {
    await startCaseEdit(id);
    return;
  }
  if (btn.classList.contains("case-print-btn")) {
    openInvoicePrintPreviewFromCase(id);
    return;
  }
  if (btn.classList.contains("case-xlsx-btn")) {
    exportInvoiceDataForCase(id);
    return;
  }

  if (btn.classList.contains("delete-btn")) {
    await deleteCase(id);
  }
}

async function startCaseEdit(caseId) {
    const target = state.cases.find((entry) => entry.id === caseId);
    if (!target) return;
    subtabState.cases = "entry";
    activateTab("cases");
    editState.caseId = target.id;
    caseForm.elements.caseClientId.value = target.clientId || "";
    caseForm.elements.caseTemplateId.value = target.templateId || "";
    caseForm.elements.customerName.value = target.customerName;
    caseForm.elements.caseName.value = target.caseName;
    caseForm.elements.amount.value = formatNumberInput(target.estimateAmount ?? "");
    caseForm.elements.receivedDate.value = target.receivedDate || "";
    caseForm.elements.dueDate.value = target.dueDate || "";
    caseForm.elements.requiredDocuments.value = target.requiredDocuments || "";
    caseForm.elements.taskList.value = target.taskList || "";
    caseForm.elements.workMemo.value = sanitizeLegacyEstimateMemo(target.workMemo) || "";
    caseForm.elements.nextActionDate.value = target.nextActionDate || "";
    caseForm.elements.nextAction.value = target.nextAction || "";
    caseForm.elements.documentUrl.value = target.documentUrl || "";
    caseForm.elements.invoiceUrl.value = target.invoiceUrl || "";
    caseForm.elements.receiptUrl.value = target.receiptUrl || "";
    caseForm.elements.status.value = normalizeStatus(target.status);
    caseSubmitBtn.textContent = "案件を更新";
    caseForm.scrollIntoView({ behavior: "smooth", block: "start" });
}

function handleCaseTemplateChange(event) {
  const templateId = event?.target?.value || "";
  if (!templateId || !caseForm) return;
  const found = state.workTemplates.find((entry) => entry.id === templateId);
  if (!found) return;
  const baseDate = caseForm.elements.receivedDate.value ? new Date(caseForm.elements.receivedDate.value) : new Date();
  const baseTime = Number.isNaN(baseDate.getTime()) ? new Date() : baseDate;
  const dueDate = Number.isFinite(found.defaultDueDays) && found.defaultDueDays >= 0
    ? toDateString(new Date(baseTime.getTime() + found.defaultDueDays * 24 * 60 * 60 * 1000))
    : "";
  if (!asTrimmedText(caseForm.elements.caseName.value)) {
    caseForm.elements.caseName.value = found.name;
  }
  if (dueDate) caseForm.elements.dueDate.value = dueDate;
  caseForm.elements.requiredDocuments.value = found.requiredDocuments || "";
  caseForm.elements.taskList.value = convertTasksToChecklist(found.defaultTasks);
  if (!asTrimmedText(caseForm.elements.workMemo.value)) {
    caseForm.elements.workMemo.value = found.memo || "";
  }
}

async function handleWorkTemplateListAction(event) {
  const btn = event.target.closest("button");
  if (!(btn instanceof HTMLButtonElement)) return;
  const item = btn.closest("[data-id]");
  if (!item || !currentUser) return;
  const id = item.dataset.id;
  if (!id) return;
  if (btn.classList.contains("edit-btn")) {
    startWorkTemplateEdit(id);
    return;
  }
  if (btn.classList.contains("delete-btn")) {
    await deleteWorkTemplate(id);
  }
}

function startWorkTemplateEdit(templateId) {
  const target = state.workTemplates.find((entry) => entry.id === templateId);
  if (!target || !workTemplateForm) return;
  editState.workTemplateId = target.id;
  workTemplateForm.elements.templateName.value = target.name;
  workTemplateForm.elements.templateDefaultDueDays.value = target.defaultDueDays ?? "";
  workTemplateForm.elements.templateRequiredDocuments.value = target.requiredDocuments || "";
  workTemplateForm.elements.templateDefaultTasks.value = target.defaultTasks || "";
  workTemplateForm.elements.templateMemo.value = target.memo || "";
  workTemplateSubmitBtn.textContent = "テンプレートを更新";
  activateTab("work-templates");
  workTemplateForm.scrollIntoView({ behavior: "smooth", block: "start" });
}


async function handleSalesListAction(event) {
  const btn = event.target.closest("button");
  if (!(btn instanceof HTMLButtonElement)) return;
  const item = btn.closest("[data-id]");
  if (!item || !currentUser) return;
  const id = item.dataset.id;
  if (!id) {
    showAppMessage("対象データIDを取得できませんでした。", true);
    return;
  }

  if (btn.classList.contains("edit-btn")) {
    await editSale(id);
    return;
  }

  if (btn.classList.contains("delete-btn")) {
    await deleteSale(id);
    return;
  }

  if (btn.classList.contains("record-reminder-btn")) {
    await recordSaleReminder(id);
  }
}

async function handleExpensesListAction(event) {
  const btn = event.target.closest("button");
  if (!(btn instanceof HTMLButtonElement)) return;
  const item = btn.closest("[data-id]");
  if (!item || !currentUser) return;
  const id = item.dataset.id;
  if (!id) {
    showAppMessage("対象データIDを取得できませんでした。", true);
    return;
  }

  if (btn.classList.contains("edit-btn")) {
    await editExpense(id);
    return;
  }

  if (btn.classList.contains("delete-btn")) {
    await deleteExpense(id);
  }
}

function handleUnpaidListAction(event) {
  const btn = event.target.closest("button");
  if (!(btn instanceof HTMLButtonElement)) return;
  const saleId = btn.dataset.saleId;
  if (!saleId) return;
  if (btn.classList.contains("edit-sale-btn")) {
    editSale(saleId).catch(() => {});
    return;
  }
  if (btn.classList.contains("record-reminder-btn")) {
    recordSaleReminder(saleId).catch(() => {});
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
    return;
  }
  if (target === "sale-reminder") {
    recordSaleReminder(taskId).catch(() => {});
    return;
  }
  if (target === "daily-report") {
    editDailyReport(taskId).catch(() => {});
  }
}

function handleBillingLeakAlertAction(event) {
  const btn = event.target.closest("button");
  if (!(btn instanceof HTMLButtonElement)) return;
  const caseId = btn.dataset.caseId;
  if (!caseId) return;
  if (btn.classList.contains("register-sale-btn")) {
    openSaleFormForCase(caseId);
  }
}

function handlePendingEstimateAction(event) {
  const btn = event.target.closest("button");
  if (!(btn instanceof HTMLButtonElement)) return;
  const estimateId = btn.dataset.estimateId;
  if (!estimateId) return;
  if (btn.classList.contains("edit-pending-estimate-btn")) {
    startEstimateEdit(estimateId).catch(() => {});
  }
}

async function recordSaleReminder(saleId) {
  if (!currentUser) {
    showAppMessage("ログイン状態を確認できません。", true);
    return;
  }
  const target = state.sales.find((entry) => entry.id === saleId);
  if (!target) {
    showAppMessage("対象の売上データが見つかりません。", true);
    return;
  }
  if (!isReminderRecordableSale(target)) {
    showAppMessage("督促記録は未入金または一部入金の売上のみ登録できます。", true);
    return;
  }

  const defaultDate = toDateString(new Date());
  const reminderDate = window.prompt("督促日を入力してください（YYYY-MM-DD）", target.lastReminderDate || defaultDate);
  if (reminderDate === null) return;
  const normalizedReminderDate = parseFlexibleDate(reminderDate);
  if (!normalizedReminderDate) {
    showAppMessage("督促日の形式が不正です。YYYY-MM-DD形式で入力してください。", true);
    return;
  }

  const methodPrompt = `督促方法を入力してください（${REMINDER_METHODS.join(" / ")}）`;
  const methodInput = window.prompt(methodPrompt, target.reminderMethod || REMINDER_METHODS[0]);
  if (methodInput === null) return;
  const normalizedMethod = normalizeReminderMethod(methodInput);

  const memoInput = window.prompt("督促メモを入力してください（任意）", target.reminderMemo || "");
  if (memoInput === null) return;
  const normalizedMemo = asTrimmedText(memoInput) || null;

  startLoading("督促記録");
  try {
    clearAppMessage();
    const payload = {
      reminder_count: (Number(target.reminderCount) || 0) + 1,
      last_reminder_date: normalizedReminderDate,
      reminder_method: normalizedMethod,
      reminder_memo: normalizedMemo,
      updated_at: new Date().toISOString(),
    };
    const { data, error } = await sbClient
      .from("sales")
      .update(payload)
      .eq("id", saleId)
      .eq("user_id", currentUser.id)
      .select()
      .single();
    if (error) throw error;
    if (!data) throw new Error("督促記録の更新結果を取得できませんでした。");
    await loadAllData();
    renderAfterDataChanged();
    showAppMessage("督促記録を保存しました。", false);
  } catch (error) {
    showAppMessage(`督促記録の保存に失敗しました。${error?.message || ""} ${error?.details || ""} ${error?.hint || ""} ${error?.code || ""}`, true);
  } finally {
    forceHideLoading();
  }
}

async function startSaleEdit(saleId) {
    const target = state.sales.find((entry) => entry.id === saleId);
    if (!target) return;
    subtabState.sales = "entry";
    activateTab("sales");
    editState.saleId = target.id;
    saleCaseSelect.value = target.caseId;
    saleForm.elements.invoiceAmount.value = formatNumberInput(target.invoiceAmount);
    saleForm.elements.paidAmount.value = formatNumberInput(target.paidAmount ?? "");
    saleForm.elements.paidDate.value = target.paidDate || "";
    saleForm.elements.dueDate.value = target.dueDate || "";
    setSaleInvoiceNumberDisplay(target.invoiceNumber || "");
    saleSubmitBtn.textContent = "売上を更新";
    saleForm.scrollIntoView({ behavior: "smooth", block: "start" });
}


async function startExpenseEdit(expenseId) {
    const target = state.expenses.find((entry) => entry.id === expenseId);
    if (!target) return;
    subtabState.expenses = "entry";
    activateTab("expenses");
    editState.expenseId = target.id;
    expenseForm.elements.expenseDate.value = target.date;
    expenseForm.elements.expenseContent.value = target.content;
    expenseForm.elements.expensePayee.value = target.payee || "";
    expenseForm.elements.expensePaymentMethod.value = target.paymentMethod || "";
    expenseForm.elements.expenseAmount.value = formatNumberInput(target.amount);
    expenseForm.elements.expenseReceiptUrl.value = target.receiptUrl || "";
    expenseCaseSelect.value = target.caseId || "";
    expenseSubmitBtn.textContent = "経費を更新";
    expenseForm.scrollIntoView({ behavior: "smooth", block: "start" });
}

function openSaleFormForCase(caseId) {
  const targetCase = state.cases.find((entry) => entry.id === caseId);
  if (!targetCase || !saleCaseSelect) return;
  activateTab("sales");
  resetSaleForm();
  saleCaseSelect.value = targetCase.id;
  saleForm.elements.invoiceAmount.value = formatNumberInput(targetCase.estimateAmount ?? "");
  saleForm.elements.paidAmount.value = "";
  saleForm.elements.paidDate.value = "";
  saleForm.elements.dueDate.value = "";
  setSaleInvoiceNumberDisplay("");
  saleForm.scrollIntoView({ behavior: "smooth", block: "start" });
}

async function startFixedExpenseEdit(fixedExpenseId) {
    const target = state.fixedExpenses.find((entry) => entry.id === fixedExpenseId);
    if (!target) return;
    activateTab("expenses");
    editState.fixedExpenseId = target.id;
    fixedExpenseForm.elements.fixedExpenseContent.value = target.content;
    fixedExpenseForm.elements.fixedExpenseAmount.value = formatNumberInput(target.amount);
    fixedExpenseForm.elements.fixedExpenseDayOfMonth.value = target.dayOfMonth;
    fixedExpenseForm.elements.fixedExpenseStartDate.value = target.startDate || "";
    fixedExpenseForm.elements.fixedExpenseActive.checked = Boolean(target.active);
    fixedExpenseSubmitBtn.textContent = "固定費を更新";
    fixedExpenseForm.scrollIntoView({ behavior: "smooth", block: "start" });
}


async function handleFixedExpensesListAction(event) {
  const btn = event.target.closest("button");
  if (!(btn instanceof HTMLButtonElement)) return;
  const item = btn.closest("[data-id]");
  if (!item || !currentUser) return;
  const id = item.dataset.id;
  if (!id) {
    showAppMessage("対象データIDを取得できませんでした。", true);
    return;
  }

  if (btn.classList.contains("edit-btn")) {
    await startFixedExpenseEdit(id);
    return;
  }

  if (btn.classList.contains("toggle-btn")) {
    const target = state.fixedExpenses.find((entry) => entry.id === id);
    if (!target) return;

    try {
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
        await refreshAfterMutation("固定費更新", "固定費の有効/無効を更新しました。");
      });
    } finally {
      forceHideLoading();
    }
    return;
  }

  if (btn.classList.contains("delete-btn")) {
    await deleteFixedExpense(id);
  }
}

async function handleDailyReportsListAction(event) {
  const btn = event.target.closest("button");
  if (!(btn instanceof HTMLButtonElement) || !currentUser) return;
  const item = btn.closest("[data-id]");
  const id = item?.dataset.id;
  if (!id) {
    showAppMessage("対象データIDを取得できませんでした。", true);
    return;
  }

  if (btn.classList.contains("edit-daily-report-btn")) {
    await editDailyReport(id);
    return;
  }

  if (btn.classList.contains("delete-daily-report-btn")) {
    await deleteDailyReport(id);
  }
}

async function startDailyReportEdit(dailyReportId) {
    const target = state.dailyReports.find((entry) => entry.id === dailyReportId);
    if (!target) return;
    subtabState["daily-reports"] = "entry";
    activateTab("daily-reports");
    editState.dailyReportId = target.id;
    dailyReportForm.elements.reportDate.value = target.reportDate || "";
    reportCaseSelect.value = target.caseId || "";
    if (reportClientSelect) reportClientSelect.value = target.clientId || "";
    if (reportInteractionTypeSelect) reportInteractionTypeSelect.value = normalizeDailyReportInteractionType(target.interactionType);
    dailyReportForm.elements.reportWorkContent.value = target.workContent || "";
    dailyReportForm.elements.reportWorkMinutes.value = target.workMinutes ?? 0;
    dailyReportForm.elements.reportNextAction.value = target.nextAction || "";
    dailyReportForm.elements.reportNextActionDate.value = target.nextActionDate || "";
    dailyReportForm.elements.reportMemo.value = target.memo || "";
    syncDailyReportClientLabel();
    dailyReportSubmitBtn.textContent = "日報を更新";
    dailyReportForm.scrollIntoView({ behavior: "smooth", block: "start" });
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
  await deleteAllData();
}

async function deleteClient(id) {
  if (!currentUser) return;
  if (!id) {
    showAppMessage("削除対象の顧客IDがありません。", true);
    return;
  }
  if (!window.confirm("この顧客を削除しますか？案件や見積は削除されません。")) return;
  const taskName = "顧客削除";
  startLoading(taskName);
  try {
    clearAppMessage();
    console.log("ACTION START", taskName, id);

    const casesUpdateRes = await sbClient
      .from("cases")
      .update({ client_id: null })
      .eq("client_id", id)
      .eq("user_id", currentUser.id);
    if (casesUpdateRes.error) throw casesUpdateRes.error;

    const estimatesUpdateRes = await sbClient
      .from("estimates")
      .update({ client_id: null })
      .eq("client_id", id)
      .eq("user_id", currentUser.id);
    if (estimatesUpdateRes.error) throw estimatesUpdateRes.error;

    const reportsUpdateRes = await sbClient
      .from("daily_reports")
      .update({ client_id: null })
      .eq("client_id", id)
      .eq("user_id", currentUser.id);
    if (reportsUpdateRes.error) throw reportsUpdateRes.error;

    const clientDeleteRes = await sbClient
      .from("clients")
      .delete()
      .eq("id", id)
      .eq("user_id", currentUser.id);
    if (clientDeleteRes.error) throw clientDeleteRes.error;
    console.log("DB DONE", taskName, id);

    if (editState.clientId === id) {
      editState.clientId = null;
      resetClientForm();
    }

    await loadAllData();
    console.log("LOAD DONE", taskName);
    renderAfterDataChanged();
    console.log("RENDER DONE", taskName);
    showAppMessage("顧客を削除しました。", false);
  } catch (error) {
    console.error("顧客削除に失敗しました。", error);
    showAppMessage(`顧客削除に失敗しました。${error.message || ""}`, true);
  } finally {
    console.log("ACTION FINALLY", taskName);
    forceHideLoading();
  }
}

async function deleteCase(id) {
  if (!currentUser) return;
  if (!id) {
    showAppMessage("削除対象の案件IDがありません。", true);
    return;
  }
  if (!window.confirm("この案件を削除しますか？関連する売上・経費も削除されます。")) return;
  const taskName = "案件削除";
  startLoading(taskName);
  try {
    clearAppMessage();
    console.log("ACTION START", taskName, id);

    const salesDeleteRes = await sbClient.from("sales").delete().eq("case_id", id).eq("user_id", currentUser.id);
    if (salesDeleteRes.error) throw salesDeleteRes.error;
    const expensesDeleteRes = await sbClient.from("expenses").delete().eq("case_id", id).eq("user_id", currentUser.id);
    if (expensesDeleteRes.error) throw expensesDeleteRes.error;
    const reportsUpdateRes = await sbClient.from("daily_reports").update({ case_id: null }).eq("case_id", id).eq("user_id", currentUser.id);
    if (reportsUpdateRes.error) throw reportsUpdateRes.error;
    const estimatesUpdateRes = await sbClient.from("estimates").update({ case_id: null }).eq("case_id", id).eq("user_id", currentUser.id);
    if (estimatesUpdateRes.error) throw estimatesUpdateRes.error;
    const caseDeleteRes = await sbClient.from("cases").delete().eq("id", id).eq("user_id", currentUser.id);
    if (caseDeleteRes.error) throw caseDeleteRes.error;
    console.log("DB DONE", taskName, id);

    if (editState.caseId === id) {
      editState.caseId = null;
      resetCaseForm();
    }

    await loadAllData();
    console.log("LOAD DONE", taskName);
    renderAfterDataChanged();
    console.log("RENDER DONE", taskName);
    showAppMessage("案件を削除しました。", false);
  } catch (error) {
    console.error("案件削除に失敗しました。", error);
    showAppMessage(`案件削除に失敗しました。${error.message || ""}`, true);
  } finally {
    console.log("ACTION FINALLY", taskName);
    forceHideLoading();
  }
}

async function deleteWorkTemplate(id) {
  if (!currentUser) return;
  if (!window.confirm("この業務テンプレートを削除しますか？")) return;
  const taskName = "業務テンプレート削除";
  startLoading(taskName);
  try {
    const caseUpdateRes = await sbClient.from("cases")
      .update({ template_id: null })
      .eq("template_id", id)
      .eq("user_id", currentUser.id);
    if (caseUpdateRes.error) throw caseUpdateRes.error;
    const { error } = await sbClient.from("work_templates").delete().eq("id", id).eq("user_id", currentUser.id);
    if (error) throw error;
    if (editState.workTemplateId === id) resetWorkTemplateForm();
    await loadAllData();
    renderAfterDataChanged();
    showAppMessage("業務テンプレートを削除しました。", false);
  } catch (error) {
    showAppMessage(`業務テンプレート削除に失敗しました。${error?.message || ""}`, true);
  } finally {
    forceHideLoading();
  }
}

async function deleteSale(id) {
  if (!currentUser) return;
  if (!id) {
    showAppMessage("削除対象の売上IDがありません。", true);
    return;
  }
  if (!window.confirm("この売上を削除しますか？")) return;
  const taskName = "売上削除";
  startLoading(taskName);
  try {
    clearAppMessage();
    console.log("ACTION START", taskName, id);
    const { error } = await sbClient.from("sales").delete().eq("id", id).eq("user_id", currentUser.id);
    if (error) throw error;
    console.log("DB DONE", taskName, id);

    if (editState.saleId === id) {
      editState.saleId = null;
      resetSaleForm();
    }
    await loadAllData();
    console.log("LOAD DONE", taskName);
    renderAfterDataChanged();
    console.log("RENDER DONE", taskName);
    showAppMessage("削除しました。", false);
  } catch (error) {
    console.error("削除に失敗しました。", error);
    showAppMessage(
      `削除に失敗しました。${error?.message || ""} ${error?.details || ""} ${error?.hint || ""} ${error?.code || ""}`.trim(),
      true
    );
  } finally {
    console.log("ACTION FINALLY", taskName);
    forceHideLoading();
  }
}

async function deleteExpense(id) {
  if (!currentUser) return;
  if (!id) {
    showAppMessage("削除対象の経費IDがありません。", true);
    return;
  }
  if (!window.confirm("この経費を削除しますか？")) return;
  const taskName = "経費削除";
  startLoading(taskName);
  try {
    clearAppMessage();
    console.log("ACTION START", taskName, id);
    const { error } = await sbClient.from("expenses").delete().eq("id", id).eq("user_id", currentUser.id);
    if (error) throw error;
    console.log("DB DONE", taskName, id);

    if (editState.expenseId === id) {
      editState.expenseId = null;
      resetExpenseForm();
    }
    await loadAllData();
    console.log("LOAD DONE", taskName);
    renderAfterDataChanged();
    console.log("RENDER DONE", taskName);
    showAppMessage("削除しました。", false);
  } catch (error) {
    console.error("削除に失敗しました。", error);
    showAppMessage(`削除に失敗しました。${error.message || ""}`, true);
  } finally {
    console.log("ACTION FINALLY", taskName);
    forceHideLoading();
  }
}

async function deleteFixedExpense(id) {
  if (!currentUser) return;
  if (!id) {
    showAppMessage("削除対象の固定費IDがありません。", true);
    return;
  }
  if (!window.confirm("この固定費を削除しますか？")) return;
  const taskName = "固定費削除";
  startLoading(taskName);
  try {
    clearAppMessage();
    console.log("ACTION START", taskName, id);
    const { error } = await sbClient.from("fixed_expenses").delete().eq("id", id).eq("user_id", currentUser.id);
    if (error) throw error;
    console.log("DB DONE", taskName, id);

    if (editState.fixedExpenseId === id) {
      editState.fixedExpenseId = null;
      resetFixedExpenseForm();
    }
    await loadAllData();
    console.log("LOAD DONE", taskName);
    renderAfterDataChanged();
    console.log("RENDER DONE", taskName);
    showAppMessage("削除しました。", false);
  } catch (error) {
    console.error("削除に失敗しました。", error);
    showAppMessage(`削除に失敗しました。${error.message || ""}`, true);
  } finally {
    console.log("ACTION FINALLY", taskName);
    forceHideLoading();
  }
}

async function deleteDailyReport(id) {
  if (!currentUser) return;
  if (!id) {
    showAppMessage("削除対象の日報IDがありません。", true);
    return;
  }
  if (!window.confirm("この日報を削除しますか？")) return;
  const taskName = "日報削除";
  startLoading(taskName);
  try {
    clearAppMessage();
    console.log("ACTION START", taskName, id);
    const { error } = await sbClient.from("daily_reports").delete().eq("id", id).eq("user_id", currentUser.id);
    if (error) throw error;
    console.log("DB DONE", taskName, id);

    if (editState.dailyReportId === id) {
      editState.dailyReportId = null;
      resetDailyReportForm();
    }
    await loadAllData();
    console.log("LOAD DONE", taskName);
    renderAfterDataChanged();
    console.log("RENDER DONE", taskName);
    showAppMessage("削除しました。", false);
  } catch (error) {
    console.error("削除に失敗しました。", error);
    showAppMessage(`削除に失敗しました。${error.message || ""}`, true);
  } finally {
    console.log("ACTION FINALLY", taskName);
    forceHideLoading();
  }
}

async function deleteAllData() {
  if (!currentUser) return;
  if (!window.confirm("顧客・対応履歴・業務テンプレート・見積・案件・売上・経費・固定費・日報の全データを削除します。よろしいですか？")) return;
  const taskName = "全件削除";
  startLoading(taskName);
  try {
    clearAppMessage();
    console.log("DELETE START", taskName, "all");
    const [estimateItemsRes, estimatesRes, salesRes, expensesRes, fixedExpensesRes, dailyReportsRes, casesRes, workTemplatesRes, clientsRes] = await Promise.all([
      sbClient.from("estimate_items").delete().eq("user_id", currentUser.id),
      sbClient.from("estimates").delete().eq("user_id", currentUser.id),
      sbClient.from("sales").delete().eq("user_id", currentUser.id),
      sbClient.from("expenses").delete().eq("user_id", currentUser.id),
      sbClient.from("fixed_expenses").delete().eq("user_id", currentUser.id),
      sbClient.from("daily_reports").delete().eq("user_id", currentUser.id),
      sbClient.from("cases").delete().eq("user_id", currentUser.id),
      sbClient.from("work_templates").delete().eq("user_id", currentUser.id),
      sbClient.from("clients").delete().eq("user_id", currentUser.id),
    ]);

    if (salesRes.error) throw salesRes.error;
    if (expensesRes.error) throw expensesRes.error;
    if (estimateItemsRes.error) throw estimateItemsRes.error;
    if (estimatesRes.error) throw estimatesRes.error;
    if (fixedExpensesRes.error) throw fixedExpensesRes.error;
    if (dailyReportsRes.error) throw dailyReportsRes.error;
    if (casesRes.error) throw casesRes.error;
    if (workTemplatesRes.error) throw workTemplatesRes.error;
    if (clientsRes.error) throw clientsRes.error;
    console.log("DELETE DB DONE", taskName, "all");
    resetClientForm();
    resetCaseForm();
    resetWorkTemplateForm();
    resetEstimateForm();
    resetSaleForm();
    resetExpenseForm();
    resetFixedExpenseForm();
    resetDailyReportForm();
    await loadAllData();
    console.log("DELETE LOAD DONE", taskName, "all");
    renderAfterDataChanged();
    console.log("DELETE RENDER DONE", taskName, "all");
    showAppMessage("削除しました。", false);
  } catch (error) {
    console.error("削除に失敗しました。", error);
    showAppMessage(`削除に失敗しました。${error.message || ""}`, true);
  } finally {
    console.log("DELETE FINALLY", taskName, "all");
    forceHideLoading();
  }
}

function handleExportCasesCsv() {
  const rows = state.cases.map((entry) => ({
    client_id: entry.clientId || "",
    template_id: entry.templateId || "",
    customer_name: entry.customerName,
    case_name: entry.caseName,
    estimate_amount: entry.estimateAmount ?? "",
    received_date: entry.receivedDate || "",
    due_date: entry.dueDate || "",
    required_documents: entry.requiredDocuments || "",
    task_list: entry.taskList || "",
    status: normalizeStoredStatus(entry.status),
    work_memo: sanitizeLegacyEstimateMemo(entry.workMemo) || "",
    next_action_date: entry.nextActionDate || "",
    next_action: entry.nextAction || "",
    document_url: entry.documentUrl || "",
    invoice_url: entry.invoiceUrl || "",
    receipt_url: entry.receiptUrl || "",
  }));
  downloadCsvFile("cases.csv", ["client_id", "template_id", "customer_name", "case_name", "estimate_amount", "received_date", "due_date", "required_documents", "task_list", "status", "work_memo", "next_action_date", "next_action", "document_url", "invoice_url", "receipt_url"], rows);
}

function handleExportSalesCsv() {
  const rows = state.sales.map((entry) => {
    const foundCase = state.cases.find((c) => c.id === entry.caseId);
    return {
      customer_name: foundCase?.customerName || "",
      case_name: foundCase?.caseName || "",
      invoice_number: entry.invoiceNumber || "",
      invoice_amount: entry.invoiceAmount ?? "",
      paid_amount: entry.paidAmount ?? "",
      remaining_amount: getRemainingAmount(entry),
      paid_date: entry.paidDate || "",
      due_date: entry.dueDate || "",
      payment_status: entry.paymentStatus || "",
      is_unpaid: entry.paymentStatus !== "入金済" ? "true" : "false",
      reminder_count: entry.reminderCount ?? 0,
      last_reminder_date: entry.lastReminderDate || "",
      reminder_method: entry.reminderMethod || "",
      reminder_memo: entry.reminderMemo || "",
    };
  });
  downloadCsvFile("sales.csv", ["customer_name", "case_name", "invoice_number", "invoice_amount", "paid_amount", "remaining_amount", "paid_date", "due_date", "payment_status", "is_unpaid", "reminder_count", "last_reminder_date", "reminder_method", "reminder_memo"], rows);
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
    "estimate_number", "invoice_number", "invoice_amount", "paid_amount", "remaining_amount", "paid_date", "due_date", "payment_status", "is_unpaid", "reminder_count", "last_reminder_date", "reminder_method", "reminder_memo",
    "date", "content", "amount", "payee", "payment_method",
    "day_of_month", "start_date", "active",
    "report_date", "report_client_id", "report_client_name", "report_case_name", "report_interaction_type", "report_work_content", "report_work_minutes", "report_next_action", "report_next_action_date", "report_memo",
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
      remaining_amount: getRemainingAmount(entry),
      paid_date: entry.paidDate || "",
      due_date: entry.dueDate || "",
      payment_status: entry.paymentStatus || "",
      is_unpaid: entry.paymentStatus !== "入金済" ? "true" : "false",
      reminder_count: entry.reminderCount ?? 0,
      last_reminder_date: entry.lastReminderDate || "",
      reminder_method: entry.reminderMethod || "",
      reminder_memo: entry.reminderMemo || "",
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
  state.dailyReports.forEach((entry) => {
    const foundCase = state.cases.find((c) => c.id === entry.caseId);
    rows.push({
      data_type: "daily_report",
      report_date: entry.reportDate || "",
      report_client_id: entry.clientId || "",
      report_client_name: resolveDailyReportClientName(entry.clientId, entry.caseId),
      report_case_name: foundCase?.caseName || "",
      report_interaction_type: entry.interactionType || "",
      report_work_content: entry.workContent || "",
      report_work_minutes: entry.workMinutes ?? "",
      report_next_action: entry.nextAction || "",
      report_next_action_date: entry.nextActionDate || "",
      report_memo: entry.memo || "",
    });
  });

  downloadCsvFile("all_data.csv", headers, rows);
}

function handleExportClientAnalysisCsv() {
  const { clientRows } = buildClientAndReferralAnalytics(getActiveDashboardFilter());
  const rows = clientRows.map((row) => ({
    client_name: row.clientName,
    rank: row.rank,
    case_count: row.caseCount,
    sales_total: row.salesTotal,
    expense_total: row.expenseTotal,
    profit: row.profit,
    unpaid_total: row.unpaidTotal,
    last_contact_date: row.lastContactDate || "",
  }));
  downloadCsvFile("client_analysis.csv", ["client_name", "rank", "case_count", "sales_total", "expense_total", "profit", "unpaid_total", "last_contact_date"], rows);
}

function handleExportReferralAnalysisCsv() {
  const { referralRows } = buildClientAndReferralAnalytics(getActiveDashboardFilter());
  const rows = referralRows.map((row) => ({
    referral_source: row.referralSource,
    client_count: row.clientCount,
    case_count: row.caseCount,
    total_sales: row.salesTotal,
    total_expenses: row.expenseTotal,
    profit: row.profit,
    unpaid_amount: row.unpaidTotal,
    average_price: row.averagePrice,
    profit_margin: row.profitMargin,
    comment: row.comment,
  }));
  downloadCsvFile("referral_analysis.csv", ["referral_source", "client_count", "case_count", "total_sales", "total_expenses", "profit", "unpaid_amount", "average_price", "profit_margin", "comment"], rows);
}

async function handleCsvImportSubmit(event) {
  event.preventDefault();
  if (!currentUser || !csvImportForm) return;
  const importType = csvImportForm.elements.csvImportType.value;
  const file = csvImportForm.elements.csvImportFile.files?.[0];
  if (!file) return;
  if (!window.confirm(`「${file.name}」を${importTypeToLabel(importType)}として取り込みます。よろしいですか？`)) return;

  try {
    await withLoading("CSV取込", async () => {
      const text = await readCsvFileTextWithEncodingFallback(file);
      const result = await importCsvByType(importType, text);
      console.log("DB success", "CSV取込");
      await refreshAfterMutation("CSV取込", `CSV取込完了: 登録件数 ${result.insertedCount}件 / スキップ件数 ${result.skippedCount}件 / エラー件数 ${result.errorCount}件`);
      if (result.errorCount > 0) showAppMessage(`CSV取込完了: 登録件数 ${result.insertedCount}件 / スキップ件数 ${result.skippedCount}件 / エラー件数 ${result.errorCount}件`, true);
      csvImportForm.reset();
    }, { triggerButton: event.submitter });
  } finally {
    forceHideLoading();
  }
}

function handleExportExcel() {
  try {
    clearAppMessage();
    exportExcel();
  } catch (error) {
    console.error("Excel出力に失敗しました。", error);
    showAppMessage("Excel出力に失敗しました。", true);
  }
}

function handleExportAnalysisExcel() {
  try {
    clearAppMessage();
    exportAnalysisExcel();
  } catch (error) {
    console.error("分析Excel出力に失敗しました。", error);
    showAppMessage("分析Excel出力に失敗しました。", true);
  }
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
    const saleHeaders = ["customer_name", "case_name", "invoice_number", "invoice_amount", "paid_amount", "remaining_amount", "paid_date", "due_date", "payment_status", "is_unpaid", "reminder_count", "last_reminder_date", "reminder_method", "reminder_memo"];
    const expenseHeaders = ["case_name", "date", "content", "amount", "payee", "payment_method", "receipt_url"];
    const fixedExpenseHeaders = ["content", "amount", "day_of_month", "start_date", "active"];
    const dailyReportHeaders = ["report_date", "client_id", "client_name", "case_name", "interaction_type", "work_content", "work_minutes", "next_action", "next_action_date", "memo"];
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
        customer_name: foundCase?.customerName || "",
        case_name: foundCase?.caseName || "",
        invoice_number: entry.invoiceNumber || "",
        invoice_amount: entry.invoiceAmount ?? "",
        paid_amount: entry.paidAmount ?? "",
        remaining_amount: getRemainingAmount(entry),
        paid_date: entry.paidDate || "",
        due_date: entry.dueDate || "",
        payment_status: entry.paymentStatus || "",
        is_unpaid: entry.paymentStatus !== "入金済" ? "true" : "false",
        reminder_count: entry.reminderCount ?? 0,
        last_reminder_date: entry.lastReminderDate || "",
        reminder_method: entry.reminderMethod || "",
        reminder_memo: entry.reminderMemo || "",
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
    const dailyReportRows = state.dailyReports.map((entry) => ({
      report_date: entry.reportDate || "",
      client_id: entry.clientId || "",
      client_name: resolveDailyReportClientName(entry.clientId, entry.caseId),
      case_name: resolveDailyReportCaseName(entry.caseId),
      interaction_type: entry.interactionType || "",
      work_content: entry.workContent || "",
      work_minutes: entry.workMinutes ?? "",
      next_action: entry.nextAction || "",
      next_action_date: entry.nextActionDate || "",
      memo: entry.memo || "",
    }));

    XLSX.utils.book_append_sheet(workbook, createExcelSheet(clientRows, clientHeaders), "顧客");
    XLSX.utils.book_append_sheet(workbook, createExcelSheet(caseRows, caseHeaders), "案件");
    XLSX.utils.book_append_sheet(workbook, createExcelSheet(saleRows, saleHeaders), "売上");
    XLSX.utils.book_append_sheet(workbook, createExcelSheet(expenseRows, expenseHeaders), "経費");
    XLSX.utils.book_append_sheet(workbook, createExcelSheet(fixedExpenseRows, fixedExpenseHeaders), "固定費");
    XLSX.utils.book_append_sheet(workbook, createExcelSheet(dailyReportRows, dailyReportHeaders), "日報");
    XLSX.writeFile(workbook, "gyosei-app-export.xlsx");
    showAppMessage("Excelファイルを出力しました。", false);
  } catch (error) {
    console.error("Excel出力に失敗しました。", error);
    showAppMessage("Excel出力に失敗しました。", true);
  }
}

function exportAnalysisExcel() {
  if (!window.XLSX) {
    showAppMessage("Excel出力ライブラリの読み込みに失敗しました", true);
    return;
  }

  try {
    const filter = getActiveDashboardFilter();
    const analytics = buildClientAndReferralAnalytics(filter);
    const workbook = XLSX.utils.book_new();
    const clientRows = analytics.clientRows.map((row) => ({
      client_name: row.clientName,
      rank: row.rank,
      case_count: row.caseCount,
      sales_total: row.salesTotal,
      expense_total: row.expenseTotal,
      profit: row.profit,
      unpaid_total: row.unpaidTotal,
      last_contact_date: row.lastContactDate || "",
    }));
    const referralRows = analytics.referralRows.map((row) => ({
      referral_source: row.referralSource,
      client_count: row.clientCount,
      case_count: row.caseCount,
      total_sales: row.salesTotal,
      total_expenses: row.expenseTotal,
      profit: row.profit,
      unpaid_amount: row.unpaidTotal,
      average_price: row.averagePrice,
      profit_margin: row.profitMargin,
      comment: row.comment,
    }));
    XLSX.utils.book_append_sheet(
      workbook,
      createExcelSheet(clientRows, ["client_name", "rank", "case_count", "sales_total", "expense_total", "profit", "unpaid_total", "last_contact_date"]),
      "顧客別分析",
    );
    XLSX.utils.book_append_sheet(
      workbook,
      createExcelSheet(referralRows, ["referral_source", "client_count", "case_count", "total_sales", "total_expenses", "profit", "unpaid_amount", "average_price", "profit_margin", "comment"]),
      "紹介元別分析",
    );
    XLSX.writeFile(workbook, "gyosei-analysis-export.xlsx");
    showAppMessage("分析Excelを出力しました。", false);
  } catch (error) {
    console.error("分析Excel出力に失敗しました。", error);
    showAppMessage("分析Excel出力に失敗しました。", true);
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

  try {
    await withLoading("Excel取込", async () => {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array", cellDates: false });
      const result = await importWorkbookBySheet(workbook);
      console.log("DB success", "Excel取込");
      await refreshAfterMutation("Excel取込", `Excel取込完了: 登録件数 ${result.insertedCount}件 / スキップ件数 ${result.skippedCount}件 / エラー件数 ${result.errorCount}件`);
      if (result.errorCount > 0) showAppMessage(`Excel取込完了: 登録件数 ${result.insertedCount}件 / スキップ件数 ${result.skippedCount}件 / エラー件数 ${result.errorCount}件`, true);
      excelImportForm.reset();
    }, { triggerButton: event.submitter });
  } finally {
    forceHideLoading();
  }
}

async function handleExportBackupJson() {
  if (!currentUser) return;
  try {
    clearAppMessage();
    const payload = await buildBackupPayload();
    const today = new Date().toISOString().slice(0, 10);
    downloadJsonFile(`gyosei-app-backup-${today}.json`, payload);
    showAppMessage("バックアップJSONを出力しました。", false);
  } catch (error) {
    console.error("バックアップJSON出力に失敗しました。", error);
    showAppMessage(`バックアップJSONの出力に失敗しました。${error?.message || ""}`, true);
  } finally {
    forceHideLoading();
  }
}

async function handleBackupRestoreSubmit(event) {
  event.preventDefault();
  if (!currentUser || !backupRestoreForm) return;
  const mode = backupRestoreForm.elements.backupRestoreMode.value;
  const file = backupRestoreForm.elements.backupRestoreFile.files?.[0];
  if (!file) {
    showAppMessage("復元するJSONファイルを選択してください。", true);
    return;
  }

  if (mode === "replace") {
    const confirmed = window.confirm("現在の全データを削除してバックアップから復元します。本当に実行しますか？");
    if (!confirmed) return;
  } else {
    const confirmed = window.confirm(`「${file.name}」から追加復元を実行します。よろしいですか？`);
    if (!confirmed) return;
  }

  try {
    await withLoading("バックアップ復元", async () => {
      const text = await file.text();
      let parsed;
      try {
        parsed = JSON.parse(text);
      } catch (_error) {
        throw new Error("JSON形式が不正です。");
      }
      validateBackupJsonPayload(parsed);
      const result = await restoreBackupData(parsed.data, mode);
      await loadAllData();
      renderAfterDataChanged();
      const message = `復元完了: 登録件数 ${result.insertedCount}件 / スキップ件数 ${result.skippedCount}件 / エラー件数 ${result.errorCount}件`;
      showAppMessage(message, result.errorCount > 0);
      backupRestoreForm.reset();
    }, { triggerButton: event.submitter });
  } catch (error) {
    showAppMessage(`復元に失敗しました。${error?.message || ""}`, true);
  } finally {
    forceHideLoading();
  }
}

async function buildBackupPayload() {
  if (!currentUser) throw new Error("ログイン情報が取得できません。");
  const data = {};
  for (const tableName of BACKUP_TABLE_KEYS) {
    const { data: rows, error } = await sbClient.from(tableName).select("*").eq("user_id", currentUser.id);
    if (error) throw error;
    data[tableName] = Array.isArray(rows) ? rows : [];
  }
  return {
    app: "gyosei-app",
    version: "1.0",
    exported_at: new Date().toISOString(),
    data,
  };
}

function validateBackupJsonPayload(payload) {
  if (!payload || typeof payload !== "object") throw new Error("バックアップJSONの形式が不正です。");
  if (payload.app !== "gyosei-app") throw new Error("このJSONは gyosei-app のバックアップではないため復元できません。");
  if (!payload.data || typeof payload.data !== "object") throw new Error("バックアップJSONに data が存在しません。");

  for (const tableName of BACKUP_TABLE_KEYS) {
    if (!Array.isArray(payload.data[tableName])) {
      throw new Error(`バックアップJSONの data.${tableName} が配列ではありません。`);
    }
  }
}

async function restoreBackupData(rawData, mode) {
  if (!currentUser) throw new Error("ログイン情報が取得できません。");
  const result = { insertedCount: 0, skippedCount: 0, errorCount: 0 };

  if (mode === "replace") {
    for (const tableName of RESTORE_DELETE_ORDER) {
      const { error } = await sbClient.from(tableName).delete().eq("user_id", currentUser.id);
      if (error) throw new Error(`${tableName} の既存データ削除に失敗しました。`);
    }
  }

  for (const tableName of RESTORE_INSERT_ORDER) {
    const sourceRows = Array.isArray(rawData[tableName]) ? rawData[tableName] : [];
    const existingIdSet = new Set();
    if (mode === "append") {
      const { data: existingRows, error } = await sbClient.from(tableName).select("id").eq("user_id", currentUser.id);
      if (error) throw new Error(`${tableName} の重複確認に失敗しました。`);
      (existingRows || []).forEach((entry) => {
        if (entry?.id) existingIdSet.add(entry.id);
      });
    }

    for (const row of sourceRows) {
      if (!row || typeof row !== "object") {
        result.errorCount += 1;
        continue;
      }
      if (mode === "append" && row.id && existingIdSet.has(row.id)) {
        result.skippedCount += 1;
        continue;
      }
      const payload = { ...row, user_id: currentUser.id };
      const { error } = await sbClient.from(tableName).insert(payload);
      if (error) {
        result.errorCount += 1;
        continue;
      }
      if (payload.id) existingIdSet.add(payload.id);
      result.insertedCount += 1;
    }
  }
  return result;
}

function safeRender(name, fn) {
  try {
    if (typeof fn === "function") fn();
  } catch (error) {
    console.error(`render failed: ${name}`, error);
    showAppMessage(`画面更新中にエラーが発生しました: ${name} ${error?.message || ""}`, true);
  }
}

function renderAfterDataChanged() {
  safeRender("clients", renderClients);
  safeRender("clientOptions", renderClientOptions);
  safeRender("clientHistory", renderClientHistory);
  safeRender("workTemplates", renderWorkTemplates);
  safeRender("workTemplateOptions", renderWorkTemplateOptions);
  safeRender("caseOptions", renderCaseOptions);
  safeRender("cases", renderCases);
  safeRender("estimates", renderEstimates);
  safeRender("sales", renderSales);
  safeRender("expenses", renderExpenses);
  safeRender("fixedExpenses", renderFixedExpenses);
  safeRender("dailyReports", renderDailyReports);
  safeRender("dashboard", renderDashboard);
  safeRender("todayTasks", renderTodayTasks);
  safeRender("unpaidAlerts", renderUnpaidAlerts);
  safeRender("billingLeakAlerts", renderBillingLeakAlerts);
  safeRender("deadlineAlerts", renderDeadlineAlerts);
  safeRender("nextActionAlerts", renderNextActionAlerts);
  safeRender("pendingEstimates", renderPendingEstimates);
  safeRender("clientAnalysis", renderClientAnalysis);
  safeRender("referralAnalysis", renderReferralAnalysis);
}

function renderTodayTasks() {
  if (!todayTaskCard || !todayTaskSummary || !todayTaskBody || !todayTaskEmpty || !todayTaskListWrap) return;
  renderTodayTaskCard();
}

function renderUnpaidAlerts() {
  if (!unpaidAlertCard || !unpaidAlertSummary || !unpaidAlertBody || !unpaidAlertEmpty || !unpaidAlertListWrap) return;
  renderUnpaidAlert({
    mode: state.selectedAggregation,
    monthKey: state.selectedMonth,
    year: state.selectedYear,
  });
}

function renderBillingLeakAlerts() {
  if (!billingLeakAlertCard || !billingLeakAlertSummary || !billingLeakAlertBody || !billingLeakAlertEmpty || !billingLeakAlertListWrap) return;
  renderBillingLeakAlert();
}

function renderDeadlineAlerts() {
  if (!deadlineAlertCard || !deadlineAlertSummary || !deadlineAlertBody || !deadlineAlertEmpty || !deadlineAlertListWrap) return;
  renderDeadlineAlertCard();
}

function renderNextActionAlerts() {
  if (!nextActionAlertCard || !nextActionAlertSummary || !nextActionAlertBody || !nextActionAlertEmpty || !nextActionAlertListWrap) return;
  renderNextActionAlertCard();
}

function renderPendingEstimates() {
  if (!pendingEstimatesCard || !pendingEstimatesSummary || !pendingEstimatesBody || !pendingEstimatesEmpty || !pendingEstimatesListWrap) return;
  const targets = getPendingEstimates(state.estimates)
    .slice()
    .sort((a, b) => b.elapsedDays - a.elapsedDays || toSortTimestamp(a.baseDate) - toSortTimestamp(b.baseDate));
  pendingEstimatesCard.classList.toggle("has-alert", targets.length > 0);
  pendingEstimatesSummary.textContent = `対象 ${targets.length}件`;
  pendingEstimatesBody.innerHTML = "";
  pendingEstimatesEmpty.hidden = targets.length > 0;
  pendingEstimatesListWrap.hidden = targets.length === 0;
  if (!targets.length) return;
  targets.slice(0, 5).forEach((entry) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(entry.customerName || "未設定")}</td>
      <td>${escapeHtml(entry.estimateTitle || "未設定")}</td>
      <td>${formatDate(entry.baseDate)}</td>
      <td>${entry.elapsedDays}日</td>
      <td><button type="button" class="secondary-btn edit-pending-estimate-btn" data-estimate-id="${entry.id}">編集</button></td>
    `;
    pendingEstimatesBody.appendChild(tr);
  });
}

function renderClientAnalysis() {
  if (!clientAnalysisBody || !clientAnalysisEmpty || !clientAnalysisWrap) return;
  renderAnalyticsSection();
}

function renderReferralAnalysis() {
  if (!referralAnalysisBody || !referralAnalysisEmpty || !referralAnalysisWrap) return;
  renderAnalyticsSection();
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
  if (["work-templates", "workTemplates", "template", "templates", "業務テンプレート", "テンプレート"].includes(key)) return "work-templates";
  if (["sales", "売上"].includes(key)) return "sales";
  if (["expenses", "経費"].includes(key)) return "expenses";
  return "cases";
}

function renderDashboard() {
  if (!summaryGrid) return;
  state.selectedMonth = normalizeMonthKey(state.selectedMonth) || toMonthKey(new Date());
  state.selectedYear = normalizeYear(state.selectedYear) || new Date().getFullYear();
  state.selectedAggregation = ["all", "month", "year"].includes(state.selectedAggregation) ? state.selectedAggregation : "month";

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
  const isAllMode = state.selectedAggregation === "all";

  const sales = isAllMode ? sumValues(salesByMonth) : (isYearMode ? sumYearFromMonthlyMap(salesByMonth, state.selectedYear) : salesByMonth[state.selectedMonth] || 0);
  const expenses = isAllMode ? sumValues(expenseByMonth) : (isYearMode ? sumYearFromMonthlyMap(expenseByMonth, state.selectedYear) : expenseByMonth[state.selectedMonth] || 0);
  const workMinutes = isAllMode ? sumValues(reportByMonth) : (isYearMode ? sumYearFromMonthlyMap(reportByMonth, state.selectedYear) : reportByMonth[state.selectedMonth] || 0);
  const profit = sales - expenses;

  summaryGrid.innerHTML = "";
  const labelPrefix = isAllMode ? "全期間" : (isYearMode ? "年別" : "月別");
  const targetLabel = isAllMode ? "全期間" : (isYearMode ? `${state.selectedYear}年` : monthLabel(state.selectedMonth));
  [
    { label: isAllMode ? "対象期間" : (isYearMode ? "対象年" : "対象月"), value: targetLabel, cls: "" },
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
  renderAnalyticsSection();
  renderPendingEstimates();
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
        <td>${row.remainingAmountLabel || "-"}</td>
        <td>${row.dueDateLabel || formatDate(row.date)}</td>
        <td>${row.lastReminderDateLabel || "-"}</td>
        <td>${row.reminderCountLabel || "-"}</td>
        <td><button type="button" class="secondary-btn" data-task-target="${row.target}" data-task-id="${row.id}">編集</button></td>
        <td>${row.showReminderButton ? `<button type="button" class="secondary-btn" data-task-target="sale-reminder" data-task-id="${row.id}">督促記録</button>` : "-"}</td>
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
  getCollectionTargetSales()
    .filter((sale) => isReminderTaskTarget(sale))
    .forEach((sale) => {
      const linked = state.cases.find((entry) => entry.id === sale.caseId);
      const reminderStatus = getReminderTaskStatus(sale);
      const urgencyClass = reminderStatus === "overdue_no_reminder"
        ? "task-reminder-danger"
        : reminderStatus === "reminder_elapsed"
          ? "task-reminder-warning"
          : "task-reminder-upcoming";
      tasks.push({
        id: sale.id,
        target: "sale",
        type: sale.paymentStatus === "一部入金" ? "督促（一部入金）" : "督促（未入金）",
        customerName: linked?.customerName || "（削除済み顧客）",
        caseName: linked?.caseName || "（削除済み案件）",
        date: sale.dueDate || sale.paidDate || toDateString(sale.createdAt),
        urgencyClass,
        remainingAmountLabel: formatCurrency(getRemainingAmount(sale)),
        dueDateLabel: sale.dueDate ? formatDate(sale.dueDate) : "未設定",
        lastReminderDateLabel: sale.lastReminderDate ? formatDate(sale.lastReminderDate) : "未記録",
        reminderCountLabel: `${Number(sale.reminderCount || 0)}回`,
        showReminderButton: isReminderRecordableSale(sale),
      });
    });
  getBillingLeakCandidates().forEach((entry) => {
    tasks.push({
      id: entry.id, target: "case", type: "請求漏れ", customerName: entry.customerName, caseName: entry.caseName,
      action: "売上登録が未実施", date: entry.updatedAt ? toDateString(entry.updatedAt) : toDateString(new Date()), urgencyClass: "task-soon",
    });
  });
  state.dailyReports.forEach((entry) => {
    const nextDateTs = toDateOnlyTimestamp(entry.nextActionDate);
    if (!entry.nextActionDate || !Number.isFinite(nextDateTs) || nextDateTs > todayLimit) return;
    const linkedCase = state.cases.find((caseEntry) => caseEntry.id === entry.caseId);
    const foundClient = state.clients.find((client) => client.id === entry.clientId)
      || (linkedCase?.clientId ? state.clients.find((client) => client.id === linkedCase.clientId) : null);
    tasks.push({
      id: entry.id,
      target: "daily-report",
      type: entry.clientId ? "顧客対応" : "日報対応",
      customerName: foundClient?.name || "未設定",
      caseName: linkedCase?.caseName || "―",
      action: entry.nextAction || "次回対応",
      date: entry.nextActionDate,
      urgencyClass: nextDateTs <= todayLimit ? "task-overdue" : "task-soon",
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
  const unpaidSales = getCollectionTargetSales();
  const count = unpaidSales.length;
  const totalUnpaidAmount = unpaidSales.reduce((sum, sale) => sum + getRemainingAmount(sale), 0);
  const unpaidCount = unpaidSales.filter((sale) => sale.paymentStatus === "未入金").length;
  const partialCount = unpaidSales.filter((sale) => sale.paymentStatus === "一部入金").length;
  const overdueCount = unpaidSales.filter((sale) => isSaleOverdue(sale)).length;
  const filteredUnpaid = unpaidSales.filter((sale) => isWithinFilterDate(sale.paidDate || sale.createdAt, filter));
  const periodCount = filteredUnpaid.length;
  const periodAmount = filteredUnpaid.reduce((sum, sale) => sum + getRemainingAmount(sale), 0);

  if (unpaidAlertCard) unpaidAlertCard.classList.toggle("has-unpaid", count > 0);
  if (unpaidAlertSummary) {
    unpaidAlertSummary.textContent = count > 0
      ? `未入金 ${unpaidCount}件 / 一部入金 ${partialCount}件 / 期限超過 ${overdueCount}件 / 未回収合計 ${formatCurrency(totalUnpaidAmount)}`
      : "要対応の入金はありません。";
  }
  if (unpaidAlertPeriod) {
    unpaidAlertPeriod.textContent = filter.mode === "year"
      ? `対象年(${filter.year}年): ${periodCount}件 / ${formatCurrency(periodAmount)}`
      : filter.mode === "month"
        ? `対象月(${monthLabel(filter.monthKey || state.selectedMonth)}): ${periodCount}件 / ${formatCurrency(periodAmount)}`
        : `全期間: ${periodCount}件 / ${formatCurrency(periodAmount)}`;
  }
  if (unpaidAlertEmpty) unpaidAlertEmpty.hidden = count > 0;
  if (unpaidAlertListWrap) unpaidAlertListWrap.hidden = count === 0;
  if (!unpaidAlertBody) return;
  unpaidAlertBody.innerHTML = "";

  unpaidSales
    .slice()
    .sort((a, b) => (b.updatedAt ?? b.createdAt) - (a.updatedAt ?? a.createdAt))
    .forEach((sale) => {
      const linkedCase = state.cases.find((entry) => entry.id === sale.caseId);
      const linkedEstimate = sale.estimateId ? state.estimates.find((entry) => entry.id === sale.estimateId) : null;
      const customerName = linkedCase?.customerName || linkedEstimate?.customerName || "顧客不明";
      const caseName = linkedCase?.caseName || "案件なし";
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${escapeHtml(customerName)}</td>
        <td>${escapeHtml(caseName)}</td>
        <td>${formatCurrency(sale.invoiceAmount)}</td>
        <td>${formatCurrency(sale.paidAmount)}</td>
        <td>${formatCurrency(getRemainingAmount(sale))}</td>
        <td>${formatDate(sale.dueDate)}</td>
        <td><span class="${getSaleStatusClass(sale)}">${escapeHtml(sale.paymentStatus)}</span></td>
        <td>${getSaleDeadlineDaysLabel(sale)}</td>
        <td>${Number(sale.reminderCount || 0)}回</td>
        <td>${sale.lastReminderDate ? formatDate(sale.lastReminderDate) : "-"}</td>
        <td>${escapeHtml(sale.reminderMethod || "-")}</td>
        <td>${escapeHtml(sale.reminderMemo || "-")}</td>
        <td>
          <button type="button" class="secondary-btn edit-sale-btn" data-sale-id="${sale.id}">編集</button>
          ${isReminderRecordableSale(sale) ? `<button type="button" class="secondary-btn record-reminder-btn" data-sale-id="${sale.id}">督促記録</button>` : ""}
        </td>
      `;
      tr.classList.add(getSaleRowClass(sale));
      unpaidAlertBody.appendChild(tr);
    });
}

function renderUnpaidList() {
  const unpaidSales = getCollectionTargetSales().slice().sort((a, b) => (b.updatedAt ?? b.createdAt) - (a.updatedAt ?? a.createdAt));
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
      <td>${formatCurrency(getRemainingAmount(sale))}</td>
      <td>${formatDate(sale.dueDate)}</td>
      <td>${formatDate(sale.paidDate)}</td>
      <td><span class="${getSaleStatusClass(sale)}">${escapeHtml(sale.paymentStatus)}</span></td>
      <td>${getSaleDeadlineDaysLabel(sale)}</td>
      <td>${Number(sale.reminderCount || 0)}回</td>
      <td>${sale.lastReminderDate ? formatDate(sale.lastReminderDate) : "-"}</td>
      <td>${escapeHtml(sale.reminderMethod || "-")}</td>
      <td>${escapeHtml(sale.reminderMemo || "-")}</td>
      <td><button type="button" class="secondary-btn edit-sale-btn" data-sale-id="${sale.id}">編集</button> ${isReminderRecordableSale(sale) ? `<button type="button" class="secondary-btn record-reminder-btn" data-sale-id="${sale.id}">督促記録</button>` : ""}</td>
    `;
    tr.classList.add(getSaleRowClass(sale));
    unpaidListBody.appendChild(tr);
  });
}

function getActiveDashboardFilter() {
  if (state.selectedAggregation === "month") {
    return { mode: "month", monthKey: state.selectedMonth };
  }
  if (state.selectedAggregation === "year") {
    return { mode: "year", year: state.selectedYear };
  }
  return { mode: "all" };
}

function normalizeNameKey(value) {
  return String(value || "").trim().toLowerCase().replace(/\s+/g, " ");
}

function getClientRank(totalSales) {
  if (totalSales >= 500000) return "A";
  if (totalSales >= 100000) return "B";
  return "C";
}

function toSafeNumber(value) {
  const num = Number(value);
  return Number.isFinite(num) ? num : 0;
}

function normalizeReferralSourceLabel(value, fallback = "未設定") {
  const normalized = asTrimmedText(value);
  return normalized || fallback;
}

function calculateProfitMarginPercent(salesTotal, expenseTotal) {
  const safeSales = toSafeNumber(salesTotal);
  if (safeSales <= 0) return 0;
  return ((safeSales - toSafeNumber(expenseTotal)) / safeSales) * 100;
}

function classifyReferralProfitability(row) {
  if (!row || row.clientCount <= 0 || row.caseCount <= 0) return "データ不足：継続観察";
  if (row.profit < 0) return "赤字：対応範囲見直し";
  if (row.caseCount >= 3 && row.averagePrice < 50000) return "案件数多いが売上低い：メニュー見直し";
  if (row.salesTotal > 0 && row.profitMargin >= 0 && row.profitMargin < 30) return "売上あり利益低い：単価見直し";
  if (row.profitMargin >= 30) return "高利益：重点営業";
  return "データ不足：継続観察";
}

function resolveClientForCase(caseEntry, clientById, clientsByNameKey) {
  const byId = caseEntry.clientId ? clientById.get(caseEntry.clientId) : null;
  if (byId) return byId;
  const matched = clientsByNameKey.get(normalizeNameKey(caseEntry.customerName));
  return matched || null;
}

function buildClientAndReferralAnalytics(filter = {}) {
  const clientById = new Map(state.clients.map((client) => [client.id, client]));
  const knownClientNames = new Set(state.clients.map((client) => String(client.name || "")));
  const clientsByNameKey = new Map();
  state.clients.forEach((client) => {
    const key = normalizeNameKey(client.name);
    if (key && !clientsByNameKey.has(key)) clientsByNameKey.set(key, client);
  });
  const caseById = new Map(state.cases.map((entry) => [entry.id, entry]));
  const clientMap = new Map();

  const ensureClientRow = (name) => {
    const key = String(name || "未設定");
    if (!clientMap.has(key)) {
      clientMap.set(key, {
        clientName: key,
        referralSource: "未設定",
        caseCount: 0,
        salesTotal: 0,
        expenseTotal: 0,
        unpaidTotal: 0,
        lastCaseDate: "",
        lastCreatedCaseDate: null,
        lastContactDate: null,
      });
    }
    return clientMap.get(key);
  };

  state.cases.forEach((entry) => {
    const client = resolveClientForCase(entry, clientById, clientsByNameKey);
    const clientName = client?.name || entry.customerName || "未設定";
    const row = ensureClientRow(clientName);
    row.referralSource = client?.referralSource || row.referralSource || "未設定";
    const caseCreatedDate = toDateString(entry.createdAt);
    if (caseCreatedDate && (!row.lastCreatedCaseDate || toSortTimestamp(caseCreatedDate) > toSortTimestamp(row.lastCreatedCaseDate))) {
      row.lastCreatedCaseDate = caseCreatedDate;
    }
    const caseDateForCount = entry.receivedDate || caseCreatedDate;
    if (isWithinFilterDate(caseDateForCount, filter)) {
      row.caseCount += 1;
      if (!row.lastCaseDate || toSortTimestamp(caseDateForCount) > toSortTimestamp(row.lastCaseDate)) {
        row.lastCaseDate = caseDateForCount;
      }
    }
  });

  state.sales.forEach((sale) => {
    const linkedCase = caseById.get(sale.caseId);
    if (!isWithinFilterDate(sale.paidDate || sale.createdAt, filter)) return;
    const client = linkedCase ? resolveClientForCase(linkedCase, clientById, clientsByNameKey) : null;
    const clientName = client?.name || linkedCase?.customerName || "不明";
    const row = ensureClientRow(clientName);
    row.referralSource = normalizeReferralSourceLabel(client?.referralSource, linkedCase ? "未設定" : "不明");
    row.salesTotal += toSafeNumber(sale.invoiceAmount);
    row.unpaidTotal += toSafeNumber(getRemainingAmount(sale));
  });

  state.expenses.forEach((expense) => {
    const linkedCase = caseById.get(expense.caseId);
    if (!isWithinFilterDate(expense.date, filter)) return;
    const client = linkedCase ? resolveClientForCase(linkedCase, clientById, clientsByNameKey) : null;
    const clientName = client?.name || linkedCase?.customerName || "不明";
    const row = ensureClientRow(clientName);
    row.referralSource = normalizeReferralSourceLabel(client?.referralSource, linkedCase ? "未設定" : "不明");
    row.expenseTotal += toSafeNumber(expense.amount);
  });

  state.dailyReports.forEach((report) => {
    const linkedCase = report.caseId ? caseById.get(report.caseId) : null;
    const directClient = report.clientId ? clientById.get(report.clientId) : null;
    const caseClient = linkedCase ? resolveClientForCase(linkedCase, clientById, clientsByNameKey) : null;
    const clientName = directClient?.name || caseClient?.name || linkedCase?.customerName || "未設定";
    const row = ensureClientRow(clientName);
    const reportDate = report.reportDate || null;
    if (reportDate && (!row.lastContactDate || toSortTimestamp(reportDate) > toSortTimestamp(row.lastContactDate))) {
      row.lastContactDate = reportDate;
    }
  });

  const clientRows = Array.from(clientMap.values())
    .map((row) => ({
      ...row,
      salesTotal: toSafeNumber(row.salesTotal),
      expenseTotal: toSafeNumber(row.expenseTotal),
      unpaidTotal: toSafeNumber(row.unpaidTotal),
      referralSource: normalizeReferralSourceLabel(row.referralSource),
      profit: toSafeNumber(row.salesTotal) - toSafeNumber(row.expenseTotal),
      rank: getClientRank(toSafeNumber(row.salesTotal)),
      lastContactDate: row.lastContactDate || row.lastCreatedCaseDate || null,
    }))
    .filter((row) => row.caseCount > 0 || row.salesTotal > 0 || row.expenseTotal > 0 || row.unpaidTotal > 0)
    .sort((a, b) => {
      const rankOrder = { A: 0, B: 1, C: 2 };
      const rankDiff = (rankOrder[a.rank] ?? 9) - (rankOrder[b.rank] ?? 9);
      if (rankDiff !== 0) return rankDiff;
      if (b.salesTotal !== a.salesTotal) return b.salesTotal - a.salesTotal;
      return a.clientName.localeCompare(b.clientName, "ja");
    });

  const referralMap = new Map();
  state.clients.forEach((client) => {
    const key = normalizeReferralSourceLabel(client?.referralSource);
    if (!referralMap.has(key)) {
      referralMap.set(key, { referralSource: key, clientCount: 0, caseCount: 0, salesTotal: 0, expenseTotal: 0, unpaidTotal: 0 });
    }
    referralMap.get(key).clientCount += 1;
  });
  clientRows.forEach((row) => {
    const key = normalizeReferralSourceLabel(row.referralSource);
    if (!referralMap.has(key)) {
      referralMap.set(key, { referralSource: key, clientCount: 0, caseCount: 0, salesTotal: 0, expenseTotal: 0, unpaidTotal: 0 });
    }
    const current = referralMap.get(key);
    if (!knownClientNames.has(String(row.clientName || ""))) current.clientCount += 1;
    current.caseCount += row.caseCount;
    current.salesTotal += row.salesTotal;
    current.expenseTotal += row.expenseTotal;
    current.unpaidTotal += row.unpaidTotal;
  });

  const referralRows = Array.from(referralMap.values())
    .map((row) => {
      const salesTotal = toSafeNumber(row.salesTotal);
      const expenseTotal = toSafeNumber(row.expenseTotal);
      const caseCount = toSafeNumber(row.caseCount);
      const profit = salesTotal - expenseTotal;
      const averagePrice = caseCount > 0 ? salesTotal / caseCount : 0;
      const profitMargin = calculateProfitMarginPercent(salesTotal, expenseTotal);
      const rowWithMetrics = {
        ...row,
        salesTotal,
        expenseTotal,
        caseCount,
        unpaidTotal: toSafeNumber(row.unpaidTotal),
        profit,
        averagePrice,
        profitMargin,
      };
      return {
        ...rowWithMetrics,
        comment: classifyReferralProfitability(rowWithMetrics),
      };
    })
    .sort((a, b) => b.profit - a.profit);

  return { clientRows, referralRows };
}

function renderAnalyticsSection() {
  if (!clientAnalysisBody || !referralAnalysisBody || !analyticsRankingGrid) return;
  const filter = getActiveDashboardFilter();
  const { clientRows, referralRows } = buildClientAndReferralAnalytics(filter);
  if (analysisPeriodLabel) {
    analysisPeriodLabel.textContent = `対象期間: ${filter.mode === "year" ? `${filter.year}年` : filter.mode === "month" ? monthLabel(filter.monthKey) : "全期間"}`;
  }

  clientAnalysisBody.innerHTML = "";
  if (clientAnalysisWrap) clientAnalysisWrap.hidden = clientRows.length === 0;
  if (clientAnalysisEmpty) clientAnalysisEmpty.hidden = clientRows.length > 0;
  clientRows.forEach((row) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(row.clientName)}</td>
      <td><span class="client-rank-badge rank-${row.rank}">${escapeHtml(row.rank)}</span></td>
      <td>${row.caseCount}件</td>
      <td>${formatCurrency(row.salesTotal)}</td>
      <td>${formatCurrency(row.expenseTotal)}</td>
      <td class="${row.profit < 0 ? "loss-text" : ""}">${formatCurrency(row.profit)}</td>
      <td class="${row.unpaidTotal > 0 ? "analysis-unpaid-text" : ""}">${formatCurrency(row.unpaidTotal)}${row.unpaidTotal > 0 ? " ⚠️" : ""}</td>
      <td>${formatDate(row.lastContactDate)}</td>
    `;
    clientAnalysisBody.appendChild(tr);
  });

  referralAnalysisBody.innerHTML = "";
  if (referralAnalysisWrap) referralAnalysisWrap.hidden = referralRows.length === 0;
  if (referralAnalysisEmpty) referralAnalysisEmpty.hidden = referralRows.length > 0;
  referralRows.forEach((row) => {
    const marginClass = row.profit < 0 ? "referral-margin-negative" : row.profitMargin >= 30 ? "referral-margin-good" : "referral-margin-normal";
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(row.referralSource || "未設定")}</td>
      <td>${row.clientCount}件</td>
      <td>${row.caseCount}件</td>
      <td>${formatCurrency(row.salesTotal)}</td>
      <td>${formatCurrency(row.expenseTotal)}</td>
      <td class="${row.profit < 0 ? "loss-text" : ""}">${formatCurrency(row.profit)}</td>
      <td class="${row.unpaidTotal > 0 ? "analysis-unpaid-text" : ""}">${formatCurrency(row.unpaidTotal)}${row.unpaidTotal > 0 ? " ⚠️" : ""}</td>
      <td>${formatCurrency(row.averagePrice)}</td>
      <td class="${marginClass}">${row.profitMargin.toFixed(1)}%</td>
      <td>${escapeHtml(row.comment)}</td>
    `;
    referralAnalysisBody.appendChild(tr);
  });

  renderAnalyticsRankings(clientRows, referralRows);
  renderImportantClients(clientRows);
}

function renderImportantClients(clientRows) {
  if (!importantClientsBody || !importantClientsEmpty || !importantClientsWrap) return;
  const rows = (clientRows || []).filter((row) => row.rank === "A").slice(0, 5);
  importantClientsBody.innerHTML = "";
  importantClientsWrap.hidden = rows.length === 0;
  importantClientsEmpty.hidden = rows.length > 0;
  rows.forEach((row) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(row.clientName)}</td>
      <td>${formatCurrency(row.salesTotal)}</td>
      <td class="${row.profit < 0 ? "loss-text" : ""}">${formatCurrency(row.profit)}</td>
      <td>${formatDate(row.lastContactDate)}</td>
    `;
    importantClientsBody.appendChild(tr);
  });
}

function renderAnalyticsRankings(clientRows, referralRows) {
  if (!analyticsRankingGrid) return;
  analyticsRankingGrid.innerHTML = "";
  const rankingDefs = [
    { label: "売上上位顧客", rows: clientRows.slice().sort((a, b) => b.salesTotal - a.salesTotal), accessor: (row) => row.salesTotal },
    { label: "利益上位顧客", rows: clientRows.slice().sort((a, b) => b.profit - a.profit), accessor: (row) => row.profit },
    { label: "未入金額上位顧客", rows: clientRows.slice().sort((a, b) => b.unpaidTotal - a.unpaidTotal), accessor: (row) => row.unpaidTotal },
    { label: "紹介元別利益ランキング", rows: referralRows.slice().sort((a, b) => b.profit - a.profit), accessor: (row) => row.profit, referral: true },
  ];
  rankingDefs.forEach((def) => {
    const wrap = document.createElement("section");
    wrap.className = "summary-card analytics-ranking-card";
    const topRows = def.rows.filter((row) => def.accessor(row) > 0).slice(0, 5);
    wrap.innerHTML = `<p class="label">${def.label}</p>`;
    if (!topRows.length) {
      wrap.innerHTML += `<p class="meta">データなし</p>`;
    } else {
      const ol = document.createElement("ol");
      ol.className = "analytics-ranking-list";
      topRows.forEach((row, idx) => {
        const value = def.accessor(row);
        const name = def.referral ? (row.referralSource || "未設定") : row.clientName;
        const li = document.createElement("li");
        li.className = value < 0 ? "loss-text" : "";
        li.textContent = `${idx + 1}. ${name}：${formatCurrency(value)}`;
        ol.appendChild(li);
      });
      wrap.appendChild(ol);
    }
    analyticsRankingGrid.appendChild(wrap);
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

function getCollectionTargetSales() {
  return state.sales.filter((sale) => {
    return sale.paymentStatus !== "入金済"
      || sale.isUnpaid === true
      || Number(sale.paidAmount || 0) < Number(sale.invoiceAmount || 0);
  });
}

function getRemainingAmount(sale) {
  return Math.max((sale.invoiceAmount ?? 0) - (sale.paidAmount ?? 0), 0);
}

function renderCaseOptions() {
  if (!saleCaseSelect || !expenseCaseSelect || !reportCaseSelect) return;
  const options = state.cases
    .slice()
    .sort((a, b) => a.createdAt - b.createdAt)
    .map((c) => `<option value="${c.id}">${escapeHtml(c.customerName)}｜${escapeHtml(c.caseName)}</option>`)
    .join("");

  saleCaseSelect.innerHTML = `<option value="">案件なし</option>${options}`;
  saleCaseSelect.disabled = false;
  expenseCaseSelect.innerHTML = `<option value="">案件に紐付けしない</option>${options}`;
  reportCaseSelect.innerHTML = `<option value="">案件なし</option>${options}`;
}

function renderWorkTemplateOptions() {
  if (!caseTemplateSelect) return;
  const currentValue = caseTemplateSelect.value;
  const options = state.workTemplates
    .slice()
    .sort((a, b) => (b.updatedAt ?? b.createdAt) - (a.updatedAt ?? a.createdAt))
    .map((entry) => `<option value="${entry.id}">${escapeHtml(entry.name)}${Number.isFinite(entry.defaultDueDays) ? `（${entry.defaultDueDays}日）` : ""}</option>`)
    .join("");
  caseTemplateSelect.innerHTML = `<option value="">テンプレートを選択してください</option>${options}`;
  if (currentValue) caseTemplateSelect.value = currentValue;
}

function renderClientOptions() {
  if (!caseClientSelect && !estimateClientSelect && !reportClientSelect && !clientHistoryClientSelect) return;
  const options = state.clients
    .slice()
    .sort((a, b) => (b.updatedAt ?? b.createdAt) - (a.updatedAt ?? a.createdAt))
    .map((client) => `<option value="${client.id}">${escapeHtml(client.name)}${client.clientType ? `（${escapeHtml(client.clientType)}）` : ""}</option>`)
    .join("");
  if (caseClientSelect) caseClientSelect.innerHTML = `<option value="">選択しない</option>${options}`;
  if (estimateClientSelect) estimateClientSelect.innerHTML = `<option value="">選択しない</option>${options}`;
  if (reportClientSelect) reportClientSelect.innerHTML = `<option value="">選択しない</option>${options}`;
  if (clientHistoryClientSelect) clientHistoryClientSelect.innerHTML = `<option value="">顧客を選択してください</option>${options}`;
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

function syncDailyReportClientFromCase() {
  if (!reportCaseSelect || !reportClientSelect) return;
  const caseId = reportCaseSelect.value;
  if (!caseId) return;
  const found = state.cases.find((entry) => entry.id === caseId);
  if (found?.clientId) reportClientSelect.value = found.clientId;
}

function syncDailyReportClientLabel() {
  if (!reportClientSelect) return;
  const selectedClientId = reportClientSelect.value;
  const found = selectedClientId ? state.clients.find((entry) => entry.id === selectedClientId) : null;
  reportClientSelect.title = found ? `選択中: ${found.name}` : "";
}

function renderClients() {
  if (!clientsList || !clientsEmpty) return;
  clientsList.innerHTML = "";
  const caseById = new Map(state.cases.map((entry) => [entry.id, entry]));
  const todayTs = toDateOnlyTimestamp(toDateString(new Date()));
  const sorted = state.clients.slice().sort((a, b) => (b.updatedAt ?? b.createdAt) - (a.updatedAt ?? a.createdAt));
  sorted.forEach((client) => {
    const related = state.dailyReports.filter((entry) => {
      if (entry.clientId === client.id) return true;
      const linkedCase = entry.caseId ? caseById.get(entry.caseId) : null;
      if (!linkedCase) return false;
      if (linkedCase?.clientId === client.id) return true;
      return linkedCase?.customerName === client.name;
    });
    const lastInteraction = related
      .map((entry) => entry.reportDate)
      .filter(Boolean)
      .sort((a, b) => toSortTimestamp(b) - toSortTimestamp(a))[0] || "";
    const nextCandidates = related
      .map((entry) => entry.nextActionDate)
      .filter(Boolean);
    const upcomingNext = nextCandidates
      .filter((dateText) => toDateOnlyTimestamp(dateText) >= todayTs)
      .sort((a, b) => toSortTimestamp(a) - toSortTimestamp(b))[0] || "";
    const nearestAny = nextCandidates.sort((a, b) => toSortTimestamp(a) - toSortTimestamp(b))[0] || "";
    const nextInteraction = upcomingNext || nearestAny;
    const li = document.createElement("li");
    li.className = "item";
    li.dataset.id = client.id;
    li.innerHTML = `
      <div>
        <p class="title">${escapeHtml(client.name)}${client.clientType ? `（${escapeHtml(client.clientType)}）` : ""}</p>
        <p class="meta">住所: ${escapeHtml(client.address || "未設定")} / 電話: ${escapeHtml(client.tel || "未設定")} / メール: ${escapeHtml(client.email || "未設定")}</p>
        <p class="meta">紹介元: ${escapeHtml(client.referralSource || "未設定")} / メモ: ${escapeHtml(truncateText(client.memo || "未設定", 60))}</p>
        <p class="meta">最終対応日: ${lastInteraction ? formatDate(lastInteraction) : "未設定"} / 次回対応日: ${nextInteraction ? formatDate(nextInteraction) : "未設定"} / 対応履歴件数: ${related.length}件</p>
      </div>
      <div class="row-actions"><button type="button" class="secondary-btn edit-client-btn">編集</button><button type="button" class="danger-btn delete-client-btn">削除</button></div>
    `;
    clientsList.appendChild(li);
  });
  clientsEmpty.hidden = sorted.length > 0;
}

function renderClientHistory() {
  if (!clientHistoryBody || !clientHistoryWrap || !clientHistoryEmpty) return;
  const selectedClientId = clientHistoryClientSelect?.value || "";
  const filtered = selectedClientId
    ? state.dailyReports.filter((entry) => entry.clientId === selectedClientId)
    : [];
  const sorted = filtered.slice().sort((a, b) => toSortTimestamp(b.reportDate) - toSortTimestamp(a.reportDate));
  clientHistoryBody.innerHTML = "";
  sorted.forEach((entry) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${formatDate(entry.reportDate)}</td>
      <td>${escapeHtml(resolveDailyReportCaseName(entry.caseId))}</td>
      <td>${escapeHtml(entry.interactionType || "作業")}</td>
      <td>${escapeHtml(truncateText(entry.workContent || "", 80))}</td>
      <td>${escapeHtml(truncateText(entry.nextAction || "", 60))}</td>
      <td>${entry.nextActionDate ? formatDate(entry.nextActionDate) : "未設定"}</td>
      <td>${escapeHtml(truncateText(entry.memo || "", 60))}</td>
    `;
    clientHistoryBody.appendChild(tr);
  });
  clientHistoryWrap.hidden = sorted.length === 0;
  clientHistoryEmpty.hidden = sorted.length > 0;
  if (!selectedClientId) clientHistoryEmpty.textContent = "顧客を選択すると履歴を表示します。";
  else if (!sorted.length) clientHistoryEmpty.textContent = "選択した顧客の日報履歴はまだありません。";
}

function renderWorkTemplates() {
  if (!workTemplatesList || !workTemplatesEmpty || !workTemplateItemTemplate) return;
  workTemplatesList.innerHTML = "";
  const sorted = state.workTemplates.slice().sort((a, b) => (b.updatedAt ?? b.createdAt) - (a.updatedAt ?? a.createdAt));
  sorted.forEach((entry) => {
    const node = workTemplateItemTemplate.content.cloneNode(true);
    const item = node.querySelector(".item");
    const title = node.querySelector(".title");
    const metas = node.querySelectorAll(".meta");
    item.dataset.id = entry.id;
    title.textContent = entry.name;
    if (metas[0]) metas[0].textContent = `標準期限: ${Number.isFinite(entry.defaultDueDays) ? `${entry.defaultDueDays}日後` : "未設定"} / 必要書類: ${truncateText(entry.requiredDocuments || "未設定", 80)}`;
    if (metas[1]) metas[1].textContent = `標準タスク: ${truncateText(entry.defaultTasks || "未設定", 90)} / メモ: ${truncateText(entry.memo || "未設定", 60)}`;
    workTemplatesList.appendChild(node);
  });
  workTemplatesEmpty.hidden = sorted.length > 0;
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
    card.dataset.id = entry.id;
    card.innerHTML = `
      <header class="daily-report-card-header">
        <div>
          <p class="daily-report-card-date">${formatDate(entry.reportDate)}</p>
          <p class="daily-report-card-case">顧客: ${escapeHtml(resolveDailyReportClientName(entry.clientId, entry.caseId))}</p>
          <p class="daily-report-card-case">${escapeHtml(resolveDailyReportCaseName(entry.caseId))}</p>
          <p class="daily-report-card-case">対応種別: ${escapeHtml(entry.interactionType || "作業")}</p>
        </div>
        <div class="daily-report-card-actions">
          <button type="button" class="secondary-btn edit-daily-report-btn">編集</button>
          <button type="button" class="danger-btn delete-daily-report-btn">削除</button>
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
        <div class="daily-report-field-inline">
          <dt>次回対応日</dt>
          <dd>${entry.nextActionDate ? formatDate(entry.nextActionDate) : "未設定"}</dd>
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
    resolveDailyReportClientName(entry.clientId, entry.caseId),
    entry.interactionType,
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
    <input type="text" inputmode="decimal" pattern="[0-9.,]*" data-key="quantity" placeholder="数量" value="${defaultItem.quantity ?? 1}" />
    <input type="text" inputmode="numeric" pattern="[0-9,]*" data-key="unitPrice" placeholder="単価" value="${defaultItem.unitPrice ?? 0}" />
    <p class="meta item-amount">${formatCurrency(defaultItem.amount ?? 0)}</p>
    <button type="button" class="danger-btn estimate-item-remove-btn">削除</button>
  `;
  estimateItemsWrap.appendChild(row);
  bindCommaInput(row.querySelector('[data-key="unitPrice"]'));
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
      const quantity = parseDecimalInput(row.querySelector('[data-key="quantity"]')?.value);
      const unitPrice = parseNumberInput(row.querySelector('[data-key="unitPrice"]')?.value);
      const amount = Math.floor((Number.isFinite(quantity) ? quantity : 0) * unitPrice);
      return { itemName, quantity, unitPrice, amount, sortOrder: idx };
    })
    .filter((item) => item.itemName);
}

function recalcEstimateTotals() {
  let subtotal = 0;
  Array.from(estimateItemsWrap?.querySelectorAll(".estimate-item-row") || []).forEach((row) => {
    const quantity = parseDecimalInput(row.querySelector('[data-key="quantity"]')?.value);
    const unitPrice = parseNumberInput(row.querySelector('[data-key="unitPrice"]')?.value);
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
  console.log("ESTIMATE SUBMIT FIRED");
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
  const actionName = taskName;
  console.log("EDIT STATE", editState);
  console.log("ACTION START", actionName, editState.estimateId || "new");
  console.log("PAYLOAD", payload);
  try {
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
      console.log("DB DONE", actionName, estimateId);
      await refreshAfterMutation(taskName, isUpdate ? "見積を更新しました。" : "見積を登録しました。");
      resetEstimateForm();
      editState.estimateId = null;
      subtabState.estimates = "list";
      activateTab("estimates");
    }, { triggerButton: event.submitter });
  } catch (error) {
    showAppMessage(`見積保存に失敗しました。${formatSupabaseError(error)}`, true);
  } finally {
    console.log("ACTION FINALLY", actionName);
    forceHideLoading();
  }
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
    const isCaseCreated = state.cases.some((c) => c.estimateId === entry.id) || Boolean(entry.caseId);
    const hasCreatedInvoice = state.sales.some((sale) => sale.estimateId === entry.id);
    const li = document.createElement("li");
    li.className = "item estimate-card";
    li.dataset.id = entry.id;
    const itemCount = state.estimateItems.filter((row) => row.estimateId === entry.id).length;
    const casedLabel = isCaseCreated ? "案件化済み" : "未案件化";
    li.innerHTML = `
      <div class="estimate-card-body">
        <p class="title estimate-card-title">${escapeHtml(entry.estimateTitle)}</p>
        <div class="estimate-card-grid">
          <p class="meta"><span>顧客名:</span> ${escapeHtml(entry.customerName)}</p>
          <p class="meta"><span>見積番号:</span> ${escapeHtml(entry.estimateNumber || "未採番")}</p>
          <p class="meta"><span>見積日:</span> ${formatDate(entry.estimateDate)}</p>
          <p class="meta"><span>有効期限:</span> ${formatDate(entry.validUntil)}</p>
          <p class="meta"><span>ステータス:</span> ${escapeHtml(entry.status)}</p>
          <p class="meta"><span>合計金額:</span> ${formatCurrency(entry.totalAmount ?? entry.total ?? 0)}</p>
          <p class="meta"><span>明細件数:</span> ${itemCount}件</p>
          <p class="meta"><span>案件化:</span> ${casedLabel}</p>
        </div>
      </div>
      <div class="row-actions estimate-card-actions">
        <button type="button" class="secondary-btn estimate-edit-btn">編集</button>
        <button type="button" class="danger-btn estimate-delete-btn">削除</button>
        <button type="button" class="secondary-btn create-case-btn" ${isCaseCreated ? "disabled" : ""}>${isCaseCreated ? "案件化済み" : "案件化"}</button>
        <button type="button" class="secondary-btn estimate-estimate-print-btn">見積書出力</button>
        <button type="button" class="secondary-btn estimate-print-btn">請求書出力</button>
        <button type="button" class="secondary-btn estimate-estimate-xlsx-btn">見積Excel</button>
        <button type="button" class="secondary-btn estimate-xlsx-btn">請求Excel</button>
        <button type="button" class="secondary-btn create-invoice-btn" ${hasCreatedInvoice ? "disabled" : ""}>${hasCreatedInvoice ? "請求済み" : "請求作成"}</button>
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
  const button = event.target.closest("button");
  if (!(button instanceof HTMLButtonElement) || !currentUser) return;
  const item = button.closest("[data-id]");
  const estimateId = item?.dataset.id;
  console.log("ESTIMATE LIST BUTTON CLICKED", button.className, estimateId);
  if (!estimateId) {
    showAppMessage("対象データIDを取得できませんでした。", true);
    return;
  }
  if (button.classList.contains("estimate-edit-btn")) return startEstimateEdit(estimateId);
  if (button.classList.contains("estimate-delete-btn")) {
    await deleteEstimate(estimateId);
    return;
  }
  if (button.classList.contains("create-case-btn")) {
    console.log("CREATE CASE CLICKED", estimateId);
    await handleCreateCaseFromEstimate(estimateId);
    return;
  }
  if (button.classList.contains("estimate-estimate-print-btn")) return openEstimatePrintPreview(estimateId);
  if (button.classList.contains("estimate-print-btn")) return openInvoicePrintPreviewFromEstimate(estimateId);
  if (button.classList.contains("estimate-estimate-xlsx-btn")) return exportEstimateDataForEstimate(estimateId);
  if (button.classList.contains("estimate-xlsx-btn")) return exportInvoiceDataForEstimate(estimateId);
  if (button.classList.contains("create-invoice-btn") || button.classList.contains("estimate-create-invoice-btn")) {
    console.log("CREATE INVOICE CLICKED", estimateId);
    await handleCreateInvoiceFromEstimate(estimateId);
  }
}

async function handleEstimateConversionAction(options) {
  const {
    estimateId,
    loadingLabel,
    actionLabel,
    execute,
    successMessage,
    successTab,
    successSubtab,
  } = options;
  if (!currentUser) {
    showAppMessage("ログイン状態を確認できません。", true);
    return null;
  }
  if (!estimateId) {
    showAppMessage("対象見積IDを取得できませんでした。", true);
    return null;
  }
  let isSuccess = false;
  startLoading(loadingLabel);
  try {
    clearAppMessage();
    const result = await execute();
    isSuccess = true;
    return result;
  } catch (error) {
    console.error(`${actionLabel}に失敗しました。`, error);
    if (["DUPLICATE_ESTIMATE_CASE", "DUPLICATE_ESTIMATE_INVOICE"].includes(error?.code)) {
      showAppMessage(error.message || "すでに作成済みです。", true);
    } else {
      showAppMessage(`${actionLabel}に失敗しました。${formatSupabaseError(error)}`, true);
    }
    return null;
  } finally {
    if (isSuccess && successMessage) {
      await loadAllData();
      renderAfterDataChanged();
      if (successTab) activateTab(successTab);
      if (successTab && successSubtab) activateSubtab(successTab, successSubtab);
      showAppMessage(successMessage, false);
    }
    forceHideLoading();
  }
}

async function createCaseFromEstimate(estimateId) {
  const estimate = state.estimates.find((entry) => entry.id === estimateId);
  if (!estimate) throw new Error("対象見積が見つかりません。");
  const existingCase = state.cases.find((c) => c.estimateId === estimateId || c.id === estimate.caseId);
  if (existingCase) {
    const duplicatedError = new Error("この見積はすでに案件化済みです。");
    duplicatedError.code = "DUPLICATE_ESTIMATE_CASE";
    throw duplicatedError;
  }
  const customerName = estimate.customerName || estimate.customer_name || "顧客不明";
  const caseName = estimate.subject || estimate.caseName || estimate.title || estimate.estimateTitle || "見積から作成した案件";
  const estimateAmount = Number(estimate.totalAmount || estimate.total || estimate.amount || estimate.estimateAmount || 0);
  const rawPayload = {
    user_id: currentUser.id,
    client_id: estimate.clientId || null,
    estimate_id: estimate.id,
    customer_name: customerName,
    case_name: caseName,
    estimate_amount: estimateAmount,
    received_date: new Date().toISOString().slice(0, 10),
    due_date: null,
    status: "未着手",
    work_memo: "見積から案件化",
    next_action_date: null,
    next_action: null,
    template_id: null,
    required_documents: null,
    task_list: null,
    document_url: null,
    invoice_url: null,
    receipt_url: null,
  };
  const payload = pickObjectKeys(rawPayload, CASE_MUTATION_COLUMNS);
  console.log("CREATE CASE PAYLOAD", payload);
  const { data, error } = await sbClient.from("cases").insert(payload).select().single();
  if (error) throw error;
  if (!data) throw new Error("案件化の登録結果を取得できませんでした。");
  const updateRes = await sbClient.from("estimates").update({ case_id: data.id }).eq("id", estimateId).eq("user_id", currentUser.id);
  if (updateRes.error) throw updateRes.error;
  return data;
}

async function handleCreateCaseFromEstimate(estimateId) {
  console.log("CREATE CASE START", estimateId);
  return handleEstimateConversionAction({
    estimateId,
    loadingLabel: "案件化",
    actionLabel: "案件化",
    successMessage: "見積から案件を作成しました。",
    successTab: "cases",
    successSubtab: "list",
    execute: async () => {
      const data = await createCaseFromEstimate(estimateId);
      console.log("CREATE CASE SUCCESS", data);
      return data;
    },
  });
}

async function createInvoiceFromEstimate(estimateId) {
  const estimate = state.estimates.find((entry) => entry.id === estimateId);
  if (!estimate) throw new Error("対象見積が見つかりません。");
  const existingSale = state.sales.find((sale) => sale.estimateId === estimateId);
  if (existingSale) {
    const duplicatedError = new Error("この見積はすでに請求作成済みです。");
    duplicatedError.code = "DUPLICATE_ESTIMATE_INVOICE";
    throw duplicatedError;
  }
  const invoiceAmount = Number(estimate.totalAmount || estimate.total || estimate.amount || estimate.estimateAmount || 0);
  if (!invoiceAmount || invoiceAmount <= 0) throw new Error("見積金額が0円のため請求を作成できません。");
  const rawPayload = {
    user_id: currentUser.id,
    estimate_id: estimate.id,
    case_id: estimate.caseId || null,
    invoice_amount: invoiceAmount,
    paid_amount: 0,
    paid_date: null,
    due_date: addDaysToDate(new Date(), 7),
    payment_status: "未入金",
    is_unpaid: true,
    invoice_number: await generateInvoiceNumberIfNeeded(),
  };
  const payload = pickObjectKeys(rawPayload, SALES_MUTATION_COLUMNS);
  console.log("CREATE INVOICE PAYLOAD", payload);
  const { data, error } = await sbClient.from("sales").insert(payload).select().single();
  if (error) throw error;
  if (!data) throw new Error("請求作成結果を取得できませんでした。");
  return data;
}

async function handleCreateInvoiceFromEstimate(estimateId) {
  if (!currentUser) {
    showAppMessage("ログイン状態を確認できません。", true);
    return;
  }
  if (!estimateId) {
    showAppMessage("対象見積IDを取得できませんでした。", true);
    return;
  }
  console.log("CREATE INVOICE START", estimateId);
  return handleEstimateConversionAction({
    estimateId,
    loadingLabel: "請求作成",
    actionLabel: "請求作成",
    successMessage: "請求を作成しました。",
    successTab: "sales",
    successSubtab: "list",
    execute: async () => {
      const data = await createInvoiceFromEstimate(estimateId);
      console.log("CREATE INVOICE SUCCESS", data);
      return data;
    },
  });
}

async function deleteEstimate(id) {
  if (!currentUser) return;
  if (!id) {
    showAppMessage("削除対象の見積IDがありません。", true);
    return;
  }
  if (!window.confirm("この見積を削除しますか？")) return;
  const taskName = "見積削除";
  startLoading(taskName);
  try {
    clearAppMessage();
    console.log("ACTION START", taskName, id);

    const estimateItemsRes = await sbClient
      .from("estimate_items")
      .delete()
      .eq("estimate_id", id)
      .eq("user_id", currentUser.id);
    if (estimateItemsRes.error) throw estimateItemsRes.error;

    const estimateDeleteRes = await sbClient
      .from("estimates")
      .delete()
      .eq("id", id)
      .eq("user_id", currentUser.id);
    if (estimateDeleteRes.error) throw estimateDeleteRes.error;
    console.log("DB DONE", taskName, id);

    if (editState.estimateId === id) {
      editState.estimateId = null;
      resetEstimateForm();
    }
    await loadAllData();
    console.log("LOAD DONE", taskName);
    renderAfterDataChanged();
    console.log("RENDER DONE", taskName);
    showAppMessage("削除しました。", false);
  } catch (error) {
    console.error("削除に失敗しました。", error);
    showAppMessage(`削除に失敗しました。${error.message || ""}`, true);
  } finally {
    console.log("ACTION FINALLY", taskName);
    forceHideLoading();
  }
}


function renderCases() {
  if (!caseList || !caseEmpty || !caseItemTemplate) return;
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
    const templateName = state.workTemplates.find((template) => template.id === entry.templateId)?.name || "未設定";
    const customerName = entry.customerName || "顧客不明";
    const caseName = entry.caseName || "案件名未設定";
    const incompleteTasks = getIncompleteTaskCount(entry.taskList);

    item.dataset.id = entry.id;
    title.textContent = `${customerName}｜${caseName}`;
    const urlLinks = [
      entry.documentUrl ? `<a href="${escapeHtml(entry.documentUrl)}" target="_blank" rel="noopener noreferrer">関連書類を開く</a>` : "",
      entry.invoiceUrl ? `<a href="${escapeHtml(entry.invoiceUrl)}" target="_blank" rel="noopener noreferrer">請求書を開く</a>` : "",
      entry.receiptUrl ? `<a href="${escapeHtml(entry.receiptUrl)}" target="_blank" rel="noopener noreferrer">領収書を開く</a>` : "",
    ].filter(Boolean).join(" / ");
    meta.innerHTML = `見積: ${formatCurrency(entry.estimateAmount)} / ステータス: ${escapeHtml(entry.status)} / 受付日: ${formatDate(entry.receivedDate)} / 期限日: ${formatDate(entry.dueDate)} / 次回対応日: ${formatDate(entry.nextActionDate)} / 次回対応内容: ${escapeHtml(entry.nextAction || "未設定")}${urlLinks ? ` / ${urlLinks}` : ""}`;
    if (caseWorkMeta) {
      caseWorkMeta.textContent = `テンプレート: ${templateName} / 必要書類: ${truncateText(entry.requiredDocuments || "未設定", 50)} / タスク: ${truncateText(entry.taskList || "未設定", 50)} / 未完了タスク: ${incompleteTasks}件 / 作業メモ: ${truncateText(sanitizeLegacyEstimateMemo(entry.workMemo) || "未設定", 40)}`;
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
  if (!salesListBody || !salesEmpty || !salesListWrap) return;
  salesListBody.innerHTML = "";
  const sales = Array.isArray(state.sales) ? state.sales : [];
  const sorted = sales.slice().sort((a, b) => (b.updatedAt ?? b.createdAt ?? 0) - (a.updatedAt ?? a.createdAt ?? 0));
  const filteredSales = sorted.filter((sale) => {
    if (!state.salesSearchQuery) return true;
    return matchesSalesSearch(sale, state.salesSearchQuery);
  });

  filteredSales.forEach((sale) => {
    try {
      if (!sale || typeof sale !== "object") return;
      const invoiceAmount = Number(sale.invoiceAmount || 0);
      const paidAmount = Number(sale.paidAmount || 0);
      const remainingAmount = Math.max(0, invoiceAmount - paidAmount);
      const paymentStatus = sale.paymentStatus || calculatePaymentStatus(invoiceAmount, paidAmount);
      const linkedCase = sale.caseId ? state.cases.find((entry) => entry.id === sale.caseId) : null;
      const linkedClient = linkedCase?.clientId ? state.clients.find((entry) => entry.id === linkedCase.clientId) : null;
      const customerLabel = linkedClient?.name || linkedCase?.customerName || "顧客不明";
      const caseLabel = linkedCase?.caseName || linkedCase?.name || "案件なし";
      const dueDateLabel = sale.dueDate ? formatDate(sale.dueDate) : "未設定";
      const paidDateLabel = sale.paidDate ? formatDate(sale.paidDate) : "未設定";
      const safeSale = { ...sale, invoiceAmount, paidAmount, paymentStatus };
      const reminderCount = Number(sale.reminderCount) || 0;
      const reminderMethod = sale.reminderMethod || "-";
      const reminderMemo = sale.reminderMemo || "-";
      const lastReminderDateLabel = sale.lastReminderDate ? formatDate(sale.lastReminderDate) : "-";
      const canRecordReminder = isReminderRecordableSale(safeSale);
      const tr = document.createElement("tr");
      tr.dataset.id = sale.id || "";
      tr.classList.add(getSaleRowClass(safeSale));
      tr.innerHTML = `
        <td>${escapeHtml(customerLabel)}</td>
        <td>${escapeHtml(caseLabel)}<br /><small>${escapeHtml(sale.invoiceNumber || "未採番")}</small></td>
        <td>${formatCurrency(invoiceAmount)}</td>
        <td>${formatCurrency(paidAmount)}</td>
        <td>${formatCurrency(remainingAmount)}</td>
        <td><span class="${getSaleStatusClass(safeSale)}">${escapeHtml(paymentStatus)}</span></td>
        <td>${dueDateLabel}</td>
        <td>${paidDateLabel}</td>
        <td>${reminderCount}回</td>
        <td>${lastReminderDateLabel}</td>
        <td>${escapeHtml(reminderMethod)}</td>
        <td>${escapeHtml(reminderMemo)}</td>
        <td>${canRecordReminder ? '<button type="button" class="secondary-btn record-reminder-btn">督促記録</button>' : '-'}</td>
        <td><button type="button" class="edit-btn secondary-btn">編集</button></td>
        <td><button type="button" class="danger-btn delete-btn">削除</button></td>
      `;
      salesListBody.appendChild(tr);
    } catch (error) {
      console.error("renderSales item failed", sale, error);
    }
  });

  if (salesFilterCount) {
    salesFilterCount.textContent = `表示中 ${filteredSales.length}件 / 全${sorted.length}件`;
  }
  salesListWrap.hidden = filteredSales.length === 0;
  salesEmpty.hidden = filteredSales.length > 0;
  salesEmpty.textContent = filteredSales.length || !state.salesSearchQuery
    ? "売上データはまだありません。"
    : "条件に一致する売上データはありません。";
  console.log("RENDER SALES DONE");
}

function renderExpenses() {
  if (!expensesList || !expensesEmpty || !expenseItemTemplate) return;
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
  if (!fixedExpensesList || !fixedExpensesEmpty || !fixedExpenseItemTemplate) return;
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

function resolveDailyReportClientName(clientId, caseId = null) {
  const directClient = clientId ? state.clients.find((client) => client.id === clientId) : null;
  if (directClient?.name) return directClient.name;

  const linkedCase = caseId ? state.cases.find((entry) => entry.id === caseId) : null;
  if (linkedCase?.clientId) {
    const caseClient = state.clients.find((client) => client.id === linkedCase.clientId);
    if (caseClient?.name) return caseClient.name;
  }
  if (linkedCase?.customerName) return linkedCase.customerName;
  return "顧客なし";
}

async function startEstimateEdit(estimateId) {
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
  if (caseForm?.elements?.caseTemplateId) caseForm.elements.caseTemplateId.value = "";
  caseForm.elements.status.value = "未着手";
  caseSubmitBtn.textContent = "案件を追加";
}

function resetWorkTemplateForm() {
  resetEditMode("workTemplate");
  workTemplateForm?.reset();
  if (workTemplateSubmitBtn) workTemplateSubmitBtn.textContent = "テンプレートを登録";
}

function resetSaleForm() {
  resetEditMode("sale");
  saleForm.reset();
  setSaleInvoiceNumberDisplay("");
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
  if (reportInteractionTypeSelect) reportInteractionTypeSelect.value = "作業";
  if (reportClientSelect) reportClientSelect.value = "";
  dailyReportSubmitBtn.textContent = "日報を登録";
}

function resetEditMode(target) {
  if (target === "client") editState.clientId = null;
  if (target === "case") editState.caseId = null;
  if (target === "workTemplate") editState.workTemplateId = null;
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
  const created = await createCaseFromEstimate(estimateId);
  return created?.id || null;
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
  await createInvoiceFromEstimate(estimateId);
}

async function generateInvoiceNumberIfNeeded() {
  return getNextMonthlyNumber("sales", "invoice_number", "S");
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
  const year = Number(String(value ?? "").replace(/,/g, "").trim());
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

function sumValues(valueMap) {
  return Object.values(valueMap || {}).reduce((sum, value) => sum + (Number(value) || 0), 0);
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

function addDaysToDate(date, days) {
  const d = new Date(date);
  d.setDate(d.getDate() + days);
  return d.toISOString().slice(0, 10);
}

function normalizeAmount(raw) {
  if (raw === "" || raw === null || raw === undefined) return null;
  const normalized = String(raw).replace(/,/g, "").trim();
  if (!normalized) return null;
  const value = Number(normalized);
  if (!Number.isFinite(value) || value < 0) return null;
  return Math.floor(value);
}

function parseNumberInput(value) {
  return Number(String(value || "").replace(/,/g, "").trim()) || 0;
}

function formatNumberInput(value) {
  const num = String(value || "").replace(/[^\d]/g, "");
  if (!num) return "";
  return num.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}

function bindCommaInput(input) {
  if (!(input instanceof HTMLInputElement)) return;
  if (input.dataset.commaBound === "true") return;
  input.dataset.commaBound = "true";
  input.addEventListener("input", () => {
    const cursor = input.selectionStart ?? input.value.length;
    const beforeLength = input.value.length;
    input.value = formatNumberInput(input.value);
    const afterLength = input.value.length;
    const diff = afterLength - beforeLength;
    const nextCursor = Math.max(0, cursor + diff);
    input.setSelectionRange(nextCursor, nextCursor);
  });
  input.value = formatNumberInput(input.value);
}

function isCommaInputTarget(input) {
  if (!(input instanceof HTMLInputElement)) return false;
  if (input.type !== "text") return false;
  if (input.readOnly) return false;
  if (input.dataset.commaFormat === "true") return true;
  const key = `${input.id || ""} ${input.name || ""}`.toLowerCase();
  if (key.includes("invoice-number") || key.includes("invoicenumber")) return false;
  return /(amount|price|invoice|paid|total|unit)/.test(key);
}

function bindCommaInputFields(root = document) {
  const scope = root instanceof HTMLElement || root instanceof Document ? root : document;
  const targets = Array.from(scope.querySelectorAll('input[type="text"]')).filter(isCommaInputTarget);
  targets.forEach((input) => bindCommaInput(input));
}

function parseDecimalInput(raw) {
  if (raw === "" || raw === null || raw === undefined) return 0;
  const normalized = String(raw).replace(/,/g, "").trim();
  const value = Number(normalized);
  if (!Number.isFinite(value) || value < 0) return 0;
  return value;
}

function handleNumberInputWheel(event) {
  const activeElement = document.activeElement;
  if (!(activeElement instanceof HTMLInputElement)) return;
  if (activeElement.type !== "number") return;
  if (event.target !== activeElement) return;
  activeElement.blur();
}

function normalizeStatus(status) {
  return STATUS_ORDER.includes(status) ? status : "未着手";
}

function normalizeStoredStatus(status) {
  if (status === null || status === undefined) return "未着手";
  const normalized = String(status).trim();
  return normalized || "未着手";
}

function normalizeDailyReportInteractionType(type) {
  const normalized = String(type || "").trim();
  if (!normalized) return "作業";
  return DAILY_REPORT_INTERACTION_TYPES.includes(normalized) ? normalized : "その他";
}

function normalizeClientInteractionType(type) {
  return normalizeDailyReportInteractionType(type);
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

function calculatePaymentStatus(invoiceAmount, paidAmount) {
  const normalizedInvoice = Math.max(Number(invoiceAmount) || 0, 0);
  const normalizedPaid = Math.max(Number(paidAmount) || 0, 0);
  if (!normalizedPaid || normalizedPaid <= 0) return "未入金";
  if (normalizedPaid < normalizedInvoice) return "一部入金";
  return "入金済";
}

function getSalePaymentStatus(invoiceAmount, paidAmount) {
  return calculatePaymentStatus(invoiceAmount, paidAmount);
}

function normalizeSalePaymentStatus(value, invoiceAmount, paidAmount) {
  return value || calculatePaymentStatus(invoiceAmount, paidAmount);
}

function normalizeReminderMethod(value) {
  const normalized = asTrimmedText(value);
  if (!normalized) return "";
  return REMINDER_METHODS.includes(normalized) ? normalized : "その他";
}

function isReminderRecordableSale(sale) {
  const status = sale?.paymentStatus || calculatePaymentStatus(sale?.invoiceAmount || 0, sale?.paidAmount || 0);
  return status === "未入金" || status === "一部入金";
}

function getReminderTaskStatus(sale) {
  const dueTs = toDateOnlyTimestamp(sale?.dueDate);
  const todayTs = getTodayTimestamp();
  const lastReminderTs = toDateOnlyTimestamp(sale?.lastReminderDate);
  const hasReminder = Number(sale?.reminderCount || 0) > 0 || Boolean(sale?.lastReminderDate);
  const isOverdue = Number.isFinite(dueTs) && Number.isFinite(todayTs) && dueTs < todayTs;
  const reminderElapsed = Number.isFinite(lastReminderTs) && Number.isFinite(todayTs) && (todayTs - lastReminderTs) >= 7 * 86400000;

  if (isOverdue && !hasReminder) return "overdue_no_reminder";
  if (reminderElapsed) return "reminder_elapsed";
  return "before_due";
}

function isReminderTaskTarget(sale) {
  if (!isReminderRecordableSale(sale)) return false;
  const dueTs = toDateOnlyTimestamp(sale?.dueDate);
  const todayTs = getTodayTimestamp();
  const lastReminderTs = toDateOnlyTimestamp(sale?.lastReminderDate);
  const hasOverdueDueDate = Number.isFinite(dueTs) && Number.isFinite(todayTs) && dueTs < todayTs;
  const reminderElapsed = Number.isFinite(lastReminderTs) && Number.isFinite(todayTs) && (todayTs - lastReminderTs) >= 7 * 86400000;
  return hasOverdueDueDate || reminderElapsed;
}

function isSaleOverdue(sale) {
  if (!sale?.dueDate || sale.paymentStatus === "入金済") return false;
  const dueTimestamp = toDateOnlyTimestamp(sale.dueDate);
  const todayTimestamp = toDateOnlyTimestamp(new Date());
  if (!Number.isFinite(dueTimestamp) || !Number.isFinite(todayTimestamp)) return false;
  return dueTimestamp < todayTimestamp;
}

function getSaleStatusClass(sale) {
  if (isSaleOverdue(sale)) return "sale-status-overdue";
  if (sale.paymentStatus === "一部入金") return "sale-status-partial";
  if (sale.paymentStatus === "入金済") return "sale-status-paid";
  return "sale-status-unpaid";
}

function getSaleRowClass(sale) {
  if (isSaleOverdue(sale)) return "sale-row-overdue";
  if (sale.paymentStatus === "一部入金") return "sale-row-partial";
  if (sale.paymentStatus === "未入金") return "sale-row-unpaid";
  return "";
}

function getSaleDeadlineDaysLabel(sale) {
  if (!sale?.dueDate) return "-";
  const dueTimestamp = toDateOnlyTimestamp(sale.dueDate);
  const todayTimestamp = getTodayTimestamp();
  if (!Number.isFinite(dueTimestamp) || !Number.isFinite(todayTimestamp)) return "-";
  const diffDays = Math.floor((dueTimestamp - todayTimestamp) / 86400000);
  return diffDays < 0 ? `${Math.abs(diffDays)}日経過` : diffDays === 0 ? "今日まで" : `期限まで${diffDays}日`;
}

function getSalesPaymentLabel(sale) {
  return sale.paymentStatus || getSalePaymentStatus(sale.invoiceAmount ?? 0, sale.paidAmount ?? 0);
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
    sale.reminderCount,
    sale.lastReminderDate,
    sale.reminderMethod,
    sale.reminderMemo,
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
  const value = Number(String(raw ?? "").replace(/,/g, "").trim());
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
  if (type === "client_interactions") return "対応履歴CSV";
  return "CSV";
}

function downloadCsvFile(filename, headers, rows) {
  const csv = buildCsvString(headers, rows);
  const bom = "\uFEFF";
  const blob = new Blob([bom + csv], { type: "text/csv;charset=utf-8;" });
  downloadBlobFile(filename, blob);
}

function downloadJsonFile(filename, value) {
  const pretty = JSON.stringify(value, null, 2);
  const blob = new Blob([pretty], { type: "application/json;charset=utf-8;" });
  downloadBlobFile(filename, blob);
}

function downloadBlobFile(filename, blob) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
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

function parseFlexibleInteger(value, fallback = 0) {
  const normalized = String(value ?? "").trim().replaceAll(",", "");
  if (!normalized) return fallback;
  const num = Number(normalized);
  if (!Number.isFinite(num)) return fallback;
  return Math.max(0, Math.floor(num));
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
  const clientInteractionsData = getSheetRows("対応履歴");
  if (clientInteractionsData.rows.length) {
    mergeResult(await importRowsByType("client_interactions", clientInteractionsData));
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
          template_id: asTrimmedText(row.template_id) || null,
          customer_name: customerName,
          case_name: caseName,
          estimate_amount: parseFlexibleAmount(row.estimate_amount),
          received_date: parseFlexibleDate(row.received_date),
          due_date: parseFlexibleDate(row.due_date),
          required_documents: asTrimmedText(row.required_documents) || null,
          task_list: asTrimmedText(row.task_list) || null,
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
    validateRequiredHeaders(headers, ["case_name", "invoice_amount", "paid_amount"]);
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
        const normalizedPaidAmount = paidAmount ?? 0;
        const paymentStatus = normalizeSalePaymentStatus(
          asTrimmedText(row.payment_status),
          invoiceAmount,
          normalizedPaidAmount,
        );
        payloads.push({
          user_id: currentUser.id,
          case_id: caseId,
          invoice_number: asTrimmedText(row.invoice_number) || null,
          invoice_amount: invoiceAmount,
          paid_amount: normalizedPaidAmount,
          paid_date: parseFlexibleDate(row.paid_date),
          due_date: parseFlexibleDate(row.due_date),
          payment_status: paymentStatus,
          is_unpaid: parseBooleanLike(row.is_unpaid) || paymentStatus !== "入金済",
          reminder_count: parseFlexibleInteger(row.reminder_count, 0),
          last_reminder_date: parseFlexibleDate(row.last_reminder_date),
          reminder_method: normalizeReminderMethod(row.reminder_method),
          reminder_memo: asTrimmedText(row.reminder_memo) || null,
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

  if (importType === "client_interactions") {
    validateRequiredHeaders(headers, ["interaction_date", "summary"]);
    const payloads = [];
    rows.forEach((row) => {
      try {
        const interactionDate = parseFlexibleDate(row.interaction_date);
        const summary = asTrimmedText(row.summary);
        if (!interactionDate || !summary) {
          result.errorCount += 1;
          return;
        }
        payloads.push({
          user_id: currentUser.id,
          client_id: asTrimmedText(row.client_id) || null,
          interaction_date: interactionDate,
          interaction_type: normalizeClientInteractionType(row.interaction_type),
          summary,
          next_action: asTrimmedText(row.next_action || row.interaction_next_action) || null,
          next_action_date: parseFlexibleDate(row.next_action_date || row.interaction_next_action_date),
          memo: asTrimmedText(row.memo || row.interaction_memo) || null,
        });
      } catch (_error) {
        result.errorCount += 1;
      }
    });
    if (payloads.length) {
      const { error } = await sbClient.from("client_interactions").insert(payloads);
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

function splitMultilineItems(text) {
  return String(text || "")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);
}

function convertTasksToChecklist(text) {
  const rows = splitMultilineItems(text);
  if (!rows.length) return "";
  return rows.map((line) => (/^\[( |x|X)\]/.test(line) ? line : `[ ] ${line}`)).join("\n");
}

function parseTaskList(taskList) {
  return splitMultilineItems(taskList).map((line) => {
    const done = /^\[(x|X)\]/.test(line) || /^✅/.test(line);
    const task = line.replace(/^\[( |x|X)\]\s*/, "").replace(/^✅\s*/, "").trim();
    return { done, task: task || "（名称未設定）" };
  });
}

function getIncompleteTaskCount(taskList) {
  return parseTaskList(taskList).filter((entry) => !entry.done).length;
}

function mapCaseFromDb(row) {
  return {
    id: row.id,
    userId: row.user_id,
    clientId: row.client_id || null,
    estimateId: row.estimate_id || null,
    templateId: row.template_id || null,
    customerName: row.customer_name || "",
    caseName: row.case_name || "",
    estimateAmount: normalizeAmount(row.estimate_amount),
    receivedDate: row.received_date || "",
    dueDate: row.due_date || "",
    status: normalizeStoredStatus(row.status),
    requiredDocuments: row.required_documents || "",
    taskList: row.task_list || "",
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

function mapWorkTemplateFromDb(row) {
  return {
    id: row.id,
    name: row.name || "",
    defaultDueDays: Number.isFinite(Number(row.default_due_days)) ? Number(row.default_due_days) : null,
    requiredDocuments: row.required_documents || "",
    defaultTasks: row.default_tasks || "",
    memo: row.memo || "",
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

function mapClientInteractionFromDb(row) {
  return {
    id: row.id,
    clientId: row.client_id || null,
    interactionDate: row.interaction_date || "",
    interactionType: normalizeClientInteractionType(row.interaction_type),
    summary: row.summary || "",
    nextAction: row.next_action || "",
    nextActionDate: row.next_action_date || "",
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

async function seedDefaultWorkTemplates() {
  if (!currentUser) return [];
  const defaults = [
    {
      user_id: currentUser.id,
      name: "建設業許可",
      default_due_days: 30,
      required_documents: "登記事項証明書\n納税証明書\n決算書\n経管資料\n専技資料",
      default_tasks: "要件確認\n必要書類案内\n書類回収\n申請書作成\n提出\n補正対応",
      memo: "建設業許可の標準テンプレート",
    },
    {
      user_id: currentUser.id,
      name: "車庫証明",
      default_due_days: 7,
      required_documents: "自認書または使用承諾書\n配置図\n所在図\n車検証情報",
      default_tasks: "書類確認\n現地確認\n警察署提出\n受取",
      memo: "車庫証明の標準テンプレート",
    },
    {
      user_id: currentUser.id,
      name: "古物商許可",
      default_due_days: 40,
      required_documents: "住民票\n身分証明書\n略歴書\n誓約書\nURL資料",
      default_tasks: "ヒアリング\n必要書類案内\n申請書作成\n警察署提出\n補正対応",
      memo: "古物商許可の標準テンプレート",
    },
  ];
  const { data, error } = await sbClient.from("work_templates").insert(defaults).select("*");
  if (error) throw error;
  return (data || []).map(mapWorkTemplateFromDb);
}

function mapEstimateFromDb(row) {
  const totalAmount = normalizeAmount(row.total) ?? 0;
  return {
    id: row.id,
    clientId: row.client_id || null,
    customerName: row.customer_name || "",
    estimateNumber: row.estimate_number || "",
    estimateTitle: row.estimate_title || "",
    estimateDate: row.estimate_date || "",
    sentDate: row.sent_date || row.estimate_date || "",
    validUntil: row.valid_until || "",
    status: normalizeEstimateStatus(row.status),
    memo: row.memo || "",
    subtotal: normalizeAmount(row.subtotal) ?? 0,
    tax: normalizeAmount(row.tax) ?? 0,
    totalAmount,
    total: totalAmount,
    caseId: row.case_id || null,
    createdAt: Date.parse(row.created_at) || Date.now(),
    updatedAt: Date.parse(row.updated_at) || Date.now(),
  };
}

function getPendingEstimates(estimates) {
  const today = new Date();
  return (Array.isArray(estimates) ? estimates : []).reduce((acc, estimate) => {
    if (!estimate || estimate.status !== "未回答") return acc;
    const baseDate = estimate.sentDate || estimate.estimateDate;
    if (!baseDate) return acc;
    const sent = new Date(baseDate);
    const sentTs = sent.getTime();
    if (!Number.isFinite(sentTs)) return acc;
    const diffDays = (today.getTime() - sentTs) / (1000 * 60 * 60 * 24);
    if (!Number.isFinite(diffDays) || diffDays < 3) return acc;
    acc.push({
      ...estimate,
      baseDate,
      elapsedDays: Math.floor(diffDays),
    });
    return acc;
  }, []);
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
  // DB拡張が未適用の場合は以下を実行してください:
  // alter table sales alter column case_id drop not null;
  // alter table sales add column if not exists paid_amount bigint default 0;
  // alter table sales add column if not exists payment_status text;
  // alter table sales add column if not exists due_date date;
  // alter table sales add column if not exists invoice_number text;
  // alter table sales add column if not exists estimate_id uuid;
  // alter table sales add column if not exists reminder_count int default 0;
  // alter table sales add column if not exists last_reminder_date date;
  // alter table sales add column if not exists reminder_method text;
  // alter table sales add column if not exists reminder_memo text;
  const invoiceAmount = Number(row.invoice_amount || 0);
  const paidAmount = Number(row.paid_amount || 0);
  return {
    id: row.id,
    userId: row.user_id,
    estimateId: row.estimate_id || null,
    caseId: row.case_id || null,
    invoiceAmount,
    paidAmount,
    paidDate: row.paid_date || null,
    dueDate: row.due_date || null,
    paymentStatus: row.payment_status || calculatePaymentStatus(invoiceAmount, paidAmount),
    isUnpaid: row.is_unpaid !== false,
    invoiceNumber: row.invoice_number || "",
    reminderCount: parseFlexibleInteger(row.reminder_count, 0),
    lastReminderDate: row.last_reminder_date || null,
    reminderMethod: normalizeReminderMethod(row.reminder_method),
    reminderMemo: row.reminder_memo || "",
    createdAt: Date.parse(row.created_at) || null,
    updatedAt: Date.parse(row.updated_at) || null,
  };
}

function mapExpenseFromDb(row) {
  return {
    id: row.id,
    date: row.date || row.expense_date || "",
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
    userId: row.user_id || null,
    reportDate: row.report_date || "",
    caseId: row.case_id || null,
    clientId: row.client_id || null,
    interactionType: normalizeDailyReportInteractionType(row.interaction_type || "作業"),
    workContent: row.work_content || "",
    workMinutes: normalizeAmount(row.work_minutes) ?? 0,
    nextAction: row.next_action || "",
    nextActionDate: row.next_action_date || "",
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

function setDataMutationControlsEnabled(enabled) {
  const controls = [
    clientSubmitBtn,
    caseSubmitBtn,
    workTemplateSubmitBtn,
    estimateSubmitBtn,
    saleSubmitBtn,
    expenseSubmitBtn,
    fixedExpenseSubmitBtn,
    dailyReportSubmitBtn,
    csvImportForm?.querySelector('button[type="submit"]'),
    excelImportForm?.querySelector('button[type="submit"]'),
    backupRestoreForm?.querySelector('button[type="submit"]'),
  ].filter(Boolean);
  controls.forEach((control) => {
    control.disabled = !enabled;
  });
}

function updateLoadingOverlay() {
  if (!loadingOverlay) return;
  const isLoading = loadingCount > 0;
  loadingOverlay.hidden = !isLoading;
  loadingOverlay.style.display = isLoading ? "grid" : "none";
}

function startLoading(taskName) {
  loadingCount += 1;
  console.log("LOADING START", taskName, loadingCount);
  updateLoadingOverlay();
  if (loadingTimeoutId) window.clearTimeout(loadingTimeoutId);
  loadingTimeoutId = window.setTimeout(() => {
    console.warn("LOADING TIMEOUT FORCE HIDE", taskName);
    forceHideLoading();
    showAppMessage("読み込みが長時間継続したため自動解除しました。", true);
  }, 10000);
}

function endLoading(taskName) {
  loadingCount = Math.max(0, loadingCount - 1);
  console.log("LOADING END", taskName, loadingCount);
  if (loadingCount === 0 && loadingTimeoutId) {
    window.clearTimeout(loadingTimeoutId);
    loadingTimeoutId = null;
  }
  updateLoadingOverlay();
}

function forceHideLoading() {
  loadingCount = 0;
  if (loadingTimeoutId) {
    window.clearTimeout(loadingTimeoutId);
    loadingTimeoutId = null;
  }
  updateLoadingOverlay();
}

async function withLoading(taskName, asyncFn, options = {}) {
  const { messageTarget = "app", triggerButton = null } = options;
  startLoading(taskName);
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
    endLoading(taskName);
  }
}

function setLoading(isLoading) {
  if (isLoading) {
    startLoading("setLoading");
  } else {
    forceHideLoading();
  }
}

function clearLoadingState() {
  forceHideLoading();
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

function resetViewState() {
  state.selectedAggregation = "month";
  state.selectedMonth = toMonthKey(new Date());
  state.selectedYear = new Date().getFullYear();
  state.caseStatusFilter = "all";
  state.caseSearchQuery = "";
  state.caseDeadlineFilter = "all";
  state.salesSearchQuery = "";
  state.expensesSearchQuery = "";
  state.dailyReportSearchQuery = "";
  state.dailyReportDateFilter = "all";
  state.estimateCustomerQuery = "";
  state.estimateTitleQuery = "";
  state.estimateStatusFilter = "all";
  state.estimateExpiredFilter = "all";
}

function resetEditState() {
  editState.clientId = null;
  editState.caseId = null;
  editState.workTemplateId = null;
  editState.saleId = null;
  editState.expenseId = null;
  editState.fixedExpenseId = null;
  editState.dailyReportId = null;
  editState.estimateId = null;
}

function clearAppState() {
  currentUser = null;
  state.isInitialDataReady = false;
  state.clients = [];
  state.workTemplates = [];
  state.cases = [];
  state.estimates = [];
  state.estimateItems = [];
  state.sales = [];
  state.expenses = [];
  state.fixedExpenses = [];
  state.dailyReports = [];
  resetViewState();
  resetEditState();
  resetClientForm();
  resetCaseForm();
  resetWorkTemplateForm();
  resetEstimateForm();
  resetSaleForm();
  resetExpenseForm();
  resetFixedExpenseForm();
  resetDailyReportForm();
  clearAppMessage();
  clearLoadingState();
  setDataMutationControlsEnabled(false);
  authView.hidden = false;
  appView.hidden = true;
  userLabel.textContent = "";
}
