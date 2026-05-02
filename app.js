const SUPABASE_URL = "https://ueelzyftlbnvjvpsmpyt.supabase.co";
const SUPABASE_KEY = "sb_publishable_0DrKsieUcCyEZN_HRg8LhQ_QqFTPMtp";
const STATUS_ORDER = ["未着手", "進行中", "完了"];
const STATUS_FILTER_KEYS = [...STATUS_ORDER, "その他"];
const DEADLINE_FILTER_KEYS = ["all", "overdue", "within3", "within7", "within30"];
const SALES_PAYMENT_STATUSES = ["未入金", "一部入金", "入金済"];
const ESTIMATE_STATUS_ORDER = ["作成中", "提出済", "未回答", "受注", "失注"];
const EXPENSE_PAYMENT_METHODS = ["現金", "クレジットカード", "銀行振込", "電子マネー", "口座振替", "その他"];
const DAILY_REPORT_INTERACTION_TYPES = ["作業", "電話", "メール", "面談", "訪問", "LINE", "その他"];
const REMINDER_METHODS = ["電話", "メール", "LINE", "郵送", "訪問", "その他"];
const PAYMENT_METHODS = ["現金", "振込", "その他"];
const CASE_DOCUMENT_STATUSES = ["未回収", "回収済", "確認済", "不備あり", "不要"];
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
const DEFAULT_APP_SETTINGS = {
  officeName: OFFICE_INFO.name,
  postalCode: OFFICE_INFO.zip,
  address: OFFICE_INFO.address,
  tel: OFFICE_INFO.tel,
  email: OFFICE_INFO.email,
  invoiceRegistrationNumber: OFFICE_INFO.registrationNumber,
  bankInfo: OFFICE_INFO.transferInfo,
  defaultInvoiceDueDays: 7,
  taxRate: 0.1,
  estimateNote: OFFICE_INFO.estimateNote,
  invoiceNote: OFFICE_INFO.invoiceNote,
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
  payments: [],
  expenses: [],
  fixedExpenses: [],
  dailyReports: [],
  caseTasks: [],
  caseDocuments: [],
  permitHearings: [],
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
  alerts: [],
  selectedIntegrityCheckKey: "",
  appSettings: { ...DEFAULT_APP_SETTINGS },
  changeLogs: [],
  isInitialDataReady: false,
  pendingAlertSelections: {},
};
const editState = { clientId: null, caseId: null, workTemplateId: null, saleId: null, expenseId: null, fixedExpenseId: null, dailyReportId: null, estimateId: null, caseTaskId: null, caseDocumentId: null };

const CLICK_ACTION_HANDLERS = {
  activate_tab: (event, button) => activateTab(button?.dataset?.tab),
  activate_subtab: (event, button) => activateSubtab(button?.dataset?.parentTab, button?.dataset?.subtab),
  signup: handleSignup,
  logout: handleLogout,
  manual_reload: handleManualReload,
  clear_all: handleClearAll,
  clear_status_filter: () => applyCaseStatusFilter("all"),
  clear_case_filters: clearCaseFilters,
  clear_sales_search: clearSalesSearch,
  clear_expenses_search: clearExpensesSearch,
  clear_daily_report_filters: clearDailyReportFilters,
  force_close_loading: () => {
    forceHideLoading();
    showAppMessage("読み込みを強制解除しました。必要なら再操作してください。", true);
  },
  export_cases_csv: handleExportCasesCsv,
  export_sales_csv: handleExportSalesCsv,
  export_payments_csv: handleExportPaymentsCsv,
  export_expenses_csv: handleExportExpensesCsv,
  export_fixed_expenses_csv: handleExportFixedExpensesCsv,
  export_case_documents_csv: handleExportCaseDocumentsCsv,
  export_all_csv: handleExportAllCsv,
  export_client_analysis_csv: handleExportClientAnalysisCsv,
  export_referral_analysis_csv: handleExportReferralAnalysisCsv,
  export_excel: handleExportExcel,
  export_analysis_excel: handleExportAnalysisExcel,
  export_backup_json: handleExportBackupJson,
  apply_permit_documents: applyPermitDocumentsToCase,
  apply_permit_tasks: applyPermitTasksToCase,
  add_estimate_item_row: () => addEstimateItemRow(),
  remove_estimate_item_row: handleEstimateItemsClick,
  status_summary_filter: handleStatusSummaryClick,
  deadline_alert_click: handleDeadlineAlertClick,
  alert_task_edit: handleTodayTaskAction,
  edit_client: handleClientListAction,
  delete_client: handleClientListAction,
  edit_case: handleCaseListAction,
  delete_case: handleCaseListAction,
  print_case_invoice: handleCaseListAction,
  export_case_invoice_excel: handleCaseListAction,
  edit_work_template: handleWorkTemplateListAction,
  delete_work_template: handleWorkTemplateListAction,
  edit_estimate: handleEstimateListAction,
  delete_estimate: handleEstimateListAction,
  create_case_from_estimate: handleEstimateListAction,
  create_invoice_from_estimate: handleEstimateListAction,
  print_estimate: handleEstimateListAction,
  print_invoice_from_estimate: handleEstimateListAction,
  print_delivery_note: handleEstimateListAction,
  print_purchase_order: handleEstimateListAction,
  print_order_confirmation: handleEstimateListAction,
  print_acceptance_certificate: handleEstimateListAction,
  print_case_delivery_note: (event, button) => {
    const caseId = button.dataset.caseId;
    return openCaseBusinessDocumentPrintPreview(caseId, "delivery_note");
  },
  print_case_purchase_order: (event, button) => {
    const caseId = button.dataset.caseId;
    return openCaseBusinessDocumentPrintPreview(caseId, "purchase_order");
  },
  print_case_order_confirmation: (event, button) => {
    const caseId = button.dataset.caseId;
    return openCaseBusinessDocumentPrintPreview(caseId, "order_confirmation");
  },
  print_case_acceptance_certificate: (event, button) => {
    const caseId = button.dataset.caseId;
    return openCaseBusinessDocumentPrintPreview(caseId, "acceptance_certificate");
  },
  export_estimate_excel: handleEstimateListAction,
  export_invoice_excel_from_estimate: handleEstimateListAction,
  edit_sale: handleSalesListAction,
  delete_sale: handleSalesListAction,
  record_payment: handleSalesListAction,
  delete_payment: handleSalesListAction,
  print_receipt: handleSalesListAction,
  print_cumulative_receipt: (event, button) => {
    const saleId = button.dataset.saleId;
    return openCumulativeReceiptPrintPreview(saleId);
  },
  record_reminder: handleSalesListAction,
  edit_expense: handleExpensesListAction,
  delete_expense: handleExpensesListAction,
  edit_fixed_expense: handleFixedExpensesListAction,
  toggle_fixed_expense: handleFixedExpensesListAction,
  delete_fixed_expense: handleFixedExpensesListAction,
  edit_daily_report: handleDailyReportsListAction,
  delete_daily_report: handleDailyReportsListAction,
  edit_case_task: handleCaseTaskListAction,
  complete_case_task: handleCaseTaskListAction,
  delete_case_task: handleCaseTaskListAction,
  register_sale: handleBillingLeakAlertAction,
  edit_pending_estimate: handlePendingEstimateAction,
  edit_case_document: handleCaseDocumentsListAction,
  delete_case_document: handleCaseDocumentsListAction,
  select_integrity_check: handleDataIntegrityAction,
  pending_alert_bulk_clear_next_action_date: () => handlePendingAlertBulkAction("clear_next_action_date"),
  pending_alert_bulk_complete_cases: () => handlePendingAlertBulkAction("complete_cases"),
  pending_alert_bulk_complete_tasks: () => handlePendingAlertBulkAction("complete_tasks"),
  pending_alert_bulk_complete_documents: () => handlePendingAlertBulkAction("complete_documents"),
  pending_alert_bulk_mark_invoiced_cases: () => handlePendingAlertBulkAction("mark_invoiced_cases"),
  pending_alert_open_sales: () => handlePendingAlertBulkAction("open_sales"),
  pending_alert_open_expenses: () => handlePendingAlertBulkAction("open_expenses"),
};

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
const exportPaymentsCsvBtn = document.getElementById("export-payments-csv-btn");
const exportExpensesCsvBtn = document.getElementById("export-expenses-csv-btn");
const exportFixedExpensesCsvBtn = document.getElementById("export-fixed-expenses-csv-btn");
const exportCaseDocumentsCsvBtn = document.getElementById("export-case-documents-csv-btn");
const exportAllCsvBtn = document.getElementById("export-all-csv-btn");
const csvImportForm = document.getElementById("csv-import-form");
const exportExcelBtn = document.getElementById("export-excel-btn");
const exportAnalysisExcelBtn = document.getElementById("export-analysis-excel-btn");
const exportClientAnalysisCsvBtn = document.getElementById("export-client-analysis-csv-btn");
const exportReferralAnalysisCsvBtn = document.getElementById("export-referral-analysis-csv-btn");
const excelImportForm = document.getElementById("excel-import-form");
const exportBackupJsonBtn = document.getElementById("export-backup-json-btn");
const backupRestoreForm = document.getElementById("backup-restore-form");
const manualReloadBtn = document.getElementById("manual-reload-btn");
const changeLogsEmpty = document.getElementById("change-logs-empty");
const changeLogsWrap = document.getElementById("change-logs-wrap");
const changeLogsBody = document.getElementById("change-logs-body");

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
  settings: document.getElementById("tab-settings"),
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
const pendingAlertCard = document.getElementById("pending-alert-card");
const pendingAlertSummary = document.getElementById("pending-alert-summary");
const pendingAlertCounts = document.getElementById("pending-alert-counts");
const pendingAlertBody = document.getElementById("pending-alert-body");
const pendingAlertEmpty = document.getElementById("pending-alert-empty");
const pendingAlertListWrap = document.getElementById("pending-alert-list-wrap");
const todayTaskCard = document.getElementById("today-task-card");
const todayTaskSummary = document.getElementById("today-task-summary");
const todayTaskBody = document.getElementById("today-task-body");
const todayTaskEmpty = document.getElementById("today-task-empty");
const todayTaskListWrap = document.getElementById("today-task-list-wrap");
const documentAlertCard = document.getElementById("document-alert-card");
const documentAlertSummary = document.getElementById("document-alert-summary");
const documentAlertBody = document.getElementById("document-alert-body");
const documentAlertEmpty = document.getElementById("document-alert-empty");
const documentAlertListWrap = document.getElementById("document-alert-list-wrap");
const pendingEstimatesCard = document.getElementById("pending-estimates-card");
const pendingEstimatesSummary = document.getElementById("pending-estimates-summary");
const pendingEstimatesBody = document.getElementById("pending-estimates-body");
const pendingEstimatesEmpty = document.getElementById("pending-estimates-empty");
const pendingEstimatesListWrap = document.getElementById("pending-estimates-list-wrap");
const dataIntegrityCard = document.getElementById("data-integrity-card");
const dataIntegritySummary = document.getElementById("data-integrity-summary");
const dataIntegrityBackupStatus = document.getElementById("data-integrity-backup-status");
const dataIntegrityList = document.getElementById("data-integrity-list");
const dataIntegrityDetailWrap = document.getElementById("data-integrity-detail-wrap");
const dataIntegrityDetailTitle = document.getElementById("data-integrity-detail-title");
const dataIntegrityDetailList = document.getElementById("data-integrity-detail-list");

const BACKUP_AT_STORAGE_KEY = "gyosei_last_backup_at";

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
const caseTaskForm = document.getElementById("case-task-form");
const caseTaskCaseSelect = document.getElementById("case-task-case-id");
const caseTaskSubmitBtn = document.getElementById("case-task-submit-btn");
const caseTasksBody = document.getElementById("case-tasks-body");
const caseTasksEmpty = document.getElementById("case-tasks-empty");
const caseTasksListWrap = document.getElementById("case-tasks-list-wrap");
const caseDocumentForm = document.getElementById("case-document-form");
const caseDocumentCaseSelect = document.getElementById("case-document-case-id");
const caseDocumentSubmitBtn = document.getElementById("case-document-submit-btn");
const caseDocumentsBody = document.getElementById("case-documents-body");
const caseDocumentsEmpty = document.getElementById("case-documents-empty");
const caseDocumentsListWrap = document.getElementById("case-documents-list-wrap");
const permitHearingForm = document.getElementById("permit-hearing-form");
const permitCaseSelect = document.getElementById("permit-case-id");
const permitResult = document.getElementById("permit-hearing-result");
const permitSummary = document.getElementById("permit-summary");
const permitDocumentsList = document.getElementById("permit-documents-list");
const permitTasksList = document.getElementById("permit-tasks-list");
const permitApplyDocumentsBtn = document.getElementById("permit-apply-documents-btn");
const permitApplyTasksBtn = document.getElementById("permit-apply-tasks-btn");
let lastPermitGenerated = null;

const saleForm = document.getElementById("sale-form");
const saleCaseSelect = document.getElementById("sale-case-id");
const saleInvoiceNumberInput = document.getElementById("sale-invoice-number");
const salesListBody = document.getElementById("sales-list-body");
const salesListWrap = document.getElementById("sales-list-wrap");
const salesEmpty = document.getElementById("sales-empty");
const saleSubmitBtn = document.getElementById("sale-submit-btn") || saleForm.querySelector('button[type="submit"]');
const saleInvoiceMemoInput = document.getElementById("sale-invoice-memo");
const salePaymentHistory = document.getElementById("sale-payment-history");
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
const dailyReportAlertBody = document.getElementById("daily-report-alert-body");
const dailyReportAlertEmpty = document.getElementById("daily-report-alert-empty");
const dailyReportAlertListWrap = document.getElementById("daily-report-alert-list-wrap");
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
const estimateTaxLabel = document.getElementById("estimate-tax-label");
const estimateTotal = document.getElementById("estimate-total");
const estimateList = document.getElementById("estimate-list");
const estimateEmpty = document.getElementById("estimate-empty");
const estimateCustomerSearch = document.getElementById("estimate-customer-search");
const estimateTitleSearch = document.getElementById("estimate-title-search");
const estimateStatusFilter = document.getElementById("estimate-status-filter");
const estimateExpiredFilter = document.getElementById("estimate-expired-filter");
const settingsForm = document.getElementById("settings-form");

let currentUser = null;
let isLoggingOut = false;
let eventsBound = false;
let loadingCount = 0;
let loadingTimer = null;
let isApplyingAuthState = false;
const BACKUP_TABLE_KEYS = ["clients", "work_templates", "cases", "case_tasks", "case_documents", "permit_hearings", "estimates", "estimate_items", "sales", "payments", "expenses", "fixed_expenses", "daily_reports", "app_settings"];
const RESTORE_INSERT_ORDER = ["clients", "work_templates", "cases", "case_tasks", "case_documents", "permit_hearings", "estimates", "estimate_items", "sales", "payments", "expenses", "fixed_expenses", "daily_reports", "app_settings"];
const RESTORE_DELETE_ORDER = ["payments", "estimate_items", "sales", "expenses", "fixed_expenses", "daily_reports", "permit_hearings", "case_documents", "case_tasks", "estimates", "cases", "work_templates", "clients", "app_settings"];
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
    await applyAuthState();

    sbClient.auth.onAuthStateChange((_event, sessionState) => {
      if (isLoggingOut) return;
      console.log("AUTH STATE CHANGED", _event);
      currentUser = sessionState?.user ?? null;
      setTimeout(() => {
        applyAuthState({ silent: true })
          .catch((error) => {
            console.error("認証状態変更処理に失敗しました", error);
          })
          .finally(() => {
            forceHideLoading();
          });
      }, 0);
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
  tabs.forEach((btn) => { btn.dataset.action = "activate_tab"; });
  subtabButtons.forEach((btn) => { btn.dataset.action = "activate_subtab"; });
  authForm.addEventListener("submit", handleLogin);
  signupBtn.dataset.action = "signup";
  logoutBtn.dataset.action = "logout";
  if (manualReloadBtn) manualReloadBtn.dataset.action = "manual_reload";

  clientForm?.addEventListener("submit", handleClientSubmit);
  caseForm.addEventListener("submit", handleCaseSubmit);
  caseTaskForm?.addEventListener("submit", handleCaseTaskSubmit);
  caseDocumentForm?.addEventListener("submit", handleCaseDocumentSubmit);
  permitHearingForm?.addEventListener("submit", handlePermitHearingSubmit);
  if (permitApplyDocumentsBtn) permitApplyDocumentsBtn.dataset.action = "apply_permit_documents";
  if (permitApplyTasksBtn) permitApplyTasksBtn.dataset.action = "apply_permit_tasks";
  workTemplateForm?.addEventListener("submit", handleWorkTemplateSubmit);
  console.log("SALE FORM FOUND", !!saleForm);
  console.log("EXPENSE FORM FOUND", !!expenseForm);
  console.log("DAILY REPORT FORM FOUND", !!dailyReportForm);
  saleForm?.addEventListener("submit", handleSaleSubmit);
  expenseForm?.addEventListener("submit", handleExpenseSubmit);
  fixedExpenseForm.addEventListener("submit", handleFixedExpenseSubmit);
  dailyReportForm?.addEventListener("submit", handleDailyReportSubmit);
  estimateForm?.addEventListener("submit", handleEstimateSubmit);
  settingsForm?.addEventListener("submit", handleSettingsSubmit);

  clearBtn.dataset.action = "clear_all";
  reportCaseSelect?.addEventListener("change", syncDailyReportClientFromCase);
  reportClientSelect?.addEventListener("change", syncDailyReportClientLabel);
  clientHistoryClientSelect?.addEventListener("change", renderClientHistory);
  targetMonthInput?.addEventListener("input", handleTargetMonthChange);
  targetYearInput?.addEventListener("input", handleTargetYearChange);
  aggregationRadios.forEach((radio) => radio.addEventListener("change", handleAggregationChange));
  if (statusFilterClearBtn) statusFilterClearBtn.dataset.action = "clear_status_filter";
  document.addEventListener("click", dispatchAction);
  caseSearchInput?.addEventListener("input", handleCaseSearchInput);
  caseStatusFilterSelect?.addEventListener("change", handleCaseStatusFilterChange);
  caseDeadlineFilterSelect?.addEventListener("change", handleCaseDeadlineFilterChange);
  if (caseFilterClearBtn) caseFilterClearBtn.dataset.action = "clear_case_filters";
  salesSearchInput?.addEventListener("input", handleSalesSearchInput);
  if (salesFilterClearBtn) salesFilterClearBtn.dataset.action = "clear_sales_search";
  expensesSearchInput?.addEventListener("input", handleExpensesSearchInput);
  if (expensesFilterClearBtn) expensesFilterClearBtn.dataset.action = "clear_expenses_search";
  dailyReportSearchInput?.addEventListener("input", handleDailyReportSearchInput);
  dailyReportDateFilterSelect?.addEventListener("change", handleDailyReportDateFilterChange);
  if (dailyReportFilterClearBtn) dailyReportFilterClearBtn.dataset.action = "clear_daily_report_filters";
  window.addEventListener("pageshow", forceHideLoading);
  window.addEventListener("focus", forceHideLoading);
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "visible") {
      forceHideLoading();
    }
  });
  if (loadingForceCloseBtn) loadingForceCloseBtn.dataset.action = "force_close_loading";
  if (exportCasesCsvBtn) exportCasesCsvBtn.dataset.action = "export_cases_csv";
  if (exportSalesCsvBtn) exportSalesCsvBtn.dataset.action = "export_sales_csv";
  if (exportPaymentsCsvBtn) exportPaymentsCsvBtn.dataset.action = "export_payments_csv";
  if (exportExpensesCsvBtn) exportExpensesCsvBtn.dataset.action = "export_expenses_csv";
  if (exportFixedExpensesCsvBtn) exportFixedExpensesCsvBtn.dataset.action = "export_fixed_expenses_csv";
  if (exportCaseDocumentsCsvBtn) exportCaseDocumentsCsvBtn.dataset.action = "export_case_documents_csv";
  if (exportAllCsvBtn) exportAllCsvBtn.dataset.action = "export_all_csv";
  if (exportClientAnalysisCsvBtn) exportClientAnalysisCsvBtn.dataset.action = "export_client_analysis_csv";
  if (exportReferralAnalysisCsvBtn) exportReferralAnalysisCsvBtn.dataset.action = "export_referral_analysis_csv";
  csvImportForm?.addEventListener("submit", handleCsvImportSubmit);
  if (exportExcelBtn) exportExcelBtn.dataset.action = "export_excel";
  if (exportAnalysisExcelBtn) exportAnalysisExcelBtn.dataset.action = "export_analysis_excel";
  excelImportForm?.addEventListener("submit", handleExcelImportSubmit);
  console.log("BACKUP BUTTON FOUND", !!exportBackupJsonBtn);
  if (exportBackupJsonBtn) exportBackupJsonBtn.dataset.action = "export_backup_json";
  backupRestoreForm?.addEventListener("submit", handleBackupRestoreSubmit);
  document.addEventListener("wheel", handleNumberInputWheel, { passive: true });
  if (estimateAddItemBtn) estimateAddItemBtn.dataset.action = "add_estimate_item_row";
  estimateItemsWrap?.addEventListener("input", handleEstimateItemsInput);
  
  caseClientSelect?.addEventListener("change", syncCaseCustomerFromClient);
  caseTemplateSelect?.addEventListener("change", handleCaseTemplateChange);
  estimateClientSelect?.addEventListener("change", syncEstimateCustomerFromClient);
  estimateCustomerSearch?.addEventListener("input", (event) => {
    state.estimateCustomerQuery = String(event.target.value || "").trim().toLowerCase();
    safeRender("estimates", renderEstimates);
  });
  estimateTitleSearch?.addEventListener("input", (event) => {
    state.estimateTitleQuery = String(event.target.value || "").trim().toLowerCase();
    safeRender("estimates", renderEstimates);
  });
  estimateStatusFilter?.addEventListener("change", (event) => {
    state.estimateStatusFilter = event.target.value || "all";
    safeRender("estimates", renderEstimates);
  });
  estimateExpiredFilter?.addEventListener("change", (event) => {
    state.estimateExpiredFilter = event.target.value || "all";
    safeRender("estimates", renderEstimates);
  });
  hydrateActionButtons();
  eventsBound = true;
  console.log("EVENTS BOUND");
  activateTab("cases");
}

function dispatchAction(event) {
  forceHideLoading();
  const button = event.target.closest("button");
  if (!(button instanceof HTMLButtonElement)) return;
  ensureButtonDataAction(button);
  const action = button.dataset.action;
  if (!action) return;
  event.preventDefault();
  const handler = CLICK_ACTION_HANDLERS[action];
  if (typeof handler !== "function") {
    console.error("未登録data-action", action, button);
    showAppMessage(`未登録の操作です: ${action}`, true);
    return;
  }
  Promise.resolve(handler(event, button)).catch((error) => {
    console.error("click action failed", action, error);
    showAppMessage(`操作に失敗しました: ${action}`, true);
  });
}

function hydrateActionButtons(root = document) {
  if (!root) return;
  root.querySelectorAll("button").forEach((button) => ensureButtonDataAction(button));
}

function ensureButtonDataAction(button) {
  if (!(button instanceof HTMLButtonElement)) return;
}

function handleAggregationChange(event) {
  const next = event?.target?.value;
  if (next === "all") state.selectedAggregation = "all";
  else state.selectedAggregation = next === "year" ? "year" : "month";
  safeRender("dashboard", renderDashboard);
}

function handleTargetMonthChange(event) {
  const nextValue = normalizeMonthKey(event?.target?.value);
  state.selectedMonth = nextValue || toMonthKey(new Date());
  safeRender("dashboard", renderDashboard);
}

function handleTargetYearChange(event) {
  const nextValue = normalizeYear(event?.target?.value);
  state.selectedYear = nextValue || new Date().getFullYear();
  safeRender("dashboard", renderDashboard);
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
  safeRender("dashboard", renderDashboard);
  safeRender("cases", renderCases);
}

function handleCaseSearchInput(event) {
  state.caseSearchQuery = String(event?.target?.value ?? "").trim().toLowerCase();
  safeRender("cases", renderCases);
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
  safeRender("cases", renderCases);
}

const PERMIT_SCENARIO_MASTER = {
  construction_corp: { label: "大阪府 建設業許可 知事許可 新規 法人", docs: ["定款", "履歴事項全部証明書", "役員全員の住民票", "営業所写真", "専任技術者の資格証"], tasks: ["要件確認", "必要書類案内", "役員・技術者情報の確認", "申請書作成", "提出・補正対応"] },
  construction_solo: { label: "大阪府 建設業許可 知事許可 新規 個人", docs: ["本人住民票", "身分証明書", "営業所写真", "専任技術者の資格証", "確定申告書控え"], tasks: ["本人要件確認", "必要書類案内", "技術者要件確認", "申請書作成", "提出・補正対応"] },
  takken_corp: { label: "大阪府 宅建業免許 新規 法人", docs: ["定款", "履歴事項全部証明書", "専任宅建士の登録証", "事務所使用権限書類", "役員の略歴書"], tasks: ["免許要件確認", "専任宅建士の確認", "必要書類回収", "申請書作成", "提出・審査対応"] },
  takken_solo: { label: "大阪府 宅建業免許 新規 個人", docs: ["本人住民票", "身分証明書", "専任宅建士の登録証", "事務所使用権限書類", "略歴書"], tasks: ["免許要件確認", "専任宅建士の確認", "必要書類回収", "申請書作成", "提出・審査対応"] },
  kobutsu_corp: { label: "大阪府警 古物商許可 新規 法人", docs: ["履歴事項全部証明書", "定款", "役員全員の住民票", "URL使用権限疎明資料（該当時）", "営業所の賃貸借契約書"], tasks: ["欠格要件確認", "営業所情報確認", "必要書類回収", "申請書作成", "警察署へ提出"] },
  kobutsu_solo: { label: "大阪府警 古物商許可 新規 個人", docs: ["本人住民票", "身分証明書", "URL使用権限疎明資料（該当時）", "営業所の賃貸借契約書", "略歴書"], tasks: ["欠格要件確認", "営業所情報確認", "必要書類回収", "申請書作成", "警察署へ提出"] },
};

async function handlePermitHearingSubmit(event) {
  event.preventDefault();
  const formData = new FormData(event.currentTarget);
  const caseId = String(formData.get("permitCaseId") || "");
  const linkedCase = state.cases.find((entry) => entry.id === caseId);
  if (!currentUser || !caseId || !linkedCase) {
    showAppMessage("案件選択は必須です。", true);
    return;
  }
  const scenarioKey = String(formData.get("permitScenario") || "construction_corp");
  const scenario = PERMIT_SCENARIO_MASTER[scenarioKey] || PERMIT_SCENARIO_MASTER.construction_corp;
  const applicantName = String(formData.get("permitApplicantName") || "").trim() || "（未入力）";
  const officeAddress = String(formData.get("permitOfficeAddress") || "").trim() || "（未入力）";
  const officerCount = Number(formData.get("permitOfficerCount") || 0);
  const qualifiedCount = Number(formData.get("permitQualifiedCount") || 0);
  const online = String(formData.get("permitOnlineApplication") || "false") === "true" ? "希望する" : "希望しない";
  const urgency = String(formData.get("permitUrgency") || "通常");
  permitSummary.textContent = `${scenario.label} / 申請者: ${applicantName} / 所在地: ${officeAddress} / 人員: ${officerCount}名 / 有資格者: ${qualifiedCount}名 / 電子申請: ${online} / 優先度: ${urgency}`;
  permitDocumentsList.innerHTML = scenario.docs.map((doc) => `<li>${escapeHtml(doc)}</li>`).join("");
  permitTasksList.innerHTML = scenario.tasks.map((task) => `<li>${escapeHtml(task)}</li>`).join("");
  permitResult.hidden = false;
  const linkedCustomerName = linkedCase.customer_name ?? linkedCase.customerName ?? "";
  const linkedCaseName = linkedCase.case_name ?? linkedCase.caseName ?? "";
  const linkedClientId = linkedCase.client_id ?? linkedCase.clientId ?? null;
  lastPermitGenerated = { case_id: caseId, customer_name: linkedCustomerName, case_name: linkedCaseName, docs: [...scenario.docs], tasks: [...scenario.tasks] };
  const { error } = await sbClient.from("permit_hearings").insert({
    user_id: currentUser.id,
    case_id: caseId,
    client_id: linkedClientId,
    permit_category: scenario.label,
    application_type: "新規",
    jurisdiction_prefecture: "大阪府",
    applicant_type: scenarioKey.endsWith("_corp") ? "法人" : "個人",
    applicant_name: applicantName === "（未入力）" ? null : applicantName,
    office_address: officeAddress === "（未入力）" ? null : officeAddress,
    officer_count: officerCount,
    qualified_person_count: qualifiedCount,
    online_application: online === "希望する",
    urgency,
    answers: { permitScenario: scenarioKey, permitApplicantName: applicantName, permitOfficeAddress: officeAddress },
    generated_documents: scenario.docs,
    generated_tasks: scenario.tasks,
    source_urls: [],
  });
  if (error) return showAppMessage(`許認可ヒアリング保存エラー: ${formatSupabaseError(error)}`, true);
  await loadPermitHearings();
}

async function applyPermitDocumentsToCase() {
  if (!currentUser || !lastPermitGenerated?.case_id) return;
  const existing = new Set(state.caseDocuments
    .filter((d) => (d.case_id ?? d.caseId) === lastPermitGenerated.case_id)
    .map((d) => String((d.document_name ?? d.documentName) || "").trim()));
  const payload = lastPermitGenerated.docs.filter((name) => !existing.has(name)).map((name) => ({
    user_id: currentUser.id,
    case_id: lastPermitGenerated.case_id,
    document_name: name,
    status: "未回収",
    received_date: null,
    checked_date: null,
    memo: null,
    file_url: null,
  }));
  if (!payload.length) return showAppMessage("同一案件に同名書類があるため追加対象はありません。", false);
  const { error } = await sbClient.from("case_documents").insert(payload);
  if (error) return showAppMessage(`書類反映エラー: ${formatSupabaseError(error)}`, true);
  await loadCaseDocuments();
  safeRender("cases", renderCaseDocuments);
}

async function applyPermitTasksToCase() {
  if (!currentUser || !lastPermitGenerated?.case_id) return;
  const existing = new Set(state.caseTasks
    .filter((t) => (t.case_id ?? t.caseId) === lastPermitGenerated.case_id)
    .map((t) => String((t.task_title ?? t.taskTitle) || "").trim()));
  // case_tasks の実カラム名は task_title（既存の登録・表示・CSV・取込処理と同一）
  const payload = lastPermitGenerated.tasks.filter((title) => !existing.has(title)).map((title) => ({
    user_id: currentUser.id,
    case_id: lastPermitGenerated.case_id,
    task_title: title,
    task_memo: null,
    due_date: null,
    status: "未着手",
    completed_at: null,
  }));
  if (!payload.length) return showAppMessage("同一案件に同名タスクがあるため追加対象はありません。", false);
  const { error } = await sbClient.from("case_tasks").insert(payload);
  if (error) return showAppMessage(`タスク反映エラー: ${formatSupabaseError(error)}`, true);
  await loadCaseTasks();
  safeRender("cases", renderCaseTasks);
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
  const taskId = button.dataset.taskId;
  if (taskId) startCaseTaskEdit(taskId);
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
  safeRender("cases", renderCases);
}

function handleSalesSearchInput(event) {
  state.salesSearchQuery = String(event?.target?.value ?? "").trim().toLowerCase();
  safeRender("sales", renderSales);
}

function clearSalesSearch() {
  state.salesSearchQuery = "";
  if (salesSearchInput) salesSearchInput.value = "";
  safeRender("sales", renderSales);
}

function handleExpensesSearchInput(event) {
  state.expensesSearchQuery = String(event?.target?.value ?? "").trim().toLowerCase();
  safeRender("expenses", renderExpenses);
}

function clearExpensesSearch() {
  state.expensesSearchQuery = "";
  if (expensesSearchInput) expensesSearchInput.value = "";
  safeRender("expenses", renderExpenses);
}

function handleDailyReportSearchInput(event) {
  state.dailyReportSearchQuery = String(event?.target?.value ?? "").trim().toLowerCase();
  safeRender("dailyReports", renderDailyReports);
}

function handleDailyReportDateFilterChange(event) {
  applyDailyReportDateFilter(event?.target?.value || "all");
}

function applyDailyReportDateFilter(nextFilter) {
  const normalized = ["all", "today", "month"].includes(nextFilter) ? nextFilter : "all";
  state.dailyReportDateFilter = normalized;
  if (dailyReportDateFilterSelect) dailyReportDateFilterSelect.value = normalized;
  safeRender("dailyReports", renderDailyReports);
}

function clearDailyReportFilters() {
  state.dailyReportSearchQuery = "";
  applyDailyReportDateFilter("all");
  if (dailyReportSearchInput) dailyReportSearchInput.value = "";
}


async function handleVisibilityChange() {
  if (document.hidden) return;
  forceHideLoading();
}

async function handleWindowFocus() {
  forceHideLoading();
}

async function handlePageShow() {
  forceHideLoading();
}

async function applyAuthState(options = {}) {
  if (isApplyingAuthState) {
    console.warn("applyAuthState skipped: already running");
    return;
  }

  const silent = options.silent === true;
  isApplyingAuthState = true;

  try {
    if (!currentUser) {
      clearAppState();
      return;
    }

    if (!silent) setDataMutationControlsEnabled(false);
    state.isInitialDataReady = false;
    userLabel.textContent = currentUser.email || "ログイン中";

    await loadAllDataSafely();
    resetCaseForm();
    resetCaseTaskForm();
    resetCaseDocumentForm();
    resetEstimateForm();
    resetSaleForm();
    resetExpenseForm();
    resetFixedExpenseForm();
    resetDailyReportForm();
    safeRender("renderAfterDataChanged", renderAfterDataChanged);
    state.isInitialDataReady = true;
    setDataMutationControlsEnabled(true);
  } catch (error) {
    console.error("applyAuthState failed", error);
    throw error;
  } finally {
    isApplyingAuthState = false;
    authView.hidden = true;
    appView.hidden = false;
    forceHideLoading();
  }
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

async function handleManualReload(event) {
  if (event) event.preventDefault();
  if (!currentUser) return;
  if (manualReloadBtn) manualReloadBtn.disabled = true;
  try {
    await loadAllDataSafely();
    renderAfterDataChanged();
    showAppMessage("最新データを読み込みました", false);
  } catch (error) {
    showAppMessage(`最新データ再読込に失敗しました。${formatSupabaseError(error)}`, true);
  } finally {
    if (manualReloadBtn) manualReloadBtn.disabled = false;
  }
}

async function loadClients() {
  if (!currentUser || isLoggingOut) {
    state.clients = [];
    return;
  }
  try {
    const { data, error } = await sbClient.from("clients").select("*").eq("user_id", currentUser.id);
    if (error) {
      console.error("LOAD clients ERROR", error);
      state.clients = [];
      return;
    }
    state.clients = (data || []).map(mapClientFromDb);
  } catch (error) {
    console.error("LOAD clients ERROR", error);
    state.clients = [];
  }
}

async function loadWorkTemplates() {
  if (!currentUser || isLoggingOut) {
    state.workTemplates = [];
    return;
  }
  try {
    const { data, error } = await sbClient.from("work_templates").select("*").eq("user_id", currentUser.id);
    if (error) {
      console.error("LOAD workTemplates ERROR", error);
      state.workTemplates = [];
      return;
    }
    state.workTemplates = (data || []).map(mapWorkTemplateFromDb);
    if (!state.workTemplates.length) {
      try {
        const seeded = await seedDefaultWorkTemplates();
        if (seeded.length) state.workTemplates = seeded;
      } catch (seedError) {
        console.error("LOAD workTemplates seed ERROR", seedError);
      }
    }
  } catch (error) {
    console.error("LOAD workTemplates ERROR", error);
    state.workTemplates = [];
  }
}

async function loadCases() {
  if (!currentUser || isLoggingOut) {
    state.cases = [];
    return;
  }
  try {
    const { data, error } = await sbClient.from("cases").select("*").eq("user_id", currentUser.id);
    if (error) {
      console.error("LOAD cases ERROR", error);
      state.cases = [];
      return;
    }
    state.cases = (data || []).map(mapCaseFromDb);
  } catch (error) {
    console.error("LOAD cases ERROR", error);
    state.cases = [];
  }
}

async function loadCaseTasks() {
  if (!currentUser || isLoggingOut) {
    state.caseTasks = [];
    return;
  }
  try {
    const { data, error } = await sbClient.from("case_tasks").select("*").eq("user_id", currentUser.id);
    if (error) {
      console.error("LOAD caseTasks ERROR", error);
      state.caseTasks = [];
      return;
    }
    state.caseTasks = (data || []).map(mapCaseTaskFromDb);
  } catch (error) {
    console.error("LOAD caseTasks ERROR", error);
    state.caseTasks = [];
  }
}

async function loadCaseDocuments() {
  if (!currentUser || isLoggingOut) {
    state.caseDocuments = [];
    return;
  }
  try {
    const { data, error } = await sbClient.from("case_documents").select("*").eq("user_id", currentUser.id);
    if (error) {
      console.error("LOAD caseDocuments ERROR", error);
      state.caseDocuments = [];
      return;
    }
    state.caseDocuments = (data || []).map(mapCaseDocumentFromDb);
  } catch (error) {
    console.error("LOAD caseDocuments ERROR", error);
    state.caseDocuments = [];
  }
}

async function loadPermitHearings() {
  if (!currentUser || isLoggingOut) {
    state.permitHearings = [];
    return;
  }
  try {
    const { data, error } = await sbClient.from("permit_hearings").select("*").eq("user_id", currentUser.id);
    if (error) {
      console.error("LOAD permitHearings ERROR", error);
      state.permitHearings = [];
      return;
    }
    state.permitHearings = data || [];
  } catch (error) {
    console.error("LOAD permitHearings ERROR", error);
    state.permitHearings = [];
  }
}

async function loadEstimates() {
  if (!currentUser || isLoggingOut) {
    state.estimates = [];
    return;
  }
  try {
    const { data, error } = await sbClient.from("estimates").select("*").eq("user_id", currentUser.id);
    if (error) {
      console.error("LOAD estimates ERROR", error);
      state.estimates = [];
      return;
    }
    state.estimates = (data || []).map(mapEstimateFromDb);
  } catch (error) {
    console.error("LOAD estimates ERROR", error);
    state.estimates = [];
  }
}

async function loadEstimateItems() {
  if (!currentUser || isLoggingOut) {
    state.estimateItems = [];
    return;
  }
  try {
    const { data, error } = await sbClient.from("estimate_items").select("*").eq("user_id", currentUser.id);
    if (error) {
      console.error("LOAD estimateItems ERROR", error);
      state.estimateItems = [];
      return;
    }
    state.estimateItems = (data || []).map(mapEstimateItemFromDb);
  } catch (error) {
    console.error("LOAD estimateItems ERROR", error);
    state.estimateItems = [];
  }
}

async function loadSales() {
  if (!currentUser || isLoggingOut) {
    state.sales = [];
    return;
  }
  try {
    const { data, error } = await sbClient.from("sales").select("*").eq("user_id", currentUser.id);
    if (error) {
      console.error("LOAD sales ERROR", error);
      state.sales = [];
      return;
    }
    state.sales = (data || []).map(mapSaleFromDb);
  } catch (error) {
    console.error("LOAD sales ERROR", error);
    state.sales = [];
  }
}

async function loadPayments() {
  if (!currentUser || isLoggingOut) {
    state.payments = [];
    return;
  }
  try {
    const { data, error } = await sbClient.from("payments").select("*").eq("user_id", currentUser.id);
    if (error) {
      console.error("LOAD payments ERROR", error);
      state.payments = [];
      return;
    }
    state.payments = (data || []).map(mapPaymentFromDb);
  } catch (error) {
    console.error("LOAD payments ERROR", error);
    state.payments = [];
  }
}

async function loadExpenses() {
  if (!currentUser || isLoggingOut) {
    state.expenses = [];
    return;
  }
  try {
    const { data, error } = await sbClient.from("expenses").select("*").eq("user_id", currentUser.id);
    if (error) {
      console.error("LOAD expenses ERROR", error);
      state.expenses = [];
      return;
    }
    state.expenses = (data || []).map(mapExpenseFromDb);
  } catch (error) {
    console.error("LOAD expenses ERROR", error);
    state.expenses = [];
  }
}

async function loadFixedExpenses() {
  if (!currentUser || isLoggingOut) {
    state.fixedExpenses = [];
    return;
  }
  try {
    const { data, error } = await sbClient.from("fixed_expenses").select("*").eq("user_id", currentUser.id);
    if (error) {
      console.error("LOAD fixedExpenses ERROR", error);
      state.fixedExpenses = [];
      return;
    }
    state.fixedExpenses = (data || []).map(mapFixedExpenseFromDb);
  } catch (error) {
    console.error("LOAD fixedExpenses ERROR", error);
    state.fixedExpenses = [];
  }
}

async function loadDailyReports() {
  if (!currentUser || isLoggingOut) {
    state.dailyReports = [];
    return;
  }
  try {
    const { data, error } = await sbClient.from("daily_reports").select("*").eq("user_id", currentUser.id);
    if (error) {
      console.error("LOAD dailyReports ERROR", error);
      state.dailyReports = [];
      return;
    }
    state.dailyReports = (data || []).map(mapDailyReportFromDb);
  } catch (error) {
    console.error("LOAD dailyReports ERROR", error);
    state.dailyReports = [];
  }
}

async function loadAppSettings() {
  if (!currentUser || isLoggingOut) {
    state.appSettings = { ...DEFAULT_APP_SETTINGS };
    return;
  }
  try {
    const { data, error } = await sbClient.from("app_settings").select("*").eq("user_id", currentUser.id);
    if (error) {
      console.error("LOAD appSettings ERROR", error);
      state.appSettings = { ...DEFAULT_APP_SETTINGS };
      return;
    }
    const loadedSettings = (data || []).map(mapAppSettingsFromDb);
    state.appSettings = loadedSettings[0] ? { ...DEFAULT_APP_SETTINGS, ...loadedSettings[0] } : { ...DEFAULT_APP_SETTINGS };
  } catch (error) {
    console.error("LOAD appSettings ERROR", error);
    state.appSettings = { ...DEFAULT_APP_SETTINGS };
  }
}

async function loadChangeLogs() {
  if (!currentUser || isLoggingOut) {
    state.changeLogs = [];
    return;
  }
  try {
    const { data, error } = await sbClient.from("operation_logs").select("*").eq("user_id", currentUser.id).order("created_at", { ascending: false }).limit(200);
    if (error) {
      console.error("LOAD changeLogs ERROR", error);
      state.changeLogs = [];
      return;
    }
    state.changeLogs = Array.isArray(data) ? data : [];
  } catch (error) {
    console.error("LOAD changeLogs ERROR", error);
    state.changeLogs = [];
  }
}

async function appendChangeLog({ action, table, recordId, before, after }) {
  if (!currentUser) return;
  const payload = {
    user_id: currentUser.id,
    action_type: action,
    target_type: table,
    target_id: recordId || null,
    detail: JSON.stringify({ before: before || null, after: after || null }),
  };
  const { error } = await sbClient.from("operation_logs").insert(payload);
  if (error) throw error;
}

async function loadAllDataSafely() {
  if (!currentUser || isLoggingOut) return;

  const loaders = [
    ["clients", loadClients],
    ["workTemplates", loadWorkTemplates],
    ["cases", loadCases],
    ["caseTasks", loadCaseTasks],
    ["caseDocuments", loadCaseDocuments],
    ["permitHearings", loadPermitHearings],
    ["estimates", loadEstimates],
    ["estimateItems", loadEstimateItems],
    ["sales", loadSales],
    ["payments", loadPayments],
    ["expenses", loadExpenses],
    ["fixedExpenses", loadFixedExpenses],
    ["dailyReports", loadDailyReports],
    ["appSettings", loadAppSettings],
    ["changeLogs", loadChangeLogs],
  ];

  for (const [name, loader] of loaders) {
    try {
      if (typeof loader === "function") {
        await loader();
      }
    } catch (error) {
      console.error(`LOAD ${name} ERROR`, error);
      if (Array.isArray(state[name])) {
        state[name] = [];
      }
    }
  }

  console.log("LOAD COUNTS", {
    clients: state.clients.length,
    cases: state.cases.length,
    sales: state.sales.length,
    expenses: state.expenses.length,
    dailyReports: state.dailyReports.length,
    fixedExpenses: state.fixedExpenses.length,
    estimates: state.estimates.length,
    caseTasks: state.caseTasks.length,
    caseDocuments: state.caseDocuments.length,
    payments: state.payments.length,
    appSettings: state.appSettings?.id ? 1 : 0,
  });

  try {
    await cleanupLegacyEstimateMemoMarkers();
  } catch (error) {
    console.error("LOAD cleanupLegacyEstimateMemoMarkers ERROR", error);
  }

  try {
    const createdCount = await ensureMonthlyFixedExpenses();
    if (createdCount > 0) {
      await loadExpenses();
      showAppMessage(`${createdCount}件の固定費を当月分として自動計上しました。`, false);
    }
  } catch (error) {
    console.error("LOAD ensureMonthlyFixedExpenses ERROR", error);
  }

  state.isInitialDataReady = true;
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
  console.log("EDIT STATE", editState);
  console.log("ACTION START", taskName, editState.clientId || "new");
  console.log("PAYLOAD", payload);
  try {
    await runMutation(taskName, async () => {
      if (editState.clientId) {
        const { data, error } = await sbClient.from("clients").update(payload).eq("id", editState.clientId).eq("user_id", currentUser.id).select().single();
        if (error) throw error;
        if (!data) throw new Error("更新結果を取得できませんでした。");
        console.log("DB DONE", taskName, editState.clientId);
        return data;
      }
      const { data, error } = await sbClient.from("clients").insert(payload).select().single();
      if (error) throw error;
      if (!data) throw new Error("登録結果を取得できませんでした。");
      console.log("DB DONE", taskName, data.id || "new");
      return data;
    }, {
      successMessage: editState.clientId ? "顧客を更新しました。" : "顧客を登録しました。",
      resetForm: resetClientForm,
    });
  } catch (error) {
    showAppMessage(`顧客保存に失敗しました。${formatSupabaseError(error)}`, true);
  }
}

async function handleCaseSubmit(event) {
  event.preventDefault();
  console.log("CASE SUBMIT FIRED");
  if (!currentUser || !ensureInitialDataReady("案件登録")) return;

  const taskName = editState.caseId ? "案件更新" : "案件登録";
  const isEdit = Boolean(editState.caseId);
  try {
    await runMutation(taskName, async () => {
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
        await createCaseTasksFromTemplate(data, payload.template_id);
        await createCaseDocumentsFromTemplate(data, payload.template_id);
        console.log("CASE INSERT SUCCESS", data);
      }
      return true;
    }, {
      successMessage: isEdit ? "案件を更新しました。" : "案件を登録しました。",
      resetForm: resetCaseForm,
      afterSuccess: () => {
        subtabState.cases = "list";
        activateTab("cases");
      },
    });
  } catch (error) {
    console.error("案件登録に失敗しました", error);
    showAppMessage(`案件保存に失敗しました。${formatSupabaseError(error)}`, true);
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
    standard_estimate_amount: normalizeAmount(workTemplateForm.elements.templateStandardEstimateAmount.value),
    default_due_days: parseNumberInput(workTemplateForm.elements.templateDefaultDueDays.value),
    required_documents: asTrimmedText(workTemplateForm.elements.templateRequiredDocuments.value) || null,
    default_tasks: asTrimmedText(workTemplateForm.elements.templateDefaultTasks.value) || null,
    memo: asTrimmedText(workTemplateForm.elements.templateMemo.value) || null,
  };
  const taskName = editState.workTemplateId ? "業務テンプレート更新" : "業務テンプレート登録";
  console.log("EDIT STATE", editState);
  try {
    await runMutation(taskName, async () => {
      if (editState.workTemplateId) {
        const { data, error } = await sbClient.from("work_templates").update(payload).eq("id", editState.workTemplateId).eq("user_id", currentUser.id).select().single();
        if (error) throw error;
        if (!data) throw new Error("更新結果を取得できませんでした。");
        return data;
      }
      const { data, error } = await sbClient.from("work_templates").insert(payload).select().single();
      if (error) throw error;
      if (!data) throw new Error("登録結果を取得できませんでした。");
      return data;
    }, {
      successMessage: editState.workTemplateId ? "業務テンプレートを更新しました。" : "業務テンプレートを登録しました。",
      resetForm: resetWorkTemplateForm,
    });
  } catch (error) {
    showAppMessage(`業務テンプレート保存に失敗しました。${formatSupabaseError(error)}`, true);
  }
}

async function createCaseTasksFromTemplate(caseRow, templateId) {
  if (!currentUser || !caseRow?.id || !templateId) return;
  const template = state.workTemplates.find((entry) => entry.id === templateId);
  if (!template) return;
  const taskSource = asTrimmedText(template.defaultTasks) || asTrimmedText(template.taskList);
  const taskLines = splitMultilineItems(taskSource);
  if (!taskLines.length) return;
  const dueDate = caseRow.due_date || null;
  const { data: existingRows, error: existingError } = await sbClient
    .from("case_tasks")
    .select("task_title")
    .eq("user_id", currentUser.id)
    .eq("case_id", caseRow.id);
  if (existingError) throw existingError;
  const existingTaskNames = new Set((existingRows || []).map((row) => normalizeAutoCreateItemName(row.task_title)));
  const insertingRows = [];
  for (const line of taskLines) {
    const title = asTrimmedText(line);
    const normalizedTitle = normalizeAutoCreateItemName(title);
    if (!normalizedTitle || existingTaskNames.has(normalizedTitle)) continue;
    existingTaskNames.add(normalizedTitle);
    insertingRows.push({
      user_id: currentUser.id,
      case_id: caseRow.id,
      task_title: title,
      task_memo: null,
      due_date: dueDate,
      status: "未完了",
      completed_at: null,
    });
  }
  if (!insertingRows.length) return;
  const { error } = await sbClient.from("case_tasks").insert(insertingRows);
  if (error) throw error;
}

async function createCaseDocumentsFromTemplate(caseRow, templateId) {
  if (!currentUser || !caseRow?.id || !templateId) return;
  const template = state.workTemplates.find((entry) => entry.id === templateId);
  if (!template) return;
  const docLines = splitMultilineItems(template.requiredDocuments);
  if (!docLines.length) return;
  const { data: existingRows, error: existingError } = await sbClient
    .from("case_documents")
    .select("document_name")
    .eq("user_id", currentUser.id)
    .eq("case_id", caseRow.id);
  if (existingError) throw existingError;
  const existingDocumentNames = new Set((existingRows || []).map((row) => normalizeAutoCreateItemName(row.document_name)));
  const insertingRows = [];
  for (const line of docLines) {
    const documentName = asTrimmedText(line);
    const normalizedDocumentName = normalizeAutoCreateItemName(documentName);
    if (!normalizedDocumentName || existingDocumentNames.has(normalizedDocumentName)) continue;
    existingDocumentNames.add(normalizedDocumentName);
    insertingRows.push({
      user_id: currentUser.id,
      case_id: caseRow.id,
      document_name: documentName,
      status: "未回収",
      received_date: null,
      checked_date: null,
      memo: null,
      file_url: null,
    });
  }
  if (!insertingRows.length) return;
  const { error } = await sbClient.from("case_documents").insert(insertingRows);
  if (error) throw error;
}

async function handleSaleSubmit(event) {
  event.preventDefault();
  if (!currentUser) {
    showAppMessage("ログイン状態を確認できません。", true);
    return;
  }
  if (!saleForm) return;

  try {
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
    const rawPayload = {
      user_id: currentUser.id,
      case_id: caseId || null,
      invoice_amount: invoiceAmount,
      paid_amount: paidAmount || 0,
      paid_date: paidDate || null,
      due_date: dueDate || null,
      payment_status: paymentStatus,
      is_unpaid: paymentStatus !== "入金済",
    };
    const payload = pickObjectKeys(rawPayload, SALES_MUTATION_COLUMNS);

    console.log("SALE PAYLOAD", payload);

    await runMutation(isEdit ? "売上更新" : "売上登録", async () => {
      if (isEdit) {
        const { data, error } = await sbClient
          .from("sales")
          .update(payload)
          .eq("id", editState.saleId)
          .eq("user_id", currentUser.id)
          .select()
          .single();
        if (error) throw error;
        if (!data) throw new Error("売上更新結果を取得できませんでした。");
        const beforeSale = state.sales.find((entry) => entry.id === editState.saleId) || null;
        await appendChangeLog({ action: "update", table: "sales", recordId: data.id, before: beforeSale, after: mapSaleFromDb(data) });
        console.log("SALE UPDATE SUCCESS", data);
        return data;
      }

      const invoiceNumber = await generateInvoiceNumberIfNeeded();
      const insertPayload = pickObjectKeys({
        ...payload,
        invoice_number: invoiceNumber || null,
      }, SALES_MUTATION_COLUMNS);
      const { data, error } = await sbClient.from("sales").insert(insertPayload).select().single();
      if (error) throw error;
      if (!data) throw new Error("売上登録結果を取得できませんでした。");
      await appendChangeLog({ action: "create", table: "sales", recordId: data.id, before: null, after: mapSaleFromDb(data) });
      console.log("SALE INSERT SUCCESS", data);
      return data;
    }, {
      successMessage: isEdit ? "売上を更新しました。" : "売上を登録しました。",
      resetForm: () => {
        resetSaleForm();
        editState.saleId = null;
      },
      afterSuccess: () => {
        activateSubtab("sales", "list");
      },
    });
  } catch (error) {
    console.error("売上登録に失敗しました。", error);
    showAppMessage(`売上登録に失敗しました。${error?.message || ""} ${error?.details || ""} ${error?.hint || ""} ${error?.code || ""}`, true);
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
  console.log("EDIT STATE", editState);
  console.log("ACTION START", taskName, editState.expenseId || "new");
  console.log("PAYLOAD", payload);
  try {
    await runMutation(taskName, async () => {
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
        await appendChangeLog({ action: "update", table: "expenses", recordId: data.id, before: currentExpense || null, after: mapExpenseFromDb(data) });
        console.log("DB DONE", taskName, editState.expenseId);
        return data;
      }

      const { data, error } = await sbClient.from("expenses").insert(payload).select().single();
      if (error) {
        showAppMessage(`経費登録に失敗しました。\n詳細: ${error.message || error}`, true);
        throw error;
      }
      if (!data) throw new Error("登録結果を取得できませんでした。");
      await appendChangeLog({ action: "create", table: "expenses", recordId: data.id, before: null, after: mapExpenseFromDb(data) });
      console.log("DB DONE", taskName, data.id || "new");
      return data;
    }, {
      successMessage: editState.expenseId ? "経費を更新しました。" : "経費を登録しました。",
      resetForm: resetExpenseForm,
      afterSuccess: () => {
        subtabState.expenses = "list";
        activateSubtab("expenses", "list");
      },
    });
  } catch (error) {
    showAppMessage(`経費保存に失敗しました。${formatSupabaseError(error)}`, true);
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
  console.log("EDIT STATE", editState);
  console.log("ACTION START", taskName, editState.fixedExpenseId || "new");
  console.log("PAYLOAD", payload);
  try {
    await runMutation(taskName, async () => {
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
        console.log("DB DONE", taskName, editState.fixedExpenseId);
        return data;
      }
      const { data, error } = await sbClient.from("fixed_expenses").insert(payload).select().single();
      if (error) throw error;
      if (!data) throw new Error("登録結果を取得できませんでした。");
      console.log("DB DONE", taskName, data.id || "new");
      return data;
    }, {
      successMessage: editState.fixedExpenseId ? "固定費を更新しました。" : "固定費を登録しました。",
      resetForm: resetFixedExpenseForm,
    });
  } catch (error) {
    showAppMessage(`固定費保存に失敗しました。${formatSupabaseError(error)}`, true);
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
    await runMutation(taskName, async () => {
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
      return true;
    }, {
      successMessage: isEdit ? "日報を更新しました。" : "日報を登録しました。",
      resetForm: resetDailyReportForm,
      afterSuccess: () => activateSubtab("daily-reports", "list"),
    });
  } catch (error) {
    showAppMessage(`日報登録に失敗しました。${error?.message || ""} ${error?.details || ""} ${error?.hint || ""} ${error?.code || ""}`, true);
  }
}

async function handleClientListAction(event) {
  const btn = event.target.closest("button");
  if (!(btn instanceof HTMLButtonElement) || !currentUser) return;
  const listAction = btn.dataset.listAction;
  const item = btn.closest("[data-id]");
  const id = item?.dataset.id;
  if (!id) {
    showAppMessage("対象データIDを取得できませんでした。", true);
    return;
  }
  if (listAction === "edit_client") {
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
  if (listAction === "delete_client") {
    await deleteClient(id);
  }
}

async function handleCaseListAction(event) {
  const btn = event.target.closest("button");
  if (!(btn instanceof HTMLButtonElement)) return;
  const listAction = btn.dataset.listAction;
  const item = btn.closest("[data-id]");
  if (!item || !currentUser) return;
  const id = item?.dataset.id || btn.dataset.caseId || btn.dataset.taskId;
  if (!id) {
    showAppMessage("対象データIDを取得できませんでした。", true);
    return;
  }

  if (listAction === "edit_case") {
    await startCaseEdit(id);
    return;
  }
  if (listAction === "print_case_invoice") {
    openInvoicePrintPreviewFromCase(id);
    return;
  }
  if (listAction === "export_case_invoice_excel") {
    exportInvoiceDataForCase(id);
    return;
  }
  if (listAction === "print_case_delivery_note") return openCaseBusinessDocumentPrintPreview(id, "delivery_note");
  if (listAction === "print_case_purchase_order") return openCaseBusinessDocumentPrintPreview(id, "purchase_order");
  if (listAction === "print_case_order_confirmation") return openCaseBusinessDocumentPrintPreview(id, "order_confirmation");
  if (listAction === "print_case_acceptance_certificate") return openCaseBusinessDocumentPrintPreview(id, "acceptance_certificate");

  if (listAction === "delete_case") {
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

function startCaseTaskEdit(taskId) {
  if (!caseTaskForm) return;
  const target = state.caseTasks.find((entry) => String(entry.id) === String(taskId));
  if (!target) return;
  editState.caseTaskId = target.id;
  caseTaskForm.elements.caseTaskCaseId.value = target.caseId || "";
  caseTaskForm.elements.caseTaskTitle.value = target.taskTitle || "";
  caseTaskForm.elements.caseTaskMemo.value = target.taskMemo || "";
  caseTaskForm.elements.caseTaskDueDate.value = target.dueDate || "";
  caseTaskForm.elements.caseTaskStatus.value = target.status === "完了" ? "完了" : "未完了";
  if (caseTaskSubmitBtn) caseTaskSubmitBtn.textContent = "タスクを更新";
  subtabState.cases = "tasks";
  activateTab("cases");
}

async function completeCaseTask(taskId) {
  if (!currentUser) return;
  try {
    await runMutation("案件タスク完了", async () => {
      const payload = { status: "完了", completed_at: toDateString(new Date()) };
      const { data, error } = await sbClient.from("case_tasks").update(payload).eq("id", taskId).eq("user_id", currentUser.id).select().single();
      if (error) throw error;
      if (!data) throw new Error("更新結果を取得できませんでした。");
      return data;
    }, { successMessage: "案件タスクを完了にしました。" });
  } catch (error) {
    showAppMessage(`案件タスク完了に失敗しました。${formatSupabaseError(error)}`, true);
  }
}

async function deleteCaseTask(taskId) {
  await deleteRecord({
    table: "case_tasks",
    id: taskId,
    actionName: "案件タスク削除",
    confirmMessage: "この案件タスクを削除しますか？",
    afterSuccess: () => {
      if (editState.caseTaskId === taskId) resetCaseTaskForm();
    },
  });
}

function buildCaseDocumentPayloadFromForm() {
  return {
    user_id: currentUser.id,
    case_id: caseDocumentForm.elements.caseDocumentCaseId.value || null,
    document_name: asTrimmedText(caseDocumentForm.elements.caseDocumentName.value),
    status: normalizeCaseDocumentStatus(caseDocumentForm.elements.caseDocumentStatus.value),
    received_date: caseDocumentForm.elements.caseDocumentReceivedDate.value || null,
    checked_date: caseDocumentForm.elements.caseDocumentCheckedDate.value || null,
    memo: asTrimmedText(caseDocumentForm.elements.caseDocumentMemo.value) || null,
    file_url: asTrimmedText(caseDocumentForm.elements.caseDocumentFileUrl.value) || null,
  };
}

async function handleCaseDocumentSubmit(event) {
  event.preventDefault();
  if (!currentUser || !caseDocumentForm || !ensureInitialDataReady("書類登録")) return;
  const payload = buildCaseDocumentPayloadFromForm();
  if (!payload.document_name) return;
  const isEdit = Boolean(editState.caseDocumentId);
  const taskName = isEdit ? "書類更新" : "書類登録";
  try {
    await runMutation(taskName, async () => {
      if (isEdit) {
        const { data, error } = await sbClient.from("case_documents").update(payload).eq("id", editState.caseDocumentId).eq("user_id", currentUser.id).select().single();
        if (error) throw error;
        if (!data) throw new Error("更新結果を取得できませんでした。");
        return data;
      }
      const { data, error } = await sbClient.from("case_documents").insert(payload).select().single();
      if (error) throw error;
      if (!data) throw new Error("登録結果を取得できませんでした。");
      return data;
    }, {
      successMessage: isEdit ? "書類を更新しました。" : "書類を登録しました。",
      resetForm: resetCaseDocumentForm,
      afterSuccess: () => {
        subtabState.cases = "documents";
        activateTab("cases");
      },
    });
  } catch (error) {
    showAppMessage(`書類保存に失敗しました。${formatSupabaseError(error)}`, true);
  }
}

async function handleCaseDocumentsListAction(event) {
  const button = event.target.closest("button");
  if (!(button instanceof HTMLButtonElement)) return;
  const listAction = button.dataset.listAction;
  const id = button.dataset.caseDocumentId || button.closest("[data-id]")?.dataset.id;
  if (!id) return;
  if (listAction === "edit_case_document") {
    startCaseDocumentEdit(id);
    return;
  }
  if (listAction === "delete_case_document") {
    await deleteCaseDocument(id);
  }
}

function startCaseDocumentEdit(caseDocumentId) {
  const target = state.caseDocuments.find((entry) => String(entry.id) === String(caseDocumentId));
  if (!target || !caseDocumentForm) return;
  editState.caseDocumentId = target.id;
  caseDocumentForm.elements.caseDocumentCaseId.value = target.caseId || "";
  caseDocumentForm.elements.caseDocumentName.value = target.documentName || "";
  caseDocumentForm.elements.caseDocumentStatus.value = normalizeCaseDocumentStatus(target.status);
  caseDocumentForm.elements.caseDocumentReceivedDate.value = target.receivedDate || "";
  caseDocumentForm.elements.caseDocumentCheckedDate.value = target.checkedDate || "";
  caseDocumentForm.elements.caseDocumentMemo.value = target.memo || "";
  caseDocumentForm.elements.caseDocumentFileUrl.value = target.fileUrl || "";
  if (caseDocumentSubmitBtn) caseDocumentSubmitBtn.textContent = "書類を更新";
  subtabState.cases = "documents";
  activateTab("cases");
}

async function deleteCaseDocument(caseDocumentId) {
  if (!currentUser || !caseDocumentId) return;
  if (!window.confirm("この書類データを削除しますか？")) return;
  try {
    await runMutation("書類削除", async () => {
      const { data, error } = await sbClient.from("case_documents").delete().eq("id", caseDocumentId).eq("user_id", currentUser.id).select().single();
      if (error) throw error;
      if (!data) throw new Error("削除結果を取得できませんでした。");
      return data;
    }, {
      successMessage: "書類を削除しました。",
      afterSuccess: () => {
        if (editState.caseDocumentId === caseDocumentId) resetCaseDocumentForm();
      },
    });
  } catch (error) {
    showAppMessage(`書類削除に失敗しました。${formatSupabaseError(error)}`, true);
  }
}

function handleCaseTemplateChange(event) {
  const templateId = event?.target?.value || "";
  if (!caseForm) return;
  if (!templateId) {
    caseForm.elements.requiredDocuments.value = "";
    caseForm.elements.taskList.value = "";
    return;
  }
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
  if (Number.isFinite(found.standardEstimateAmount) && found.standardEstimateAmount > 0) {
    caseForm.elements.amount.value = String(found.standardEstimateAmount);
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
  const listAction = btn.dataset.listAction;
  const item = btn.closest("[data-id]");
  if (!item || !currentUser) return;
  const id = item.dataset.id;
  if (!id) return;
  if (listAction === "edit_work_template") {
    startWorkTemplateEdit(id);
    return;
  }
  if (listAction === "delete_work_template") {
    await deleteWorkTemplate(id);
  }
}

function startWorkTemplateEdit(templateId) {
  const target = state.workTemplates.find((entry) => entry.id === templateId);
  if (!target || !workTemplateForm) return;
  editState.workTemplateId = target.id;
  workTemplateForm.elements.templateName.value = target.name;
  workTemplateForm.elements.templateStandardEstimateAmount.value = target.standardEstimateAmount ?? "";
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
  const listAction = btn.dataset.listAction;
  if (!currentUser) return;
  const item = btn.closest("[data-id]");
  const id = item?.dataset.id || btn.dataset.saleId || btn.closest("[data-sale-id]")?.dataset.saleId;
  if (!id) {
    showAppMessage("対象データIDを取得できませんでした。", true);
    return;
  }

  if (listAction === "edit_sale") {
    await editSale(id);
    return;
  }

  if (listAction === "delete_sale") {
    await deleteSale(id);
    return;
  }
  if (listAction === "record_payment") {
    await handleRecordPayment(id);
    return;
  }
  if (listAction === "record_reminder") {
    await handleRecordReminder(id);
    return;
  }
  if (listAction === "delete_payment") {
    const paymentId = btn.dataset.paymentId;
    if (!paymentId) return;
    await deletePayment(paymentId);
    return;
  }
  if (listAction === "print_receipt") {
    const paymentId = btn.dataset.paymentId || null;
    openReceiptPrintPreview(id, paymentId);
    return;
  }
  if (listAction === "print_cumulative_receipt") {
    openReceiptPrintPreview(id, null, { cumulative: true });
  }

}

async function handleExpensesListAction(event) {
  const btn = event.target.closest("button");
  if (!(btn instanceof HTMLButtonElement)) return;
  const listAction = btn.dataset.listAction;
  const item = btn.closest("[data-id]");
  if (!item || !currentUser) return;
  const id = item.dataset.id;
  if (!id) {
    showAppMessage("対象データIDを取得できませんでした。", true);
    return;
  }

  if (listAction === "edit_expense") {
    await editExpense(id);
    return;
  }

  if (listAction === "delete_expense") {
    await deleteExpense(id);
  }
}

function handleUnpaidListAction(event) {
  const btn = event.target.closest("button");
  if (!(btn instanceof HTMLButtonElement)) return;
  const listAction = btn.dataset.listAction;
  const saleId = btn.dataset.saleId;
  if (!saleId) return;
  if (listAction === "edit_sale") {
    editSale(saleId).catch(() => {});
    return;
  }
  if (listAction === "delete_payment") {
    const paymentId = btn.dataset.paymentId;
    if (!paymentId) return;
    deletePayment(paymentId).catch(() => {});
  }
}

function handleTodayTaskAction(event) {
  const btn = event.target.closest("button");
  if (!(btn instanceof HTMLButtonElement)) return;
  if (btn.dataset.listAction === "record_reminder") {
    const saleId = btn.dataset.saleId;
    if (!saleId) return;
    handleRecordReminder(saleId).catch(() => {});
    return;
  }
  const taskId = btn.dataset.taskId;
  const target = btn.dataset.taskTarget;
  if (!taskId || !target) return;
  if (target === "case-task-complete") {
    completeCaseTask(taskId).catch(() => {});
    return;
  }
  if (target === "case") {
    startCaseEdit(taskId).catch(() => {});
    return;
  }
  if (target === "sale") {
    editSale(taskId).catch(() => {});
    return;
  }
  if (target === "case-task") {
    startCaseTaskEdit(taskId);
    return;
  }
  if (target === "case-document") {
    startCaseDocumentEdit(taskId);
    return;
  }
  if (target === "daily-report") {
    editDailyReport(taskId).catch(() => {});
  }
}

function handleDocumentAlertAction(event) {
  const btn = event.target.closest("button");
  if (!(btn instanceof HTMLButtonElement)) return;
  const caseDocumentId = btn.dataset.caseDocumentId;
  if (!caseDocumentId) return;
  startCaseDocumentEdit(caseDocumentId);
}

async function handleCaseTaskSubmit(event) {
  event.preventDefault();
  if (!currentUser || !caseTaskForm || !ensureInitialDataReady("案件タスク登録")) return;
  const payload = {
    user_id: currentUser.id,
    case_id: caseTaskForm.elements.caseTaskCaseId.value || null,
    task_title: asTrimmedText(caseTaskForm.elements.caseTaskTitle.value),
    task_memo: asTrimmedText(caseTaskForm.elements.caseTaskMemo.value) || null,
    due_date: caseTaskForm.elements.caseTaskDueDate.value || null,
    status: caseTaskForm.elements.caseTaskStatus.value === "完了" ? "完了" : "未完了",
    completed_at: null,
  };
  if (!payload.task_title) return;
  if (payload.status === "完了") payload.completed_at = toDateString(new Date());
  const taskName = editState.caseTaskId ? "案件タスク更新" : "案件タスク登録";
  const isEdit = Boolean(editState.caseTaskId);
  try {
    await runMutation(taskName, async () => {
      if (isEdit) {
        const { data, error } = await sbClient.from("case_tasks").update(payload).eq("id", editState.caseTaskId).eq("user_id", currentUser.id).select().single();
        if (error) throw error;
        if (!data) throw new Error("更新結果を取得できませんでした。");
        return data;
      } else {
        const { data, error } = await sbClient.from("case_tasks").insert(payload).select().single();
        if (error) throw error;
        if (!data) throw new Error("登録結果を取得できませんでした。");
        return data;
      }
    }, {
      successMessage: isEdit ? "案件タスクを更新しました。" : "案件タスクを登録しました。",
      resetForm: resetCaseTaskForm,
    });
  } catch (error) {
    showAppMessage(`案件タスク保存に失敗しました。${formatSupabaseError(error)}`, true);
  }
}

function handleCaseTaskListAction(event) {
  const btn = event.target.closest("button");
  if (!(btn instanceof HTMLButtonElement)) return;
  const listAction = btn.dataset.listAction;
  const taskId = btn.dataset.taskId || btn.closest("[data-id]")?.dataset.id;
  if (!taskId) return;
  if (listAction === "edit_case_task") {
    startCaseTaskEdit(taskId);
    return;
  }
  if (listAction === "complete_case_task") {
    completeCaseTask(taskId).catch(() => {});
    return;
  }
  if (listAction === "delete_case_task") {
    deleteCaseTask(taskId).catch(() => {});
  }
}

function handleBillingLeakAlertAction(event) {
  const btn = event.target.closest("button");
  if (!(btn instanceof HTMLButtonElement)) return;
  const caseId = btn.dataset.caseId;
  if (!caseId) return;
  if (btn.dataset.listAction === "register_sale") {
    openSaleFormForCase(caseId);
  }
}

function handlePendingEstimateAction(event) {
  const btn = event.target.closest("button");
  if (!(btn instanceof HTMLButtonElement)) return;
  const estimateId = btn.dataset.estimateId;
  if (!estimateId) return;
  if (btn.dataset.listAction === "edit_pending_estimate") {
    startEstimateEdit(estimateId).catch(() => {});
  }
}

function handleDataIntegrityAction(event) {
  const btn = event.target.closest("button");
  if (!(btn instanceof HTMLButtonElement)) return;
  const checkKey = btn.dataset.checkKey;
  if (!checkKey) return;
  state.selectedIntegrityCheckKey = checkKey;
  safeRender("dataIntegrity", renderDataIntegrityCheck);
}

async function handleRecordReminder(saleId) {
  if (!currentUser) {
    showAppMessage("ログイン状態を確認できません。", true);
    return;
  }
  const sale = state.sales.find((entry) => entry.id === saleId);
  if (!sale) {
    showAppMessage("対象の売上が見つかりません。", true);
    return;
  }
  if (!isReminderRecordableSale(sale)) {
    showAppMessage("督促記録は未入金または一部入金の売上のみ登録できます。", true);
    return;
  }

  const reminderDate = window.prompt("督促日を入力してください（YYYY-MM-DD）", toDateString(new Date()));
  if (!reminderDate) return;
  const normalizedReminderDate = parseFlexibleDate(reminderDate);
  if (!normalizedReminderDate) {
    showAppMessage("督促日の形式が不正です。YYYY-MM-DD形式で入力してください。", true);
    return;
  }

  const methodPrompt = `督促方法を入力してください（${REMINDER_METHODS.join("・")}）`;
  const reminderMethod = window.prompt(methodPrompt, sale.reminderMethod || REMINDER_METHODS[0]);
  if (!reminderMethod) return;
  const normalizedMethod = normalizeReminderMethod(reminderMethod);

  const reminderMemo = window.prompt("督促メモを入力してください", sale.reminderMemo || "");
  if (reminderMemo === null) return;
  const normalizedMemo = asTrimmedText(reminderMemo) || "";

  try {
    await runMutation("督促記録", async () => {
      const currentCount = Number(sale.reminderCount || 0);
      const payload = {
        reminder_count: currentCount + 1,
        last_reminder_date: normalizedReminderDate,
        reminder_method: normalizedMethod,
        reminder_memo: normalizedMemo,
      };
      console.log("REMINDER PAYLOAD", saleId, payload);
      const { data, error } = await sbClient
        .from("sales")
        .update(payload)
        .eq("id", saleId)
        .eq("user_id", currentUser.id)
        .select()
        .single();
      if (error) throw error;
      if (!data) throw new Error("督促記録の更新結果を取得できませんでした。");
      return data;
    }, { successMessage: "督促記録を保存しました。" });
  } catch (error) {
    console.error("督促記録に失敗しました。", error);
    showAppMessage(`督促記録に失敗しました。${error?.message || ""} ${error?.details || ""} ${error?.hint || ""} ${error?.code || ""}`, true);
  }
}

async function handleRecordPayment(saleId) {
  if (!currentUser) {
    showAppMessage("ログイン状態を確認できません。", true);
    return;
  }
  const sale = state.sales.find((entry) => entry.id === saleId);
  if (!sale) {
    showAppMessage("対象の売上が見つかりません。", true);
    return;
  }

  const paymentDateInput = window.prompt("入金日を入力してください（YYYY-MM-DD）", toDateString(new Date()));
  if (!paymentDateInput) return;
  const paymentDate = parseFlexibleDate(paymentDateInput);
  if (!paymentDate) {
    showAppMessage("入金日の形式が不正です。YYYY-MM-DD形式で入力してください。", true);
    return;
  }

  const amountInput = window.prompt("入金額を入力してください", "");
  if (amountInput === null) return;
  if (!String(amountInput).trim()) {
    showAppMessage("入金額を入力してください。", true);
    return;
  }
  const amount = Number(String(amountInput).replace(/,/g, "").trim());
  if (!Number.isFinite(amount) || Number.isNaN(amount)) {
    showAppMessage("入金額は数値で入力してください。", true);
    return;
  }
  if (amount <= 0) {
    showAppMessage("入金額は0より大きい値を入力してください。", true);
    return;
  }

  const methodInput = window.prompt("入金方法を入力してください（現金・振込・その他）", "振込");
  if (!methodInput) return;
  const method = normalizePaymentMethod(methodInput);

  const memoInput = window.prompt("メモを入力してください", "");
  if (memoInput === null) return;
  const memo = asTrimmedText(memoInput) || null;

  try {
    await runMutation("入金登録", async () => {
      console.log("PAYMENT START", { saleId, paymentDate, amount, method, memo });
      const {
        data: { user },
        error: userError,
      } = await sbClient.auth.getUser();
      if (userError || !user) {
        throw userError || new Error("ログインユーザーを確認できません。");
      }
      const authUserId = user.id;
      console.log("PAYMENT AUTH USER", authUserId);
      const { data: saleRow, error: saleLoadError } = await sbClient
        .from("sales")
        .select("id,user_id,invoice_amount,paid_amount,paid_date,payment_status")
        .eq("id", saleId)
        .eq("user_id", authUserId)
        .single();
      if (saleLoadError) throw saleLoadError;
      if (!saleRow) throw new Error("対象売上の取得に失敗しました。");
      const invoiceAmount = Number(saleRow.invoice_amount || 0);

      const { data: paymentRowsBefore, error: paymentLoadError } = await sbClient
        .from("payments")
        .select("amount")
        .eq("sale_id", saleId)
        .eq("user_id", authUserId);
      if (paymentLoadError) throw paymentLoadError;
      const historyPaidBefore = (paymentRowsBefore || []).reduce((sum, payment) => sum + Number(payment.amount || 0), 0);
      const manualPaidBefore = Number(saleRow.paid_amount || 0);
      const paidBaseBefore = Math.max(manualPaidBefore, historyPaidBefore);
      const newPaidAmount = paidBaseBefore + amount;
      if (newPaidAmount > invoiceAmount) {
        throw new Error("入金額が請求額を超えています。");
      }

      const paymentPayload = {
        user_id: authUserId,
        sale_id: saleId,
        payment_date: paymentDate,
        amount: amount,
        method: method,
        memo: memo || null,
      };
      console.log("PAYMENT INSERT PAYLOAD", paymentPayload);

      const { data: insertedPayment, error: paymentError } = await sbClient
        .from("payments")
        .insert(paymentPayload)
        .select()
        .single();

      if (paymentError) {
        console.error("PAYMENT INSERT ERROR", paymentError);
        throw paymentError;
      }
      if (insertedPayment) {
        await appendChangeLog({ action: "create", table: "payments", recordId: insertedPayment.id, before: null, after: mapPaymentFromDb(insertedPayment) });
      }

      const paymentStatus = calculatePaymentStatus(invoiceAmount, newPaidAmount);
      const saleUpdatePayload = {
        paid_amount: newPaidAmount,
        paid_date: paymentDate,
        payment_status: paymentStatus,
        is_unpaid: newPaidAmount < invoiceAmount,
      };
      const { error: saleUpdateError } = await sbClient
        .from("sales")
        .update(saleUpdatePayload)
        .eq("id", saleId)
        .eq("user_id", authUserId);
      if (saleUpdateError) throw saleUpdateError;
      return null;
    }, { successMessage: "入金を登録しました。" });
  } catch (error) {
    console.error("入金登録に失敗しました。", error);
    showAppMessage(`入金登録に失敗しました。${error?.message || ""} ${error?.details || ""} ${error?.hint || ""} ${error?.code || ""}`, true);
  }
}


async function deletePayment(paymentId) {
  if (!currentUser) return;
  const target = state.payments.find((entry) => entry.id === paymentId);
  if (!target) {
    showAppMessage("削除対象の入金履歴が見つかりません。", true);
    return;
  }
  if (!window.confirm("この入金履歴を削除しますか？")) return;

  await runMutation("入金削除", async () => {
    const {
      data: { user },
      error: userError,
    } = await sbClient.auth.getUser();
    if (userError || !user) {
      throw userError || new Error("ログインユーザーを確認できません。");
    }
    const authUserId = user.id;
    const { data: saleRow, error: saleLoadError } = await sbClient
      .from("sales")
      .select("id,user_id,invoice_amount,paid_amount")
      .eq("id", target.saleId)
      .eq("user_id", authUserId)
      .single();
    if (saleLoadError) throw saleLoadError;
    if (!saleRow) throw new Error("対象売上の取得に失敗しました。");

    const { data, error } = await sbClient.from("payments").delete().eq("id", paymentId).eq("user_id", currentUser.id).select().single();
    if (error) throw error;
    if (!data) throw new Error("入金削除結果を取得できませんでした。");
    await appendChangeLog({ action: "delete", table: "payments", recordId: data.id, before: mapPaymentFromDb(data), after: null });
    const deletedAmount = Number(data.amount || 0);
    const invoiceAmount = Number(saleRow.invoice_amount || 0);
    const currentSalePaid = Number(saleRow.paid_amount || 0);
    const newPaidAmount = Math.max(0, currentSalePaid - deletedAmount);
    const paymentStatus = calculatePaymentStatus(invoiceAmount, newPaidAmount);
    const saleUpdatePayload = {
      paid_amount: newPaidAmount,
      paid_date: newPaidAmount > 0 ? (target.paymentDate || null) : null,
      payment_status: paymentStatus,
      is_unpaid: newPaidAmount < invoiceAmount,
    };
    const { error: saleUpdateError } = await sbClient
      .from("sales")
      .update(saleUpdatePayload)
      .eq("id", target.saleId)
      .eq("user_id", authUserId);
    if (saleUpdateError) throw saleUpdateError;
    return data;
  }, { successMessage: "入金履歴を削除しました。" });
}

async function startSaleEdit(saleId) {
    const target = state.sales.find((entry) => entry.id === saleId);
    if (!target) return;
    subtabState.sales = "entry";
    activateTab("sales");
    editState.saleId = target.id;
    if (saleCaseSelect) saleCaseSelect.value = target.caseId || "";
    saleForm.elements.invoiceAmount.value = formatNumberInput(target.invoiceAmount);
    saleForm.elements.paidAmount.value = formatNumberInput(target.paidAmount ?? "");
    saleForm.elements.paidDate.value = target.paidDate || "";
    saleForm.elements.dueDate.value = target.dueDate || "";
    setSaleInvoiceNumberDisplay(target.invoiceNumber || "");
    saleSubmitBtn.textContent = "売上を更新";
    renderPayments();
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
    if (expenseCaseSelect) expenseCaseSelect.value = target.caseId || "";
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
  const listAction = btn.dataset.listAction;
  const item = btn.closest("[data-id]");
  if (!item || !currentUser) return;
  const id = item.dataset.id;
  if (!id) {
    showAppMessage("対象データIDを取得できませんでした。", true);
    return;
  }

  if (listAction === "edit_fixed_expense") {
    await startFixedExpenseEdit(id);
    return;
  }

  if (listAction === "toggle_fixed_expense") {
    const target = state.fixedExpenses.find((entry) => entry.id === id);
    if (!target) return;

    try {
      await runMutation("固定費更新", async () => {
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
        return data;
      }, {
        successMessage: "固定費の有効/無効を更新しました。",
      });
    } catch (error) {
      showAppMessage(`固定費更新に失敗しました。${formatSupabaseError(error)}`, true);
    }
    return;
  }

  if (listAction === "delete_fixed_expense") {
    await deleteFixedExpense(id);
  }
}

async function handleDailyReportsListAction(event) {
  const btn = event.target.closest("button");
  if (!(btn instanceof HTMLButtonElement) || !currentUser) return;
  const listAction = btn.dataset.listAction;
  const item = btn.closest("[data-id]");
  const id = item?.dataset.id;
  if (!id) {
    showAppMessage("対象データIDを取得できませんでした。", true);
    return;
  }

  if (listAction === "edit_daily_report") {
    await editDailyReport(id);
    return;
  }

  if (listAction === "delete_daily_report") {
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
  const sale = state.sales.find((entry) => entry.id === saleId);
  if (!sale) {
    showAppMessage("編集対象の売上が見つかりません。", true);
    return;
  }
  editState.saleId = saleId;
  if (saleCaseSelect) saleCaseSelect.value = sale.caseId || "";
  document.getElementById("invoice-amount").value = formatNumberInput(String(sale.invoiceAmount || 0));
  document.getElementById("paid-amount").value = formatNumberInput(String(sale.paidAmount || 0));
  document.getElementById("paid-date").value = sale.paidDate || "";
  document.getElementById("sale-due-date").value = sale.dueDate || "";
  setSaleInvoiceNumberDisplay(sale.invoiceNumber || "");
  saleSubmitBtn.textContent = "売上を更新";
  activateTab("sales");
  activateSubtab("sales", "entry");
  renderPayments();
  saleForm.scrollIntoView({ behavior: "smooth", block: "start" });
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

async function runMutation(actionName, mutationFn, options = {}) {
  startLoading(actionName);

  try {
    clearAppMessage();

    console.log("MUTATION START", actionName);

    const result = await mutationFn();

    await loadAllDataSafely();
    renderAfterDataChanged();

    if (typeof options.afterSuccess === "function") {
      await options.afterSuccess(result);
    }

    if (typeof options.resetForm === "function") {
      options.resetForm();
    }

    if (options.log) {
      try {
        const logPayload = typeof options.log === "function"
          ? options.log(result)
          : options.log;
        await addOperationLog(logPayload);
      } catch (logError) {
        console.error("operation log failed", logError);
      }
    }

    console.log("MUTATION SUCCESS", actionName);
    showAppMessage(options.successMessage || `${actionName}が完了しました。`, false);

    return result;
  } catch (error) {
    console.error(`${actionName} failed`, error);
    showAppMessage(
      `${actionName}に失敗しました。${error.message || ""} ${error.details || ""} ${error.hint || ""} ${error.code || ""}`,
      true
    );
    return null;
  } finally {
    forceHideLoading();
  }
}

async function deleteRecord({ table, id, actionName, confirmMessage, beforeDelete, afterSuccess }) {
  if (!currentUser) {
    showAppMessage("ログイン状態を確認できません。", true);
    return;
  }
  if (!id) {
    showAppMessage(`${actionName}対象IDを取得できませんでした。`, true);
    return;
  }
  if (confirmMessage && !window.confirm(confirmMessage)) return;

  await runMutation(actionName, async () => {
    if (beforeDelete) await beforeDelete(id);
    const { data, error } = await sbClient
      .from(table)
      .delete()
      .eq("id", id)
      .eq("user_id", currentUser.id)
      .select()
      .single();
    if (error) throw error;
    return data;
  }, {
    successMessage: `${actionName}しました。`,
    afterSuccess,
  });
}

async function deleteClient(id) {
  await deleteRecord({
    table: "clients",
    id,
    actionName: "顧客削除",
    confirmMessage: "この顧客を削除しますか？案件や見積は削除されません。",
    beforeDelete: async (clientId) => {
      const casesUpdateRes = await sbClient.from("cases").update({ client_id: null }).eq("client_id", clientId).eq("user_id", currentUser.id);
      if (casesUpdateRes.error) throw casesUpdateRes.error;
      const estimatesUpdateRes = await sbClient.from("estimates").update({ client_id: null }).eq("client_id", clientId).eq("user_id", currentUser.id);
      if (estimatesUpdateRes.error) throw estimatesUpdateRes.error;
      const reportsUpdateRes = await sbClient.from("daily_reports").update({ client_id: null }).eq("client_id", clientId).eq("user_id", currentUser.id);
      if (reportsUpdateRes.error) throw reportsUpdateRes.error;
    },
    afterSuccess: () => {
      if (editState.clientId === id) {
        editState.clientId = null;
        resetClientForm();
      }
    },
  });
}

async function deleteCase(id) {
  await deleteRecord({
    table: "cases",
    id,
    actionName: "案件削除",
    confirmMessage: "この案件を削除しますか？関連する売上・経費も削除されます。",
    beforeDelete: async (caseId) => {
      const caseSales = state.sales.filter((entry) => entry.caseId === caseId);
      for (const sale of caseSales) {
        const paymentDeleteRes = await sbClient.from("payments").delete().eq("sale_id", sale.id).eq("user_id", currentUser.id);
        if (paymentDeleteRes.error) throw paymentDeleteRes.error;
      }
      const salesDeleteRes = await sbClient.from("sales").delete().eq("case_id", caseId).eq("user_id", currentUser.id);
      if (salesDeleteRes.error) throw salesDeleteRes.error;
      const caseTasksDeleteRes = await sbClient.from("case_tasks").delete().eq("case_id", caseId).eq("user_id", currentUser.id);
      if (caseTasksDeleteRes.error) throw caseTasksDeleteRes.error;
      const caseDocumentsDeleteRes = await sbClient.from("case_documents").delete().eq("case_id", caseId).eq("user_id", currentUser.id);
      if (caseDocumentsDeleteRes.error) throw caseDocumentsDeleteRes.error;
      const expensesDeleteRes = await sbClient.from("expenses").delete().eq("case_id", caseId).eq("user_id", currentUser.id);
      if (expensesDeleteRes.error) throw expensesDeleteRes.error;
      const reportsUpdateRes = await sbClient.from("daily_reports").update({ case_id: null }).eq("case_id", caseId).eq("user_id", currentUser.id);
      if (reportsUpdateRes.error) throw reportsUpdateRes.error;
      const estimatesUpdateRes = await sbClient.from("estimates").update({ case_id: null }).eq("case_id", caseId).eq("user_id", currentUser.id);
      if (estimatesUpdateRes.error) throw estimatesUpdateRes.error;
    },
    afterSuccess: () => {
      if (editState.caseId === id) {
        editState.caseId = null;
        resetCaseForm();
      }
    },
  });
}

async function deleteWorkTemplate(id) {
  await deleteRecord({
    table: "work_templates",
    id,
    actionName: "業務テンプレート削除",
    confirmMessage: "この業務テンプレートを削除しますか？",
    beforeDelete: async (templateId) => {
      const caseUpdateRes = await sbClient.from("cases").update({ template_id: null }).eq("template_id", templateId).eq("user_id", currentUser.id);
      if (caseUpdateRes.error) throw caseUpdateRes.error;
    },
    afterSuccess: () => {
      if (editState.workTemplateId === id) resetWorkTemplateForm();
    },
  });
}

async function deleteSale(id) {
  console.log("DELETE SALE CLICKED", id);
  await deleteRecord({
    table: "sales",
    id,
    actionName: "売上削除",
    confirmMessage: "この売上を削除しますか？関連する入金履歴も削除されます。",
    beforeDelete: async (saleId) => {
      const { error } = await sbClient.from("payments").delete().eq("sale_id", saleId).eq("user_id", currentUser.id);
      if (error) console.error("payments削除エラー", error);
    },
    afterSuccess: () => {
      editState.saleId = null;
      resetSaleForm?.();
      activateSubtab?.("sales", "list");
    },
  });
}

async function deleteExpense(id) {
  const targetExpense = state.expenses.find((entry) => entry.id === id) || null;
  await deleteRecord({
    table: "expenses",
    id,
    actionName: "経費削除",
    confirmMessage: "この経費を削除しますか？",
    afterSuccess: () => {
      appendChangeLog({ action: "delete", table: "expenses", recordId: id, before: targetExpense, after: null }).catch((error) => {
        console.error("change log failed", error);
      });
      if (editState.expenseId === id) {
        editState.expenseId = null;
        resetExpenseForm();
      }
    },
  });
}

async function deleteFixedExpense(id) {
  await deleteRecord({
    table: "fixed_expenses",
    id,
    actionName: "固定費削除",
    confirmMessage: "この固定費を削除しますか？",
    afterSuccess: () => {
      if (editState.fixedExpenseId === id) {
        editState.fixedExpenseId = null;
        resetFixedExpenseForm();
      }
    },
  });
}

async function deleteDailyReport(id) {
  await deleteRecord({
    table: "daily_reports",
    id,
    actionName: "日報削除",
    confirmMessage: "この日報を削除しますか？",
    afterSuccess: () => {
      if (editState.dailyReportId === id) {
        editState.dailyReportId = null;
        resetDailyReportForm();
      }
    },
  });
}

async function deleteAllData() {
  if (!currentUser) return;
  if (!window.confirm("顧客・対応履歴・業務テンプレート・見積・案件・売上・入金・経費・固定費・日報の全データを削除します。よろしいですか？")) return;
  try {
    await runMutation("全件削除", async () => {
      const [estimateItemsRes, estimatesRes, paymentsRes, salesRes, expensesRes, fixedExpensesRes, dailyReportsRes, caseDocumentsRes, casesRes, workTemplatesRes, clientsRes] = await Promise.all([
        sbClient.from("estimate_items").delete().eq("user_id", currentUser.id),
        sbClient.from("estimates").delete().eq("user_id", currentUser.id),
        sbClient.from("payments").delete().eq("user_id", currentUser.id),
        sbClient.from("sales").delete().eq("user_id", currentUser.id),
        sbClient.from("expenses").delete().eq("user_id", currentUser.id),
        sbClient.from("fixed_expenses").delete().eq("user_id", currentUser.id),
        sbClient.from("daily_reports").delete().eq("user_id", currentUser.id),
        sbClient.from("case_documents").delete().eq("user_id", currentUser.id),
        sbClient.from("cases").delete().eq("user_id", currentUser.id),
        sbClient.from("work_templates").delete().eq("user_id", currentUser.id),
        sbClient.from("clients").delete().eq("user_id", currentUser.id),
      ]);
      if (salesRes.error) throw salesRes.error;
      if (paymentsRes.error) throw paymentsRes.error;
      if (expensesRes.error) throw expensesRes.error;
      if (estimateItemsRes.error) throw estimateItemsRes.error;
      if (estimatesRes.error) throw estimatesRes.error;
      if (fixedExpensesRes.error) throw fixedExpensesRes.error;
      if (dailyReportsRes.error) throw dailyReportsRes.error;
      if (caseDocumentsRes.error) throw caseDocumentsRes.error;
      if (casesRes.error) throw casesRes.error;
      if (workTemplatesRes.error) throw workTemplatesRes.error;
      if (clientsRes.error) throw clientsRes.error;
      return true;
    }, {
      successMessage: "削除しました。",
      afterSuccess: () => {
        resetClientForm();
        resetCaseForm();
        resetCaseDocumentForm();
        resetWorkTemplateForm();
        resetEstimateForm();
        resetSaleForm();
        resetExpenseForm();
        resetFixedExpenseForm();
        resetDailyReportForm();
      },
    });
  } catch (error) {
    console.error("削除に失敗しました。", error);
    showAppMessage(`削除に失敗しました。${error.message || ""}`, true);
  }
}

function handleExportCasesCsv() {
  safeFileExport("案件CSV", () => {
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
  });
}

function handleExportSalesCsv() {
  safeFileExport("売上CSV", () => {
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
  });
}

function handleExportPaymentsCsv() {
  safeFileExport("入金CSV", () => {
    const rows = state.payments.map((entry) => {
      const sale = state.sales.find((row) => row.id === entry.saleId);
      const linkedCase = sale?.caseId ? state.cases.find((c) => c.id === sale.caseId) : null;
      return {
        sale_invoice_number: sale?.invoiceNumber || "",
        customer_name: linkedCase?.customerName || "",
        case_name: linkedCase?.caseName || "",
        payment_date: entry.paymentDate || "",
        amount: entry.amount ?? "",
        method: entry.method || "",
        memo: entry.memo || "",
      };
    });
    downloadCsvFile("payments.csv", ["sale_invoice_number", "customer_name", "case_name", "payment_date", "amount", "method", "memo"], rows);
  });
}

function handleExportExpensesCsv() {
  safeFileExport("経費CSV", () => {
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
  });
}

function handleExportFixedExpensesCsv() {
  safeFileExport("固定費CSV", () => {
    const rows = state.fixedExpenses.map((entry) => ({
      content: entry.content,
      amount: entry.amount ?? "",
      day_of_month: entry.dayOfMonth ?? "",
      start_date: entry.startDate || "",
      active: entry.active ? "true" : "false",
    }));
    downloadCsvFile("fixed_expenses.csv", ["content", "amount", "day_of_month", "start_date", "active"], rows);
  });
}

function handleExportCaseDocumentsCsv() {
  safeFileExport("書類管理CSV", () => {
    const rows = state.caseDocuments.map((entry) => {
      const linkedCase = state.cases.find((row) => row.id === entry.caseId);
      return {
        customer_name: linkedCase?.customerName || "",
        case_name: linkedCase?.caseName || "",
        document_name: entry.documentName || "",
        status: entry.status || "未回収",
        received_date: entry.receivedDate || "",
        checked_date: entry.checkedDate || "",
        memo: entry.memo || "",
        file_url: entry.fileUrl || "",
      };
    });
    downloadCsvFile("case_documents.csv", ["customer_name", "case_name", "document_name", "status", "received_date", "checked_date", "memo", "file_url"], rows);
  });
}

function handleExportAllCsv() {
  safeFileExport("全データCSV", () => {
  const headers = [
    "data_type",
    "client_name", "client_type", "address", "tel", "email", "referral_source", "client_memo",
    "client_id", "customer_name", "case_name", "estimate_amount", "received_date", "due_date", "status", "work_memo", "next_action_date", "next_action", "document_url", "invoice_url", "receipt_url",
    "estimate_number", "invoice_number", "invoice_amount", "paid_amount", "remaining_amount", "paid_date", "due_date", "payment_status", "is_unpaid", "reminder_count", "last_reminder_date", "reminder_method", "reminder_memo",
    "sale_invoice_number", "payment_date", "payment_amount", "payment_method", "payment_memo",
    "date", "content", "amount", "payee", "payment_method",
    "day_of_month", "start_date", "active",
    "report_date", "report_client_id", "report_client_name", "report_case_name", "report_interaction_type", "report_work_content", "report_work_minutes", "report_next_action", "report_next_action_date", "report_memo",
    "task_case_name", "task_title", "task_memo", "task_due_date", "task_status", "task_completed_at",
    "doc_customer_name", "doc_case_name", "doc_document_name", "doc_status", "doc_received_date", "doc_checked_date", "doc_memo", "doc_file_url",
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
  state.payments.forEach((entry) => {
    const sale = state.sales.find((row) => row.id === entry.saleId);
    rows.push({
      data_type: "payment",
      sale_invoice_number: sale?.invoiceNumber || "",
      payment_date: entry.paymentDate || "",
      payment_amount: entry.amount ?? "",
      payment_method: entry.method || "",
      payment_memo: entry.memo || "",
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
  state.caseDocuments.forEach((entry) => {
    const linkedCase = state.cases.find((row) => row.id === entry.caseId);
    rows.push({
      data_type: "case_document",
      doc_customer_name: linkedCase?.customerName || "",
      doc_case_name: linkedCase?.caseName || "",
      doc_document_name: entry.documentName || "",
      doc_status: entry.status || "未回収",
      doc_received_date: entry.receivedDate || "",
      doc_checked_date: entry.checkedDate || "",
      doc_memo: entry.memo || "",
      doc_file_url: entry.fileUrl || "",
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
  state.caseTasks.forEach((entry) => {
    const foundCase = state.cases.find((c) => c.id === entry.caseId);
    rows.push({
      data_type: "case_task",
      customer_name: foundCase?.customerName || "顧客不明",
      task_case_name: foundCase?.caseName || "案件なし",
      task_title: entry.taskTitle || "",
      task_memo: entry.taskMemo || "",
      task_due_date: entry.dueDate || "",
      task_status: entry.status || "未完了",
      task_completed_at: entry.completedAt || "",
    });
  });

  downloadCsvFile("all_data.csv", headers, rows);
  });
}

function handleExportClientAnalysisCsv() {
  safeFileExport("顧客別分析CSV", () => {
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
  });
}

function handleExportReferralAnalysisCsv() {
  safeFileExport("紹介元別分析CSV", () => {
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
  });
}

async function handleCsvImportSubmit(event) {
  event.preventDefault();
  if (!currentUser || !csvImportForm) return;
  const importType = csvImportForm.elements.csvImportType.value;
  const file = csvImportForm.elements.csvImportFile.files?.[0];
  if (!file) return;
  if (!window.confirm(`「${file.name}」を${importTypeToLabel(importType)}として取り込みます。よろしいですか？`)) return;

  await runMutation("CSV取込", async () => {
      const text = await readCsvFileTextWithEncodingFallback(file);
      const result = await importCsvByType(importType, text);
      return result;
    }, {
      successMessage: "CSV取込が完了しました。",
      resetForm: () => csvImportForm.reset(),
      afterSuccess: (result) => {
        if (!result) return;
        const message = `CSV取込完了: 登録件数 ${result.insertedCount}件 / スキップ件数 ${result.skippedCount}件 / エラー件数 ${result.errorCount}件`;
        showAppMessage(message, result.errorCount > 0);
      },
    });
}

function handleExportExcel() {
  safeFileExport("Excel出力", () => {
    exportExcel();
  });
}

function handleExportAnalysisExcel() {
  safeFileExport("分析Excel出力", () => {
    exportAnalysisExcel();
  });
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
    const paymentHeaders = ["sale_invoice_number", "payment_date", "amount", "method", "memo"];
    const expenseHeaders = ["case_name", "date", "content", "amount", "payee", "payment_method", "receipt_url"];
    const fixedExpenseHeaders = ["content", "amount", "day_of_month", "start_date", "active"];
    const dailyReportHeaders = ["report_date", "client_id", "client_name", "case_name", "interaction_type", "work_content", "work_minutes", "next_action", "next_action_date", "memo"];
    const caseTaskHeaders = ["customer_name", "case_name", "task_title", "task_memo", "due_date", "status", "completed_at"];
    const caseDocumentHeaders = ["customer_name", "case_name", "document_name", "status", "received_date", "checked_date", "memo", "file_url"];
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
    const paymentRows = state.payments.map((entry) => {
      const sale = state.sales.find((row) => row.id === entry.saleId);
      return {
        sale_invoice_number: sale?.invoiceNumber || "",
        payment_date: entry.paymentDate || "",
        amount: entry.amount ?? "",
        method: entry.method || "",
        memo: entry.memo || "",
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
    const caseTaskRows = state.caseTasks.map((entry) => {
      const linkedCase = state.cases.find((row) => row.id === entry.caseId);
      return {
        customer_name: linkedCase?.customerName || "顧客不明",
        case_name: linkedCase?.caseName || "案件なし",
        task_title: entry.taskTitle || "",
        task_memo: entry.taskMemo || "",
        due_date: entry.dueDate || "",
        status: entry.status || "未完了",
        completed_at: entry.completedAt || "",
      };
    });
    const caseDocumentRows = state.caseDocuments.map((entry) => {
      const linkedCase = state.cases.find((row) => row.id === entry.caseId);
      return {
        customer_name: linkedCase?.customerName || "",
        case_name: linkedCase?.caseName || "",
        document_name: entry.documentName || "",
        status: entry.status || "未回収",
        received_date: entry.receivedDate || "",
        checked_date: entry.checkedDate || "",
        memo: entry.memo || "",
        file_url: entry.fileUrl || "",
      };
    });

    XLSX.utils.book_append_sheet(workbook, createExcelSheet(clientRows, clientHeaders), "顧客");
    XLSX.utils.book_append_sheet(workbook, createExcelSheet(caseRows, caseHeaders), "案件");
    XLSX.utils.book_append_sheet(workbook, createExcelSheet(saleRows, saleHeaders), "売上");
    XLSX.utils.book_append_sheet(workbook, createExcelSheet(paymentRows, paymentHeaders), "入金");
    XLSX.utils.book_append_sheet(workbook, createExcelSheet(expenseRows, expenseHeaders), "経費");
    XLSX.utils.book_append_sheet(workbook, createExcelSheet(fixedExpenseRows, fixedExpenseHeaders), "固定費");
    XLSX.utils.book_append_sheet(workbook, createExcelSheet(dailyReportRows, dailyReportHeaders), "日報");
    XLSX.utils.book_append_sheet(workbook, createExcelSheet(caseTaskRows, caseTaskHeaders), "案件タスク");
    XLSX.utils.book_append_sheet(workbook, createExcelSheet(caseDocumentRows, caseDocumentHeaders), "書類管理");
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

  await runMutation("Excel取込", async () => {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array", cellDates: false });
      const result = await importWorkbookBySheet(workbook);
      return result;
    }, {
      successMessage: "Excel取込が完了しました。",
      resetForm: () => excelImportForm.reset(),
      afterSuccess: (result) => {
        if (!result) return;
        const message = `Excel取込完了: 登録件数 ${result.insertedCount}件 / スキップ件数 ${result.skippedCount}件 / エラー件数 ${result.errorCount}件`;
        showAppMessage(message, result.errorCount > 0);
      },
    });
}

function handleExportBackupJson(event) {
  if (event) event.preventDefault();
  if (!currentUser) return;
  console.log("BACKUP CLICK FIRED");
  safeFileExport("全データバックアップJSON", () => {
    const backup = buildBackupJson();
    const date = new Date().toISOString().slice(0, 10);
    const filename = `gyosei-app-backup-${date}.json`;
    const counts = {
      clients: backup.data.clients.length,
      cases: backup.data.cases.length,
      sales: backup.data.sales.length,
      expenses: backup.data.expenses.length,
      daily_reports: backup.data.daily_reports.length,
      case_documents: backup.data.case_documents.length,
      permit_hearings: backup.data.permit_hearings.length,
      payments: backup.data.payments.length,
      case_tasks: backup.data.case_tasks.length,
      app_settings: backup.data.app_settings.length,
    };
    console.log("BACKUP JSON COUNTS", counts);
    downloadJsonFile(filename, backup);
    safeSetLocalStorage(BACKUP_AT_STORAGE_KEY, new Date().toISOString());
    safeRender("dataIntegrity", renderDataIntegrityCheck);
  });
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

  await runMutation("バックアップ復元", async () => {
      const text = await file.text();
      let parsed;
      try {
        parsed = JSON.parse(text);
      } catch (_error) {
        throw new Error("JSON形式が不正です。");
      }
      validateBackupJsonPayload(parsed);
      const result = await restoreBackupData(parsed.data, mode);
      return result;
    }, {
      successMessage: "バックアップ復元が完了しました。",
      resetForm: () => backupRestoreForm.reset(),
      afterSuccess: (result) => {
        if (!result) return;
        const message = `復元完了: 登録件数 ${result.insertedCount}件 / スキップ件数 ${result.skippedCount}件 / エラー件数 ${result.errorCount}件`;
        showAppMessage(message, result.errorCount > 0);
      },
    });
}

function buildBackupJson() {
  return {
    app: "gyosei-app",
    version: "1.0",
    exported_at: new Date().toISOString(),
    data: {
      clients: Array.isArray(state.clients) ? state.clients : [],
      cases: Array.isArray(state.cases) ? state.cases : [],
      estimates: Array.isArray(state.estimates) ? state.estimates : [],
      estimate_items: Array.isArray(state.estimateItems) ? state.estimateItems : [],
      sales: Array.isArray(state.sales) ? state.sales : [],
      payments: Array.isArray(state.payments) ? state.payments : [],
      expenses: Array.isArray(state.expenses) ? state.expenses : [],
      fixed_expenses: Array.isArray(state.fixedExpenses) ? state.fixedExpenses : [],
      daily_reports: Array.isArray(state.dailyReports) ? state.dailyReports : [],
      work_templates: Array.isArray(state.workTemplates) ? state.workTemplates : [],
      case_tasks: Array.isArray(state.caseTasks) ? state.caseTasks : [],
      case_documents: Array.isArray(state.caseDocuments) ? state.caseDocuments : [],
      permit_hearings: Array.isArray(state.permitHearings) ? state.permitHearings : [],
      app_settings: state.appSettings?.id ? [{
        id: state.appSettings.id,
        user_id: currentUser?.id || state.appSettings.userId || null,
        office_name: state.appSettings.officeName || null,
        postal_code: state.appSettings.postalCode || null,
        address: state.appSettings.address || null,
        tel: state.appSettings.tel || null,
        email: state.appSettings.email || null,
        invoice_registration_number: state.appSettings.invoiceRegistrationNumber || null,
        bank_info: state.appSettings.bankInfo || null,
        default_invoice_due_days: getDefaultInvoiceDueDays(),
        tax_rate: getCurrentTaxRate(),
        estimate_note: state.appSettings.estimateNote || null,
        invoice_note: state.appSettings.invoiceNote || null,
      }] : [],
    },
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

function safeClassName(className) {
  return String(className || "").trim().replace(/[^\w\-\s]/g, "").replace(/\s+/g, " ").trim();
}

function safeAddClass(element, className) {
  if (!element || !className) return;
  const cleaned = safeClassName(className);
  if (!cleaned) return;
  element.className = safeClassName(`${element.className || ""} ${cleaned}`);
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
  safeRender("caseTasks", renderCaseTasks);
  safeRender("caseTaskAlerts", renderCaseTaskAlerts);
  safeRender("caseDocuments", renderCaseDocuments);
  safeRender("documentAlerts", renderDocumentAlerts);
  safeRender("cases", renderCases);
  safeRender("estimates", renderEstimates);
  safeRender("estimateTotals", recalcEstimateTotals);
  safeRender("sales", renderSales);
  safeRender("payments", renderPayments);
  safeRender("expenses", renderExpenses);
  safeRender("fixedExpenses", renderFixedExpenses);
  safeRender("dailyReports", renderDailyReports);
  safeRender("settingsForm", renderSettingsForm);
  safeRender("changeLogs", renderChangeLogs);
  safeRender("todayTasks", renderTodayTasks);
  safeRender("dashboard", renderDashboard);
  safeRender("unpaidAlerts", renderUnpaidAlerts);
  safeRender("billingLeakAlerts", renderBillingLeakAlerts);
  safeRender("deadlineAlerts", renderDeadlineAlerts);
  safeRender("nextActionAlerts", renderNextActionAlerts);
  safeRender("pendingEstimates", renderPendingEstimates);
  safeRender("dataIntegrity", renderDataIntegrityCheck);
  safeRender("clientAnalysis", renderClientAnalysis);
  safeRender("referralAnalysis", renderReferralAnalysis);
  hydrateActionButtons();
  console.log("RENDER DONE");
}

function formatShortId(value) {
  const text = `${value || ""}`.trim();
  if (!text) return "-";
  if (text.length <= 16) return text;
  return `${text.slice(0, 8)}...${text.slice(-6)}`;
}

function renderChangeLogs() {
  if (!changeLogsBody || !changeLogsEmpty || !changeLogsWrap) return;
  const logs = Array.isArray(state.changeLogs) ? state.changeLogs : [];
  const actionLabelMap = { create: "登録", update: "更新", delete: "削除" };
  const tableLabelMap = { sales: "売上", payments: "入金", expenses: "経費" };
  changeLogsBody.innerHTML = "";
  changeLogsEmpty.hidden = logs.length > 0;
  changeLogsWrap.hidden = logs.length === 0;
  logs.forEach((log) => {
    let detail = {};
    try { detail = log?.detail ? JSON.parse(log.detail) : {}; } catch (_) { detail = {}; }
    const actionType = `${log?.action_type || ""}`.toLowerCase();
    const targetType = `${log?.target_type || ""}`.toLowerCase();
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(formatDateTimeValue(log?.created_at || log?.createdAt || ""))}</td>
      <td>${escapeHtml(formatShortId(log?.user_id))}</td>
      <td>${escapeHtml(actionLabelMap[actionType] || (log?.action_type || "-"))}</td>
      <td>${escapeHtml(tableLabelMap[targetType] || (log?.target_type || "-"))}</td>
      <td>${escapeHtml(formatShortId(log?.target_id))}</td>
      <td><pre>${escapeHtml(JSON.stringify(detail?.before ?? null, null, 2))}</pre></td>
      <td><pre>${escapeHtml(JSON.stringify(detail?.after ?? null, null, 2))}</pre></td>
    `;
    changeLogsBody.appendChild(tr);
  });
}

function renderPayments() {
  if (!salePaymentHistory) return;
  const targetSaleId = editState.saleId;
  if (!targetSaleId) {
    salePaymentHistory.innerHTML = '<p class="meta">売上編集時に入金履歴を表示します。</p>';
    return;
  }
  const html = renderPaymentHistoryInline(targetSaleId);
  salePaymentHistory.innerHTML = html === "-" ? '<p class="meta">入金履歴はありません。</p>' : html;
}

function renderTodayTasks() {
  if (!todayTaskCard || !todayTaskSummary || !todayTaskBody || !todayTaskEmpty || !todayTaskListWrap) return;
  safeRender("todayTaskCard", () => {
    const alerts = buildIntegratedAlerts();
    state.alerts = alerts;
    todayTaskCard.classList.toggle("has-alert", alerts.length > 0);
    todayTaskSummary.textContent = `対象 ${alerts.length}件`;
    todayTaskBody.innerHTML = "";
    todayTaskEmpty.hidden = alerts.length > 0;
    todayTaskListWrap.hidden = alerts.length === 0;
    if (!alerts.length) return;
    alerts.forEach((alert) => {
      const tr = document.createElement("tr");
      tr.className = safeClassName(`task-priority-${alert.priority || "low"}`);
      tr.innerHTML = `
        <td>${escapeHtml(alert.typeLabel || "通知")}</td>
        <td>${escapeHtml(alert.clientName || "顧客不明")}</td>
        <td>${escapeHtml(alert.caseName || "案件なし")}</td>
        <td>${escapeHtml(alert.title || "-")}</td>
        <td>${escapeHtml(alert.dateLabel || "-")}</td>
        <td>${escapeHtml(alert.subInfo || "-")}</td>
        <td><button type="button" class="secondary-btn" data-action="alert_task_edit" data-task-target="${alert.target || "case"}" data-task-id="${alert.id || ""}">編集</button></td>
        <td>${alert.actionButtonHtml || "-"}</td>
      `;
      todayTaskBody.appendChild(tr);
    });
  });
}

function renderCaseTasks() {
  renderCaseTaskOptions();
  renderCaseTasksTable();
}

function renderCaseTaskAlerts() {
  renderDeadlineAlertCard();
}

function renderDocumentAlerts() {
  if (!documentAlertCard || !documentAlertSummary || !documentAlertBody || !documentAlertEmpty || !documentAlertListWrap) return;
  const targets = getDocumentAlertTargets();
  documentAlertCard.classList.toggle("has-alert", targets.length > 0);
  documentAlertSummary.textContent = `対象 ${targets.length}件`;
  documentAlertBody.innerHTML = "";
  documentAlertEmpty.hidden = targets.length > 0;
  documentAlertListWrap.hidden = targets.length === 0;
  targets.forEach((entry) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(entry.typeLabel)}</td>
      <td>${escapeHtml(entry.clientName)}</td>
      <td>${escapeHtml(entry.caseName)}</td>
      <td>${escapeHtml(entry.documentName)}</td>
      <td>${escapeHtml(entry.subInfo)}</td>
      <td><button type="button" class="secondary-btn" data-action="edit_case_document" data-list-action="edit_case_document" data-case-document-id="${entry.id}">編集</button></td>
    `;
    documentAlertBody.appendChild(tr);
  });
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
      <td><button type="button" class="secondary-btn edit-pending-estimate-btn" data-action="edit_pending_estimate" data-list-action="edit_pending_estimate" data-estimate-id="${entry.id}">編集</button></td>
    `;
    pendingEstimatesBody.appendChild(tr);
  });
}

function renderDataIntegrityCheck() {
  if (!dataIntegrityCard || !dataIntegritySummary || !dataIntegrityList || !dataIntegrityDetailWrap || !dataIntegrityDetailTitle || !dataIntegrityDetailList) return;
  const checks = getDataIntegrityChecks();
  const criticalCount = checks.filter((entry) => entry.severity === "critical").reduce((sum, entry) => sum + entry.items.length, 0);
  const warningCount = checks.filter((entry) => entry.severity === "warning").reduce((sum, entry) => sum + entry.items.length, 0);
  const infoCount = checks.filter((entry) => entry.severity === "info").reduce((sum, entry) => sum + entry.items.length, 0);
  dataIntegritySummary.textContent = `重大 ${criticalCount}件 / 要確認 ${warningCount}件 / 参考 ${infoCount}件`;
  renderBackupStatusForIntegrityCheck();

  dataIntegrityCard.classList.toggle("has-alert", checks.some((entry) => entry.items.length > 0));
  dataIntegrityList.innerHTML = "";
  checks.forEach((entry) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = `data-integrity-item severity-${entry.severity}`;
    button.dataset.checkKey = entry.key;
    button.dataset.action = "select_integrity_check";
    button.innerHTML = `<span>${escapeHtml(entry.label)}</span><span class="count">${entry.items.length}件</span>`;
    dataIntegrityList.appendChild(button);
  });

  const selected = checks.find((entry) => entry.key === state.selectedIntegrityCheckKey)
    || checks.find((entry) => entry.items.length > 0)
    || checks[0];
  if (!selected) {
    dataIntegrityDetailWrap.hidden = true;
    return;
  }
  state.selectedIntegrityCheckKey = selected.key;
  dataIntegrityDetailTitle.textContent = `${selected.label}（${selected.items.length}件）`;
  dataIntegrityDetailList.innerHTML = "";
  if (!selected.items.length) {
    const li = document.createElement("li");
    li.textContent = "対象データはありません。";
    dataIntegrityDetailList.appendChild(li);
  } else {
    selected.items.forEach((item) => {
      const li = document.createElement("li");
      li.textContent = item;
      dataIntegrityDetailList.appendChild(li);
    });
  }
  dataIntegrityDetailWrap.hidden = false;
}

function getDataIntegrityChecks() {
  const caseMap = new Map((Array.isArray(state.cases) ? state.cases : []).map((entry) => [entry.id, entry]));
  const clientMap = new Map((Array.isArray(state.clients) ? state.clients : []).map((entry) => [entry.id, entry]));
  const paidStatuses = new Set(["入金済"]);
  const todayTs = getTodayTimestamp();

  const checks = [
    { key: "cases_without_clients", label: "顧客なし案件", severity: "warning", items: (Array.isArray(state.cases) ? state.cases : []).filter((entry) => !entry?.clientId || !clientMap.has(entry.clientId)).map((entry) => `${entry?.caseName || "案件名未設定"} / ID:${entry?.id || "-"}`) },
    { key: "sales_without_cases", label: "案件なし売上", severity: "warning", items: (Array.isArray(state.sales) ? state.sales : []).filter((entry) => !entry?.caseId || !caseMap.has(entry.caseId)).map((entry) => `${entry?.invoiceNumber || "請求番号未設定"} / 売上ID:${entry?.id || "-"}`) },
    { key: "expenses_without_cases", label: "案件なし経費", severity: "info", items: (Array.isArray(state.expenses) ? state.expenses : []).filter((entry) => !entry?.caseId || !caseMap.has(entry.caseId)).map((entry) => `${entry?.content || "経費名未設定"} / 経費ID:${entry?.id || "-"}`) },
    { key: "sales_zero_invoice_amount", label: "請求額0円の売上", severity: "warning", items: (Array.isArray(state.sales) ? state.sales : []).filter((entry) => Number(entry?.invoiceAmount || 0) <= 0).map((entry) => `${entry?.invoiceNumber || "請求番号未設定"} / 売上ID:${entry?.id || "-"}`) },
    { key: "sales_overpaid", label: "入金額が請求額を超えている売上", severity: "critical", items: (Array.isArray(state.sales) ? state.sales : []).filter((entry) => Number(entry?.paidAmount || 0) > Number(entry?.invoiceAmount || 0)).map((entry) => `${entry?.invoiceNumber || "請求番号未設定"} / 請求:${formatCurrency(entry?.invoiceAmount || 0)} / 入金:${formatCurrency(entry?.paidAmount || 0)}`) },
    { key: "overdue_open_tasks", label: "期限切れ未完了タスク", severity: "critical", items: (Array.isArray(state.caseTasks) ? state.caseTasks : []).filter((entry) => entry?.status !== "完了" && Number.isFinite(toDateOnlyTimestamp(entry?.dueDate)) && toDateOnlyTimestamp(entry?.dueDate) < todayTs).map((entry) => `${entry?.taskTitle || "タスク名未設定"} / 期限:${entry?.dueDate || "-"} / ID:${entry?.id || "-"}`) },
    { key: "near_due_uncollected_documents", label: "案件期限7日以内で未回収書類がある案件", severity: "critical", items: getNearDueUncollectedDocumentItems(caseMap, clientMap, todayTs) },
    { key: "estimates_zero_amount", label: "見積金額0円の見積", severity: "warning", items: (Array.isArray(state.estimates) ? state.estimates : []).filter((entry) => Number(entry?.totalAmount || 0) <= 0).map((entry) => `${entry?.estimateNumber || "見積番号未設定"} / ${entry?.estimateTitle || "件名未設定"}`) },
    { key: "sales_marked_paid_but_unpaid", label: "未入金なのに入金済扱いになっている売上", severity: "critical", items: (Array.isArray(state.sales) ? state.sales : []).filter((entry) => paidStatuses.has(entry?.paymentStatus) && Number(entry?.paidAmount || 0) < Number(entry?.invoiceAmount || 0)).map((entry) => `${entry?.invoiceNumber || "請求番号未設定"} / 請求:${formatCurrency(entry?.invoiceAmount || 0)} / 入金:${formatCurrency(entry?.paidAmount || 0)}`) },
    { key: "payment_due_overdue_unpaid", label: "支払期限切れ未入金", severity: "critical", items: (Array.isArray(state.sales) ? state.sales : []).filter((entry) => isSaleOverdue(entry) && Number(getRemainingAmount(entry)) > 0).map((entry) => `${entry?.invoiceNumber || "請求番号未設定"} / 期限:${entry?.dueDate || "-"} / 残額:${formatCurrency(getRemainingAmount(entry))}`) },
    { key: "clients_without_history", label: "対応履歴のない顧客", severity: "info", items: getClientsWithoutHistoryItems(caseMap) },
    { key: "backup_old_or_missing", label: "最終バックアップ日が未設定または7日超過", severity: "info", items: getBackupStalenessItems() },
    { key: "preservation_requirements", label: "保存要件チェック", severity: "warning", items: getPreservationRequirementItems(caseMap) },
  ];
  return checks;
}

function getPreservationRequirementItems(caseMap) {
  const sales = Array.isArray(state.sales) ? state.sales : [];
  const payments = Array.isArray(state.payments) ? state.payments : [];
  const expenses = Array.isArray(state.expenses) ? state.expenses : [];
  const appSettings = getAppSettings();
  const hasInvoiceRegistration = Boolean(asTrimmedText(appSettings?.invoiceRegistrationNumber || ""));
  const paymentSaleIds = new Set(payments.map((entry) => entry?.saleId).filter(Boolean));

  const formatMissingList = (entries, fallbackLabel) => entries.slice(0, 5).join("、") + (entries.length > 5 ? ` ほか${entries.length - 5}件` : "") || fallbackLabel;
  const statusLabel = (level, message) => `${level}｜${message}`;
  const result = [];

  const noInvoiceNumber = sales.filter((entry) => !asTrimmedText(entry?.invoiceNumber)).map((entry) => `売上ID:${entry?.id || "-"}`);
  result.push(statusLabel(noInvoiceNumber.length ? "不足" : "OK", noInvoiceNumber.length ? `請求書: invoice_number 未設定（${formatMissingList(noInvoiceNumber, "対象なし")}）` : "請求書: invoice_number"));

  const noCustomerName = sales.filter((entry) => !asTrimmedText(entry?.customerName)).map((entry) => `売上ID:${entry?.id || "-"}`);
  result.push(statusLabel(noCustomerName.length ? "不足" : "OK", noCustomerName.length ? `請求書: customer_name 未設定（${formatMissingList(noCustomerName, "対象なし")}）` : "請求書: customer_name"));

  const noInvoiceAmount = sales.filter((entry) => Number(entry?.invoiceAmount || 0) <= 0).map((entry) => `売上ID:${entry?.id || "-"}`);
  result.push(statusLabel(noInvoiceAmount.length ? "不足" : "OK", noInvoiceAmount.length ? `請求書: invoice_amount 未設定/0（${formatMissingList(noInvoiceAmount, "対象なし")}）` : "請求書: invoice_amount"));

  const noInvoiceOutput = sales.filter((entry) => !entry?.caseId && !asTrimmedText(entry?.invoiceUrl)).map((entry) => `売上ID:${entry?.id || "-"}`);
  result.push(statusLabel(noInvoiceOutput.length ? "要確認" : "OK", noInvoiceOutput.length ? `請求書: invoice_url または請求書出力元案件が不足（${formatMissingList(noInvoiceOutput, "対象なし")}）` : "請求書: 出力手段"));
  result.push(statusLabel(hasInvoiceRegistration ? "OK" : "不足", hasInvoiceRegistration ? "請求書: インボイス登録番号" : "請求書: インボイス登録番号が未設定"));

  const paidSales = sales.filter((entry) => Number(entry?.paidAmount || 0) > 0);
  const paidSalesWithoutReceipt = paidSales.filter((entry) => !entry?.id && !entry?.caseId).map((entry) => `売上ID:${entry?.id || "-"}`);
  result.push(statusLabel(paidSalesWithoutReceipt.length ? "要確認" : "OK", paidSalesWithoutReceipt.length ? `領収書: paid_amount>0 の出力手段不足（${formatMissingList(paidSalesWithoutReceipt, "対象なし")}）` : "領収書: paid_amount>0 の出力手段"));

  const paymentHistoryMissing = paidSales.filter((entry) => paymentSaleIds.has(entry?.id) && !entry?.id).map((entry) => `売上ID:${entry?.id || "-"}`);
  result.push(statusLabel(paymentHistoryMissing.length ? "不足" : "OK", paymentHistoryMissing.length ? `領収書: payment履歴ありで個別領収書不可（${formatMissingList(paymentHistoryMissing, "対象なし")}）` : "領収書: payment履歴あり個別領収書"));

  const cumulativeMissing = paidSales.filter((entry) => !paymentSaleIds.has(entry?.id) && !entry?.id).map((entry) => `売上ID:${entry?.id || "-"}`);
  result.push(statusLabel(cumulativeMissing.length ? "不足" : "OK", cumulativeMissing.length ? `領収書: payment履歴なしで累計領収書不可（${formatMissingList(cumulativeMissing, "対象なし")}）` : "領収書: payment履歴なし累計領収書"));

  const expenseNoAmount = expenses.filter((entry) => Number(entry?.amount || 0) <= 0).map((entry) => `経費ID:${entry?.id || "-"}`);
  result.push(statusLabel(expenseNoAmount.length ? "不足" : "OK", expenseNoAmount.length ? `経費証憑: amount 未設定/0（${formatMissingList(expenseNoAmount, "対象なし")}）` : "経費証憑: amount"));

  const expenseNoContent = expenses.filter((entry) => !asTrimmedText(entry?.content)).map((entry) => `経費ID:${entry?.id || "-"}`);
  result.push(statusLabel(expenseNoContent.length ? "不足" : "OK", expenseNoContent.length ? `経費証憑: content 未設定（${formatMissingList(expenseNoContent, "対象なし")}）` : "経費証憑: content"));

  const expenseNoDate = expenses.filter((entry) => !entry?.date && !entry?.expenseDate).map((entry) => `経費ID:${entry?.id || "-"}`);
  result.push(statusLabel(expenseNoDate.length ? "不足" : "OK", expenseNoDate.length ? `経費証憑: date/expense_date 未設定（${formatMissingList(expenseNoDate, "対象なし")}）` : "経費証憑: date/expense_date"));

  const expenseNoReceiptUrl = expenses.filter((entry) => !asTrimmedText(entry?.receiptUrl)).map((entry) => `経費ID:${entry?.id || "-"}`);
  result.push(statusLabel(expenseNoReceiptUrl.length ? "要確認" : "OK", expenseNoReceiptUrl.length ? `経費証憑: receipt_url 空欄（${formatMissingList(expenseNoReceiptUrl, "対象なし")}）` : "経費証憑: receipt_url"));

  const salesSearchReady = sales.every((entry) => entry?.invoiceAmount != null && Boolean(asTrimmedText(entry?.customerName)) && Boolean(entry?.dueDate || entry?.paidDate || (entry?.caseId && caseMap.get(entry.caseId)?.receivedDate)));
  result.push(statusLabel(salesSearchReady ? "OK" : "要確認", salesSearchReady ? "検索性: 売上（日付・金額・顧客名）" : "検索性: 売上の検索キー不足データあり"));

  const expenseSearchReady = expenses.every((entry) => entry?.amount != null && Boolean(entry?.date || entry?.expenseDate) && Boolean(asTrimmedText(entry?.payee) || asTrimmedText(entry?.content)));
  result.push(statusLabel(expenseSearchReady ? "OK" : "要確認", expenseSearchReady ? "検索性: 経費（日付・金額・支払先/内容）" : "検索性: 経費の検索キー不足データあり"));

  return result;
}

function getNearDueUncollectedDocumentItems(caseMap, clientMap, todayTs) {
  const caseIds = new Set(
    (Array.isArray(state.caseDocuments) ? state.caseDocuments : [])
      .filter((entry) => entry?.status === "未回収" && entry?.caseId && caseMap.has(entry.caseId))
      .map((entry) => entry.caseId)
  );
  return Array.from(caseIds).map((caseId) => caseMap.get(caseId)).filter(Boolean).filter((entry) => {
    const dueTs = toDateOnlyTimestamp(entry?.dueDate);
    return Number.isFinite(dueTs) && dueTs <= todayTs + (7 * 86400000);
  }).map((entry) => {
    const clientName = entry?.clientId && clientMap.has(entry.clientId) ? clientMap.get(entry.clientId)?.name : (entry?.customerName || "顧客不明");
    return `${clientName || "顧客不明"} / ${entry?.caseName || "案件名未設定"} / 期限:${entry?.dueDate || "-"}`;
  });
}

function getClientsWithoutHistoryItems(caseMap) {
  return (Array.isArray(state.clients) ? state.clients : []).filter((client) => {
    const hasDailyReport = (Array.isArray(state.dailyReports) ? state.dailyReports : []).some((entry) => {
      if (entry?.clientId === client.id) return true;
      const linkedCase = entry?.caseId ? caseMap.get(entry.caseId) : null;
      return linkedCase?.clientId === client.id;
    });
    return !hasDailyReport;
  }).map((entry) => `${entry?.name || "顧客名未設定"} / 顧客ID:${entry?.id || "-"}`);
}

function getBackupStalenessItems() {
  const info = getLastBackupInfo();
  if (!info.hasValue) return ["最終バックアップ日が未設定です。"];
  if (info.isStale) return [`最終バックアップ日: ${info.label}（${info.daysAgo}日前）`];
  return [];
}

function renderBackupStatusForIntegrityCheck() {
  if (!dataIntegrityBackupStatus) return;
  const info = getLastBackupInfo();
  dataIntegrityBackupStatus.classList.toggle("is-warning", !info.hasValue || info.isStale);
  if (!info.hasValue) {
    dataIntegrityBackupStatus.textContent = "最終バックアップ日: 未設定（要バックアップ）";
    return;
  }
  const staleLabel = info.isStale ? "（7日超過・要確認）" : "";
  dataIntegrityBackupStatus.textContent = `最終バックアップ日: ${info.label}（${info.daysAgo}日前）${staleLabel}`;
}

function getLastBackupInfo() {
  const raw = safeGetLocalStorage(BACKUP_AT_STORAGE_KEY);
  if (!raw) return { hasValue: false, isStale: true, label: "未設定", daysAgo: null };
  const valueDate = new Date(raw);
  const time = valueDate.getTime();
  if (!Number.isFinite(time)) return { hasValue: false, isStale: true, label: "未設定", daysAgo: null };
  const today = new Date(toDateString(new Date()));
  const diff = Math.floor((today.getTime() - time) / 86400000);
  return {
    hasValue: true,
    isStale: diff >= 7,
    label: toDateString(valueDate),
    daysAgo: Math.max(0, diff),
  };
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
  if (["settings", "setting", "設定"].includes(key)) return "settings";
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
  renderPendingAlerts();
  renderDocumentAlerts();
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
  renderDailyReportAlerts();
  renderDataIntegrityCheck();
}

function renderPendingAlerts() {
  if (!pendingAlertCard || !pendingAlertSummary || !pendingAlertCounts || !pendingAlertBody || !pendingAlertEmpty || !pendingAlertListWrap) return;
  const todayTimestamp = getTodayTimestamp();
  const alerts = [];
  const countMap = new Map();
  const addAlert = (category, entry) => {
    alerts.push({ category, ...entry });
    countMap.set(category, (countMap.get(category) || 0) + 1);
  };
  const addCaseAlert = (category, entry, action) => addAlert(category, {
    targetName: `${entry.customerName || "顧客不明"} / ${entry.caseName || "案件なし"}`,
    dateLabel: entry.dueDate || entry.nextActionDate || "-",
    amountLabel: "-",
    status: entry.status || "-",
    nextAction: action,
    rowId: entry?.id,
    rowType: "case",
  });

  state.cases.forEach((entry) => {
    const dueTs = toDateOnlyTimestamp(entry?.dueDate);
    if (entry?.dueDate && Number.isFinite(dueTs) && dueTs < todayTimestamp && getStatusCategory(entry.status) !== "完了") {
      addCaseAlert("期限日を過ぎた案件", entry, "案件の期限・進捗を更新");
    }
    const nextTs = toDateOnlyTimestamp(entry?.nextActionDate);
    if (entry?.nextActionDate && Number.isFinite(nextTs) && nextTs < todayTimestamp && getStatusCategory(entry.status) !== "完了") {
      addCaseAlert("次回対応日を過ぎた案件", entry, entry.nextAction || "次回対応を実施");
    }
  });
  state.caseTasks.forEach((entry) => {
    if (entry?.status === "完了") return;
    const linkedCase = state.cases.find((row) => row.id === entry.caseId);
    addAlert("未完了タスク", {
      targetName: `${linkedCase?.customerName || "顧客不明"} / ${(entry.taskTitle || "タスク未設定")}`,
      dateLabel: entry?.dueDate || "-",
      amountLabel: "-",
      status: entry?.status || "未着手",
      nextAction: "タスクを完了または更新",
      rowId: entry?.id,
      rowType: "case-task",
    });
  });
  state.caseDocuments.forEach((entry) => {
    if (entry?.status !== "未回収") return;
    const linkedCase = state.cases.find((row) => row.id === entry.caseId);
    addAlert("未回収書類", {
      targetName: `${linkedCase?.customerName || "顧客不明"} / ${entry?.documentName || "書類名未設定"}`,
      dateLabel: entry?.dueDate || "-",
      amountLabel: "-",
      status: entry?.status || "-",
      nextAction: "回収状況を更新",
      rowId: entry?.id,
      rowType: "case-document",
    });
  });
  getBillingLeakCandidates().forEach((entry) => {
    addCaseAlert("未請求の案件", entry, "売上を登録");
  });
  state.sales.forEach((entry) => {
    const remaining = getRemainingAmount(entry);
    const dueTs = toDateOnlyTimestamp(entry?.dueDate);
    const name = `${entry?.customerName || "顧客不明"} / ${entry?.caseName || "案件なし"}`;
    const invoiceRegistrationStatus = getInvoiceRegistrationStatus(entry);
    const invoiceDeliveryStatus = getInvoiceDeliveryStatus(entry);
    const paymentStatus = entry?.paymentStatus || calculatePaymentStatus(entry?.invoiceAmount || 0, entry?.paidAmount || 0);
    if (invoiceRegistrationStatus === "未作成") {
      addAlert("未請求の売上", {
        targetName: name,
        dateLabel: entry?.dueDate || "-",
        amountLabel: formatCurrency(entry?.invoiceAmount || 0),
        status: invoiceRegistrationStatus,
        nextAction: "請求情報を登録",
        rowId: entry?.id,
        rowType: "sale-uninvoiced",
      });
    }
    if (invoiceDeliveryStatus === "未送付") {
      addAlert("未送付の売上", {
        targetName: name,
        dateLabel: entry?.dueDate || "-",
        amountLabel: formatCurrency(entry?.invoiceAmount || 0),
        status: invoiceDeliveryStatus,
        nextAction: "請求書URLを登録",
        rowId: entry?.id,
        rowType: "sale-undelivered",
      });
    }
    if (paymentStatus === "未入金") {
      addAlert("未入金の売上", {
        targetName: name,
        dateLabel: entry?.dueDate || "-",
        amountLabel: formatCurrency(remaining),
        status: paymentStatus,
        nextAction: "入金登録または督促",
        rowId: entry?.id,
        rowType: "sale-unpaid-only",
      });
    }
    if (["未入金", "一部入金"].includes(entry?.paymentStatus) && remaining > 0) {
      addAlert("未入金・一部入金の売上", {
        targetName: name,
        dateLabel: entry?.dueDate || "-",
        amountLabel: formatCurrency(remaining),
        status: entry?.paymentStatus || "-",
        nextAction: "入金登録または督促",
        rowId: entry?.id,
        rowType: "sale-unpaid",
      });
    }
    if (entry?.dueDate && Number.isFinite(dueTs) && dueTs < todayTimestamp && remaining > 0) {
      addAlert("支払期限を過ぎた売上", {
        targetName: name,
        dateLabel: entry?.dueDate,
        amountLabel: formatCurrency(remaining),
        status: entry?.paymentStatus || "-",
        nextAction: "期限超過分を督促",
      });
    }
  });
  state.expenses.forEach((entry) => {
    if (String(entry?.receiptUrl || "").trim()) return;
    addAlert("receipt_url が空の経費", {
      targetName: `${entry?.category || "分類未設定"} / ${entry?.description || "内容未設定"}`,
      dateLabel: entry?.date || "-",
      amountLabel: formatCurrency(entry?.amount || 0),
      status: "領収書未登録",
      nextAction: "領収書URLを登録",
      rowId: entry?.id,
      rowType: "expense-no-receipt",
    });
  });
  state.pendingAlertSelections = {};

  pendingAlertCard.classList.toggle("has-alert", alerts.length > 0);
  pendingAlertSummary.textContent = `対象 ${alerts.length}件`;
  pendingAlertCounts.innerHTML = Array.from(countMap.entries()).map(([category, count]) => `<span class="pending-alert-chip">${escapeHtml(category)}: ${count}件</span>`).join("");
  pendingAlertBody.innerHTML = "";
  pendingAlertEmpty.hidden = alerts.length > 0;
  pendingAlertListWrap.hidden = alerts.length === 0;
  alerts.forEach((entry) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td><input type="checkbox" class="pending-alert-checkbox" data-pending-alert-id="${escapeHtml(`${entry.category}:${entry.rowType || "case"}:${entry.rowId || entry.targetName}`)}"></td>
      <td>${escapeHtml(entry.category)}</td>
      <td>${escapeHtml(entry.targetName)}</td>
      <td>${escapeHtml(entry.dateLabel ? formatDate(entry.dateLabel) : "-")}</td>
      <td>${escapeHtml(entry.amountLabel || "-")}</td>
      <td>${escapeHtml(entry.status || "-")}</td>
      <td>${escapeHtml(entry.nextAction || "-")}</td>
    `;
    const checkbox = tr.querySelector(".pending-alert-checkbox");
    if (checkbox) {
      const key = checkbox.dataset.pendingAlertId;
      if (key) {
        checkbox.dataset.rowType = entry.rowType || "";
        checkbox.dataset.rowId = entry.rowId || "";
        checkbox.dataset.category = entry.category || "";
        checkbox.addEventListener("change", () => {
          state.pendingAlertSelections[key] = checkbox.checked ? {
            rowType: checkbox.dataset.rowType || "",
            rowId: checkbox.dataset.rowId || "",
            category: checkbox.dataset.category || "",
          } : null;
          if (!checkbox.checked) delete state.pendingAlertSelections[key];
        });
      }
    }
    pendingAlertBody.appendChild(tr);
  });
}

async function handlePendingAlertBulkAction(action) {
  const selected = Object.values(state.pendingAlertSelections || {}).filter(Boolean);
  if (action === "open_sales") {
    subtabState.sales = "list";
    activateTab("sales");
    showAppMessage("売上一覧へ移動しました。未入金/一部入金の確認を行ってください。", false);
    return;
  }
  if (action === "open_expenses") {
    subtabState.expenses = "list";
    activateTab("expenses");
    showAppMessage("経費一覧へ移動しました。証憑URL未登録の確認を行ってください。", false);
    return;
  }
  if (!currentUser || !selected.length) {
    showAppMessage("一括更新対象を選択してください。", true);
    return;
  }
  try {
    if (action === "clear_next_action_date" || action === "complete_cases" || action === "mark_invoiced_cases") {
      const caseIds = [...new Set(selected.filter((entry) => entry.category === "次回対応日を過ぎた案件" || entry.category === "未請求の案件").map((entry) => entry.rowId).filter(Boolean))];
      if (!caseIds.length) return showAppMessage("対象分類の案件が選択されていません。", true);
      const payload = action === "clear_next_action_date" ? { next_action_date: null } : action === "complete_cases" ? { status: "完了" } : null;
      if (action === "mark_invoiced_cases") {
        const targets = state.cases.filter((entry) => caseIds.includes(entry.id) && asTrimmedText(entry.invoiceUrl));
        if (!targets.length) return showAppMessage("invoice_url がある未請求案件が選択されていません。", true);
        await Promise.all(targets.map((entry) => sbClient.from("cases").update({ status: "完了" }).eq("id", entry.id).eq("user_id", currentUser.id)));
      } else {
        await sbClient.from("cases").update(payload).in("id", caseIds).eq("user_id", currentUser.id);
      }
    } else if (action === "complete_tasks") {
      const ids = [...new Set(selected.filter((entry) => entry.rowType === "case-task").map((entry) => entry.rowId).filter(Boolean))];
      if (!ids.length) return showAppMessage("未完了タスクが選択されていません。", true);
      await sbClient.from("case_tasks").update({ status: "完了", completed_at: toDateString(new Date()) }).in("id", ids).eq("user_id", currentUser.id);
    } else if (action === "complete_documents") {
      const ids = [...new Set(selected.filter((entry) => entry.rowType === "case-document").map((entry) => entry.rowId).filter(Boolean))];
      if (!ids.length) return showAppMessage("未回収書類が選択されていません。", true);
      await sbClient.from("case_documents").update({ status: "回収済" }).in("id", ids).eq("user_id", currentUser.id);
    }
    await loadAllData();
    showAppMessage("一括処理を実行しました。", false);
  } catch (error) {
    showAppMessage(`一括処理に失敗しました。${formatSupabaseError(error)}`, true);
  }
}

function renderTodayTaskCard() {
  renderTodayTasks();
}

function buildIntegratedAlerts() {
  const todayTs = getTodayTimestamp();
  const threeDaysLater = todayTs + (3 * 86400000);
  const priorityOrder = { high: 0, medium: 1, low: 2 };
  const alerts = [];
  const caseMap = new Map((Array.isArray(state.cases) ? state.cases : []).map((entry) => [entry.id, entry]));
  const clientMap = new Map((Array.isArray(state.clients) ? state.clients : []).map((entry) => [entry.id, entry]));

  (Array.isArray(state.caseTasks) ? state.caseTasks : []).forEach((task) => {
    const dueTs = toDateOnlyTimestamp(task?.dueDate);
    if (task?.status === "完了" || !Number.isFinite(dueTs) || dueTs > todayTs) return;
    const linkedCase = task?.caseId ? caseMap.get(task.caseId) : null;
    const linkedClient = linkedCase?.clientId ? clientMap.get(linkedCase.clientId) : null;
    alerts.push({
      id: task?.id,
      type: "task",
      typeLabel: "案件タスク",
      priority: "high",
      title: task?.taskTitle || "未設定",
      clientName: linkedClient?.name || linkedCase?.customerName || "顧客不明",
      caseName: linkedCase?.caseName || "案件なし",
      dueDate: task?.dueDate || "",
      dateSort: dueTs,
      dateLabel: task?.dueDate ? formatDate(task.dueDate) : "未設定",
      subInfo: dueTs < todayTs ? "期限切れ" : "本日期限",
      target: "case-task",
      actionButtonHtml: `<button type="button" class="secondary-btn" data-action="complete_case_task" data-list-action="complete_case_task" data-task-target="case-task-complete" data-task-id="${task?.id || ""}">完了</button>`,
    });
  });

  (Array.isArray(state.sales) ? state.sales : []).forEach((sale) => {
    const dueTs = toDateOnlyTimestamp(sale?.dueDate);
    const isUnpaid = sale?.paymentStatus !== "入金済";
    if (!isUnpaid || !Number.isFinite(dueTs)) return;
    const linkedCase = sale?.caseId ? caseMap.get(sale.caseId) : null;
    const linkedClient = linkedCase?.clientId ? clientMap.get(linkedCase.clientId) : null;
    if (dueTs <= todayTs) {
      alerts.push({
        id: sale?.id,
        type: "payment",
        typeLabel: "未入金",
        priority: "high",
        title: "未入金",
        amount: getRemainingAmount(sale),
        clientName: linkedClient?.name || linkedCase?.customerName || "顧客不明",
        caseName: linkedCase?.caseName || "案件なし",
        dueDate: sale?.dueDate || "",
        dateSort: dueTs,
        dateLabel: sale?.dueDate ? formatDate(sale.dueDate) : "未設定",
        subInfo: `残額: ${formatCurrency(getRemainingAmount(sale))}`,
        target: "sale",
        actionButtonHtml: isReminderRecordableSale(sale) ? `<button type="button" class="secondary-btn record-reminder-btn" data-action="record_reminder" data-list-action="record_reminder" data-sale-id="${sale?.id || ""}">督促記録</button>` : "-",
      });
    }
    const reminderTs = toDateOnlyTimestamp(sale?.lastReminderDate);
    const reminderElapsed = Number.isFinite(reminderTs) && ((todayTs - reminderTs) / 86400000 >= 7);
    if (reminderElapsed) {
      alerts.push({
        id: sale?.id,
        type: "reminder",
        typeLabel: "督促",
        priority: "medium",
        title: "督促必要",
        clientName: linkedClient?.name || linkedCase?.customerName || "顧客不明",
        caseName: linkedCase?.caseName || "案件なし",
        lastReminderDate: sale?.lastReminderDate || "",
        dateSort: reminderTs,
        dateLabel: sale?.lastReminderDate ? `最終督促: ${formatDate(sale.lastReminderDate)}` : "最終督促日なし",
        subInfo: `残額: ${formatCurrency(getRemainingAmount(sale))}`,
        target: "sale",
        actionButtonHtml: isReminderRecordableSale(sale) ? `<button type="button" class="secondary-btn record-reminder-btn" data-action="record_reminder" data-list-action="record_reminder" data-sale-id="${sale?.id || ""}">督促記録</button>` : "-",
      });
    }
  });

  (Array.isArray(state.cases) ? state.cases : []).forEach((caseRow) => {
    if (getStatusCategory(caseRow?.status) === "完了") return;
    const dueTs = toDateOnlyTimestamp(caseRow?.dueDate);
    if (!Number.isFinite(dueTs) || dueTs > threeDaysLater) return;
    const linkedClient = caseRow?.clientId ? clientMap.get(caseRow.clientId) : null;
    alerts.push({
      id: caseRow?.id,
      type: "deadline",
      typeLabel: "期限",
      priority: "medium",
      title: "期限間近",
      clientName: linkedClient?.name || caseRow?.customerName || "顧客不明",
      caseName: caseRow?.caseName || "案件なし",
      dueDate: caseRow?.dueDate || "",
      dateSort: dueTs,
      dateLabel: caseRow?.dueDate ? formatDate(caseRow.dueDate) : "未設定",
      subInfo: formatRemainingDays(Math.floor((dueTs - todayTs) / 86400000)),
      target: "case",
    });
  });

  (Array.isArray(state.caseDocuments) ? state.caseDocuments : []).forEach((doc) => {
    const linkedCase = doc?.caseId ? caseMap.get(doc.caseId) : null;
    const linkedClient = linkedCase?.clientId ? clientMap.get(linkedCase.clientId) : null;
    if (doc?.status === "不備あり") {
      alerts.push({
        id: doc?.id,
        type: "document",
        typeLabel: "書類不備",
        priority: "high",
        title: doc?.documentName || "書類名未設定",
        clientName: linkedClient?.name || linkedCase?.customerName || "顧客不明",
        caseName: linkedCase?.caseName || "案件なし",
        dateSort: todayTs,
        dateLabel: "要対応",
        subInfo: "不備あり",
        target: "case-document",
      });
    } else if (doc?.status === "未回収") {
      alerts.push({
        id: doc?.id,
        type: "document",
        typeLabel: "未回収書類",
        priority: "medium",
        title: doc?.documentName || "書類名未設定",
        clientName: linkedClient?.name || linkedCase?.customerName || "顧客不明",
        caseName: linkedCase?.caseName || "案件なし",
        dateSort: todayTs,
        dateLabel: "未回収",
        subInfo: "回収待ち",
        target: "case-document",
      });
    }
  });

  return alerts
    .slice()
    .sort((a, b) => {
      const priorityDiff = (priorityOrder[a.priority] ?? 99) - (priorityOrder[b.priority] ?? 99);
      if (priorityDiff !== 0) return priorityDiff;
      return (a.dateSort || 0) - (b.dateSort || 0);
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
        taskTitle: "未入金督促",
        date: sale.dueDate || sale.paidDate || toDateString(sale.createdAt),
        urgencyClass,
        dueDateLabel: sale.dueDate ? formatDate(sale.dueDate) : "未設定",
        subInfo: `残額: ${formatCurrency(getRemainingAmount(sale))} / 最終督促: ${sale.lastReminderDate ? formatDate(sale.lastReminderDate) : "未記録"} / ${Number(sale.reminderCount || 0)}回`,
        showReminderButton: isReminderRecordableSale(sale),
      });
    });
  state.caseTasks.forEach((entry) => {
    const dueTs = toDateOnlyTimestamp(entry.dueDate);
    if (!entry.dueDate || entry.status === "完了" || !Number.isFinite(dueTs) || dueTs > todayLimit) return;
    const linkedCase = state.cases.find((row) => row.id === entry.caseId);
    tasks.push({
      id: entry.id,
      target: "case-task",
      type: "案件タスク",
      customerName: linkedCase?.customerName || "顧客不明",
      caseName: linkedCase?.caseName || "案件なし",
      taskTitle: entry.taskTitle || "未設定",
      date: entry.dueDate,
      dueDateLabel: formatDate(entry.dueDate),
      subInfo: entry.status,
      urgencyClass: getCaseTaskUrgencyClass(entry),
      actionButtonHtml: `<button type="button" class="secondary-btn" data-action="complete_case_task" data-list-action="complete_case_task" data-task-target="case-task-complete" data-task-id="${entry.id}">完了</button>`,
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
  const counts = { overdue: 0, within3: 0, within7: 0, within30: 0 };
  targets.forEach((item) => {
    counts[item.deadlineStatus] += 1;
  });

  deadlineAlertCard.classList.toggle("has-alert", targets.length > 0);
  deadlineAlertSummary.innerHTML = "";

  [
    { key: "overdue", label: "期限切れ", count: counts.overdue },
    { key: "within3", label: "3日以内", count: counts.within3 },
    { key: "within7", label: "7日以内", count: counts.within7 },
    { key: "within30", label: "30日以内", count: counts.within30 },
  ].forEach((entry) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = `deadline-alert-chip ${entry.key}`;
    button.dataset.deadlineFilter = entry.key;
    button.dataset.action = "deadline_alert_click";
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
        entry.deadlineStatus === "within3" ? "deadline-within3" :
        entry.deadlineStatus === "within7" ? "deadline-within7" :
        "deadline-within30";
      tr.innerHTML = `
        <td>${escapeHtml(entry.customerName)}</td>
        <td>${escapeHtml(entry.caseName)}</td>
        <td>${formatDate(entry.dueDate)}</td>
        <td>${escapeHtml(entry.label || entry.status)}</td>
        <td>${formatRemainingDays(entry.remainingDays)}</td>
        <td>${entry.alertType === "case-task"
    ? `<button type="button" class="secondary-btn edit-case-task-btn" data-action="edit_case_task" data-list-action="edit_case_task" data-task-id="${entry.id}">編集</button>`
    : `<button type="button" class="secondary-btn" data-action="edit_case" data-list-action="edit_case" data-task-target="case" data-task-id="${entry.id}" data-case-id="${entry.id}">編集</button>`}</td>
      `;
      deadlineAlertBody.appendChild(tr);
    });
}

function getDeadlineAlertTargets() {
  const caseTargets = state.cases
    .map((entry) => {
      const info = getCaseDeadlineInfo(entry);
      if (!info) return null;
      return { ...entry, ...info, alertType: "case", label: entry.status };
    })
    .filter(Boolean);
  const taskTargets = state.caseTasks
    .map((entry) => {
      const info = getCaseTaskDeadlineInfo(entry);
      if (!info) return null;
      const linkedCase = state.cases.find((row) => row.id === entry.caseId);
      return {
        id: entry.id,
        customerName: linkedCase?.customerName || "顧客不明",
        caseName: linkedCase?.caseName || "案件なし",
        dueDate: entry.dueDate,
        status: entry.status,
        deadlineStatus: info.deadlineStatus,
        remainingDays: info.remainingDays,
        alertType: "case-task",
        label: `案件タスク: ${entry.taskTitle || "未設定"}`,
      };
    })
    .filter(Boolean);
  return [...caseTargets, ...taskTargets];
}

function getCaseDeadlineInfo(entry) {
  if (!entry?.dueDate) return null;
  if (getStatusCategory(entry.status) === "完了") return null;
  const dueTimestamp = toDateOnlyTimestamp(entry.dueDate);
  const todayTimestamp = getTodayTimestamp();
  if (!Number.isFinite(dueTimestamp) || !Number.isFinite(todayTimestamp)) return null;
  const remainingDays = Math.floor((dueTimestamp - todayTimestamp) / 86400000);
  if (remainingDays < 0) return { deadlineStatus: "overdue", remainingDays };
  if (remainingDays <= 3) return { deadlineStatus: "within3", remainingDays };
  if (remainingDays <= 7) return { deadlineStatus: "within7", remainingDays };
  if (remainingDays <= 30) return { deadlineStatus: "within30", remainingDays };
  return null;
}

function getCaseTaskDeadlineInfo(entry) {
  if (!entry?.dueDate) return null;
  if (entry.status === "完了") return null;
  const dueTimestamp = toDateOnlyTimestamp(entry.dueDate);
  const todayTimestamp = getTodayTimestamp();
  if (!Number.isFinite(dueTimestamp) || !Number.isFinite(todayTimestamp)) return null;
  const remainingDays = Math.floor((dueTimestamp - todayTimestamp) / 86400000);
  if (remainingDays < 0) return { deadlineStatus: "overdue", remainingDays };
  if (remainingDays <= 3) return { deadlineStatus: "within3", remainingDays };
  if (remainingDays <= 7) return { deadlineStatus: "within7", remainingDays };
  return null;
}

function getCaseTaskUrgencyClass(entry) {
  const info = getCaseTaskDeadlineInfo(entry);
  if (!info) return "";
  if (info.deadlineStatus === "overdue") return "task-overdue";
  if (info.deadlineStatus === "within3") return "task-soon";
  return "task-upcoming";
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
        <td>${entry.sourceType === "daily-report"
        ? `<button type="button" class="secondary-btn" data-action="edit_daily_report" data-list-action="edit_daily_report" data-daily-report-id="${entry.id}">編集</button>`
        : `<button type="button" class="secondary-btn" data-action="edit_case" data-list-action="edit_case" data-task-target="case" data-task-id="${entry.id}" data-case-id="${entry.id}">編集</button>`}</td>
      `;
      nextActionAlertBody.appendChild(tr);
    });
}

function getNextActionAlertTargets() {
  const caseTargets = state.cases
    .map((entry) => {
      const info = getNextActionInfo(entry);
      if (!info) return null;
      return { ...entry, ...info, sourceType: "case" };
    })
    .filter((entry) => entry && entry.remainingDays <= 0 && getStatusCategory(entry.status) !== "完了");

  const dailyReportTargets = state.dailyReports
    .map((entry) => {
      const info = getNextActionInfo(entry);
      if (!info) return null;
      const hasContext = String(entry?.nextAction || "").trim()
        || String(entry?.workContent || "").trim()
        || String(entry?.memo || "").trim();
      if (!hasContext || info.remainingDays >= 0) return null;
      const linkedCase = state.cases.find((caseEntry) => caseEntry.id === entry.caseId);
      return {
        ...entry,
        ...info,
        sourceType: "daily-report",
        customerName: linkedCase?.customerName || "日報",
        caseName: linkedCase?.caseName || "日報",
        status: "対応期限超過",
      };
    })
    .filter(Boolean);

  return [...caseTargets, ...dailyReportTargets];
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


function renderDailyReportAlerts() {
  if (!dailyReportAlertBody || !dailyReportAlertEmpty || !dailyReportAlertListWrap) return;
  const targets = getNextActionAlertTargets();
  dailyReportAlertBody.innerHTML = "";
  dailyReportAlertEmpty.hidden = targets.length > 0;
  dailyReportAlertListWrap.hidden = targets.length === 0;
  targets
    .slice()
    .sort((a, b) => (a.remainingDays || 0) - (b.remainingDays || 0))
    .forEach((entry) => {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>次回対応</td>
        <td>${escapeHtml(entry.sourceType === "daily-report" ? (entry.caseName || "日報") : `${entry.customerName || "顧客不明"} / ${entry.caseName || "案件なし"}`)}</td>
        <td>${formatDate(entry.nextActionDate)}</td>
        <td>${escapeHtml(entry.sourceType === "daily-report" ? "対応期限超過" : "次回対応期限超過")}</td>
        <td>${escapeHtml(entry.nextAction || "-")}</td>
      `;
      dailyReportAlertBody.appendChild(tr);
    });
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
    button.dataset.action = "status_summary_filter";
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
      const paymentHistory = renderPaymentHistoryInline(sale.id);
      tr.innerHTML = `
        <td>${escapeHtml(customerName)}</td>
        <td>${escapeHtml(caseName)}</td>
        <td>${formatCurrency(sale.invoiceAmount)}</td>
        <td>${formatCurrency(sale.paidAmount)}</td>
        <td>${formatCurrency(getRemainingAmount(sale))}</td>
        <td>${formatDate(sale.dueDate)}</td>
        <td><span class="${safeClassName(getSaleStatusClass(sale))}">${escapeHtml(sale.paymentStatus)}</span></td>
        <td>${getSaleDeadlineDaysLabel(sale)}</td>
        <td>${Number(sale.reminderCount || 0)}回</td>
        <td>${sale.lastReminderDate ? formatDate(sale.lastReminderDate) : "-"}</td>
        <td>${escapeHtml(sale.reminderMethod || "-")}</td>
        <td>${escapeHtml(sale.reminderMemo || "-")}</td>
        <td>${paymentHistory}${renderSaleCumulativeReceiptButton(sale)}</td>
        <td>
          <button type="button" class="secondary-btn edit-sale-btn" data-action="edit_sale" data-list-action="edit_sale" data-sale-id="${sale.id}">編集</button>
          <button type="button" class="secondary-btn register-payment-btn record-payment-btn" data-action="record_payment" data-list-action="record_payment" data-sale-id="${sale.id}">入金登録</button>
          ${isReminderRecordableSale(sale) ? `<button type="button" class="secondary-btn record-reminder-btn" data-action="record_reminder" data-list-action="record_reminder" data-sale-id="${sale.id}">督促記録</button>` : ""}
        </td>
      `;
      safeAddClass(tr, getSaleRowClass(sale));
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
    const paymentHistory = renderPaymentHistoryInline(sale.id);
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(linkedCase?.customerName || "（削除済み顧客）")}</td>
      <td>${escapeHtml(linkedCase?.caseName || "（削除済み案件）")}</td>
      <td>${formatCurrency(sale.invoiceAmount)}</td>
      <td>${formatCurrency(sale.paidAmount)}</td>
      <td>${formatCurrency(getRemainingAmount(sale))}</td>
      <td>${formatDate(sale.dueDate)}</td>
      <td>${formatDate(sale.paidDate)}</td>
      <td><span class="${safeClassName(getSaleStatusClass(sale))}">${escapeHtml(sale.paymentStatus)}</span></td>
      <td>${getSaleDeadlineDaysLabel(sale)}</td>
      <td>${Number(sale.reminderCount || 0)}回</td>
      <td>${sale.lastReminderDate ? formatDate(sale.lastReminderDate) : "-"}</td>
      <td>${escapeHtml(sale.reminderMethod || "-")}</td>
      <td>${escapeHtml(sale.reminderMemo || "-")}</td>
      <td>${paymentHistory}${renderSaleCumulativeReceiptButton(sale)}</td>
      <td><button type="button" class="secondary-btn edit-sale-btn" data-action="edit_sale" data-list-action="edit_sale" data-sale-id="${sale.id}">編集</button> <button type="button" class="secondary-btn register-payment-btn record-payment-btn" data-action="record_payment" data-list-action="record_payment" data-sale-id="${sale.id}">入金登録</button> ${isReminderRecordableSale(sale) ? `<button type="button" class="secondary-btn record-reminder-btn" data-action="record_reminder" data-list-action="record_reminder" data-sale-id="${sale.id}">督促記録</button>` : ""}</td>
    `;
    safeAddClass(tr, getSaleRowClass(sale));
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
      <td class="${safeClassName(row.profit < 0 ? "loss-text" : "")}">${formatCurrency(row.profit)}</td>
      <td class="${safeClassName(row.unpaidTotal > 0 ? "analysis-unpaid-text" : "")}">${formatCurrency(row.unpaidTotal)}${row.unpaidTotal > 0 ? " ⚠️" : ""}</td>
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
      <td class="${safeClassName(row.profit < 0 ? "loss-text" : "")}">${formatCurrency(row.profit)}</td>
      <td class="${safeClassName(row.unpaidTotal > 0 ? "analysis-unpaid-text" : "")}">${formatCurrency(row.unpaidTotal)}${row.unpaidTotal > 0 ? " ⚠️" : ""}</td>
      <td>${formatCurrency(row.averagePrice)}</td>
      <td class="${safeClassName(marginClass)}">${row.profitMargin.toFixed(1)}%</td>
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
      <td class="${safeClassName(row.profit < 0 ? "loss-text" : "")}">${formatCurrency(row.profit)}</td>
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
      <td><button type="button" class="secondary-btn register-sale-btn" data-action="register_sale" data-list-action="register_sale" data-case-id="${entry.id}">売上登録</button></td>
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

function renderPaymentHistoryInline(saleId) {
  const payments = getSalePayments(saleId);
  if (!payments.length) return "-";
  return payments.map((entry) => (
    `<div class="payment-history-row">${formatDate(entry.paymentDate)} / ${formatCurrency(entry.amount)} / ${escapeHtml(entry.method || "その他")} ${entry.memo ? `/ ${escapeHtml(entry.memo)}` : ""} <button type="button" class="danger-btn delete-payment-btn" data-action="delete_payment" data-list-action="delete_payment" data-sale-id="${saleId}" data-payment-id="${entry.id}">削除</button> <button type="button" class="secondary-btn" data-action="print_receipt" data-list-action="print_receipt" data-sale-id="${saleId}" data-payment-id="${entry.id}">領収書</button></div>`
  )).join("");
}

function renderSaleCumulativeReceiptButton(sale) {
  if (Number(sale?.paidAmount || 0) <= 0) return "";
  return ' <button type="button" class="secondary-btn" data-action="print_cumulative_receipt" data-list-action="print_cumulative_receipt" data-sale-id="' + escapeHtml(sale.id) + '">累計領収書</button>';
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
  if (caseTaskCaseSelect) {
    const currentValue = caseTaskCaseSelect.value;
    caseTaskCaseSelect.innerHTML = `<option value="">案件なし</option>${options}`;
    if (currentValue) caseTaskCaseSelect.value = currentValue;
  }
  if (caseDocumentCaseSelect) {
    const currentValue = caseDocumentCaseSelect.value;
    caseDocumentCaseSelect.innerHTML = `<option value="">案件なし</option>${options}`;
    if (currentValue) caseDocumentCaseSelect.value = currentValue;
  }
  if (permitCaseSelect) {
    const currentValue = permitCaseSelect.value;
    permitCaseSelect.innerHTML = `<option value="">案件を選択してください</option>${options}`;
    if (currentValue) permitCaseSelect.value = currentValue;
  }
}

function renderCaseTaskOptions() {
  if (!caseTaskCaseSelect) return;
  const options = state.cases
    .slice()
    .sort((a, b) => a.createdAt - b.createdAt)
    .map((c) => `<option value="${c.id}">${escapeHtml(c.customerName)}｜${escapeHtml(c.caseName)}</option>`)
    .join("");
  const currentValue = caseTaskCaseSelect.value;
  caseTaskCaseSelect.innerHTML = `<option value="">案件なし</option>${options}`;
  if (currentValue) caseTaskCaseSelect.value = currentValue;
}

function renderCaseTasksTable() {
  if (!caseTasksBody || !caseTasksEmpty || !caseTasksListWrap) return;
  caseTasksBody.innerHTML = "";
  const rows = state.caseTasks.slice().sort((a, b) => toSortTimestamp(a.dueDate) - toSortTimestamp(b.dueDate));
  caseTasksEmpty.hidden = rows.length > 0;
  caseTasksListWrap.hidden = rows.length === 0;
  rows.forEach((entry) => {
    const linkedCase = state.cases.find((row) => row.id === entry.caseId);
    const customerName = linkedCase?.customerName || "顧客不明";
    const caseName = linkedCase?.caseName || "案件なし";
    const tr = document.createElement("tr");
    tr.className = getCaseTaskUrgencyClass(entry);
    tr.dataset.id = entry.id;
    tr.innerHTML = `
      <td>${escapeHtml(customerName)}</td>
      <td>${escapeHtml(caseName)}</td>
      <td>${escapeHtml(entry.taskTitle || "未設定")}</td>
      <td>${formatDate(entry.dueDate)}</td>
      <td>${escapeHtml(entry.status)}</td>
      <td>${formatDate(entry.completedAt)}</td>
      <td><button type="button" class="secondary-btn edit-case-task-btn" data-action="edit_case_task" data-list-action="edit_case_task" data-task-id="${entry.id}">編集</button></td>
      <td>${entry.status !== "完了" ? `<button type="button" class="secondary-btn complete-case-task-btn" data-action="complete_case_task" data-list-action="complete_case_task" data-task-id="${entry.id}">完了</button>` : "-"}</td>
      <td><button type="button" class="danger-btn delete-case-task-btn" data-action="delete_case_task" data-list-action="delete_case_task" data-task-id="${entry.id}">削除</button></td>
    `;
    caseTasksBody.appendChild(tr);
  });
}

function renderCaseDocuments() {
  if (!caseDocumentsBody || !caseDocumentsEmpty || !caseDocumentsListWrap) return;
  caseDocumentsBody.innerHTML = "";
  const rows = (Array.isArray(state.caseDocuments) ? state.caseDocuments : [])
    .slice()
    .sort((a, b) => toSortTimestamp(a.receivedDate || a.createdAt) - toSortTimestamp(b.receivedDate || b.createdAt));
  caseDocumentsEmpty.hidden = rows.length > 0;
  caseDocumentsListWrap.hidden = rows.length === 0;
  rows.forEach((entry) => {
    const linkedCase = state.cases.find((row) => row.id === entry.caseId);
    const customerName = linkedCase?.customerName || "顧客不明";
    const caseName = linkedCase?.caseName || "案件なし";
    const tr = document.createElement("tr");
    tr.dataset.id = entry.id;
    tr.innerHTML = `
      <td>${escapeHtml(customerName)}</td>
      <td>${escapeHtml(caseName)}</td>
      <td>${escapeHtml(entry.documentName || "未設定")}</td>
      <td>${escapeHtml(entry.status || "未回収")}</td>
      <td>${formatDate(entry.receivedDate)}</td>
      <td>${formatDate(entry.checkedDate)}</td>
      <td>${escapeHtml(truncateText(entry.memo || "", 80) || "-")}</td>
      <td>${entry.fileUrl ? `<a href="${escapeHtml(entry.fileUrl)}" target="_blank" rel="noopener noreferrer">開く</a>` : "-"}</td>
      <td><button type="button" class="secondary-btn edit-btn" data-action="edit_case_document" data-list-action="edit_case_document" data-case-document-id="${entry.id}">編集</button></td>
      <td><button type="button" class="danger-btn delete-btn" data-action="delete_case_document" data-list-action="delete_case_document" data-case-document-id="${entry.id}">削除</button></td>
    `;
    caseDocumentsBody.appendChild(tr);
  });
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
      <div class="row-actions"><button type="button" class="secondary-btn edit-client-btn" data-action="edit_client" data-list-action="edit_client">編集</button><button type="button" class="danger-btn delete-client-btn" data-action="delete_client" data-list-action="delete_client">削除</button></div>
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
    if (metas[0]) metas[0].textContent = `標準見積: ${Number.isFinite(entry.standardEstimateAmount) && entry.standardEstimateAmount > 0 ? formatCurrency(entry.standardEstimateAmount) : "未設定"} / 標準期限: ${Number.isFinite(entry.defaultDueDays) ? `${entry.defaultDueDays}日後` : "未設定"} / 必要書類: ${truncateText(entry.requiredDocuments || "未設定", 80)}`;
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
          <button type="button" class="secondary-btn edit-daily-report-btn" data-action="edit_daily_report" data-list-action="edit_daily_report">編集</button>
          <button type="button" class="danger-btn delete-daily-report-btn" data-action="delete_daily_report" data-list-action="delete_daily_report">削除</button>
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
    <button type="button" class="danger-btn estimate-item-remove-btn" data-action="remove_estimate_item_row">削除</button>
  `;
  estimateItemsWrap.appendChild(row);
  hydrateActionButtons(row);
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

function getAppSettings() {
  return { ...DEFAULT_APP_SETTINGS, ...(state.appSettings || {}) };
}

function normalizeTaxRate(value) {
  const parsed = Number(value);
  if (!Number.isFinite(parsed) || parsed < 0) return DEFAULT_APP_SETTINGS.taxRate;
  return parsed <= 1 ? parsed : parsed / 100;
}

function getPositiveInt(value, fallback) {
  const parsed = Number(value);
  if (!Number.isFinite(parsed) || parsed <= 0) return fallback;
  return Math.floor(parsed);
}

function getCurrentTaxRate() {
  return normalizeTaxRate(getAppSettings().taxRate);
}

function getDefaultInvoiceDueDays() {
  const parsed = Number(getAppSettings().defaultInvoiceDueDays);
  if (!Number.isFinite(parsed) || parsed <= 0) return DEFAULT_APP_SETTINGS.defaultInvoiceDueDays;
  return Math.floor(parsed);
}

function formatTaxRatePercent(taxRate = getCurrentTaxRate()) {
  return `${(normalizeTaxRate(taxRate) * 100).toFixed(1).replace(/\.0$/, "")}%`;
}

function renderSettingsForm() {
  if (!settingsForm) return;
  const settings = getAppSettings();
  settingsForm.elements.officeName.value = settings.officeName || "";
  settingsForm.elements.postalCode.value = settings.postalCode || "";
  settingsForm.elements.address.value = settings.address || "";
  settingsForm.elements.tel.value = settings.tel || "";
  settingsForm.elements.email.value = settings.email || "";
  settingsForm.elements.invoiceRegistrationNumber.value = settings.invoiceRegistrationNumber || "";
  settingsForm.elements.bankInfo.value = settings.bankInfo || "";
  settingsForm.elements.defaultInvoiceDueDays.value = getDefaultInvoiceDueDays();
  settingsForm.elements.taxRate.value = String(getCurrentTaxRate());
  settingsForm.elements.estimateNote.value = settings.estimateNote || "";
  settingsForm.elements.invoiceNote.value = settings.invoiceNote || "";
}

async function handleSettingsSubmit(event) {
  event.preventDefault();
  if (!currentUser || !settingsForm) return;
  const payload = {
    user_id: currentUser.id,
    office_name: asTrimmedText(settingsForm.elements.officeName.value) || null,
    postal_code: asTrimmedText(settingsForm.elements.postalCode.value) || null,
    address: asTrimmedText(settingsForm.elements.address.value) || null,
    tel: asTrimmedText(settingsForm.elements.tel.value) || null,
    email: asTrimmedText(settingsForm.elements.email.value) || null,
    invoice_registration_number: asTrimmedText(settingsForm.elements.invoiceRegistrationNumber.value) || null,
    bank_info: asTrimmedText(settingsForm.elements.bankInfo.value) || null,
    default_invoice_due_days: getPositiveInt(settingsForm.elements.defaultInvoiceDueDays.value, DEFAULT_APP_SETTINGS.defaultInvoiceDueDays),
    tax_rate: normalizeTaxRate(settingsForm.elements.taxRate.value),
    estimate_note: asTrimmedText(settingsForm.elements.estimateNote.value) || null,
    invoice_note: asTrimmedText(settingsForm.elements.invoiceNote.value) || null,
    updated_at: new Date().toISOString(),
  };
  try {
    await runMutation("設定保存", async () => {
      const existingSettingsId = state.appSettings?.id || null;
      if (existingSettingsId) {
        const { data, error } = await sbClient.from("app_settings").update(payload).eq("id", existingSettingsId).eq("user_id", currentUser.id).select().single();
        if (error) throw error;
        return data;
      }
      const { data: existingRows, error: existingError } = await sbClient.from("app_settings").select("id").eq("user_id", currentUser.id).limit(1);
      if (existingError) throw existingError;
      if (existingRows?.[0]?.id) {
        const { data, error } = await sbClient.from("app_settings").update(payload).eq("id", existingRows[0].id).eq("user_id", currentUser.id).select().single();
        if (error) throw error;
        return data;
      }
      const { data, error } = await sbClient.from("app_settings").insert(payload).select().single();
      if (error) throw error;
      return data;
    }, { successMessage: "設定を保存しました。" });
  } catch (error) {
    showAppMessage(`設定保存に失敗しました。${formatSupabaseError(error)}`, true);
  }
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
  const taxRate = getCurrentTaxRate();
  const tax = Math.floor(subtotal * taxRate);
  const total = subtotal + tax;
  if (estimateSubtotal) estimateSubtotal.textContent = formatCurrency(subtotal);
  if (estimateTax) estimateTax.textContent = formatCurrency(tax);
  if (estimateTaxLabel) estimateTaxLabel.textContent = `消費税(${formatTaxRatePercent(taxRate)})`;
  if (estimateTotal) estimateTotal.textContent = formatCurrency(total);
  return { subtotal, tax, total, taxRate };
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
  console.log("EDIT STATE", editState);
  console.log("ACTION START", taskName, editState.estimateId || "new");
  console.log("PAYLOAD", payload);
  try {
    await runMutation(taskName, async () => {
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
      console.log("DB DONE", taskName, estimateId);
      return { estimateId, isUpdate };
    }, {
      successMessage: editState.estimateId ? "見積を更新しました。" : "見積を登録しました。",
      resetForm: resetEstimateForm,
      afterSuccess: () => {
        editState.estimateId = null;
        subtabState.estimates = "list";
        activateTab("estimates");
      },
    });
  } catch (error) {
    showAppMessage(`見積保存に失敗しました。${formatSupabaseError(error)}`, true);
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
        <button type="button" class="secondary-btn estimate-edit-btn" data-action="edit_estimate" data-list-action="edit_estimate">編集</button>
        <button type="button" class="danger-btn estimate-delete-btn" data-action="delete_estimate" data-list-action="delete_estimate">削除</button>
        <button type="button" class="secondary-btn create-case-btn" data-action="create_case_from_estimate" data-list-action="create_case_from_estimate" ${isCaseCreated ? "disabled" : ""}>${isCaseCreated ? "案件化済み" : "案件化"}</button>
        <button type="button" class="secondary-btn estimate-estimate-print-btn" data-action="print_estimate" data-list-action="print_estimate">見積書出力</button>
        <button type="button" class="secondary-btn estimate-print-btn" data-action="print_invoice_from_estimate" data-list-action="print_invoice_from_estimate">請求書出力</button>
        <button type="button" class="secondary-btn estimate-delivery-note-btn" data-action="print_delivery_note" data-list-action="print_delivery_note">納品書</button>
        <button type="button" class="secondary-btn estimate-purchase-order-btn" data-action="print_purchase_order" data-list-action="print_purchase_order">注文書</button>
        <button type="button" class="secondary-btn estimate-order-confirmation-btn" data-action="print_order_confirmation" data-list-action="print_order_confirmation">注文請書</button>
        <button type="button" class="secondary-btn estimate-inspection-report-btn" data-action="print_acceptance_certificate" data-list-action="print_acceptance_certificate">検収書</button>
        <button type="button" class="secondary-btn estimate-estimate-xlsx-btn" data-action="export_estimate_excel" data-list-action="export_estimate_excel">見積Excel</button>
        <button type="button" class="secondary-btn estimate-xlsx-btn" data-action="export_invoice_excel_from_estimate" data-list-action="export_invoice_excel_from_estimate">請求Excel</button>
        <button type="button" class="secondary-btn create-invoice-btn" data-action="create_invoice_from_estimate" data-list-action="create_invoice_from_estimate" ${hasCreatedInvoice ? "disabled" : ""}>${hasCreatedInvoice ? "請求済み" : "請求作成"}</button>
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
  const listAction = button.dataset.listAction;
  if (listAction === "edit_estimate") return startEstimateEdit(estimateId);
  if (listAction === "delete_estimate") {
    await deleteEstimate(estimateId);
    return;
  }
  if (listAction === "create_case_from_estimate") {
    console.log("CREATE CASE CLICKED", estimateId);
    await handleCreateCaseFromEstimate(estimateId);
    return;
  }
  if (listAction === "print_estimate") return openEstimatePrintPreview(estimateId);
  if (listAction === "print_invoice_from_estimate") return openInvoicePrintPreviewFromEstimate(estimateId);
  if (listAction === "print_delivery_note") return openPeripheralDocumentPrintPreviewFromEstimate(estimateId, "delivery_note");
  if (listAction === "print_purchase_order") return openPeripheralDocumentPrintPreviewFromEstimate(estimateId, "purchase_order");
  if (listAction === "print_order_confirmation") return openPeripheralDocumentPrintPreviewFromEstimate(estimateId, "order_confirmation");
  if (listAction === "print_acceptance_certificate") return openPeripheralDocumentPrintPreviewFromEstimate(estimateId, "inspection_report");
  if (listAction === "export_estimate_excel") return exportEstimateDataForEstimate(estimateId);
  if (listAction === "export_invoice_excel_from_estimate") return exportInvoiceDataForEstimate(estimateId);
  if (listAction === "create_invoice_from_estimate") {
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
  try {
    return await runMutation(loadingLabel, execute, {
      successMessage,
      afterSuccess: () => {
        if (successTab) activateTab(successTab);
        if (successTab && successSubtab) activateSubtab(successTab, successSubtab);
      },
    });
  } catch (error) {
    console.error(`${actionLabel}に失敗しました。`, error);
    if (["DUPLICATE_ESTIMATE_CASE", "DUPLICATE_ESTIMATE_INVOICE"].includes(error?.code)) {
      showAppMessage(error.message || "すでに作成済みです。", true);
    } else {
      showAppMessage(`${actionLabel}に失敗しました。${formatSupabaseError(error)}`, true);
    }
    return null;
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
    due_date: addDaysToDate(new Date(), getDefaultInvoiceDueDays()),
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
  try {
    await runMutation("見積削除", async () => {
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
      return true;
    }, {
      successMessage: "削除しました。",
      afterSuccess: () => {
        if (editState.estimateId === id) {
          editState.estimateId = null;
          resetEstimateForm();
        }
      },
    });
  } catch (error) {
    console.error("削除に失敗しました。", error);
    showAppMessage(`削除に失敗しました。${error.message || ""}`, true);
  }
}


function getCaseDocumentStats(caseId) {
  const docs = (Array.isArray(state.caseDocuments) ? state.caseDocuments : []).filter((entry) => entry.caseId === caseId);
  const total = docs.length;
  const received = docs.filter((entry) => ["回収済", "確認済"].includes(entry.status)).length;
  const unreceived = docs.filter((entry) => entry.status === "未回収").length;
  const defective = docs.filter((entry) => entry.status === "不備あり").length;
  return { total, received, unreceived, defective };
}

function getDocumentAlertTargets() {
  const todayTs = getTodayTimestamp();
  return (Array.isArray(state.caseDocuments) ? state.caseDocuments : []).reduce((acc, doc) => {
    const linkedCase = state.cases.find((entry) => entry.id === doc.caseId);
    const linkedClient = linkedCase?.clientId ? state.clients.find((entry) => entry.id === linkedCase.clientId) : null;
    if (doc.status === "不備あり") {
      acc.push({
        id: doc.id,
        typeLabel: "不備あり書類",
        clientName: linkedClient?.name || linkedCase?.customerName || "顧客不明",
        caseName: linkedCase?.caseName || "案件なし",
        documentName: doc.documentName || "未設定",
        subInfo: "修正・再提出が必要です",
        priority: 0,
      });
    }
    if (doc.status === "未回収") {
      const nearDeadline = Number.isFinite(toDateOnlyTimestamp(linkedCase?.dueDate))
        && toDateOnlyTimestamp(linkedCase?.dueDate) <= todayTs + (7 * 86400000);
      acc.push({
        id: doc.id,
        typeLabel: nearDeadline ? "期限7日以内・未回収" : "未回収書類",
        clientName: linkedClient?.name || linkedCase?.customerName || "顧客不明",
        caseName: linkedCase?.caseName || "案件なし",
        documentName: doc.documentName || "未設定",
        subInfo: nearDeadline ? `期限: ${formatDate(linkedCase?.dueDate)}` : "回収待ち",
        priority: nearDeadline ? 1 : 2,
      });
    }
    return acc;
  }, []).sort((a, b) => (a.priority || 99) - (b.priority || 99));
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
    const docStats = getCaseDocumentStats(entry.id);

    item.dataset.id = entry.id;
    title.textContent = `${customerName}｜${caseName}`;
    const urlLinks = [
      entry.documentUrl ? `<a href="${escapeHtml(entry.documentUrl)}" target="_blank" rel="noopener noreferrer">関連書類を開く</a>` : "",
      entry.invoiceUrl ? `<a href="${escapeHtml(entry.invoiceUrl)}" target="_blank" rel="noopener noreferrer">請求書を開く</a>` : "",
      entry.receiptUrl ? `<a href="${escapeHtml(entry.receiptUrl)}" target="_blank" rel="noopener noreferrer">領収書を開く</a>` : "",
    ].filter(Boolean).join(" / ");
    meta.innerHTML = `見積: ${formatCurrency(entry.estimateAmount)} / ステータス: ${escapeHtml(entry.status)} / 受付日: ${formatDate(entry.receivedDate)} / 期限日: ${formatDate(entry.dueDate)} / 次回対応日: ${formatDate(entry.nextActionDate)} / 次回対応内容: ${escapeHtml(entry.nextAction || "未設定")}${urlLinks ? ` / ${urlLinks}` : ""}`;
    if (caseWorkMeta) {
      caseWorkMeta.textContent = `テンプレート: ${templateName} / 必要書類: ${truncateText(entry.requiredDocuments || "未設定", 50)} / 書類管理: 必要書類 ${docStats.total}件 / 回収済 ${docStats.received}件 / 未回収 ${docStats.unreceived}件 / 不備 ${docStats.defective}件 / タスク: ${truncateText(entry.taskList || "未設定", 50)} / 未完了タスク: ${incompleteTasks}件 / 作業メモ: ${truncateText(sanitizeLegacyEstimateMemo(entry.workMemo) || "未設定", 40)}`;
      caseWorkMeta.classList.remove("next-action-overdue", "next-action-within3", "next-action-within7");
      if (nextActionInfo) safeAddClass(caseWorkMeta, nextActionInfo.urgencyClass);
    }
    profitMeta.textContent = `売上合計: ${formatCurrency(totals.sales)} / 経費合計: ${formatCurrency(totals.expenses)} / 利益: ${formatCurrency(totals.profit)}`;
    profitMeta.classList.toggle("loss-text", totals.profit < 0);
    if (rowActions && !rowActions.querySelector(".case-print-btn")) {
      const printBtn = document.createElement("button");
      printBtn.type = "button";
      printBtn.className = "secondary-btn case-print-btn";
      printBtn.dataset.action = "print_case_invoice";
      printBtn.dataset.listAction = "print_case_invoice";
      printBtn.textContent = "請求書出力";
      rowActions.appendChild(printBtn);
    }
    if (rowActions && !rowActions.querySelector(".case-xlsx-btn")) {
      const xlsxBtn = document.createElement("button");
      xlsxBtn.type = "button";
      xlsxBtn.className = "secondary-btn case-xlsx-btn";
      xlsxBtn.dataset.action = "export_case_invoice_excel";
      xlsxBtn.dataset.listAction = "export_case_invoice_excel";
      xlsxBtn.textContent = "請求Excel";
      rowActions.appendChild(xlsxBtn);
    }
    const caseDocumentButtons = [
      { selector: ".case-delivery-note-btn", action: "print_case_delivery_note", label: "納品書" },
      { selector: ".case-purchase-order-btn", action: "print_case_purchase_order", label: "注文書" },
      { selector: ".case-order-confirmation-btn", action: "print_case_order_confirmation", label: "注文請書" },
      { selector: ".case-inspection-report-btn", action: "print_case_acceptance_certificate", label: "検収書" },
    ];
    caseDocumentButtons.forEach((config) => {
      if (!rowActions || rowActions.querySelector(config.selector)) return;
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = `secondary-btn ${config.selector.slice(1)}`;
      btn.dataset.action = config.action;
      btn.dataset.listAction = config.action;
      btn.dataset.caseId = entry.id;
      btn.textContent = config.label;
      rowActions.appendChild(btn);
    });
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
        <p class="meta">売上合計: ${formatCurrency(totals.sales)} / 経費合計: ${formatCurrency(totals.expenses)} / 作業時間: ${formatMinutes(totals.workMinutes)} / <span class="${safeClassName(totals.profit < 0 ? "loss-text" : "")}">利益: ${formatCurrency(totals.profit)}</span></p>
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
      <td class="${safeClassName(profit < 0 ? "loss-text" : "")}">${formatCurrency(profit)}</td>
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
      const remainingAmount = calculateSaleRemainingAmount(invoiceAmount, paidAmount);
      const paymentStatus = sale.paymentStatus || calculatePaymentStatus(invoiceAmount, paidAmount);
      const invoiceRegistrationStatus = getInvoiceRegistrationStatus(sale);
      const invoiceDeliveryStatus = getInvoiceDeliveryStatus(sale);
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
      const paymentHistory = renderPaymentHistoryInline(sale.id);
      const tr = document.createElement("tr");
      tr.dataset.id = sale.id || "";
      safeAddClass(tr, getSaleRowClass(safeSale));
      tr.innerHTML = `
        <td>${escapeHtml(customerLabel)}</td>
        <td>${escapeHtml(caseLabel)}<br /><small>${escapeHtml(sale.invoiceNumber || "未採番")}</small></td>
        <td>${escapeHtml(invoiceRegistrationStatus)}</td>
        <td>${escapeHtml(invoiceDeliveryStatus)}</td>
        <td>${escapeHtml(paymentStatus)}</td>
        <td>${formatCurrency(invoiceAmount)}</td>
        <td>${formatCurrency(paidAmount)}</td>
        <td>${formatCurrency(remainingAmount)}</td>
        <td><span class="${safeClassName(getSaleStatusClass(safeSale))}">${escapeHtml(paymentStatus)}</span></td>
        <td>${dueDateLabel}</td>
        <td>${paidDateLabel}</td>
        <td>${reminderCount}回</td>
        <td>${lastReminderDateLabel}</td>
        <td>${escapeHtml(reminderMethod)}</td>
        <td>${escapeHtml(reminderMemo)}</td>
        <td>${paymentHistory}${renderSaleCumulativeReceiptButton(sale)}</td>
        <td>${canRecordReminder ? `<button type="button" class="secondary-btn record-reminder-btn" data-action="record_reminder" data-list-action="record_reminder" data-sale-id="${sale.id}">督促記録</button>` : "-"}</td>
        <td><button type="button" class="secondary-btn register-payment-btn record-payment-btn" data-action="record_payment" data-list-action="record_payment" data-sale-id="${sale.id}">入金登録</button></td>
        <td><button type="button" class="edit-sale-btn secondary-btn" data-action="edit_sale" data-list-action="edit_sale" data-sale-id="${sale.id}">編集</button></td>
        <td><button type="button" class="danger-btn delete-btn delete-sale-btn" data-action="delete_sale" data-list-action="delete_sale" data-sale-id="${sale.id}">削除</button></td>
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

function resetCaseTaskForm() {
  resetEditMode("caseTask");
  caseTaskForm?.reset();
  if (caseTaskSubmitBtn) caseTaskSubmitBtn.textContent = "タスクを追加";
  if (caseTaskForm?.elements?.caseTaskStatus) caseTaskForm.elements.caseTaskStatus.value = "未完了";
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
  renderPayments();
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

function resetCaseDocumentForm() {
  resetEditMode("caseDocument");
  caseDocumentForm?.reset();
  if (caseDocumentForm?.elements?.caseDocumentStatus) caseDocumentForm.elements.caseDocumentStatus.value = "未回収";
  if (caseDocumentSubmitBtn) caseDocumentSubmitBtn.textContent = "書類を登録";
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
  if (target === "caseTask") editState.caseTaskId = null;
  if (target === "caseDocument") editState.caseDocumentId = null;
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
  const appSettings = getAppSettings();
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
    companyName: appSettings.officeName,
    companyZip: appSettings.postalCode,
    companyAddress: appSettings.address,
    companyPhone: appSettings.tel,
    companyEmail: appSettings.email,
    registrationNumber: appSettings.invoiceRegistrationNumber,
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
    transferInfo: appSettings.bankInfo,
    taxRate: getCurrentTaxRate(),
    note: asTrimmedText(noteOverride ?? "") || estimate.memo || appSettings.invoiceNote || "",
  };
}

function exportInvoiceDataForEstimate(estimateId) {
  const estimate = state.estimates.find((entry) => entry.id === estimateId);
  if (!estimate) return;
  safeFileExport("請求Excel出力", () => {
    const note = requestInvoiceMemo(estimate.memo || "");
    if (note === null) return false;
    const invoiceData = buildInvoiceDocumentFromEstimate(estimate, note);
    downloadInvoiceWorkbook(invoiceData);
    return true;
  });
}

function buildInvoiceDocumentFromCase(foundCase, noteOverride = null) {
  const subtotal = foundCase.estimateAmount ?? 0;
  const appSettings = getAppSettings();
  const taxRate = getCurrentTaxRate();
  const tax = Math.floor(subtotal * taxRate);
  const total = subtotal + tax;
  const invoiceDate = toDateString(new Date());
  const linkedSale = state.sales.find((sale) => sale.caseId === foundCase.id);
  return {
    customerName: foundCase.customerName || "顧客名未設定",
    subject: foundCase.caseName || "請求内容",
    invoiceDate,
    invoiceNumber: linkedSale?.invoiceNumber || "未採番",
    companyName: appSettings.officeName,
    companyZip: appSettings.postalCode,
    companyAddress: appSettings.address,
    companyPhone: appSettings.tel,
    companyEmail: appSettings.email,
    registrationNumber: appSettings.invoiceRegistrationNumber,
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
    transferInfo: appSettings.bankInfo,
    taxRate,
    note: asTrimmedText(noteOverride ?? "") || appSettings.invoiceNote || "",
  };
}

function exportInvoiceDataForCase(caseId) {
  const foundCase = state.cases.find((entry) => entry.id === caseId);
  if (!foundCase) return;
  safeFileExport("請求Excel出力", () => {
    const note = requestInvoiceMemo(asTrimmedText(saleInvoiceMemoInput?.value || ""));
    if (note === null) return false;
    const invoiceData = buildInvoiceDocumentFromCase(foundCase, note);
    downloadInvoiceWorkbook(invoiceData);
    return true;
  });
}

function openEstimatePrintPreview(estimateId) {
  const estimate = state.estimates.find((entry) => entry.id === estimateId);
  if (!estimate) {
    showAppMessage("出力対象データが見つかりません", true);
    return;
  }
  safeFileExport("見積書出力", () => {
    const documentData = buildEstimateDocumentFromEstimate(estimate);
    openBusinessDocumentPrintWindow(documentData, { type: "estimate" });
  });
}

function openInvoicePrintPreviewFromEstimate(estimateId) {
  const estimate = state.estimates.find((entry) => entry.id === estimateId);
  if (!estimate) {
    showAppMessage("出力対象データが見つかりません", true);
    return;
  }
  safeFileExport("請求書出力", () => {
    const note = requestInvoiceMemo(estimate.memo || "");
    if (note === null) return false;
    const documentData = buildInvoiceDocumentFromEstimate(estimate, note);
    openBusinessDocumentPrintWindow(documentData, { type: "invoice" });
    return true;
  });
}

function openInvoicePrintPreviewFromCase(caseId) {
  const foundCase = state.cases.find((entry) => entry.id === caseId);
  if (!foundCase) {
    showAppMessage("出力対象データが見つかりません", true);
    return;
  }
  safeFileExport("請求書出力", () => {
    const note = requestInvoiceMemo(asTrimmedText(saleInvoiceMemoInput?.value || ""));
    if (note === null) return false;
    const documentData = buildInvoiceDocumentFromCase(foundCase, note);
    openBusinessDocumentPrintWindow(documentData, { type: "invoice" });
    return true;
  });
}


function openReceiptPrintPreview(saleId, paymentId, options = {}) {
  const sale = state.sales.find((entry) => entry.id === saleId);
  if (!sale) {
    showAppMessage("出力対象データが見つかりません", true);
    return;
  }
  const payment = paymentId
    ? state.payments.find((entry) => entry.id === paymentId && entry.saleId === saleId)
    : buildReceiptPaymentFromSale(sale, options);
  if (!payment) {
    showAppMessage("出力対象データが見つかりません", true);
    return;
  }
  safeFileExport("領収書出力", () => {
    const documentData = buildReceiptDocumentData(sale, payment);
    openBusinessDocumentPrintWindow(documentData, { type: "receipt" });
  });
}

function openCumulativeReceiptPrintPreview(saleId) {
  if (!saleId) {
    showAppMessage("出力対象データIDを取得できませんでした。", true);
    return;
  }
  return openReceiptPrintPreview(saleId, null, { cumulative: true });
}

function buildReceiptPaymentFromSale(sale, options = {}) {
  const amount = Math.max(Number(sale?.paidAmount || 0), 0);
  if (amount <= 0) return null;
  const latestPayment = getSalePayments(sale?.id || "")[0] || null;
  const paymentDate = sale?.paidDate || latestPayment?.paymentDate || toDateString(new Date());
  const fallbackId = String(sale?.id || "").replace(/[^a-zA-Z0-9]/g, "").slice(-6).toUpperCase() || "000001";
  const receiptNumberSource = sale?.invoiceNumber ? `R-${sale.invoiceNumber}` : `R-${fallbackId}`;
  return {
    id: `sale-${sale.id || fallbackId}`,
    amount,
    paymentDate,
    method: latestPayment?.method || sale?.paymentMethod || "未設定",
    memo: "上記金額を領収いたしました。",
    receiptNumber: receiptNumberSource,
    receiptType: options?.cumulative ? "cumulative" : "payment",
  };
}

function buildReceiptDocumentData(sale, payment) {
  const appSettings = getAppSettings();
  const taxRate = getCurrentTaxRate() || 0.1;
  const total = Math.max(Number(payment.amount || 0), 0);
  const subtotal = Math.floor(total / (1 + taxRate));
  const tax = total - subtotal;
  const paymentDate = payment.paymentDate || toDateString(new Date());
  const dateToken = String(paymentDate || "").replace(/-/g, "").slice(0, 6) || toDateString(new Date()).replace(/-/g, "").slice(0, 6);
  const paymentIdText = String(payment.id || "").replace(/[^a-zA-Z0-9]/g, "").slice(-6).toUpperCase() || "000001";
  const receiptNumber = payment.receiptNumber || `R-${dateToken}-${paymentIdText}`;
  return {
    receiptType: payment.receiptType || "payment",
    customerName: sale.customerName || resolveCustomerNameBySale(sale),
    subject: sale.caseName || resolveCaseName(sale.caseId) || "案件名未設定",
    receiptDate: paymentDate,
    receiptNumber,
    paymentMethod: payment.method || "その他",
    paymentDate,
    referenceInvoiceNumber: sale.invoiceNumber || "未採番",
    companyName: appSettings.officeName,
    companyZip: appSettings.postalCode,
    companyAddress: appSettings.address,
    companyPhone: appSettings.tel,
    companyEmail: appSettings.email,
    registrationNumber: appSettings.invoiceRegistrationNumber,
    details: [{ itemName: `行政書士業務報酬として / ${resolveCaseName(sale.caseId) || "案件名未設定"}`, quantity: 1, unitPrice: total, amount: total }],
    subtotal,
    tax,
    total,
    taxRate,
    note: payment.memo || "",
  };
}

function resolveCustomerNameBySale(sale) {
  const linkedCase = sale?.caseId ? state.cases.find((entry) => entry.id === sale.caseId) : null;
  const linkedClient = linkedCase?.clientId ? state.clients.find((entry) => entry.id === linkedCase.clientId) : null;
  return linkedClient?.name || linkedCase?.customerName || "顧客名未設定";
}


function buildPeripheralDocumentFromEstimate(estimate, documentType) {
  const appSettings = getAppSettings();
  const taxRate = getCurrentTaxRate();
  const subtotal = Number(estimate.subtotal ?? Math.floor(Number(estimate.total ?? 0) / (1 + taxRate)) ?? 0);
  const tax = Number(estimate.tax ?? Math.floor(subtotal * taxRate));
  const total = Number(estimate.total ?? (subtotal + tax));
  const linkedSale = estimate.caseId ? state.sales.find((sale) => sale.caseId === estimate.caseId) : null;
  return {
    documentType,
    customerName: estimate.customerName || "顧客名未設定",
    subject: estimate.estimateTitle || "案件名未設定",
    issueDate: toDateString(new Date()),
    documentNumber: linkedSale?.invoiceNumber || estimate.estimateNumber || "未採番",
    companyName: appSettings.officeName,
    companyZip: appSettings.postalCode,
    companyAddress: appSettings.address,
    companyPhone: appSettings.tel,
    companyEmail: appSettings.email,
    registrationNumber: appSettings.invoiceRegistrationNumber,
    subtotal,
    tax,
    total,
    taxRate,
    note: estimate.memo || "",
    details: [{ itemName: estimate.estimateTitle || "業務一式", quantity: 1, unitPrice: subtotal, amount: subtotal }],
  };
}

function buildPeripheralDocumentFromCase(foundCase, documentType) {
  const appSettings = getAppSettings();
  const taxRate = getCurrentTaxRate();
  const subtotal = Number(foundCase.estimateAmount ?? 0);
  const tax = Math.floor(subtotal * taxRate);
  const total = subtotal + tax;
  const linkedSale = state.sales.find((sale) => sale.caseId === foundCase.id);
  return {
    documentType,
    customerName: foundCase.customerName || "顧客名未設定",
    subject: foundCase.caseName || "案件名未設定",
    issueDate: toDateString(new Date()),
    documentNumber: linkedSale?.invoiceNumber || "未採番",
    companyName: appSettings.officeName,
    companyZip: appSettings.postalCode,
    companyAddress: appSettings.address,
    companyPhone: appSettings.tel,
    companyEmail: appSettings.email,
    registrationNumber: appSettings.invoiceRegistrationNumber,
    subtotal,
    tax,
    total,
    taxRate,
    note: sanitizeLegacyEstimateMemo(foundCase.workMemo) || "",
    details: [{ itemName: foundCase.caseName || "業務一式", quantity: 1, unitPrice: subtotal, amount: subtotal }],
  };
}

function openPeripheralDocumentPrintPreviewFromEstimate(estimateId, documentType) {
  const estimate = state.estimates.find((entry) => entry.id === estimateId);
  if (!estimate) {
    showAppMessage("出力対象データが見つかりません", true);
    return;
  }
  safeFileExport("帳票出力", () => {
    const documentData = buildPeripheralDocumentFromEstimate(estimate, documentType);
    openBusinessDocumentPrintWindow(documentData, { type: "peripheral" });
  });
}

function openPeripheralDocumentPrintPreviewFromCase(caseId, documentType) {
  const foundCase = state.cases.find((entry) => entry.id === caseId);
  if (!foundCase) {
    showAppMessage("出力対象データが見つかりません", true);
    return;
  }
  safeFileExport("帳票出力", () => {
    const documentData = buildPeripheralDocumentFromCase(foundCase, documentType);
    openBusinessDocumentPrintWindow(documentData, { type: "peripheral" });
  });
}

function openCaseBusinessDocumentPrintPreview(caseId, documentType) {
  return openPeripheralDocumentPrintPreviewFromCase(caseId, documentType);
}

function openBusinessDocumentPrintWindow(documentData, options = { type: "invoice" }) {
  let html = "";
  try {
    html = buildBusinessDocumentHtml(documentData, options);
  } catch (error) {
    console.error("帳票HTMLの生成に失敗しました。", error);
    throw new Error("帳票の生成に失敗しました。再度お試しください。");
  }

  if (typeof html !== "string" || !html.trim()) {
    console.error("帳票HTMLの生成結果が不正です。", { documentData, options, html });
    throw new Error("帳票の生成に失敗しました。再度お試しください。");
  }

  const printWindow = window.open("", "_blank");
  if (!printWindow) {
    throw new Error("ポップアップがブロックされています。ブラウザ設定を確認してください。");
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
  const appSettings = getAppSettings();
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
    companyName: appSettings.officeName,
    companyZip: appSettings.postalCode,
    companyAddress: appSettings.address,
    companyPhone: appSettings.tel,
    companyEmail: appSettings.email,
    details: rows,
    subtotal: estimate.subtotal ?? 0,
    tax: estimate.tax ?? 0,
    total: estimate.total ?? 0,
    paymentTerms: `お支払い条件：請求書受領後${getDefaultInvoiceDueDays()}日以内に銀行振込`,
    taxRate: getCurrentTaxRate(),
    note: estimate.memo || appSettings.estimateNote,
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
  safeFileExport("見積Excel出力", () => {
    const estimateData = buildEstimateDocumentFromEstimate(estimate);
    downloadEstimateWorkbook(estimateData);
  });
}

function createBusinessDocumentSheet(documentData, options = { type: "invoice" }) {
  const isInvoice = options.type === "invoice";
  const isReceipt = options.type === "receipt";
  const title = isReceipt ? "領収書" : (isInvoice ? "請求書" : "見積書");
  const amountLabel = isInvoice ? "ご請求金額（税込）" : "お見積金額（税込）";
  const sentence = isInvoice ? "下記の通りご請求申し上げます。" : "下記の通りお見積り申し上げます。";
  const leftLabel = isInvoice ? "請求先" : "宛先";
  const dateLabel = isInvoice ? "発行日" : "見積日";
  const numberLabel = isReceipt ? "領収書番号" : (isInvoice ? "請求書番号" : "見積番号");
  const numberValue = isReceipt ? documentData.receiptNumber : (isInvoice ? documentData.invoiceNumber : documentData.estimateNumber);
  const dateValue = isReceipt ? documentData.receiptDate : (isInvoice ? documentData.invoiceDate : documentData.estimateDate);
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
  rows[27][5] = `消費税${formatTaxRatePercent(documentData.taxRate)}`;
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
  const isReceipt = options.type === "receipt";
  const peripheralMap = {
    delivery_note: { title: "納品書", description: "下記のとおり納品いたしました。" },
    purchase_order: { title: "注文書", description: "下記のとおり注文いたします。" },
    order_confirmation: { title: "注文請書", description: "下記のとおりご注文を承りました。" },
    inspection_report: { title: "検収書", description: "下記のとおり検収いたしました。" },
  };
  const peripheralConfig = peripheralMap[documentData.documentType] || { title: "帳票", description: "下記のとおり発行いたしました。" };
  const isPeripheral = options.type === "peripheral";
  const title = isPeripheral ? peripheralConfig.title : (isReceipt ? "領収書" : (isInvoice ? "請求書" : "見積書"));
  const description = isPeripheral ? peripheralConfig.description : (isReceipt ? "下記のとおり領収いたしました。" : (isInvoice ? "下記のとおりご請求申し上げます。" : "下記のとおり、お見積もり申し上げます。"));
  const emphasizedLabel = isPeripheral ? "金額（税込）" : (isReceipt ? "領収金額（税込）" : (isInvoice ? "ご請求金額（税込）" : "お見積金額（税込）"));
  const numberLabel = isPeripheral ? "書類番号" : (isReceipt ? "領収書番号" : (isInvoice ? "請求書番号" : "見積番号"));
  const numberValue = isPeripheral ? documentData.documentNumber : (isReceipt ? documentData.receiptNumber : (isInvoice ? documentData.invoiceNumber : documentData.estimateNumber));
  const dateLabel = isInvoice ? "日付" : "日付";
  const dateValue = isPeripheral ? documentData.issueDate : (isReceipt ? documentData.receiptDate : (isInvoice ? documentData.invoiceDate : documentData.estimateDate));
  const hasCompanyZip = Boolean(asTrimmedText(documentData.companyZip || ""));
  const hasCompanyAddress = Boolean(asTrimmedText(documentData.companyAddress || ""));
  const hasCompanyPhone = Boolean(asTrimmedText(documentData.companyPhone || ""));
  const hasRegistrationNumber = Boolean(asTrimmedText(documentData.registrationNumber || ""));
  const hasTransferInfo = Boolean(asTrimmedText(documentData.transferInfo || ""));
  const extraLabel = isPeripheral ? "登録番号" : (isReceipt ? "対象請求番号" : (isInvoice ? "登録番号" : "有効期限"));
  const extraValue = isPeripheral ? documentData.registrationNumber : (isReceipt ? documentData.referenceInvoiceNumber : (isInvoice ? documentData.registrationNumber : documentData.validUntil));
  const extraRowHtml = isPeripheral || isReceipt || !isInvoice || hasRegistrationNumber
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
  const receiptMetaHtml = isReceipt
    ? `<section class="info-section"><h3>領収情報</h3><div class="note-box">但し書き: ${escapeHtml(`行政書士業務報酬として / ${documentData.subject || "案件名未設定"}`)}\n支払方法: ${escapeHtml(documentData.paymentMethod || "その他")}\n入金日: ${escapeHtml(formatDate(documentData.paymentDate))}\n案件名: ${escapeHtml(documentData.subject || "案件名未設定")}</div></section>`
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
  <div class="print-toolbar no-print"><button type="button" class="print-btn" id="print-preview-btn">印刷 / PDF保存</button></div>
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
      <tr><th>消費税${formatTaxRatePercent(documentData.taxRate)}</th><td class="align-right">${formatCurrencyCompact(documentData.tax)}</td></tr>
      <tr class="total-row"><th>合計</th><td class="align-right">${formatCurrencyCompact(documentData.total)}</td></tr>
    </table>
    <section class="document-footer">
      ${isInvoice ? transferSectionHtml : ""}
      ${receiptMetaHtml}
      <section class="info-section">
        <h3>${isInvoice ? "備考" : "備考"}</h3>
        <div class="note-box">${escapeHtml(isInvoice ? (documentData.note || "") : `${documentData.note || ""}\n${documentData.paymentTerms || "支払条件: 記載なし"}\n有効期限: ${documentData.validUntil || "記載なし"}`.trim())}</div>
      </section>
    </section>
</main>
<script>
  document.getElementById("print-preview-btn")?.addEventListener("click", () => window.print());
</script>
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

function formatDateTimeValue(value) {
  if (!value) return "未設定";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "未設定";
  return new Intl.DateTimeFormat("ja-JP", { dateStyle: "short", timeStyle: "medium" }).format(date);
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
  if (filterKey === "within3") return info.deadlineStatus === "within3";
  if (filterKey === "within7") return info.deadlineStatus === "within7";
  if (filterKey === "within30") return ["within30", "within7", "within3"].includes(info.deadlineStatus);
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

function getInvoiceRegistrationStatus(sale) {
  return asTrimmedText(sale?.invoiceNumber) || asTrimmedText(sale?.invoiceUrl) ? "作成済" : "未作成";
}

function getInvoiceDeliveryStatus(sale) {
  return asTrimmedText(sale?.invoiceUrl) ? "送付済" : "未送付";
}

function normalizePaymentMethod(value) {
  const normalized = asTrimmedText(value);
  if (!normalized) return "その他";
  return PAYMENT_METHODS.includes(normalized) ? normalized : "その他";
}

function getSalePayments(saleId) {
  return state.payments
    .filter((entry) => entry.saleId === saleId)
    .slice()
    .sort((a, b) => toSortTimestamp(b.paymentDate) - toSortTimestamp(a.paymentDate));
}

function calculateSaleRemainingAmount(invoiceAmount, paidAmount) {
  return (Number(invoiceAmount) || 0) - (Number(paidAmount) || 0);
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

function normalizeCaseDocumentStatus(raw) {
  const value = asTrimmedText(raw);
  if (!value) return "未回収";
  return CASE_DOCUMENT_STATUSES.includes(value) ? value : "未回収";
}

function importTypeToLabel(type) {
  if (type === "cases") return "案件CSV";
  if (type === "sales") return "売上CSV";
  if (type === "payments") return "入金CSV";
  if (type === "expenses") return "経費CSV";
  if (type === "fixed_expenses") return "固定費CSV";
  if (type === "case_documents") return "書類管理CSV";
  if (type === "client_interactions") return "対応履歴CSV";
  return "CSV";
}

function safeFileExport(actionName, exportFn) {
  try {
    clearAppMessage();
    console.log("FILE EXPORT START", actionName);
    const result = exportFn();
    if (result === false) {
      console.log("FILE EXPORT CANCELED", actionName);
      return;
    }
    showAppMessage(`${actionName}を出力しました。`, false);
    console.log("FILE EXPORT DONE", actionName);
  } catch (error) {
    console.error(`${actionName}の出力に失敗しました`, error);
    showAppMessage(`${actionName}の出力に失敗しました。${error?.message || ""}`, true);
  } finally {
    forceHideLoading();
  }
}

async function safeFileExportAsync(actionName, exportFn) {
  try {
    clearAppMessage();
    console.log("FILE EXPORT START", actionName);
    const result = await exportFn();
    if (result === false) {
      console.log("FILE EXPORT CANCELED", actionName);
      return;
    }
    showAppMessage(`${actionName}を出力しました。`, false);
    console.log("FILE EXPORT DONE", actionName);
  } catch (error) {
    console.error(`${actionName}の出力に失敗しました`, error);
    showAppMessage(`${actionName}の出力に失敗しました。${error?.message || ""}`, true);
  } finally {
    forceHideLoading();
  }
}

function safeSetLocalStorage(key, value) {
  try {
    if (!key) return;
    window.localStorage.setItem(String(key), String(value ?? ""));
  } catch (error) {
    console.warn("localStorage set failed", error);
  }
}

function safeGetLocalStorage(key) {
  try {
    if (!key) return "";
    return window.localStorage.getItem(String(key)) || "";
  } catch (error) {
    console.warn("localStorage get failed", error);
    return "";
  }
}

function downloadCsvFile(filename, headers, rows) {
  const csv = buildCsvString(headers, rows);
  const bom = "\uFEFF";
  downloadTextFile(filename, bom + csv, "text/csv;charset=utf-8");
}

function downloadJsonFile(filename, value) {
  const pretty = JSON.stringify(value, null, 2);
  downloadTextFile(filename, pretty, "application/json;charset=utf-8");
}

function downloadTextFile(filename, content, mimeType = "text/plain;charset=utf-8") {
  if (!filename) throw new Error("ファイル名がありません。");
  if (content === undefined || content === null) throw new Error("出力内容がありません。");
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  console.log("DOWNLOAD START", filename);
  link.href = url;
  link.download = filename;
  link.style.display = "none";
  document.body.appendChild(link);
  link.click();
  console.log("DOWNLOAD CLICKED", filename);
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
  const paymentsData = getSheetRows("入金");
  if (paymentsData.rows.length) {
    mergeResult(await importRowsByType("payments", paymentsData));
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
  const caseTasksData = getSheetRows("案件タスク");
  if (caseTasksData.rows.length) {
    mergeResult(await importRowsByType("case_tasks", caseTasksData));
  }
  const caseDocumentsData = getSheetRows("書類管理");
  if (caseDocumentsData.rows.length) {
    mergeResult(await importRowsByType("case_documents", caseDocumentsData));
  }
  return total;
}

async function loadCasesOnly() {
  if (!currentUser || isLoggingOut) {
    state.cases = [];
    return;
  }
  try {
    const casesRes = await sbClient.from("cases").select("*").eq("user_id", currentUser.id);
    if (casesRes.error) {
      console.error("LOAD casesOnly ERROR", casesRes.error);
      state.cases = [];
      return;
    }
    state.cases = (casesRes.data || []).map(mapCaseFromDb);
  } catch (error) {
    console.error("LOAD casesOnly ERROR", error);
    state.cases = [];
  }
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

  if (importType === "payments") {
    validateRequiredHeaders(headers, ["sale_invoice_number", "payment_date", "amount"]);
    const saleMap = new Map(state.sales.map((s) => [s.invoiceNumber, s.id]));
    const payloads = [];
    rows.forEach((row) => {
      try {
        const saleInvoiceNumber = asTrimmedText(row.sale_invoice_number);
        const saleId = saleMap.get(saleInvoiceNumber);
        if (!saleId) {
          result.skippedCount += 1;
          return;
        }
        const paymentDate = parseFlexibleDate(row.payment_date);
        const amount = parseFlexibleAmount(row.amount);
        if (!paymentDate || amount === null || amount <= 0) {
          result.errorCount += 1;
          return;
        }
        payloads.push({
          user_id: currentUser.id,
          sale_id: saleId,
          payment_date: paymentDate,
          amount,
          method: normalizePaymentMethod(row.method || row.payment_method),
          memo: asTrimmedText(row.memo || row.payment_memo) || null,
        });
      } catch (_error) {
        result.errorCount += 1;
      }
    });
    if (payloads.length) {
      const { error } = await sbClient.from("payments").insert(payloads);
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

  if (importType === "case_tasks") {
    validateRequiredHeaders(headers, ["task_title", "status"]);
    const payloads = [];
    rows.forEach((row) => {
      try {
        const taskTitle = asTrimmedText(row.task_title);
        if (!taskTitle) {
          result.skippedCount += 1;
          return;
        }
        const caseName = asTrimmedText(row.case_name);
        const linkedCase = state.cases.find((entry) => entry.caseName === caseName && (!row.customer_name || entry.customerName === asTrimmedText(row.customer_name)));
        const status = asTrimmedText(row.status) === "完了" ? "完了" : "未完了";
        payloads.push({
          user_id: currentUser.id,
          case_id: linkedCase?.id || null,
          task_title: taskTitle,
          task_memo: asTrimmedText(row.task_memo) || null,
          due_date: parseFlexibleDate(row.due_date),
          status,
          completed_at: status === "完了" ? parseFlexibleDate(row.completed_at) || toDateString(new Date()) : null,
        });
      } catch (_error) {
        result.errorCount += 1;
      }
    });
    if (payloads.length) {
      const { error } = await sbClient.from("case_tasks").insert(payloads);
      if (error) throw error;
      result.insertedCount += payloads.length;
    }
    return result;
  }

  if (importType === "case_documents") {
    validateRequiredHeaders(headers, ["document_name", "status"]);
    const payloads = [];
    rows.forEach((row) => {
      try {
        const documentName = asTrimmedText(row.document_name);
        if (!documentName) {
          result.skippedCount += 1;
          return;
        }
        const caseName = asTrimmedText(row.case_name);
        const linkedCase = state.cases.find((entry) => entry.caseName === caseName && (!row.customer_name || entry.customerName === asTrimmedText(row.customer_name)));
        payloads.push({
          user_id: currentUser.id,
          case_id: linkedCase?.id || null,
          document_name: documentName,
          status: normalizeCaseDocumentStatus(row.status),
          received_date: parseFlexibleDate(row.received_date),
          checked_date: parseFlexibleDate(row.checked_date),
          memo: asTrimmedText(row.memo) || null,
          file_url: asTrimmedText(row.file_url) || null,
        });
      } catch (_error) {
        result.errorCount += 1;
      }
    });
    if (payloads.length) {
      const { error } = await sbClient.from("case_documents").insert(payloads);
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

function normalizeAutoCreateItemName(text) {
  return asTrimmedText(text).toLowerCase().replace(/\s+/g, " ");
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

function mapAppSettingsFromDb(row) {
  return {
    id: row.id,
    userId: row.user_id,
    officeName: row.office_name || "",
    postalCode: row.postal_code || "",
    address: row.address || "",
    tel: row.tel || "",
    email: row.email || "",
    invoiceRegistrationNumber: row.invoice_registration_number || "",
    bankInfo: row.bank_info || "",
    defaultInvoiceDueDays: getPositiveInt(row.default_invoice_due_days, DEFAULT_APP_SETTINGS.defaultInvoiceDueDays),
    taxRate: normalizeTaxRate(row.tax_rate),
    estimateNote: row.estimate_note || "",
    invoiceNote: row.invoice_note || "",
    createdAt: Date.parse(row.created_at) || Date.now(),
    updatedAt: Date.parse(row.updated_at) || Date.now(),
  };
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

function mapCaseTaskFromDb(row) {
  return {
    id: row.id,
    userId: row.user_id,
    caseId: row.case_id || null,
    taskTitle: row.task_title || "",
    taskMemo: row.task_memo || "",
    dueDate: row.due_date || "",
    status: row.status === "完了" ? "完了" : "未完了",
    completedAt: row.completed_at || "",
    createdAt: Date.parse(row.created_at) || Date.now(),
    updatedAt: Date.parse(row.updated_at) || Date.now(),
  };
}

function mapCaseDocumentFromDb(row) {
  return {
    id: row.id,
    userId: row.user_id || null,
    caseId: row.case_id || null,
    documentName: row.document_name || "",
    status: normalizeCaseDocumentStatus(row.status),
    receivedDate: row.received_date || null,
    checkedDate: row.checked_date || null,
    memo: row.memo || "",
    fileUrl: row.file_url || "",
    createdAt: Date.parse(row.created_at) || Date.now(),
    updatedAt: Date.parse(row.updated_at) || Date.now(),
  };
}

function mapWorkTemplateFromDb(row) {
  // DB拡張が未適用の場合は以下を実行してください:
  // alter table work_templates add column if not exists standard_estimate_amount bigint;
  // alter table work_templates add column if not exists task_list text;
  return {
    id: row.id,
    name: row.name || "",
    standardEstimateAmount: normalizeAmount(row.standard_estimate_amount),
    defaultDueDays: Number.isFinite(Number(row.default_due_days)) ? Number(row.default_due_days) : null,
    requiredDocuments: row.required_documents || "",
    defaultTasks: row.default_tasks || row.task_list || "",
    taskList: row.task_list || "",
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
      standard_estimate_amount: 120000,
      default_due_days: 30,
      required_documents: "登記事項証明書\n納税証明書\n決算書\n経管資料\n専技資料",
      default_tasks: "要件確認\n必要書類案内\n書類回収\n申請書作成\n提出\n補正対応",
      memo: "建設業許可の標準テンプレート",
    },
    {
      user_id: currentUser.id,
      name: "車庫証明",
      standard_estimate_amount: 55000,
      default_due_days: 7,
      required_documents: "自認書または使用承諾書\n配置図\n所在図\n車検証情報",
      default_tasks: "書類確認\n現地確認\n警察署提出\n受取",
      memo: "車庫証明の標準テンプレート",
    },
    {
      user_id: currentUser.id,
      name: "古物商許可",
      standard_estimate_amount: 95000,
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
  // create table if not exists payments (
  //   id uuid primary key default gen_random_uuid(),
  //   user_id uuid not null,
  //   sale_id uuid not null,
  //   payment_date date,
  //   amount numeric not null,
  //   method text,
  //   memo text,
  //   created_at timestamptz default now()
  // );
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
    reminderCount: Number(row.reminder_count || 0),
    lastReminderDate: row.last_reminder_date || null,
    reminderMethod: row.reminder_method || "",
    reminderMemo: row.reminder_memo || "",
    createdAt: Date.parse(row.created_at) || null,
    updatedAt: Date.parse(row.updated_at) || null,
  };
}

function mapPaymentFromDb(row) {
  return {
    id: row.id,
    userId: row.user_id,
    saleId: row.sale_id,
    paymentDate: row.payment_date || null,
    amount: Number(row.amount || 0),
    method: row.method || "",
    memo: row.memo || "",
    createdAt: row.created_at || null,
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
    caseTaskSubmitBtn,
    workTemplateSubmitBtn,
    estimateSubmitBtn,
    saleSubmitBtn,
    expenseSubmitBtn,
    fixedExpenseSubmitBtn,
    dailyReportSubmitBtn,
    manualReloadBtn,
    csvImportForm?.querySelector('button[type="submit"]'),
    excelImportForm?.querySelector('button[type="submit"]'),
    backupRestoreForm?.querySelector('button[type="submit"]'),
  ].filter(Boolean);
  controls.forEach((control) => {
    control.disabled = !enabled;
  });
}

function startLoading(label = "処理中") {
  loadingCount = 1;

  if (loadingOverlay) {
    loadingOverlay.hidden = false;
    loadingOverlay.style.display = "flex";
    loadingOverlay.setAttribute("aria-busy", "true");
    loadingOverlay.style.pointerEvents = "auto";
  }

  document.body.classList.add("is-loading");
  document.body.setAttribute("aria-busy", "true");
  console.log("LOADING START", label, loadingCount);

}

function forceHideLoading() {
  loadingCount = 0;
  if (loadingTimer) {
    clearTimeout(loadingTimer);
    loadingTimer = null;
  }

  if (loadingOverlay) {
    loadingOverlay.hidden = true;
    loadingOverlay.style.display = "none";
    loadingOverlay.style.pointerEvents = "none";
  }

  document.body.classList.remove("is-loading");
  document.body.style.pointerEvents = "";
  document.body.removeAttribute("aria-busy");
  setSubmitButtonsDisabled(false);
}

async function withLoading(label, fn, options = {}) {
  const { messageTarget = "app", triggerButton = null } = options;
  startLoading(label);
  setSubmitButtonsDisabled(true);
  if (triggerButton instanceof HTMLButtonElement) triggerButton.disabled = true;

  try {
    clearAppMessage();
    if (messageTarget === "auth") showAuthMessage("", false);
    return await fn();
  } catch (error) {
    console.error(`${label} failed`, error);
    const message = `${label}に失敗しました。${error.message || ""}`;
    if (messageTarget === "auth") {
      showAuthMessage(message, true);
    } else {
      showAppMessage(message, true);
    }
    return null;
  } finally {
    setSubmitButtonsDisabled(false);
    if (triggerButton instanceof HTMLButtonElement) triggerButton.disabled = false;
    forceHideLoading();
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
  editState.caseTaskId = null;
  editState.caseDocumentId = null;
}

function clearAppState() {
  currentUser = null;
  state.isInitialDataReady = false;
  state.clients = [];
  state.workTemplates = [];
  state.cases = [];
  state.caseTasks = [];
  state.caseDocuments = [];
  state.estimates = [];
  state.estimateItems = [];
  state.sales = [];
  state.payments = [];
  state.expenses = [];
  state.fixedExpenses = [];
  state.dailyReports = [];
  state.appSettings = { ...DEFAULT_APP_SETTINGS };
  resetViewState();
  resetEditState();
  resetClientForm();
  resetCaseForm();
  resetCaseTaskForm();
  resetCaseDocumentForm();
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
