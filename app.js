const SUPABASE_URL = "https://ueelzyftlbnvjvpsmpyt.supabase.co";
const SUPABASE_KEY = "sb_publishable_0DrKsieUcCyEZN_HRg8LhQ_QqFTPMtp";
const STATUS_ORDER = ["未着手", "進行中", "完了"];

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
  targetMonthInput?.addEventListener("input", handleTargetMonthChange);
  targetYearInput?.addEventListener("input", handleTargetYearChange);
  aggregationRadios.forEach((radio) => radio.addEventListener("change", handleAggregationChange));
  document.addEventListener("visibilitychange", handleVisibilityChange);
  window.addEventListener("focus", handleWindowFocus);
  window.addEventListener("pageshow", handlePageShow);
  exportCasesCsvBtn?.addEventListener("click", handleExportCasesCsv);
  exportSalesCsvBtn?.addEventListener("click", handleExportSalesCsv);
  exportExpensesCsvBtn?.addEventListener("click", handleExportExpensesCsv);
  exportFixedExpensesCsvBtn?.addEventListener("click", handleExportFixedExpensesCsv);
  exportAllCsvBtn?.addEventListener("click", handleExportAllCsv);
  csvImportForm?.addEventListener("submit", handleCsvImportSubmit);
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
    status: normalizeStatus(entry.status),
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
      status: normalizeStatus(entry.status),
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

  renderYearlyBreakdown(salesByMonth, expenseByMonth);
  renderCaseProfitList({
    mode: state.selectedAggregation,
    monthKey: state.selectedMonth,
    year: state.selectedYear,
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
  const profitsByCaseId = buildCaseProfitMap();

  sorted.forEach((entry) => {
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

  caseEmpty.hidden = sorted.length > 0;
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

function normalizeAmount(raw) {
  if (raw === "" || raw === null || raw === undefined) return null;
  const value = Number(raw);
  if (!Number.isFinite(value) || value < 0) return null;
  return Math.floor(value);
}

function normalizeStatus(status) {
  return STATUS_ORDER.includes(status) ? status : "未着手";
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
  const trimmed = String(value || "").trim();
  if (!trimmed) return null;
  if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) return trimmed;
  if (/^\d{4}\/\d{2}\/\d{2}$/.test(trimmed)) return trimmed.replaceAll("/", "-");
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

async function importCsvByType(importType, csvText) {
  const result = { insertedCount: 0, skippedCount: 0, errorCount: 0 };
  const { headers, rows } = parseCsvToObjects(csvText);
  if (!rows.length) return result;

  if (importType === "cases") {
    validateRequiredHeaders(headers, ["customer_name", "case_name", "estimate_amount", "received_date", "due_date", "status"]);
    const payloads = [];
    rows.forEach((row) => {
      try {
        const customerName = row.customer_name?.trim();
        const caseName = row.case_name?.trim();
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
          status: normalizeStatus(row.status),
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
        const caseName = row.case_name?.trim();
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
        const content = row.content?.trim();
        const amount = parseFlexibleAmount(row.amount);
        if (!date || !content || amount === null) {
          result.errorCount += 1;
          return;
        }
        const caseName = row.case_name?.trim();
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
        const content = row.content?.trim();
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
