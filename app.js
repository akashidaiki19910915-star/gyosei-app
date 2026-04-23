const STORAGE_KEY = "gyosei-cases";
const STATUS_ORDER = ["未着手", "進行中", "完了"];

const form = document.getElementById("case-form");
const caseList = document.getElementById("case-list");
const emptyState = document.getElementById("empty-state");
const clearBtn = document.getElementById("clear-btn");
const csvBtn = document.getElementById("csv-btn");
const caseTemplate = document.getElementById("case-template");

const customerNameInput = document.getElementById("customer-name");
const caseNameInput = document.getElementById("case-name");
const amountInput = document.getElementById("amount");
const receivedDateInput = document.getElementById("received-date");
const dueDateInput = document.getElementById("due-date");
const statusInput = document.getElementById("status");

let cases = loadCases();

render();

form.addEventListener("submit", (event) => {
  event.preventDefault();

  const customerName = customerNameInput.value.trim();
  const caseName = caseNameInput.value.trim();

  if (!customerName || !caseName) return;

  cases.push({
    id: crypto.randomUUID(),
    customerName,
    caseName,
    amount: normalizeAmount(amountInput.value),
    receivedDate: receivedDateInput.value || "",
    dueDate: dueDateInput.value || "",
    status: normalizeStatus(statusInput.value),
    createdAt: Date.now(),
  });

  persist();
  render();
  form.reset();
  statusInput.value = "未着手";
  customerNameInput.focus();
});

caseList.addEventListener("click", (event) => {
  const target = event.target;
  if (!(target instanceof HTMLButtonElement)) return;

  const item = target.closest(".case-item");
  if (!item) return;

  const id = item.dataset.id;

  if (target.classList.contains("delete-btn")) {
    cases = cases.filter((entry) => entry.id !== id);
    persist();
    render();
    return;
  }

  if (target.classList.contains("edit-btn")) {
    toggleEditMode(item, true);
    return;
  }

  if (target.classList.contains("cancel-btn")) {
    toggleEditMode(item, false);
  }
});

caseList.addEventListener("submit", (event) => {
  const target = event.target;
  if (!(target instanceof HTMLFormElement) || !target.classList.contains("edit-form")) return;

  event.preventDefault();

  const item = target.closest(".case-item");
  if (!item) return;

  const id = item.dataset.id;

  const customerName = target.elements.customerName.value.trim();
  const caseName = target.elements.caseName.value.trim();

  if (!customerName || !caseName) return;

  cases = cases.map((entry) => {
    if (entry.id !== id) return entry;

    return {
      ...entry,
      customerName,
      caseName,
      amount: normalizeAmount(target.elements.amount.value),
      receivedDate: target.elements.receivedDate.value || "",
      dueDate: target.elements.dueDate.value || "",
      status: normalizeStatus(target.elements.status.value),
    };
  });

  persist();
  render();
});

clearBtn.addEventListener("click", () => {
  cases = [];
  persist();
  render();
});

csvBtn.addEventListener("click", () => {
  const rows = [
    ["顧客名", "案件名", "金額", "受付日", "期限日", "ステータス", "登録日時"],
    ...sortCases(cases).map((entry) => [
      entry.customerName,
      entry.caseName,
      entry.amount === null ? "" : String(entry.amount),
      entry.receivedDate,
      entry.dueDate,
      entry.status,
      formatDateTime(entry.createdAt),
    ]),
  ];

  const csvContent = rows.map((row) => row.map(escapeCsvCell).join(",")).join("\r\n");
  const blob = new Blob(["\uFEFF" + csvContent], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);

  const a = document.createElement("a");
  const today = new Date().toISOString().slice(0, 10);
  a.href = url;
  a.download = `案件一覧_${today}.csv`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
});

function render() {
  caseList.innerHTML = "";

  const sorted = sortCases(cases);

  sorted.forEach((entry) => {
    const node = caseTemplate.content.cloneNode(true);
    const item = node.querySelector(".case-item");
    const customerName = node.querySelector(".customer-name");
    const caseName = node.querySelector(".case-name");
    const amount = node.querySelector(".amount");
    const receivedDate = node.querySelector(".received-date");
    const dueDate = node.querySelector(".due-date");
    const status = node.querySelector(".status-badge");
    const editForm = node.querySelector(".edit-form");

    item.dataset.id = entry.id;
    customerName.textContent = entry.customerName;
    caseName.textContent = entry.caseName;
    amount.textContent = `金額: ${formatCurrency(entry.amount)}`;
    receivedDate.textContent = `受付日: ${formatDate(entry.receivedDate)}`;
    dueDate.textContent = `期限日: ${formatDate(entry.dueDate)}`;
    status.textContent = entry.status;
    status.dataset.status = entry.status;

    editForm.elements.customerName.value = entry.customerName;
    editForm.elements.caseName.value = entry.caseName;
    editForm.elements.amount.value = entry.amount === null ? "" : String(entry.amount);
    editForm.elements.receivedDate.value = entry.receivedDate;
    editForm.elements.dueDate.value = entry.dueDate;
    editForm.elements.status.value = entry.status;

    caseList.appendChild(node);
  });

  emptyState.hidden = sorted.length > 0;
}

function toggleEditMode(item, enabled) {
  const main = item.querySelector(".case-main");
  const actions = item.querySelector(".item-actions");
  const form = item.querySelector(".edit-form");

  main.hidden = enabled;
  actions.hidden = enabled;
  form.hidden = !enabled;
}

function sortCases(data) {
  return [...data].sort((a, b) => {
    const dueA = toSortTimestamp(a.dueDate);
    const dueB = toSortTimestamp(b.dueDate);

    if (dueA !== dueB) return dueA - dueB;

    const statusA = STATUS_ORDER.indexOf(normalizeStatus(a.status));
    const statusB = STATUS_ORDER.indexOf(normalizeStatus(b.status));

    if (statusA !== statusB) return statusA - statusB;

    return a.createdAt - b.createdAt;
  });
}

function toSortTimestamp(dateText) {
  if (!dateText) return Number.MAX_SAFE_INTEGER;
  const t = Date.parse(dateText);
  return Number.isNaN(t) ? Number.MAX_SAFE_INTEGER : t;
}

function formatCurrency(value) {
  if (typeof value !== "number" || Number.isNaN(value)) return "未入力";
  return new Intl.NumberFormat("ja-JP").format(value) + " 円";
}

function formatDate(dateText) {
  if (!dateText) return "未設定";
  const date = new Date(dateText);
  if (Number.isNaN(date.getTime())) return "未設定";
  return new Intl.DateTimeFormat("ja-JP").format(date);
}

function formatDateTime(timestamp) {
  return new Intl.DateTimeFormat("ja-JP", {
    dateStyle: "short",
    timeStyle: "short",
  }).format(new Date(timestamp));
}

function normalizeAmount(rawValue) {
  if (rawValue === "" || rawValue === null || rawValue === undefined) return null;
  const number = Number(rawValue);
  if (!Number.isFinite(number) || number < 0) return null;
  return Math.floor(number);
}

function normalizeStatus(status) {
  return STATUS_ORDER.includes(status) ? status : "未着手";
}

function escapeCsvCell(value) {
  const str = String(value ?? "");
  if (!/[",\r\n]/.test(str)) return str;
  return `"${str.replaceAll('"', '""')}"`;
}

function persist() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(cases));
}

function loadCases() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    const parsed = raw ? JSON.parse(raw) : [];

    if (!Array.isArray(parsed)) return [];

    return parsed
      .filter((entry) => entry && typeof entry === "object")
      .map((entry) => ({
        id: entry.id || crypto.randomUUID(),
        customerName: String(entry.customerName || "").trim(),
        caseName: String(entry.caseName || "").trim(),
        amount: normalizeAmount(entry.amount),
        receivedDate: String(entry.receivedDate || ""),
        dueDate: String(entry.dueDate || ""),
        status: normalizeStatus(entry.status),
        createdAt: Number.isFinite(entry.createdAt) ? entry.createdAt : Date.now(),
      }))
      .filter((entry) => entry.customerName && entry.caseName);
  } catch {
    return [];
  }
}
