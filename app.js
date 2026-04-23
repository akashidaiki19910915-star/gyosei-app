const STORAGE_KEY = "admin-services-requests";

const form = document.getElementById("request-form");
const requesterInput = document.getElementById("requester");
const serviceInput = document.getElementById("service");
const notesInput = document.getElementById("notes");
const requestList = document.getElementById("request-list");
const emptyState = document.getElementById("empty-state");
const clearBtn = document.getElementById("clear-btn");
const requestTemplate = document.getElementById("request-template");

let requests = loadRequests();

render();

form.addEventListener("submit", (event) => {
  event.preventDefault();

  const requester = requesterInput.value.trim();
  const service = serviceInput.value;
  const notes = notesInput.value.trim();

  if (!requester || !service) return;

  requests.unshift({
    id: crypto.randomUUID(),
    requester,
    service,
    notes,
    createdAt: Date.now(),
  });

  persist();
  render();
  form.reset();
  requesterInput.focus();
});

requestList.addEventListener("click", (event) => {
  const target = event.target;
  if (!(target instanceof HTMLButtonElement)) return;

  if (!target.classList.contains("done-btn")) return;

  const item = target.closest(".request-item");
  if (!item) return;

  requests = requests.filter((request) => request.id !== item.dataset.id);
  persist();
  render();
});

clearBtn.addEventListener("click", () => {
  requests = [];
  persist();
  render();
});

function render() {
  requestList.innerHTML = "";

  requests.forEach((request) => {
    const node = requestTemplate.content.cloneNode(true);
    const item = node.querySelector(".request-item");
    const requester = node.querySelector(".requester");
    const service = node.querySelector(".service");
    const created = node.querySelector(".created");
    const notes = node.querySelector(".notes");

    item.dataset.id = request.id;
    requester.textContent = request.requester;
    service.textContent = request.service;
    created.textContent = new Date(request.createdAt).toLocaleString();
    notes.textContent = request.notes || "No notes provided.";

    requestList.appendChild(node);
  });

  emptyState.hidden = requests.length > 0;
}

function persist() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(requests));
}

function loadRequests() {
  try {
    const value = localStorage.getItem(STORAGE_KEY);
    const parsed = value ? JSON.parse(value) : [];
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}
