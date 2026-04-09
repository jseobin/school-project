const elements = {
  entries: document.getElementById("entries"),
  entryCount: document.getElementById("entry-count"),
  healthLabel: document.getElementById("health-label"),
  functionsChip: document.getElementById("functions-chip"),
  databaseChip: document.getElementById("database-chip"),
  form: document.getElementById("entry-form"),
  author: document.getElementById("author"),
  content: document.getElementById("content"),
  submitButton: document.getElementById("submit-button"),
  formStatus: document.getElementById("form-status")
};

const dateFormatter = new Intl.DateTimeFormat(undefined, {
  dateStyle: "medium",
  timeStyle: "short"
});

async function requestJSON(url, init) {
  const response = await fetch(url, init);
  const payload = await response.json().catch(() => ({}));

  if (!response.ok) {
    const message = payload.error || payload.message || `Request failed with ${response.status}`;
    throw new Error(message);
  }

  return payload;
}

function setChip(target, text, tone) {
  target.textContent = text;
  target.className = "status-chip";

  if (tone) {
    target.classList.add(tone);
  }
}

function renderEntries(items) {
  elements.entries.replaceChildren();
  elements.entryCount.textContent = `${items.length} ${items.length === 1 ? "entry" : "entries"}`;

  if (!items.length) {
    const empty = document.createElement("div");
    empty.className = "empty";
    empty.textContent = "No entries yet. Submit the first record to D1.";
    elements.entries.append(empty);
    return;
  }

  for (const item of items) {
    const article = document.createElement("article");
    article.className = "entry";

    const header = document.createElement("header");
    const author = document.createElement("strong");
    const time = document.createElement("time");
    const message = document.createElement("p");

    author.textContent = item.author;
    time.textContent = dateFormatter.format(new Date(item.created_at));
    message.textContent = item.content;

    header.append(author, time);
    article.append(header, message);
    elements.entries.append(article);
  }
}

async function loadHealth() {
  try {
    const payload = await requestJSON("/api/health");

    elements.healthLabel.textContent = `Running in ${payload.runtime}`;
    setChip(elements.functionsChip, "Functions: online", "good");
    setChip(elements.databaseChip, `D1: ${payload.database.status}`, "good");
  } catch (error) {
    elements.healthLabel.textContent = "Edge runtime check failed";
    setChip(elements.functionsChip, "Functions: error", "bad");
    setChip(elements.databaseChip, "D1: unavailable", "bad");
    elements.formStatus.textContent = error.message;
  }
}

async function loadEntries() {
  try {
    const payload = await requestJSON("/api/messages");
    renderEntries(payload.entries || []);
  } catch (error) {
    elements.entries.replaceChildren();

    const empty = document.createElement("div");
    empty.className = "empty";
    empty.textContent = error.message;
    elements.entries.append(empty);
  }
}

async function submitEntry(event) {
  event.preventDefault();

  elements.submitButton.disabled = true;
  elements.formStatus.textContent = "Writing entry to Cloudflare D1...";

  try {
    const payload = await requestJSON("/api/messages", {
      method: "POST",
      headers: {
        "content-type": "application/json"
      },
      body: JSON.stringify({
        author: elements.author.value,
        content: elements.content.value
      })
    });

    elements.author.value = "";
    elements.content.value = "";
    elements.formStatus.textContent = `Saved entry #${payload.entry.id} to D1.`;

    await loadEntries();
    await loadHealth();
  } catch (error) {
    elements.formStatus.textContent = error.message;
  } finally {
    elements.submitButton.disabled = false;
  }
}

elements.form.addEventListener("submit", submitEntry);

loadHealth();
loadEntries();
