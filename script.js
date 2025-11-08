 const $ = (sel) => document.querySelector(sel);

const micBtn = $("#micBtn");
const manualInput = $("#manualInput");
const addBtn = $("#addBtn");
const filterInput = $("#filterInput");
const clearFilter = $("#clearFilter");
const mainTableBody = $("#mainTable tbody");
const filteredTableBody = $("#filteredTable tbody");
const grandTotal = $("#grandTotal");
const filteredTotal = $("#filteredTotal");
const flatTotals = $("#flatTotals");
const exportMainBtn = $("#exportMain");
const exportFilteredBtn = $("#exportFiltered");

let data = JSON.parse(localStorage.getItem("whisperData") || "[]");
let filter = "";

// 🧩 Utility
const saveData = () =>
  localStorage.setItem("whisperData", JSON.stringify(data));

function parseInput(text) {
  const parts = text.split("-").map((p) => p.trim());
  if (parts.length < 4) return null;
  return {
    date: parts[0],
    flat: parts[1],
    item: parts[2],
    amount: parseFloat(parts[3]) || 0,
    id: Date.now() + Math.random(),
  };
}

function renderTables() {
  mainTableBody.innerHTML = "";
  filteredTableBody.innerHTML = "";

  const filtered = filter
    ? data.filter((r) =>
        r.flat.toLowerCase().includes(filter.toLowerCase().trim())
      )
    : data;

  data.forEach((row) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td contenteditable="true">${row.date}</td>
      <td contenteditable="true">${row.flat}</td>
      <td contenteditable="true">${row.item}</td>
      <td contenteditable="true">${row.amount}</td>
      <td><button class="del">❌</button></td>
    `;
    tr.querySelector(".del").onclick = () => {
      data = data.filter((r) => r.id !== row.id);
      saveData();
      renderTables();
    };
    tr.querySelectorAll("[contenteditable]").forEach((td, i) => {
      td.onblur = () => {
        const keys = ["date", "flat", "item", "amount"];
        row[keys[i]] =
          i === 3 ? parseFloat(td.innerText.trim()) || 0 : td.innerText.trim();
        saveData();
        renderTables();
      };
    });
    mainTableBody.appendChild(tr);
  });

  filtered.forEach((row) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${row.date}</td><td>${row.flat}</td><td>${row.item}</td><td>${row.amount}</td><td></td>
    `;
    filteredTableBody.appendChild(tr);
  });

  const total = data.reduce((s, r) => s + Number(r.amount || 0), 0);
  grandTotal.innerText = total.toFixed(2);

  const ftotal = filtered.reduce((s, r) => s + Number(r.amount || 0), 0);
  filteredTotal.innerText = ftotal.toFixed(2);

  const flats = {};
  data.forEach((r) => {
    flats[r.flat] = (flats[r.flat] || 0) + Number(r.amount);
  });
  flatTotals.innerHTML = Object.entries(flats)
    .map(([f, sum]) => `<div><b>${f}:</b> ${sum.toFixed(2)}</div>`)
    .join("");

  saveData();
}

function addEntry(input) {
  const parsed = parseInput(input);
  if (!parsed) {
    alert("Format: date - flat - item - amount");
    return;
  }
  data.push(parsed);
  saveData();
  renderTables();
}

addBtn.onclick = () => {
  if (manualInput.value.trim()) addEntry(manualInput.value.trim());
  manualInput.value = "";
};

filterInput.oninput = (e) => {
  filter = e.target.value;
  renderTables();
};

clearFilter.onclick = () => {
  filter = "";
  filterInput.value = "";
  renderTables();
};

// 🎙️ Voice input
let recognition;
function setupSpeech() {
  const SpeechRecognition =
    window.SpeechRecognition || window.webkitSpeechRecognition;
  if (!SpeechRecognition) {
    micBtn.disabled = true;
    micBtn.textContent = "🎙️ Not supported";
    return;
  }

  recognition = new SpeechRecognition();
  recognition.lang = "en-US";
  recognition.interimResults = false;

  recognition.onstart = () => (micBtn.textContent = "🎧 Listening...");
  recognition.onend = () => (micBtn.textContent = "🎙️ Start Listening");
  recognition.onerror = () => alert("Speech recognition error!");

  recognition.onresult = (e) => {
    const spoken = e.results[0][0].transcript;
    manualInput.value = spoken;
    addEntry(spoken);
  };
}

micBtn.onclick = () => {
  if (!recognition) setupSpeech();
  recognition && recognition.start();
};

// 📄 Export DOCX
async function exportDoc(rows, filename) {
  const { Document, Packer, Paragraph, Table, TableRow, TableCell } =
    window.docx;

  const tableRows = [
    new TableRow({
      children: ["Date", "Flat", "Item", "Amount"].map(
        (h) => new TableCell({ children: [new Paragraph(h)] })
      ),
    }),
    ...rows.map(
      (r) =>
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph(r.date)] }),
            new TableCell({ children: [new Paragraph(r.flat)] }),
            new TableCell({ children: [new Paragraph(r.item)] }),
            new TableCell({ children: [new Paragraph(String(r.amount))] }),
          ],
        })
    ),
  ];

  const doc = new Document({
    sections: [{ children: [new Table({ rows: tableRows })] }],
  });

  const blob = await Packer.toBlob(doc);
  saveAs(blob, filename);
}

exportMainBtn.onclick = () => exportDoc(data, "main_table.docx");
exportFilteredBtn.onclick = () => {
  const filtered = filter
    ? data.filter((r) =>
        r.flat.toLowerCase().includes(filter.toLowerCase().trim())
      )
    : data;
  exportDoc(filtered, "filtered_table.docx");
};

// Initialize
setupSpeech();
renderTables();
