// app.js
// Basit Excel okuma ve tabloya yazdırma mantığı
// HTML'deki id'lere dikkat: excelFileInput, parseExcelBtn, fileNameInfo, messageBox, rowCountBadge, emptyState, excelTableWrapper

let selectedFile = null;

// DOM referansları
const fileInput = document.getElementById("excelFileInput");
const parseBtn = document.getElementById("parseExcelBtn");
const fileNameInfo = document.getElementById("fileNameInfo");
const messageBox = document.getElementById("messageBox");
const rowCountBadge = document.getElementById("rowCountBadge");
const emptyState = document.getElementById("emptyState");
const excelTableWrapper = document.getElementById("excelTableWrapper");

// Mesaj gösterme fonksiyonu
function showMessage(text, type = "info") {
  if (!messageBox) return;
  messageBox.classList.remove("hidden");

  let baseClasses =
    "mt-2 rounded-lg px-3 py-2 text-xs border ";
  let typeClasses = "";

  switch (type) {
    case "error":
      typeClasses =
        "bg-red-900/40 border-red-500/60 text-red-200";
      break;
    case "success":
      typeClasses =
        "bg-emerald-900/40 border-emerald-500/60 text-emerald-200";
      break;
    default:
      typeClasses =
        "bg-slate-800/70 border-slate-500/60 text-slate-200";
  }

  messageBox.className = baseClasses + typeClasses;
  messageBox.textContent = text;
}

// Satır sayısını güncelle
function updateRowCount(count) {
  if (!rowCountBadge) return;
  rowCountBadge.textContent = `Toplam Satır: ${count}`;
}

// Excel'den gelen JSON verisini tabloya çevir
function renderTableFromJson(jsonData) {
  if (!excelTableWrapper) return;

  // Eğer veri yoksa eski tabloyu temizle, empty state göster
  if (!jsonData || jsonData.length === 0) {
    excelTableWrapper.innerHTML = "";
    excelTableWrapper.classList.add("hidden");
    emptyState.classList.remove("hidden");
    updateRowCount(0);
    return;
  }

  // Header'ları JSON'un ilk satırının key'lerinden al
  const headers = Object.keys(jsonData[0]);

  // <table> oluştur
  const table = document.createElement("table");
  table.className =
    "min-w-full text-xs text-left text-slate-200 border-collapse";

  // THEAD
  const thead = document.createElement("thead");
  thead.className = "bg-slate-800 sticky top-0 z-10";

  const headerRow = document.createElement("tr");
  headerRow.className = "border-b border-slate-700";

  headers.forEach((header) => {
    const th = document.createElement("th");
    th.className =
      "px-3 py-2 font-semibold uppercase tracking-wide text-[0.7rem] text-slate-300 border-r border-slate-700 last:border-r-0";
    th.textContent = header;
    headerRow.appendChild(th);
  });

  thead.appendChild(headerRow);
  table.appendChild(thead);

  // TBODY
  const tbody = document.createElement("tbody");
  tbody.className = "divide-y divide-slate-800";

  jsonData.forEach((row, rowIndex) => {
    const tr = document.createElement("tr");
    tr.className =
      rowIndex % 2 === 0
        ? "bg-slate-900/40 hover:bg-slate-800/70"
        : "bg-slate-900/20 hover:bg-slate-800/70";

    headers.forEach((header) => {
      const td = document.createElement("td");
      td.className =
        "px-3 py-1.5 align-middle border-r border-slate-800 last:border-r-0";
      let cellValue = row[header];

      // undefined/null ise boş string yap
      if (cellValue === undefined || cellValue === null) {
        cellValue = "";
      }

      td.textContent = cellValue;
      tr.appendChild(td);
    });

    tbody.appendChild(tr);
  });

  table.appendChild(tbody);

  // Wrapper'a yerleştir
  excelTableWrapper.innerHTML = "";
  excelTableWrapper.appendChild(table);

  // Görünürlükleri ayarla
  emptyState.classList.add("hidden");
  excelTableWrapper.classList.remove("hidden");

  // Satır sayısını güncelle
  updateRowCount(jsonData.length);
}

// Excel dosyasını parse et
function parseExcelFile(file) {
  if (!file) {
    showMessage("Lütfen önce bir Excel dosyası seç.", "error");
    return;
  }

  // Boyut kontrolü (opsiyonel, 10MB üstüne uyarı)
  const maxSizeMB = 10;
  if (file.size > maxSizeMB * 1024 * 1024) {
    showMessage(
      `Dosya çok büyük görünüyor (${maxSizeMB}MB üzeri). Lütfen daha küçük bir dosya deneyin.`,
      "error"
    );
    return;
  }

  const reader = new FileReader();

  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });

      // İlk sayfayı kullan
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];

      // Sayfayı JSON'a çevir
      const jsonData = XLSX.utils.sheet_to_json(worksheet, {
        defval: "",
        raw: false,
      });

      if (!jsonData || jsonData.length === 0) {
        showMessage(
          "Excel dosyasında veri bulunamadı. Lütfen içeriği kontrol edin.",
          "error"
        );
        renderTableFromJson([]);
        return;
      }

      // Tabloyu oluştur
      renderTableFromJson(jsonData);
      showMessage(
        `Excel başarıyla okundu. Toplam ${jsonData.length} satır yüklendi.`,
        "success"
      );
    } catch (err) {
      console.error("Excel parse hatası:", err);
      showMessage(
        "Excel dosyası okunurken bir hata oluştu. Formatı kontrol edin.",
        "error"
      );
      renderTableFromJson([]);
    }
  };

  reader.onerror = (e) => {
    console.error("FileReader hatası:", e);
    showMessage(
      "Dosya okunurken bir hata oluştu. Lütfen tekrar deneyin.",
      "error"
    );
  };

  // Dosyayı ArrayBuffer olarak oku
  reader.readAsArrayBuffer(file);
}

// Dosya input değiştiğinde
if (fileInput) {
  fileInput.addEventListener("change", (event) => {
    const file = event.target.files[0];

    if (!file) {
      selectedFile = null;
      fileNameInfo.textContent = "Henüz dosya seçilmedi.";
      showMessage("Herhangi bir dosya seçilmedi.", "info");
      renderTableFromJson([]);
      return;
    }

    selectedFile = file;
    fileNameInfo.textContent = `Seçilen dosya: ${file.name}`;
    showMessage(
      "Dosya seçildi. Şimdi 'Excel’i Oku ve Listele' butonuna basabilirsin.",
      "info"
    );
  });
}

// Butona tıklanınca Excel'i işle
if (parseBtn) {
  parseBtn.addEventListener("click", () => {
    parseExcelFile(selectedFile);
  });
}

// (Opsiyonel) Sürükle-Bırak desteği: label alanı üzerinden
const dropZoneLabel = document.querySelector("label[for='excelFileInput']");
if (dropZoneLabel && fileInput) {
  ["dragenter", "dragover"].forEach((eventName) => {
    dropZoneLabel.addEventListener(
      eventName,
      (e) => {
        e.preventDefault();
        e.stopPropagation();
        dropZoneLabel.classList.add("border-sky-500", "bg-slate-900/80");
      },
      false
    );
  });

  ["dragleave", "drop"].forEach((eventName) => {
    dropZoneLabel.addEventListener(
      eventName,
      (e) => {
        e.preventDefault();
        e.stopPropagation();
        dropZoneLabel.classList.remove("border-sky-500", "bg-slate-900/80");
      },
      false
    );
  });

  dropZoneLabel.addEventListener(
    "drop",
    (e) => {
      const dt = e.dataTransfer;
      const files = dt.files;
      if (files && files[0]) {
        fileInput.files = files; // input'a da set et
        const event = new Event("change");
        fileInput.dispatchEvent(event);
      }
    },
    false
  );
}

// Sayfa ilk açıldığında başlangıç durumu
renderTableFromJson([]);
showMessage(
  "Başlamak için sol taraftan bir Excel dosyası seç ve ardından 'Excel’i Oku ve Listele' butonuna tıkla.",
  "info"
);
