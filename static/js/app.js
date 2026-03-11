const resetDataBtn = document.getElementById("resetDataBtn");
const excelUploadForm = document.getElementById("excelUploadForm");
const excelFileInput = document.getElementById("excelFile");
const uploadResult = document.getElementById("uploadResult");

const folderCard = document.getElementById("folderCard");
const folderImportForm = document.getElementById("folderImportForm");
const folderInput = document.getElementById("folderInput");
const folderResult = document.getElementById("folderResult");

// const summaryCard = document.getElementById("summaryCard");
// const summaryContent = document.getElementById("summaryContent");

const imageSearchCard = document.getElementById("imageSearchCard");
const pasteZone = document.getElementById("pasteZone");
const pastedPreview = document.getElementById("pastedPreview");
const searchStatus = document.getElementById("searchStatus");
const matchResults = document.getElementById("matchResults");

const dataCard = document.getElementById("dataCard");
const sheetTabs = document.getElementById("sheetTabs");
const excelDataContainer = document.getElementById("excelDataContainer");

const navButtons = document.querySelectorAll(".sidebar-nav-btn");
const featurePanels = document.querySelectorAll(".feature-panel");

let workbookData = null;

excelUploadForm.addEventListener("submit", handleUploadExcel);
folderImportForm.addEventListener("submit", handleImportFolder);
pasteZone.addEventListener("click", () => pasteZone.focus());
pasteZone.addEventListener("paste", handlePasteImage);
resetDataBtn.addEventListener("click", handleResetAllData);
document.addEventListener("DOMContentLoaded", restoreStateOnReload);

navButtons.forEach((button) => {
  button.addEventListener("click", () => {
    const targetId = button.dataset.target;
    activatePanel(targetId);
  });
});

updateSidebarAvailability();
activatePanel("uploadCard");

async function restoreStateOnReload() {
  try {
    const response = await fetch("/api/state");
    const result = await response.json();

    if (!response.ok || !result.success) {
      return;
    }

    const workbook = result.data?.workbook;
    if (!workbook) {
      return;
    }

    workbookData = workbook;

    renderSheets(workbookData.sheets);

    folderCard.classList.remove("hidden");
    imageSearchCard.classList.remove("hidden");

    updateSidebarAvailability();

    uploadResult.textContent = `Đã khôi phục dữ liệu: ${workbookData.fileName}`;
    folderResult.textContent =
      workbookData.mappedFolderCount > 0
        ? `Đã map ${workbookData.mappedFolderCount} dòng với folder`
        : "Chưa import folder";

    searchStatus.textContent = "Dữ liệu đã được khôi phục. Bạn có thể tiếp tục dán ảnh để tìm kiếm.";

    activatePanel("dataCard");
  } catch (error) {
    console.error("Không khôi phục được state:", error);
  }
}

async function handleResetAllData() {
  const confirmed = window.confirm(
    "Bạn có chắc muốn reset toàn bộ dữ liệu đã import không?"
  );

  if (!confirmed) {
    return;
  }

  try {
    const response = await fetch("/api/reset", {
      method: "POST",
    });

    const result = await response.json();

    if (!response.ok || !result.success) {
      alert(result.message || "Reset thất bại");
      return;
    }

    workbookData = null;

    excelUploadForm.reset();
    folderImportForm.reset();

    uploadResult.textContent = "Chưa có dữ liệu";
    folderResult.textContent = "Chưa import folder";
    searchStatus.textContent = "Chưa dán ảnh";

    folderCard.classList.add("hidden");
    imageSearchCard.classList.add("hidden");
    dataCard.classList.add("hidden");

    sheetTabs.innerHTML = "";
    excelDataContainer.innerHTML = "";
    matchResults.innerHTML = "";

    pastedPreview.classList.add("hidden");
    pastedPreview.removeAttribute("src");

    updateSidebarAvailability();
    activatePanel("uploadCard");

    alert("Đã reset toàn bộ dữ liệu");
  } catch (error) {
    console.error(error);
    alert("Có lỗi khi reset dữ liệu");
  }
}

function updateSidebarAvailability() {
  navButtons.forEach((button) => {
    const targetId = button.dataset.target;
    const panel = document.getElementById(targetId);

    if (!panel) return;

    button.disabled = panel.classList.contains("hidden");
  });
}

function activatePanel(panelId) {
  const targetPanel = document.getElementById(panelId);
  if (!targetPanel || targetPanel.classList.contains("hidden")) {
    return;
  }

  featurePanels.forEach((panel) => panel.classList.add("panel-collapsed"));

  navButtons.forEach((button) => {
    button.classList.toggle("active", button.dataset.target === panelId);
  });

  targetPanel.classList.remove("panel-collapsed");
}

async function handleUploadExcel(event) {
  event.preventDefault();

  const file = excelFileInput.files[0];
  if (!file) {
    uploadResult.textContent = "Bạn chưa chọn file Excel";
    return;
  }

  const formData = new FormData();
  formData.append("file", file);

  uploadResult.textContent = "Đang import file Excel...";
  folderResult.textContent = "Chưa import folder";
  searchStatus.textContent = "Chưa dán ảnh";

  workbookData = null;
  folderCard.classList.add("hidden");
  imageSearchCard.classList.add("hidden");
  dataCard.classList.add("hidden");
  // summaryCard.classList.add("hidden");

  // summaryContent.innerHTML = "";
  sheetTabs.innerHTML = "";
  excelDataContainer.innerHTML = "";
  matchResults.innerHTML = "";

  pastedPreview.classList.add("hidden");
  pastedPreview.removeAttribute("src");

  updateSidebarAvailability();
  activatePanel("uploadCard");

  try {
    const response = await fetch("/api/excel/upload", {
      method: "POST",
      body: formData,
    });

    const result = await response.json();

    if (!response.ok || !result.success) {
      uploadResult.textContent = result.message || "Import Excel thất bại";
      return;
    }

    workbookData = result.data;
    uploadResult.textContent = result.message;

    // renderSummary(workbookData);
    renderSheets(workbookData.sheets);

    folderCard.classList.remove("hidden");
    imageSearchCard.classList.remove("hidden");
    updateSidebarAvailability();
    activatePanel("folderCard");

    searchStatus.textContent = "Excel đã import xong. Bạn có thể dán ảnh hoặc import folder để map chi tiết.";
  } catch (error) {
    console.error(error);
    uploadResult.textContent = "Có lỗi khi upload file Excel";
  }
}

async function handleImportFolder(event) {
  event.preventDefault();

  if (!workbookData) {
    folderResult.textContent = "Bạn cần import Excel trước";
    return;
  }

  const files = Array.from(folderInput.files || []);
  if (files.length === 0) {
    folderResult.textContent = "Bạn chưa chọn folder";
    return;
  }

  const rootFolderName =
    files[0]?.webkitRelativePath?.split("/")[0] || "Pattern Manual";

  const formData = new FormData();
  formData.append("rootFolderName", rootFolderName);

  files.forEach((file) => {
    formData.append("files", file);
    formData.append("relativePaths", file.webkitRelativePath || file.name);
  });

  folderResult.textContent = "Đang import folder và map với Style No...";

  try {
    const response = await fetch("/api/folder/import", {
      method: "POST",
      body: formData,
    });

    const result = await response.json();

    if (!response.ok || !result.success) {
      folderResult.textContent = result.message || "Import folder thất bại";
      return;
    }

    workbookData = result.data.workbook;

    // renderSummary(workbookData);
    renderSheets(workbookData.sheets);

    folderResult.textContent =
      `Import folder thành công. Tổng folder con: ${result.data.folderCount}. ` +
      `Map được: ${result.data.mappedCount} dòng.`;

    updateSidebarAvailability();
    activatePanel("dataCard");
  } catch (error) {
    console.error(error);
    folderResult.textContent = "Có lỗi khi import folder";
  }
}

function handlePasteImage(event) {
  const items = Array.from(event.clipboardData?.items || []);
  const imageItem = items.find((item) => item.type.startsWith("image/"));

  if (!imageItem) {
    searchStatus.textContent = "Clipboard hiện không có ảnh";
    return;
  }

  event.preventDefault();

  const imageFile = imageItem.getAsFile();
  if (!imageFile) {
    searchStatus.textContent = "Không đọc được ảnh từ clipboard";
    return;
  }

  const previewUrl = URL.createObjectURL(imageFile);
  pastedPreview.src = previewUrl;
  pastedPreview.classList.remove("hidden");

  searchByImage(imageFile);
}

async function searchByImage(imageFile) {
  const formData = new FormData();
  formData.append("image", imageFile, "pasted-image.png");

  searchStatus.textContent = "Đang so khớp ảnh...";
  matchResults.innerHTML = "";

  try {
    const response = await fetch("/api/search-by-image", {
      method: "POST",
      body: formData,
    });

    const result = await response.json();

    if (!response.ok || !result.success) {
      searchStatus.textContent = result.message || "Tìm ảnh thất bại";
      return;
    }

    searchStatus.textContent = `Tìm thấy ${result.data.results.length} hàng gần giống nhất`;
    renderMatchResults(result.data.results);
    activatePanel("imageSearchCard");
  } catch (error) {
    console.error(error);
    searchStatus.textContent = "Có lỗi khi tìm ảnh tương tự";
  }
}

function renderMatchResults(results) {
  if (!Array.isArray(results) || results.length === 0) {
    matchResults.innerHTML = `<p class="empty-state">Không tìm thấy kết quả phù hợp</p>`;
    return;
  }

  matchResults.innerHTML = results
    .map((item) => {
      const row = item.row || {};
      const title = row["Style No"] || row["Style Name"] || "Kết quả tương tự";

      return `
        <div class="match-card">
          <div class="match-card-image">
            <img src="${item.matchedImage}" alt="Ảnh khớp" />
          </div>

          <div class="match-card-body">
            <div class="match-card-head">
              <h3>${escapeHtml(title)}</h3>
              <span class="score-badge">${escapeHtml(item.scorePercent)}%</span>
            </div>

            <p class="match-meta">
              Sheet: <strong>${escapeHtml(item.sheetName)}</strong>
              | Dòng Excel: <strong>${escapeHtml(item.excelRow)}</strong>
            </p>

            <div class="match-fields">
              ${renderField("Style No", row["Style No"])}
              ${renderField("Style Name", row["Style Name"])}
              ${renderField("Customer", row["Customer"])}
              ${renderField("Season", row["Season"])}
              ${renderField("Staff", row["Staff"])}
              ${renderField("Product", row["Product"])}
              ${renderField("Categories", row["Categories"])}
              ${renderField("Folder", item.folderName || "-")}
            </div>

            <div class="match-link-row">
              ${
                item.detailUrl
                  ? `<a class="folder-link-btn" href="${escapeAttribute(item.detailUrl)}" target="_blank" rel="noopener noreferrer">Chi tiết</a>`
                  : `<span class="muted-text">Chưa map folder</span>`
              }
            </div>
          </div>
        </div>
      `;
    })
    .join("");
}

function renderField(label, value) {
  const safeValue =
    value === null || value === undefined || value === ""
      ? "-"
      : value;

  return `
    <div class="field-chip">
      <span>${escapeHtml(label)}</span>
      <strong>${escapeHtml(safeValue)}</strong>
    </div>
  `;
}

// function renderSummary(data) {
//   // summaryCard.classList.remove("hidden");

//   const items = [
//     { label: "Tên file", value: data.fileName },
//     { label: "Số sheet", value: data.sheetCount },
//     { label: "Tổng số dòng dữ liệu", value: data.totalRows },
//     { label: "Số ảnh index để tìm kiếm", value: data.imageIndexCount || 0 },
//     { label: "Số row đã map folder", value: data.mappedFolderCount || 0 },
//     { label: "Folder đã import", value: data.folderImportName || "-" },
//   ];

//   summaryContent.innerHTML = items
//     .map(
//       (item) => `
//         <div class="summary-item">
//           <span>${escapeHtml(item.label)}</span>
//           <strong>${escapeHtml(item.value)}</strong>
//         </div>
//       `
//     )
//     .join("");
// }

function renderSheets(sheets) {
  if (!Array.isArray(sheets) || sheets.length === 0) {
    excelDataContainer.innerHTML = "<p>Không có sheet nào để hiển thị</p>";
    dataCard.classList.remove("hidden");
    updateSidebarAvailability();
    return;
  }

  dataCard.classList.remove("hidden");
  updateSidebarAvailability();

  sheetTabs.innerHTML = sheets
    .map(
      (sheet, index) => `
        <button
          type="button"
          class="tab-button ${index === 0 ? "active" : ""}"
          data-sheet-target="sheet-${index}"
        >
          ${escapeHtml(sheet.sheetName)}
        </button>
      `
    )
    .join("");

  excelDataContainer.innerHTML = sheets
    .map((sheet, index) => createSheetPanel(sheet, index === 0, index))
    .join("");

  bindTabEvents();
}

function createSheetPanel(sheet, isActive, index) {
  const headers = [...sheet.headers, "Chi tiết"];

  const tableHead = headers
    .map((header) => `<th>${escapeHtml(header)}</th>`)
    .join("");

  const tableBody = sheet.rows
    .map((row) => {
      const cells = headers.map((header) => createTableCell(header, row)).join("");
      return `<tr>${cells}</tr>`;
    })
    .join("");

  return `
    <div class="sheet-panel ${isActive ? "active" : ""}" id="sheet-${index}">
      <p class="sheet-meta">
        Sheet: <strong>${escapeHtml(sheet.sheetName)}</strong>
        | Số dòng: <strong>${escapeHtml(sheet.rowCount)}</strong>
      </p>

      <div class="table-wrapper">
        <table class="data-table">
          <thead>
            <tr>${tableHead}</tr>
          </thead>
          <tbody>
            ${tableBody}
          </tbody>
        </table>
      </div>
    </div>
  `;
}

function createTableCell(header, row) {
  if (header === "Sketch Design") {
    const images = Array.isArray(row.__images) ? row.__images : [];

    if (images.length === 0) {
      return `<td class="image-cell"><span class="empty-value">Không có ảnh</span></td>`;
    }

    const imageHtml = images
      .map((src, index) => `<img src="${src}" alt="Sketch ${index + 1}" />`)
      .join("");

    return `<td class="image-cell"><div class="image-list">${imageHtml}</div></td>`;
  }

  if (header === "Chi tiết") {
    if (!row.__detailUrl) {
      return `<td><span class="empty-value">Chưa map</span></td>`;
    }

    return `
      <td>
        <a class="table-folder-link" href="${escapeAttribute(row.__detailUrl)}" target="_blank" rel="noopener noreferrer">
          ${escapeHtml(row.__folderName || "Chi tiết")}
        </a>
      </td>
    `;
  }

  const value = row[header];

  if (value === null || value === undefined || value === "") {
    return `<td><span class="empty-value">-</span></td>`;
  }

  return `<td>${escapeHtml(value)}</td>`;
}

function bindTabEvents() {
  const buttons = document.querySelectorAll(".tab-button");
  const panels = document.querySelectorAll(".sheet-panel");

  buttons.forEach((button) => {
    button.addEventListener("click", () => {
      const targetId = button.dataset.sheetTarget;

      buttons.forEach((item) => item.classList.remove("active"));
      panels.forEach((panel) => panel.classList.remove("active"));

      button.classList.add("active");
      document.getElementById(targetId)?.classList.add("active");
    });
  });
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function escapeAttribute(value) {
  return String(value ?? "").replace(/"/g, "&quot;");
}