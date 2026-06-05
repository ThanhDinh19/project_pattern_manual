const basePath = document.body?.dataset.basePath || "";

const navExcelBtn = document.getElementById("navExcelBtn");
const navFolderBtn = document.getElementById("navFolderBtn");
const navImageBtn = document.getElementById("navImageBtn");
const navDataBtn = document.getElementById("navDataBtn");
const navResetScopeBtn = document.getElementById("navResetScopeBtn");
const navAdminBtn = document.getElementById("navAdminBtn");
const logoutBtn = document.getElementById("logoutBtn");
const currentUserLabel = document.getElementById("currentUserLabel");

const createUserForm = document.getElementById("createUserForm");
const createUserStatus = document.getElementById("createUserStatus");
const usersStatus = document.getElementById("usersStatus");
const usersTableWrap = document.getElementById("usersTableWrap");
const adminCard = document.getElementById("adminCard");

const openCreateUserModalBtn = document.getElementById("openCreateUserModalBtn");

const createUserModal = document.getElementById("createUserModal");
const createUserModalCloseBtn = document.getElementById("createUserModalCloseBtn");
const cancelCreateUserBtn = document.getElementById("cancelCreateUserBtn");

const editUserModal = document.getElementById("editUserModal");
const editUserModalCloseBtn = document.getElementById("editUserModalCloseBtn");
const cancelEditUserBtn = document.getElementById("cancelEditUserBtn");
const saveEditUserBtn = document.getElementById("saveEditUserBtn");
const resetUserPasswordBtn = document.getElementById("resetUserPasswordBtn");
const editUserStatus = document.getElementById("editUserStatus");

const editUserId = document.getElementById("editUserId");
const editUserEmail = document.getElementById("editUserEmail");
const editFullName = document.getElementById("editFullName");
const editIsActive = document.getElementById("editIsActive");
const editIsAdmin = document.getElementById("editIsAdmin");
const editCanImportExcel = document.getElementById("editCanImportExcel");
const editCanImportFolder = document.getElementById("editCanImportFolder");
const editCanSearchImage = document.getElementById("editCanSearchImage");
const editCanViewData = document.getElementById("editCanViewData");
const editCanResetData = document.getElementById("editCanResetData");
const editCanManageUsers = document.getElementById("editCanManageUsers");
const editNewPassword = document.getElementById("editNewPassword");
const editUserModalTitle = document.getElementById("editUserModalTitle");
const editUserModalSubTitle = document.getElementById("editUserModalSubTitle");

const adminUserSearch = document.getElementById("adminUserSearch");
const adminReloadBtn = document.getElementById("adminReloadBtn");

const navAuditBtn = document.getElementById("navAuditBtn");
const auditCard = document.getElementById("auditCard");
const reloadAuditLogsBtn = document.getElementById("reloadAuditLogsBtn");
const auditLogsTableBody = document.getElementById("auditLogsTableBody");
const editCanViewAuditLogs = document.getElementById("editCanViewAuditLogs");
const permViewAuditLogs = document.getElementById("permViewAuditLogs");

function formatAuditDate(value) {
  if (!value) return "-";

  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;

  const dd = String(date.getDate()).padStart(2, "0");
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const yyyy = date.getFullYear();
  const hh = String(date.getHours()).padStart(2, "0");
  const mi = String(date.getMinutes()).padStart(2, "0");

  return `${dd}/${mm}/${yyyy} ${hh}:${mi}`;
}

function parseAuditDetails(detailsJson) {
  if (!detailsJson) return "";

  let details = detailsJson;

  if (typeof detailsJson === "string") {
    try {
      details = JSON.parse(detailsJson);
    } catch {
      return detailsJson;
    }
  }

  if (typeof details !== "object" || details === null) {
    return String(details);
  }

  if (details.removed_count) {
    return `Xóa ${details.removed_count} file Excel`;
  }

  if (details.file_name && details.sheet_count) {
    return `${details.file_name} / ${details.sheet_count} sheet / ${details.total_rows || 0} dòng`;
  }

  if (details.mapped_count) {
    return `Map ${details.mapped_count} dòng`;
  }

  const parts = Object.entries(details).map(([key, value]) => `${key}: ${value}`);
  return parts.join(" | ");
}

function renderAuditLogs(logs) {
  if (!auditLogsTableBody) return;

  if (!Array.isArray(logs) || logs.length === 0) {
    auditLogsTableBody.innerHTML = `
      <tr>
        <td colspan="5" class="admin-history-empty">Chưa có lịch sử thao tác</td>
      </tr>
    `;
    return;
  }

  auditLogsTableBody.innerHTML = logs.map((log) => {
    const target = log.target_value || "-";
    const detailText = parseAuditDetails(log.details_json);

    return `
      <tr>
        <td class="admin-history-time">${escapeHtml(formatAuditDate(log.created_at))}</td>
        <td class="admin-history-user">${escapeHtml(log.user_email || "-")}</td>
        <td><span class="admin-history-action">${escapeHtml(log.action_label || log.action_type || "-")}</span></td>
        <td class="admin-history-target">${escapeHtml(target)}</td>
        <td class="admin-history-detail">${escapeHtml(detailText || "-")}</td>
      </tr>
    `;
  }).join("");
}

async function loadAuditLogs() {
  if (!auditLogsTableBody) return;

  auditLogsTableBody.innerHTML = `
    <tr>
      <td colspan="5" class="admin-history-empty">Đang tải lịch sử thao tác...</td>
    </tr>
  `;

  try {
    const response = await fetch(`${basePath}/api/admin/audit-logs`);
    const result = await response.json();

    if (!response.ok || !result.success) {
      auditLogsTableBody.innerHTML = `
        <tr>
          <td colspan="5" class="admin-history-empty">
            ${escapeHtml(result.message || "Không tải được lịch sử thao tác")}
          </td>
        </tr>
      `;
      return;
    }

    const logs = Array.isArray(result.data?.logs) ? result.data.logs : [];
    renderAuditLogs(logs);
  } catch (error) {
    console.error(error);
    auditLogsTableBody.innerHTML = `
      <tr>
        <td colspan="5" class="admin-history-empty">Có lỗi khi tải lịch sử thao tác</td>
      </tr>
    `;
  }
}

reloadAuditLogsBtn?.addEventListener("click", () => {
  loadAuditLogs();
});


let adminUsersCache = [];

function getUserInitials(user) {
  const source = (user.fullName || user.email || "U").trim();
  const words = source.split(/\s+/).filter(Boolean);

  if (words.length >= 2) {
    return `${words[0][0] || ""}${words[1][0] || ""}`.toUpperCase();
  }

  return source.slice(0, 2).toUpperCase();
}

function getUserRoleLabel(user) {
  if (user.isAdmin) return "Admin";
  if (user.permissions?.canManageUsers) return "Quản lý";
  if (user.permissions?.canResetData) return "Reset";
  return "Người dùng";
}

function renderAdminUsers(users) {
  if (!usersTableWrap) return;

  if (!users.length) {
    usersTableWrap.innerHTML = `
      <div class="admin-empty-state">Không có tài khoản nào phù hợp.</div>
    `;
    return;
  }

  usersTableWrap.innerHTML = `
    <div class="admin-user-grid">
      ${users.map(renderUserGridCard).join("")}
    </div>
  `;

  bindAdminGridActions();
}

function renderUserGridCard(user) {
  const roleLabel = getUserRoleLabel(user);
  const activeClass = user.isActive ? "green" : "red";
  const activeText = user.isActive ? "Đang hoạt động" : "Ngưng hoạt động";

  return `
    <article class="admin-user-card" data-user-id="${user.id}">
      <div class="admin-user-card-top">
        <div class="admin-user-avatar">${escapeHtml(getUserInitials(user))}</div>

        <div class="admin-user-meta">
          <div class="admin-user-name">${escapeHtml(user.fullName || "Chưa có tên")}</div>

          <div class="admin-user-role-row">
            <span class="admin-tag">${escapeHtml(roleLabel)}</span>
            ${user.isAdmin ? `<span class="admin-tag gray">Quản trị</span>` : ""}
          </div>
        </div>
      </div>

      <div class="admin-user-email">${escapeHtml(user.email || "")}</div>

      <div class="admin-user-status">
        <span class="admin-dot ${user.isActive ? "" : "inactive"}"></span>
        <span>${escapeHtml(activeText)}</span>
      </div>
    </article>
  `;
}

function bindAdminGridActions() {
  document.querySelectorAll(".admin-user-card").forEach((card) => {
    card.addEventListener("click", () => {
      const userId = Number(card.dataset.userId);
      const user = adminUsersCache.find((item) => Number(item.id) === userId);
      if (!user) return;
      openEditUserModal(user);
    });
  });
}

function openAdminModal(modal) {
  if (!modal) return;
  modal.classList.remove("hidden");
  document.body.classList.add("modal-open");
}

function closeAdminModal(modal) {
  if (!modal) return;
  modal.classList.add("hidden");

  const hasOpenModal = !createUserModal?.classList.contains("hidden") || !editUserModal?.classList.contains("hidden");
  if (!hasOpenModal) {
    document.body.classList.remove("modal-open");
  }
}

function resetCreateUserFormState() {
  createUserForm?.reset();
  if (document.getElementById("permImportExcel")) document.getElementById("permImportExcel").checked = true;
  if (document.getElementById("permImportFolder")) document.getElementById("permImportFolder").checked = true;
  if (document.getElementById("permSearchImage")) document.getElementById("permSearchImage").checked = true;
  if (document.getElementById("permViewData")) document.getElementById("permViewData").checked = true;
  if (document.getElementById("permIsActive")) document.getElementById("permIsActive").checked = true;
  if (createUserStatus) createUserStatus.textContent = "";
}

function openEditUserModal(user) {
  editUserId.value = user.id;
  editUserEmail.value = user.email || "";
  editFullName.value = user.fullName || "";
  editIsActive.checked = !!user.isActive;
  editIsAdmin.checked = !!user.isAdmin;

  editCanImportExcel.checked = !!user.permissions?.canImportExcel;
  editCanImportFolder.checked = !!user.permissions?.canImportFolder;
  editCanSearchImage.checked = !!user.permissions?.canSearchImage;
  editCanViewData.checked = !!user.permissions?.canViewData;
  editCanResetData.checked = !!user.permissions?.canResetData;
  editCanManageUsers.checked = !!user.permissions?.canManageUsers;
  editCanViewAuditLogs.checked = !!user.permissions?.canViewAuditLogs;

  editNewPassword.value = "";
  editUserStatus.textContent = "";
  editUserModalTitle.textContent = user.fullName || user.email || "Cập nhật tài khoản";
  editUserModalSubTitle.textContent = user.email || "";

  const editUserAvatar = document.getElementById("editUserAvatar");
  if (editUserAvatar) {
    editUserAvatar.textContent = getUserInitials(user);
  }

  openAdminModal(editUserModal);
}

async function loadUsers() {
  if (!usersTableWrap || !usersStatus) return;

  usersStatus.textContent = "Đang tải danh sách tài khoản...";

  try {
    const response = await fetch(`${basePath}/api/admin/users`);
    const result = await response.json();

    if (!response.ok || !result.success) {
      usersStatus.textContent = result.message || "Không tải được danh sách user";
      return;
    }

    adminUsersCache = Array.isArray(result.data?.users) ? result.data.users : [];
    usersStatus.textContent = `Có ${adminUsersCache.length} tài khoản`;
    renderAdminUsers(adminUsersCache);
  } catch (error) {
    console.error(error);
    usersStatus.textContent = "Có lỗi khi tải danh sách user";
  }
}

function filterAdminUsers() {
  const keyword = String(adminUserSearch?.value || "").trim().toLowerCase();

  if (!keyword) {
    renderAdminUsers(adminUsersCache);
    return;
  }

  const filtered = adminUsersCache.filter((user) => {
    const email = String(user.email || "").toLowerCase();
    const fullName = String(user.fullName || "").toLowerCase();
    return email.includes(keyword) || fullName.includes(keyword);
  });

  renderAdminUsers(filtered);
}

openCreateUserModalBtn?.addEventListener("click", () => {
  resetCreateUserFormState();
  openAdminModal(createUserModal);
});

createUserModalCloseBtn?.addEventListener("click", () => closeAdminModal(createUserModal));
cancelCreateUserBtn?.addEventListener("click", () => closeAdminModal(createUserModal));
editUserModalCloseBtn?.addEventListener("click", () => closeAdminModal(editUserModal));
cancelEditUserBtn?.addEventListener("click", () => closeAdminModal(editUserModal));

document.addEventListener("click", (event) => {
  const closeCreate = event.target.closest('[data-admin-close="create"]');
  const closeEdit = event.target.closest('[data-admin-close="edit"]');

  if (closeCreate) closeAdminModal(createUserModal);
  if (closeEdit) closeAdminModal(editUserModal);
});

adminUserSearch?.addEventListener("input", filterAdminUsers);
adminReloadBtn?.addEventListener("click", loadUsers);

createUserForm?.addEventListener("submit", async (event) => {
  event.preventDefault();

  const payload = {
    email: document.getElementById("createEmail").value.trim(),
    password: document.getElementById("createPassword").value,
    full_name: document.getElementById("createFullName").value.trim(),
    can_import_excel: document.getElementById("permImportExcel").checked,
    can_import_folder: document.getElementById("permImportFolder").checked,
    can_search_image: document.getElementById("permSearchImage").checked,
    can_view_data: document.getElementById("permViewData").checked,
    can_reset_data: document.getElementById("permResetData").checked,
    can_manage_users: document.getElementById("permManageUsers").checked,
    is_admin: document.getElementById("permIsAdmin").checked,
    is_active: document.getElementById("permIsActive").checked,
    can_view_audit_logs: editCanViewAuditLogs.checked
  };

  createUserStatus.textContent = "Đang tạo tài khoản...";

  try {
    const response = await fetch(`${basePath}/api/admin/users`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });

    const result = await response.json();

    if (!response.ok || !result.success) {
      createUserStatus.textContent = result.message || "Tạo tài khoản thất bại";
      return;
    }

    createUserStatus.textContent = result.message || "Tạo tài khoản thành công";
    await loadUsers();
    setTimeout(() => {
      closeAdminModal(createUserModal);
      resetCreateUserFormState();
    }, 500);
  } catch (error) {
    console.error(error);
    createUserStatus.textContent = "Có lỗi khi tạo tài khoản";
  }
});

saveEditUserBtn?.addEventListener("click", async () => {
  const userId = editUserId.value;

  const payload = {
    full_name: editFullName.value.trim(),
    is_active: editIsActive.checked,
    is_admin: editIsAdmin.checked,
    can_import_excel: editCanImportExcel.checked,
    can_import_folder: editCanImportFolder.checked,
    can_search_image: editCanSearchImage.checked,
    can_view_data: editCanViewData.checked,
    can_reset_data: editCanResetData.checked,
    can_manage_users: editCanManageUsers.checked,
    can_view_audit_logs: editCanViewAuditLogs.checked
  };

  editUserStatus.textContent = "Đang lưu thay đổi...";

  try {
    const response = await fetch(`${basePath}/api/admin/users/${userId}/permissions`, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });

    const result = await response.json();

    if (!response.ok || !result.success) {
      editUserStatus.textContent = result.message || "Lưu thay đổi thất bại";
      return;
    }

    editUserStatus.textContent = result.message || "Đã cập nhật thành công";
    await loadUsers();
  } catch (error) {
    console.error(error);
    editUserStatus.textContent = "Có lỗi khi cập nhật tài khoản";
  }
});

resetUserPasswordBtn?.addEventListener("click", async () => {
  const userId = editUserId.value;
  const password = editNewPassword.value.trim();

  if (!password) {
    editUserStatus.textContent = "Bạn chưa nhập mật khẩu mới";
    return;
  }

  editUserStatus.textContent = "Đang đổi mật khẩu...";

  try {
    const response = await fetch(`${basePath}/api/admin/users/${userId}/password`, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ password })
    });

    const result = await response.json();

    if (!response.ok || !result.success) {
      editUserStatus.textContent = result.message || "Đổi mật khẩu thất bại";
      return;
    }

    editUserStatus.textContent = result.message || "Đổi mật khẩu thành công";
    editNewPassword.value = "";
  } catch (error) {
    console.error(error);
    editUserStatus.textContent = "Có lỗi khi đổi mật khẩu";
  }
});

let currentUser = null;
let currentPermissions = {
  canImportExcel: false,
  canImportFolder: false,
  canSearchImage: false,
  canViewData: false,
  canResetData: false,
  canManageUsers: false,
  canViewAuditLogs: false,
};

const resetDataBtn = document.getElementById("resetDataBtn");
const excelUploadForm = document.getElementById("excelUploadForm");
const excelFileInput = document.getElementById("excelFile");
const uploadResult = document.getElementById("uploadResult");

const folderCard = document.getElementById("folderCard");
const folderImportForm = document.getElementById("folderImportForm");
const folderInput = document.getElementById("folderInput");
const folderResult = document.getElementById("folderResult");

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

const excelSearchInput = document.getElementById("excelSearchInput");
const clearExcelSearchBtn = document.getElementById("clearExcelSearchBtn");
const excelSearchStatus = document.getElementById("excelSearchStatus");

let currentExcelSearch = "";
let activeSheetPanelId = null;

let workbookList = [];
let workbookData = null;

const resetScopeCard = document.getElementById("resetScopeCard");
const resetCustomerSelect = document.getElementById("resetCustomerSelect");
const resetSeasonSelect = document.getElementById("resetSeasonSelect");
const resetCustomerBtn = document.getElementById("resetCustomerBtn");
const resetSeasonBtn = document.getElementById("resetSeasonBtn");
const resetScopeStatus = document.getElementById("resetScopeStatus");

let resetOptionsState = {
  customers: [],
  seasonsByCustomer: {},
};

const confirmModal = document.getElementById("confirmModal");
const confirmModalTitle = document.getElementById("confirmModalTitle");
const confirmModalMessage = document.getElementById("confirmModalMessage");
const confirmModalOkBtn = document.getElementById("confirmModalOkBtn");
const confirmModalCancelBtn = document.getElementById("confirmModalCancelBtn");
const confirmModalCloseBtn = document.getElementById("confirmModalCloseBtn");

let confirmDialogResolver = null;

const resetCustomerBox = document.getElementById("resetCustomerBox");
const resetCustomerTrigger = document.getElementById("resetCustomerTrigger");
const resetCustomerMenu = document.getElementById("resetCustomerMenu");

const resetSeasonBox = document.getElementById("resetSeasonBox");
const resetSeasonTrigger = document.getElementById("resetSeasonTrigger");
const resetSeasonMenu = document.getElementById("resetSeasonMenu");

let selectedResetCustomer = "";
let selectedResetSeason = "";

/* =========================
   FILTER CHO KẾT QUẢ TÌM ẢNH
========================= */
const MATCH_COLUMNS = [
  {
    key: "matchedImage",
    label: "Ảnh",
    filterable: false,
    getValue: (item) => item.matchedImage || "",
  },
  {
    key: "styleNo",
    label: "Style No",
    filterable: true,
    getValue: (item) => item.row?.["Style No"] || "",
  },
  {
    key: "styleName",
    label: "Style Name",
    filterable: true,
    getValue: (item) => item.row?.["Style Name"] || "",
  },
  {
    key: "customer",
    label: "Customer",
    filterable: true,
    getValue: (item) => item.row?.["Customer"] || "",
  },
  {
    key: "season",
    label: "Season",
    filterable: true,
    getValue: (item) => item.row?.["Season"] || "",
  },
  {
    key: "staff",
    label: "Staff",
    filterable: true,
    getValue: (item) => item.row?.["Staff"] || "",
  },
  {
    key: "product",
    label: "Product",
    filterable: true,
    getValue: (item) => item.row?.["Product"] || "",
  },
  {
    key: "categories",
    label: "Categories",
    filterable: true,
    getValue: (item) => item.row?.["Categories"] || "",
  },
  {
    key: "folderName",
    label: "Folder",
    filterable: true,
    getValue: (item) => item.folderName || "",
  },
  {
    key: "detail",
    label: "Chi tiết",
    filterable: false,
    getValue: (item) => item.detailUrl || "",
  },
];

let originalMatchResults = [];
let filteredMatchResults = [];
const activeColumnFilters = {};
let filterMenuState = null;
let filterMenuAnchor = null;

/* =========================
   FILTER CHO TAB DỮ LIỆU EXCEL
========================= */
const excelColumnFilters = {};
let excelFilterMenuState = null;
let excelFilterMenuAnchor = null;

/* =========================
   INIT
========================= */
excelSearchInput?.addEventListener("input", handleExcelSearch);
clearExcelSearchBtn?.addEventListener("click", clearExcelSearch);
excelUploadForm?.addEventListener("submit", handleUploadExcel);
folderImportForm?.addEventListener("submit", handleImportFolder);
pasteZone?.addEventListener("click", () => pasteZone.focus());
pasteZone?.addEventListener("paste", handlePasteImage);
resetDataBtn?.addEventListener("click", handleResetAllData);

document.addEventListener("DOMContentLoaded", async () => {
  const ok = await loadCurrentUser();
  if (!ok) return;
  await restoreStateOnReload();
});

resetCustomerSelect?.addEventListener("change", handleResetCustomerChange);
resetCustomerBtn?.addEventListener("click", handleResetCustomerOnly);
resetSeasonBtn?.addEventListener("click", handleResetCustomerSeason);

navButtons.forEach((button) => {
  button.addEventListener("click", async () => {
    const targetId = button.dataset.target;
    activatePanel(targetId);

    if (targetId === "adminCard" && currentPermissions.canManageUsers) {
      await loadUsers();
    }

    if (targetId === "auditCard" && currentPermissions.canViewAuditLogs) {
      await loadAuditLogs();
    }
  });
});

adminUserSearch?.addEventListener("input", () => {
  const keyword = adminUserSearch.value.trim().toLowerCase();

  const filtered = adminUsersCache.filter((user) => {
    const email = String(user.email || "").toLowerCase();
    const fullName = String(user.fullName || "").toLowerCase();
    return email.includes(keyword) || fullName.includes(keyword);
  });

  renderAdminUsers(filtered);
});

adminReloadBtn?.addEventListener("click", () => {
  loadUsers();
});

matchResults?.addEventListener("click", (event) => {
  const filterBtn = event.target.closest(".filter-trigger");
  if (!filterBtn) return;

  event.stopPropagation();

  const columnKey = filterBtn.dataset.filterKey;
  if (!columnKey) return;

  if (filterMenuState?.columnKey === columnKey) {
    closeFilterMenu();
    return;
  }

  openFilterMenu(columnKey, filterBtn);
});

excelDataContainer?.addEventListener("click", (event) => {
  const filterBtn = event.target.closest(".excel-data-filter-trigger");
  if (!filterBtn) return;

  event.stopPropagation();

  const header = filterBtn.dataset.excelFilterKey;
  if (!header) return;

  if (excelFilterMenuState?.header === header) {
    closeExcelFilterMenu();
    return;
  }

  openExcelFilterMenu(header, filterBtn);
});

document.addEventListener("click", (event) => {
  const clickedInsideCustomer = resetCustomerBox?.contains(event.target);
  const clickedInsideSeason = resetSeasonBox?.contains(event.target);

  if (!clickedInsideCustomer) {
    resetCustomerBox?.classList.remove("open");
    resetCustomerMenu?.classList.add("hidden");
  }

  if (!clickedInsideSeason) {
    resetSeasonBox?.classList.remove("open");
    resetSeasonMenu?.classList.add("hidden");
  }

  if (filterMenuState) {
    const popup = document.getElementById("excelFilterPopup");
    const clickedInsidePopup = popup?.contains(event.target);
    const clickedOnFilterButton = event.target.closest(".filter-trigger");

    if (!clickedInsidePopup && !clickedOnFilterButton) {
      closeFilterMenu();
    }
  }

  if (excelFilterMenuState) {
    const popup = document.getElementById("excelDataFilterPopup");
    const clickedInsidePopup = popup?.contains(event.target);
    const clickedOnFilterButton = event.target.closest(".excel-data-filter-trigger");

    if (!clickedInsidePopup && !clickedOnFilterButton) {
      closeExcelFilterMenu();
    }
  }
});

window.addEventListener("resize", () => {
  if (filterMenuState) positionFilterPopup();
  if (excelFilterMenuState) positionExcelFilterPopup();
});

window.addEventListener(
  "scroll",
  () => {
    if (filterMenuState) positionFilterPopup();
    if (excelFilterMenuState) positionExcelFilterPopup();
  },
  true
);

confirmModalOkBtn?.addEventListener("click", () => closeConfirmModal(true));
confirmModalCancelBtn?.addEventListener("click", () => closeConfirmModal(false));
confirmModalCloseBtn?.addEventListener("click", () => closeConfirmModal(false));

confirmModal?.addEventListener("click", (event) => {
  const backdrop = event.target.closest('[data-confirm-close="backdrop"]');
  if (backdrop) {
    closeConfirmModal(false);
  }
});

document.addEventListener("keydown", (event) => {
  if (event.key === "Escape" && confirmModal && !confirmModal.classList.contains("hidden")) {
    closeConfirmModal(false);
  }
});

document.addEventListener("click", (e) => {
  const img = e.target.closest(".previewable-image");
  if (!img) return;

  const lightbox = document.getElementById("imageLightbox");
  const lightboxImg = lightbox?.querySelector(".lightbox-img");
  if (!lightbox || !lightboxImg) return;

  lightboxImg.src = img.src;
  lightbox.classList.remove("hidden");
});

document.addEventListener("click", (e) => {
  const lightbox = document.getElementById("imageLightbox");
  if (!lightbox || lightbox.classList.contains("hidden")) return;

  if (
    e.target.classList.contains("lightbox") ||
    e.target.classList.contains("lightbox-close")
  ) {
    lightbox.classList.add("hidden");
  }
});

logoutBtn?.addEventListener("click", handleLogout);

resetCustomerTrigger?.addEventListener("click", () => {
  resetCustomerBox?.classList.toggle("open");
  resetCustomerMenu?.classList.toggle("hidden");
  resetSeasonBox?.classList.remove("open");
  resetSeasonMenu?.classList.add("hidden");
});

resetSeasonTrigger?.addEventListener("click", () => {
  if (!selectedResetCustomer) return;
  resetSeasonBox?.classList.toggle("open");
  resetSeasonMenu?.classList.toggle("hidden");
  resetCustomerBox?.classList.remove("open");
  resetCustomerMenu?.classList.add("hidden");
});

resetCustomerMenu?.addEventListener("click", (event) => {
  const option = event.target.closest("[data-reset-customer]");
  if (!option) return;

  selectedResetCustomer = option.dataset.resetCustomer || "";
  selectedResetSeason = "";

  renderResetCustomerOptions();
  handleResetCustomerChange();

  resetCustomerBox?.classList.remove("open");
  resetCustomerMenu?.classList.add("hidden");
});

resetSeasonMenu?.addEventListener("click", (event) => {
  const option = event.target.closest("[data-reset-season]");
  if (!option) return;

  selectedResetSeason = option.dataset.resetSeason || "";
  renderResetSeasonOptions(selectedResetCustomer);

  resetSeasonBox?.classList.remove("open");
  resetSeasonMenu?.classList.add("hidden");
});

createUserForm?.addEventListener("submit", async (event) => {
  event.preventDefault();

  const payload = {
    email: document.getElementById("createEmail").value.trim(),
    password: document.getElementById("createPassword").value,
    full_name: document.getElementById("createFullName").value.trim(),
    can_import_excel: document.getElementById("permImportExcel").checked,
    can_import_folder: document.getElementById("permImportFolder").checked,
    can_search_image: document.getElementById("permSearchImage").checked,
    can_view_data: document.getElementById("permViewData").checked,
    can_reset_data: document.getElementById("permResetData").checked,
    can_manage_users: document.getElementById("permManageUsers").checked,
    is_admin: document.getElementById("permIsAdmin").checked,
    is_active: document.getElementById("permIsActive").checked,
    can_view_audit_logs: document.getElementById("permViewAuditLogs").checked,
  };

  createUserStatus.textContent = "Đang tạo tài khoản...";

  try {
    const response = await fetch(`${basePath}/api/admin/users`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(payload),
    });

    const result = await response.json();

    if (!response.ok || !result.success) {
      createUserStatus.textContent = result.message || "Tạo tài khoản thất bại";
      return;
    }

    createUserStatus.textContent = result.message || "Tạo tài khoản thành công";
    createUserForm.reset();

    document.getElementById("permImportExcel").checked = true;
    document.getElementById("permImportFolder").checked = true;
    document.getElementById("permSearchImage").checked = true;
    document.getElementById("permViewData").checked = true;
    document.getElementById("permIsActive").checked = true;

    await loadUsers();
  } catch (error) {
    console.error(error);
    createUserStatus.textContent = "Có lỗi khi tạo tài khoản";
  }
});

updateSidebarAvailability();
activatePanel("uploadCard");

/* =========================
   CORE
========================= */
function updateResetTabAvailability() {
  if (!resetScopeCard) return;

  if (workbookList.length > 0 && currentPermissions.canResetData) {
    resetScopeCard.classList.remove("hidden");
  } else {
    resetScopeCard.classList.add("hidden");
  }

  updateSidebarAvailability();
}

async function loadCurrentUser() {
  try {
    const response = await fetch(`${basePath}/api/me`);
    const result = await response.json();

    if (!response.ok || !result.success) {
      window.location.href = `${basePath}/login`;
      return false;
    }

    currentUser = result.data;
    currentPermissions = result.data.permissions || currentPermissions;

    if (currentUserLabel) {
      currentUserLabel.textContent = currentUser.email || "";
    }

    const sidebarUserAvatar = document.getElementById("sidebarUserAvatar");
    if (sidebarUserAvatar && currentUser?.email) {
      sidebarUserAvatar.textContent = currentUser.email.trim().charAt(0).toUpperCase();
    }

    applyPermissionsToUI();
    return true;
  } catch (error) {
    console.error(error);
    window.location.href = `${basePath}/login`;
    return false;
  }
}


function renderUserCard(user) {
  const displayName = user.fullName || "Chưa có tên";
  const activeBadge = user.isActive
    ? `<span class="admin-badge green">Đang hoạt động</span>`
    : `<span class="admin-badge red">Ngưng hoạt động</span>`;

  const adminBadge = user.isAdmin
    ? `<span class="admin-badge">Admin</span>`
    : `<span class="admin-badge gray">User</span>`;

  return `
    <details class="admin-user-item" data-user-id="${user.id}">
      <summary class="admin-user-summary">
        <div class="admin-user-main">
          <div class="admin-user-name">${escapeHtml(displayName)}</div>
          <div class="admin-user-email">${escapeHtml(user.email)}</div>
        </div>

        <div class="admin-user-badges">
          ${activeBadge}
          ${adminBadge}
        </div>
      </summary>

      <div class="admin-user-detail">
        <div class="admin-user-detail-grid">
          <div class="admin-user-block">
            <h5>Thông tin chung</h5>
            <div class="admin-stack">
              <input
                class="input-full-name admin-user-input"
                type="text"
                value="${escapeAttribute(user.fullName || "")}"
                placeholder="Họ tên"
              />
            </div>
          </div>

          <div class="admin-user-block">
            <h5>Trạng thái</h5>
            <div class="admin-check-grid two-col">
              <label class="admin-check-chip">
                <input class="input-active" type="checkbox" ${user.isActive ? "checked" : ""} />
                <span>Đang hoạt động</span>
              </label>

              <label class="admin-check-chip">
                <input class="input-admin" type="checkbox" ${user.isAdmin ? "checked" : ""} />
                <span>Là admin</span>
              </label>
            </div>
          </div>

          <div class="admin-user-block">
            <h5>Đổi mật khẩu</h5>
            <div class="admin-stack">
              <input
                class="input-new-password admin-password-input"
                type="text"
                placeholder="Nhập mật khẩu mới"
              />
              <button class="btn btn-reset-password" type="button">Đổi mật khẩu</button>
            </div>
          </div>
        </div>

        <div class="admin-user-block" style="margin-top:14px;">
          <h5>Phân quyền</h5>
          <div class="admin-check-grid three-col">
            <label class="admin-check-chip">
              <input class="perm-import-excel" type="checkbox" ${user.permissions.canImportExcel ? "checked" : ""} />
              <span>Import Excel</span>
            </label>

            <label class="admin-check-chip">
              <input class="perm-import-folder" type="checkbox" ${user.permissions.canImportFolder ? "checked" : ""} />
              <span>Import Folder</span>
            </label>

            <label class="admin-check-chip">
              <input class="perm-search-image" type="checkbox" ${user.permissions.canSearchImage ? "checked" : ""} />
              <span>Tìm theo ảnh</span>
            </label>

            <label class="admin-check-chip">
              <input class="perm-view-data" type="checkbox" ${user.permissions.canViewData ? "checked" : ""} />
              <span>Xem dữ liệu</span>
            </label>

            <label class="admin-check-chip">
              <input class="perm-reset-data" type="checkbox" ${user.permissions.canResetData ? "checked" : ""} />
              <span>Reset dữ liệu</span>
            </label>

            <label class="admin-check-chip">
              <input class="perm-manage-users" type="checkbox" ${user.permissions.canManageUsers ? "checked" : ""} />
              <span>Quản lý tài khoản</span>
            </label>
          </div>

          <div class="admin-inline-actions">
            <button class="btn btn-save-user" type="button">Lưu quyền</button>
          </div>
        </div>
      </div>
    </details>
  `;
}

function bindUserActions() {
  document.querySelectorAll(".btn-save-user").forEach((btn) => {
    btn.addEventListener("click", async () => {
      const row = btn.closest("[data-user-id]");
      if (!row) return;

      const userId = row.dataset.userId;

      const payload = {
        full_name: row.querySelector(".input-full-name").value.trim(),
        is_active: row.querySelector(".input-active").checked,
        is_admin: row.querySelector(".input-admin").checked,
        can_import_excel: row.querySelector(".perm-import-excel").checked,
        can_import_folder: row.querySelector(".perm-import-folder").checked,
        can_search_image: row.querySelector(".perm-search-image").checked,
        can_view_data: row.querySelector(".perm-view-data").checked,
        can_reset_data: row.querySelector(".perm-reset-data").checked,
        can_manage_users: row.querySelector(".perm-manage-users").checked,
        can_view_audit_logs: editCanViewAuditLogs.checked
      };

      try {
        const response = await fetch(`${basePath}/api/admin/users/${userId}/permissions`, {
          method: "PUT",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(payload),
        });

        const result = await response.json();
        alert(result.message || "Đã lưu");

        if (response.ok && result.success) {
          await loadUsers();
        }
      } catch (error) {
        console.error(error);
        alert("Có lỗi khi cập nhật quyền");
      }
    });
  });

  document.querySelectorAll(".btn-reset-password").forEach((btn) => {
    btn.addEventListener("click", async () => {
      const row = btn.closest("[data-user-id]");
      if (!row) return;

      const userId = row.dataset.userId;
      const password = row.querySelector(".input-new-password").value.trim();

      if (!password) {
        alert("Nhập mật khẩu mới");
        return;
      }

      try {
        const response = await fetch(`${basePath}/api/admin/users/${userId}/password`, {
          method: "PUT",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ password }),
        });

        const result = await response.json();
        alert(result.message || "Đã đổi mật khẩu");

        if (response.ok && result.success) {
          row.querySelector(".input-new-password").value = "";
        }
      } catch (error) {
        console.error(error);
        alert("Có lỗi khi đổi mật khẩu");
      }
    });
  });
}

function applyPermissionsToUI() {
  navExcelBtn?.classList.toggle("hidden", !currentPermissions.canImportExcel);
  navFolderBtn?.classList.toggle("hidden", !currentPermissions.canImportFolder);
  navImageBtn?.classList.toggle("hidden", !currentPermissions.canSearchImage);
  navDataBtn?.classList.toggle("hidden", !currentPermissions.canViewData);
  navResetScopeBtn?.classList.toggle("hidden", !currentPermissions.canResetData);
  navAdminBtn?.classList.toggle("hidden", !currentPermissions.canManageUsers);
  navAuditBtn?.classList.toggle("hidden", !currentPermissions.canViewAuditLogs);

  if (resetScopeCard) {
    resetScopeCard.classList.toggle("hidden", !currentPermissions.canResetData);
  }

  if (auditCard) {
    auditCard.classList.toggle("hidden", !currentPermissions.canViewAuditLogs);
  }

  if (adminCard) {
    adminCard.classList.toggle("hidden", !currentPermissions.canManageUsers);
  }

  updateSidebarAvailability();
}

async function handleLogout() {
  try {
    await fetch(`${basePath}/api/logout`, { method: "POST" });
  } catch (error) {
    console.error(error);
  }
  window.location.href = `${basePath}/login`;
}

function closeConfirmModal(result = false) {
  if (!confirmModal) return;

  confirmModal.classList.add("hidden");
  document.body.classList.remove("modal-open");

  if (typeof confirmDialogResolver === "function") {
    confirmDialogResolver(result);
    confirmDialogResolver = null;
  }
}

function showConfirmDialog({
  title = "Xác nhận thao tác",
  message = "Bạn có chắc muốn thực hiện thao tác này không?",
  confirmText = "Xác nhận",
  cancelText = "Hủy",
} = {}) {
  if (!confirmModal) {
    return Promise.resolve(window.confirm(message));
  }

  confirmModalTitle.textContent = title;
  confirmModalMessage.textContent = message;
  confirmModalOkBtn.textContent = confirmText;
  confirmModalCancelBtn.textContent = cancelText;

  confirmModal.classList.remove("hidden");
  document.body.classList.add("modal-open");

  return new Promise((resolve) => {
    confirmDialogResolver = resolve;
  });
}

function setResetOptions(options) {
  resetOptionsState = {
    customers: Array.isArray(options?.customers) ? options.customers : [],
    seasonsByCustomer: options?.seasonsByCustomer || {},
  };

  if (!resetOptionsState.customers.includes(selectedResetCustomer)) {
    selectedResetCustomer = "";
    selectedResetSeason = "";
  }

  renderResetCustomerOptions();
  renderResetSeasonOptions(selectedResetCustomer);
  updateResetTabAvailability();
}

function renderResetCustomerOptions() {
  if (!resetCustomerMenu || !resetCustomerTrigger) return;

  const customers = resetOptionsState.customers || [];

  resetCustomerMenu.innerHTML = `
    <button type="button" class="custom-select-option ${selectedResetCustomer === "" ? "active" : ""}" data-reset-customer="">
      Chọn customer
    </button>
    ${customers
      .map(
        (customer) => `
          <button
            type="button"
            class="custom-select-option ${selectedResetCustomer === customer ? "active" : ""}"
            data-reset-customer="${escapeAttribute(customer)}"
          >
            ${escapeHtml(customer)}
          </button>
        `
      )
      .join("")}
  `;

  resetCustomerTrigger.textContent = selectedResetCustomer || "Chọn customer";
}

function renderResetSeasonOptions(customer) {
  if (!resetSeasonMenu || !resetSeasonTrigger) return;

  const seasons = Array.isArray(resetOptionsState.seasonsByCustomer?.[customer])
    ? resetOptionsState.seasonsByCustomer[customer]
    : [];

  if (!seasons.includes(selectedResetSeason)) {
    selectedResetSeason = "";
  }

  resetSeasonMenu.innerHTML = `
    <button type="button" class="custom-select-option ${selectedResetSeason === "" ? "active" : ""}" data-reset-season="">
      Chọn season
    </button>
    ${seasons
      .map(
        (season) => `
          <button
            type="button"
            class="custom-select-option ${selectedResetSeason === season ? "active" : ""}"
            data-reset-season="${escapeAttribute(season)}"
          >
            ${escapeHtml(season)}
          </button>
        `
      )
      .join("")}
  `;

  resetSeasonTrigger.textContent = selectedResetSeason || "Chọn season";
}

function handleResetCustomerChange() {
  renderResetSeasonOptions(selectedResetCustomer);
}

function combineImportedWorkbooks(workbooks) {
  const imports = Array.isArray(workbooks) ? workbooks : [];
  const mergedSheets = [];
  let totalRows = 0;
  let imageIndexCount = 0;
  let mappedFolderCount = 0;

  imports.forEach((workbook, workbookIndex) => {
    const sheets = Array.isArray(workbook.sheets) ? workbook.sheets : [];

    totalRows += Number(workbook.totalRows || 0);
    imageIndexCount += Number(workbook.imageIndexCount || 0);
    mappedFolderCount += Number(workbook.mappedFolderCount || 0);

    sheets.forEach((sheet, sheetIndex) => {
      mergedSheets.push({
        ...sheet,
        sheetName: `${workbook.fileName} / ${sheet.sheetName}`,
        __workbookId: workbook.id || `wb-${workbookIndex}`,
        __sheetId: `${workbook.id || workbookIndex}-${sheetIndex}`,
      });
    });
  });

  return {
    id: "combined",
    fileName: imports.length === 1 ? imports[0].fileName : `${imports.length} file Excel`,
    sheetCount: mergedSheets.length,
    totalRows,
    imageIndexCount,
    mappedFolderCount,
    sheets: mergedSheets,
  };
}

function refreshCombinedWorkbookView() {
  if (!workbookList.length) {
    workbookData = null;
    sheetTabs.innerHTML = "";
    excelDataContainer.innerHTML = "";
    dataCard?.classList.add("hidden");
    updateSidebarAvailability();
    return;
  }

  workbookData = combineImportedWorkbooks(workbookList);

  folderCard?.classList.remove("hidden");
  imageSearchCard?.classList.remove("hidden");
  dataCard?.classList.remove("hidden");

  renderSheets(workbookData.sheets);
  updateSidebarAvailability();
}

async function handleResetCustomerOnly() {
  const customer = selectedResetCustomer || "";

  if (!customer) {
    resetScopeStatus.textContent = "Bạn chưa chọn customer";
    return;
  }

  const confirmed = await showConfirmDialog({
    title: "Reset theo customer",
    message: `Bạn có chắc muốn reset toàn bộ dữ liệu của customer "${customer}" không?`,
    confirmText: "Reset",
    cancelText: "Hủy",
  });

  if (!confirmed) return;

  try {
    const response = await fetch(`${basePath}/api/reset/customer`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ customer }),
    });

    const result = await response.json();

    if (!response.ok || !result.success) {
      resetScopeStatus.textContent = result.message || "Reset customer thất bại";
      return;
    }

    workbookList = Array.isArray(result.data?.imports) ? result.data.imports : [];
    setResetOptions(result.data?.resetOptions || buildResetOptionsFromWorkbooks(workbookList));
    refreshCombinedWorkbookView();

    resetScopeStatus.textContent = result.message;
    uploadResult.textContent = `Hiện có ${workbookList.length} file Excel.`;

    if (!workbookList.length) {
      folderCard?.classList.add("hidden");
      imageSearchCard?.classList.add("hidden");
      dataCard?.classList.add("hidden");
      activatePanel("uploadCard");
    }
  } catch (error) {
    console.error(error);
    resetScopeStatus.textContent = "Có lỗi khi reset customer";
  }
}

async function handleResetCustomerSeason() {
  const customer = selectedResetCustomer || "";
  const season = selectedResetSeason || "";

  if (!customer || !season) {
    resetScopeStatus.textContent = "Bạn cần chọn customer và season";
    return;
  }

  const confirmed = await showConfirmDialog({
    title: "Reset theo customer / season",
    message: `Bạn có chắc muốn reset dữ liệu của "${customer} / ${season}" không?`,
    confirmText: "Reset",
    cancelText: "Hủy",
  });

  if (!confirmed) return;

  try {
    const response = await fetch(`${basePath}/api/reset/season`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ customer, season }),
    });

    const result = await response.json();

    if (!response.ok || !result.success) {
      resetScopeStatus.textContent = result.message || "Reset season thất bại";
      return;
    }

    workbookList = Array.isArray(result.data?.imports) ? result.data.imports : [];
    setResetOptions(result.data?.resetOptions || buildResetOptionsFromWorkbooks(workbookList));
    refreshCombinedWorkbookView();

    resetScopeStatus.textContent = result.message;
    uploadResult.textContent = `Hiện có ${workbookList.length} file Excel.`;

    if (!workbookList.length) {
      folderCard?.classList.add("hidden");
      imageSearchCard?.classList.add("hidden");
      dataCard?.classList.add("hidden");
      activatePanel("uploadCard");
    }
  } catch (error) {
    console.error(error);
    resetScopeStatus.textContent = "Có lỗi khi reset season";
  }
}

async function restoreStateOnReload() {
  try {
    const response = await fetch(`${basePath}/api/state`);
    const result = await response.json();

    if (!response.ok || !result.success) {
      return;
    }

    const imports = Array.isArray(result.data?.imports) ? result.data.imports : [];
    workbookList = imports;

    setResetOptions(result.data?.resetOptions || buildResetOptionsFromWorkbooks(workbookList));

    if (!imports.length) {
      refreshCombinedWorkbookView();
      return;
    }

    refreshCombinedWorkbookView();

    uploadResult.textContent = `Đã khôi phục ${workbookList.length} file Excel`;

    const totalMapped = workbookList.reduce(
      (sum, item) => sum + Number(item.mappedFolderCount || 0),
      0
    );

    folderResult.textContent =
      totalMapped > 0
        ? `Đã map ${totalMapped} dòng với folder`
        : "Chưa import folder";

    searchStatus.textContent =
      "Dữ liệu đã được khôi phục. Bạn có thể tiếp tục dán ảnh để tìm kiếm.";

    activatePanel("dataCard");
  } catch (error) {
    console.error("Không khôi phục được state:", error);
  }
}

async function handleResetAllData() {
  const confirmed = await showConfirmDialog({
    title: "Reset toàn bộ dữ liệu",
    message: "Bạn có chắc muốn reset toàn bộ dữ liệu đã import không? Thao tác này không thể hoàn tác.",
    confirmText: "Reset",
    cancelText: "Hủy",
  });

  if (!confirmed) {
    return;
  }

  excelSearchInput.value = "";
  currentExcelSearch = "";
  activeSheetPanelId = null;

  Object.keys(excelColumnFilters).forEach((key) => delete excelColumnFilters[key]);
  closeExcelFilterMenu();

  Object.keys(activeColumnFilters).forEach((key) => delete activeColumnFilters[key]);
  closeFilterMenu();

  try {
    const response = await fetch(`${basePath}/api/reset`, {
      method: "POST",
    });

    const result = await response.json();

    if (!response.ok || !result.success) {
      alert(result.message || "Reset thất bại");
      return;
    }

    workbookList = [];
    workbookData = null;
    originalMatchResults = [];
    filteredMatchResults = [];

    excelUploadForm?.reset();
    folderImportForm?.reset();

    uploadResult.textContent = "Chưa có dữ liệu";
    folderResult.textContent = "Chưa import folder";
    searchStatus.textContent = "Chưa dán ảnh";
    excelSearchStatus.textContent = "Nhập từ khóa để lọc dữ liệu theo tất cả cột text.";

    folderCard?.classList.add("hidden");
    imageSearchCard?.classList.add("hidden");
    dataCard?.classList.add("hidden");

    sheetTabs.innerHTML = "";
    excelDataContainer.innerHTML = "";
    matchResults.innerHTML = "";

    pastedPreview?.classList.add("hidden");
    pastedPreview?.removeAttribute("src");

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

function buildResetOptionsFromWorkbooks(workbooks) {
  const seasonsByCustomer = {};
  const customerSet = new Set();

  (workbooks || []).forEach((workbook) => {
    const customer = String(workbook.customer || "").trim();
    if (!customer) return;

    customerSet.add(customer);

    if (!seasonsByCustomer[customer]) {
      seasonsByCustomer[customer] = [];
    }

    const seasonValues = Array.isArray(workbook.seasonValues)
      ? workbook.seasonValues
      : [];

    if (seasonValues.length) {
      seasonValues.forEach((season) => {
        const seasonText = String(season || "").trim();
        if (seasonText && !seasonsByCustomer[customer].includes(seasonText)) {
          seasonsByCustomer[customer].push(seasonText);
        }
      });
    } else {
      const season = String(workbook.season || "").trim();
      if (season && !seasonsByCustomer[customer].includes(season)) {
        seasonsByCustomer[customer].push(season);
      }
    }
  });

  Object.keys(seasonsByCustomer).forEach((customer) => {
    seasonsByCustomer[customer].sort((a, b) =>
      a.localeCompare(b, "vi", { numeric: true, sensitivity: "base" })
    );
  });

  return {
    customers: Array.from(customerSet).sort((a, b) =>
      a.localeCompare(b, "vi", { numeric: true, sensitivity: "base" })
    ),
    seasonsByCustomer,
  };
}

async function handleUploadExcel(event) {
  excelSearchInput.value = "";
  currentExcelSearch = "";
  activeSheetPanelId = null;

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

  try {
    const response = await fetch(`${basePath}/api/excel/upload`, {
      method: "POST",
      body: formData,
    });

    const result = await response.json();

    if (!response.ok || !result.success) {
      uploadResult.textContent = result.message || "Import Excel thất bại";
      return;
    }

    const returnedImports = Array.isArray(result.data?.imports)
      ? result.data.imports
      : [];

    const newWorkbook =
      result.data?.workbook ||
      (result.data && Array.isArray(result.data.sheets) ? result.data : null);

    if (returnedImports.length) {
      workbookList = returnedImports;
      setResetOptions(result.data?.resetOptions || buildResetOptionsFromWorkbooks(workbookList));
    } else if (newWorkbook) {
      const existedIndex = workbookList.findIndex(
        (item) => item.id === newWorkbook.id
      );

      if (existedIndex >= 0) {
        workbookList[existedIndex] = newWorkbook;
      } else {
        workbookList.push(newWorkbook);
      }

      setResetOptions(buildResetOptionsFromWorkbooks(workbookList));
    } else {
      console.error("Upload Excel response không đúng format:", result);
      uploadResult.textContent = "Import Excel thành công nhưng dữ liệu trả về không đúng";
      return;
    }

    refreshCombinedWorkbookView();

    uploadResult.textContent =
      `${result.message} Hiện có ${workbookList.length} file Excel.`;

    folderCard?.classList.remove("hidden");
    imageSearchCard?.classList.remove("hidden");
    dataCard?.classList.remove("hidden");

    updateSidebarAvailability();
    activatePanel("folderCard");

    searchStatus.textContent =
      "Excel đã import xong. Bạn có thể dán ảnh hoặc import folder để map chi tiết.";

    excelUploadForm?.reset();
  } catch (error) {
    console.error(error);
    uploadResult.textContent = "Có lỗi khi upload file Excel";
  }
}

async function handleImportFolder(event) {
  event.preventDefault();

  if (!workbookList.length) {
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
    const response = await fetch(`${basePath}/api/folder/import`, {
      method: "POST",
      body: formData,
    });

    const result = await response.json();

    if (!response.ok || !result.success) {
      folderResult.textContent = result.message || "Import folder thất bại";
      return;
    }

    if (Array.isArray(result.data?.imports) && result.data.imports.length) {
      workbookList = result.data.imports;
    } else if (result.data?.workbook && Array.isArray(result.data.workbook.sheets)) {
      const updatedWorkbook = result.data.workbook;
      const existingIndex = workbookList.findIndex(
        (item) => item.id === updatedWorkbook.id
      );

      if (existingIndex >= 0) {
        workbookList[existingIndex] = updatedWorkbook;
      } else {
        workbookList.push(updatedWorkbook);
      }
    } else {
      console.error("Folder import response không đúng format:", result);
      folderResult.textContent =
        "Import folder thành công nhưng không lấy được dữ liệu workbook";
      return;
    }

    setResetOptions(result.data?.resetOptions || buildResetOptionsFromWorkbooks(workbookList));
    refreshCombinedWorkbookView();

    folderResult.textContent =
      `Import folder thành công. Tổng folder con: ${result.data?.folderCount ?? 0}. ` +
      `Map được: ${result.data?.mappedCount ?? 0} dòng.`;

    updateSidebarAvailability();
    activatePanel("dataCard");

    folderImportForm?.reset();
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
    const response = await fetch(`${basePath}/api/search-by-image`, {
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

/* =========================
   FILTER CHO KẾT QUẢ TÌM ẢNH
========================= */
function normalizeFilterValue(value) {
  return String(value ?? "").trim();
}

function hasActiveFilter(columnKey) {
  return activeColumnFilters[columnKey] instanceof Set;
}

function renderMatchResults(results) {
  if (!Array.isArray(results) || results.length === 0) {
    originalMatchResults = [];
    filteredMatchResults = [];
    matchResults.innerHTML = `<p class="empty-state">Không tìm thấy kết quả phù hợp</p>`;
    closeFilterMenu();
    return;
  }

  originalMatchResults = [...results];
  applyColumnFilters();
  renderMatchResultsTable();
  closeFilterMenu();
}

function applyColumnFilters() {
  filteredMatchResults = originalMatchResults.filter((item) => {
    return MATCH_COLUMNS.every((column) => {
      if (!column.filterable) return true;

      const selectedValues = activeColumnFilters[column.key];
      if (!(selectedValues instanceof Set)) return true;

      const currentValue = normalizeFilterValue(column.getValue(item));
      return selectedValues.has(currentValue);
    });
  });
}

function renderMatchResultsTable() {
  matchResults.innerHTML = `
    <div class="excel-table-wrap">
      <table class="excel-table">
        <thead>
          <tr>
            <th class="col-stt">#</th>
            ${MATCH_COLUMNS.map(renderTableHeader).join("")}
          </tr>
        </thead>
        <tbody>
          ${filteredMatchResults.length
      ? filteredMatchResults.map((item, index) => renderTableRow(item, index)).join("")
      : `
              <tr>
                <td colspan="${MATCH_COLUMNS.length + 1}">
                  <div class="empty-state-table">Không có dữ liệu phù hợp với filter hiện tại</div>
                </td>
              </tr>
            `
    }
        </tbody>
      </table>
    </div>
  `;
}

function renderTableHeader(column) {
  const active = hasActiveFilter(column.key);

  return `
    <th class="${active ? "filter-active-col" : ""}">
      <div class="th-wrap">
        <span class="th-label">${escapeHtml(column.label)}</span>
        ${column.filterable
      ? `
            <button
              type="button"
              class="filter-trigger ${active ? "active" : ""}"
              data-filter-key="${escapeAttribute(column.key)}"
              title="Lọc cột ${escapeAttribute(column.label)}"
            >
              ▾
            </button>
          `
      : ""
    }
      </div>
    </th>
  `;
}

function renderTableRow(item, index) {
  const row = item.row || {};

  return `
    <tr>
      <td class="cell-center cell-stt">${index + 1}</td>

      <td class="cell-image">
        ${item.matchedImage
      ? `<img src="${escapeAttribute(item.matchedImage)}" alt="Ảnh khớp" class="result-thumb previewable-image" />`
      : `<span class="muted-text">-</span>`
    }
      </td>

      <td>${escapeHtml(row["Style No"] || "")}</td>
      <td>${escapeHtml(row["Style Name"] || "")}</td>
      <td>${escapeHtml(row["Customer"] || "")}</td>
      <td>${escapeHtml(row["Season"] || "")}</td>
      <td>${escapeHtml(row["Staff"] || "")}</td>
      <td>${escapeHtml(row["Product"] || "")}</td>
      <td>${escapeHtml(row["Categories"] || "")}</td>
      <td>${escapeHtml(item.folderName || "-")}</td>

      <td class="cell-center">
        ${item.detailUrl
      ? `<a class="detail-link-btn" href="${escapeAttribute(item.detailUrl)}" target="_blank" rel="noopener noreferrer">Mở</a>`
      : `<span class="muted-text">-</span>`
    }
      </td>
    </tr>
  `;
}

function getBaseRowsForColumnFilter(columnKey) {
  return originalMatchResults.filter((item) => {
    return MATCH_COLUMNS.every((column) => {
      if (!column.filterable) return true;
      if (column.key === columnKey) return true;

      const selectedValues = activeColumnFilters[column.key];
      if (!(selectedValues instanceof Set)) return true;

      const currentValue = normalizeFilterValue(column.getValue(item));
      return selectedValues.has(currentValue);
    });
  });
}

function getDistinctColumnValues(columnKey) {
  const column = MATCH_COLUMNS.find((c) => c.key === columnKey);
  if (!column) return [];

  const rows = getBaseRowsForColumnFilter(columnKey);
  const map = new Map();

  rows.forEach((item) => {
    const raw = normalizeFilterValue(column.getValue(item));
    const label = raw || "(Trống)";

    if (!map.has(raw)) {
      map.set(raw, {
        value: raw,
        label,
        count: 0,
      });
    }

    map.get(raw).count += 1;
  });

  return Array.from(map.values()).sort((a, b) =>
    a.label.localeCompare(b.label, "vi", {
      numeric: true,
      sensitivity: "base",
    })
  );
}

function ensureFilterPopup() {
  let popup = document.getElementById("excelFilterPopup");

  if (!popup) {
    popup = document.createElement("div");
    popup.id = "excelFilterPopup";
    popup.className = "excel-filter-popup hidden";
    document.body.appendChild(popup);

    popup.addEventListener("input", handleFilterPopupInput);
    popup.addEventListener("change", handleFilterPopupChange);
    popup.addEventListener("click", handleFilterPopupClick);
  }

  return popup;
}

function openFilterMenu(columnKey, anchorEl) {
  filterMenuAnchor = anchorEl;

  const distinctValues = getDistinctColumnValues(columnKey);
  const selected =
    activeColumnFilters[columnKey] instanceof Set
      ? new Set(activeColumnFilters[columnKey])
      : new Set(distinctValues.map((item) => item.value));

  filterMenuState = {
    columnKey,
    search: "",
    selected,
  };

  renderFilterPopup();
}

function closeFilterMenu() {
  filterMenuState = null;
  filterMenuAnchor = null;

  const popup = document.getElementById("excelFilterPopup");
  if (popup) {
    popup.classList.add("hidden");
    popup.innerHTML = "";
  }
}

function getVisibleFilterOptions() {
  if (!filterMenuState) return [];

  const allValues = getDistinctColumnValues(filterMenuState.columnKey);
  const keyword = filterMenuState.search.trim().toLowerCase();

  if (!keyword) return allValues;

  return allValues.filter((item) =>
    item.label.toLowerCase().includes(keyword)
  );
}

function renderFilterPopup() {
  const popup = ensureFilterPopup();

  if (!filterMenuState) {
    popup.classList.add("hidden");
    popup.innerHTML = "";
    return;
  }

  const column = MATCH_COLUMNS.find((c) => c.key === filterMenuState.columnKey);
  const visibleValues = getVisibleFilterOptions();

  const checkedVisibleCount = visibleValues.filter((item) =>
    filterMenuState.selected.has(item.value)
  ).length;

  const allVisibleChecked =
    visibleValues.length > 0 && checkedVisibleCount === visibleValues.length;

  const noneVisibleChecked = checkedVisibleCount === 0;

  popup.innerHTML = `
    <div class="filter-panel">
      <div class="filter-panel-title">${escapeHtml(column?.label || "Filter")}</div>

      <input
        type="text"
        class="filter-search-input"
        placeholder="Search"
        value="${escapeAttribute(filterMenuState.search)}"
      />

      <div class="filter-value-box">
        <label class="filter-check-row select-all-row">
          <input
            type="checkbox"
            data-role="select-all"
            ${allVisibleChecked ? "checked" : ""}
          />
          <span>(Select All)</span>
        </label>

        <div class="filter-value-list">
          ${visibleValues.length
      ? visibleValues
        .map(
          (item) => `
                    <label class="filter-check-row">
                      <input
                        type="checkbox"
                        data-role="filter-value"
                        value="${escapeAttribute(item.value)}"
                        ${filterMenuState.selected.has(item.value) ? "checked" : ""}
                      />
                      <span class="filter-text">${escapeHtml(item.label)}</span>
                    </label>
                  `
        )
        .join("")
      : `<div class="filter-empty">Không có giá trị</div>`
    }
        </div>
      </div>

      <div class="filter-actions">
        <button type="button" class="filter-btn-secondary" data-action="clear">Clear</button>
        <div class="filter-actions-spacer"></div>
        <button type="button" class="filter-btn-secondary" data-action="cancel">Cancel</button>
        <button type="button" class="filter-btn-primary" data-action="ok">OK</button>
      </div>
    </div>
  `;

  popup.classList.remove("hidden");
  positionFilterPopup();

  const selectAllCheckbox = popup.querySelector('[data-role="select-all"]');
  if (selectAllCheckbox) {
    selectAllCheckbox.indeterminate =
      !allVisibleChecked && !noneVisibleChecked;
  }
}

function positionFilterPopup() {
  const popup = document.getElementById("excelFilterPopup");
  if (!popup || !filterMenuAnchor || popup.classList.contains("hidden")) return;

  const rect = filterMenuAnchor.getBoundingClientRect();

  popup.style.left = `${window.scrollX + rect.left}px`;
  popup.style.top = `${window.scrollY + rect.bottom + 6}px`;

  const popupRect = popup.getBoundingClientRect();
  const maxLeft = window.scrollX + window.innerWidth - popupRect.width - 12;

  if (window.scrollX + rect.left > maxLeft) {
    popup.style.left = `${Math.max(12, maxLeft)}px`;
  }
}

function handleFilterPopupInput(event) {
  if (!filterMenuState) return;

  if (event.target.matches(".filter-search-input")) {
    filterMenuState.search = event.target.value || "";
    renderFilterPopup();
  }
}

function handleFilterPopupChange(event) {
  if (!filterMenuState) return;

  const target = event.target;

  if (target.matches('[data-role="select-all"]')) {
    const visibleValues = getVisibleFilterOptions();

    if (target.checked) {
      visibleValues.forEach((item) => filterMenuState.selected.add(item.value));
    } else {
      visibleValues.forEach((item) => filterMenuState.selected.delete(item.value));
    }

    renderFilterPopup();
    return;
  }

  if (target.matches('[data-role="filter-value"]')) {
    const value = target.value;

    if (target.checked) {
      filterMenuState.selected.add(value);
    } else {
      filterMenuState.selected.delete(value);
    }

    renderFilterPopup();
  }
}

function handleFilterPopupClick(event) {
  if (!filterMenuState) return;

  const actionButton = event.target.closest("[data-action]");
  if (!actionButton) return;

  const action = actionButton.dataset.action;

  if (action === "cancel") {
    closeFilterMenu();
    return;
  }

  if (action === "clear") {
    delete activeColumnFilters[filterMenuState.columnKey];
    applyColumnFilters();
    renderMatchResultsTable();
    closeFilterMenu();
    return;
  }

  if (action === "ok") {
    applyFilterMenuSelection();
  }
}

function applyFilterMenuSelection() {
  if (!filterMenuState) return;

  const allValues = getDistinctColumnValues(filterMenuState.columnKey).map(
    (item) => item.value
  );

  const allSelected =
    filterMenuState.selected.size === allValues.length &&
    allValues.every((value) => filterMenuState.selected.has(value));

  if (allSelected) {
    delete activeColumnFilters[filterMenuState.columnKey];
  } else {
    activeColumnFilters[filterMenuState.columnKey] = new Set(filterMenuState.selected);
  }

  applyColumnFilters();
  renderMatchResultsTable();
  closeFilterMenu();
}

/* =========================
   FILTER CHO TAB DỮ LIỆU EXCEL
========================= */
function normalizeExcelFilterValue(value) {
  return String(value ?? "").trim();
}

function isExcelHeaderFilterable(header) {
  return header !== "Sketch Design";
}

function hasExcelActiveFilter(header) {
  return excelColumnFilters[header] instanceof Set;
}

function hasAnyExcelActiveFilter() {
  return Object.keys(excelColumnFilters).some(
    (key) => excelColumnFilters[key] instanceof Set
  );
}

function getCombinedExcelHeaders(sheets) {
  const headerSet = new Set();
  const headers = [];

  sheets.forEach((sheet) => {
    (sheet.headers || []).forEach((header) => {
      if (!headerSet.has(header)) {
        headerSet.add(header);
        headers.push(header);
      }
    });
  });

  return [...headers, "Chi tiết"];
}

function getCombinedExcelRows(sheets) {
  const rows = [];

  sheets.forEach((sheet, sourceIndex) => {
    (sheet.rows || []).forEach((row) => {
      rows.push({
        ...row,
        __sheetName: sheet.sheetName,
        __sourceIndex: sourceIndex,
      });
    });
  });

  return rows;
}

function getExcelCellValueByHeader(row, header) {
  if (header === "Chi tiết") {
    return row.__folderName || "";
  }

  return row?.[header] ?? "";
}

function rowMatchesExcelColumnFilters(row, headers, ignoreHeader = null) {
  const availableHeaders = new Set([...(headers || []), "Chi tiết"]);

  for (const [header, selectedValues] of Object.entries(excelColumnFilters)) {
    if (!(selectedValues instanceof Set)) continue;
    if (header === ignoreHeader) continue;
    if (!availableHeaders.has(header)) continue;

    const currentValue = normalizeExcelFilterValue(
      getExcelCellValueByHeader(row, header)
    );

    if (!selectedValues.has(currentValue)) {
      return false;
    }
  }

  return true;
}

function getExcelFilterBaseRows(header) {
  if (!workbookData || !Array.isArray(workbookData.sheets)) return [];

  const combinedHeaders = getCombinedExcelHeaders(workbookData.sheets);
  const combinedRows = getCombinedExcelRows(workbookData.sheets);

  return combinedRows.filter((row) => {
    return (
      rowMatchesKeyword(row, combinedHeaders, currentExcelSearch) &&
      rowMatchesExcelColumnFilters(row, combinedHeaders, header)
    );
  });
}

function getDistinctExcelColumnValues(header) {
  const rows = getExcelFilterBaseRows(header);
  const map = new Map();

  rows.forEach((row) => {
    const raw = normalizeExcelFilterValue(getExcelCellValueByHeader(row, header));
    const label = raw || "(Trống)";

    if (!map.has(raw)) {
      map.set(raw, {
        value: raw,
        label,
        count: 0,
      });
    }

    map.get(raw).count += 1;
  });

  return Array.from(map.values()).sort((a, b) =>
    a.label.localeCompare(b.label, "vi", {
      numeric: true,
      sensitivity: "base",
    })
  );
}

function ensureExcelFilterPopup() {
  let popup = document.getElementById("excelDataFilterPopup");

  if (!popup) {
    popup = document.createElement("div");
    popup.id = "excelDataFilterPopup";
    popup.className = "excel-filter-popup hidden";
    document.body.appendChild(popup);

    popup.addEventListener("input", handleExcelFilterPopupInput);
    popup.addEventListener("change", handleExcelFilterPopupChange);
    popup.addEventListener("click", handleExcelFilterPopupClick);
  }

  return popup;
}

function openExcelFilterMenu(header, anchorEl) {
  excelFilterMenuAnchor = anchorEl;

  const distinctValues = getDistinctExcelColumnValues(header);
  const selected =
    excelColumnFilters[header] instanceof Set
      ? new Set(excelColumnFilters[header])
      : new Set(distinctValues.map((item) => item.value));

  excelFilterMenuState = {
    header,
    search: "",
    selected,
  };

  renderExcelFilterPopup();
}

function closeExcelFilterMenu() {
  excelFilterMenuState = null;
  excelFilterMenuAnchor = null;

  const popup = document.getElementById("excelDataFilterPopup");
  if (popup) {
    popup.classList.add("hidden");
    popup.innerHTML = "";
  }
}

function getVisibleExcelFilterOptions() {
  if (!excelFilterMenuState) return [];

  const allValues = getDistinctExcelColumnValues(excelFilterMenuState.header);
  const keyword = excelFilterMenuState.search.trim().toLowerCase();

  if (!keyword) return allValues;

  return allValues.filter((item) =>
    item.label.toLowerCase().includes(keyword)
  );
}

function renderExcelFilterPopup() {
  const popup = ensureExcelFilterPopup();

  if (!excelFilterMenuState) {
    popup.classList.add("hidden");
    popup.innerHTML = "";
    return;
  }

  const visibleValues = getVisibleExcelFilterOptions();

  const checkedVisibleCount = visibleValues.filter((item) =>
    excelFilterMenuState.selected.has(item.value)
  ).length;

  const allVisibleChecked =
    visibleValues.length > 0 && checkedVisibleCount === visibleValues.length;

  const noneVisibleChecked = checkedVisibleCount === 0;

  popup.innerHTML = `
    <div class="filter-panel">
      <div class="filter-panel-title">${escapeHtml(excelFilterMenuState.header)}</div>

      <input
        type="text"
        class="filter-search-input"
        placeholder="Search"
        value="${escapeAttribute(excelFilterMenuState.search)}"
      />

      <div class="filter-value-box">
        <label class="filter-check-row select-all-row">
          <input
            type="checkbox"
            data-role="select-all"
            ${allVisibleChecked ? "checked" : ""}
          />
          <span>(Select All)</span>
        </label>

        <div class="filter-value-list">
          ${visibleValues.length
      ? visibleValues
        .map(
          (item) => `
                    <label class="filter-check-row">
                      <input
                        type="checkbox"
                        data-role="filter-value"
                        value="${escapeAttribute(item.value)}"
                        ${excelFilterMenuState.selected.has(item.value) ? "checked" : ""}
                      />
                      <span class="filter-text">${escapeHtml(item.label)}</span>
                    </label>
                  `
        )
        .join("")
      : `<div class="filter-empty">Không có giá trị</div>`
    }
        </div>
      </div>

      <div class="filter-actions">
        <button type="button" class="filter-btn-secondary" data-action="clear">Clear</button>
        <div class="filter-actions-spacer"></div>
        <button type="button" class="filter-btn-secondary" data-action="cancel">Cancel</button>
        <button type="button" class="filter-btn-primary" data-action="ok">OK</button>
      </div>
    </div>
  `;

  popup.classList.remove("hidden");
  positionExcelFilterPopup();

  const selectAllCheckbox = popup.querySelector('[data-role="select-all"]');
  if (selectAllCheckbox) {
    selectAllCheckbox.indeterminate =
      !allVisibleChecked && !noneVisibleChecked;
  }
}

function positionExcelFilterPopup() {
  const popup = document.getElementById("excelDataFilterPopup");
  if (!popup || !excelFilterMenuAnchor || popup.classList.contains("hidden")) return;

  const rect = excelFilterMenuAnchor.getBoundingClientRect();

  popup.style.left = `${window.scrollX + rect.left}px`;
  popup.style.top = `${window.scrollY + rect.bottom + 6}px`;

  const popupRect = popup.getBoundingClientRect();
  const maxLeft = window.scrollX + window.innerWidth - popupRect.width - 12;

  if (window.scrollX + rect.left > maxLeft) {
    popup.style.left = `${Math.max(12, maxLeft)}px`;
  }
}

function handleExcelFilterPopupInput(event) {
  if (!excelFilterMenuState) return;

  if (event.target.matches(".filter-search-input")) {
    excelFilterMenuState.search = event.target.value || "";
    renderExcelFilterPopup();
  }
}

function handleExcelFilterPopupChange(event) {
  if (!excelFilterMenuState) return;

  const target = event.target;

  if (target.matches('[data-role="select-all"]')) {
    const visibleValues = getVisibleExcelFilterOptions();

    if (target.checked) {
      visibleValues.forEach((item) => excelFilterMenuState.selected.add(item.value));
    } else {
      visibleValues.forEach((item) => excelFilterMenuState.selected.delete(item.value));
    }

    renderExcelFilterPopup();
    return;
  }

  if (target.matches('[data-role="filter-value"]')) {
    const value = target.value;

    if (target.checked) {
      excelFilterMenuState.selected.add(value);
    } else {
      excelFilterMenuState.selected.delete(value);
    }

    renderExcelFilterPopup();
  }
}

function handleExcelFilterPopupClick(event) {
  if (!excelFilterMenuState) return;

  const actionButton = event.target.closest("[data-action]");
  if (!actionButton) return;

  const action = actionButton.dataset.action;

  if (action === "cancel") {
    closeExcelFilterMenu();
    return;
  }

  if (action === "clear") {
    delete excelColumnFilters[excelFilterMenuState.header];
    renderSheets(workbookData?.sheets || []);
    closeExcelFilterMenu();
    return;
  }

  if (action === "ok") {
    applyExcelFilterMenuSelection();
  }
}

function applyExcelFilterMenuSelection() {
  if (!excelFilterMenuState) return;

  const allValues = getDistinctExcelColumnValues(
    excelFilterMenuState.header
  ).map((item) => item.value);

  const allSelected =
    excelFilterMenuState.selected.size === allValues.length &&
    allValues.every((value) => excelFilterMenuState.selected.has(value));

  if (allSelected) {
    delete excelColumnFilters[excelFilterMenuState.header];
  } else {
    excelColumnFilters[excelFilterMenuState.header] = new Set(
      excelFilterMenuState.selected
    );
  }

  renderSheets(workbookData?.sheets || []);
  closeExcelFilterMenu();
}

function createExcelTableHeader(header) {
  const filterable = isExcelHeaderFilterable(header);
  const active = hasExcelActiveFilter(header);

  return `
    <th class="${active ? "filter-active-col" : ""}">
      <div class="th-wrap">
        <span class="th-label">${escapeHtml(header)}</span>
        ${filterable
      ? `
            <button
              type="button"
              class="filter-trigger excel-data-filter-trigger ${active ? "active" : ""}"
              data-excel-filter-key="${escapeAttribute(header)}"
              title="Lọc cột ${escapeAttribute(header)}"
            >
              ▾
            </button>
          `
      : ""
    }
      </div>
    </th>
  `;
}

/* =========================
   EXCEL SEARCH + RENDER
========================= */
function handleExcelSearch() {
  currentExcelSearch = excelSearchInput.value.trim().toLowerCase();
  activeSheetPanelId = null;

  if (!workbookData) return;

  renderSheets(workbookData.sheets);
  activatePanel("dataCard");
}

function clearExcelSearch() {
  excelSearchInput.value = "";
  currentExcelSearch = "";
  activeSheetPanelId = null;

  if (!workbookData) return;

  renderSheets(workbookData.sheets);
  activatePanel("dataCard");
}

function buildRowSearchText(row, headers) {
  const values = [];

  headers.forEach((header) => {
    const value = row[header];
    if (value !== null && value !== undefined && value !== "") {
      values.push(String(value));
    }
  });

  if (row.__folderName) {
    values.push(String(row.__folderName));
  }

  if (row.__excelRow) {
    values.push(String(row.__excelRow));
  }

  return values.join(" ").toLowerCase();
}

function rowMatchesKeyword(row, headers, keyword) {
  if (!keyword) return true;
  return buildRowSearchText(row, headers).includes(keyword);
}

function renderSheets(sheets) {
  if (!Array.isArray(sheets) || sheets.length === 0) {
    excelDataContainer.innerHTML = "<p>Không có dữ liệu để hiển thị</p>";
    sheetTabs.innerHTML = "";
    dataCard?.classList.remove("hidden");
    updateSidebarAvailability();
    return;
  }

  dataCard?.classList.remove("hidden");
  updateSidebarAvailability();

  const combinedHeaders = getCombinedExcelHeaders(sheets);
  const combinedRows = getCombinedExcelRows(sheets);
  const hasExcelFilters = hasAnyExcelActiveFilter();

  const filteredRows = combinedRows.filter((row) =>
    rowMatchesKeyword(row, combinedHeaders, currentExcelSearch) &&
    rowMatchesExcelColumnFilters(row, combinedHeaders)
  );

  if (currentExcelSearch || hasExcelFilters) {
    if (filteredRows.length > 0) {
      let status = `Tìm thấy ${filteredRows.length} dòng phù hợp.`;
      if (hasExcelFilters) {
        status += ` Đang áp dụng ${Object.keys(excelColumnFilters).length} filter cột.`;
      }
      excelSearchStatus.textContent = status;
    } else {
      excelSearchStatus.textContent = "Không có dữ liệu phù hợp với điều kiện lọc hiện tại.";
    }
  } else {
    excelSearchStatus.textContent =
      "Nhập từ khóa để lọc dữ liệu theo tất cả cột text.";
  }

  sheetTabs.innerHTML = "";

  const rowsToRender =
    currentExcelSearch || hasExcelFilters ? filteredRows : combinedRows;

  const tableHead = combinedHeaders
    .map((header) => createExcelTableHeader(header))
    .join("");

  const tableBody = rowsToRender
    .map((row) => {
      const cells = combinedHeaders.map((header) => createTableCell(header, row)).join("");
      return `<tr>${cells}</tr>`;
    })
    .join("");

  excelDataContainer.innerHTML = `
    <div class="sheet-panel active" id="sheet-combined">
      <p class="sheet-meta">
        Tổng số dòng: <strong>${escapeHtml(rowsToRender.length)}</strong>
      </p>

      <div class="table-wrapper">
        <table class="data-table">
          <thead>
            <tr>${tableHead}</tr>
          </thead>
          <tbody>
            ${tableBody || `<tr><td colspan="${combinedHeaders.length}" class="empty-value">Không có dữ liệu</td></tr>`}
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
      .map(
        (src, index) =>
          `<img src="${escapeAttribute(src)}" alt="Sketch ${index + 1}" class="previewable-image" />`
      )
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

/* =========================
   UTIL
========================= */
function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function escapeAttribute(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;");
}



