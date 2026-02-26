"use strict";

const STORAGE_KEYS = {
  shipments: "logistics_shipments_v1",
  finance: "logistics_finance_v1",
  dispatchRecords: "logistics_dispatch_records_v1",
  homeSettleActionRecords: "logistics_home_settle_action_records_v1",
  catalogProfiles: "logistics_catalog_profiles_v1",
  shipmentRetain: "logistics_shipment_retain_v1",
  financeRetain: "logistics_finance_retain_v1",
};
const LOCAL_META_KEYS = {
  manualSaveState: "logistics_manual_save_state_v1",
};
const FILE_STORAGE_ENDPOINT_PREFIX = "/api/storage/";
const FILE_STORAGE_HEALTH_ENDPOINT = "/api/health";
const FILE_STORAGE_KEYS = new Set(Object.values(STORAGE_KEYS));
const FILE_STORAGE_FLUSH_MS = 180;
const MANUAL_STORAGE_FILENAME_PREFIX = "顺源物流数据包";

const TYPE_LABELS = {
  receive: "收货",
  ship: "发货",
  income: "收入",
  expense: "支出",
};

const MEASURE_UNIT_LABELS = {
  piece: "件",
  ton: "KG",
  cubic: "平方",
};
const DEFAULT_MEASURE_UNIT = "piece";
const STORAGE_PACKAGE_VERSION = "1.0.0";
const STORAGE_PACKAGE_LABELS = {
  [STORAGE_KEYS.shipments]: "收发货记录",
  [STORAGE_KEYS.finance]: "收支记录",
  [STORAGE_KEYS.dispatchRecords]: "配车记录",
  [STORAGE_KEYS.homeSettleActionRecords]: "家结处理记录",
  [STORAGE_KEYS.catalogProfiles]: "厂家/客户/产品匹配",
  [STORAGE_KEYS.shipmentRetain]: "收发表单保留项",
  [STORAGE_KEYS.financeRetain]: "收支表单保留项",
};

const PAYMENT_METHOD_LABELS = {
  signback: "签回单",
  pickup: "自提",
  delivery: "送货",
  notify: "等通知发货",
};

const HOME_SETTLE_REMINDER_MONTHS = 1;
const HOME_SETTLE_SNOOZE_MS = 60 * 60 * 1000;
const HOME_SETTLE_ACTION_LABELS = {
  snooze: "稍后处理(1小时)",
  contact: "马上联系",
  paid: "已收款",
};
const HOME_SETTLE_STATUS_CATALOG = [
  { id: "unpaid", label: "待收款", group: "pending" },
  { id: "reminded", label: "已提醒", group: "active" },
  { id: "paid", label: "已收款", group: "done" },
];
const HOME_SETTLE_STATUS_BY_ID = new Map(HOME_SETTLE_STATUS_CATALOG.map((item) => [item.id, item]));
const HOME_SETTLE_STATUS_IDS = new Set(HOME_SETTLE_STATUS_CATALOG.map((item) => item.id));

const MAIN_PAGE_SIZE = 10;
const SUMMARY_PAGE_SIZE = 10;
const DELETE_UNDO_MS = 5000;
const FORM_SUGGESTION_LIMIT = 16;
const SHIPMENT_PAYMENT_FINANCE_SOURCE = "shipment_payment";
const SHIPMENT_PAYMENT_FINANCE_SUMMARY_PREFIX = "货物收款";
const SHIPMENT_PAYMENT_FINANCE_NOTE_PREFIX = "来自收发货已收款记录";
const SHIPMENT_PAYMENT_FINANCE_SUMMARY_PREFIX_LEGACY = "货物收款（系统同步）";
const SHIPMENT_PAYMENT_FINANCE_NOTE_PREFIX_LEGACY = "系统同步：来自收发货已收款记录";
const SHIPMENT_RETAIN_FIELDS = [
  "type",
  "measureUnit",
  "manufacturer",
  "customer",
  "product",
  "unitPrice",
  "transit",
  "delivery",
  "pickup",
  "note",
];
const FINANCE_RETAIN_FIELDS = ["category", "note"];
const SHIPMENT_INLINE_EDIT_FIELDS = new Set([
  "date",
  "serialNo",
  "manufacturer",
  "customer",
  "product",
  "quantity",
  "measureUnit",
  "unitPrice",
  "amount",
  "weight",
  "cubic",
  "transit",
  "delivery",
  "pickup",
  "rebate",
  "note",
]);
const SHIPMENT_INLINE_EDIT_NUMERIC_CONFIG = {
  quantity: { step: "0.001", min: "0", digits: 3 },
  unitPrice: { step: "0.01", min: "0", digits: 2 },
  amount: { step: "0.01", min: "0", digits: 2 },
  weight: { step: "0.001", min: "0", digits: 3 },
  cubic: { step: "0.001", min: "0", digits: 3 },
  rebate: { step: "0.01", min: "0", digits: 2 },
};

const shipmentForm = document.getElementById("shipment-form");
const financeForm = document.getElementById("finance-form");

const shipmentFilter = document.getElementById("shipment-filter");
const shipmentSearch = document.getElementById("shipment-search");
const shipmentManufacturerFilter = document.getElementById("shipment-manufacturer-filter");
const dispatchFilter = document.getElementById("dispatch-filter");
const dispatchSearch = document.getElementById("dispatch-search");
const financeFilter = document.getElementById("finance-filter");
const financeSearch = document.getElementById("finance-search");
const financeMonth = document.getElementById("finance-month");
const statusFilter = document.getElementById("status-filter");
const statusSearch = document.getElementById("status-search");
const statusFromDateInput = document.getElementById("status-from-date");
const statusToDateInput = document.getElementById("status-to-date");
const statusClearDateBtn = document.getElementById("status-clear-date");
const statusStatTotal = document.getElementById("status-stat-total");
const statusStatPending = document.getElementById("status-stat-pending");
const statusStatActive = document.getElementById("status-stat-active");
const statusStatDone = document.getElementById("status-stat-done");
const financeExportSelectedBtn = document.getElementById("finance-export-selected");
const financeExportSelectAllCheckbox = document.getElementById("finance-export-select-all");
const financeExportPreviewModal = document.getElementById("finance-export-preview-modal");
const financeExportPreviewTip = document.getElementById("finance-export-preview-tip");
const financeExportPreviewBody = document.getElementById("finance-export-preview-body");
const financeExportPreviewCloseBtn = document.getElementById("finance-export-preview-close");
const financeExportPreviewCancelBtn = document.getElementById("finance-export-preview-cancel");
const financeExportPreviewConfirmBtn = document.getElementById("finance-export-preview-confirm");
const shipmentExportFilteredSummaryBtn = document.getElementById("shipment-export-filtered-summary");
const shipmentExportSelectAllCheckbox = document.getElementById("shipment-export-select-all");
const shipmentExportPreviewModal = document.getElementById("shipment-export-preview-modal");
const shipmentExportPreviewTip = document.getElementById("shipment-export-preview-tip");
const shipmentExportPreviewBody = document.getElementById("shipment-export-preview-body");
const shipmentExportPreviewCloseBtn = document.getElementById("shipment-export-preview-close");
const shipmentExportPreviewCancelBtn = document.getElementById("shipment-export-preview-cancel");
const shipmentExportPreviewConfirmBtn = document.getElementById("shipment-export-preview-confirm");
const shipmentBatchImportBtn = document.getElementById("shipment-batch-import-btn");
const shipmentBatchImportFileInput = document.getElementById("shipment-batch-import-file");
const shipmentImportPreviewModal = document.getElementById("shipment-import-preview-modal");
const shipmentImportPreviewTip = document.getElementById("shipment-import-preview-tip");
const shipmentImportPreviewBody = document.getElementById("shipment-import-preview-body");
const shipmentImportPreviewCloseBtn = document.getElementById("shipment-import-preview-close");
const shipmentImportPreviewCancelBtn = document.getElementById("shipment-import-preview-cancel");
const shipmentImportPreviewConfirmBtn = document.getElementById("shipment-import-preview-confirm");
const storageExportPackageBtn = document.getElementById("storage-export-package-btn");
const storageImportPackageBtn = document.getElementById("storage-import-package-btn");
const storageImportPackageFileInput = document.getElementById("storage-import-package-file");
const storageImportPreviewModal = document.getElementById("storage-import-preview-modal");
const storageImportPreviewTip = document.getElementById("storage-import-preview-tip");
const storageImportPreviewBody = document.getElementById("storage-import-preview-body");
const storageImportPreviewCloseBtn = document.getElementById("storage-import-preview-close");
const storageImportPreviewCancelBtn = document.getElementById("storage-import-preview-cancel");
const storageImportPreviewConfirmBtn = document.getElementById("storage-import-preview-confirm");

const shipmentTableBody = document.getElementById("shipment-table-body");
const dispatchTableBody = document.getElementById("dispatch-table-body");
const financeTableBody = document.getElementById("finance-table-body");
const statusTableBody = document.getElementById("status-table-body");
const summaryShipmentBody = document.getElementById("summary-shipment-body");
const summaryFinanceBody = document.getElementById("summary-finance-body");


const currentTimeEl = document.getElementById("current-time");
const globalSearchInput = document.getElementById("global-search-input");
const storageModeBadge = document.getElementById("storage-mode-badge");
const panels = Array.from(document.querySelectorAll("[data-panel]"));
const sidebarOpenButtons = Array.from(document.querySelectorAll("[data-open-panel]"));
const summaryForm = document.getElementById("summary-form");
const summaryFromDateInput = document.getElementById("summary-from-date");
const summaryToDateInput = document.getElementById("summary-to-date");
const summaryResetRangeBtn = document.getElementById("summary-reset-range");
const summaryShipmentPrevBtn = document.getElementById("summary-shipment-prev");
const summaryShipmentNextBtn = document.getElementById("summary-shipment-next");
const summaryShipmentPageInfo = document.getElementById("summary-shipment-page-info");
const summaryFinancePrevBtn = document.getElementById("summary-finance-prev");
const summaryFinanceNextBtn = document.getElementById("summary-finance-next");
const summaryFinancePageInfo = document.getElementById("summary-finance-page-info");
const summaryShipmentSearchInput = document.getElementById("summary-shipment-search");
const summaryShipmentClearSearchBtn = document.getElementById("summary-shipment-clear-search");
const summaryFinanceSearchInput = document.getElementById("summary-finance-search");
const summaryFinanceClearSearchBtn = document.getElementById("summary-finance-clear-search");
const summaryQuickRangeButtons = Array.from(document.querySelectorAll("[data-summary-quick-range]"));
const summaryViewButtons = Array.from(document.querySelectorAll("[data-summary-view]"));
const summaryGroups = Array.from(document.querySelectorAll("[data-summary-group]"));
const shipmentFromDateInput = document.getElementById("shipment-from-date");
const shipmentToDateInput = document.getElementById("shipment-to-date");
const shipmentClearDateBtn = document.getElementById("shipment-clear-date");
const financeFromDateInput = document.getElementById("finance-from-date");
const financeToDateInput = document.getElementById("finance-to-date");
const financeClearDateBtn = document.getElementById("finance-clear-date");
const shipmentPrevBtn = document.getElementById("shipment-prev");
const shipmentNextBtn = document.getElementById("shipment-next");
const shipmentPageInfo = document.getElementById("shipment-page-info");
const dispatchPrevBtn = document.getElementById("dispatch-prev");
const dispatchNextBtn = document.getElementById("dispatch-next");
const dispatchPageInfo = document.getElementById("dispatch-page-info");
const dispatchLoadBtn = document.getElementById("dispatch-load-btn");
const dispatchResetStepBtn = document.getElementById("dispatch-reset-step-btn");
const dispatchSelectedCount = document.getElementById("dispatch-selected-count");
const dispatchSelectAllLoad = document.getElementById("dispatch-select-all-load");
const dispatchMetaModal = document.getElementById("dispatch-meta-modal");
const dispatchMetaModalTip = document.getElementById("dispatch-meta-modal-tip");
const dispatchMetaModalCloseBtn = document.getElementById("dispatch-meta-modal-close");
const dispatchMetaCancelBtn = document.getElementById("dispatch-meta-cancel-btn");
const dispatchMetaConfirmBtn = document.getElementById("dispatch-meta-confirm-btn");
const dispatchMetaDateInput = document.getElementById("dispatch-meta-date");
const dispatchMetaTruckNoInput = document.getElementById("dispatch-meta-truck-no");
const dispatchMetaPlateNoInput = document.getElementById("dispatch-meta-plate-no");
const dispatchMetaDriverInput = document.getElementById("dispatch-meta-driver");
const dispatchMetaContactNameInput = document.getElementById("dispatch-meta-contact-name");
const statusPrevBtn = document.getElementById("status-prev");
const statusNextBtn = document.getElementById("status-next");
const statusPageInfo = document.getElementById("status-page-info");
const statusActionRecordFilter = document.getElementById("status-action-record-filter");
const statusActionRecordSearch = document.getElementById("status-action-record-search");
const statusActionRecordFromDateInput = document.getElementById("status-action-record-from-date");
const statusActionRecordToDateInput = document.getElementById("status-action-record-to-date");
const statusActionRecordClearDateBtn = document.getElementById("status-action-record-clear-date");
const statusActionRecordTableBody = document.getElementById("status-action-record-table-body");
const statusActionRecordPrevBtn = document.getElementById("status-action-record-prev");
const statusActionRecordNextBtn = document.getElementById("status-action-record-next");
const statusActionRecordPageInfo = document.getElementById("status-action-record-page-info");
const statusDetailModal = document.getElementById("status-detail-modal");
const statusDetailModalCloseBtn = document.getElementById("status-detail-modal-close");
const statusDetailModalContent = document.getElementById("status-detail-modal-content");
const confirmModal = document.getElementById("confirm-modal");
const confirmModalTitle = document.getElementById("confirm-modal-title");
const confirmModalMessage = document.getElementById("confirm-modal-message");
const confirmModalCloseBtn = document.getElementById("confirm-modal-close");
const confirmModalCancelBtn = document.getElementById("confirm-modal-cancel");
const confirmModalConfirmBtn = document.getElementById("confirm-modal-confirm");
const storageRequiredModal = document.getElementById("storage-required-modal");
const storageRequiredStatus = document.getElementById("storage-required-status");
const storageRequiredRetryBtn = document.getElementById("storage-required-retry");
const storageRequiredContinueBtn = document.getElementById("storage-required-continue");
const dispatchRecordBody = document.getElementById("dispatch-record-body");
const dispatchRecordModal = document.getElementById("dispatch-record-modal");
const dispatchRecordModalTitle = document.getElementById("dispatch-record-modal-title");
const dispatchRecordDeleteBtn = document.getElementById("dispatch-record-delete-btn");
const dispatchRecordExportBtn = document.getElementById("dispatch-record-export-btn");
const dispatchRecordEditBtn = document.getElementById("dispatch-record-edit-btn");
const dispatchRecordModalCloseBtn = document.getElementById("dispatch-record-modal-close");
const dispatchRecordModalSummary = document.getElementById("dispatch-record-modal-summary");
const dispatchRecordDetailBody = document.getElementById("dispatch-record-detail-body");
const dispatchRecordEditPanel = document.getElementById("dispatch-record-edit-panel");
const dispatchRecordAddSearchInput = document.getElementById("dispatch-record-add-search");
const dispatchRecordAddClearBtn = document.getElementById("dispatch-record-add-clear-btn");
const dispatchRecordAddSelectAll = document.getElementById("dispatch-record-add-select-all");
const dispatchRecordAddBody = document.getElementById("dispatch-record-add-body");
const dispatchRecordAddSelectedCount = document.getElementById("dispatch-record-add-selected-count");
const dispatchRecordAddBtn = document.getElementById("dispatch-record-add-btn");
const financePrevBtn = document.getElementById("finance-prev");
const financeNextBtn = document.getElementById("finance-next");
const financePageInfo = document.getElementById("finance-page-info");
const tableSortButtons = Array.from(document.querySelectorAll("[data-table-sort]"));
const undoToast = document.getElementById("undo-toast");
const undoToastText = document.getElementById("undo-toast-text");
const undoDeleteBtn = document.getElementById("undo-delete-btn");
const undoCountdownEl = document.getElementById("undo-countdown");
const openSummaryViewButtons = Array.from(document.querySelectorAll("[data-open-summary-view]"));
const closeSummaryViewButtons = Array.from(document.querySelectorAll("[data-close-summary-view]"));
const shipmentManufacturerList = document.getElementById("shipment-manufacturer-list");
const shipmentCustomerList = document.getElementById("shipment-customer-list");
const shipmentProductList = document.getElementById("shipment-product-list");
const shipmentProductSuggestionPanel = document.getElementById("shipment-product-suggestion-panel");
const shipmentQuantityLabel = document.getElementById("shipment-quantity-label");
const shipmentNoteOptionInputs = Array.from(document.querySelectorAll("input[data-shipment-note-option]"));
const shipmentNoteCustomInput = document.getElementById("shipment-note-custom");

const shipmentFieldEls = {
  measureUnit: shipmentForm?.elements?.measureUnit || null,
  manufacturer: shipmentForm?.elements?.manufacturer || null,
  customer: shipmentForm?.elements?.customer || null,
  product: shipmentForm?.elements?.product || null,
  quantity: shipmentForm?.elements?.quantity || null,
  unitPrice: shipmentForm?.elements?.unitPrice || null,
  amount: shipmentForm?.elements?.amount || null,
  weight: shipmentForm?.elements?.weight || null,
  cubic: shipmentForm?.elements?.cubic || null,
  transit: shipmentForm?.elements?.transit || null,
  delivery: shipmentForm?.elements?.delivery || null,
  pickup: shipmentForm?.elements?.pickup || null,
  rebate: shipmentForm?.elements?.rebate || null,
  note: shipmentForm?.elements?.note || null,
};

const BUILTIN_PRICE_CATALOG_ROWS = normalizeBuiltInCatalogRows(window.BUILTIN_PRICE_CATALOG);

const shipmentStatsEls = {
  receiveQty: document.getElementById("stat-receive-qty"),
  receiveWeight: document.getElementById("stat-receive-weight"),
  shipQty: document.getElementById("stat-ship-qty"),
  shipWeight: document.getElementById("stat-ship-weight"),
  totalFee: document.getElementById("stat-shipment-fee"),
  totalDeclared: document.getElementById("stat-shipment-declared"),
  netQty: document.getElementById("stat-net-qty"),
  netWeight: document.getElementById("stat-net-weight"),
};

const financeStatsEls = {
  todayIncome: document.getElementById("stat-today-income"),
  todayExpense: document.getElementById("stat-today-expense"),
  monthBalance: document.getElementById("stat-month-balance"),
  monthBreakdown: document.getElementById("stat-month-breakdown"),
};

const summaryShipmentStatsEls = {
  days: document.getElementById("summary-ship-stat-days"),
  receiveQty: document.getElementById("summary-ship-stat-receive-qty"),
  receiveWeight: document.getElementById("summary-ship-stat-receive-weight"),
  shipQty: document.getElementById("summary-ship-stat-ship-qty"),
  shipWeight: document.getElementById("summary-ship-stat-ship-weight"),
  netQty: document.getElementById("summary-ship-stat-net-qty"),
  netWeight: document.getElementById("summary-ship-stat-net-weight"),
};

const summaryFinanceStatsEls = {
  days: document.getElementById("summary-fin-stat-days"),
  income: document.getElementById("summary-fin-stat-income"),
  expense: document.getElementById("summary-fin-stat-expense"),
  balance: document.getElementById("summary-fin-stat-balance"),
  records: document.getElementById("summary-fin-stat-records"),
};

let shipments = readStorage(STORAGE_KEYS.shipments);
let financeRecords = readStorage(STORAGE_KEYS.finance);
let dispatchRecords = readStorage(STORAGE_KEYS.dispatchRecords);
let homeSettleActionRecords = sanitizeHomeSettleActionRecords(readStorage(STORAGE_KEYS.homeSettleActionRecords));
let customCatalogProfiles = sanitizeCustomCatalogProfiles(readStorage(STORAGE_KEYS.catalogProfiles));
let activePanelId = "shipment-panel";
let activeSummaryView = "shipment";
let shipmentPage = 1;
let dispatchPage = 1;
let financePage = 1;
let statusPage = 1;
let statusActionRecordPage = 1;
let summaryShipmentPage = 1;
let summaryFinancePage = 1;
let summaryShipmentKeyword = "";
let summaryFinanceKeyword = "";
let shipmentSort = { key: "date", dir: "desc" };
let financeSort = { key: "date", dir: "desc" };
let pendingDeletion = null;
let selectedShipmentIdsForExport = new Set();
let selectedFinanceIdsForExport = new Set();
let selectedDispatchIdsForLoading = new Set();
let confirmedDispatchIdsForLoading = new Set();
let dispatchLoadStage = "select";
let currentVisibleShipmentRows = [];
let currentVisibleFinanceRows = [];
let currentVisibleDispatchRows = [];
let currentVisibleStatusRows = [];
let currentVisibleStatusActionRecordRows = [];
let currentStatusDetailShipmentId = "";
let activeDispatchRecordId = "";
let dispatchRecordEditMode = false;
let dispatchRecordAddKeyword = "";
let dispatchRecordPendingAddIds = new Set();
let pendingFinanceExportContext = null;
let pendingShipmentExportContext = null;
let pendingShipmentImportContext = null;
let pendingStorageImportContext = null;
let shipmentProfileDataset = [];
let shipmentProfileByExactKey = new Map();
let shipmentProfilesByProductKey = new Map();
let shipmentProductSuggestionPool = [];
let manufacturerSuggestionPool = [];
let customerSuggestionPool = [];
let latestProductSuggestionProfiles = [];
let productSuggestionHideTimer = null;
let fileStorageEnabled = false;
let fileStorageReady = false;
let fileStorageDataDir = "";
let fileStorageWriteQueue = new Map();
let fileStorageFlushTimer = 0;
let storageLifecycleBound = false;
let pendingConfirmDialogResolve = null;
let activeShipmentInlineEditSession = null;
let storageGuardDismissed = false;
let manualSaveDirty = false;
let manualSaveLastAt = 0;
let appBootstrapped = false;
const SHIPMENT_NOTE_SEPARATOR = "；";

function restoreManualSaveState() {
  try {
    const raw = localStorage.getItem(LOCAL_META_KEYS.manualSaveState);
    if (!raw) {
      return;
    }
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== "object") {
      return;
    }
    manualSaveLastAt = Math.max(0, toFiniteNumber(parsed.lastSavedAt, 0));
    manualSaveDirty = false;
  } catch {
    manualSaveLastAt = 0;
    manualSaveDirty = false;
  }
}

function persistManualSaveState() {
  try {
    localStorage.setItem(
      LOCAL_META_KEYS.manualSaveState,
      JSON.stringify({
        lastSavedAt: manualSaveLastAt,
      })
    );
  } catch {
    // Ignore quota errors.
  }
}

function markManualSaveDirty() {
  if (!appBootstrapped) {
    return;
  }
  manualSaveDirty = true;
  updateStorageModeBadge();
}

bootstrap();

async function bootstrap() {
  restoreManualSaveState();
  fileStorageEnabled = false;
  fileStorageReady = true;
  storageGuardDismissed = true;
  updateStorageModeBadge();
  initialize();
  appBootstrapped = true;
  manualSaveDirty = false;
  persistManualSaveState();
  updateStorageModeBadge();
}

function initialize() {
  setDefaultDates();
  restoreRetainedFormValues();
  rebuildShipmentProfileIndex();
  setupEventListeners();
  setupStorageSyncLifecycle();
  setupTableMouseInteractions();
  setupShipmentNoteQuickFill();
  setupShipmentSmartFill();
  setupKeyboardShortcuts();
  setupSummaryViewControls();
  updateSortButtonsUI("shipment");
  updateSortButtonsUI("finance");
  setupSummaryLayoutViews();
  setupPanelControls();
  tickClock();
  setInterval(tickClock, 60_000);
  syncFinanceIncomeFromPaidShipments();
  renderAll();
}

function persistSanitizedState() {
  customCatalogProfiles = sanitizeCustomCatalogProfiles(customCatalogProfiles);
  homeSettleActionRecords = sanitizeHomeSettleActionRecords(homeSettleActionRecords);
  writeStorage(STORAGE_KEYS.shipments, shipments);
  writeStorage(STORAGE_KEYS.finance, financeRecords);
  writeStorage(STORAGE_KEYS.dispatchRecords, dispatchRecords);
  writeStorage(STORAGE_KEYS.homeSettleActionRecords, homeSettleActionRecords);
  writeStorage(STORAGE_KEYS.catalogProfiles, customCatalogProfiles);
}

function setupEventListeners() {
  shipmentForm.addEventListener("submit", handleShipmentSubmit);
  financeForm.addEventListener("submit", handleFinanceSubmit);
  bindFormRetention(shipmentForm, STORAGE_KEYS.shipmentRetain, SHIPMENT_RETAIN_FIELDS);
  bindFormRetention(financeForm, STORAGE_KEYS.financeRetain, FINANCE_RETAIN_FIELDS);

  shipmentFilter.addEventListener("change", () => renderShipmentTable({ resetPage: true }));
  shipmentSearch.addEventListener("input", () => {
    if (activePanelId === "shipment-panel" && globalSearchInput) {
      globalSearchInput.value = shipmentSearch.value;
    }
    renderShipmentTable({ resetPage: true });
  });
  if (shipmentManufacturerFilter) {
    shipmentManufacturerFilter.addEventListener("input", () => {
      renderShipmentTable({ resetPage: true });
    });
  }
  if (dispatchFilter) {
    dispatchFilter.addEventListener("change", () => renderDispatchTable({ resetPage: true }));
  }
  if (dispatchSearch) {
    dispatchSearch.addEventListener("input", () => {
      if (activePanelId === "dispatch-panel" && globalSearchInput) {
        globalSearchInput.value = dispatchSearch.value;
      }
      renderDispatchTable({ resetPage: true });
    });
  }

  financeFilter.addEventListener("change", () => renderFinanceTable({ resetPage: true }));
  financeSearch.addEventListener("input", () => {
    if (activePanelId === "finance-panel" && globalSearchInput) {
      globalSearchInput.value = financeSearch.value;
    }
    renderFinanceTable({ resetPage: true });
  });
  financeMonth.addEventListener("change", () => renderFinanceTable({ resetPage: true }));
  if (shipmentFromDateInput) {
    shipmentFromDateInput.addEventListener("change", () => {
      normalizeDateInputs(shipmentFromDateInput, shipmentToDateInput);
      renderShipmentTable({ resetPage: true });
    });
  }
  if (shipmentToDateInput) {
    shipmentToDateInput.addEventListener("change", () => {
      normalizeDateInputs(shipmentFromDateInput, shipmentToDateInput);
      renderShipmentTable({ resetPage: true });
    });
  }
  if (shipmentClearDateBtn) {
    shipmentClearDateBtn.addEventListener("click", clearShipmentDateRange);
  }
  if (financeFromDateInput) {
    financeFromDateInput.addEventListener("change", () => {
      normalizeDateInputs(financeFromDateInput, financeToDateInput);
      renderFinanceTable({ resetPage: true });
    });
  }
  if (financeToDateInput) {
    financeToDateInput.addEventListener("change", () => {
      normalizeDateInputs(financeFromDateInput, financeToDateInput);
      renderFinanceTable({ resetPage: true });
    });
  }
  if (financeClearDateBtn) {
    financeClearDateBtn.addEventListener("click", clearFinanceDateRange);
  }
  if (statusFilter) {
    statusFilter.addEventListener("change", () => renderStatusSection({ resetPage: true }));
  }
  if (statusSearch) {
    statusSearch.addEventListener("input", () => {
      if (activePanelId === "status-panel" && globalSearchInput) {
        globalSearchInput.value = statusSearch.value;
      }
      renderStatusSection({ resetPage: true });
    });
  }
  if (statusFromDateInput) {
    statusFromDateInput.addEventListener("change", () => {
      normalizeDateInputs(statusFromDateInput, statusToDateInput);
      renderStatusSection({ resetPage: true });
    });
  }
  if (statusToDateInput) {
    statusToDateInput.addEventListener("change", () => {
      normalizeDateInputs(statusFromDateInput, statusToDateInput);
      renderStatusSection({ resetPage: true });
    });
  }
  if (statusClearDateBtn) {
    statusClearDateBtn.addEventListener("click", clearStatusDateRange);
  }
  if (statusActionRecordFilter) {
    statusActionRecordFilter.addEventListener("change", () => renderStatusSection({ resetRecordPage: true }));
  }
  if (statusActionRecordSearch) {
    statusActionRecordSearch.addEventListener("input", () => renderStatusSection({ resetRecordPage: true }));
  }
  if (statusActionRecordFromDateInput) {
    statusActionRecordFromDateInput.addEventListener("change", () => {
      normalizeDateInputs(statusActionRecordFromDateInput, statusActionRecordToDateInput);
      renderStatusSection({ resetRecordPage: true });
    });
  }
  if (statusActionRecordToDateInput) {
    statusActionRecordToDateInput.addEventListener("change", () => {
      normalizeDateInputs(statusActionRecordFromDateInput, statusActionRecordToDateInput);
      renderStatusSection({ resetRecordPage: true });
    });
  }
  if (statusActionRecordClearDateBtn) {
    statusActionRecordClearDateBtn.addEventListener("click", clearStatusActionRecordDateRange);
  }
  if (financeExportSelectedBtn) {
    financeExportSelectedBtn.addEventListener("click", handleFinanceExportSelectedClick);
  }
  if (financeExportSelectAllCheckbox) {
    financeExportSelectAllCheckbox.addEventListener("change", handleFinanceExportSelectAllChange);
  }
  if (shipmentPrevBtn) {
    shipmentPrevBtn.addEventListener("click", () => changeMainPage("shipment", -1));
  }
  if (shipmentNextBtn) {
    shipmentNextBtn.addEventListener("click", () => changeMainPage("shipment", 1));
  }
  if (dispatchPrevBtn) {
    dispatchPrevBtn.addEventListener("click", () => changeMainPage("dispatch", -1));
  }
  if (dispatchNextBtn) {
    dispatchNextBtn.addEventListener("click", () => changeMainPage("dispatch", 1));
  }
  if (dispatchLoadBtn) {
    dispatchLoadBtn.addEventListener("click", handleBatchDispatchLoad);
  }
  if (dispatchResetStepBtn) {
    dispatchResetStepBtn.addEventListener("click", handleDispatchLoadStageReset);
  }
  if (dispatchMetaModalCloseBtn) {
    dispatchMetaModalCloseBtn.addEventListener("click", closeDispatchMetaModal);
  }
  if (dispatchMetaCancelBtn) {
    dispatchMetaCancelBtn.addEventListener("click", handleDispatchLoadStageReset);
  }
  if (dispatchMetaConfirmBtn) {
    dispatchMetaConfirmBtn.addEventListener("click", handleDispatchMetaConfirmLoad);
  }
  if (dispatchMetaModal) {
    dispatchMetaModal.addEventListener("click", (event) => {
      if (event.target === dispatchMetaModal) {
        closeDispatchMetaModal();
      }
    });
  }
  if (dispatchSelectAllLoad) {
    dispatchSelectAllLoad.addEventListener("change", handleDispatchSelectAllChange);
  }
  if (financePrevBtn) {
    financePrevBtn.addEventListener("click", () => changeMainPage("finance", -1));
  }
  if (financeNextBtn) {
    financeNextBtn.addEventListener("click", () => changeMainPage("finance", 1));
  }
  if (statusPrevBtn) {
    statusPrevBtn.addEventListener("click", () => changeMainPage("status", -1));
  }
  if (statusNextBtn) {
    statusNextBtn.addEventListener("click", () => changeMainPage("status", 1));
  }
  if (statusActionRecordPrevBtn) {
    statusActionRecordPrevBtn.addEventListener("click", () => changeStatusActionRecordPage(-1));
  }
  if (statusActionRecordNextBtn) {
    statusActionRecordNextBtn.addEventListener("click", () => changeStatusActionRecordPage(1));
  }
  if (statusDetailModalCloseBtn) {
    statusDetailModalCloseBtn.addEventListener("click", closeStatusDetailModal);
  }
  if (statusDetailModal) {
    statusDetailModal.addEventListener("click", (event) => {
      if (event.target === statusDetailModal) {
        closeStatusDetailModal();
      }
    });
  }
  if (confirmModalCloseBtn) {
    confirmModalCloseBtn.addEventListener("click", () => resolveConfirmDialog(false));
  }
  if (confirmModalCancelBtn) {
    confirmModalCancelBtn.addEventListener("click", () => resolveConfirmDialog(false));
  }
  if (confirmModalConfirmBtn) {
    confirmModalConfirmBtn.addEventListener("click", () => resolveConfirmDialog(true));
  }
  if (confirmModal) {
    confirmModal.addEventListener("click", (event) => {
      if (event.target === confirmModal) {
        resolveConfirmDialog(false);
      }
    });
  }
  if (globalSearchInput) {
    globalSearchInput.addEventListener("input", handleGlobalSearchInput);
  }
  if (storageModeBadge) {
    storageModeBadge.addEventListener("click", handleStorageModeBadgeClick);
  }
  if (storageRequiredRetryBtn) {
    storageRequiredRetryBtn.addEventListener("click", handleStorageGuardRetryClick);
  }
  if (storageRequiredContinueBtn) {
    storageRequiredContinueBtn.addEventListener("click", handleStorageGuardContinueClick);
  }
  tableSortButtons.forEach((button) => {
    button.addEventListener("click", () => handleTableSortClick(button));
  });
  openSummaryViewButtons.forEach((button) => {
    button.addEventListener("click", () => openSummaryView(button));
  });
  closeSummaryViewButtons.forEach((button) => {
    button.addEventListener("click", () => closeSummaryView(button));
  });
  if (undoDeleteBtn) {
    undoDeleteBtn.addEventListener("click", undoPendingDeletion);
  }
  if (dispatchRecordModalCloseBtn) {
    dispatchRecordModalCloseBtn.addEventListener("click", closeDispatchRecordModal);
  }
  if (dispatchRecordExportBtn) {
    dispatchRecordExportBtn.addEventListener("click", handleDispatchRecordExportClick);
  }
  if (dispatchRecordEditBtn) {
    dispatchRecordEditBtn.addEventListener("click", handleDispatchRecordEditToggle);
  }
  if (dispatchRecordDeleteBtn) {
    dispatchRecordDeleteBtn.addEventListener("click", handleDispatchRecordDeleteClick);
  }
  if (dispatchRecordModal) {
    dispatchRecordModal.addEventListener("click", (event) => {
      if (event.target === dispatchRecordModal) {
        closeDispatchRecordModal();
      }
    });
  }

  shipmentTableBody.addEventListener("click", handleShipmentTableClick);
  shipmentTableBody.addEventListener("dblclick", handleShipmentTableDblClick);
  shipmentTableBody.addEventListener("change", handleShipmentExportSelectionChange);
  if (dispatchTableBody) {
    dispatchTableBody.addEventListener("change", handleDispatchLoadSelectionChange);
    dispatchTableBody.addEventListener("click", handleDispatchTableClick);
  }
  if (dispatchRecordBody) {
    dispatchRecordBody.addEventListener("click", handleDispatchRecordClick);
  }
  if (dispatchRecordDetailBody) {
    dispatchRecordDetailBody.addEventListener("click", handleDispatchRecordDetailClick);
  }
  if (dispatchRecordAddBody) {
    dispatchRecordAddBody.addEventListener("change", handleDispatchRecordAddSelectionChange);
  }
  if (dispatchRecordAddSelectAll) {
    dispatchRecordAddSelectAll.addEventListener("change", handleDispatchRecordAddSelectAllChange);
  }
  if (dispatchRecordAddBtn) {
    dispatchRecordAddBtn.addEventListener("click", handleDispatchRecordAppendClick);
  }
  if (dispatchRecordAddSearchInput) {
    dispatchRecordAddSearchInput.addEventListener("input", handleDispatchRecordAddSearchInput);
  }
  if (dispatchRecordAddClearBtn) {
    dispatchRecordAddClearBtn.addEventListener("click", handleDispatchRecordAddSearchClear);
  }
  if (statusTableBody) {
    statusTableBody.addEventListener("click", handleStatusTableClick);
  }
  financeTableBody.addEventListener("click", handleDeleteClick);
  financeTableBody.addEventListener("change", handleFinanceExportSelectionChange);
  if (shipmentExportFilteredSummaryBtn) {
    shipmentExportFilteredSummaryBtn.addEventListener("click", handleShipmentExportPreviewClick);
  }
  if (shipmentExportSelectAllCheckbox) {
    shipmentExportSelectAllCheckbox.addEventListener("change", handleShipmentExportSelectAllChange);
  }
  if (shipmentBatchImportBtn) {
    shipmentBatchImportBtn.addEventListener("click", handleShipmentBatchImportClick);
  }
  if (shipmentBatchImportFileInput) {
    shipmentBatchImportFileInput.addEventListener("change", handleShipmentBatchImportFileChange);
  }
  if (storageExportPackageBtn) {
    storageExportPackageBtn.addEventListener("click", handleStorageExportPackageClick);
  }
  if (storageImportPackageBtn) {
    storageImportPackageBtn.addEventListener("click", handleStorageImportPackageClick);
  }
  if (storageImportPackageFileInput) {
    storageImportPackageFileInput.addEventListener("change", handleStorageImportPackageFileChange);
  }
  if (storageImportPreviewCloseBtn) {
    storageImportPreviewCloseBtn.addEventListener("click", closeStorageImportPreviewModal);
  }
  if (storageImportPreviewCancelBtn) {
    storageImportPreviewCancelBtn.addEventListener("click", closeStorageImportPreviewModal);
  }
  if (storageImportPreviewConfirmBtn) {
    storageImportPreviewConfirmBtn.addEventListener("click", handleStorageImportPreviewConfirm);
  }
  if (storageImportPreviewModal) {
    storageImportPreviewModal.addEventListener("click", (event) => {
      if (event.target === storageImportPreviewModal) {
        closeStorageImportPreviewModal();
      }
    });
  }
  if (shipmentImportPreviewCloseBtn) {
    shipmentImportPreviewCloseBtn.addEventListener("click", closeShipmentImportPreviewModal);
  }
  if (shipmentImportPreviewCancelBtn) {
    shipmentImportPreviewCancelBtn.addEventListener("click", closeShipmentImportPreviewModal);
  }
  if (shipmentImportPreviewConfirmBtn) {
    shipmentImportPreviewConfirmBtn.addEventListener("click", handleShipmentImportPreviewConfirm);
  }
  if (shipmentImportPreviewModal) {
    shipmentImportPreviewModal.addEventListener("click", (event) => {
      if (event.target === shipmentImportPreviewModal) {
        closeShipmentImportPreviewModal();
      }
    });
  }
  if (shipmentExportPreviewCloseBtn) {
    shipmentExportPreviewCloseBtn.addEventListener("click", closeShipmentExportPreviewModal);
  }
  if (shipmentExportPreviewCancelBtn) {
    shipmentExportPreviewCancelBtn.addEventListener("click", closeShipmentExportPreviewModal);
  }
  if (shipmentExportPreviewConfirmBtn) {
    shipmentExportPreviewConfirmBtn.addEventListener("click", handleShipmentExportPreviewConfirm);
  }
  if (shipmentExportPreviewModal) {
    shipmentExportPreviewModal.addEventListener("click", (event) => {
      if (event.target === shipmentExportPreviewModal) {
        closeShipmentExportPreviewModal();
      }
    });
  }
  if (financeExportPreviewCloseBtn) {
    financeExportPreviewCloseBtn.addEventListener("click", closeFinanceExportPreviewModal);
  }
  if (financeExportPreviewCancelBtn) {
    financeExportPreviewCancelBtn.addEventListener("click", closeFinanceExportPreviewModal);
  }
  if (financeExportPreviewConfirmBtn) {
    financeExportPreviewConfirmBtn.addEventListener("click", handleFinanceExportPreviewConfirm);
  }
  if (financeExportPreviewModal) {
    financeExportPreviewModal.addEventListener("click", (event) => {
      if (event.target === financeExportPreviewModal) {
        closeFinanceExportPreviewModal();
      }
    });
  }

  if (summaryForm) {
    summaryForm.addEventListener("submit", handleSummarySubmit);
  }
  if (summaryFromDateInput) {
    summaryFromDateInput.addEventListener("change", () => {
      normalizeDateInputs(summaryFromDateInput, summaryToDateInput);
      clearSummaryQuickRangeActive();
      renderDailySummarySection({ resetPages: true });
    });
  }
  if (summaryToDateInput) {
    summaryToDateInput.addEventListener("change", () => {
      normalizeDateInputs(summaryFromDateInput, summaryToDateInput);
      clearSummaryQuickRangeActive();
      renderDailySummarySection({ resetPages: true });
    });
  }
  if (summaryResetRangeBtn) {
    summaryResetRangeBtn.addEventListener("click", resetSummaryRangeToCurrentMonth);
  }
  if (summaryShipmentSearchInput) {
    summaryShipmentSearchInput.addEventListener("input", () => {
      summaryShipmentKeyword = cleanText(summaryShipmentSearchInput.value);
      if (activePanelId === "daily-summary-panel" && activeSummaryView === "shipment" && globalSearchInput) {
        globalSearchInput.value = summaryShipmentKeyword;
      }
      renderDailySummarySection({ resetPages: true });
    });
  }
  if (summaryShipmentClearSearchBtn) {
    summaryShipmentClearSearchBtn.addEventListener("click", () => {
      summaryShipmentKeyword = "";
      syncSummaryKeywordInputs();
      if (activePanelId === "daily-summary-panel" && activeSummaryView === "shipment" && globalSearchInput) {
        globalSearchInput.value = "";
      }
      renderDailySummarySection({ resetPages: true });
    });
  }
  if (summaryFinanceSearchInput) {
    summaryFinanceSearchInput.addEventListener("input", () => {
      summaryFinanceKeyword = cleanText(summaryFinanceSearchInput.value);
      if (activePanelId === "daily-summary-panel" && activeSummaryView === "finance" && globalSearchInput) {
        globalSearchInput.value = summaryFinanceKeyword;
      }
      renderDailySummarySection({ resetPages: true });
    });
  }
  if (summaryFinanceClearSearchBtn) {
    summaryFinanceClearSearchBtn.addEventListener("click", () => {
      summaryFinanceKeyword = "";
      syncSummaryKeywordInputs();
      if (activePanelId === "daily-summary-panel" && activeSummaryView === "finance" && globalSearchInput) {
        globalSearchInput.value = "";
      }
      renderDailySummarySection({ resetPages: true });
    });
  }
  if (summaryQuickRangeButtons.length) {
    summaryQuickRangeButtons.forEach((button) => {
      button.addEventListener("click", () => {
        applySummaryQuickRange(button.dataset.summaryQuickRange);
      });
    });
  }
  if (summaryShipmentPrevBtn) {
    summaryShipmentPrevBtn.addEventListener("click", () => changeSummaryPage("shipment", -1));
  }
  if (summaryShipmentNextBtn) {
    summaryShipmentNextBtn.addEventListener("click", () => changeSummaryPage("shipment", 1));
  }
  if (summaryFinancePrevBtn) {
    summaryFinancePrevBtn.addEventListener("click", () => changeSummaryPage("finance", -1));
  }
  if (summaryFinanceNextBtn) {
    summaryFinanceNextBtn.addEventListener("click", () => changeSummaryPage("finance", 1));
  }
}

function setupStorageSyncLifecycle() {
  if (storageLifecycleBound || typeof window === "undefined") {
    return;
  }
  storageLifecycleBound = true;
  window.addEventListener("beforeunload", (event) => {
    if (!manualSaveDirty) {
      return;
    }
    event.preventDefault();
    event.returnValue = "";
  });
}

function setupTableMouseInteractions() {
  const tableWraps = Array.from(document.querySelectorAll(".table-wrap"));
  if (!tableWraps.length) {
    return;
  }

  tableWraps.forEach((wrap) => {
    if (!(wrap instanceof HTMLElement) || wrap.dataset.mouseReady === "1") {
      return;
    }
    wrap.dataset.mouseReady = "1";

    let dragging = false;
    let startX = 0;
    let startScrollLeft = 0;
    const hasHorizontalOverflow = () => wrap.scrollWidth - wrap.clientWidth > 2;
    const updateScrollableState = () => {
      wrap.classList.toggle("is-x-scrollable", hasHorizontalOverflow());
    };

    const stopDrag = () => {
      if (!dragging) {
        return;
      }
      dragging = false;
      wrap.classList.remove("is-dragging");
      document.body.classList.remove("is-table-dragging");
    };

    updateScrollableState();
    window.addEventListener("resize", updateScrollableState, { passive: true });

    if (typeof ResizeObserver !== "undefined") {
      const resizeObserver = new ResizeObserver(updateScrollableState);
      resizeObserver.observe(wrap);
      const table = wrap.querySelector("table");
      if (table) {
        resizeObserver.observe(table);
      }
    }

    wrap.addEventListener(
      "wheel",
      (event) => {
        if (!hasHorizontalOverflow()) {
          return;
        }

        if (event.shiftKey && Math.abs(event.deltaY) > 0) {
          wrap.scrollLeft += event.deltaY;
          event.preventDefault();
          return;
        }

        if (Math.abs(event.deltaX) > 0) {
          wrap.scrollLeft += event.deltaX;
          event.preventDefault();
        }
      },
      { passive: false }
    );

    wrap.addEventListener("mousedown", (event) => {
      if (event.button !== 0 || !hasHorizontalOverflow() || shouldSkipDragScroll(event.target)) {
        return;
      }
      dragging = true;
      startX = event.clientX;
      startScrollLeft = wrap.scrollLeft;
      wrap.classList.add("is-dragging");
      document.body.classList.add("is-table-dragging");
      event.preventDefault();
    });

    document.addEventListener("mousemove", (event) => {
      if (!dragging) {
        return;
      }
      const deltaX = event.clientX - startX;
      wrap.scrollLeft = startScrollLeft - deltaX;
    });

    document.addEventListener("mouseup", stopDrag);
    document.addEventListener("mouseleave", stopDrag);
    wrap.addEventListener("dragstart", stopDrag);
    wrap.addEventListener("blur", stopDrag);
  });
}

function shouldSkipDragScroll(target) {
  if (!(target instanceof Element)) {
    return false;
  }
  return Boolean(
    target.closest("input, button, select, textarea, a, label, [contenteditable='true'], .th-sort")
  );
}

function setupKeyboardShortcuts() {
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") {
      if (confirmModal?.classList.contains("is-open")) {
        resolveConfirmDialog(false);
        return;
      }
      if (financeExportPreviewModal?.classList.contains("is-open")) {
        closeFinanceExportPreviewModal();
        return;
      }
      if (shipmentImportPreviewModal?.classList.contains("is-open")) {
        closeShipmentImportPreviewModal();
        return;
      }
      if (storageImportPreviewModal?.classList.contains("is-open")) {
        closeStorageImportPreviewModal();
        return;
      }
      if (shipmentExportPreviewModal?.classList.contains("is-open")) {
        closeShipmentExportPreviewModal();
        return;
      }
      if (dispatchMetaModal?.classList.contains("is-open")) {
        closeDispatchMetaModal();
        return;
      }
      if (dispatchRecordModal?.classList.contains("is-open")) {
        closeDispatchRecordModal();
        return;
      }
    }

    if (!event.altKey && event.key === "/" && !isInputLike(event.target)) {
      event.preventDefault();
      globalSearchInput?.focus();
      globalSearchInput?.select();
      return;
    }

    if (event.altKey && !event.shiftKey && !event.metaKey && !event.ctrlKey) {
      const panelKeyMap = {
        "1": "shipment-panel",
        "2": "dispatch-panel",
        "3": "finance-panel",
        "4": "daily-summary-panel",
        "5": "status-panel",
      };
      const targetPanelId = panelKeyMap[event.key];
      if (targetPanelId) {
        event.preventDefault();
        activatePanel(targetPanelId, { scroll: true });
      }
    }

    if (event.key === "Escape" && globalSearchInput && document.activeElement === globalSearchInput) {
      globalSearchInput.value = "";
      handleGlobalSearchInput();
      globalSearchInput.blur();
    }
  });
}

function setupSummaryViewControls() {
  if (!summaryViewButtons.length) {
    return;
  }

  summaryViewButtons.forEach((button) => {
    button.addEventListener("click", () => {
      activateSummaryView(button.dataset.summaryView);
    });
  });

  const initialView =
    summaryViewButtons.find((button) => button.classList.contains("is-active"))?.dataset.summaryView ||
    summaryViewButtons[0].dataset.summaryView;

  activateSummaryView(initialView);
}

function setupPanelControls() {
  sidebarOpenButtons.forEach((button) => {
    button.addEventListener("click", () => {
      activatePanel(button.dataset.openPanel, { scroll: true });
    });
  });

  const initialActiveButton = sidebarOpenButtons.find((button) =>
    button.classList.contains("is-active")
  );
  const initialPanelId = (initialActiveButton || sidebarOpenButtons[0])?.dataset.openPanel;
  if (initialPanelId) {
    activatePanel(initialPanelId, { scroll: false });
  }
}

function setupSummaryLayoutViews() {
  document.querySelectorAll("[data-summary-layout]").forEach((grid) => {
    setSummaryLayoutMode(grid, "function");
  });
}

function openSummaryView(button) {
  const grid =
    button.closest("[data-summary-layout]") ||
    button.closest(".panel")?.querySelector("[data-summary-layout]");
  if (!grid) {
    return;
  }
  setSummaryLayoutMode(grid, "summary");
}

function closeSummaryView(button) {
  const grid =
    button.closest("[data-summary-layout]") ||
    button.closest(".panel")?.querySelector("[data-summary-layout]");
  if (!grid) {
    return;
  }
  setSummaryLayoutMode(grid, "function");
}

function setSummaryLayoutMode(grid, mode) {
  const summaryMode = mode === "summary";
  grid.classList.toggle("is-summary-view", summaryMode);
  grid.classList.toggle("is-function-view", !summaryMode);
  const panel = grid.closest(".panel");
  if (panel) {
    panel.classList.toggle("is-summary-detail-view", summaryMode);
  }

  const enterButtons = panel
    ? Array.from(panel.querySelectorAll("[data-open-summary-view]"))
    : Array.from(grid.querySelectorAll("[data-open-summary-view]"));
  enterButtons.forEach((button) => {
    button.setAttribute("aria-expanded", String(summaryMode));
  });
}

function activatePanel(panelId, options = {}) {
  const targetPanel = document.getElementById(panelId || "");
  if (!targetPanel) {
    return;
  }

  if (panelId !== "dispatch-panel" && dispatchRecordModal?.classList.contains("is-open")) {
    closeDispatchRecordModal();
  }

  panels.forEach((panel) => {
    panel.classList.toggle("is-hidden", panel !== targetPanel);
  });

  activePanelId = panelId;
  updateSidebarActive(panelId);
  syncGlobalSearchFromPanel(panelId);

  if (options.scroll !== false) {
    targetPanel.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  if (panelId === "daily-summary-panel") {
    renderDailySummarySection();
    return;
  }

  if (panelId === "status-panel") {
    renderStatusSection();
  }
}

function updateSidebarActive(panelId) {
  sidebarOpenButtons.forEach((button) => {
    const active = button.dataset.openPanel === panelId;
    button.classList.toggle("is-active", active);
    if (active) {
      button.setAttribute("aria-current", "true");
    } else {
      button.removeAttribute("aria-current");
    }
  });
}

function activateSummaryView(view) {
  const normalized = view === "finance" ? "finance" : "shipment";
  activeSummaryView = normalized;
  syncSummaryKeywordInputs();

  summaryViewButtons.forEach((button) => {
    const active = button.dataset.summaryView === normalized;
    button.classList.toggle("is-active", active);
    button.setAttribute("aria-pressed", String(active));
  });

  summaryGroups.forEach((group) => {
    group.classList.toggle("is-hidden-sub", group.dataset.summaryGroup !== normalized);
  });

  if (activePanelId === "daily-summary-panel") {
    syncGlobalSearchFromPanel(activePanelId);
  }
}

function syncGlobalSearchFromPanel(panelId) {
  if (!globalSearchInput) {
    return;
  }

  if (panelId === "shipment-panel") {
    globalSearchInput.placeholder = "搜索单号、站点、单位、货名...";
    globalSearchInput.value = shipmentSearch.value;
    return;
  }

  if (panelId === "dispatch-panel") {
    globalSearchInput.placeholder = "搜索配车货物：序号、厂家、客户、产品...";
    globalSearchInput.value = dispatchSearch?.value || "";
    return;
  }

  if (panelId === "finance-panel") {
    globalSearchInput.placeholder = "搜索摘要、备注...";
    globalSearchInput.value = financeSearch.value;
    return;
  }

  if (panelId === "status-panel") {
    globalSearchInput.placeholder = "搜索家结提醒：单号、厂家、客户、产品、备注...";
    globalSearchInput.value = statusSearch?.value || "";
    return;
  }

  if (panelId === "daily-summary-panel") {
    if (activeSummaryView === "finance") {
      globalSearchInput.placeholder = "搜索日常收支：摘要、备注...";
      globalSearchInput.value = summaryFinanceKeyword;
      return;
    }
    globalSearchInput.placeholder = "搜索收发汇总：单号、单位、备注...";
    globalSearchInput.value = summaryShipmentKeyword;
    return;
  }

  globalSearchInput.placeholder = "输入关键词搜索...";
  globalSearchInput.value = "";
}

function handleGlobalSearchInput() {
  const keyword = cleanText(globalSearchInput?.value);

  if (activePanelId === "shipment-panel") {
    shipmentSearch.value = keyword;
    renderShipmentTable({ resetPage: true });
    return;
  }

  if (activePanelId === "dispatch-panel") {
    if (dispatchSearch) {
      dispatchSearch.value = keyword;
    }
    renderDispatchTable({ resetPage: true });
    return;
  }

  if (activePanelId === "finance-panel") {
    financeSearch.value = keyword;
    renderFinanceTable({ resetPage: true });
    return;
  }

  if (activePanelId === "status-panel") {
    if (statusSearch) {
      statusSearch.value = keyword;
    }
    renderStatusSection({ resetPage: true });
    return;
  }

  if (activePanelId === "daily-summary-panel") {
    if (activeSummaryView === "finance") {
      summaryFinanceKeyword = keyword;
    } else {
      summaryShipmentKeyword = keyword;
    }
    syncSummaryKeywordInputs();
    renderDailySummarySection({ resetPages: true });
  }
}

function handleTableSortClick(button) {
  const table = button.dataset.tableSort;
  const key = button.dataset.sortKey;
  if (!table || !key) {
    return;
  }

  if (table === "shipment") {
    shipmentSort = buildNextSort(shipmentSort, key);
    updateSortButtonsUI("shipment");
    renderShipmentTable({ resetPage: true });
    return;
  }

  if (table === "finance") {
    financeSort = buildNextSort(financeSort, key);
    updateSortButtonsUI("finance");
    renderFinanceTable({ resetPage: true });
  }
}

function buildNextSort(current, nextKey) {
  if (current.key !== nextKey) {
    return { key: nextKey, dir: "desc" };
  }

  return { key: nextKey, dir: current.dir === "desc" ? "asc" : "desc" };
}

function updateSortButtonsUI(table) {
  const state = table === "finance" ? financeSort : shipmentSort;
  tableSortButtons.forEach((button) => {
    if (button.dataset.tableSort !== table) {
      return;
    }

    const active = button.dataset.sortKey === state.key;
    button.classList.toggle("is-active", active);
    if (active) {
      button.dataset.sortDir = state.dir;
      return;
    }
    delete button.dataset.sortDir;
  });
}

function getShipmentUnitWeight(item) {
  return Math.max(0, toFiniteNumber(item?.weight));
}

function getShipmentTotalWeight(item) {
  const unitWeight = getShipmentUnitWeight(item);
  const quantity = Math.max(0, toFiniteNumber(item?.quantity));
  return unitWeight * quantity;
}

function getShipmentUnitCubic(item) {
  const cubic = toFiniteNumber(item?.cubic, NaN);
  if (Number.isFinite(cubic) && cubic >= 0) {
    return cubic;
  }

  const unit = normalizeMeasureUnit(item?.measureUnit, DEFAULT_MEASURE_UNIT) || DEFAULT_MEASURE_UNIT;
  return unit === "cubic" ? 1 : 0;
}

function getShipmentTotalCubic(item) {
  const unitCubic = getShipmentUnitCubic(item);
  const quantity = Math.max(0, toFiniteNumber(item?.quantity));
  return unitCubic * quantity;
}

function sortShipmentRows(rows) {
  const sorted = rows.slice();
  const direction = shipmentSort.dir === "asc" ? 1 : -1;

  sorted.sort((a, b) => {
    if (shipmentSort.key === "quantity") {
      const diff = toFiniteNumber(a.quantity) - toFiniteNumber(b.quantity);
      if (diff !== 0) {
        return diff * direction;
      }
      return sortByDateDesc(a, b);
    }

    if (shipmentSort.key === "unitWeight") {
      const diff = getShipmentUnitWeight(a) - getShipmentUnitWeight(b);
      if (diff !== 0) {
        return diff * direction;
      }
      return sortByDateDesc(a, b);
    }

    if (shipmentSort.key === "totalWeight" || shipmentSort.key === "weight") {
      const diff = getShipmentTotalWeight(a) - getShipmentTotalWeight(b);
      if (diff !== 0) {
        return diff * direction;
      }
      return sortByDateDesc(a, b);
    }

    if (shipmentSort.key === "amount") {
      const diff = toFiniteNumber(a.amount) - toFiniteNumber(b.amount);
      if (diff !== 0) {
        return diff * direction;
      }
      return sortByDateDesc(a, b);
    }

    const textCompare = String(a.date || "").localeCompare(String(b.date || ""));
    if (textCompare !== 0) {
      return textCompare * direction;
    }
    return (a.createdAt - b.createdAt) * direction;
  });

  return sorted;
}

function sortFinanceRows(rows) {
  const sorted = rows.slice();
  const direction = financeSort.dir === "asc" ? 1 : -1;

  sorted.sort((a, b) => {
    if (financeSort.key === "amount") {
      const diff = toFiniteNumber(a.amount) - toFiniteNumber(b.amount);
      if (diff !== 0) {
        return diff * direction;
      }
      return sortByDateDesc(a, b);
    }

    const textCompare = String(a.date || "").localeCompare(String(b.date || ""));
    if (textCompare !== 0) {
      return textCompare * direction;
    }
    return (a.createdAt - b.createdAt) * direction;
  });

  return sorted;
}

function setDefaultDates() {
  const today = dateValue(new Date());
  const thisMonth = today.slice(0, 7);
  const monthStart = `${thisMonth}-01`;

  shipmentForm.elements.date.value = today;
  if (shipmentForm.elements.type) {
    shipmentForm.elements.type.value = "receive";
  }
  if (shipmentFieldEls.measureUnit) {
    shipmentFieldEls.measureUnit.value = DEFAULT_MEASURE_UNIT;
  }
  configureShipmentQuantityInputByUnit();
  financeForm.elements.date.value = today;
  financeMonth.value = "";
  if (summaryFromDateInput) {
    summaryFromDateInput.value = monthStart;
  }
  if (summaryToDateInput) {
    summaryToDateInput.value = today;
  }
  setSummaryQuickRangeActive("thisMonth");
  if (dispatchMetaDateInput) {
    dispatchMetaDateInput.value = today;
  }
}

function restoreRetainedFormValues() {
  applyFormValues(shipmentForm, readObjectStorage(STORAGE_KEYS.shipmentRetain, {}), SHIPMENT_RETAIN_FIELDS);
  applyFormValues(financeForm, readObjectStorage(STORAGE_KEYS.financeRetain, {}), FINANCE_RETAIN_FIELDS);

  const today = dateValue(new Date());
  if (shipmentForm?.elements?.date) {
    shipmentForm.elements.date.value = today;
  }
  if (shipmentFieldEls.measureUnit) {
    shipmentFieldEls.measureUnit.value =
      normalizeMeasureUnit(shipmentFieldEls.measureUnit.value, DEFAULT_MEASURE_UNIT) || DEFAULT_MEASURE_UNIT;
  }
  configureShipmentQuantityInputByUnit();
  syncShipmentNoteSelectionFromField();
  if (financeForm?.elements?.date) {
    financeForm.elements.date.value = today;
  }
}

function bindFormRetention(form, storageKey, allowedFields) {
  if (!form) {
    return;
  }

  const persist = () => persistRetainedFromForm(form, storageKey, allowedFields);
  form.addEventListener("input", persist);
  form.addEventListener("change", persist);
}

function persistRetainedFromForm(form, storageKey, allowedFields) {
  if (!form) {
    return;
  }

  const retained = {};
  allowedFields.forEach((name) => {
    const field = form.elements[name];
    if (!field) {
      return;
    }
    retained[name] = cleanText(field.value);
  });
  writeObjectStorage(storageKey, retained);
}

function pickRetainedValuesFromFormData(fd, allowedFields) {
  const retained = {};
  allowedFields.forEach((name) => {
    retained[name] = cleanText(fd.get(name));
  });
  return retained;
}

function applyFormValues(form, values, allowedFields) {
  if (!form || !values || typeof values !== "object") {
    return;
  }

  const fields = Array.isArray(allowedFields) ? allowedFields : Object.keys(values);
  fields.forEach((name) => {
    if (!(name in values)) {
      return;
    }
    const field = form.elements[name];
    if (!field) {
      return;
    }
    field.value = String(values[name] ?? "");
  });
}

function setupShipmentSmartFill() {
  if (!shipmentForm || !shipmentFieldEls.product) {
    return;
  }

  const refreshSuggestions = (showPanel = false) => {
    tryAutoFillShipmentFields();
    refreshShipmentSmartFillSuggestions({ showPanel });
  };

  shipmentFieldEls.product.addEventListener("focus", () => {
    clearProductSuggestionHideTimer();
    refreshSuggestions(true);
  });
  shipmentFieldEls.product.addEventListener("click", () => {
    clearProductSuggestionHideTimer();
    refreshSuggestions(true);
  });
  shipmentFieldEls.product.addEventListener("input", () => refreshSuggestions(true));
  shipmentFieldEls.product.addEventListener("change", () => refreshSuggestions(true));
  shipmentFieldEls.product.addEventListener("blur", scheduleHideProductSuggestionPanel);

  if (shipmentFieldEls.manufacturer) {
    shipmentFieldEls.manufacturer.addEventListener("focus", () => refreshShipmentSmartFillSuggestions());
    shipmentFieldEls.manufacturer.addEventListener("input", () => refreshSuggestions(isProductFieldFocused()));
    shipmentFieldEls.manufacturer.addEventListener("change", () => refreshSuggestions(isProductFieldFocused()));
  }

  if (shipmentFieldEls.customer) {
    shipmentFieldEls.customer.addEventListener("focus", () => refreshShipmentSmartFillSuggestions());
    shipmentFieldEls.customer.addEventListener("input", () => refreshSuggestions(isProductFieldFocused()));
    shipmentFieldEls.customer.addEventListener("change", () => refreshSuggestions(isProductFieldFocused()));
  }

  if (shipmentFieldEls.quantity) {
    shipmentFieldEls.quantity.addEventListener("input", syncShipmentAmountFromUnitPrice);
  }
  if (shipmentFieldEls.measureUnit) {
    shipmentFieldEls.measureUnit.addEventListener("change", () => {
      configureShipmentQuantityInputByUnit();
      syncShipmentAmountFromUnitPrice({ force: true });
      refreshShipmentSmartFillSuggestions({ showPanel: isProductFieldFocused() });
      persistRetainedFromForm(shipmentForm, STORAGE_KEYS.shipmentRetain, SHIPMENT_RETAIN_FIELDS);
    });
  }
  if (shipmentFieldEls.unitPrice) {
    const markUnitPriceManual = () => {
      shipmentFieldEls.unitPrice.dataset.autofill = "0";
    };
    shipmentFieldEls.unitPrice.addEventListener("input", () => {
      markUnitPriceManual();
      syncShipmentAmountFromUnitPrice();
    });
    shipmentFieldEls.unitPrice.addEventListener("change", () => {
      markUnitPriceManual();
      syncShipmentAmountFromUnitPrice();
    });
  }
  if (shipmentFieldEls.amount) {
    const markManualAmount = () => {
      shipmentFieldEls.amount.dataset.manual = cleanText(shipmentFieldEls.amount.value) ? "1" : "0";
    };
    shipmentFieldEls.amount.addEventListener("input", markManualAmount);
    shipmentFieldEls.amount.addEventListener("change", markManualAmount);
    shipmentFieldEls.amount.dataset.manual = "0";
  }

  if (shipmentProductSuggestionPanel) {
    shipmentProductSuggestionPanel.addEventListener("mousedown", (event) => {
      // Keep focus on product input so clicking suggestion does not collapse first.
      event.preventDefault();
    });
    shipmentProductSuggestionPanel.addEventListener("click", handleProductSuggestionPanelClick);
  }

  refreshShipmentSmartFillSuggestions();
  syncShipmentAmountFromUnitPrice({ force: true });
}

function rebuildShipmentProfileIndex() {
  const merged = new Map();

  BUILTIN_PRICE_CATALOG_ROWS.forEach((record, index) => {
    upsertShipmentProfile(merged, {
      manufacturer: record.manufacturer,
      customer: record.customer,
      product: record.product,
      measureUnit: normalizeMeasureUnit(record.measureUnit, ""),
      unitPrice: record.unitPrice,
      rebate: record.rebate,
      weight: record.weight,
      cubic: record.cubic,
      transit: record.transit,
      delivery: record.delivery,
      pickup: "",
      note: record.note,
      sourcePriority: 1,
      updatedAt: index + 1,
      historyHits: 0,
    });
  });

  customCatalogProfiles.forEach((item, index) => {
    upsertShipmentProfile(merged, {
      manufacturer: cleanText(item?.manufacturer),
      customer: cleanText(item?.customer),
      product: cleanText(item?.product),
      measureUnit: normalizeMeasureUnit(item?.measureUnit, ""),
      unitPrice: toFiniteNumber(item?.unitPrice, NaN),
      rebate: toFiniteNumber(item?.rebate, NaN),
      weight: toFiniteNumber(item?.weight, NaN),
      cubic: toFiniteNumber(item?.cubic, NaN),
      transit: cleanText(item?.transit),
      delivery: cleanText(item?.delivery),
      pickup: cleanText(item?.pickup),
      note: cleanText(item?.note),
      sourcePriority: 2,
      updatedAt: Number(item?.updatedAt) || index + 1,
      historyHits: Math.max(0, Number(item?.historyHits) || 0),
    });
  });

  shipments.forEach((item, index) => {
    if (!cleanText(item.product)) {
      return;
    }

    upsertShipmentProfile(merged, {
      manufacturer: cleanText(item.manufacturer),
      customer: cleanText(item.customer),
      product: cleanText(item.product),
      measureUnit: normalizeMeasureUnit(item.measureUnit, ""),
      unitPrice: toFiniteNumber(item.unitPrice, NaN),
      rebate: toFiniteNumber(item.rebate, NaN),
      weight: toFiniteNumber(item.weight, NaN),
      cubic: toFiniteNumber(item.cubic, NaN),
      transit: cleanText(item.transit),
      delivery: cleanText(item.delivery),
      pickup: cleanText(item.pickup),
      note: cleanText(item.note),
      sourcePriority: 3,
      updatedAt: Number(item.createdAt) || Date.now() + index,
      historyHits: 1,
    });
  });

  shipmentProfileDataset = Array.from(merged.values())
    .map((item) => item.profile)
    .sort((a, b) => {
      if (a.sourcePriority !== b.sourcePriority) {
        return b.sourcePriority - a.sourcePriority;
      }
      if (a.historyHits !== b.historyHits) {
        return b.historyHits - a.historyHits;
      }
      return (b.updatedAt || 0) - (a.updatedAt || 0);
    });

  shipmentProfileByExactKey = new Map();
  const byProduct = new Map();
  const manufacturerSet = new Set();
  const customerSet = new Set();

  shipmentProfileDataset.forEach((profile) => {
    const exactKey = makeShipmentProfileKey(profile.manufacturer, profile.customer, profile.product);
    if (exactKey && !shipmentProfileByExactKey.has(exactKey)) {
      shipmentProfileByExactKey.set(exactKey, profile);
    }

    const productKey = normalizeSuggestText(profile.product);
    if (productKey) {
      if (!byProduct.has(productKey)) {
        byProduct.set(productKey, []);
      }
      byProduct.get(productKey).push(profile);
    }

    if (profile.manufacturer) {
      manufacturerSet.add(profile.manufacturer);
    }
    if (profile.customer) {
      customerSet.add(profile.customer);
    }
  });

  shipmentProfilesByProductKey = byProduct;
  shipmentProductSuggestionPool = buildProductSuggestionPool(byProduct);
  manufacturerSuggestionPool = Array.from(manufacturerSet).sort((a, b) => a.localeCompare(b, "zh-CN"));
  customerSuggestionPool = Array.from(customerSet).sort((a, b) => a.localeCompare(b, "zh-CN"));
  refreshShipmentSmartFillSuggestions();
}

function upsertShipmentProfile(store, rawProfile) {
  const profile = normalizeShipmentProfile(rawProfile);
  const key = makeShipmentProfileKey(profile.manufacturer, profile.customer, profile.product);
  if (!key) {
    return;
  }

  const existing = store.get(key);
  if (!existing) {
    store.set(key, { profile, rank: getShipmentProfileRank(profile) });
    return;
  }

  const existingProfile = existing.profile;
  const incomingRank = getShipmentProfileRank(profile);
  const preferred = incomingRank >= existing.rank ? profile : existingProfile;
  const fallback = incomingRank >= existing.rank ? existingProfile : profile;
  const merged = mergeShipmentProfiles(preferred, fallback);
  store.set(key, { profile: merged, rank: Math.max(incomingRank, existing.rank) });
}

function sanitizeCustomCatalogProfiles(list) {
  if (!Array.isArray(list) || list.length === 0) {
    return [];
  }

  const merged = new Map();
  list.forEach((item, index) => {
    upsertShipmentProfile(merged, {
      manufacturer: cleanText(item?.manufacturer),
      customer: cleanText(item?.customer),
      product: cleanText(item?.product),
      measureUnit: normalizeMeasureUnit(item?.measureUnit, ""),
      unitPrice: toFiniteNumber(item?.unitPrice, NaN),
      rebate: toFiniteNumber(item?.rebate, NaN),
      weight: toFiniteNumber(item?.weight, NaN),
      cubic: toFiniteNumber(item?.cubic, NaN),
      transit: cleanText(item?.transit),
      delivery: cleanText(item?.delivery),
      pickup: cleanText(item?.pickup),
      note: cleanText(item?.note),
      sourcePriority: 2,
      updatedAt: Number(item?.updatedAt) || index + 1,
      historyHits: Math.max(0, Number(item?.historyHits) || 0),
    });
  });

  return Array.from(merged.values())
    .map((entry) => serializeCatalogProfile(entry.profile))
    .sort((a, b) => (b.updatedAt || 0) - (a.updatedAt || 0));
}

function serializeCatalogProfile(profile) {
  const normalized = normalizeShipmentProfile(profile || {});
  const stored = {
    manufacturer: normalized.manufacturer,
    customer: normalized.customer,
    product: normalized.product,
    measureUnit: normalizeMeasureUnit(normalized.measureUnit, ""),
    transit: normalized.transit,
    delivery: normalized.delivery,
    pickup: normalized.pickup,
    note: normalized.note,
    updatedAt: Number(normalized.updatedAt) || 0,
    historyHits: Math.max(0, Number(normalized.historyHits) || 0),
  };

  if (Number.isFinite(normalized.unitPrice)) {
    stored.unitPrice = normalized.unitPrice;
  }
  if (Number.isFinite(normalized.rebate)) {
    stored.rebate = normalized.rebate;
  }
  if (Number.isFinite(normalized.weight)) {
    stored.weight = normalized.weight;
  }
  if (Number.isFinite(normalized.cubic)) {
    stored.cubic = normalized.cubic;
  }

  return stored;
}

function upsertCustomCatalogProfile(rawProfile) {
  const merged = new Map();
  sanitizeCustomCatalogProfiles(customCatalogProfiles).forEach((item, index) => {
    upsertShipmentProfile(merged, {
      ...item,
      sourcePriority: 2,
      updatedAt: Number(item.updatedAt) || index + 1,
      historyHits: Math.max(0, Number(item.historyHits) || 0),
    });
  });

  upsertShipmentProfile(merged, {
    manufacturer: cleanText(rawProfile?.manufacturer),
    customer: cleanText(rawProfile?.customer),
    product: cleanText(rawProfile?.product),
    measureUnit: normalizeMeasureUnit(rawProfile?.measureUnit, ""),
    unitPrice: toFiniteNumber(rawProfile?.unitPrice, NaN),
    rebate: toFiniteNumber(rawProfile?.rebate, NaN),
    weight: toFiniteNumber(rawProfile?.weight, NaN),
    cubic: toFiniteNumber(rawProfile?.cubic, NaN),
    transit: cleanText(rawProfile?.transit),
    delivery: cleanText(rawProfile?.delivery),
    pickup: cleanText(rawProfile?.pickup),
    note: cleanText(rawProfile?.note),
    sourcePriority: 2,
    updatedAt: Date.now(),
    historyHits: 1,
  });

  customCatalogProfiles = Array.from(merged.values())
    .map((entry) => serializeCatalogProfile(entry.profile))
    .sort((a, b) => (b.updatedAt || 0) - (a.updatedAt || 0));
  writeStorage(STORAGE_KEYS.catalogProfiles, customCatalogProfiles);
}

function normalizeShipmentProfile(profile) {
  return {
    manufacturer: cleanText(profile.manufacturer),
    customer: cleanText(profile.customer),
    product: cleanText(profile.product),
    measureUnit: normalizeMeasureUnit(profile.measureUnit, ""),
    unitPrice: toFiniteNumber(profile.unitPrice, NaN),
    rebate: toFiniteNumber(profile.rebate, NaN),
    weight: toFiniteNumber(profile.weight, NaN),
    cubic: toFiniteNumber(profile.cubic, NaN),
    transit: cleanText(profile.transit),
    delivery: cleanText(profile.delivery),
    pickup: cleanText(profile.pickup),
    note: cleanText(profile.note),
    sourcePriority: Number(profile.sourcePriority) || 0,
    updatedAt: Number(profile.updatedAt) || 0,
    historyHits: Number(profile.historyHits) || 0,
  };
}

function getShipmentProfileRank(profile) {
  const priority = (Number(profile.sourcePriority) || 0) * 1_000_000_000_000;
  const updatedAt = Number(profile.updatedAt) || 0;
  const completeness = getShipmentProfileCompleteness(profile);
  return priority + updatedAt + completeness;
}

function getShipmentProfileCompleteness(profile) {
  let score = 0;
  if (profile.manufacturer) {
    score += 1;
  }
  if (profile.customer) {
    score += 1;
  }
  if (profile.product) {
    score += 1;
  }
  if (profile.measureUnit) {
    score += 1;
  }
  if (Number.isFinite(profile.unitPrice)) {
    score += 2;
  }
  if (Number.isFinite(profile.rebate)) {
    score += 1;
  }
  if (Number.isFinite(profile.weight)) {
    score += 1;
  }
  if (Number.isFinite(profile.cubic)) {
    score += 1;
  }
  if (profile.transit) {
    score += 1;
  }
  if (profile.delivery) {
    score += 1;
  }
  if (profile.pickup) {
    score += 1;
  }
  if (profile.note) {
    score += 1;
  }
  return score;
}

function mergeShipmentProfiles(primary, secondary) {
  const merged = { ...primary };
  const textFields = ["manufacturer", "customer", "product", "transit", "delivery", "pickup", "note"];
  const numberFields = ["unitPrice", "rebate", "weight", "cubic"];
  const primaryUnit = normalizeMeasureUnit(merged.measureUnit, "");
  const secondaryUnit = normalizeMeasureUnit(secondary.measureUnit, "");

  if (!primaryUnit && secondaryUnit) {
    merged.measureUnit = secondaryUnit;
  }

  textFields.forEach((field) => {
    if (!cleanText(merged[field]) && cleanText(secondary[field])) {
      merged[field] = cleanText(secondary[field]);
    }
  });

  numberFields.forEach((field) => {
    if (!Number.isFinite(merged[field]) && Number.isFinite(secondary[field])) {
      merged[field] = toFiniteNumber(secondary[field], NaN);
    }
  });

  merged.sourcePriority = Math.max(Number(primary.sourcePriority) || 0, Number(secondary.sourcePriority) || 0);
  merged.updatedAt = Math.max(Number(primary.updatedAt) || 0, Number(secondary.updatedAt) || 0);
  merged.historyHits = (Number(primary.historyHits) || 0) + (Number(secondary.historyHits) || 0);
  return merged;
}

function buildProductSuggestionPool(byProduct) {
  return Array.from(byProduct.entries())
    .map(([normalizedProduct, profiles]) => {
      const sorted = profiles
        .slice()
        .sort((a, b) => {
          if (a.sourcePriority !== b.sourcePriority) {
            return b.sourcePriority - a.sourcePriority;
          }
          if (a.historyHits !== b.historyHits) {
            return b.historyHits - a.historyHits;
          }
          return (b.updatedAt || 0) - (a.updatedAt || 0);
        });
      const head = sorted[0];
      const historyHits = sorted.reduce((sum, item) => sum + (Number(item.historyHits) || 0), 0);
      return {
        normalizedProduct,
        product: head?.product || "",
        head,
        historyHits,
        count: profiles.length,
      };
    })
    .sort((a, b) => {
      if (a.historyHits !== b.historyHits) {
        return b.historyHits - a.historyHits;
      }
      if (a.count !== b.count) {
        return b.count - a.count;
      }
      return a.product.localeCompare(b.product, "zh-CN");
    });
}

function refreshShipmentSmartFillSuggestions(options = {}) {
  if (!shipmentForm) {
    return;
  }
  const showPanel = Boolean(options.showPanel);

  renderSimpleStringSuggestions(
    shipmentManufacturerList,
    filterStringSuggestions(manufacturerSuggestionPool, cleanText(shipmentFieldEls.manufacturer?.value))
  );
  renderSimpleStringSuggestions(
    shipmentCustomerList,
    filterStringSuggestions(customerSuggestionPool, cleanText(shipmentFieldEls.customer?.value))
  );
  const profiles = renderProductSuggestions(cleanText(shipmentFieldEls.product?.value));
  if (showPanel) {
    renderProductSuggestionPanel(profiles);
  } else if (!isProductFieldFocused()) {
    hideProductSuggestionPanel();
  }
}

function renderSimpleStringSuggestions(datalist, values) {
  if (!datalist) {
    return;
  }

  datalist.innerHTML = values
    .map((value) => `<option value="${escapeHtml(value)}"></option>`)
    .join("");
}

function filterStringSuggestions(pool, keyword) {
  if (!pool.length) {
    return [];
  }

  const normalizedKeyword = normalizeSuggestText(keyword);
  if (!normalizedKeyword) {
    return pool.slice(0, FORM_SUGGESTION_LIMIT);
  }

  return pool
    .filter((item) => normalizeSuggestText(item).includes(normalizedKeyword))
    .slice(0, FORM_SUGGESTION_LIMIT);
}

function renderProductSuggestions(keyword) {
  if (!shipmentProductList) {
    return [];
  }

  const matched = getContextualProductSuggestions(keyword);
  latestProductSuggestionProfiles = matched;

  shipmentProductList.innerHTML = matched
    .map((profile) => {
      const priceText = Number.isFinite(profile.unitPrice) ? `${profile.unitPrice.toFixed(2)} 元` : "未设单价";
      const unitText = profile.measureUnit ? formatMeasureUnit(profile.measureUnit) : "未设单位";
      const label = `${profile.manufacturer || "-"} / ${profile.customer || "-"} / ${unitText} / ${priceText}`;
      return `<option value="${escapeHtml(profile.product)}" label="${escapeHtml(label)}"></option>`;
    })
    .join("");

  return matched;
}

function getContextualProductSuggestions(keyword) {
  const normalizedKeyword = normalizeSuggestText(keyword);
  const manufacturer = cleanText(shipmentFieldEls.manufacturer?.value);
  const customer = cleanText(shipmentFieldEls.customer?.value);
  const manufacturerKey = normalizeSuggestText(manufacturer);
  const customerKey = normalizeSuggestText(customer);

  const scored = [];
  shipmentProfileDataset.forEach((profile) => {
    const profileProductKey = normalizeSuggestText(profile.product);
    if (!profileProductKey) {
      return;
    }

    const profileManufacturerKey = normalizeSuggestText(profile.manufacturer);
    const profileCustomerKey = normalizeSuggestText(profile.customer);

    let score = (profile.sourcePriority || 0) * 1000 + (profile.historyHits || 0) * 10;
    let scoped = true;

    if (customerKey) {
      const customerExact = profileCustomerKey === customerKey;
      const customerFuzzy = profileCustomerKey.includes(customerKey);
      if (!customerExact && !customerFuzzy) {
        scoped = false;
      } else {
        score += customerExact ? 1000 : 120;
      }
    }

    if (manufacturerKey) {
      const manufacturerExact = profileManufacturerKey === manufacturerKey;
      const manufacturerFuzzy = profileManufacturerKey.includes(manufacturerKey);
      if (!manufacturerExact && !manufacturerFuzzy) {
        scoped = false;
      } else {
        score += manufacturerExact ? 500 : 80;
      }
    }

    if (!scoped) {
      return;
    }

    if (normalizedKeyword) {
      const startsWith = profileProductKey.startsWith(normalizedKeyword);
      const includes = profileProductKey.includes(normalizedKeyword);
      if (!startsWith && !includes) {
        return;
      }
      score += startsWith ? 400 : 160;
    } else {
      score += 60;
    }

    if (profile.updatedAt) {
      score += Math.min(profile.updatedAt / 100_000_000_000, 10);
    }

    scored.push({ profile, score });
  });

  if (!scored.length) {
    if (customerKey || manufacturerKey) {
      return [];
    }
    return getFallbackProductSuggestions(normalizedKeyword);
  }

  const byProduct = new Map();
  scored.forEach((item) => {
    const productKey = normalizeSuggestText(item.profile.product);
    const existing = byProduct.get(productKey);
    if (!existing || item.score > existing.score) {
      byProduct.set(productKey, item);
    }
  });

  return Array.from(byProduct.values())
    .sort((a, b) => b.score - a.score)
    .slice(0, FORM_SUGGESTION_LIMIT)
    .map((item) => item.profile);
}

function getFallbackProductSuggestions(normalizedKeyword) {
  return shipmentProductSuggestionPool
    .filter((item) => {
      if (!normalizedKeyword) {
        return true;
      }
      if (item.normalizedProduct.includes(normalizedKeyword)) {
        return true;
      }
      const manufacturerKey = normalizeSuggestText(item.head?.manufacturer);
      const customerKey = normalizeSuggestText(item.head?.customer);
      return manufacturerKey.includes(normalizedKeyword) || customerKey.includes(normalizedKeyword);
    })
    .sort((a, b) => {
      const aStarts = normalizedKeyword && a.normalizedProduct.startsWith(normalizedKeyword) ? 1 : 0;
      const bStarts = normalizedKeyword && b.normalizedProduct.startsWith(normalizedKeyword) ? 1 : 0;
      if (aStarts !== bStarts) {
        return bStarts - aStarts;
      }
      if (a.historyHits !== b.historyHits) {
        return b.historyHits - a.historyHits;
      }
      return b.count - a.count;
    })
    .slice(0, FORM_SUGGESTION_LIMIT)
    .map((item) => item.head)
    .filter(Boolean);
}

function renderProductSuggestionPanel(profiles) {
  if (!shipmentProductSuggestionPanel) {
    return;
  }

  clearProductSuggestionHideTimer();

  if (!profiles.length) {
    shipmentProductSuggestionPanel.innerHTML = '<div class="product-suggestion-empty">该客户暂无匹配产品</div>';
    shipmentProductSuggestionPanel.classList.remove("is-hidden");
    return;
  }

  shipmentProductSuggestionPanel.innerHTML = profiles
    .map((profile, index) => {
      const priceText = Number.isFinite(profile.unitPrice) ? `${profile.unitPrice.toFixed(2)} 元` : "未设单价";
      const unitText = profile.measureUnit ? formatMeasureUnit(profile.measureUnit) : "未设";
      const meta = `${profile.manufacturer || "-"} / ${profile.customer || "-"} / ${unitText} / 单价 ${priceText}`;
      return `
        <button type="button" class="product-suggestion-item" data-product-index="${index}">
          <span class="product-suggestion-name">${escapeHtml(profile.product || "-")}</span>
          <span class="product-suggestion-meta">${escapeHtml(meta)}</span>
        </button>
      `;
    })
    .join("");
  shipmentProductSuggestionPanel.classList.remove("is-hidden");
}

function handleProductSuggestionPanelClick(event) {
  const button = event.target.closest("[data-product-index]");
  if (!button) {
    return;
  }

  const index = Number(button.dataset.productIndex);
  if (!Number.isInteger(index) || index < 0 || index >= latestProductSuggestionProfiles.length) {
    return;
  }

  const profile = latestProductSuggestionProfiles[index];
  if (!profile) {
    return;
  }

  applyShipmentProfileBySelection(profile);
  hideProductSuggestionPanel();
  shipmentFieldEls.product?.focus();
}

function applyShipmentProfileBySelection(profile) {
  if (!shipmentForm || !profile) {
    return;
  }

  const profileUnit = normalizeMeasureUnit(profile.measureUnit, "");

  if (shipmentFieldEls.product) {
    shipmentFieldEls.product.value = cleanText(profile.product);
  }
  applyShipmentFieldIfEmpty(shipmentFieldEls.manufacturer, profile.manufacturer);
  applyShipmentFieldIfEmpty(shipmentFieldEls.customer, profile.customer);
  if (shipmentFieldEls.measureUnit && profileUnit) {
    shipmentFieldEls.measureUnit.value = profileUnit;
    configureShipmentQuantityInputByUnit();
  }

  if (shipmentFieldEls.unitPrice && Number.isFinite(profile.unitPrice)) {
    shipmentFieldEls.unitPrice.value = profile.unitPrice.toFixed(2);
    shipmentFieldEls.unitPrice.dataset.autofill = "1";
  }
  if (shipmentFieldEls.rebate && Number.isFinite(profile.rebate) && !cleanText(shipmentFieldEls.rebate.value)) {
    shipmentFieldEls.rebate.value = profile.rebate.toFixed(2);
  }
  if (shipmentFieldEls.weight && Number.isFinite(profile.weight) && !cleanText(shipmentFieldEls.weight.value)) {
    shipmentFieldEls.weight.value = profile.weight.toFixed(2);
  }
  if (shipmentFieldEls.cubic && Number.isFinite(profile.cubic) && !cleanText(shipmentFieldEls.cubic.value)) {
    shipmentFieldEls.cubic.value = profile.cubic.toFixed(2);
  }
  applyShipmentFieldIfEmpty(shipmentFieldEls.transit, profile.transit);
  applyShipmentFieldIfEmpty(shipmentFieldEls.delivery, profile.delivery);
  applyShipmentFieldIfEmpty(shipmentFieldEls.note, profile.note);
  syncShipmentNoteSelectionFromField();

  if (shipmentFieldEls.amount) {
    shipmentFieldEls.amount.dataset.manual = "0";
  }
  syncShipmentAmountFromUnitPrice({ force: true });
  persistRetainedFromForm(shipmentForm, STORAGE_KEYS.shipmentRetain, SHIPMENT_RETAIN_FIELDS);
}

function scheduleHideProductSuggestionPanel() {
  clearProductSuggestionHideTimer();
  productSuggestionHideTimer = window.setTimeout(() => {
    hideProductSuggestionPanel();
  }, 140);
}

function clearProductSuggestionHideTimer() {
  if (productSuggestionHideTimer) {
    clearTimeout(productSuggestionHideTimer);
    productSuggestionHideTimer = null;
  }
}

function hideProductSuggestionPanel() {
  if (!shipmentProductSuggestionPanel) {
    return;
  }
  shipmentProductSuggestionPanel.classList.add("is-hidden");
}

function isProductFieldFocused() {
  return Boolean(shipmentFieldEls.product && document.activeElement === shipmentFieldEls.product);
}

function tryAutoFillShipmentFields() {
  if (!shipmentForm || !shipmentFieldEls.product) {
    return;
  }

  const profile = findBestShipmentProfile({
    manufacturer: cleanText(shipmentFieldEls.manufacturer?.value),
    customer: cleanText(shipmentFieldEls.customer?.value),
    product: cleanText(shipmentFieldEls.product.value),
  });

  if (!profile) {
    return;
  }

  const profileUnit = normalizeMeasureUnit(profile.measureUnit, "");
  let changed = false;
  changed = applyShipmentFieldIfEmpty(shipmentFieldEls.manufacturer, profile.manufacturer) || changed;
  changed = applyShipmentFieldIfEmpty(shipmentFieldEls.customer, profile.customer) || changed;
  if (shipmentFieldEls.measureUnit && profileUnit) {
    const currentUnit =
      normalizeMeasureUnit(shipmentFieldEls.measureUnit.value, DEFAULT_MEASURE_UNIT) || DEFAULT_MEASURE_UNIT;
    if (currentUnit === DEFAULT_MEASURE_UNIT) {
      shipmentFieldEls.measureUnit.value = profileUnit;
      configureShipmentQuantityInputByUnit();
      changed = true;
    }
  }
  changed = applyShipmentUnitPriceFromProfile(shipmentFieldEls.unitPrice, profile.unitPrice) || changed;
  changed = applyShipmentNumberFieldIfEmpty(shipmentFieldEls.rebate, profile.rebate) || changed;
  changed = applyShipmentNumberFieldIfEmpty(shipmentFieldEls.weight, profile.weight) || changed;
  changed = applyShipmentNumberFieldIfEmpty(shipmentFieldEls.cubic, profile.cubic) || changed;
  changed = applyShipmentFieldIfEmpty(shipmentFieldEls.transit, profile.transit) || changed;
  changed = applyShipmentFieldIfEmpty(shipmentFieldEls.delivery, profile.delivery) || changed;
  changed = applyShipmentFieldIfEmpty(shipmentFieldEls.pickup, profile.pickup) || changed;
  const noteChanged = applyShipmentFieldIfEmpty(shipmentFieldEls.note, profile.note);
  changed = noteChanged || changed;

  if (changed) {
    if (noteChanged) {
      syncShipmentNoteSelectionFromField();
    }
    syncShipmentAmountFromUnitPrice();
    persistRetainedFromForm(shipmentForm, STORAGE_KEYS.shipmentRetain, SHIPMENT_RETAIN_FIELDS);
    refreshShipmentSmartFillSuggestions();
  }
}

function findBestShipmentProfile(criteria) {
  const productKey = normalizeSuggestText(criteria.product);
  if (!productKey) {
    return null;
  }

  const manufacturerKey = normalizeSuggestText(criteria.manufacturer);
  const customerKey = normalizeSuggestText(criteria.customer);

  if (manufacturerKey && customerKey) {
    const exact = shipmentProfileByExactKey.get(makeShipmentProfileKey(criteria.manufacturer, criteria.customer, criteria.product));
    if (exact) {
      return exact;
    }
  }

  const candidates = shipmentProfilesByProductKey.get(productKey) || [];
  if (!candidates.length) {
    return null;
  }

  let best = null;
  let bestScore = -Infinity;
  candidates.forEach((candidate) => {
    const candidateManufacturerKey = normalizeSuggestText(candidate.manufacturer);
    const candidateCustomerKey = normalizeSuggestText(candidate.customer);

    if (manufacturerKey && candidateManufacturerKey !== manufacturerKey) {
      return;
    }
    if (customerKey && candidateCustomerKey !== customerKey) {
      return;
    }

    let score = (candidate.sourcePriority || 0) * 1000 + (candidate.historyHits || 0) * 10;

    if (candidate.updatedAt) {
      score += Math.min(candidate.updatedAt / 100_000_000_000, 10);
    }

    if (score > bestScore) {
      best = candidate;
      bestScore = score;
    }
  });

  return best;
}

function applyShipmentFieldIfEmpty(field, value) {
  if (!field) {
    return false;
  }
  if (cleanText(field.value) || !cleanText(value)) {
    return false;
  }
  field.value = cleanText(value);
  return true;
}

function applyShipmentNumberFieldIfEmpty(field, value) {
  if (!field) {
    return false;
  }
  if (cleanText(field.value) || !Number.isFinite(value)) {
    return false;
  }
  field.value = value.toFixed(2);
  return true;
}

function applyShipmentUnitPriceFromProfile(field, value) {
  if (!field || !Number.isFinite(value)) {
    return false;
  }

  const nextText = value.toFixed(2);
  const currentText = cleanText(field.value);
  if (!currentText) {
    field.value = nextText;
    field.dataset.autofill = "1";
    return true;
  }

  const currentValue = parseOptionalNonNegative(currentText);
  const autoFilled = field.dataset.autofill === "1";
  if (!autoFilled) {
    return false;
  }

  if (!Number.isFinite(currentValue) || Math.abs(currentValue - value) > 1e-9) {
    field.value = nextText;
    field.dataset.autofill = "1";
    return true;
  }

  return false;
}

function configureShipmentQuantityInputByUnit() {
  if (!shipmentFieldEls.quantity) {
    return;
  }

  const unit = normalizeMeasureUnit(shipmentFieldEls.measureUnit?.value, DEFAULT_MEASURE_UNIT) || DEFAULT_MEASURE_UNIT;
  if (shipmentQuantityLabel) {
    shipmentQuantityLabel.textContent = `计量值(${formatMeasureUnit(unit)})`;
  }
  shipmentFieldEls.quantity.setAttribute("min", "0");
  shipmentFieldEls.quantity.setAttribute("step", "0.001");
  if (unit === "piece") {
    shipmentFieldEls.quantity.setAttribute("placeholder", "按件填写（支持 0.001）");
  } else if (unit === "ton") {
    shipmentFieldEls.quantity.setAttribute("placeholder", "按KG填写（支持 0.001）");
  } else {
    shipmentFieldEls.quantity.setAttribute("placeholder", "按平方填写（支持 0.001）");
  }
}

function syncShipmentAmountFromUnitPrice(options = {}) {
  if (!shipmentFieldEls.amount) {
    return;
  }
  const force = Boolean(options.force);
  const isManual = shipmentFieldEls.amount.dataset.manual === "1";
  if (isManual && !force) {
    return;
  }

  const quantity = parseOptionalNonNegative(shipmentFieldEls.quantity?.value);
  const unitPrice = parseOptionalNonNegative(shipmentFieldEls.unitPrice?.value);
  if (!Number.isFinite(quantity) || !Number.isFinite(unitPrice) || quantity <= 0) {
    if (force && !isManual) {
      shipmentFieldEls.amount.value = "";
    }
    return;
  }
  shipmentFieldEls.amount.value = (quantity * unitPrice).toFixed(2);
  shipmentFieldEls.amount.dataset.manual = "0";
}

function makeShipmentProfileKey(manufacturer, customer, product) {
  const manufacturerKey = normalizeSuggestText(manufacturer);
  const customerKey = normalizeSuggestText(customer);
  const productKey = normalizeSuggestText(product);
  if (!productKey) {
    return "";
  }
  return `${manufacturerKey}|${customerKey}|${productKey}`;
}

function normalizeSuggestText(value) {
  return cleanText(value)
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/[()（）_\-./|:：,，]/g, "");
}

function cleanCatalogPartyName(value) {
  return cleanText(value)
    .replace(/\u3000/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function cleanCatalogProductText(value) {
  const raw = cleanText(value);
  if (!raw) {
    return "";
  }

  let text = raw
    .replace(/\u3000/g, " ")
    .replace(/[|｜]/g, " ")
    .replace(/[\t\r\n]+/g, " ");

  text = text.replace(
    /(中转另算|包转另算|运费已付|家结|提付|现付|月结|北结|南结|厂付|本单不算钱|上单库存|库存|样品|配套|纸箱|空箱|袋子|箱头|等通知放货|签回单|包送货|送货上门)/g,
    " "
  );

  text = text.replace(/\d+(?:\.\d+)?\s*(?:\*|x|X|×|＋|\+)\s*\d+(?:\.\d+)?/g, " ");
  text = text.replace(/\d+(?:\.\d+)?\s*(?:元|毛|角)(?=[\u4e00-\u9fffA-Za-z])/g, " ");
  text = text.replace(
    /\d+(?:\.\d+)?\s*(?:件|方|吨|公斤|kg|KG|g|G|克|斤|包|箱|捆|板|托|支|袋|盒|个|连|串|条|杯|桶|瓶|打|组|朵|片)/g,
    " "
  );
  text = text.replace(/(?:标|第)\d+/g, " ");
  text = text.replace(/[0-9０-９]+/g, " ");
  text = text.replace(/[*/xX×+\-]+/g, " ");
  text = text.replace(/[()（）[\]【】]/g, " ");
  text = text.replace(/[，,;；、]+/g, " ");
  text = text.replace(/\s+/g, " ").trim();
  const stripped = raw.replace(/[0-9０-９]+/g, " ").replace(/\s+/g, " ").trim();

  if (!text) {
    return stripped;
  }
  if (!/[\u4e00-\u9fffA-Za-z]/.test(text)) {
    return stripped;
  }

  return text;
}

function normalizeBuiltInCatalogRows(rows) {
  if (!Array.isArray(rows)) {
    return [];
  }

  const normalized = [];
  rows.forEach((row) => {
    if (!Array.isArray(row) || row.length < 3) {
      return;
    }
    const manufacturer = cleanCatalogPartyName(row[0]);
    const customer = cleanCatalogPartyName(row[1]);
    const { product, unitPriceValue } = resolveCatalogProductAndPrice(row);
    if (isCatalogNoiseRow(manufacturer, customer, product)) {
      return;
    }
    if (!manufacturer && !customer && !product) {
      return;
    }

    normalized.push({
      manufacturer,
      customer,
      product,
      measureUnit: normalizeMeasureUnit(row[9], ""),
      unitPrice: unitPriceValue,
      note: cleanText(row[4]),
      weight: parseCatalogNumber(row[5]),
      cubic: parseCatalogNumber(row[10]),
      delivery: cleanText(row[6]),
      rebate: parseCatalogNumber(row[7]),
      transit: cleanText(row[8]),
    });
  });
  return normalized;
}

function isCatalogNoiseRow(manufacturer, customer, product) {
  const m = cleanText(manufacturer);
  const c = cleanText(customer);
  const p = cleanText(product);

  if (!m && !c && !p) {
    return true;
  }

  const headerRow = normalizeSuggestText(m) === "厂家" && normalizeSuggestText(c) === "客户";
  if (headerRow) {
    return true;
  }

  const summaryBag = `${m} ${c} ${p}`;
  if (/(合计|月结|现付|北结|南结|小计)/.test(summaryBag)) {
    return true;
  }

  return false;
}

function resolveCatalogProductAndPrice(row) {
  const optionA = scoreCatalogLayout(row[2], row[3]);
  const optionB = scoreCatalogLayout(row[3], row[2]);
  let picked = optionA;
  let fallback = optionB;

  if (optionB.score > optionA.score) {
    picked = optionB;
    fallback = optionA;
  }

  if (!Number.isFinite(picked.unitPriceValue) && Number.isFinite(fallback.unitPriceValue)) {
    picked = fallback;
  }
  if (!picked.product && fallback.product) {
    picked = fallback;
  }

  return {
    product: picked.product || cleanCatalogProductText(picked.rawProductText),
    unitPriceValue: picked.unitPriceValue,
  };
}

function scoreCatalogLayout(productText, unitPriceRaw) {
  const rawProductText = cleanText(productText);
  const product = cleanCatalogProductText(rawProductText);
  const unitPriceValue = parseCatalogNumber(unitPriceRaw);
  let score = 0;

  if (product) {
    score += 2;
  } else if (rawProductText) {
    score -= 2;
  }
  if (/[\u4e00-\u9fffA-Za-z]/.test(product)) {
    score += 2;
  }
  if (/\d/.test(product)) {
    score -= 2;
  }
  if (isLikelyPriceCellText(rawProductText)) {
    score -= 3;
  }

  const normalizedProductKey = normalizeSuggestText(product);
  if (/^(厂家|客户|产品|品名|单价|单格|价格|price)$/.test(normalizedProductKey)) {
    score -= 20;
  }

  if (Number.isFinite(unitPriceValue)) {
    score += 2;
  }
  if (isLikelyPriceCellText(unitPriceRaw)) {
    score += 1;
  }

  return { score, unitPriceValue, product, rawProductText };
}

function isLikelyPriceCellText(value) {
  const text = cleanText(value);
  if (!text || !/\d/.test(text)) {
    return false;
  }

  return /^[0-9.+\-*/xX×()（）\s/%¥￥元方吨件公斤kgKG送货家结运费提付中转包转现付月结北结南结]+$/.test(text);
}

function parseCatalogNumber(value) {
  const rawText = cleanText(value);
  if (!rawText) {
    return NaN;
  }

  const direct = parseOptionalNonNegative(rawText);
  if (Number.isFinite(direct)) {
    return direct;
  }

  const text = rawText.replaceAll(",", "").replaceAll("¥", "").replaceAll("￥", "");
  if (!text) {
    return NaN;
  }

  const matched = text.match(/\d+(\.\d+)?/);
  if (!matched) {
    return NaN;
  }

  const parsed = Number(matched[0]);
  if (!Number.isFinite(parsed) || parsed < 0) {
    return NaN;
  }
  return parsed;
}

function handleSummarySubmit(event) {
  event.preventDefault();
  renderDailySummarySection({ resetPages: true });
}

function syncSummaryKeywordInputs() {
  if (summaryShipmentSearchInput && summaryShipmentSearchInput.value !== summaryShipmentKeyword) {
    summaryShipmentSearchInput.value = summaryShipmentKeyword;
  }
  if (summaryFinanceSearchInput && summaryFinanceSearchInput.value !== summaryFinanceKeyword) {
    summaryFinanceSearchInput.value = summaryFinanceKeyword;
  }
}

function clearSummaryQuickRangeActive() {
  if (!summaryQuickRangeButtons.length) {
    return;
  }
  summaryQuickRangeButtons.forEach((button) => {
    button.classList.remove("is-active");
    button.setAttribute("aria-pressed", "false");
  });
}

function setSummaryQuickRangeActive(rangeType) {
  if (!summaryQuickRangeButtons.length) {
    return;
  }
  summaryQuickRangeButtons.forEach((button) => {
    const active = button.dataset.summaryQuickRange === rangeType;
    button.classList.toggle("is-active", active);
    button.setAttribute("aria-pressed", String(active));
  });
}

function setSummaryRange(from, to) {
  if (summaryFromDateInput) {
    summaryFromDateInput.value = from || "";
  }
  if (summaryToDateInput) {
    summaryToDateInput.value = to || "";
  }
}

function applySummaryQuickRange(rangeType) {
  const todayDate = new Date();
  const today = dateValue(todayDate);
  setSummaryQuickRangeActive(rangeType);

  if (rangeType === "today") {
    setSummaryRange(today, today);
    renderDailySummarySection({ resetPages: true });
    return;
  }

  if (rangeType === "yesterday") {
    const yesterdayDate = new Date(todayDate);
    yesterdayDate.setDate(yesterdayDate.getDate() - 1);
    const yesterday = dateValue(yesterdayDate);
    setSummaryRange(yesterday, yesterday);
    renderDailySummarySection({ resetPages: true });
    return;
  }

  if (rangeType === "last7") {
    const startDate = new Date(todayDate);
    startDate.setDate(startDate.getDate() - 6);
    setSummaryRange(dateValue(startDate), today);
    renderDailySummarySection({ resetPages: true });
    return;
  }

  if (rangeType === "lastMonth") {
    const year = todayDate.getFullYear();
    const month = todayDate.getMonth();
    const startDate = new Date(year, month - 1, 1);
    const endDate = new Date(year, month, 0);
    setSummaryRange(dateValue(startDate), dateValue(endDate));
    renderDailySummarySection({ resetPages: true });
    return;
  }

  resetSummaryRangeToCurrentMonth();
}

function resetSummaryRangeToCurrentMonth() {
  const today = dateValue(new Date());
  const thisMonth = today.slice(0, 7);
  setSummaryQuickRangeActive("thisMonth");
  setSummaryRange(`${thisMonth}-01`, today);
  renderDailySummarySection({ resetPages: true });
}

function clearShipmentDateRange() {
  if (shipmentFromDateInput) {
    shipmentFromDateInput.value = "";
  }
  if (shipmentToDateInput) {
    shipmentToDateInput.value = "";
  }
  renderShipmentTable({ resetPage: true });
}

function clearFinanceDateRange() {
  if (financeFromDateInput) {
    financeFromDateInput.value = "";
  }
  if (financeToDateInput) {
    financeToDateInput.value = "";
  }
  renderFinanceTable({ resetPage: true });
}

function clearStatusDateRange() {
  if (statusFromDateInput) {
    statusFromDateInput.value = "";
  }
  if (statusToDateInput) {
    statusToDateInput.value = "";
  }
  renderStatusSection({ resetPage: true });
}

function clearStatusActionRecordDateRange() {
  if (statusActionRecordFromDateInput) {
    statusActionRecordFromDateInput.value = "";
  }
  if (statusActionRecordToDateInput) {
    statusActionRecordToDateInput.value = "";
  }
  renderStatusSection({ resetRecordPage: true });
}

function setupShipmentNoteQuickFill() {
  if (!shipmentForm || !shipmentFieldEls.note) {
    return;
  }

  const syncToField = () => {
    applyShipmentNoteSelectionToField();
  };

  shipmentNoteOptionInputs.forEach((input) => {
    input.addEventListener("change", syncToField);
  });

  if (shipmentNoteCustomInput) {
    shipmentNoteCustomInput.addEventListener("input", syncToField);
    shipmentNoteCustomInput.addEventListener("change", syncToField);
  }

  syncShipmentNoteSelectionFromField();
}

function splitShipmentNoteTokens(value) {
  return String(value || "")
    .split(/[;；,，、\n\r]+/)
    .map((item) => cleanText(item))
    .filter(Boolean);
}

function applyShipmentNoteSelectionToField() {
  if (!shipmentFieldEls.note) {
    return;
  }

  const merged = [];
  shipmentNoteOptionInputs.forEach((input) => {
    const text = cleanText(input.value);
    if (input.checked && text && !merged.includes(text)) {
      merged.push(text);
    }
  });

  splitShipmentNoteTokens(shipmentNoteCustomInput?.value).forEach((text) => {
    if (!merged.includes(text)) {
      merged.push(text);
    }
  });

  shipmentFieldEls.note.value = merged.join(SHIPMENT_NOTE_SEPARATOR);
}

function syncShipmentNoteSelectionFromField() {
  if (!shipmentFieldEls.note) {
    return;
  }

  const tokens = splitShipmentNoteTokens(shipmentFieldEls.note.value);
  const tokenKeys = new Set(tokens.map((token) => normalizeSuggestText(token)));
  const optionKeys = new Set();

  shipmentNoteOptionInputs.forEach((input) => {
    const key = normalizeSuggestText(input.value);
    const checked = tokenKeys.has(key);
    input.checked = checked;
    if (checked) {
      optionKeys.add(key);
    }
  });

  if (shipmentNoteCustomInput) {
    const customTokens = tokens.filter((token) => !optionKeys.has(normalizeSuggestText(token)));
    shipmentNoteCustomInput.value = customTokens.join(SHIPMENT_NOTE_SEPARATOR);
  }

  applyShipmentNoteSelectionToField();
}

function buildShipmentSubmitConfirmMessage(record) {
  const lines = [
    "请确认本次收发录入信息：",
    `日期：${record.date || "-"}`,
    `序号：${record.serialNo || "-"}`,
    `厂家：${record.manufacturer || "-"}`,
    `客户：${record.customer || "-"}`,
    `产品：${record.product || "-"}`,
    `计量值：${formatMeasureValue(record.quantity, record.measureUnit)} ${formatMeasureUnit(record.measureUnit)}`,
    `单价：¥${formatDecimal(Math.max(0, toFiniteNumber(record.unitPrice)))}`,
    `金额：¥${formatDecimal(Math.max(0, toFiniteNumber(record.amount)))}`,
    `单位重量：${formatDecimal(Math.max(0, toFiniteNumber(record.weight)))} KG`,
    `单位立方：${formatDecimal(Math.max(0, toFiniteNumber(record.cubic)))} m3`,
    `备注：${record.note || "-"}`,
    "",
    "点击“确定”后提交，点击“取消”返回修改。",
  ];
  return lines.join("\n");
}

function buildFinanceSubmitConfirmMessage(record) {
  const isIncome = record.type === "income";
  const lines = [
    "请确认本次收支录入信息：",
    `日期：${record.date || "-"}`,
    `摘要：${record.summary || "-"}`,
    `类型：${isIncome ? "收入" : "支出"}`,
    `金额：¥${formatDecimal(Math.max(0, toFiniteNumber(record.amount)))}`,
    `备注：${record.note || "-"}`,
    "",
    "点击“确定”后提交，点击“取消”返回修改。",
  ];
  return lines.join("\n");
}

function resetShipmentFormAfterSubmit() {
  if (!shipmentForm) {
    return;
  }

  shipmentForm.reset();
  if (shipmentForm.elements.date) {
    shipmentForm.elements.date.value = dateValue(new Date());
  }
  if (shipmentForm.elements.type) {
    shipmentForm.elements.type.value = "receive";
  }
  if (shipmentFieldEls.measureUnit) {
    shipmentFieldEls.measureUnit.value = DEFAULT_MEASURE_UNIT;
  }
  configureShipmentQuantityInputByUnit();
  if (shipmentFieldEls.amount) {
    shipmentFieldEls.amount.dataset.manual = "0";
  }
  syncShipmentAmountFromUnitPrice({ force: true });
  syncShipmentNoteSelectionFromField();
  hideProductSuggestionPanel();
  persistRetainedFromForm(shipmentForm, STORAGE_KEYS.shipmentRetain, SHIPMENT_RETAIN_FIELDS);
}

function resetFinanceFormAfterSubmit() {
  if (!financeForm) {
    return;
  }

  financeForm.reset();
  if (financeForm.elements.date) {
    financeForm.elements.date.value = dateValue(new Date());
  }
  persistRetainedFromForm(financeForm, STORAGE_KEYS.financeRetain, FINANCE_RETAIN_FIELDS);
}

async function handleShipmentSubmit(event) {
  event.preventDefault();

  const fd = new FormData(shipmentForm);
  const measureUnit = normalizeMeasureUnit(fd.get("measureUnit"), DEFAULT_MEASURE_UNIT) || DEFAULT_MEASURE_UNIT;
  const quantity = Number(fd.get("quantity"));
  const unitPrice = parseOptionalNonNegative(fd.get("unitPrice"));
  const amountText = cleanText(fd.get("amount"));
  const amountRaw = parseOptionalNonNegative(amountText);
  const rebate = parseOptionalNonNegative(fd.get("rebate"));
  const weight = parseOptionalNonNegative(fd.get("weight"));
  const cubic = parseOptionalNonNegative(fd.get("cubic"));

  if (!Number.isFinite(quantity) || quantity <= 0) {
    alert("计量值必须大于 0。");
    return;
  }

  if (!Number.isFinite(unitPrice)) {
    alert("单价必须是大于等于 0 的数字。");
    return;
  }
  if (!Number.isFinite(amountRaw)) {
    alert("金额必须是大于等于 0 的数字。");
    return;
  }
  if (!Number.isFinite(rebate)) {
    alert("回扣必须是大于等于 0 的数字。");
    return;
  }
  if (!Number.isFinite(weight)) {
    alert("单位重量（KG）必须是大于等于 0 的数字。");
    return;
  }
  if (!Number.isFinite(cubic)) {
    alert("单位立方（m3）必须是大于等于 0 的数字。");
    return;
  }

  const normalizedQuantity = Number(quantity.toFixed(3));
  const normalizedWeight = Number(weight.toFixed(3));
  const normalizedCubic = Number(cubic.toFixed(3));
  const amount = amountText ? amountRaw : normalizedQuantity * unitPrice;

  const record = {
    id: makeId(),
    type: normalizeShipmentType(fd.get("type")) || "receive",
    date: String(fd.get("date") || dateValue(new Date())),
    serialNo: cleanText(fd.get("serialNo")),
    manufacturer: cleanText(fd.get("manufacturer")),
    customer: cleanText(fd.get("customer")),
    product: cleanText(fd.get("product")),
    measureUnit,
    quantity: normalizedQuantity,
    unitPrice,
    amount,
    rebate,
    weight: normalizedWeight,
    cubic: normalizedCubic,
    transit: cleanText(fd.get("transit")),
    delivery: cleanText(fd.get("delivery")),
    pickup: cleanText(fd.get("pickup")),
    note: cleanText(fd.get("note")),
    isLoaded: false,
    loadedAt: 0,
    dispatchDate: "",
    dispatchTruckNo: "",
    dispatchPlateNo: "",
    dispatchDriver: "",
    dispatchContactName: "",
    dispatchDestination: "",
    destination: "",
    homeSettleStatus: "",
    homeSettleStatusUpdatedAt: 0,
    homeSettleRemindedAt: 0,
    homeSettleSnoozeUntil: 0,
    homeSettleLastAction: "",
    homeSettleLastActionAt: 0,
    isPaid: false,
    paidAt: 0,
    createdAt: Date.now(),
  };

  if (!(await confirmDialog(buildShipmentSubmitConfirmMessage(record), { title: "确认保存发货记录" }))) {
    return;
  }

  shipments.push(record);
  upsertCustomCatalogProfile(record);
  writeStorage(STORAGE_KEYS.shipments, shipments);
  resetShipmentFormAfterSubmit();
  renderShipmentSection();
}

async function handleFinanceSubmit(event) {
  event.preventDefault();

  const fd = new FormData(financeForm);
  const summary = cleanText(fd.get("category"));
  const incomeText = cleanText(fd.get("income"));
  const expenseText = cleanText(fd.get("expense"));
  const income = incomeText ? parseOptionalNonNegative(incomeText) : 0;
  const expense = expenseText ? parseOptionalNonNegative(expenseText) : 0;

  if (!summary) {
    alert("请填写摘要。");
    return;
  }

  if ((incomeText && !Number.isFinite(income)) || (expenseText && !Number.isFinite(expense))) {
    alert("收入和支出必须是大于等于 0 的数字。");
    return;
  }

  if (income > 0 && expense > 0) {
    alert("一条记录只能填写收入或支出其中一项。");
    return;
  }

  if (income <= 0 && expense <= 0) {
    alert("请至少填写一项：收入或支出。");
    return;
  }

  const isIncome = income > 0;
  const amount = isIncome ? income : expense;

  const record = {
    id: makeId(),
    type: isIncome ? "income" : "expense",
    date: String(fd.get("date") || dateValue(new Date())),
    category: summary,
    summary,
    amount,
    note: cleanText(fd.get("note")),
    createdAt: Date.now(),
  };

  if (!(await confirmDialog(buildFinanceSubmitConfirmMessage(record), { title: "确认保存收支账目" }))) {
    return;
  }

  financeRecords.push(record);
  writeStorage(STORAGE_KEYS.finance, financeRecords);
  resetFinanceFormAfterSubmit();
  renderFinanceSection();
}

function renderAll() {
  renderShipmentSection();
  renderDispatchSection();
  renderFinanceSection();
  renderStatusSection();
  renderDailySummarySection();
}

function renderShipmentSection() {
  rebuildShipmentProfileIndex();
  renderShipmentStats();
  renderShipmentTable();
  renderDispatchSection();
  renderStatusSection();
  renderDailySummarySection();
}

function renderDispatchSection() {
  if (dispatchLoadStage !== "meta") {
    closeDispatchMetaModal();
  }
  renderDispatchTable();
  renderDispatchRecordTable();
  syncDispatchRecordDetailView();
}

function renderFinanceSection() {
  renderFinanceStats();
  renderFinanceTable();
  renderStatusSection();
  renderDailySummarySection();
}

function renderStatusSection(options = {}) {
  if (!statusTableBody) {
    return;
  }

  if (options.resetPage) {
    statusPage = 1;
  }
  if (options.resetRecordPage) {
    statusActionRecordPage = 1;
  }

  normalizeDateInputs(statusFromDateInput, statusToDateInput);
  normalizeDateInputs(statusActionRecordFromDateInput, statusActionRecordToDateInput);
  const rows = getFilteredStatusRows();
  renderStatusStats(getStatusStatsSnapshot());
  const paged = paginateRows(rows, statusPage, MAIN_PAGE_SIZE);
  statusPage = paged.page;
  currentVisibleStatusRows = paged.items.slice();
  updateMainPagination("status", paged.page, paged.totalPages, paged.totalRows);

  if (!rows.length) {
    currentVisibleStatusRows = [];
    statusTableBody.innerHTML = '<tr><td class="empty" colspan="7">暂无符合条件的状态记录</td></tr>';
  } else {
    statusTableBody.innerHTML = paged.items
      .map((row) => {
        const item = row.item;
        const itemId = cleanText(item?.id);
        const isPaid = normalizeShipmentPaidFlag(item?.isPaid);
        const debtAmount = isPaid
          ? 0
          : Math.max(0, toFiniteNumber(item?.amount) - Math.max(0, toFiniteNumber(item?.rebate)));

        return `
          <tr data-shipment-id="${escapeHtml(itemId)}">
            <td>${escapeHtml(item.date || "-")}</td>
            <td class="mono">${escapeHtml(item.serialNo || "-")}</td>
            <td>${escapeHtml(item.manufacturer || "-")}</td>
            <td>${escapeHtml(item.customer || "-")}</td>
            <td>${escapeHtml(item.product || "-")}</td>
            <td class="mono">¥${formatDecimal(debtAmount)}</td>
            <td>
              <div class="status-row-actions">
                <button
                  type="button"
                  class="ghost-btn status-action-btn status-action-snooze"
                  data-status-action="snooze"
                  data-status-action-id="${escapeHtml(itemId)}"
                >
                  稍后处理(1小时)
                </button>
                <button
                  type="button"
                  class="ghost-btn status-action-btn status-action-contact"
                  data-status-action="contact"
                  data-status-action-id="${escapeHtml(itemId)}"
                >
                  马上联系
                </button>
                <button
                  type="button"
                  class="ghost-btn status-action-btn status-action-paid"
                  data-status-action="paid"
                  data-status-action-id="${escapeHtml(itemId)}"
                >
                  已收款
                </button>
                <button
                  type="button"
                  class="ghost-btn status-view-detail-btn"
                  data-status-detail-id="${escapeHtml(itemId)}"
                >
                  查看详情
                </button>
              </div>
            </td>
          </tr>
        `;
      })
      .join("");
  }

  renderStatusActionRecordTable();
  syncStatusDetailView();
}

function handleStatusTableClick(event) {
  const actionTrigger = event.target.closest("[data-status-action][data-status-action-id]");
  if (actionTrigger) {
    const shipmentId = cleanText(actionTrigger.dataset.statusActionId);
    const action = cleanText(actionTrigger.dataset.statusAction);
    if (shipmentId && action) {
      handleStatusRowAction(shipmentId, action);
    }
    return;
  }

  const detailTrigger = event.target.closest("[data-status-detail-id]");
  if (detailTrigger) {
    const shipmentId = cleanText(detailTrigger.dataset.statusDetailId);
    if (shipmentId) {
      openStatusDetailModal(shipmentId);
    }
    return;
  }

  const rowTrigger = event.target.closest("tr[data-shipment-id]");
  if (!rowTrigger) {
    return;
  }

  const shipmentId = cleanText(rowTrigger.dataset.shipmentId);
  if (!shipmentId) {
    return;
  }

  openStatusDetailModal(shipmentId);
}

function handleStatusRowAction(shipmentId, action) {
  const target = shipments.find((item) => cleanText(item?.id) === cleanText(shipmentId));
  if (!target) {
    alert("未找到对应货物，可能已删除。");
    return;
  }

  const now = Date.now();
  const debtBefore = Math.max(0, toFiniteNumber(target?.amount) - Math.max(0, toFiniteNumber(target?.rebate)));
  if (action === "snooze") {
    setShipmentHomeSettleStatusOverride(target, "reminded");
    target.homeSettleRemindedAt = now;
    target.homeSettleSnoozeUntil = now + HOME_SETTLE_SNOOZE_MS;
    target.homeSettleLastAction = "稍后处理";
    target.homeSettleLastActionAt = now;
    writeStorage(STORAGE_KEYS.shipments, shipments);
    appendHomeSettleActionRecord(target, "snooze", {
      actionAt: now,
      debtAmountBefore: debtBefore,
      debtAmountAfter: debtBefore,
    });
    renderStatusSection();
    return;
  }

  if (action === "contact") {
    setShipmentHomeSettleStatusOverride(target, "reminded");
    target.homeSettleRemindedAt = now;
    target.homeSettleSnoozeUntil = 0;
    target.homeSettleLastAction = "马上联系";
    target.homeSettleLastActionAt = now;
    writeStorage(STORAGE_KEYS.shipments, shipments);
    appendHomeSettleActionRecord(target, "contact", {
      actionAt: now,
      debtAmountBefore: debtBefore,
      debtAmountAfter: debtBefore,
    });
    renderStatusSection();
    return;
  }

  if (action === "paid") {
    clearShipmentHomeSettleStatusOverride(target);
    target.homeSettleRemindedAt = now;
    target.homeSettleSnoozeUntil = 0;
    target.homeSettleLastAction = "已收款";
    target.homeSettleLastActionAt = now;
    const { financeChanged } = setShipmentPaidState(target, true, { skipRender: true });
    appendHomeSettleActionRecord(target, "paid", {
      actionAt: now,
      debtAmountBefore: debtBefore,
      debtAmountAfter: 0,
    });
    if (currentStatusDetailShipmentId && cleanText(currentStatusDetailShipmentId) === cleanText(shipmentId)) {
      closeStatusDetailModal();
    }
    renderShipmentSection();
    if (financeChanged) {
      renderFinanceSection();
    }
    return;
  }
}

function findLinkedPaymentFinanceRecord(shipment) {
  const shipmentId = cleanText(shipment?.id);
  const shipmentSerialNo = cleanText(shipment?.serialNo);
  if (!shipmentId && !shipmentSerialNo) {
    return null;
  }

  return (
    financeRecords.find((record) => {
      if (!isShipmentPaymentAutoFinanceRecord(record)) {
        return false;
      }
      const sourceShipmentId = cleanText(record?.sourceShipmentId);
      if (sourceShipmentId && shipmentId) {
        return sourceShipmentId === shipmentId;
      }
      return shipmentSerialNo && cleanText(record?.sourceSerialNo) === shipmentSerialNo;
    }) || null
  );
}

function renderStatusDetailModal(item, row = null) {
  if (!statusDetailModalContent || !item) {
    return;
  }

  const statusResolved = row?.statusId
    ? { id: row.statusId, source: row.source || "auto" }
    : getShipmentHomeSettleStatus(item);
  const statusMeta = getHomeSettleStatusMeta(statusResolved.id);
  const linkedFinanceRecord = findLinkedPaymentFinanceRecord(item);
  const isPaid = normalizeShipmentPaidFlag(item?.isPaid);
  const amount = Math.max(0, toFiniteNumber(item?.amount));
  const rebate = Math.max(0, toFiniteNumber(item?.rebate));
  const debtAmount = isPaid ? 0 : Math.max(0, amount - rebate);
  const dueTimestamp = getShipmentHomeSettleDueTimestamp(item);
  const dueDate = dueTimestamp > 0 ? dateValue(new Date(dueTimestamp)) : "-";
  const overdueDays = dueTimestamp > 0 ? Math.max(0, Math.floor((Date.now() - dueTimestamp) / (24 * 60 * 60 * 1000))) : 0;
  const remindedAt = Math.max(0, toFiniteNumber(item?.homeSettleRemindedAt));
  const snoozeUntil = getShipmentHomeSettleSnoozeUntil(item);
  const lastAction = cleanText(item?.homeSettleLastAction);
  const lastActionAt = Math.max(0, toFiniteNumber(item?.homeSettleLastActionAt));
  const quantity = Math.max(0, toFiniteNumber(item?.quantity));
  const measureUnit = MEASURE_UNIT_LABELS[normalizeMeasureUnit(item?.measureUnit, DEFAULT_MEASURE_UNIT)] || cleanText(item?.measureUnit) || "-";
  const unitWeight = getShipmentUnitWeight(item);
  const unitCubic = getShipmentUnitCubic(item);
  const totalWeight = getShipmentTotalWeight(item);
  const totalCubic = getShipmentTotalCubic(item);
  const loadStatus = renderShipmentLoadStatusText(item);
  const paymentStatus = renderShipmentPaymentStatusText(item);
  const destination = renderShipmentDestinationText(item);
  const linkedFinanceDate = linkedFinanceRecord ? parseDateFromAny(linkedFinanceRecord.date) || "-" : "-";
  const linkedFinanceAmount = linkedFinanceRecord ? Math.max(0, toFiniteNumber(linkedFinanceRecord.amount)) : 0;
  const linkageHint = isPaid
    ? linkedFinanceRecord
      ? "已同步到收支管理"
      : "已收款但未检索到联动收入记录"
    : linkedFinanceRecord
    ? "未收款但检索到联动收入记录，请核对"
    : "未收款，暂无联动收入记录";

  const metaItems = [
    ["日期", item?.date || "-"],
    ["序号", item?.serialNo || "-"],
    ["厂家", item?.manufacturer || "-"],
    ["客户", item?.customer || "-"],
    ["产品", item?.product || "-"],
    ["备注", item?.note || "-"],
    ["提醒状态", statusMeta?.label || "-"],
    ["收款状态", paymentStatus],
    ["装车状态", loadStatus],
    ["去向", destination || "-"],
    ["计量值", formatNumber(quantity)],
    ["计量单位", measureUnit],
    ["单价(元)", formatDecimal(Math.max(0, toFiniteNumber(item?.unitPrice)))],
    ["金额(元)", formatDecimal(amount)],
    ["回扣(元)", formatDecimal(rebate)],
    ["欠款金额(元)", formatDecimal(debtAmount)],
    ["单位重量(KG)", formatDecimal(unitWeight)],
    ["单位立方(m3)", formatDecimal(unitCubic)],
    ["总重量(KG)", formatDecimal(totalWeight)],
    ["总立方(m3)", formatDecimal(totalCubic)],
    ["到期日期", dueDate],
    ["逾期天数", `${overdueDays} 天`],
    ["上次提醒时间", formatDateTime(remindedAt)],
    ["稍后提醒至", snoozeUntil > 0 ? formatDateTime(snoozeUntil) : "-"],
    ["最近操作", lastAction || "-"],
    ["最近操作时间", formatDateTime(lastActionAt)],
    ["收款时间", formatDateTime(item?.paidAt)],
    ["联动收入日期", linkedFinanceDate],
    ["联动收入金额(元)", linkedFinanceRecord ? formatDecimal(linkedFinanceAmount) : "-"],
    ["联动检索", linkageHint],
  ];

  statusDetailModalContent.innerHTML = `
    <div class="status-detail-block">
      <h4 class="status-detail-title">货物信息</h4>
      <div class="status-detail-meta-grid">
        ${metaItems
          .map(
            ([label, value]) => `
              <div>
                <strong>${escapeHtml(label)}</strong>
                <span>${escapeHtml(cleanText(value || "-"))}</span>
              </div>
            `
          )
          .join("")}
      </div>
    </div>
  `;
}

function openStatusDetailModal(shipmentId) {
  const safeShipmentId = cleanText(shipmentId);
  if (!safeShipmentId) {
    return;
  }

  const target = shipments.find((item) => cleanText(item?.id) === safeShipmentId);
  if (!target) {
    alert("未找到对应货物，可能已删除。");
    return;
  }

  const row = currentVisibleStatusRows.find((entry) => cleanText(entry?.item?.id) === safeShipmentId) || null;
  currentStatusDetailShipmentId = safeShipmentId;
  renderStatusDetailModal(target, row);
  statusDetailModal?.classList.add("is-open");
}

function closeStatusDetailModal() {
  currentStatusDetailShipmentId = "";
  statusDetailModal?.classList.remove("is-open");
}

function syncStatusDetailView() {
  if (!statusDetailModal?.classList.contains("is-open") || !currentStatusDetailShipmentId) {
    return;
  }

  const target = shipments.find((item) => cleanText(item?.id) === currentStatusDetailShipmentId);
  if (!target) {
    closeStatusDetailModal();
    return;
  }
  const row = currentVisibleStatusRows.find((entry) => cleanText(entry?.item?.id) === currentStatusDetailShipmentId) || null;
  renderStatusDetailModal(target, row);
}

function getStatusStatsSnapshot(now = Date.now()) {
  const stats = {
    total: 0,
    pending: 0,
    active: 0,
    done: 0,
  };

  shipments.forEach((item) => {
    if (!isShipmentHomeSettleCandidate(item) || !isShipmentHomeSettleDue(item, now)) {
      return;
    }
    const statusId = getShipmentHomeSettleStatus(item).id;
    if (statusId === "paid") {
      stats.done += 1;
      return;
    }
    if (statusId === "reminded") {
      stats.active += 1;
    } else {
      stats.pending += 1;
    }
    if (getShipmentHomeSettleSnoozeUntil(item) <= now) {
      stats.total += 1;
    }
  });

  return stats;
}

function renderStatusStats(stats = {}) {
  if (!statusStatTotal || !statusStatPending || !statusStatActive || !statusStatDone) {
    return;
  }

  const total = Math.max(0, toFiniteNumber(stats.total));
  const pending = Math.max(0, toFiniteNumber(stats.pending));
  const active = Math.max(0, toFiniteNumber(stats.active));
  const done = Math.max(0, toFiniteNumber(stats.done));

  statusStatTotal.textContent = formatNumber(total);
  statusStatPending.textContent = formatNumber(pending);
  statusStatActive.textContent = formatNumber(active);
  statusStatDone.textContent = formatNumber(done);
}

function renderStatusActionRecordTable() {
  if (!statusActionRecordTableBody) {
    return;
  }

  const rows = getFilteredStatusActionRecords();
  const paged = paginateRows(rows, statusActionRecordPage, MAIN_PAGE_SIZE);
  statusActionRecordPage = paged.page;
  currentVisibleStatusActionRecordRows = paged.items.slice();
  updateStatusActionRecordPagination(paged.page, paged.totalPages, paged.totalRows);

  if (!rows.length) {
    currentVisibleStatusActionRecordRows = [];
    statusActionRecordTableBody.innerHTML = '<tr><td class="empty" colspan="10">暂无处理记录</td></tr>';
    return;
  }

  statusActionRecordTableBody.innerHTML = paged.items
    .map((item) => {
      const actionType = normalizeHomeSettleActionType(item?.actionType);
      const actionLabel = getHomeSettleActionLabel(actionType);
      const resultText = cleanText(item?.resultText) || "-";
      const debtAfter = Math.max(0, toFiniteNumber(item?.debtAmountAfter, toFiniteNumber(item?.debtAmountBefore)));

      return `
        <tr data-status-action-record-id="${escapeHtml(cleanText(item?.id))}">
          <td>${escapeHtml(formatDateTime(item?.actionAt))}</td>
          <td>${escapeHtml(parseDateFromAny(item?.shipmentDate) || "-")}</td>
          <td class="mono">${escapeHtml(cleanText(item?.shipmentSerialNo) || "-")}</td>
          <td>${escapeHtml(cleanText(item?.manufacturer) || "-")}</td>
          <td>${escapeHtml(cleanText(item?.customer) || "-")}</td>
          <td>${escapeHtml(cleanText(item?.product) || "-")}</td>
          <td><span class="status-node-tag ${escapeHtml(getHomeSettleActionTagClass(actionType))}">${escapeHtml(actionLabel)}</span></td>
          <td>${escapeHtml(resultText)}</td>
          <td class="mono">¥${formatDecimal(debtAfter)}</td>
          <td>${escapeHtml(cleanText(item?.sourceNote) || "-")}</td>
        </tr>
      `;
    })
    .join("");
}

function updateStatusActionRecordPagination(page, totalPages, totalRows) {
  if (!statusActionRecordPrevBtn || !statusActionRecordNextBtn || !statusActionRecordPageInfo) {
    return;
  }

  if (totalRows === 0) {
    statusActionRecordPageInfo.textContent = "暂无数据";
    statusActionRecordPrevBtn.disabled = true;
    statusActionRecordNextBtn.disabled = true;
    return;
  }

  statusActionRecordPageInfo.textContent = `第 ${page} / ${totalPages} 页（共 ${formatNumber(totalRows)} 条）`;
  statusActionRecordPrevBtn.disabled = page <= 1;
  statusActionRecordNextBtn.disabled = page >= totalPages;
}

function getFilteredStatusActionRecords() {
  const filterValue = cleanText(statusActionRecordFilter?.value) || "all";
  const keyword = cleanText(statusActionRecordSearch?.value).toLowerCase();
  const { from, to } = getDateRangeFromInputs(statusActionRecordFromDateInput, statusActionRecordToDateInput);

  return homeSettleActionRecords
    .filter((item) => {
      const actionType = normalizeHomeSettleActionType(item?.actionType);
      if (filterValue !== "all" && actionType !== filterValue) {
        return false;
      }

      const actionDate = parseDateFromAny(item?.actionDate || item?.actionAt);
      if (!isDateInRange(actionDate, from, to)) {
        return false;
      }

      if (!keyword) {
        return true;
      }

      const actionLabel = getHomeSettleActionLabel(actionType);
      const bag = `${item?.shipmentDate || ""} ${item?.shipmentSerialNo || ""} ${item?.manufacturer || ""} ${
        item?.customer || ""
      } ${item?.product || ""} ${actionLabel} ${item?.resultText || ""} ${item?.sourceNote || ""}`.toLowerCase();
      return bag.includes(keyword);
    })
    .sort((a, b) => toFiniteNumber(b?.actionAt) - toFiniteNumber(a?.actionAt));
}

function getFilteredStatusRows() {
  const filterValue = cleanText(statusFilter?.value) || "all";
  const keyword = cleanText(statusSearch?.value).toLowerCase();
  const { from, to } = getDateRangeFromInputs(statusFromDateInput, statusToDateInput);
  const now = Date.now();

  return shipments
    .filter((item) => {
      if (!isShipmentHomeSettleCandidate(item) || !isShipmentHomeSettleDue(item, now)) {
        return false;
      }
      if (getShipmentHomeSettleSnoozeUntil(item) > now) {
        return false;
      }
      const resolved = getShipmentHomeSettleStatus(item);
      return resolved.id !== "paid";
    })
    .map((item) => {
      const resolved = getShipmentHomeSettleStatus(item);
      return {
        item,
        statusId: resolved.id,
        source: resolved.source,
        statusMeta: getHomeSettleStatusMeta(resolved.id),
      };
    })
    .filter((row) => {
      if (!isDateInRange(row.item?.date, from, to)) {
        return false;
      }
      if (filterValue !== "all" && row.statusId !== filterValue) {
        return false;
      }
      if (!keyword) {
        return true;
      }
      const bag = `${row.item?.date || ""} ${row.item?.serialNo || ""} ${row.item?.manufacturer || ""} ${
        row.item?.customer || ""
      } ${row.item?.product || ""} ${row.statusMeta?.label || ""} ${row.item?.note || ""}`.toLowerCase();
      return bag.includes(keyword);
    })
    .sort((a, b) => sortByDateDesc(a.item, b.item));
}

function isShipmentHomeSettleCandidate(item) {
  const note = cleanText(item?.note);
  if (!note) {
    return false;
  }
  const tokenHit = splitShipmentNoteTokens(note).some((token) => normalizeSuggestText(token).includes("家结"));
  if (tokenHit) {
    return true;
  }
  return normalizeSuggestText(note).includes("家结");
}

function getShipmentHomeSettleDueTimestamp(item) {
  const baseDate = parseDateFromAny(item?.date);
  if (!baseDate) {
    return 0;
  }
  const dueDate = new Date(`${baseDate}T00:00:00`);
  dueDate.setMonth(dueDate.getMonth() + HOME_SETTLE_REMINDER_MONTHS);
  return dueDate.getTime();
}

function isShipmentHomeSettleDue(item, now = Date.now()) {
  const dueTimestamp = getShipmentHomeSettleDueTimestamp(item);
  return dueTimestamp > 0 && now >= dueTimestamp;
}

function getShipmentHomeSettleSnoozeUntil(item) {
  return Math.max(0, toFiniteNumber(item?.homeSettleSnoozeUntil));
}

function getShipmentHomeSettleStatus(item) {
  if (normalizeShipmentPaidFlag(item?.isPaid)) {
    return {
      id: "paid",
      source: "auto",
    };
  }

  const manualStatusId = normalizeHomeSettleStatusId(item?.homeSettleStatus);
  if (manualStatusId && manualStatusId !== "paid") {
    return {
      id: manualStatusId,
      source: "manual",
    };
  }
  if (Math.max(0, toFiniteNumber(item?.homeSettleRemindedAt)) > 0) {
    return {
      id: "reminded",
      source: "auto",
    };
  }
  return {
    id: "unpaid",
    source: "auto",
  };
}

function normalizeHomeSettleActionType(value) {
  const text = cleanText(value).toLowerCase();
  if (!text) {
    return "contact";
  }
  if (text === "snooze" || text.includes("稍后")) {
    return "snooze";
  }
  if (text === "paid" || text.includes("已收")) {
    return "paid";
  }
  if (text === "contact" || text.includes("联系")) {
    return "contact";
  }
  return "contact";
}

function getHomeSettleActionLabel(actionType) {
  const normalized = normalizeHomeSettleActionType(actionType);
  return HOME_SETTLE_ACTION_LABELS[normalized] || HOME_SETTLE_ACTION_LABELS.contact;
}

function getHomeSettleActionTagClass(actionType) {
  const normalized = normalizeHomeSettleActionType(actionType);
  if (normalized === "paid") {
    return "is-done";
  }
  if (normalized === "snooze") {
    return "is-pending";
  }
  return "is-active";
}

function buildHomeSettleActionResultText(actionType) {
  const normalized = normalizeHomeSettleActionType(actionType);
  if (normalized === "snooze") {
    return "已延后1小时，超时后自动再次提醒";
  }
  if (normalized === "paid") {
    return "已收款，已同步联动收支并停止提醒";
  }
  return "已标记为马上联系，后续可继续跟进";
}

function sanitizeHomeSettleActionRecords(list) {
  if (!Array.isArray(list) || list.length === 0) {
    return [];
  }

  return list
    .map((item, idx) => {
      const actionType = normalizeHomeSettleActionType(item?.actionType);
      const actionAt = Math.max(0, toFiniteNumber(item?.actionAt || item?.createdAt || item?.timestamp));
      const shipmentDate = parseDateFromAny(item?.shipmentDate || item?.date || item?.shipment_time) || "";
      const actionDate = parseDateFromAny(item?.actionDate || actionAt) || "";
      const debtAmountBefore = Math.max(0, toFiniteNumber(item?.debtAmountBefore));
      const debtAmountAfter = Math.max(
        0,
        toFiniteNumber(item?.debtAmountAfter, toFiniteNumber(item?.debtAmountBefore))
      );

      return {
        id: cleanText(item?.id) || `status_action_${Date.now()}_${idx}`,
        shipmentId: cleanText(item?.shipmentId),
        shipmentDate,
        shipmentSerialNo: cleanText(item?.shipmentSerialNo || item?.serialNo),
        manufacturer: cleanText(item?.manufacturer),
        customer: cleanText(item?.customer),
        product: cleanText(item?.product),
        actionType,
        actionLabel: getHomeSettleActionLabel(actionType),
        resultText: cleanText(item?.resultText) || buildHomeSettleActionResultText(actionType),
        debtAmountBefore,
        debtAmountAfter,
        sourceNote: cleanText(item?.sourceNote || item?.note),
        actionAt: actionAt || Date.now() + idx,
        actionDate,
      };
    })
    .sort((a, b) => toFiniteNumber(b?.actionAt) - toFiniteNumber(a?.actionAt));
}

function appendHomeSettleActionRecord(shipment, actionType, options = {}) {
  if (!shipment) {
    return;
  }
  const normalized = normalizeHomeSettleActionType(actionType);
  const now = Math.max(1, toFiniteNumber(options.actionAt, Date.now()));
  const debtBefore = Math.max(
    0,
    toFiniteNumber(options.debtAmountBefore, Math.max(0, toFiniteNumber(shipment?.amount) - Math.max(0, toFiniteNumber(shipment?.rebate))))
  );
  const debtAfter = Math.max(0, toFiniteNumber(options.debtAmountAfter, normalized === "paid" ? 0 : debtBefore));

  const record = {
    id: makeId(),
    shipmentId: cleanText(shipment?.id),
    shipmentDate: parseDateFromAny(shipment?.date) || "",
    shipmentSerialNo: cleanText(shipment?.serialNo),
    manufacturer: cleanText(shipment?.manufacturer),
    customer: cleanText(shipment?.customer),
    product: cleanText(shipment?.product),
    actionType: normalized,
    actionLabel: getHomeSettleActionLabel(normalized),
    resultText: cleanText(options.resultText) || buildHomeSettleActionResultText(normalized),
    debtAmountBefore: debtBefore,
    debtAmountAfter: debtAfter,
    sourceNote: cleanText(shipment?.note),
    actionAt: now,
    actionDate: parseDateFromAny(now) || dateValue(new Date()),
  };

  homeSettleActionRecords.unshift(record);
  writeStorage(STORAGE_KEYS.homeSettleActionRecords, homeSettleActionRecords);
}

function getHomeSettleStatusMeta(statusId) {
  return HOME_SETTLE_STATUS_BY_ID.get(statusId) || HOME_SETTLE_STATUS_BY_ID.get("unpaid");
}

function normalizeHomeSettleStatusId(value) {
  const raw = cleanText(value);
  const text = raw.toLowerCase();
  if (!text) {
    return "";
  }
  if (HOME_SETTLE_STATUS_IDS.has(text)) {
    return text;
  }
  if (raw.includes("待收款")) {
    return "unpaid";
  }
  if (raw.includes("已提醒")) {
    return "reminded";
  }
  if (raw.includes("已收款")) {
    return "paid";
  }
  return "";
}

function clearShipmentHomeSettleStatusOverride(item) {
  if (!item) {
    return;
  }
  item.homeSettleStatus = "";
  item.homeSettleStatusUpdatedAt = 0;
}

function setShipmentHomeSettleStatusOverride(item, statusId) {
  if (!item) {
    return;
  }
  const normalized = normalizeHomeSettleStatusId(statusId);
  if (!normalized) {
    clearShipmentHomeSettleStatusOverride(item);
    return;
  }
  item.homeSettleStatus = normalized;
  item.homeSettleStatusUpdatedAt = Date.now();
}

function renderDailySummarySection(options = {}) {
  if (!summaryShipmentBody || !summaryFinanceBody) {
    return;
  }

  if (options.resetPages) {
    summaryShipmentPage = 1;
    summaryFinancePage = 1;
  }
  syncSummaryKeywordInputs();
  normalizeDateInputs(summaryFromDateInput, summaryToDateInput);

  const { from, to } = getSummaryRange();
  const rangedShipments = filterSummaryShipments(
    shipments.filter((item) => isDateInRange(item.date, from, to)),
    summaryShipmentKeyword
  );
  const rangedFinance = filterSummaryFinance(
    financeRecords.filter((item) => isDateInRange(item.date, from, to)),
    summaryFinanceKeyword
  );

  renderSummaryShipmentStats(rangedShipments);
  renderSummaryFinanceStats(rangedFinance);
  renderSummaryShipmentTable(rangedShipments);
  renderSummaryFinanceTable(rangedFinance);
}

function renderSummaryShipmentStats(rangedShipments) {
  let receiveQty = 0;
  let shipQty = 0;
  let receiveWeight = 0;
  let shipWeight = 0;
  let receiveCubic = 0;
  let shipCubic = 0;
  let totalAmount = 0;
  let totalRebate = 0;

  rangedShipments.forEach((item) => {
    const quantity = toFiniteNumber(item.quantity);
    const amount = toFiniteNumber(item.amount);
    const rebate = toFiniteNumber(item.rebate);
    const totalWeight = getShipmentTotalWeight(item);
    const totalCubic = getShipmentTotalCubic(item);
    const isReceive = item.type === "receive";
    const isShip = item.type === "ship";

    totalAmount += amount;
    totalRebate += rebate;

    if (isReceive) {
      receiveQty += quantity;
      receiveWeight += totalWeight;
      receiveCubic += totalCubic;
      return;
    }

    if (isShip) {
      shipQty += quantity;
      shipWeight += totalWeight;
      shipCubic += totalCubic;
    }
  });

  const netAmount = totalAmount - totalRebate;
  const days = new Set(rangedShipments.map((item) => String(item.date || "")).filter(Boolean)).size;

  if (!summaryShipmentStatsEls.days) {
    return;
  }

  summaryShipmentStatsEls.days.textContent = formatNumber(days);
  summaryShipmentStatsEls.receiveQty.textContent = formatNumber(receiveQty);
  summaryShipmentStatsEls.receiveWeight.textContent = `总重量 ${formatDecimal(receiveWeight)} KG / 总立方 ${formatDecimal(receiveCubic)} m3`;
  summaryShipmentStatsEls.shipQty.textContent = formatNumber(shipQty);
  summaryShipmentStatsEls.shipWeight.textContent = `总重量 ${formatDecimal(shipWeight)} KG / 总立方 ${formatDecimal(shipCubic)} m3`;
  summaryShipmentStatsEls.netQty.textContent = formatCurrency(totalAmount);
  summaryShipmentStatsEls.netWeight.textContent = `回扣 ${formatCurrency(totalRebate)} / 净额 ${formatCurrency(netAmount)}`;
}

function renderSummaryFinanceStats(rangedFinance) {
  let income = 0;
  let expense = 0;
  const daySet = new Set();

  rangedFinance.forEach((item) => {
    const amount = toFiniteNumber(item.amount);
    if (item.type === "income") {
      income += amount;
    } else if (item.type === "expense") {
      expense += amount;
    }

    const day = String(item.date || "");
    if (day) {
      daySet.add(day);
    }
  });
  const days = daySet.size;

  if (!summaryFinanceStatsEls.days) {
    return;
  }

  summaryFinanceStatsEls.days.textContent = formatNumber(days);
  summaryFinanceStatsEls.income.textContent = formatCurrency(income);
  summaryFinanceStatsEls.expense.textContent = formatCurrency(expense);
  summaryFinanceStatsEls.balance.textContent = formatCurrency(income - expense);
  summaryFinanceStatsEls.records.textContent = `记录 ${formatNumber(rangedFinance.length)}`;
}

function filterSummaryShipments(rows, keyword) {
  const normalizedKeyword = cleanText(keyword).toLowerCase();
  if (!normalizedKeyword) {
    return rows;
  }

  return rows.filter((item) => {
    const bag =
      `${item.date} ${TYPE_LABELS[item.type] || ""} ${item.serialNo} ${item.manufacturer} ${item.customer} ${item.product} ${formatMeasureUnit(item.measureUnit)} ${item.transit} ${item.delivery} ${item.pickup} ${item.note} ${renderShipmentLoadStatusText(item)} ${renderShipmentPaymentStatusText(item)}`.toLowerCase();
    return bag.includes(normalizedKeyword);
  });
}

function filterSummaryFinance(rows, keyword) {
  const normalizedKeyword = cleanText(keyword).toLowerCase();
  if (!normalizedKeyword) {
    return rows;
  }

  return rows.filter((item) => {
    const bag = `${item.date} ${TYPE_LABELS[item.type] || ""} ${item.summary || item.category} ${item.note}`.toLowerCase();
    return bag.includes(normalizedKeyword);
  });
}

function buildSummaryShipmentQuantityText(rows) {
  const unitTotals = {
    piece: 0,
    ton: 0,
    cubic: 0,
    other: 0,
  };

  rows.forEach((item) => {
    const quantity = toFiniteNumber(item.quantity);
    if (quantity <= 0) {
      return;
    }
    const unit = normalizeMeasureUnit(item.measureUnit, DEFAULT_MEASURE_UNIT) || DEFAULT_MEASURE_UNIT;
    if (unit === "piece" || unit === "ton" || unit === "cubic") {
      unitTotals[unit] += quantity;
      return;
    }
    unitTotals.other += quantity;
  });

  const parts = [];
  if (unitTotals.piece > 0) {
    parts.push(`件:${formatMeasureNumber(unitTotals.piece)}`);
  }
  if (unitTotals.ton > 0) {
    parts.push(`KG:${formatMeasureNumber(unitTotals.ton)}`);
  }
  if (unitTotals.cubic > 0) {
    parts.push(`平方:${formatMeasureNumber(unitTotals.cubic)}`);
  }
  if (unitTotals.other > 0) {
    parts.push(`其他:${formatMeasureNumber(unitTotals.other)}`);
  }

  return parts.length ? parts.join(" / ") : "-";
}

function renderSummaryShipmentTable(rangedShipments) {
  const rows = rangedShipments.slice().sort(sortByDateDesc);

  const paged = paginateRows(rows, summaryShipmentPage, SUMMARY_PAGE_SIZE);
  summaryShipmentPage = paged.page;
  updateSummaryPagination("shipment", paged.page, paged.totalPages, paged.totalRows);

  if (rows.length === 0) {
    summaryShipmentBody.innerHTML =
      '<tr><td class="empty" colspan="16">当前区间暂无收发记录</td></tr>';
    return;
  }

  const totalAmount = rows.reduce((sum, item) => sum + toFiniteNumber(item.amount), 0);
  const totalRebate = rows.reduce((sum, item) => sum + toFiniteNumber(item.rebate), 0);
  const totalWeight = rows.reduce((sum, item) => sum + getShipmentTotalWeight(item), 0);
  const totalCubic = rows.reduce((sum, item) => sum + getShipmentTotalCubic(item), 0);
  const loadedCount = rows.filter((item) => normalizeShipmentLoadedFlag(item.isLoaded)).length;
  const pendingCount = rows.filter((item) => item.type === "receive" && !normalizeShipmentLoadedFlag(item.isLoaded)).length;
  const paidCount = rows.filter((item) => normalizeShipmentPaidFlag(item.isPaid)).length;
  const unpaidCount = rows.length - paidCount;
  const quantityText = buildSummaryShipmentQuantityText(rows);

  summaryShipmentBody.innerHTML = paged.items
    .map(
      (item) => `
      <tr>
        <td>${escapeHtml(item.date)}</td>
        <td><span class="tag ${escapeHtml(item.type)}">${TYPE_LABELS[item.type]}</span></td>
        <td class="mono">${escapeHtml(item.serialNo || "-")}</td>
        <td>${escapeHtml(item.manufacturer || "-")}</td>
        <td>${escapeHtml(item.customer || "-")}</td>
        <td>${escapeHtml(item.product || "-")}</td>
        <td>${formatMeasureValue(item.quantity, item.measureUnit)}</td>
        <td>${formatMeasureUnit(item.measureUnit)}</td>
        <td>${formatCurrency(toFiniteNumber(item.amount))}</td>
        <td>${formatDecimal(getShipmentUnitWeight(item))}</td>
        <td>${formatDecimal(getShipmentUnitCubic(item))}</td>
        <td>${formatDecimal(getShipmentTotalWeight(item))}</td>
        <td>${formatDecimal(getShipmentTotalCubic(item))}</td>
        <td>${escapeHtml(item.note || "-")}</td>
        <td>${renderShipmentLoadStatus(item, { interactive: false })}</td>
        <td>${renderShipmentPaymentStatusTag(item)}</td>
      </tr>
    `
    )
    .join("");

  summaryShipmentBody.innerHTML += `
    <tr class="summary-total-row">
      <td>合计</td>
      <td>-</td>
      <td class="mono">记录 ${formatNumber(rows.length)} 条</td>
      <td>-</td>
      <td>-</td>
      <td>-</td>
      <td>${escapeHtml(quantityText)}</td>
      <td>混合</td>
      <td>${formatCurrency(totalAmount)}</td>
      <td>-</td>
      <td>-</td>
      <td>${formatDecimal(totalWeight)}</td>
      <td>${formatDecimal(totalCubic)}</td>
      <td>回扣 ${formatCurrency(totalRebate)}</td>
      <td>已装 ${formatNumber(loadedCount)} / 待配 ${formatNumber(pendingCount)}</td>
      <td>已收 ${formatNumber(paidCount)} / 未收 ${formatNumber(unpaidCount)}</td>
    </tr>
  `;
}

function renderSummaryFinanceTable(rangedFinance) {
  const rows = rangedFinance.slice().sort(sortByDateDesc);
  const balanceById = buildFinanceBalanceMap(rows);

  const paged = paginateRows(rows, summaryFinancePage, SUMMARY_PAGE_SIZE);
  summaryFinancePage = paged.page;
  updateSummaryPagination("finance", paged.page, paged.totalPages, paged.totalRows);

  if (rows.length === 0) {
    summaryFinanceBody.innerHTML =
      '<tr><td class="empty" colspan="5">当前区间暂无收支记录</td></tr>';
    return;
  }

  const totalIncome = rows
    .filter((item) => item.type === "income")
    .reduce((sum, item) => sum + toFiniteNumber(item.amount), 0);
  const totalExpense = rows
    .filter((item) => item.type === "expense")
    .reduce((sum, item) => sum + toFiniteNumber(item.amount), 0);
  const totalBalance = totalIncome - totalExpense;

  summaryFinanceBody.innerHTML = paged.items
    .map(
      (item) => `
      <tr>
        <td>${escapeHtml(item.date)}</td>
        <td class="finance-ledger-summary">${escapeHtml(resolveFinanceSummary(item))}</td>
        <td class="finance-ledger-income">${item.type === "income" ? formatCurrency(toFiniteNumber(item.amount)) : "-"}</td>
        <td class="finance-ledger-expense">${item.type === "expense" ? formatCurrency(toFiniteNumber(item.amount)) : "-"}</td>
        <td class="finance-ledger-balance">${formatCurrency(toFiniteNumber(balanceById.get(item.id)))}</td>
      </tr>
    `
    )
    .join("");

  summaryFinanceBody.innerHTML += `
    <tr class="summary-total-row">
      <td>合计</td>
      <td class="finance-ledger-summary">记录 ${formatNumber(rows.length)} 条</td>
      <td class="finance-ledger-income">${formatCurrency(totalIncome)}</td>
      <td class="finance-ledger-expense">${formatCurrency(totalExpense)}</td>
      <td class="finance-ledger-balance">${formatCurrency(totalBalance)}</td>
    </tr>
  `;
}

function changeSummaryPage(table, offset) {
  if (table === "shipment") {
    summaryShipmentPage = Math.max(1, summaryShipmentPage + offset);
  }
  if (table === "finance") {
    summaryFinancePage = Math.max(1, summaryFinancePage + offset);
  }
  renderDailySummarySection();
}

function updateSummaryPagination(table, page, totalPages, totalRows) {
  const isShipment = table === "shipment";
  const prevBtn = isShipment ? summaryShipmentPrevBtn : summaryFinancePrevBtn;
  const nextBtn = isShipment ? summaryShipmentNextBtn : summaryFinanceNextBtn;
  const pageInfo = isShipment ? summaryShipmentPageInfo : summaryFinancePageInfo;

  if (!prevBtn || !nextBtn || !pageInfo) {
    return;
  }

  if (totalRows === 0) {
    pageInfo.textContent = "暂无数据";
    prevBtn.disabled = true;
    nextBtn.disabled = true;
    return;
  }

  pageInfo.textContent = `第 ${page} / ${totalPages} 页（共 ${formatNumber(totalRows)} 条）`;
  prevBtn.disabled = page <= 1;
  nextBtn.disabled = page >= totalPages;
}

function paginateRows(rows, page, pageSize) {
  const totalRows = rows.length;
  const totalPages = Math.max(1, Math.ceil(totalRows / pageSize));
  const safePage = Math.min(Math.max(1, page), totalPages);
  const start = (safePage - 1) * pageSize;

  return {
    items: rows.slice(start, start + pageSize),
    page: safePage,
    totalPages,
    totalRows,
  };
}

function getDateRangeFromInputs(fromInput, toInput) {
  const from = cleanText(fromInput?.value);
  const to = cleanText(toInput?.value);

  if (from && to && from > to) {
    return { from: to, to: from };
  }

  return { from, to };
}

function normalizeDateInputs(fromInput, toInput) {
  if (!fromInput || !toInput) {
    return;
  }

  const from = cleanText(fromInput.value);
  const to = cleanText(toInput.value);
  if (from && to && from > to) {
    fromInput.value = to;
    toInput.value = from;
  }
}

function isInputLike(target) {
  if (!(target instanceof HTMLElement)) {
    return false;
  }

  if (target.isContentEditable) {
    return true;
  }

  const tagName = target.tagName.toLowerCase();
  return tagName === "input" || tagName === "textarea" || tagName === "select";
}

function getSummaryRange() {
  if (!summaryFromDateInput || !summaryToDateInput) {
    return { from: "", to: "" };
  }

  const from = cleanText(summaryFromDateInput.value);
  const to = cleanText(summaryToDateInput.value);
  if (from && to && from > to) {
    return { from: to, to: from };
  }
  return { from, to };
}

function isDateInRange(date, from, to) {
  const normalized = String(date || "");
  if (!normalized) {
    return false;
  }
  if (from && normalized < from) {
    return false;
  }
  if (to && normalized > to) {
    return false;
  }
  return true;
}

function renderShipmentStats() {
  const receive = sumShipment("receive");
  const ship = sumShipment("ship");
  const totalAmount = shipments.reduce((sum, item) => sum + toFiniteNumber(item.amount), 0);
  const totalRebate = shipments.reduce((sum, item) => sum + toFiniteNumber(item.rebate), 0);
  const netAmount = totalAmount - totalRebate;

  shipmentStatsEls.receiveQty.textContent = formatNumber(receive.quantity);
  shipmentStatsEls.receiveWeight.textContent = `总重量 ${formatDecimal(receive.weight)} KG / 总立方 ${formatDecimal(receive.cubic)} m3`;
  shipmentStatsEls.shipQty.textContent = formatNumber(ship.quantity);
  shipmentStatsEls.shipWeight.textContent = `总重量 ${formatDecimal(ship.weight)} KG / 总立方 ${formatDecimal(ship.cubic)} m3`;
  if (shipmentStatsEls.totalFee) {
    shipmentStatsEls.totalFee.textContent = formatCurrency(totalAmount);
  }
  if (shipmentStatsEls.totalDeclared) {
    shipmentStatsEls.totalDeclared.textContent = `回扣合计 ${formatCurrency(totalRebate)}`;
  }
  shipmentStatsEls.netQty.textContent = formatCurrency(netAmount);
  shipmentStatsEls.netWeight.textContent = `记录 ${formatNumber(shipments.length)} 条`;
}

function renderFinanceStats() {
  const month = dateValue(new Date()).slice(0, 7);
  let totalIncome = 0;
  let totalExpense = 0;
  let monthIncome = 0;
  let monthExpense = 0;

  financeRecords.forEach((item) => {
    const amount = toFiniteNumber(item.amount);
    const isCurrentMonth = String(item.date || "").startsWith(month);

    if (item.type === "income") {
      totalIncome += amount;
      if (isCurrentMonth) {
        monthIncome += amount;
      }
      return;
    }

    if (item.type === "expense") {
      totalExpense += amount;
      if (isCurrentMonth) {
        monthExpense += amount;
      }
    }
  });

  const totalBalance = totalIncome - totalExpense;
  const monthBalance = monthIncome - monthExpense;

  financeStatsEls.todayIncome.textContent = formatCurrency(totalIncome);
  financeStatsEls.todayExpense.textContent = formatCurrency(totalExpense);
  financeStatsEls.monthBalance.textContent = formatCurrency(totalBalance);
  financeStatsEls.monthBreakdown.textContent = `本月收入 ${formatCurrency(monthIncome)} / 本月支出 ${formatCurrency(monthExpense)} / 本月结余 ${formatCurrency(monthBalance)} / 记录 ${formatNumber(financeRecords.length)} 条`;
}

function normalizeShipmentEditedFieldsMap(raw) {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) {
    return {};
  }
  const next = {};
  Object.entries(raw).forEach(([key, value]) => {
    const normalizedKey = cleanText(key);
    if (!normalizedKey || value !== true) {
      return;
    }
    next[normalizedKey] = true;
  });
  return next;
}

function isShipmentFieldEdited(item, fieldKey) {
  if (!item || !fieldKey) {
    return false;
  }
  const map = normalizeShipmentEditedFieldsMap(item.editedFields);
  return Boolean(map[fieldKey]);
}

function setShipmentFieldEdited(item, fieldKey, edited = true) {
  if (!item || !fieldKey) {
    return;
  }
  const map = normalizeShipmentEditedFieldsMap(item.editedFields);
  if (edited) {
    map[fieldKey] = true;
  } else {
    delete map[fieldKey];
  }
  item.editedFields = map;
}

function buildShipmentCellClassAttr(item, fieldKey, baseClass = "") {
  const classes = [];
  if (baseClass) {
    classes.push(baseClass);
  }
  if (isShipmentFieldEdited(item, fieldKey)) {
    classes.push("shipment-cell-edited");
  }
  if (!classes.length) {
    return "";
  }
  return ` class="${classes.join(" ")}"`;
}

function buildShipmentEditKeyAttr(fieldKey) {
  if (!SHIPMENT_INLINE_EDIT_FIELDS.has(fieldKey)) {
    return "";
  }
  return ` data-edit-key="${fieldKey}" title="双击编辑"`;
}

function renderShipmentTable(options = {}) {
  if (options.resetPage) {
    shipmentPage = 1;
  }

  const rows = sortShipmentRows(getFilteredShipments());
  pruneSelectedShipmentExportIds(rows);
  const paged = paginateRows(rows, shipmentPage, MAIN_PAGE_SIZE);
  shipmentPage = paged.page;
  currentVisibleShipmentRows = paged.items.slice();
  updateMainPagination("shipment", paged.page, paged.totalPages, paged.totalRows);

  if (rows.length === 0) {
    currentVisibleShipmentRows = [];
    syncShipmentExportSelectAllState();
    shipmentTableBody.innerHTML =
      '<tr><td class="empty" colspan="24">暂无符合条件的收发记录</td></tr>';
    return;
  }

  shipmentTableBody.innerHTML = paged.items
    .map((item) => {
      const isLoaded = normalizeShipmentLoadedFlag(item.isLoaded);
      const isSelectedForExport = selectedShipmentIdsForExport.has(item.id);
      const rowClassName = [isLoaded ? "shipment-loaded-row" : "", isSelectedForExport ? "shipment-selected-row" : ""]
        .filter(Boolean)
        .join(" ");
      return `
      <tr class="${rowClassName}" data-shipment-id="${escapeHtml(cleanText(item.id))}">
        <td class="shipment-load-select-cell">
          <input
            type="checkbox"
            class="shipment-load-checkbox shipment-export-checkbox"
            data-id="${escapeHtml(item.id)}"
            ${isSelectedForExport ? "checked" : ""}
            aria-label="勾选用于导出"
          />
        </td>
        <td${buildShipmentEditKeyAttr("date")}${buildShipmentCellClassAttr(item, "date")}>${escapeHtml(item.date)}</td>
        <td${buildShipmentEditKeyAttr("serialNo")}${buildShipmentCellClassAttr(item, "serialNo", "mono")}>${escapeHtml(
        item.serialNo || "-"
      )}</td>
        <td><span class="tag ${escapeHtml(item.type)}">${TYPE_LABELS[item.type]}</span></td>
        <td${buildShipmentEditKeyAttr("manufacturer")}${buildShipmentCellClassAttr(item, "manufacturer")}>${escapeHtml(
        item.manufacturer || "-"
      )}</td>
        <td${buildShipmentEditKeyAttr("customer")}${buildShipmentCellClassAttr(item, "customer")}>${escapeHtml(
        item.customer || "-"
      )}</td>
        <td${buildShipmentEditKeyAttr("product")}${buildShipmentCellClassAttr(item, "product")}>${escapeHtml(
        item.product || "-"
      )}</td>
        <td${buildShipmentEditKeyAttr("quantity")}${buildShipmentCellClassAttr(item, "quantity")}>${formatMeasureValue(
        item.quantity,
        item.measureUnit
      )}</td>
        <td${buildShipmentEditKeyAttr("measureUnit")}${buildShipmentCellClassAttr(item, "measureUnit")}>${formatMeasureUnit(
        item.measureUnit
      )}</td>
        <td${buildShipmentEditKeyAttr("unitPrice")}${buildShipmentCellClassAttr(item, "unitPrice")}>${toFiniteNumber(
        item.unitPrice
      ).toFixed(2)}</td>
        <td${buildShipmentEditKeyAttr("amount")}${buildShipmentCellClassAttr(item, "amount")}>${formatCurrency(
        toFiniteNumber(item.amount)
      )}</td>
        <td${buildShipmentEditKeyAttr("weight")}${buildShipmentCellClassAttr(item, "weight")}>${formatDecimal(
        getShipmentUnitWeight(item)
      )}</td>
        <td${buildShipmentEditKeyAttr("cubic")}${buildShipmentCellClassAttr(item, "cubic")}>${formatDecimal(
        getShipmentUnitCubic(item)
      )}</td>
        <td>${formatDecimal(getShipmentTotalWeight(item))}</td>
        <td>${formatDecimal(getShipmentTotalCubic(item))}</td>
        <td${buildShipmentEditKeyAttr("transit")}${buildShipmentCellClassAttr(item, "transit")}>${escapeHtml(
        item.transit || "-"
      )}</td>
        <td${buildShipmentEditKeyAttr("delivery")}${buildShipmentCellClassAttr(item, "delivery")}>${escapeHtml(
        item.delivery || "-"
      )}</td>
        <td${buildShipmentEditKeyAttr("pickup")}${buildShipmentCellClassAttr(item, "pickup")}>${escapeHtml(
        item.pickup || "-"
      )}</td>
        <td${buildShipmentEditKeyAttr("rebate")}${buildShipmentCellClassAttr(item, "rebate")}>${toFiniteNumber(item.rebate).toFixed(
        2
      )}</td>
        <td${buildShipmentEditKeyAttr("note")}${buildShipmentCellClassAttr(item, "note")}>${escapeHtml(item.note || "-")}</td>
        <td>${renderShipmentLoadStatus(item)}</td>
        <td>${renderShipmentDestination(item)}</td>
        <td>${renderShipmentPaymentToggle(item)}</td>
        <td><button class="delete-btn" data-table="shipment" data-id="${item.id}">删除</button></td>
      </tr>
    `;
    })
    .join("");

  syncShipmentExportSelectAllState();
}

function pruneSelectedShipmentExportIds(filteredRows = []) {
  const validIds = new Set((Array.isArray(filteredRows) ? filteredRows : []).map((item) => cleanText(item?.id)));
  selectedShipmentIdsForExport.forEach((id) => {
    if (!validIds.has(id)) {
      selectedShipmentIdsForExport.delete(id);
    }
  });
}

function syncShipmentExportSelectAllState() {
  if (!shipmentExportSelectAllCheckbox) {
    return;
  }
  const checkboxes = Array.from(shipmentTableBody.querySelectorAll(".shipment-export-checkbox"));
  if (!checkboxes.length) {
    shipmentExportSelectAllCheckbox.checked = false;
    shipmentExportSelectAllCheckbox.indeterminate = false;
    return;
  }
  const checkedCount = checkboxes.filter((checkbox) => checkbox.checked).length;
  shipmentExportSelectAllCheckbox.checked = checkedCount > 0 && checkedCount === checkboxes.length;
  shipmentExportSelectAllCheckbox.indeterminate = checkedCount > 0 && checkedCount < checkboxes.length;
}

function pruneSelectedFinanceExportIds(filteredRows = []) {
  const validIds = new Set((Array.isArray(filteredRows) ? filteredRows : []).map((item) => cleanText(item?.id)));
  selectedFinanceIdsForExport.forEach((id) => {
    if (!validIds.has(id)) {
      selectedFinanceIdsForExport.delete(id);
    }
  });
}

function syncFinanceExportSelectAllState() {
  if (!financeExportSelectAllCheckbox) {
    return;
  }
  const checkboxes = Array.from(financeTableBody.querySelectorAll(".finance-export-checkbox"));
  if (!checkboxes.length) {
    financeExportSelectAllCheckbox.checked = false;
    financeExportSelectAllCheckbox.indeterminate = false;
    return;
  }
  const checkedCount = checkboxes.filter((checkbox) => checkbox.checked).length;
  financeExportSelectAllCheckbox.checked = checkedCount > 0 && checkedCount === checkboxes.length;
  financeExportSelectAllCheckbox.indeterminate = checkedCount > 0 && checkedCount < checkboxes.length;
}

function handleFinanceExportSelectionChange(event) {
  const checkbox = event.target.closest(".finance-export-checkbox");
  if (!checkbox) {
    return;
  }
  const id = cleanText(checkbox.dataset.id);
  if (!id) {
    return;
  }
  if (checkbox.checked) {
    selectedFinanceIdsForExport.add(id);
  } else {
    selectedFinanceIdsForExport.delete(id);
  }
  const row = checkbox.closest("tr");
  if (row) {
    row.classList.toggle("finance-selected-row", checkbox.checked);
  }
  syncFinanceExportSelectAllState();
}

function handleFinanceExportSelectAllChange(event) {
  const shouldSelect = Boolean(event.target?.checked);
  const visibleRows = Array.isArray(currentVisibleFinanceRows) ? currentVisibleFinanceRows : [];
  visibleRows.forEach((item) => {
    const id = cleanText(item?.id);
    if (!id) {
      return;
    }
    if (shouldSelect) {
      selectedFinanceIdsForExport.add(id);
    } else {
      selectedFinanceIdsForExport.delete(id);
    }
  });
  renderFinanceTable();
}

function renderShipmentLoadStatus(item, options = {}) {
  const interactive = options.interactive !== false;
  const isLoaded = normalizeShipmentLoadedFlag(item.isLoaded);
  const canToggle = item.type === "receive" || (item.type === "ship" && isLoaded);
  if (!canToggle) {
    return '<span class="shipment-load-status is-other">非入库</span>';
  }
  const statusClass = isLoaded ? "is-loaded" : "is-pending";
  const statusText = isLoaded ? "已装车" : "待配车";
  if (!interactive) {
    return `<span class="shipment-load-status ${statusClass}">${statusText}</span>`;
  }
  return `<button type="button" class="shipment-load-toggle ${statusClass}" data-id="${escapeHtml(
    item.id
  )}" aria-label="切换装车状态">${statusText}</button>`;
}

function renderShipmentLoadStatusText(item) {
  const isLoaded = normalizeShipmentLoadedFlag(item.isLoaded);
  const canToggle = item.type === "receive" || (item.type === "ship" && isLoaded);
  if (!canToggle) {
    return "非入库";
  }
  return isLoaded ? "已装车" : "待配车";
}

function buildDispatchDestinationText(meta = {}) {
  const dispatchDate = parseDateFromAny(meta.dispatchDate) || cleanText(meta.dispatchDate);
  const dispatchTruckNo = cleanText(meta.dispatchTruckNo || meta.truckNo);
  const dispatchPlateNo = cleanText(meta.dispatchPlateNo || meta.plateNo).toUpperCase();
  const dispatchDriver = cleanText(meta.dispatchDriver || meta.driver);
  const dispatchContactName = cleanText(meta.dispatchContactName || meta.contactName || meta.name);
  const segments = [];
  if (dispatchDate) {
    segments.push(dispatchDate);
  }
  if (dispatchTruckNo) {
    segments.push(dispatchTruckNo);
  }
  if (dispatchPlateNo) {
    segments.push(dispatchPlateNo);
  }
  if (dispatchDriver) {
    segments.push(`司机姓名:${dispatchDriver}`);
  }
  if (dispatchContactName) {
    segments.push(`联系方式:${dispatchContactName}`);
  }
  return segments.join(" / ");
}

function normalizeDispatchMeta(meta = {}) {
  const dispatchDate = parseDateFromAny(meta.dispatchDate) || cleanText(meta.dispatchDate);
  const dispatchTruckNo = cleanText(meta.dispatchTruckNo || meta.truckNo);
  const dispatchPlateNo = cleanText(meta.dispatchPlateNo || meta.plateNo).toUpperCase();
  const dispatchDriver = cleanText(meta.dispatchDriver || meta.driver);
  const dispatchContactName = cleanText(meta.dispatchContactName || meta.contactName || meta.name);
  return {
    dispatchDate,
    dispatchTruckNo,
    dispatchPlateNo,
    dispatchDriver,
    dispatchContactName,
    dispatchDestination: buildDispatchDestinationText({
      dispatchDate,
      dispatchTruckNo,
      dispatchPlateNo,
      dispatchDriver,
      dispatchContactName,
    }),
  };
}

function buildDispatchProductsSummary(items = [], maxCount = 5) {
  const list = Array.isArray(items) ? items : [];
  const uniqueProducts = [];

  list.forEach((item) => {
    const product = cleanText(item?.product);
    if (!product || uniqueProducts.includes(product)) {
      return;
    }
    uniqueProducts.push(product);
  });

  if (uniqueProducts.length === 0) {
    return "-";
  }

  if (uniqueProducts.length <= maxCount) {
    return uniqueProducts.join("、");
  }

  return `${uniqueProducts.slice(0, maxCount).join("、")} 等${uniqueProducts.length}种`;
}

function getDispatchProductsSummary(record = {}) {
  const direct = cleanText(record.productsSummary || record.products);
  if (direct) {
    return direct;
  }
  return buildDispatchProductsSummary(record.items);
}

function getDispatchMetaFromRecord(record = {}) {
  const items = Array.isArray(record.items) ? record.items : [];
  const firstItem = items[0] || {};
  const fallbackDispatchDate = parseDateFromAny(record.loadedAt || record.createdAt) || "";
  const baseMeta = normalizeDispatchMeta({
    dispatchDate: record.dispatchDate || firstItem.dispatchDate || fallbackDispatchDate,
    dispatchTruckNo: record.dispatchTruckNo || firstItem.dispatchTruckNo,
    dispatchPlateNo: record.dispatchPlateNo || firstItem.dispatchPlateNo,
    dispatchDriver: record.dispatchDriver || firstItem.dispatchDriver,
    dispatchContactName: record.dispatchContactName || firstItem.dispatchContactName,
  });
  const directDestination = cleanText(record.dispatchDestination);
  return {
    ...baseMeta,
    dispatchDestination: directDestination || baseMeta.dispatchDestination,
  };
}

function getDispatchMetaFromInputs() {
  const fallbackDate = dateValue(new Date());
  const dispatchDate = parseDateFromAny(dispatchMetaDateInput?.value) || fallbackDate;
  const dispatchTruckNo = cleanText(dispatchMetaTruckNoInput?.value);
  const dispatchPlateNo = cleanText(dispatchMetaPlateNoInput?.value).toUpperCase();
  const dispatchDriver = cleanText(dispatchMetaDriverInput?.value);
  const dispatchContactName = cleanText(dispatchMetaContactNameInput?.value);
  const missing = [];

  if (!dispatchDate) {
    missing.push("日期");
  }
  if (!dispatchTruckNo) {
    missing.push("几号车");
  }
  if (!dispatchPlateNo) {
    missing.push("车牌号");
  }
  if (!dispatchDriver) {
    missing.push("司机姓名");
  }
  if (!dispatchContactName) {
    missing.push("联系方式");
  }

  if (missing.length) {
    alert(`请先填写配车信息：${missing.join("、")}。`);
    return null;
  }

  if (dispatchMetaDateInput) {
    dispatchMetaDateInput.value = dispatchDate;
  }
  if (dispatchMetaPlateNoInput) {
    dispatchMetaPlateNoInput.value = dispatchPlateNo;
  }

  return normalizeDispatchMeta({
    dispatchDate,
    dispatchTruckNo,
    dispatchPlateNo,
    dispatchDriver,
    dispatchContactName,
  });
}

function getShipmentDispatchDestination(item) {
  const isLoaded = normalizeShipmentLoadedFlag(item?.isLoaded);
  const canShowDestination = item?.type === "receive" || (item?.type === "ship" && isLoaded);
  if (!canShowDestination || !isLoaded) {
    return "";
  }
  const direct = cleanText(item.dispatchDestination || item.destination);
  if (direct) {
    return direct;
  }
  return buildDispatchDestinationText(item);
}

function renderShipmentDestination(item) {
  const destination = getShipmentDispatchDestination(item);
  if (!destination) {
    return "-";
  }
  return `<span class="dispatch-destination-text">${escapeHtml(destination)}</span>`;
}

function renderShipmentDestinationText(item) {
  return getShipmentDispatchDestination(item) || "-";
}

function renderShipmentPaymentStatusText(item) {
  return normalizeShipmentPaidFlag(item?.isPaid) ? "已收款" : "未收款";
}

function renderShipmentPaymentStatusTag(item) {
  const isPaid = normalizeShipmentPaidFlag(item?.isPaid);
  return `<span class="shipment-payment-status ${isPaid ? "is-paid" : "is-unpaid"}">${renderShipmentPaymentStatusText(item)}</span>`;
}

function renderShipmentPaymentToggle(item) {
  const isPaid = normalizeShipmentPaidFlag(item?.isPaid);
  return `<button type="button" class="shipment-payment-toggle ${isPaid ? "is-paid" : "is-unpaid"}" data-id="${escapeHtml(
    item.id
  )}" aria-label="切换收款状态">${renderShipmentPaymentStatusText(item)}</button>`;
}

function buildShipmentPaymentFinanceSummary(item) {
  const serialNo = cleanText(item?.serialNo);
  const manufacturer = cleanText(item?.manufacturer);
  const customer = cleanText(item?.customer);
  const product = cleanText(item?.product);
  const segments = [serialNo, manufacturer, customer, product].filter(Boolean);
  if (!segments.length) {
    return SHIPMENT_PAYMENT_FINANCE_SUMMARY_PREFIX;
  }
  return `${SHIPMENT_PAYMENT_FINANCE_SUMMARY_PREFIX}：${segments.join(" / ")}`;
}

function buildShipmentPaymentFinanceNote(item) {
  const shipmentNote = cleanText(item?.note);
  return shipmentNote;
}

function stripShipmentPaymentFinanceHintText(noteText) {
  let text = cleanText(noteText);
  if (!text) {
    return "";
  }

  if (text.includes(SHIPMENT_PAYMENT_FINANCE_NOTE_PREFIX_LEGACY)) {
    text = text.replace(SHIPMENT_PAYMENT_FINANCE_NOTE_PREFIX_LEGACY, "");
  }
  if (text.includes(SHIPMENT_PAYMENT_FINANCE_NOTE_PREFIX)) {
    text = text.replace(SHIPMENT_PAYMENT_FINANCE_NOTE_PREFIX, "");
  }

  text = text.replace(/^[：:；;，,\s]+/u, "");
  text = text.replace(/^原备注[：:]?/u, "");
  return cleanText(text);
}

function isShipmentPaymentAutoFinanceRecord(record) {
  const sourceType = cleanText(record?.sourceType);
  if (sourceType === SHIPMENT_PAYMENT_FINANCE_SOURCE) {
    return true;
  }

  const summary = cleanText(record?.summary || record?.category);
  if (
    summary.startsWith(SHIPMENT_PAYMENT_FINANCE_SUMMARY_PREFIX) ||
    summary.startsWith(SHIPMENT_PAYMENT_FINANCE_SUMMARY_PREFIX_LEGACY)
  ) {
    return true;
  }

  const note = cleanText(record?.note);
  if (
    note.includes(SHIPMENT_PAYMENT_FINANCE_NOTE_PREFIX) ||
    note.includes(SHIPMENT_PAYMENT_FINANCE_NOTE_PREFIX_LEGACY)
  ) {
    return true;
  }

  const sourceShipmentId = cleanText(record?.sourceShipmentId);
  const sourceSerialNo = cleanText(record?.sourceSerialNo);
  if ((sourceShipmentId || sourceSerialNo) && normalizeFinanceType(record?.type) === "income") {
    return true;
  }

  return false;
}

function resolveShipmentIdForPaymentFinanceRecord(record, shipmentIdBySerialNo = new Map()) {
  const directShipmentId = cleanText(record?.sourceShipmentId);
  if (directShipmentId) {
    return directShipmentId;
  }

  const sourceSerialNo = cleanText(record?.sourceSerialNo);
  if (!sourceSerialNo) {
    return "";
  }

  return cleanText(shipmentIdBySerialNo.get(sourceSerialNo));
}

function buildShipmentPaymentFinanceRecord(item, previousRecord = null) {
  const shipmentId = cleanText(item?.id);
  const amount = Math.max(0, toFiniteNumber(item?.amount));
  const paidAt = Math.max(0, toFiniteNumber(item?.paidAt));
  const date = parseDateFromAny(paidAt) || parseDateFromAny(item?.date) || dateValue(new Date());
  const summary = buildShipmentPaymentFinanceSummary(item);
  return {
    id: cleanText(previousRecord?.id) || makeId(),
    type: "income",
    date,
    category: summary,
    summary,
    amount,
    note: buildShipmentPaymentFinanceNote(item),
    createdAt: Math.max(0, toFiniteNumber(previousRecord?.createdAt)) || paidAt || Date.now(),
    sourceType: SHIPMENT_PAYMENT_FINANCE_SOURCE,
    sourceShipmentId: shipmentId,
    sourceSerialNo: cleanText(item?.serialNo),
  };
}

function isSameShipmentPaymentFinanceRecord(prev, next) {
  return (
    cleanText(prev?.type) === cleanText(next?.type) &&
    cleanText(prev?.date) === cleanText(next?.date) &&
    cleanText(prev?.category) === cleanText(next?.category) &&
    cleanText(prev?.summary) === cleanText(next?.summary) &&
    Math.max(0, toFiniteNumber(prev?.amount)) === Math.max(0, toFiniteNumber(next?.amount)) &&
    cleanText(prev?.note) === cleanText(next?.note) &&
    cleanText(prev?.sourceType) === cleanText(next?.sourceType) &&
    cleanText(prev?.sourceShipmentId) === cleanText(next?.sourceShipmentId) &&
    cleanText(prev?.sourceSerialNo) === cleanText(next?.sourceSerialNo)
  );
}

function syncFinanceIncomeFromPaidShipments() {
  const shipmentIdBySerialNo = new Map();
  shipments.forEach((item) => {
    const shipmentId = cleanText(item?.id);
    const serialNo = cleanText(item?.serialNo);
    if (!shipmentId || !serialNo || shipmentIdBySerialNo.has(serialNo)) {
      return;
    }
    shipmentIdBySerialNo.set(serialNo, shipmentId);
  });

  const manualRecords = [];
  const autoRecordsByShipmentId = new Map();
  let changed = false;

  financeRecords.forEach((record) => {
    if (!isShipmentPaymentAutoFinanceRecord(record)) {
      manualRecords.push(record);
      return;
    }
    const shipmentId = resolveShipmentIdForPaymentFinanceRecord(record, shipmentIdBySerialNo);
    if (!shipmentId) {
      changed = true;
      return;
    }
    if (autoRecordsByShipmentId.has(shipmentId)) {
      changed = true;
      return;
    }
    const normalizedRecord = {
      ...record,
      sourceType: SHIPMENT_PAYMENT_FINANCE_SOURCE,
      sourceShipmentId: shipmentId,
      sourceSerialNo: cleanText(record?.sourceSerialNo),
    };
    if (!isSameShipmentPaymentFinanceRecord(record, normalizedRecord)) {
      changed = true;
    }
    autoRecordsByShipmentId.set(shipmentId, normalizedRecord);
  });

  const nextFinanceRecords = manualRecords.slice();
  const activePaidShipmentIds = new Set();

  shipments.forEach((item) => {
    if (!normalizeShipmentPaidFlag(item.isPaid)) {
      return;
    }
    const shipmentId = cleanText(item.id);
    if (!shipmentId) {
      return;
    }
    const amount = Math.max(0, toFiniteNumber(item.amount));
    if (amount <= 0) {
      return;
    }

    activePaidShipmentIds.add(shipmentId);
    const previousRecord = autoRecordsByShipmentId.get(shipmentId) || null;
    const nextRecord = buildShipmentPaymentFinanceRecord(item, previousRecord);
    if (!previousRecord || !isSameShipmentPaymentFinanceRecord(previousRecord, nextRecord)) {
      changed = true;
    }
    nextFinanceRecords.push(nextRecord);
  });

  autoRecordsByShipmentId.forEach((_record, shipmentId) => {
    if (!activePaidShipmentIds.has(shipmentId)) {
      changed = true;
    }
  });

  if (!changed) {
    return false;
  }

  financeRecords = nextFinanceRecords;
  writeStorage(STORAGE_KEYS.finance, financeRecords);
  return true;
}

function escapeSelectorValue(value) {
  const text = cleanText(value);
  if (!text) {
    return "";
  }
  if (typeof CSS !== "undefined" && typeof CSS.escape === "function") {
    return CSS.escape(text);
  }
  return text.replace(/["\\]/g, "\\$&");
}

function findShipmentEditCell(shipmentId, fieldKey) {
  const safeShipmentId = escapeSelectorValue(shipmentId);
  const safeFieldKey = escapeSelectorValue(fieldKey);
  if (!safeShipmentId || !safeFieldKey || !shipmentTableBody) {
    return null;
  }
  return shipmentTableBody.querySelector(
    `tr[data-shipment-id="${safeShipmentId}"] td[data-edit-key="${safeFieldKey}"]`
  );
}

function getShipmentInlineEditRawValue(item, fieldKey) {
  if (!item || !fieldKey) {
    return "";
  }
  if (fieldKey === "date") {
    return parseDateFromAny(item.date) || "";
  }
  if (fieldKey === "measureUnit") {
    return normalizeMeasureUnit(item.measureUnit, DEFAULT_MEASURE_UNIT) || DEFAULT_MEASURE_UNIT;
  }
  if (fieldKey in SHIPMENT_INLINE_EDIT_NUMERIC_CONFIG) {
    const config = SHIPMENT_INLINE_EDIT_NUMERIC_CONFIG[fieldKey];
    const digits = Number.isFinite(config?.digits) ? config.digits : 3;
    return toFiniteNumber(item[fieldKey]).toFixed(digits);
  }
  return cleanText(item[fieldKey]);
}

function normalizeShipmentInlineEditValue(fieldKey, rawValue) {
  const text = cleanText(rawValue);

  if (fieldKey === "date") {
    const normalizedDate = parseDateFromAny(text);
    if (!normalizedDate) {
      return { ok: false, error: "日期格式不正确，请使用 年/月/日。" };
    }
    return { ok: true, value: normalizedDate };
  }

  if (fieldKey === "measureUnit") {
    const unit = normalizeMeasureUnit(text, "");
    if (!unit) {
      return { ok: false, error: "计量单位仅支持：件 / KG / m3(平方)。" };
    }
    return { ok: true, value: unit };
  }

  if (fieldKey in SHIPMENT_INLINE_EDIT_NUMERIC_CONFIG) {
    const parsed = parseNumericValue(text);
    if (!Number.isFinite(parsed) || parsed < 0) {
      return { ok: false, error: "请输入大于等于 0 的数字。" };
    }
    const config = SHIPMENT_INLINE_EDIT_NUMERIC_CONFIG[fieldKey];
    const digits = Number.isFinite(config?.digits) ? config.digits : 3;
    return { ok: true, value: Number(parsed.toFixed(digits)) };
  }

  return { ok: true, value: cleanText(text) };
}

function areShipmentInlineValuesEqual(fieldKey, prevValue, nextValue) {
  if (fieldKey in SHIPMENT_INLINE_EDIT_NUMERIC_CONFIG) {
    const config = SHIPMENT_INLINE_EDIT_NUMERIC_CONFIG[fieldKey];
    const digits = Number.isFinite(config?.digits) ? config.digits : 3;
    return Number(toFiniteNumber(prevValue).toFixed(digits)) === Number(toFiniteNumber(nextValue).toFixed(digits));
  }
  return cleanText(prevValue) === cleanText(nextValue);
}

function buildShipmentInlineEditor(fieldKey, item) {
  if (fieldKey === "measureUnit") {
    const select = document.createElement("select");
    select.className = "shipment-inline-editor shipment-inline-select";
    Object.entries(MEASURE_UNIT_LABELS).forEach(([value, label]) => {
      const option = document.createElement("option");
      option.value = value;
      option.textContent = label;
      select.appendChild(option);
    });
    select.value = getShipmentInlineEditRawValue(item, fieldKey);
    return select;
  }

  const input = document.createElement("input");
  input.className = "shipment-inline-editor";
  input.autocomplete = "off";

  if (fieldKey === "date") {
    input.type = "date";
  } else if (fieldKey in SHIPMENT_INLINE_EDIT_NUMERIC_CONFIG) {
    const config = SHIPMENT_INLINE_EDIT_NUMERIC_CONFIG[fieldKey];
    input.type = "number";
    input.step = config.step || "0.001";
    input.min = config.min || "0";
    input.inputMode = "decimal";
  } else {
    input.type = "text";
  }

  input.value = getShipmentInlineEditRawValue(item, fieldKey);
  return input;
}

function reportShipmentInlineEditorError(session, message) {
  if (!session?.editor) {
    return;
  }
  const editor = session.editor;
  editor.classList.add("is-invalid");
  if (typeof editor.setCustomValidity === "function") {
    editor.setCustomValidity(message || "输入有误");
    if (typeof editor.reportValidity === "function") {
      editor.reportValidity();
    }
    editor.setCustomValidity("");
  }
}

function clearShipmentInlineEditorError(session) {
  if (!session?.editor) {
    return;
  }
  session.editor.classList.remove("is-invalid");
  if (typeof session.editor.setCustomValidity === "function") {
    session.editor.setCustomValidity("");
  }
}

function syncDispatchRecordsForShipment(shipment) {
  const shipmentId = cleanText(shipment?.id);
  if (!shipmentId) {
    return false;
  }

  let changed = false;
  dispatchRecords = dispatchRecords
    .map((record) => {
      const items = Array.isArray(record?.items) ? record.items : [];
      if (!items.length) {
        return record;
      }

      let rowChanged = false;
      const recordMeta = getDispatchMetaFromRecord(record);
      const nextItems = items.map((item) => {
        const currentShipmentId = cleanText(item?.shipmentId || item?.id);
        if (currentShipmentId !== shipmentId) {
          return item;
        }
        rowChanged = true;
        return buildDispatchRecordItem(shipment, recordMeta);
      });

      if (!rowChanged) {
        return record;
      }

      changed = true;
      return rebuildDispatchRecordWithItems(record, nextItems);
    })
    .filter(Boolean);

  return changed;
}

function restoreShipmentInlineCell(session) {
  if (!session?.cell || !session.cell.isConnected) {
    return;
  }
  session.cell.innerHTML = session.originalHtml;
}

function clearShipmentInlineEditSession(options = {}) {
  const { restore = false } = options;
  if (!activeShipmentInlineEditSession) {
    return null;
  }
  const session = activeShipmentInlineEditSession;
  activeShipmentInlineEditSession = null;
  if (restore) {
    restoreShipmentInlineCell(session);
  }
  return session;
}

function startShipmentInlineEdit(cell, shipmentId, fieldKey) {
  const safeShipmentId = cleanText(shipmentId);
  if (!cell || !safeShipmentId || !SHIPMENT_INLINE_EDIT_FIELDS.has(fieldKey)) {
    return;
  }

  if (activeShipmentInlineEditSession) {
    const current = activeShipmentInlineEditSession;
    if (current.shipmentId === safeShipmentId && current.fieldKey === fieldKey) {
      return;
    }
    commitShipmentInlineEdit({ reason: "switch", skipFocusBack: true });
    const refreshedCell = findShipmentEditCell(safeShipmentId, fieldKey);
    if (!refreshedCell) {
      return;
    }
    startShipmentInlineEdit(refreshedCell, safeShipmentId, fieldKey);
    return;
  }

  const target = shipments.find((item) => cleanText(item?.id) === safeShipmentId);
  if (!target) {
    return;
  }

  const editor = buildShipmentInlineEditor(fieldKey, target);
  const originalHtml = cell.innerHTML;
  cell.innerHTML = "";
  cell.appendChild(editor);

  const session = {
    shipmentId: safeShipmentId,
    fieldKey,
    cell,
    editor,
    originalHtml,
  };
  activeShipmentInlineEditSession = session;

  const handleKeydown = (event) => {
    if (event.key === "Escape") {
      event.preventDefault();
      cancelShipmentInlineEdit();
      return;
    }
    if (event.key === "Enter" && !(event.shiftKey && editor.tagName === "TEXTAREA")) {
      event.preventDefault();
      commitShipmentInlineEdit();
    }
  };

  const handleBlur = () => {
    commitShipmentInlineEdit({ reason: "blur", skipFocusBack: true });
  };

  editor.addEventListener("keydown", handleKeydown);
  editor.addEventListener("blur", handleBlur);
  if (editor.tagName === "SELECT") {
    editor.addEventListener("change", () => commitShipmentInlineEdit({ reason: "change", skipFocusBack: true }));
  }

  requestAnimationFrame(() => {
    if (!editor.isConnected) {
      return;
    }
    editor.focus();
    if (typeof editor.select === "function" && editor.tagName === "INPUT") {
      editor.select();
    }
  });
}

function cancelShipmentInlineEdit() {
  const session = clearShipmentInlineEditSession({ restore: true });
  if (!session) {
    return false;
  }
  return true;
}

function commitShipmentInlineEdit(options = {}) {
  const { reason = "manual", skipFocusBack = false } = options;
  const session = activeShipmentInlineEditSession;
  if (!session) {
    return false;
  }

  const target = shipments.find((item) => cleanText(item?.id) === session.shipmentId);
  if (!target) {
    clearShipmentInlineEditSession({ restore: true });
    return false;
  }

  clearShipmentInlineEditorError(session);
  const parsed = normalizeShipmentInlineEditValue(session.fieldKey, session.editor.value);
  if (!parsed.ok) {
    reportShipmentInlineEditorError(session, parsed.error);
    if (!skipFocusBack) {
      requestAnimationFrame(() => {
        if (session.editor?.isConnected) {
          session.editor.focus();
        }
      });
    }
    return false;
  }

  const prevValue = target[session.fieldKey];
  const nextValue = parsed.value;
  const fieldChanged = !areShipmentInlineValuesEqual(session.fieldKey, prevValue, nextValue);

  if (!fieldChanged) {
    clearShipmentInlineEditSession({ restore: true });
    return true;
  }

  target[session.fieldKey] = nextValue;
  setShipmentFieldEdited(target, session.fieldKey, true);

  let amountAutoUpdated = false;
  if (
    (session.fieldKey === "quantity" || session.fieldKey === "unitPrice") &&
    !isShipmentFieldEdited(target, "amount")
  ) {
    const nextAmount = Number((Math.max(0, toFiniteNumber(target.quantity)) * Math.max(0, toFiniteNumber(target.unitPrice))).toFixed(2));
    if (!areShipmentInlineValuesEqual("amount", target.amount, nextAmount)) {
      target.amount = nextAmount;
      setShipmentFieldEdited(target, "amount", true);
      amountAutoUpdated = true;
    }
  }

  writeStorage(STORAGE_KEYS.shipments, shipments);
  const dispatchChanged = syncDispatchRecordsForShipment(target);
  if (dispatchChanged) {
    writeStorage(STORAGE_KEYS.dispatchRecords, dispatchRecords);
  }
  const financeChanged = syncFinanceIncomeFromPaidShipments();

  clearShipmentInlineEditSession({ restore: false });
  renderShipmentSection();
  if (financeChanged || session.fieldKey === "amount" || session.fieldKey === "date" || session.fieldKey === "serialNo" || amountAutoUpdated) {
    renderFinanceSection();
  }

  if (!skipFocusBack && reason !== "blur") {
    const nextCell = findShipmentEditCell(session.shipmentId, session.fieldKey);
    if (nextCell) {
      nextCell.focus();
    }
  }
  return true;
}

function handleShipmentTableDblClick(event) {
  if (event.target.closest(".shipment-inline-editor")) {
    return;
  }
  const cell = event.target.closest("td[data-edit-key]");
  if (!cell) {
    return;
  }
  const row = cell.closest("tr[data-shipment-id]");
  if (!row) {
    return;
  }
  const shipmentId = cleanText(row.dataset.shipmentId);
  const fieldKey = cleanText(cell.dataset.editKey);
  if (!shipmentId || !fieldKey || !SHIPMENT_INLINE_EDIT_FIELDS.has(fieldKey)) {
    return;
  }
  event.preventDefault();
  startShipmentInlineEdit(cell, shipmentId, fieldKey);
}

async function handleShipmentTableClick(event) {
  if (await handleShipmentLoadToggleClick(event, "shipment")) {
    return;
  }
  if (handleShipmentPaymentToggleClick(event)) {
    return;
  }
  await handleDeleteClick(event);
}

function handleShipmentExportSelectionChange(event) {
  const checkbox = event.target.closest(".shipment-export-checkbox");
  if (!checkbox) {
    return;
  }
  const id = cleanText(checkbox.dataset.id);
  if (!id) {
    return;
  }
  if (checkbox.checked) {
    selectedShipmentIdsForExport.add(id);
  } else {
    selectedShipmentIdsForExport.delete(id);
  }
  const row = checkbox.closest("tr");
  if (row) {
    row.classList.toggle("shipment-selected-row", checkbox.checked);
  }
  syncShipmentExportSelectAllState();
}

function handleShipmentExportSelectAllChange(event) {
  const shouldSelect = Boolean(event.target?.checked);
  const visibleRows = Array.isArray(currentVisibleShipmentRows) ? currentVisibleShipmentRows : [];
  visibleRows.forEach((item) => {
    const id = cleanText(item?.id);
    if (!id) {
      return;
    }
    if (shouldSelect) {
      selectedShipmentIdsForExport.add(id);
    } else {
      selectedShipmentIdsForExport.delete(id);
    }
  });
  renderShipmentTable();
}

async function handleDispatchTableClick(event) {
  if (dispatchLoadStage === "meta" && event.target.closest(".shipment-load-toggle")) {
    return;
  }
  await handleShipmentLoadToggleClick(event, "dispatch");
}

function setShipmentPaidState(target, nextPaid, options = {}) {
  if (!target) {
    return {
      changed: false,
      financeChanged: false,
    };
  }
  const normalizedNext = Boolean(nextPaid);
  const prevPaid = normalizeShipmentPaidFlag(target.isPaid);
  if (prevPaid === normalizedNext) {
    return {
      changed: false,
      financeChanged: false,
    };
  }

  target.isPaid = normalizedNext;
  target.paidAt = normalizedNext ? Date.now() : 0;
  writeStorage(STORAGE_KEYS.shipments, shipments);
  const financeChanged = syncFinanceIncomeFromPaidShipments();
  if (!options.skipRender) {
    renderShipmentSection();
    if (financeChanged) {
      renderFinanceSection();
    }
  }

  return {
    changed: true,
    financeChanged,
  };
}

function handleShipmentPaymentToggleClick(event) {
  const button = event.target.closest(".shipment-payment-toggle");
  if (!button) {
    return false;
  }

  const id = cleanText(button.dataset.id);
  if (!id) {
    return true;
  }

  const target = shipments.find((item) => item.id === id);
  if (!target) {
    return true;
  }

  const nextPaid = !normalizeShipmentPaidFlag(target.isPaid);
  setShipmentPaidState(target, nextPaid);
  return true;
}

function getDispatchMetaDraftFromInputs() {
  return normalizeDispatchMeta({
    dispatchDate: parseDateFromAny(dispatchMetaDateInput?.value) || "",
    dispatchTruckNo: cleanText(dispatchMetaTruckNoInput?.value),
    dispatchPlateNo: cleanText(dispatchMetaPlateNoInput?.value).toUpperCase(),
    dispatchDriver: cleanText(dispatchMetaDriverInput?.value),
    dispatchContactName: cleanText(dispatchMetaContactNameInput?.value),
  });
}

function getDispatchMetaForManualToggle(item) {
  const itemMeta = normalizeDispatchMeta({
    dispatchDate: item.dispatchDate,
    dispatchTruckNo: item.dispatchTruckNo,
    dispatchPlateNo: item.dispatchPlateNo,
    dispatchDriver: item.dispatchDriver,
    dispatchContactName: item.dispatchContactName,
  });
  const draftMeta = getDispatchMetaDraftFromInputs();
  return normalizeDispatchMeta({
    dispatchDate: draftMeta.dispatchDate || itemMeta.dispatchDate || dateValue(new Date()),
    dispatchTruckNo: draftMeta.dispatchTruckNo || itemMeta.dispatchTruckNo,
    dispatchPlateNo: draftMeta.dispatchPlateNo || itemMeta.dispatchPlateNo,
    dispatchDriver: draftMeta.dispatchDriver || itemMeta.dispatchDriver,
    dispatchContactName: draftMeta.dispatchContactName || itemMeta.dispatchContactName,
  });
}

function rebuildDispatchRecordWithItems(record, items) {
  if (!Array.isArray(items) || items.length === 0) {
    return null;
  }
  const totals = computeDispatchRecordTotals(items);
  const dispatchMeta = getDispatchMetaFromRecord({
    ...record,
    items,
  });
  return {
    ...record,
    itemCount: items.length,
    totalAmount: totals.totalAmount,
    totalWeight: totals.totalWeight,
    totalCubic: totals.totalCubic,
    totalPieces: totals.totalPieces,
    dispatchDate: dispatchMeta.dispatchDate,
    dispatchTruckNo: dispatchMeta.dispatchTruckNo,
    dispatchPlateNo: dispatchMeta.dispatchPlateNo,
    dispatchDriver: dispatchMeta.dispatchDriver,
    dispatchContactName: dispatchMeta.dispatchContactName,
    dispatchDestination: dispatchMeta.dispatchDestination,
    productsSummary: buildDispatchProductsSummary(items),
    items,
  };
}

function removeShipmentFromDispatchRecords(shipmentId) {
  const safeShipmentId = cleanText(shipmentId);
  if (!safeShipmentId) {
    return;
  }
  dispatchRecords = dispatchRecords
    .map((record) => {
      const items = (Array.isArray(record.items) ? record.items : []).filter(
        (item) => cleanText(item.shipmentId || item.id) !== safeShipmentId
      );
      return rebuildDispatchRecordWithItems(record, items);
    })
    .filter(Boolean);
}

async function handleShipmentLoadToggleClick(event, source = "shipment") {
  const button = event.target.closest(".shipment-load-toggle");
  if (!button) {
    return false;
  }
  const id = cleanText(button.dataset.id);
  if (!id) {
    return true;
  }
  const target = shipments.find((item) => item.id === id);
  const currentlyLoaded = target ? normalizeShipmentLoadedFlag(target.isLoaded) : false;
  const isToggleable = Boolean(target) && (target.type === "receive" || (target.type === "ship" && currentlyLoaded));
  if (!isToggleable) {
    return true;
  }
  const nextLoaded = !currentlyLoaded;

  if (!nextLoaded) {
    const confirmed = await confirmDialog("确认将此货物改为“待配车”吗？去向与配车关联会一并清空。", {
      title: "确认变更装车状态",
    });
    if (!confirmed) {
      return true;
    }
  }

  if (nextLoaded) {
    const dispatchMeta = getDispatchMetaForManualToggle(target);
    const loadedAt = Date.now();
    target.type = "ship";
    target.isLoaded = true;
    target.loadedAt = loadedAt;
    target.dispatchDate = dispatchMeta.dispatchDate;
    target.dispatchTruckNo = dispatchMeta.dispatchTruckNo;
    target.dispatchPlateNo = dispatchMeta.dispatchPlateNo;
    target.dispatchDriver = dispatchMeta.dispatchDriver;
    target.dispatchContactName = dispatchMeta.dispatchContactName;
    target.dispatchDestination = dispatchMeta.dispatchDestination;
    target.destination = dispatchMeta.dispatchDestination;

    const record = buildDispatchRecord([target], loadedAt, dispatchMeta);
    if (record) {
      dispatchRecords.push(record);
    }
    selectedDispatchIdsForLoading.delete(target.id);
    confirmedDispatchIdsForLoading.delete(target.id);
  } else {
    target.type = "receive";
    target.isLoaded = false;
    target.loadedAt = 0;
    target.dispatchDate = "";
    target.dispatchTruckNo = "";
    target.dispatchPlateNo = "";
    target.dispatchDriver = "";
    target.dispatchContactName = "";
    target.dispatchDestination = "";
    target.destination = "";
    removeShipmentFromDispatchRecords(target.id);
    selectedDispatchIdsForLoading.delete(target.id);
    confirmedDispatchIdsForLoading.delete(target.id);
  }

  writeStorage(STORAGE_KEYS.shipments, shipments);
  writeStorage(STORAGE_KEYS.dispatchRecords, dispatchRecords);
  renderShipmentSection();
  if (source === "dispatch" && dispatchSelectAllLoad) {
    dispatchSelectAllLoad.checked = false;
  }
  return true;
}

function applyBatchLoadForIds(selectedIds, dispatchMeta = {}) {
  const selectedSet = new Set(selectedIds);
  const loadedAt = Date.now();
  const loadedRows = [];
  const normalizedMeta = normalizeDispatchMeta(dispatchMeta);

  shipments.forEach((item) => {
    if (!selectedSet.has(item.id)) {
      return;
    }
    if (item.type !== "receive" || normalizeShipmentLoadedFlag(item.isLoaded)) {
      return;
    }
    item.type = "ship";
    item.isLoaded = true;
    item.loadedAt = loadedAt;
    item.dispatchDate = normalizedMeta.dispatchDate;
    item.dispatchTruckNo = normalizedMeta.dispatchTruckNo;
    item.dispatchPlateNo = normalizedMeta.dispatchPlateNo;
    item.dispatchDriver = normalizedMeta.dispatchDriver;
    item.dispatchContactName = normalizedMeta.dispatchContactName;
    item.dispatchDestination = normalizedMeta.dispatchDestination;
    item.destination = normalizedMeta.dispatchDestination;
    loadedRows.push(item);
  });

  if (!loadedRows.length) {
    return { updatedCount: 0, record: null };
  }

  const record = buildDispatchRecord(loadedRows, loadedAt, normalizedMeta);
  if (record) {
    dispatchRecords.push(record);
    writeStorage(STORAGE_KEYS.dispatchRecords, dispatchRecords);
  }

  writeStorage(STORAGE_KEYS.shipments, shipments);
  return { updatedCount: loadedRows.length, record };
}

function renderDispatchTable(options = {}) {
  if (!dispatchTableBody) {
    return;
  }

  if (options.resetPage) {
    dispatchPage = 1;
  }

  const rows = getFilteredDispatchRows().sort(sortByDateDesc);
  const paged = paginateRows(rows, dispatchPage, MAIN_PAGE_SIZE);
  dispatchPage = paged.page;
  currentVisibleDispatchRows = paged.items.slice();
  updateMainPagination("dispatch", paged.page, paged.totalPages, paged.totalRows);

  if (rows.length === 0) {
    currentVisibleDispatchRows = [];
    dispatchTableBody.innerHTML = '<tr><td class="empty" colspan="16">暂无符合条件的入库记录</td></tr>';
    updateDispatchLoadToolbar([]);
    return;
  }

  dispatchTableBody.innerHTML = paged.items
    .map((item) => {
      const isLoaded = normalizeShipmentLoadedFlag(item.isLoaded);
      const selectedSet = dispatchLoadStage === "meta" ? confirmedDispatchIdsForLoading : selectedDispatchIdsForLoading;
      const isSelected = !isLoaded && selectedSet.has(item.id);
      const rowClassName = [isLoaded ? "dispatch-loaded-row" : "", isSelected ? "dispatch-selected-row" : ""]
        .filter(Boolean)
        .join(" ");
      const checkboxAttrs = dispatchLoadStage === "meta" ? "disabled" : "";

      return `
      <tr class="${rowClassName}">
        <td>${escapeHtml(item.date)}</td>
        <td class="mono">${escapeHtml(item.serialNo || "-")}</td>
        <td>${escapeHtml(item.manufacturer || "-")}</td>
        <td>${escapeHtml(item.customer || "-")}</td>
        <td>${escapeHtml(item.product || "-")}</td>
        <td>${formatMeasureValue(item.quantity, item.measureUnit)}</td>
        <td>${formatMeasureUnit(item.measureUnit)}</td>
        <td>${toFiniteNumber(item.unitPrice).toFixed(2)}</td>
        <td>${formatCurrency(toFiniteNumber(item.amount))}</td>
        <td>${formatDecimal(getShipmentUnitWeight(item))}</td>
        <td>${formatDecimal(getShipmentUnitCubic(item))}</td>
        <td>${formatDecimal(getShipmentTotalWeight(item))}</td>
        <td>${formatDecimal(getShipmentTotalCubic(item))}</td>
        <td>${escapeHtml(item.note || "-")}</td>
        <td class="shipment-load-select-cell">
          ${
            isLoaded
              ? '<span class="shipment-load-placeholder">-</span>'
              : `<input type="checkbox" class="dispatch-load-checkbox" data-id="${escapeHtml(item.id)}" ${checkboxAttrs} ${
                  isSelected ? "checked" : ""
                } />`
          }
        </td>
        <td>${renderDispatchLoadStatus(item)}</td>
      </tr>
      `;
    })
    .join("");

  updateDispatchLoadToolbar(paged.items);
}

function pruneSelectedDispatchLoadIds() {
  const validIds = new Set(
    shipments
      .filter((item) => item.type === "receive" && !normalizeShipmentLoadedFlag(item.isLoaded))
      .map((item) => item.id)
  );

  if (dispatchLoadStage === "meta") {
    confirmedDispatchIdsForLoading.forEach((id) => {
      if (!validIds.has(id)) {
        confirmedDispatchIdsForLoading.delete(id);
      }
    });

    selectedDispatchIdsForLoading = new Set(confirmedDispatchIdsForLoading);
    if (!confirmedDispatchIdsForLoading.size) {
      dispatchLoadStage = "select";
      closeDispatchMetaModal();
    }
    return;
  }

  if (!selectedDispatchIdsForLoading.size) {
    return;
  }

  selectedDispatchIdsForLoading.forEach((id) => {
    if (!validIds.has(id)) {
      selectedDispatchIdsForLoading.delete(id);
    }
  });
}

function updateDispatchMetaInputsState(disabled) {
  const inputs = [
    dispatchMetaDateInput,
    dispatchMetaTruckNoInput,
    dispatchMetaPlateNoInput,
    dispatchMetaDriverInput,
    dispatchMetaContactNameInput,
  ];

  inputs.forEach((input) => {
    if (input) {
      input.disabled = disabled;
    }
  });

  if (!disabled && dispatchMetaDateInput && !dispatchMetaDateInput.value) {
    dispatchMetaDateInput.value = dateValue(new Date());
  }
}

function resetDispatchLoadFlow(options = {}) {
  const { clearSelected = false } = options;
  dispatchLoadStage = "select";
  confirmedDispatchIdsForLoading.clear();
  if (clearSelected) {
    selectedDispatchIdsForLoading.clear();
  }
}

function updateDispatchLoadToolbar(visibleRows = []) {
  pruneSelectedDispatchLoadIds();

  const isMetaStage = dispatchLoadStage === "meta";
  const effectiveSelectedSet = isMetaStage ? confirmedDispatchIdsForLoading : selectedDispatchIdsForLoading;
  const selectedCount = effectiveSelectedSet.size;

  if (dispatchSelectedCount) {
    dispatchSelectedCount.textContent = isMetaStage
      ? `已确认 ${selectedCount} 条（第2步：填写车辆信息）`
      : `已选 ${selectedCount} 条（第1步：先选货）`;
  }
  if (dispatchLoadBtn) {
    dispatchLoadBtn.textContent = isMetaStage ? "填写并确认装车" : "确认配货";
    dispatchLoadBtn.disabled = selectedCount === 0;
  }
  if (dispatchResetStepBtn) {
    dispatchResetStepBtn.hidden = !isMetaStage;
    dispatchResetStepBtn.disabled = !isMetaStage;
  }

  updateDispatchMetaInputsState(!isMetaStage);

  if (!dispatchSelectAllLoad) {
    return;
  }

  if (isMetaStage) {
    dispatchSelectAllLoad.disabled = true;
    dispatchSelectAllLoad.checked = false;
    dispatchSelectAllLoad.indeterminate = false;
    return;
  }

  const selectableVisibleIds = visibleRows
    .filter((item) => item.type === "receive" && !normalizeShipmentLoadedFlag(item.isLoaded))
    .map((item) => item.id);
  const selectedVisibleCount = selectableVisibleIds.filter((id) => selectedDispatchIdsForLoading.has(id)).length;

  dispatchSelectAllLoad.disabled = selectableVisibleIds.length === 0;
  dispatchSelectAllLoad.checked = selectableVisibleIds.length > 0 && selectedVisibleCount === selectableVisibleIds.length;
  dispatchSelectAllLoad.indeterminate = selectedVisibleCount > 0 && selectedVisibleCount < selectableVisibleIds.length;
}

function handleDispatchLoadSelectionChange(event) {
  const checkbox = event.target.closest(".dispatch-load-checkbox");
  if (!checkbox) {
    return;
  }

  if (dispatchLoadStage === "meta") {
    renderDispatchTable();
    return;
  }

  const { id } = checkbox.dataset;
  if (!id) {
    return;
  }

  if (checkbox.checked) {
    selectedDispatchIdsForLoading.add(id);
  } else {
    selectedDispatchIdsForLoading.delete(id);
  }
  const row = checkbox.closest("tr");
  if (row) {
    row.classList.toggle("dispatch-selected-row", checkbox.checked);
  }

  updateDispatchLoadToolbar(currentVisibleDispatchRows);
}

function handleDispatchSelectAllChange(event) {
  if (dispatchLoadStage === "meta") {
    return;
  }

  if (!currentVisibleDispatchRows.length) {
    return;
  }

  const checked = Boolean(event.target.checked);
  currentVisibleDispatchRows.forEach((item) => {
    const selectable = item.type === "receive" && !normalizeShipmentLoadedFlag(item.isLoaded);
    if (!selectable) {
      return;
    }
    if (checked) {
      selectedDispatchIdsForLoading.add(item.id);
    } else {
      selectedDispatchIdsForLoading.delete(item.id);
    }
  });

  renderDispatchTable();
}

async function handleBatchDispatchLoad() {
  pruneSelectedDispatchLoadIds();

  if (dispatchLoadStage !== "meta") {
    const selectedIds = Array.from(selectedDispatchIdsForLoading);
    if (selectedIds.length === 0) {
      alert("请先勾选待配车的产品。");
      updateDispatchLoadToolbar(currentVisibleDispatchRows);
      return;
    }

    if (
      !(await confirmDialog(`已选择 ${selectedIds.length} 条货物，确认配货无误并进入车辆信息填写吗？`, {
        title: "确认配货",
      }))
    ) {
      return;
    }

    confirmedDispatchIdsForLoading = new Set(selectedIds);
    dispatchLoadStage = "meta";
    renderDispatchTable();
    openDispatchMetaModal();
    return;
  }

  const selectedIds = Array.from(confirmedDispatchIdsForLoading);
  if (!selectedIds.length) {
    closeDispatchMetaModal();
    resetDispatchLoadFlow({ clearSelected: true });
    renderDispatchTable();
    alert("已确认货物为空，请重新选择后再配车。");
    return;
  }

  openDispatchMetaModal();
}

async function handleDispatchLoadStageReset() {
  if (dispatchLoadStage !== "meta") {
    return;
  }

  if (
    !(await confirmDialog("确认返回重选货物吗？已填写的车辆信息会保留。", {
      title: "返回重选确认",
    }))
  ) {
    return;
  }

  closeDispatchMetaModal();
  resetDispatchLoadFlow({ clearSelected: false });
  renderDispatchTable();
}

async function handleDispatchMetaConfirmLoad() {
  pruneSelectedDispatchLoadIds();

  const selectedIds = Array.from(confirmedDispatchIdsForLoading);
  if (!selectedIds.length) {
    closeDispatchMetaModal();
    resetDispatchLoadFlow({ clearSelected: true });
    renderDispatchTable();
    alert("已确认货物为空，请重新选择后再配车。");
    return;
  }

  const dispatchMeta = getDispatchMetaFromInputs();
  if (!dispatchMeta) {
    return;
  }

  if (
    !(await confirmDialog(
      `确认将选中的 ${selectedIds.length} 条记录执行配货装车吗？\n` +
        `配车信息：${dispatchMeta.dispatchDestination || "-"}`,
      { title: "确认装车" }
    ))
  ) {
    return;
  }

  const { updatedCount, record } = applyBatchLoadForIds(selectedIds, dispatchMeta);
  closeDispatchMetaModal();
  resetDispatchLoadFlow({ clearSelected: true });

  if (updatedCount === 0) {
    renderDispatchTable();
    alert("未找到可配车的入库记录。");
    return;
  }

  renderShipmentSection();
  alert(`已完成配货装车：${updatedCount} 条。${record ? "已新增 1 条配车记录。" : ""}`);
}

function buildDispatchRecord(rows, loadedAt, dispatchMeta = {}) {
  if (!Array.isArray(rows) || rows.length === 0) {
    return null;
  }

  const normalizedMeta = normalizeDispatchMeta(dispatchMeta);
  if (!normalizedMeta.dispatchDate) {
    normalizedMeta.dispatchDate = parseDateFromAny(loadedAt) || "";
    normalizedMeta.dispatchDestination = buildDispatchDestinationText(normalizedMeta);
  }
  const items = rows.map((item) => buildDispatchRecordItem(item, normalizedMeta));
  const totals = computeDispatchRecordTotals(items);
  const productsSummary = buildDispatchProductsSummary(items);

  return {
    id: makeId(),
    loadedAt,
    createdAt: loadedAt,
    itemCount: items.length,
    totalAmount: totals.totalAmount,
    totalWeight: totals.totalWeight,
    totalCubic: totals.totalCubic,
    totalPieces: totals.totalPieces,
    dispatchDate: normalizedMeta.dispatchDate,
    dispatchTruckNo: normalizedMeta.dispatchTruckNo,
    dispatchPlateNo: normalizedMeta.dispatchPlateNo,
    dispatchDriver: normalizedMeta.dispatchDriver,
    dispatchContactName: normalizedMeta.dispatchContactName,
    dispatchDestination: normalizedMeta.dispatchDestination,
    productsSummary,
    items,
  };
}

function buildDispatchRecordItem(item, dispatchMeta = {}) {
  const measureUnit = normalizeMeasureUnit(item.measureUnit, DEFAULT_MEASURE_UNIT) || DEFAULT_MEASURE_UNIT;
  const normalizedMeta = normalizeDispatchMeta({
    dispatchDate: item.dispatchDate || dispatchMeta.dispatchDate,
    dispatchTruckNo: item.dispatchTruckNo || dispatchMeta.dispatchTruckNo,
    dispatchPlateNo: item.dispatchPlateNo || dispatchMeta.dispatchPlateNo,
    dispatchDriver: item.dispatchDriver || dispatchMeta.dispatchDriver,
    dispatchContactName: item.dispatchContactName || dispatchMeta.dispatchContactName,
  });

  return {
    shipmentId: cleanText(item.id),
    date: cleanText(item.date),
    serialNo: cleanText(item.serialNo),
    manufacturer: cleanText(item.manufacturer),
    customer: cleanText(item.customer),
    product: cleanText(item.product),
    measureUnit,
    quantity: toFiniteNumber(item.quantity),
    unitPrice: Math.max(0, toFiniteNumber(item.unitPrice)),
    amount: Math.max(0, toFiniteNumber(item.amount)),
    unitWeight: Math.max(0, getShipmentUnitWeight(item)),
    unitCubic: Math.max(0, getShipmentUnitCubic(item)),
    totalWeight: Math.max(0, getShipmentTotalWeight(item)),
    totalCubic: Math.max(0, getShipmentTotalCubic(item)),
    note: cleanText(item.note),
    dispatchDate: normalizedMeta.dispatchDate,
    dispatchTruckNo: normalizedMeta.dispatchTruckNo,
    dispatchPlateNo: normalizedMeta.dispatchPlateNo,
    dispatchDriver: normalizedMeta.dispatchDriver,
    dispatchContactName: normalizedMeta.dispatchContactName,
    dispatchDestination: normalizedMeta.dispatchDestination,
  };
}

function computeDispatchRecordTotals(items) {
  return items.reduce(
    (acc, item) => {
      const unit = normalizeMeasureUnit(item.measureUnit, DEFAULT_MEASURE_UNIT) || DEFAULT_MEASURE_UNIT;
      const quantity = Math.max(0, toFiniteNumber(item.quantity));
      acc.totalAmount += Math.max(0, toFiniteNumber(item.amount));
      acc.totalWeight += Math.max(0, toFiniteNumber(item.totalWeight));
      acc.totalCubic += Math.max(0, toFiniteNumber(item.totalCubic));
      if (unit === "piece") {
        acc.totalPieces += quantity;
      }
      return acc;
    },
    { totalAmount: 0, totalWeight: 0, totalCubic: 0, totalPieces: 0 }
  );
}

function getDispatchRecordsForView() {
  return dispatchRecords.slice().sort((a, b) => {
    const aTime = Math.max(0, toFiniteNumber(a.loadedAt) || toFiniteNumber(a.createdAt));
    const bTime = Math.max(0, toFiniteNumber(b.loadedAt) || toFiniteNumber(b.createdAt));
    if (aTime !== bTime) {
      return bTime - aTime;
    }
    return String(b.id || "").localeCompare(String(a.id || ""));
  });
}

function renderDispatchRecordTable() {
  if (!dispatchRecordBody) {
    return;
  }

  const rows = getDispatchRecordsForView();
  if (rows.length === 0) {
    dispatchRecordBody.innerHTML = '<tr><td class="empty" colspan="13">暂无配车记录</td></tr>';
    return;
  }

  dispatchRecordBody.innerHTML = rows
    .map((record) => {
      const recordId = escapeHtml(record.id);
      const dispatchMeta = getDispatchMetaFromRecord(record);
      const productsSummary = getDispatchProductsSummary(record);
      return `
      <tr class="dispatch-record-row" data-dispatch-record-id="${recordId}">
        <td>${formatDateTime(record.loadedAt || record.createdAt)}</td>
        <td>${escapeHtml(dispatchMeta.dispatchDate || "-")}</td>
        <td>${escapeHtml(dispatchMeta.dispatchTruckNo || "-")}</td>
        <td>${escapeHtml(dispatchMeta.dispatchPlateNo || "-")}</td>
        <td>${escapeHtml(dispatchMeta.dispatchDriver || "-")}</td>
        <td>${escapeHtml(dispatchMeta.dispatchContactName || "-")}</td>
        <td title="${escapeHtml(productsSummary)}">${escapeHtml(productsSummary)}</td>
        <td>${formatNumber(Math.max(0, toFiniteNumber(record.itemCount || record.items?.length)))}</td>
        <td>${formatCurrency(Math.max(0, toFiniteNumber(record.totalAmount)))}</td>
        <td>${formatDecimal(Math.max(0, toFiniteNumber(record.totalWeight)))}</td>
        <td>${formatDecimal(Math.max(0, toFiniteNumber(record.totalCubic)))}</td>
        <td>${formatNumber(Math.max(0, toFiniteNumber(record.totalPieces)))}</td>
        <td>
          <div class="dispatch-record-actions-inline">
            <button type="button" class="ghost-btn dispatch-record-view-btn" data-dispatch-record-id="${recordId}">详情</button>
            <button type="button" class="delete-btn dispatch-record-delete-btn" data-dispatch-record-id="${recordId}">删除</button>
          </div>
        </td>
      </tr>
      `;
    })
    .join("");
}

async function handleDispatchRecordClick(event) {
  const deleteBtn = event.target.closest(".dispatch-record-delete-btn");
  if (deleteBtn) {
    const recordId = cleanText(deleteBtn.dataset.dispatchRecordId);
    if (!recordId) {
      return;
    }
    await deleteDispatchRecordById(recordId);
    return;
  }

  const trigger = event.target.closest("[data-dispatch-record-id]");
  if (!trigger) {
    return;
  }
  const recordId = cleanText(trigger.dataset.dispatchRecordId);
  if (!recordId) {
    return;
  }
  openDispatchRecordModal(recordId);
}

async function handleDispatchRecordDeleteClick() {
  if (!activeDispatchRecordId) {
    alert("请先选择配车记录。");
    return;
  }
  await deleteDispatchRecordById(activeDispatchRecordId);
}

function resetDispatchRecordEditDraft() {
  dispatchRecordEditMode = false;
  dispatchRecordAddKeyword = "";
  dispatchRecordPendingAddIds.clear();
  if (dispatchRecordAddSearchInput) {
    dispatchRecordAddSearchInput.value = "";
  }
}

function setDispatchRecordEditMode(nextMode) {
  const enabled = Boolean(nextMode) && Boolean(activeDispatchRecordId);
  dispatchRecordEditMode = enabled;
  dispatchRecordPendingAddIds.clear();
  if (!enabled) {
    dispatchRecordAddKeyword = "";
    if (dispatchRecordAddSearchInput) {
      dispatchRecordAddSearchInput.value = "";
    }
  }
  if (dispatchRecordEditBtn) {
    dispatchRecordEditBtn.textContent = enabled ? "退出二次配车" : "二次配车";
  }
  syncDispatchRecordDetailView();
}

function handleDispatchRecordEditToggle() {
  if (!activeDispatchRecordId) {
    alert("请先选择一条配车记录。");
    return;
  }
  setDispatchRecordEditMode(!dispatchRecordEditMode);
}

function getDispatchRecordPendingRows(record) {
  if (!record) {
    return [];
  }

  const keyword = cleanText(dispatchRecordAddKeyword).toLowerCase();
  return shipments
    .filter((item) => item.type === "receive" && !normalizeShipmentLoadedFlag(item.isLoaded))
    .filter((item) => {
      if (!keyword) {
        return true;
      }
      const haystack = [
        item.date,
        item.serialNo,
        item.manufacturer,
        item.customer,
        item.product,
        item.note,
      ]
        .map((field) => cleanText(field).toLowerCase())
        .join(" ");
      return haystack.includes(keyword);
    })
    .sort(sortByDateDesc);
}

function pruneDispatchRecordPendingAddIds(record) {
  const pendingIds = new Set(getDispatchRecordPendingRows(record).map((item) => cleanText(item.id)));
  dispatchRecordPendingAddIds.forEach((id) => {
    if (!pendingIds.has(id)) {
      dispatchRecordPendingAddIds.delete(id);
    }
  });
}

function renderDispatchRecordEditPanel(record) {
  if (!dispatchRecordEditPanel || !dispatchRecordAddBody) {
    return;
  }

  dispatchRecordEditPanel.classList.toggle("is-hidden", !dispatchRecordEditMode);

  if (!dispatchRecordEditMode) {
    if (dispatchRecordAddBody) {
      dispatchRecordAddBody.innerHTML = "";
    }
    if (dispatchRecordAddSelectedCount) {
      dispatchRecordAddSelectedCount.textContent = "已选 0 条";
    }
    if (dispatchRecordAddBtn) {
      dispatchRecordAddBtn.disabled = true;
    }
    if (dispatchRecordAddSelectAll) {
      dispatchRecordAddSelectAll.checked = false;
      dispatchRecordAddSelectAll.indeterminate = false;
      dispatchRecordAddSelectAll.disabled = true;
    }
    return;
  }

  const pendingRows = getDispatchRecordPendingRows(record);
  pruneDispatchRecordPendingAddIds(record);
  const selectedVisibleCount = pendingRows.filter((item) => dispatchRecordPendingAddIds.has(cleanText(item.id))).length;

  if (!pendingRows.length) {
    dispatchRecordAddBody.innerHTML =
      '<tr><td class="empty" colspan="9">暂无可追加的待配车货物</td></tr>';
  } else {
    dispatchRecordAddBody.innerHTML = pendingRows
      .map((item) => {
        const id = cleanText(item.id);
        const checked = dispatchRecordPendingAddIds.has(id);
        return `
          <tr>
            <td class="shipment-load-select-cell">
              <input
                type="checkbox"
                class="dispatch-record-add-checkbox"
                data-id="${escapeHtml(id)}"
                ${checked ? "checked" : ""}
                aria-label="勾选追加货物"
              />
            </td>
            <td>${escapeHtml(item.date || "-")}</td>
            <td class="mono">${escapeHtml(item.serialNo || "-")}</td>
            <td>${escapeHtml(item.manufacturer || "-")}</td>
            <td>${escapeHtml(item.customer || "-")}</td>
            <td>${escapeHtml(item.product || "-")}</td>
            <td>${formatMeasureValue(item.quantity, item.measureUnit)}</td>
            <td>${formatMeasureUnit(item.measureUnit)}</td>
            <td>${formatCurrency(Math.max(0, toFiniteNumber(item.amount)))}</td>
          </tr>
        `;
      })
      .join("");
  }

  if (dispatchRecordAddSelectedCount) {
    dispatchRecordAddSelectedCount.textContent = `已选 ${dispatchRecordPendingAddIds.size} 条`;
  }
  if (dispatchRecordAddBtn) {
    dispatchRecordAddBtn.disabled = dispatchRecordPendingAddIds.size === 0;
  }
  if (dispatchRecordAddSelectAll) {
    dispatchRecordAddSelectAll.disabled = pendingRows.length === 0;
    dispatchRecordAddSelectAll.checked = pendingRows.length > 0 && selectedVisibleCount === pendingRows.length;
    dispatchRecordAddSelectAll.indeterminate =
      selectedVisibleCount > 0 && selectedVisibleCount < pendingRows.length;
  }
}

function handleDispatchRecordAddSearchInput(event) {
  dispatchRecordAddKeyword = cleanText(event?.target?.value);
  syncDispatchRecordDetailView();
}

function handleDispatchRecordAddSearchClear() {
  dispatchRecordAddKeyword = "";
  if (dispatchRecordAddSearchInput) {
    dispatchRecordAddSearchInput.value = "";
  }
  dispatchRecordPendingAddIds.clear();
  syncDispatchRecordDetailView();
}

function handleDispatchRecordAddSelectionChange(event) {
  const checkbox = event.target.closest(".dispatch-record-add-checkbox");
  if (!checkbox || !dispatchRecordEditMode) {
    return;
  }
  const id = cleanText(checkbox.dataset.id);
  if (!id) {
    return;
  }
  if (checkbox.checked) {
    dispatchRecordPendingAddIds.add(id);
  } else {
    dispatchRecordPendingAddIds.delete(id);
  }
  syncDispatchRecordDetailView();
}

function handleDispatchRecordAddSelectAllChange(event) {
  if (!dispatchRecordEditMode || !activeDispatchRecordId) {
    return;
  }
  const record = getDispatchRecordById(activeDispatchRecordId);
  if (!record) {
    return;
  }
  const pendingRows = getDispatchRecordPendingRows(record);
  if (!pendingRows.length) {
    return;
  }
  const shouldSelect = Boolean(event.target?.checked);
  pendingRows.forEach((item) => {
    const id = cleanText(item.id);
    if (!id) {
      return;
    }
    if (shouldSelect) {
      dispatchRecordPendingAddIds.add(id);
    } else {
      dispatchRecordPendingAddIds.delete(id);
    }
  });
  syncDispatchRecordDetailView();
}

function appendShipmentsToDispatchRecord(recordId, shipmentIds = []) {
  const safeRecordId = cleanText(recordId);
  const uniqueShipmentIds = Array.from(new Set((Array.isArray(shipmentIds) ? shipmentIds : []).map((id) => cleanText(id)).filter(Boolean)));
  if (!safeRecordId || !uniqueShipmentIds.length) {
    return { addedCount: 0, updatedRecord: null };
  }

  const recordIndex = dispatchRecords.findIndex((item) => cleanText(item.id) === safeRecordId);
  if (recordIndex < 0) {
    return { addedCount: 0, updatedRecord: null };
  }

  const currentRecord = dispatchRecords[recordIndex];
  const currentItems = Array.isArray(currentRecord.items) ? currentRecord.items.slice() : [];
  const existingIds = new Set(currentItems.map((item) => cleanText(item.shipmentId || item.id)).filter(Boolean));
  const recordMeta = getDispatchMetaFromRecord(currentRecord);
  const addedAt = Date.now();
  const appendedItems = [];

  uniqueShipmentIds.forEach((shipmentId) => {
    if (existingIds.has(shipmentId)) {
      return;
    }
    const target = shipments.find((item) => cleanText(item.id) === shipmentId);
    if (!target || target.type !== "receive" || normalizeShipmentLoadedFlag(target.isLoaded)) {
      return;
    }
    applyShipmentDispatchState(target, { timestamp: addedAt, meta: recordMeta });
    const item = buildDispatchRecordItem(target, recordMeta);
    currentItems.push(item);
    appendedItems.push(item);
    existingIds.add(shipmentId);
    selectedDispatchIdsForLoading.delete(shipmentId);
    confirmedDispatchIdsForLoading.delete(shipmentId);
  });

  if (!appendedItems.length) {
    return { addedCount: 0, updatedRecord: currentRecord };
  }

  const rebuiltRecord = rebuildDispatchRecordWithItems(currentRecord, currentItems);
  if (rebuiltRecord) {
    dispatchRecords[recordIndex] = rebuiltRecord;
  }

  writeStorage(STORAGE_KEYS.shipments, shipments);
  writeStorage(STORAGE_KEYS.dispatchRecords, dispatchRecords);
  return { addedCount: appendedItems.length, updatedRecord: rebuiltRecord };
}

async function handleDispatchRecordAppendClick() {
  if (!dispatchRecordEditMode || !activeDispatchRecordId) {
    alert("请先进入二次配车模式。");
    return;
  }
  const selectedIds = Array.from(dispatchRecordPendingAddIds);
  if (!selectedIds.length) {
    alert("请先勾选要追加的货物。");
    return;
  }

  const confirmed = await confirmDialog(`确认将选中的 ${selectedIds.length} 条货物追加到当前车次吗？`, {
    title: "确认追加配货",
    confirmText: "确认追加",
    cancelText: "取消"
  });
  if (!confirmed) {
    return;
  }

  const { addedCount } = appendShipmentsToDispatchRecord(activeDispatchRecordId, selectedIds);
  dispatchRecordPendingAddIds.clear();
  renderShipmentSection();
  if (addedCount <= 0) {
    alert("未追加成功：所选货物可能已被其它车次占用。");
    return;
  }
  alert(`已追加 ${addedCount} 条货物到当前车次。`);
}

function removeShipmentFromDispatchRecord(recordId, shipmentId) {
  const safeRecordId = cleanText(recordId);
  const safeShipmentId = cleanText(shipmentId);
  if (!safeRecordId || !safeShipmentId) {
    return { removed: false, recordDeleted: false, updatedRecord: null };
  }

  const recordIndex = dispatchRecords.findIndex((item) => cleanText(item.id) === safeRecordId);
  if (recordIndex < 0) {
    return { removed: false, recordDeleted: false, updatedRecord: null };
  }

  const currentRecord = dispatchRecords[recordIndex];
  const currentItems = Array.isArray(currentRecord.items) ? currentRecord.items : [];
  const nextItems = currentItems.filter((item) => cleanText(item.shipmentId || item.id) !== safeShipmentId);
  if (nextItems.length === currentItems.length) {
    return { removed: false, recordDeleted: false, updatedRecord: currentRecord };
  }

  let updatedRecord = null;
  let recordDeleted = false;
  if (!nextItems.length) {
    dispatchRecords.splice(recordIndex, 1);
    recordDeleted = true;
  } else {
    updatedRecord = rebuildDispatchRecordWithItems(currentRecord, nextItems);
    dispatchRecords[recordIndex] = updatedRecord;
  }

  const shipmentTarget = shipments.find((item) => cleanText(item.id) === safeShipmentId);
  if (shipmentTarget) {
    const latestReference = getLatestDispatchReferenceByShipmentId(safeShipmentId);
    if (latestReference) {
      applyShipmentDispatchState(shipmentTarget, latestReference);
    } else {
      resetShipmentDispatchState(shipmentTarget);
    }
    selectedDispatchIdsForLoading.delete(safeShipmentId);
    confirmedDispatchIdsForLoading.delete(safeShipmentId);
  }

  writeStorage(STORAGE_KEYS.shipments, shipments);
  writeStorage(STORAGE_KEYS.dispatchRecords, dispatchRecords);
  return { removed: true, recordDeleted, updatedRecord };
}

async function handleDispatchRecordDetailClick(event) {
  const removeBtn = event.target.closest(".dispatch-record-item-remove-btn");
  if (!removeBtn) {
    return;
  }
  if (!dispatchRecordEditMode) {
    alert("请先点击“二次配车”进入编辑模式。");
    return;
  }
  const shipmentId = cleanText(removeBtn.dataset.shipmentId);
  if (!shipmentId || !activeDispatchRecordId) {
    return;
  }
  const confirmed = await confirmDialog("确认从当前车次移除该货物吗？", {
    title: "确认移除货物",
    confirmText: "确认移除",
    cancelText: "取消"
  });
  if (!confirmed) {
    return;
  }

  const { removed, recordDeleted } = removeShipmentFromDispatchRecord(activeDispatchRecordId, shipmentId);
  if (!removed) {
    alert("移除失败：货物不存在或已变更。");
    return;
  }

  dispatchRecordPendingAddIds.delete(shipmentId);
  renderShipmentSection();
  if (recordDeleted) {
    closeDispatchRecordModal();
    alert("该车次已无货物，记录已自动删除。");
    return;
  }
  alert("货物已从当前车次移除。");
}

function getDispatchRecordTimestamp(record = {}) {
  return Math.max(0, toFiniteNumber(record.loadedAt) || toFiniteNumber(record.createdAt));
}

function getLatestDispatchReferenceByShipmentId(shipmentId) {
  const safeShipmentId = cleanText(shipmentId);
  if (!safeShipmentId) {
    return null;
  }

  let latestReference = null;
  dispatchRecords.forEach((record) => {
    const items = Array.isArray(record.items) ? record.items : [];
    if (!items.length) {
      return;
    }
    const recordTimestamp = getDispatchRecordTimestamp(record);
    const recordMeta = getDispatchMetaFromRecord(record);
    items.forEach((item) => {
      const itemShipmentId = cleanText(item.shipmentId || item.id);
      if (itemShipmentId !== safeShipmentId) {
        return;
      }
      const metaBase = normalizeDispatchMeta({
        dispatchDate: item.dispatchDate || recordMeta.dispatchDate,
        dispatchTruckNo: item.dispatchTruckNo || recordMeta.dispatchTruckNo,
        dispatchPlateNo: item.dispatchPlateNo || recordMeta.dispatchPlateNo,
        dispatchDriver: item.dispatchDriver || recordMeta.dispatchDriver,
        dispatchContactName: item.dispatchContactName || recordMeta.dispatchContactName,
      });
      const directDestination = cleanText(item.dispatchDestination || record.dispatchDestination);
      const dispatchDestination = directDestination || metaBase.dispatchDestination;
      const reference = {
        timestamp: recordTimestamp,
        meta: {
          ...metaBase,
          dispatchDestination,
        },
      };
      if (!latestReference || reference.timestamp >= latestReference.timestamp) {
        latestReference = reference;
      }
    });
  });

  return latestReference;
}

function resetShipmentDispatchState(item) {
  if (!item) {
    return;
  }
  item.type = "receive";
  item.isLoaded = false;
  item.loadedAt = 0;
  item.dispatchDate = "";
  item.dispatchTruckNo = "";
  item.dispatchPlateNo = "";
  item.dispatchDriver = "";
  item.dispatchContactName = "";
  item.dispatchDestination = "";
  item.destination = "";
}

function applyShipmentDispatchState(item, reference) {
  if (!item || !reference) {
    return;
  }
  const meta = normalizeDispatchMeta(reference.meta || {});
  const dispatchDestination = cleanText(reference.meta?.dispatchDestination) || meta.dispatchDestination;
  item.type = "ship";
  item.isLoaded = true;
  item.loadedAt = Math.max(0, toFiniteNumber(reference.timestamp)) || Date.now();
  item.dispatchDate = meta.dispatchDate;
  item.dispatchTruckNo = meta.dispatchTruckNo;
  item.dispatchPlateNo = meta.dispatchPlateNo;
  item.dispatchDriver = meta.dispatchDriver;
  item.dispatchContactName = meta.dispatchContactName;
  item.dispatchDestination = dispatchDestination;
  item.destination = dispatchDestination;
}

function syncShipmentsAfterDispatchRecordDelete(removedRecords = []) {
  const records = Array.isArray(removedRecords) ? removedRecords : [];
  const affectedShipmentIds = new Set();
  records.forEach((record) => {
    const items = Array.isArray(record.items) ? record.items : [];
    items.forEach((item) => {
      const shipmentId = cleanText(item.shipmentId || item.id);
      if (shipmentId) {
        affectedShipmentIds.add(shipmentId);
      }
    });
  });

  if (!affectedShipmentIds.size) {
    return;
  }

  affectedShipmentIds.forEach((shipmentId) => {
    const target = shipments.find((item) => cleanText(item.id) === shipmentId);
    if (!target) {
      return;
    }
    const latestReference = getLatestDispatchReferenceByShipmentId(shipmentId);
    if (latestReference) {
      applyShipmentDispatchState(target, latestReference);
    } else {
      resetShipmentDispatchState(target);
    }
    selectedDispatchIdsForLoading.delete(shipmentId);
    confirmedDispatchIdsForLoading.delete(shipmentId);
  });
}

async function deleteDispatchRecordById(recordId) {
  const safeRecordId = cleanText(recordId);
  if (!safeRecordId) {
    return false;
  }

  const recordIndex = dispatchRecords.findIndex((item) => cleanText(item.id) === safeRecordId);
  if (recordIndex < 0) {
    alert("当前配车记录不存在或已被删除。");
    return false;
  }

  const record = dispatchRecords[recordIndex];
  const affectedCount = Math.max(0, toFiniteNumber(record.itemCount || record.items?.length));
  const confirmed = await confirmDialog(
    `确认删除这条配车记录吗？\n删除后会更新总库货物状态。${affectedCount ? `\n影响货物：${affectedCount} 条` : ""}`,
    {
      title: "确认删除配车记录",
      confirmText: "确认删除",
      cancelText: "取消"
    }
  );
  if (!confirmed) {
    return false;
  }

  const removedRecords = dispatchRecords.splice(recordIndex, 1);
  syncShipmentsAfterDispatchRecordDelete(removedRecords);

  writeStorage(STORAGE_KEYS.shipments, shipments);
  writeStorage(STORAGE_KEYS.dispatchRecords, dispatchRecords);

  if (activeDispatchRecordId === safeRecordId) {
    closeDispatchRecordModal();
  }

  renderShipmentSection();
  alert("配车记录已删除。");
  return true;
}

function handleDispatchRecordExportClick() {
  if (!activeDispatchRecordId) {
    alert("请先选择配车记录。");
    return;
  }

  const record = getDispatchRecordById(activeDispatchRecordId);
  if (!record) {
    alert("当前配车记录不存在或已被删除。");
    return;
  }

  exportDispatchRecordExcel(record);
}

function getDispatchRecordById(id) {
  const recordId = cleanText(id);
  if (!recordId) {
    return null;
  }
  return dispatchRecords.find((item) => item.id === recordId) || null;
}

function openConfirmModal(options = {}) {
  if (!confirmModal) {
    return;
  }
  const title = cleanText(options.title) || "操作确认";
  const message = cleanText(options.message) || "请确认是否继续该操作。";
  const confirmText = cleanText(options.confirmText) || "确认";
  const cancelText = cleanText(options.cancelText) || "取消";

  if (confirmModalTitle) {
    confirmModalTitle.textContent = title;
  }
  if (confirmModalMessage) {
    confirmModalMessage.textContent = message;
  }
  if (confirmModalConfirmBtn) {
    confirmModalConfirmBtn.textContent = confirmText;
  }
  if (confirmModalCancelBtn) {
    confirmModalCancelBtn.textContent = cancelText;
  }

  confirmModal.classList.add("is-open");
  syncModalBodyState();
  if (confirmModalConfirmBtn) {
    confirmModalConfirmBtn.focus();
  }
}

function closeConfirmModal() {
  if (!confirmModal) {
    return;
  }
  confirmModal.classList.remove("is-open");
  syncModalBodyState();
}

function resolveConfirmDialog(confirmed) {
  const resolver = pendingConfirmDialogResolve;
  pendingConfirmDialogResolve = null;
  closeConfirmModal();
  if (resolver) {
    resolver(Boolean(confirmed));
  }
}

function confirmDialog(message, options = {}) {
  const text = cleanText(message) || "请确认是否继续该操作。";
  if (!confirmModal || !confirmModalConfirmBtn || !confirmModalCancelBtn || !confirmModalMessage) {
    alert("确认弹窗组件未加载，当前操作已取消。");
    return Promise.resolve(false);
  }

  if (pendingConfirmDialogResolve) {
    resolveConfirmDialog(false);
  }

  return new Promise((resolve) => {
    pendingConfirmDialogResolve = resolve;
    openConfirmModal({ ...options, message: text });
  });
}

function syncModalBodyState() {
  const hasOpenModal = Boolean(
    document.querySelector(
      ".dispatch-record-modal.is-open, .dispatch-meta-modal.is-open, .confirm-modal.is-open, .shipment-import-preview-modal.is-open, .shipment-export-preview-modal.is-open, .finance-export-preview-modal.is-open, .storage-import-preview-modal.is-open, .storage-required-modal.is-open"
    )
  );
  document.body.classList.toggle("modal-open", hasOpenModal);
}

function openDispatchMetaModal() {
  if (!dispatchMetaModal) {
    return;
  }

  if (dispatchMetaDateInput && !dispatchMetaDateInput.value) {
    dispatchMetaDateInput.value = dateValue(new Date());
  }
  if (dispatchMetaModalTip) {
    dispatchMetaModalTip.textContent = `已确认货物 ${confirmedDispatchIdsForLoading.size} 条，请填写车辆信息后确认装车。`;
  }

  dispatchMetaModal.classList.add("is-open");
  syncModalBodyState();
  if (dispatchMetaDateInput) {
    dispatchMetaDateInput.focus();
  }
}

function closeDispatchMetaModal() {
  if (!dispatchMetaModal) {
    return;
  }

  dispatchMetaModal.classList.remove("is-open");
  syncModalBodyState();
}

function openDispatchRecordModal(recordId) {
  const record = getDispatchRecordById(recordId);
  if (!record || !dispatchRecordModal) {
    return;
  }

  resetDispatchRecordEditDraft();
  activeDispatchRecordId = record.id;
  renderDispatchRecordDetail(record);
  dispatchRecordModal.classList.add("is-open");
  syncModalBodyState();
}

function syncDispatchRecordDetailView() {
  if (!activeDispatchRecordId || !dispatchRecordModal?.classList.contains("is-open")) {
    return;
  }
  const record = getDispatchRecordById(activeDispatchRecordId);
  if (!record) {
    closeDispatchRecordModal();
    return;
  }
  renderDispatchRecordDetail(record);
}

function renderDispatchRecordDetail(record) {
  if (!dispatchRecordModalSummary || !dispatchRecordDetailBody) {
    return;
  }
  const dispatchMeta = getDispatchMetaFromRecord(record);
  const productsSummary = getDispatchProductsSummary(record);

  if (dispatchRecordModalTitle) {
    dispatchRecordModalTitle.textContent = `配车记录详情（${formatDateTime(record.loadedAt || record.createdAt)}）`;
  }
  if (dispatchRecordEditBtn) {
    dispatchRecordEditBtn.textContent = dispatchRecordEditMode ? "退出二次配车" : "二次配车";
  }

  dispatchRecordModalSummary.innerHTML = `
    <article class="dispatch-record-summary-item">
      <p>配车日期</p>
      <strong>${escapeHtml(dispatchMeta.dispatchDate || "-")}</strong>
    </article>
    <article class="dispatch-record-summary-item">
      <p>几号车</p>
      <strong>${escapeHtml(dispatchMeta.dispatchTruckNo || "-")}</strong>
    </article>
    <article class="dispatch-record-summary-item">
      <p>车牌号</p>
      <strong>${escapeHtml(dispatchMeta.dispatchPlateNo || "-")}</strong>
    </article>
    <article class="dispatch-record-summary-item">
      <p>司机姓名</p>
      <strong>${escapeHtml(dispatchMeta.dispatchDriver || "-")}</strong>
    </article>
    <article class="dispatch-record-summary-item">
      <p>联系方式</p>
      <strong>${escapeHtml(dispatchMeta.dispatchContactName || "-")}</strong>
    </article>
    <article class="dispatch-record-summary-item">
      <p>配好产品</p>
      <strong>${escapeHtml(productsSummary)}</strong>
    </article>
    <article class="dispatch-record-summary-item">
      <p>总价</p>
      <strong>${formatCurrency(Math.max(0, toFiniteNumber(record.totalAmount)))}</strong>
    </article>
    <article class="dispatch-record-summary-item">
      <p>总重量</p>
      <strong>${formatDecimal(Math.max(0, toFiniteNumber(record.totalWeight)))} 吨</strong>
    </article>
    <article class="dispatch-record-summary-item">
      <p>总立方</p>
      <strong>${formatDecimal(Math.max(0, toFiniteNumber(record.totalCubic)))}</strong>
    </article>
    <article class="dispatch-record-summary-item">
      <p>总件数(中件数)</p>
      <strong>${formatNumber(Math.max(0, toFiniteNumber(record.totalPieces)))}</strong>
    </article>
    <article class="dispatch-record-summary-item">
      <p>配车条数</p>
      <strong>${formatNumber(Math.max(0, toFiniteNumber(record.itemCount || record.items?.length)))}</strong>
    </article>
  `;

  const items = Array.isArray(record.items) ? record.items : [];
  if (!items.length) {
    dispatchRecordDetailBody.innerHTML = '<tr><td class="empty" colspan="12">当前配车记录无明细</td></tr>';
    renderDispatchRecordEditPanel(record);
    return;
  }

  const detailTotals = computeDispatchDetailTotals(items);
  const detailRowsHtml = items
    .map(
      (item) => `
      <tr>
        <td>${escapeHtml(item.date || "-")}</td>
        <td class="mono">${escapeHtml(item.serialNo || "-")}</td>
        <td>${escapeHtml(item.manufacturer || "-")}</td>
        <td>${escapeHtml(item.customer || "-")}</td>
        <td>${escapeHtml(item.product || "-")}</td>
        <td>${formatMeasureValue(item.quantity, item.measureUnit)}</td>
        <td>${formatMeasureUnit(item.measureUnit)}</td>
        <td>${formatCurrency(Math.max(0, toFiniteNumber(item.amount)))}</td>
        <td>${formatDecimal(Math.max(0, toFiniteNumber(item.totalWeight)))}</td>
        <td>${formatDecimal(Math.max(0, toFiniteNumber(item.totalCubic)))}</td>
        <td>${escapeHtml(item.note || "-")}</td>
        <td>
          ${
            dispatchRecordEditMode
              ? `<button type="button" class="delete-btn dispatch-record-item-remove-btn" data-shipment-id="${escapeHtml(
                  cleanText(item.shipmentId || item.id)
                )}">移除</button>`
              : "-"
          }
        </td>
      </tr>
    `
    )
    .join("");

  const totalRowHtml = `
    <tr class="dispatch-record-total-row">
      <td>合计</td>
      <td>-</td>
      <td>-</td>
      <td>-</td>
      <td>-</td>
      <td>${formatDecimal(detailTotals.totalQuantity)}</td>
      <td>-</td>
      <td>${formatCurrency(detailTotals.totalAmount)}</td>
      <td>${formatDecimal(detailTotals.totalWeight)}</td>
      <td>${formatDecimal(detailTotals.totalCubic)}</td>
      <td>-</td>
      <td>-</td>
    </tr>
  `;

  dispatchRecordDetailBody.innerHTML = `${detailRowsHtml}${totalRowHtml}`;
  renderDispatchRecordEditPanel(record);
}

function computeDispatchDetailTotals(items) {
  return (Array.isArray(items) ? items : []).reduce(
    (acc, item) => {
      acc.totalQuantity += Math.max(0, toFiniteNumber(item.quantity));
      acc.totalUnitPrice += Math.max(0, toFiniteNumber(item.unitPrice));
      acc.totalAmount += Math.max(0, toFiniteNumber(item.amount));
      acc.totalUnitWeight += Math.max(0, toFiniteNumber(item.unitWeight));
      acc.totalUnitCubic += Math.max(0, toFiniteNumber(item.unitCubic));
      acc.totalWeight += Math.max(0, toFiniteNumber(item.totalWeight));
      acc.totalCubic += Math.max(0, toFiniteNumber(item.totalCubic));
      return acc;
    },
    {
      totalQuantity: 0,
      totalUnitPrice: 0,
      totalAmount: 0,
      totalUnitWeight: 0,
      totalUnitCubic: 0,
      totalWeight: 0,
      totalCubic: 0,
    }
  );
}

function exportDispatchRecordExcel(record) {
  if (!record) {
    return;
  }

  if (typeof XLSX === "undefined") {
    alert("未加载 Excel 导出库，请刷新页面后重试。");
    return;
  }

  const workbook = XLSX.utils.book_new();
  const items = Array.isArray(record.items) ? record.items : [];
  const safeTotalAmount = Math.max(0, toFiniteNumber(record.totalAmount));
  const safeTotalWeight = Math.max(0, toFiniteNumber(record.totalWeight));
  const safeTotalCubic = Math.max(0, toFiniteNumber(record.totalCubic));
  const safeTotalPieces = Math.max(0, toFiniteNumber(record.totalPieces));
  const safeItemCount = Math.max(0, toFiniteNumber(record.itemCount || items.length));
  const detailTotals = computeDispatchDetailTotals(items);

  const rows = [
    ["配车汇总"],
    ["配车时间", formatDateTime(record.loadedAt || record.createdAt)],
    ["配车条数", safeItemCount],
    ["总件数", safeTotalPieces],
    ["总价", safeTotalAmount],
    ["总重量", safeTotalWeight],
    ["总立方", safeTotalCubic],
    [],
    ["配车明细"],
    [
      "日期",
      "序号",
      "厂家",
      "客户",
      "产品",
      "计量值",
      "计量单位",
      "单价",
      "金额",
      "单位重量(KG)",
      "单位立方(m3)",
      "总重量(KG)",
      "总立方",
      "备注",
    ],
    ...items.map((item) => [
      cleanText(item.date || "-"),
      cleanText(item.serialNo || "-"),
      cleanText(item.manufacturer || "-"),
      cleanText(item.customer || "-"),
      cleanText(item.product || "-"),
      toFiniteNumber(item.quantity),
      formatMeasureUnit(item.measureUnit),
      Math.max(0, toFiniteNumber(item.unitPrice)),
      Math.max(0, toFiniteNumber(item.amount)),
      Math.max(0, toFiniteNumber(item.unitWeight)),
      Math.max(0, toFiniteNumber(item.unitCubic)),
      Math.max(0, toFiniteNumber(item.totalWeight)),
      Math.max(0, toFiniteNumber(item.totalCubic)),
      cleanText(item.note || "-"),
    ]),
    [
      "合计",
      "",
      "",
      "",
      "",
      detailTotals.totalQuantity,
      "",
      detailTotals.totalUnitPrice,
      detailTotals.totalAmount,
      detailTotals.totalUnitWeight,
      detailTotals.totalUnitCubic,
      detailTotals.totalWeight,
      detailTotals.totalCubic,
      "",
    ],
  ];

  const sheet = XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(workbook, sheet, "配车汇总");
  XLSX.writeFile(workbook, getDispatchRecordExportFilename(record));
}

function getDispatchRecordExportFilename(record) {
  const rawTs = toFiniteNumber(record?.loadedAt || record?.createdAt);
  const sourceTime = rawTs > 0 ? rawTs : Date.now();
  const date = new Date(sourceTime);
  if (Number.isNaN(date.getTime())) {
    return `配车记录详情_${dateValue(new Date()).replaceAll("-", "")}.xlsx`;
  }

  const yyyy = String(date.getFullYear());
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const dd = String(date.getDate()).padStart(2, "0");
  const hh = String(date.getHours()).padStart(2, "0");
  const mi = String(date.getMinutes()).padStart(2, "0");
  const ss = String(date.getSeconds()).padStart(2, "0");
  return `配车记录详情_${yyyy}${mm}${dd}_${hh}${mi}${ss}.xlsx`;
}

function closeDispatchRecordModal() {
  resetDispatchRecordEditDraft();
  activeDispatchRecordId = "";
  if (dispatchRecordModal) {
    dispatchRecordModal.classList.remove("is-open");
  }
  syncModalBodyState();
}

function renderFinanceTable(options = {}) {
  if (options.resetPage) {
    financePage = 1;
  }

  const rows = sortFinanceRows(getFilteredFinanceRecords());
  pruneSelectedFinanceExportIds(rows);
  const balanceById = buildFinanceBalanceMap(rows);
  const paged = paginateRows(rows, financePage, MAIN_PAGE_SIZE);
  financePage = paged.page;
  currentVisibleFinanceRows = paged.items.slice();
  updateMainPagination("finance", paged.page, paged.totalPages, paged.totalRows);

  if (rows.length === 0) {
    currentVisibleFinanceRows = [];
    syncFinanceExportSelectAllState();
    financeTableBody.innerHTML =
      '<tr><td class="empty" colspan="7">暂无符合条件的收支记录</td></tr>';
    return;
  }

  financeTableBody.innerHTML = paged.items
    .map(
      (item) => {
        const itemId = cleanText(item.id);
        const amount = toFiniteNumber(item.amount);
        const income = item.type === "income" ? amount : 0;
        const expense = item.type === "expense" ? amount : 0;
        const balance = toFiniteNumber(balanceById.get(item.id));
        const isSelectedForExport = selectedFinanceIdsForExport.has(itemId);
        const rowClassName = isSelectedForExport ? "finance-selected-row" : "";
        return `
      <tr class="${rowClassName}">
        <td class="finance-export-select-cell">
          <input
            type="checkbox"
            class="finance-export-checkbox"
            data-id="${escapeHtml(itemId)}"
            ${isSelectedForExport ? "checked" : ""}
            aria-label="勾选用于导出Excel"
          />
        </td>
        <td>${escapeHtml(item.date)}</td>
        <td class="finance-ledger-summary">${escapeHtml(resolveFinanceSummary(item))}</td>
        <td class="finance-ledger-income">${income > 0 ? formatCurrency(income) : "-"}</td>
        <td class="finance-ledger-expense">${expense > 0 ? formatCurrency(expense) : "-"}</td>
        <td class="finance-ledger-balance">${formatCurrency(balance)}</td>
        <td><button class="delete-btn" data-table="finance" data-id="${item.id}">删除</button></td>
      </tr>
    `;
      }
    )
    .join("");

  syncFinanceExportSelectAllState();
}

function resolveFinanceSummary(item) {
  const summary = normalizeFinanceSummaryText(cleanText(item.summary || item.category));
  return summary || resolveFinanceNote(item, { fallback: "-" }) || "-";
}

function normalizeFinanceSummaryText(summaryText) {
  const text = cleanText(summaryText);
  if (!text) {
    return "";
  }
  if (text.startsWith(SHIPMENT_PAYMENT_FINANCE_SUMMARY_PREFIX_LEGACY)) {
    return `${SHIPMENT_PAYMENT_FINANCE_SUMMARY_PREFIX}${text.slice(SHIPMENT_PAYMENT_FINANCE_SUMMARY_PREFIX_LEGACY.length)}`;
  }
  return text;
}

function resolveFinanceNote(item, options = {}) {
  const fallback = Object.prototype.hasOwnProperty.call(options, "fallback") ? options.fallback : "-";
  let note = cleanText(item?.note);
  if (!note) {
    return fallback;
  }

  if (isShipmentPaymentAutoFinanceRecord(item)) {
    note = stripShipmentPaymentFinanceHintText(note);
  }
  if (!note) {
    return fallback;
  }

  return note;
}

function buildFinanceBalanceMap(rows) {
  const chronological = rows
    .slice()
    .sort((a, b) => {
      const dateDiff = String(a.date || "").localeCompare(String(b.date || ""));
      if (dateDiff !== 0) {
        return dateDiff;
      }
      return toFiniteNumber(a.createdAt) - toFiniteNumber(b.createdAt);
    });

  let running = 0;
  const map = new Map();

  chronological.forEach((item) => {
    const amount = toFiniteNumber(item.amount);
    if (item.type === "income") {
      running += amount;
    } else {
      running -= amount;
    }
    map.set(item.id, running);
  });

  return map;
}

function getFilteredDispatchRows() {
  const status = dispatchFilter?.value || "pending";
  const keyword = cleanText(dispatchSearch?.value).toLowerCase();

  return shipments.filter((item) => {
    const isLoaded = normalizeShipmentLoadedFlag(item.isLoaded);
    const isPendingCandidate = item.type === "receive" && !isLoaded;
    const isLoadedCandidate = isLoaded;

    if (!isPendingCandidate && !isLoadedCandidate) {
      return false;
    }

    if (status === "pending" && !isPendingCandidate) {
      return false;
    }
    if (status === "loaded" && !isLoadedCandidate) {
      return false;
    }
    if (status === "all" && !isPendingCandidate && !isLoadedCandidate) {
      return false;
    }

    if (!keyword) {
      return true;
    }

    const bag =
      `${item.date} ${item.serialNo} ${item.manufacturer} ${item.customer} ${item.product} ${item.note} ${renderDispatchLoadStatusText(item)} ${renderShipmentDestinationText(item)} ${cleanText(
        item.dispatchDate
      )} ${cleanText(item.dispatchTruckNo)} ${cleanText(item.dispatchPlateNo)} ${cleanText(item.dispatchDriver)} ${cleanText(
        item.dispatchContactName
      )}`.toLowerCase();
    return bag.includes(keyword);
  });
}

function renderDispatchLoadStatus(item) {
  const isLoaded = normalizeShipmentLoadedFlag(item.isLoaded);
  const statusClass = isLoaded ? "is-loaded" : "is-pending";
  const statusText = isLoaded ? "已装车" : "待配车";
  return `<button type="button" class="shipment-load-toggle ${statusClass}" data-id="${escapeHtml(
    item.id
  )}" aria-label="切换装车状态">${statusText}</button>`;
}

function renderDispatchLoadStatusText(item) {
  return normalizeShipmentLoadedFlag(item.isLoaded) ? "已装车" : "待配车";
}

function getFilteredShipments() {
  const type = shipmentFilter.value;
  const keyword = cleanText(shipmentSearch.value).toLowerCase();
  const manufacturerKeyword = cleanText(shipmentManufacturerFilter?.value).toLowerCase();
  const { from, to } = getDateRangeFromInputs(shipmentFromDateInput, shipmentToDateInput);

  return shipments
    .filter((item) => {
      if (type !== "all" && item.type !== type) {
        return false;
      }

      if (!isDateInRange(item.date, from, to)) {
        return false;
      }

      if (manufacturerKeyword) {
        const manufacturerText = cleanText(item.manufacturer).toLowerCase();
        if (!manufacturerText.includes(manufacturerKeyword)) {
          return false;
        }
      }

      if (!keyword) {
        return true;
      }

      const bag =
        `${item.serialNo} ${TYPE_LABELS[item.type] || ""} ${item.manufacturer} ${item.customer} ${item.product} ${formatMeasureUnit(
          item.measureUnit
        )} ${item.transit} ${item.delivery} ${item.pickup} ${item.note} ${renderShipmentLoadStatusText(item)} ${renderShipmentPaymentStatusText(
          item
        )} ${renderShipmentDestinationText(item)} ${cleanText(item.dispatchDate)} ${cleanText(item.dispatchTruckNo)} ${cleanText(
          item.dispatchPlateNo
        )} ${cleanText(item.dispatchDriver)} ${cleanText(item.dispatchContactName)}`.toLowerCase();
      return bag.includes(keyword);
    });
}

function getFilteredFinanceRecords() {
  const type = financeFilter.value;
  const month = financeMonth.value;
  const keyword = cleanText(financeSearch.value).toLowerCase();
  const { from, to } = getDateRangeFromInputs(financeFromDateInput, financeToDateInput);

  return financeRecords
    .filter((item) => {
      if (type !== "all" && item.type !== type) {
        return false;
      }

      if (month && !item.date.startsWith(month)) {
        return false;
      }

      if (!isDateInRange(item.date, from, to)) {
        return false;
      }

      if (!keyword) {
        return true;
      }

      const bag = `${item.summary || item.category} ${item.note}`.toLowerCase();
      return bag.includes(keyword);
    });
}

function changeMainPage(table, offset) {
  if (table === "shipment") {
    shipmentPage = Math.max(1, shipmentPage + offset);
    renderShipmentTable();
    return;
  }

  if (table === "dispatch") {
    dispatchPage = Math.max(1, dispatchPage + offset);
    renderDispatchTable();
    return;
  }

  if (table === "finance") {
    financePage = Math.max(1, financePage + offset);
    renderFinanceTable();
    return;
  }

  if (table === "status") {
    statusPage = Math.max(1, statusPage + offset);
    renderStatusSection();
  }
}

function changeStatusActionRecordPage(offset) {
  statusActionRecordPage = Math.max(1, statusActionRecordPage + offset);
  renderStatusSection();
}

function updateMainPagination(table, page, totalPages, totalRows) {
  let prevBtn = null;
  let nextBtn = null;
  let pageInfo = null;

  if (table === "shipment") {
    prevBtn = shipmentPrevBtn;
    nextBtn = shipmentNextBtn;
    pageInfo = shipmentPageInfo;
  } else if (table === "dispatch") {
    prevBtn = dispatchPrevBtn;
    nextBtn = dispatchNextBtn;
    pageInfo = dispatchPageInfo;
  } else if (table === "status") {
    prevBtn = statusPrevBtn;
    nextBtn = statusNextBtn;
    pageInfo = statusPageInfo;
  } else {
    prevBtn = financePrevBtn;
    nextBtn = financeNextBtn;
    pageInfo = financePageInfo;
  }

  if (!prevBtn || !nextBtn || !pageInfo) {
    return;
  }

  if (totalRows === 0) {
    pageInfo.textContent = "暂无数据";
    prevBtn.disabled = true;
    nextBtn.disabled = true;
    return;
  }

  pageInfo.textContent = `第 ${page} / ${totalPages} 页（共 ${formatNumber(totalRows)} 条）`;
  prevBtn.disabled = page <= 1;
  nextBtn.disabled = page >= totalPages;
}

async function handleDeleteClick(event) {
  const btn = event.target.closest(".delete-btn");
  if (!btn) {
    return;
  }

  const { id, table } = btn.dataset;
  if (!id || !table) {
    return;
  }

  const confirmed = await confirmDialog("确定删除这条记录吗？", {
    title: "确认删除",
    confirmText: "确认删除",
    cancelText: "取消"
  });
  if (!confirmed) {
    return;
  }

  queueDeleteWithUndo(table, id);
}

function queueDeleteWithUndo(table, id) {
  if (table === "shipment") {
    const index = shipments.findIndex((item) => item.id === id);
    if (index < 0) {
      return;
    }
    const [record] = shipments.splice(index, 1);
    let dispatchRecordsSnapshot = null;
    const beforeDispatchRecords = JSON.stringify(dispatchRecords);
    removeShipmentFromDispatchRecords(record.id);
    const afterDispatchRecords = JSON.stringify(dispatchRecords);
    if (beforeDispatchRecords !== afterDispatchRecords) {
      dispatchRecordsSnapshot = safeCloneForUndo(JSON.parse(beforeDispatchRecords));
      writeStorage(STORAGE_KEYS.dispatchRecords, dispatchRecords);
    }
    writeStorage(STORAGE_KEYS.shipments, shipments);
    const financeChanged = syncFinanceIncomeFromPaidShipments();
    renderShipmentSection();
    if (financeChanged) {
      renderFinanceSection();
    }
    openUndoToast({ table, record, index, dispatchRecordsSnapshot });
    return;
  }

  if (table === "finance") {
    const index = financeRecords.findIndex((item) => item.id === id);
    if (index < 0) {
      return;
    }
    const [record] = financeRecords.splice(index, 1);
    let linkedShipmentSnapshot = null;
    let shipmentChanged = false;

    if (isShipmentPaymentAutoFinanceRecord(record)) {
      const shipmentIdBySerialNo = new Map();
      shipments.forEach((item) => {
        const shipmentId = cleanText(item?.id);
        const serialNo = cleanText(item?.serialNo);
        if (!shipmentId || !serialNo || shipmentIdBySerialNo.has(serialNo)) {
          return;
        }
        shipmentIdBySerialNo.set(serialNo, shipmentId);
      });

      const linkedShipmentId = resolveShipmentIdForPaymentFinanceRecord(record, shipmentIdBySerialNo);
      const linkedShipment = shipments.find((item) => cleanText(item?.id) === linkedShipmentId);
      if (linkedShipment) {
        linkedShipmentSnapshot = {
          shipmentId: cleanText(linkedShipment.id),
          isPaid: normalizeShipmentPaidFlag(linkedShipment.isPaid),
          paidAt: Math.max(0, toFiniteNumber(linkedShipment.paidAt)),
        };
        if (normalizeShipmentPaidFlag(linkedShipment.isPaid) || Math.max(0, toFiniteNumber(linkedShipment.paidAt)) > 0) {
          linkedShipment.isPaid = false;
          linkedShipment.paidAt = 0;
          shipmentChanged = true;
        }
      }
    }

    if (shipmentChanged) {
      writeStorage(STORAGE_KEYS.shipments, shipments);
    }
    writeStorage(STORAGE_KEYS.finance, financeRecords);
    syncFinanceIncomeFromPaidShipments();
    if (shipmentChanged) {
      renderShipmentSection();
    }
    renderFinanceSection();
    openUndoToast({ table, record, index, linkedShipmentSnapshot });
  }
}

function openUndoToast(payload) {
  clearUndoState();
  pendingDeletion = {
    ...payload,
    expiresAt: Date.now() + DELETE_UNDO_MS,
  };

  if (undoToastText) {
    undoToastText.textContent = payload.table === "shipment" ? "已删除 1 条收发记录" : "已删除 1 条收支记录";
  }
  if (undoToast) {
    undoToast.classList.remove("is-hidden");
  }

  updateUndoCountdown();
  pendingDeletion.intervalId = window.setInterval(updateUndoCountdown, 250);
  pendingDeletion.timeoutId = window.setTimeout(finalizeUndoState, DELETE_UNDO_MS);
}

function updateUndoCountdown() {
  if (!pendingDeletion || !undoCountdownEl) {
    return;
  }
  const seconds = Math.max(0, Math.ceil((pendingDeletion.expiresAt - Date.now()) / 1000));
  undoCountdownEl.textContent = `${seconds}s`;
}

function undoPendingDeletion() {
  if (!pendingDeletion) {
    return;
  }

  const { table, record, index, dispatchRecordsSnapshot } = pendingDeletion;
  if (table === "shipment") {
    shipments.splice(Math.min(index, shipments.length), 0, record);
    writeStorage(STORAGE_KEYS.shipments, shipments);
    if (Array.isArray(dispatchRecordsSnapshot)) {
      dispatchRecords = safeCloneForUndo(dispatchRecordsSnapshot) || dispatchRecordsSnapshot;
      writeStorage(STORAGE_KEYS.dispatchRecords, dispatchRecords);
    }
    const financeChanged = syncFinanceIncomeFromPaidShipments();
    renderShipmentSection();
    if (financeChanged) {
      renderFinanceSection();
    }
  } else if (table === "finance") {
    const linkedShipmentSnapshot = pendingDeletion?.linkedShipmentSnapshot;
    let shipmentRestored = false;
    if (linkedShipmentSnapshot && cleanText(linkedShipmentSnapshot.shipmentId)) {
      const targetShipment = shipments.find((item) => cleanText(item?.id) === cleanText(linkedShipmentSnapshot.shipmentId));
      if (targetShipment) {
        targetShipment.isPaid = normalizeShipmentPaidFlag(linkedShipmentSnapshot.isPaid);
        targetShipment.paidAt = targetShipment.isPaid ? Math.max(0, toFiniteNumber(linkedShipmentSnapshot.paidAt)) : 0;
        shipmentRestored = true;
      }
    }
    financeRecords.splice(Math.min(index, financeRecords.length), 0, record);
    if (shipmentRestored) {
      writeStorage(STORAGE_KEYS.shipments, shipments);
    }
    writeStorage(STORAGE_KEYS.finance, financeRecords);
    const financeChanged = syncFinanceIncomeFromPaidShipments();
    if (shipmentRestored) {
      renderShipmentSection();
    }
    renderFinanceSection();
  }

  clearUndoState();
}

function finalizeUndoState() {
  clearUndoState();
}

function clearUndoState() {
  if (pendingDeletion?.timeoutId) {
    clearTimeout(pendingDeletion.timeoutId);
  }
  if (pendingDeletion?.intervalId) {
    clearInterval(pendingDeletion.intervalId);
  }
  pendingDeletion = null;

  if (undoToast) {
    undoToast.classList.add("is-hidden");
  }
}

function safeCloneForUndo(value) {
  try {
    return JSON.parse(JSON.stringify(value));
  } catch {
    return null;
  }
}

const SHIPMENT_EXPORT_HEADERS = [
  "日期",
  "序号",
  "类型",
  "厂家",
  "客户",
  "产品",
  "计量值",
  "计量单位",
  "单价(元)",
  "金额(元)",
  "单位重量(KG)",
  "单位立方(m3)",
  "总重量(KG)",
  "总立方",
  "中转",
  "送货",
  "提货",
  "回扣(元)",
  "备注",
  "装车状态",
  "去向",
  "收款状态",
];

function buildShipmentExportRows(rows) {
  return rows.map((item) => [
    item.date,
    item.serialNo,
    TYPE_LABELS[item.type],
    item.manufacturer,
    item.customer,
    item.product,
    formatMeasureValue(item.quantity, item.measureUnit),
    formatMeasureUnit(item.measureUnit),
    toFiniteNumber(item.unitPrice).toFixed(2),
    toFiniteNumber(item.amount).toFixed(2),
    formatDecimal(getShipmentUnitWeight(item)),
    formatDecimal(getShipmentUnitCubic(item)),
    formatDecimal(getShipmentTotalWeight(item)),
    formatDecimal(getShipmentTotalCubic(item)),
    item.transit,
    item.delivery,
    item.pickup,
    toFiniteNumber(item.rebate).toFixed(2),
    item.note,
    renderShipmentLoadStatusText(item),
    renderShipmentDestinationText(item),
    renderShipmentPaymentStatusText(item),
  ]);
}

function sanitizeExportPart(value, fallback = "全部") {
  const text = cleanText(value);
  if (!text) {
    return fallback;
  }
  const normalized = text
    .replace(/[\\/:*?"<>|]+/g, "-")
    .replace(/\s+/g, "")
    .replace(/-+/g, "-")
    .replace(/^[-_.]+|[-_.]+$/g, "");
  return normalized || fallback;
}

function formatDateRangePart(from, to) {
  if (from && to) {
    return `${from}_${to}`;
  }
  if (from) {
    return `${from}_至今`;
  }
  if (to) {
    return `截至_${to}`;
  }
  return "全部日期";
}

const FINANCE_EXPORT_HEADERS = ["日期", "摘要", "收入(元)", "支出(元)", "余额(元)", "备注"];

function buildFinanceExportRows(rows, balanceById) {
  return rows.map((item) => {
    const amount = toFiniteNumber(item.amount);
    const income = item.type === "income" ? amount : 0;
    const expense = item.type === "expense" ? amount : 0;
    const balance = toFiniteNumber(balanceById.get(item.id));
    return [
      item.date || "",
      resolveFinanceSummary(item),
      Number(income.toFixed(2)),
      Number(expense.toFixed(2)),
      Number(balance.toFixed(2)),
      resolveFinanceNote(item, { fallback: "" }),
    ];
  });
}

function buildSelectedFinanceExportContext() {
  const filteredRows = sortFinanceRows(getFilteredFinanceRecords());
  if (!filteredRows.length) {
    alert("当前筛选条件下没有可导出的收支记录。");
    return null;
  }

  pruneSelectedFinanceExportIds(filteredRows);
  const rows = filteredRows.filter((item) => selectedFinanceIdsForExport.has(cleanText(item.id)));
  if (!rows.length) {
    alert("请先在总账表格中勾选需要导出的记录。");
    return null;
  }

  const balanceById = buildFinanceBalanceMap(filteredRows);
  const typeValue = cleanText(financeFilter?.value || "all");
  const typeLabelMap = {
    all: "全部",
    income: "收入",
    expense: "支出",
  };
  const typePart = sanitizeExportPart(typeLabelMap[typeValue] || "全部");
  const monthPart = sanitizeExportPart(financeMonth?.value || "", "");
  const keywordPart = sanitizeExportPart(financeSearch?.value || "", "");
  const { from, to } = getDateRangeFromInputs(financeFromDateInput, financeToDateInput);
  const datePart = sanitizeExportPart(formatDateRangePart(from, to), "全部日期");
  const todayPart = dateValue(new Date());
  const filenameParts = ["收支总账", "勾选", typePart];
  if (monthPart) {
    filenameParts.push(monthPart);
  }
  if (keywordPart) {
    filenameParts.push(keywordPart);
  }
  filenameParts.push(datePart, `已选${rows.length}条`, todayPart);

  return {
    rows,
    balanceById,
    headers: FINANCE_EXPORT_HEADERS,
    filename: `${filenameParts.join("_")}.xlsx`,
  };
}

function handleFinanceExportSelectedClick() {
  const context = buildSelectedFinanceExportContext();
  if (!context) {
    return;
  }
  openFinanceExportPreviewModal(context);
}

function exportSelectedFinanceExcel(context) {
  if (typeof XLSX === "undefined") {
    alert("未加载 Excel 导出库，请刷新页面后重试。");
    return;
  }
  const rows = buildFinanceExportRows(context.rows, context.balanceById);
  const totalIncome = rows.reduce((sum, row) => sum + toFiniteNumber(row[2]), 0);
  const totalExpense = rows.reduce((sum, row) => sum + toFiniteNumber(row[3]), 0);
  const totalBalance = totalIncome - totalExpense;
  const aoa = [
    context.headers,
    ...rows,
    ["合计", `记录 ${context.rows.length} 条`, Number(totalIncome.toFixed(2)), Number(totalExpense.toFixed(2)), Number(totalBalance.toFixed(2)), ""],
  ];
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet(aoa);
  sheet["!cols"] = [{ wch: 12 }, { wch: 32 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 30 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "收支勾选导出");
  XLSX.writeFile(workbook, context.filename);
}

function renderFinanceExportPreview(context) {
  if (!financeExportPreviewBody) {
    return;
  }
  if (!context || !Array.isArray(context.rows) || !context.rows.length) {
    financeExportPreviewBody.innerHTML = '<tr><td colspan="6" class="empty">暂无可预览数据</td></tr>';
    return;
  }

  financeExportPreviewBody.innerHTML = context.rows
    .map((item) => {
      const amount = toFiniteNumber(item.amount);
      const income = item.type === "income" ? amount : 0;
      const expense = item.type === "expense" ? amount : 0;
      const balance = toFiniteNumber(context.balanceById.get(item.id));
      return `
      <tr>
        <td>${escapeHtml(item.date || "-")}</td>
        <td class="finance-ledger-summary">${escapeHtml(resolveFinanceSummary(item))}</td>
        <td>${income > 0 ? toFiniteNumber(income).toFixed(2) : "-"}</td>
        <td>${expense > 0 ? toFiniteNumber(expense).toFixed(2) : "-"}</td>
        <td>${toFiniteNumber(balance).toFixed(2)}</td>
        <td>${escapeHtml(resolveFinanceNote(item, { fallback: "-" }))}</td>
      </tr>`;
    })
    .join("");
}

function openFinanceExportPreviewModal(context) {
  if (!financeExportPreviewModal) {
    return;
  }
  pendingFinanceExportContext = context;
  renderFinanceExportPreview(context);
  if (financeExportPreviewTip) {
    financeExportPreviewTip.textContent = `将导出 ${context.rows.length} 条记录，文件名：${context.filename}`;
  }
  financeExportPreviewModal.classList.add("is-open");
  syncModalBodyState();
}

function closeFinanceExportPreviewModal() {
  if (!financeExportPreviewModal) {
    return;
  }
  financeExportPreviewModal.classList.remove("is-open");
  pendingFinanceExportContext = null;
  syncModalBodyState();
}

function handleFinanceExportPreviewConfirm() {
  if (!pendingFinanceExportContext) {
    return;
  }
  const context = pendingFinanceExportContext;
  exportSelectedFinanceExcel(context);
  closeFinanceExportPreviewModal();
}

const STORAGE_PACKAGE_KEY_ORDER = [
  STORAGE_KEYS.shipments,
  STORAGE_KEYS.finance,
  STORAGE_KEYS.dispatchRecords,
  STORAGE_KEYS.homeSettleActionRecords,
  STORAGE_KEYS.catalogProfiles,
  STORAGE_KEYS.shipmentRetain,
  STORAGE_KEYS.financeRetain,
];

function getStoragePackageDefaultValue(key) {
  if (key === STORAGE_KEYS.shipmentRetain || key === STORAGE_KEYS.financeRetain) {
    return {};
  }
  return [];
}

function readRawStorageValue(key, fallback) {
  try {
    const raw = localStorage.getItem(key);
    if (!raw) {
      return cloneForStorage(fallback);
    }
    const parsed = JSON.parse(raw);
    return parsed == null ? cloneForStorage(fallback) : parsed;
  } catch {
    return cloneForStorage(fallback);
  }
}

function countStorageRecords(value) {
  if (Array.isArray(value)) {
    return value.length;
  }
  if (value && typeof value === "object") {
    return Object.keys(value).length;
  }
  return 0;
}

function buildStoragePackageSnapshot() {
  const datasets = {};
  STORAGE_PACKAGE_KEY_ORDER.forEach((key) => {
    datasets[key] = readRawStorageValue(key, getStoragePackageDefaultValue(key));
  });
  return {
    version: STORAGE_PACKAGE_VERSION,
    exportedAt: new Date().toISOString(),
    source: "顺源物流管理系统",
    datasets,
  };
}

function buildStoragePackageExportFilename() {
  return `${MANUAL_STORAGE_FILENAME_PREFIX}.json`;
}

function createStoragePackageText(payload) {
  return `${JSON.stringify(payload, null, 2)}\n`;
}

function downloadJson(filename, payload) {
  const text = createStoragePackageText(payload);
  const blob = new Blob([text], { type: "application/json;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}

async function saveStoragePackageToFile(filename, payload) {
  const text = createStoragePackageText(payload);
  const pickerSupported =
    typeof window !== "undefined" &&
    typeof window.showSaveFilePicker === "function" &&
    window.isSecureContext;

  if (pickerSupported) {
    try {
      const handle = await window.showSaveFilePicker({
        suggestedName: filename,
        types: [
          {
            description: "JSON 数据包",
            accept: {
              "application/json": [".json"],
            },
          },
        ],
      });
      const writable = await handle.createWritable();
      await writable.write(text);
      await writable.close();
      return { saved: true, method: "picker", fileName: handle?.name || filename };
    } catch (error) {
      if (error?.name === "AbortError") {
        return { saved: false, canceled: true };
      }
    }
  }

  downloadJson(filename, payload);
  return { saved: true, method: "download", fileName: filename };
}

async function handleStorageExportPackageClick() {
  const snapshot = buildStoragePackageSnapshot();
  const filename = buildStoragePackageExportFilename();
  const result = await saveStoragePackageToFile(filename, snapshot);
  if (!result?.saved) {
    return;
  }
  manualSaveDirty = false;
  manualSaveLastAt = Date.now();
  persistManualSaveState();
  updateStorageModeBadge();
}

function handleStorageImportPackageClick() {
  if (!storageImportPackageFileInput) {
    alert("读取控件未找到，请刷新页面后重试。");
    return;
  }
  storageImportPackageFileInput.value = "";
  storageImportPackageFileInput.click();
}

async function handleStorageImportPackageFileChange(event) {
  const input = event?.currentTarget || storageImportPackageFileInput;
  const file = input?.files?.[0];
  if (!file) {
    return;
  }
  try {
    const context = await parseStorageImportPackage(file);
    if (!context) {
      return;
    }
    openStorageImportPreviewModal(context);
  } finally {
    if (input) {
      input.value = "";
    }
  }
}

async function parseStorageImportPackage(file) {
  let parsed = null;
  let normalizedText = "";
  try {
    const text = await file.text();
    normalizedText = normalizeStoragePackageText(text);
    parsed = JSON.parse(normalizedText);
  } catch (error) {
    const fileName = cleanText(file?.name) || "未知文件";
    const suffix = getStorageReadErrorHint(fileName, error);
    alert(`数据包读取失败：${suffix}`);
    return null;
  }

  const rawDatasets = buildImportDatasetsFromParsed(parsed, file?.name);
  if (!rawDatasets || typeof rawDatasets !== "object" || Array.isArray(rawDatasets)) {
    alert("数据包格式无效，请选择“手动保存数据包”导出的 JSON 文件。");
    return null;
  }

  const importValues = {};
  const rows = STORAGE_PACKAGE_KEY_ORDER.map((key) => {
    const hasValue = Object.prototype.hasOwnProperty.call(rawDatasets, key);
    const currentValue = readRawStorageValue(key, getStoragePackageDefaultValue(key));
    const incomingValue = hasValue ? rawDatasets[key] : undefined;
    const incomingCount = hasValue ? countStorageRecords(incomingValue) : 0;
    const currentCount = countStorageRecords(currentValue);

    if (hasValue) {
      importValues[key] = cloneForStorage(incomingValue);
    }

    return {
      key,
      label: STORAGE_PACKAGE_LABELS[key] || key,
      hasValue,
      currentCount,
      incomingCount,
      message: hasValue ? "将覆盖当前数据" : "数据包未包含，保持现有",
    };
  });

  const importCount = Object.keys(importValues).length;
  if (!importCount) {
    alert("数据包中没有可导入的数据项。请确认选择了本系统导出的数据包，或选择正确的单项数据文件。");
    return null;
  }

  return {
    fileName: cleanText(file?.name) || `${MANUAL_STORAGE_FILENAME_PREFIX}.json`,
    packageVersion: cleanText(parsed.version) || "未知",
    exportedAt: cleanText(parsed.exportedAt),
    rows,
    importValues,
  };
}

function normalizeStoragePackageText(text) {
  if (typeof text !== "string") {
    return "";
  }
  let normalized = text;
  if (normalized.charCodeAt(0) === 0xfeff) {
    normalized = normalized.slice(1);
  }
  return normalized.trim();
}

function getStorageReadErrorHint(fileName, error) {
  const lower = String(fileName || "").toLowerCase();
  if (!lower.endsWith(".json")) {
    return "请选择 JSON 文件（不是 Excel/CSV）。";
  }
  if (error && typeof error === "object" && error.name === "SyntaxError") {
    return "文件内容不是合法 JSON，可能不是本系统导出的数据包。";
  }
  return "请确认文件未损坏，并且是本系统导出的数据包。";
}

function inferStorageKeyFromFileName(fileName = "") {
  const name = String(fileName || "").toLowerCase();
  if (!name) {
    return "";
  }
  if (name.includes("logistics_shipments_v1")) {
    return STORAGE_KEYS.shipments;
  }
  if (name.includes("logistics_finance_v1")) {
    return STORAGE_KEYS.finance;
  }
  if (name.includes("logistics_dispatch_records_v1")) {
    return STORAGE_KEYS.dispatchRecords;
  }
  if (name.includes("logistics_home_settle_action_records_v1")) {
    return STORAGE_KEYS.homeSettleActionRecords;
  }
  if (name.includes("logistics_catalog_profiles_v1")) {
    return STORAGE_KEYS.catalogProfiles;
  }
  if (name.includes("logistics_shipment_retain_v1")) {
    return STORAGE_KEYS.shipmentRetain;
  }
  if (name.includes("logistics_finance_retain_v1")) {
    return STORAGE_KEYS.financeRetain;
  }
  return "";
}

function normalizeLegacyDatasetKeys(raw = {}) {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) {
    return {};
  }
  const mapped = {};
  const keyAliases = {
    shipments: STORAGE_KEYS.shipments,
    finance: STORAGE_KEYS.finance,
    dispatchRecords: STORAGE_KEYS.dispatchRecords,
    homeSettleActionRecords: STORAGE_KEYS.homeSettleActionRecords,
    catalogProfiles: STORAGE_KEYS.catalogProfiles,
    shipmentRetain: STORAGE_KEYS.shipmentRetain,
    financeRetain: STORAGE_KEYS.financeRetain,
    [STORAGE_KEYS.shipments]: STORAGE_KEYS.shipments,
    [STORAGE_KEYS.finance]: STORAGE_KEYS.finance,
    [STORAGE_KEYS.dispatchRecords]: STORAGE_KEYS.dispatchRecords,
    [STORAGE_KEYS.homeSettleActionRecords]: STORAGE_KEYS.homeSettleActionRecords,
    [STORAGE_KEYS.catalogProfiles]: STORAGE_KEYS.catalogProfiles,
    [STORAGE_KEYS.shipmentRetain]: STORAGE_KEYS.shipmentRetain,
    [STORAGE_KEYS.financeRetain]: STORAGE_KEYS.financeRetain,
  };

  Object.entries(raw).forEach(([key, value]) => {
    const target = keyAliases[key];
    if (target) {
      mapped[target] = value;
    }
  });
  return mapped;
}

function buildImportDatasetsFromParsed(parsed, fileName = "") {
  if (Array.isArray(parsed)) {
    const inferredKey = inferStorageKeyFromFileName(fileName);
    if (!inferredKey) {
      return null;
    }
    return { [inferredKey]: parsed };
  }
  if (!parsed || typeof parsed !== "object") {
    return null;
  }

  const candidates = [];
  if (parsed.datasets && typeof parsed.datasets === "object" && !Array.isArray(parsed.datasets)) {
    candidates.push(parsed.datasets);
  }
  if (parsed.data && typeof parsed.data === "object" && !Array.isArray(parsed.data)) {
    if (parsed.data.datasets && typeof parsed.data.datasets === "object" && !Array.isArray(parsed.data.datasets)) {
      candidates.push(parsed.data.datasets);
    }
    candidates.push(parsed.data);
  }
  if (parsed.payload && typeof parsed.payload === "object" && !Array.isArray(parsed.payload)) {
    if (
      parsed.payload.datasets &&
      typeof parsed.payload.datasets === "object" &&
      !Array.isArray(parsed.payload.datasets)
    ) {
      candidates.push(parsed.payload.datasets);
    }
    candidates.push(parsed.payload);
  }
  candidates.push(parsed);

  for (const candidate of candidates) {
    const normalized = normalizeLegacyDatasetKeys(candidate);
    if (Object.keys(normalized).length) {
      return normalized;
    }
  }
  return null;
}

function renderStorageImportPreview(context) {
  if (!storageImportPreviewBody) {
    return;
  }
  if (!context || !Array.isArray(context.rows) || !context.rows.length) {
    storageImportPreviewBody.innerHTML = '<tr><td colspan="3" class="empty">暂无可预览数据</td></tr>';
    return;
  }

  storageImportPreviewBody.innerHTML = context.rows
    .map((row) => {
      const countText = row.hasValue ? `新 ${row.incomingCount} / 现 ${row.currentCount}` : `现 ${row.currentCount}`;
      const rowClass = row.hasValue ? "shipment-import-row-ok" : "";
      return `
        <tr class="${rowClass}">
          <td>${escapeHtml(row.label)}</td>
          <td class="mono">${escapeHtml(countText)}</td>
          <td>${escapeHtml(row.message)}</td>
        </tr>
      `;
    })
    .join("");
}

function openStorageImportPreviewModal(context) {
  if (!storageImportPreviewModal) {
    return;
  }
  pendingStorageImportContext = context;
  renderStorageImportPreview(context);
  if (storageImportPreviewTip) {
    const exportedAtText = context.exportedAt ? formatDateTime(context.exportedAt) : "-";
    storageImportPreviewTip.textContent = `文件：${context.fileName} ｜版本：${context.packageVersion} ｜导出时间：${exportedAtText}。确认后将覆盖对应数据项。`;
  }
  storageImportPreviewModal.classList.add("is-open");
  syncModalBodyState();
}

function closeStorageImportPreviewModal() {
  if (!storageImportPreviewModal) {
    return;
  }
  storageImportPreviewModal.classList.remove("is-open");
  pendingStorageImportContext = null;
  syncModalBodyState();
}

async function handleStorageImportPreviewConfirm() {
  const context = pendingStorageImportContext;
  if (!context || !context.importValues || typeof context.importValues !== "object") {
    return;
  }

  Object.entries(context.importValues).forEach(([key, value]) => {
    writeStorage(key, value);
  });

  shipments = readStorage(STORAGE_KEYS.shipments);
  financeRecords = readStorage(STORAGE_KEYS.finance);
  dispatchRecords = readStorage(STORAGE_KEYS.dispatchRecords);
  homeSettleActionRecords = sanitizeHomeSettleActionRecords(readStorage(STORAGE_KEYS.homeSettleActionRecords));
  customCatalogProfiles = sanitizeCustomCatalogProfiles(readStorage(STORAGE_KEYS.catalogProfiles));

  selectedShipmentIdsForExport.clear();
  selectedFinanceIdsForExport.clear();
  selectedDispatchIdsForLoading.clear();
  confirmedDispatchIdsForLoading.clear();
  dispatchRecordPendingAddIds.clear();
  dispatchLoadStage = "select";
  activeDispatchRecordId = "";
  dispatchRecordEditMode = false;
  dispatchRecordAddKeyword = "";

  restoreRetainedFormValues();
  persistSanitizedState();
  await flushFileStorageWrites();
  manualSaveDirty = false;
  const importedAt = context.exportedAt ? Date.parse(context.exportedAt) : NaN;
  manualSaveLastAt = Number.isFinite(importedAt) ? importedAt : Date.now();
  persistManualSaveState();
  updateStorageModeBadge();
  renderAll();
  closeStorageImportPreviewModal();

  alert("数据包已读取并完成联动刷新。");
}

const SHIPMENT_IMPORT_FALLBACK_COLUMN_ORDER = [
  "date",
  "serialNo",
  "manufacturer",
  "customer",
  "product",
  "quantity",
  "measureUnit",
  "unitPrice",
  "amount",
  "weight",
  "cubic",
  "rebate",
  "transit",
  "delivery",
  "pickup",
  "note",
];

function handleShipmentBatchImportClick() {
  if (!shipmentBatchImportFileInput) {
    alert("导入控件未找到，请刷新页面后重试。");
    return;
  }
  shipmentBatchImportFileInput.value = "";
  shipmentBatchImportFileInput.click();
}

async function handleShipmentBatchImportFileChange(event) {
  const input = event?.currentTarget || shipmentBatchImportFileInput;
  const file = input?.files?.[0];
  if (!file) {
    return;
  }
  try {
    const context = await parseShipmentImportWorkbook(file);
    if (!context) {
      return;
    }
    openShipmentImportPreviewModal(context);
  } finally {
    if (input) {
      input.value = "";
    }
  }
}

async function parseShipmentImportWorkbook(file) {
  if (typeof XLSX === "undefined") {
    alert("未加载 Excel 解析库，请刷新页面后重试。");
    return null;
  }

  let workbook = null;
  try {
    const fileBuffer = await file.arrayBuffer();
    workbook = XLSX.read(fileBuffer, {
      type: "array",
      cellDates: true,
      raw: false,
    });
  } catch (error) {
    console.error("读取Excel失败:", error);
    alert("Excel 文件读取失败，请检查文件格式。");
    return null;
  }

  const firstSheetName = cleanText(workbook?.SheetNames?.[0]);
  if (!firstSheetName) {
    alert("Excel 文件没有可读取的工作表。");
    return null;
  }

  const sheet = workbook.Sheets[firstSheetName];
  if (!sheet) {
    alert("无法读取工作表内容。");
    return null;
  }

  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: "",
    blankrows: false,
    raw: false,
  });

  if (!Array.isArray(rows) || !rows.length) {
    alert("未读取到可导入的数据。");
    return null;
  }

  const context = buildShipmentImportPreviewContext(rows, file?.name || "导入文件", firstSheetName);
  if (!context.previewRows.length) {
    alert("未发现可导入行，请检查模板列名或数据内容。");
    return null;
  }
  return context;
}

function buildShipmentImportPreviewContext(rows, fileName, sheetName) {
  const headerRow = Array.isArray(rows[0]) ? rows[0] : [];
  const { isHeaderRow, fieldIndexMap } = resolveShipmentImportFieldIndexMap(headerRow);
  const startIndex = isHeaderRow ? 1 : 0;
  const previewRows = [];
  const today = dateValue(new Date());

  for (let rowIndex = startIndex; rowIndex < rows.length; rowIndex += 1) {
    const rawRow = Array.isArray(rows[rowIndex]) ? rows[rowIndex] : [];
    if (isShipmentImportRowBlank(rawRow)) {
      continue;
    }

    const importDraft = buildShipmentImportDraftFromRow(rawRow, fieldIndexMap, today);
    const issues = validateShipmentImportDraft(importDraft, rawRow);
    const valid = issues.length === 0;
    previewRows.push({
      rowNo: rowIndex + 1,
      valid,
      message: valid ? "可导入" : issues.join("；"),
      draft: importDraft,
    });
  }

  const validRows = previewRows.filter((item) => item.valid);
  const invalidRows = previewRows.filter((item) => !item.valid);
  return {
    fileName: cleanText(fileName) || "导入文件",
    sheetName: cleanText(sheetName) || "Sheet1",
    previewRows,
    validRows,
    invalidRows,
  };
}

function resolveShipmentImportFieldIndexMap(headerRow) {
  const map = new Map();
  let recognizedCount = 0;

  headerRow.forEach((cell, index) => {
    const field = resolveShipmentImportFieldName(cell);
    if (!field || map.has(field)) {
      return;
    }
    map.set(field, index);
    recognizedCount += 1;
  });

  const hasCoreHeaders =
    map.has("manufacturer") && map.has("customer") && map.has("product") && map.has("quantity");

  if (recognizedCount >= 3 || hasCoreHeaders) {
    return { isHeaderRow: true, fieldIndexMap: map };
  }

  const fallbackMap = new Map();
  SHIPMENT_IMPORT_FALLBACK_COLUMN_ORDER.forEach((field, index) => {
    fallbackMap.set(field, index);
  });
  return { isHeaderRow: false, fieldIndexMap: fallbackMap };
}

function resolveShipmentImportFieldName(headerValue) {
  const normalized = normalizeShipmentImportHeader(headerValue);
  if (!normalized) {
    return "";
  }

  if (normalized.includes("单位重量") || normalized.includes("重量kg") || normalized.includes("重量吨")) {
    return "weight";
  }
  if (normalized.includes("单位立方") || normalized.includes("立方m3") || normalized.includes("体积")) {
    return "cubic";
  }
  if (normalized.includes("计量单位") || normalized === "单位" || normalized === "计量") {
    return "measureUnit";
  }
  if (normalized.includes("计量值") || normalized.includes("件数") || normalized.includes("数量")) {
    return "quantity";
  }
  if (normalized.includes("单价")) {
    return "unitPrice";
  }
  if (normalized.includes("金额")) {
    return "amount";
  }
  if (normalized.includes("回扣")) {
    return "rebate";
  }
  if (normalized.includes("序号") || normalized.includes("单号") || normalized.includes("运单号") || normalized === "编号") {
    return "serialNo";
  }
  if (normalized.includes("日期")) {
    return "date";
  }
  if (normalized.includes("厂家") || normalized.includes("厂商") || normalized.includes("供货商") || normalized.includes("发货单位")) {
    return "manufacturer";
  }
  if (normalized.includes("客户") || normalized.includes("收货人") || normalized.includes("收货单位") || normalized.includes("门店")) {
    return "customer";
  }
  if (normalized.includes("产品") || normalized.includes("品名") || normalized.includes("货名")) {
    return "product";
  }
  if (normalized.includes("中转")) {
    return "transit";
  }
  if (normalized.includes("送货")) {
    return "delivery";
  }
  if (normalized.includes("提货")) {
    return "pickup";
  }
  if (normalized.includes("备注") || normalized.includes("说明")) {
    return "note";
  }
  return "";
}

function normalizeShipmentImportHeader(value) {
  return cleanText(value)
    .toLowerCase()
    .replace(/[（]/g, "(")
    .replace(/[）]/g, ")")
    .replace(/\s+/g, "")
    .replace(/[^\u4e00-\u9fffa-z0-9()]/g, "");
}

function isShipmentImportRowBlank(row) {
  if (!Array.isArray(row) || !row.length) {
    return true;
  }
  return row.every((cell) => !cleanText(cell));
}

function buildShipmentImportDraftFromRow(rawRow, fieldIndexMap, defaultDate) {
  const getCell = (field) => {
    const index = fieldIndexMap.get(field);
    if (!Number.isInteger(index) || index < 0) {
      return "";
    }
    return rawRow[index];
  };

  const rawDate = getCell("date");
  const parsedDate = parseDateFromAny(rawDate);
  const quantityRaw = parseOptionalNonNegative(getCell("quantity"));
  const unitPriceRaw = parseOptionalNonNegative(getCell("unitPrice"));
  const amountCell = getCell("amount");
  const amountText = cleanText(amountCell);
  const amountRaw = parseOptionalNonNegative(amountCell);
  const rebateRaw = parseOptionalNonNegative(getCell("rebate"));
  const weightRaw = parseOptionalNonNegative(getCell("weight"));
  const cubicRaw = parseOptionalNonNegative(getCell("cubic"));
  const normalizedQuantity = Number((Number.isFinite(quantityRaw) ? quantityRaw : 0).toFixed(3));
  const normalizedWeight = Number((Number.isFinite(weightRaw) ? weightRaw : 0).toFixed(3));
  const normalizedCubic = Number((Number.isFinite(cubicRaw) ? cubicRaw : 0).toFixed(3));
  const computedAmount = amountText
    ? amountRaw
    : normalizedQuantity * (Number.isFinite(unitPriceRaw) ? unitPriceRaw : 0);

  return {
    dateRaw: rawDate,
    date: parsedDate || defaultDate,
    serialNo: cleanText(getCell("serialNo")),
    manufacturer: cleanText(getCell("manufacturer")),
    customer: cleanText(getCell("customer")),
    product: cleanText(getCell("product")),
    measureUnit: normalizeMeasureUnit(getCell("measureUnit"), DEFAULT_MEASURE_UNIT) || DEFAULT_MEASURE_UNIT,
    quantity: normalizedQuantity,
    unitPrice: unitPriceRaw,
    amount: computedAmount,
    rebate: rebateRaw,
    weight: normalizedWeight,
    cubic: normalizedCubic,
    transit: cleanText(getCell("transit")),
    delivery: cleanText(getCell("delivery")),
    pickup: cleanText(getCell("pickup")),
    note: cleanText(getCell("note")),
    amountText,
    rawValues: {
      quantity: getCell("quantity"),
      unitPrice: getCell("unitPrice"),
      amount: getCell("amount"),
      rebate: getCell("rebate"),
      weight: getCell("weight"),
      cubic: getCell("cubic"),
    },
  };
}

function validateShipmentImportDraft(draft) {
  const issues = [];
  const hasDateInput = cleanText(draft.dateRaw);
  if (hasDateInput && !parseDateFromAny(draft.dateRaw)) {
    issues.push("日期格式无效");
  }

  if (!Number.isFinite(draft.quantity) || draft.quantity <= 0) {
    issues.push("计量值必须大于0");
  }
  if (!Number.isFinite(draft.unitPrice)) {
    issues.push("单价格式无效");
  }
  if (!Number.isFinite(draft.amount) || draft.amount < 0) {
    issues.push("金额格式无效");
  }
  if (!Number.isFinite(draft.rebate)) {
    issues.push("回扣格式无效");
  }
  if (!Number.isFinite(draft.weight)) {
    issues.push("单位重量格式无效");
  }
  if (!Number.isFinite(draft.cubic)) {
    issues.push("单位立方格式无效");
  }

  if (!draft.manufacturer && !draft.customer && !draft.product) {
    issues.push("厂家/客户/产品至少填写一项");
  }
  return issues;
}

function buildShipmentImportRecordFromDraft(draft) {
  return {
    id: makeId(),
    type: "receive",
    date: draft.date || dateValue(new Date()),
    serialNo: cleanText(draft.serialNo),
    manufacturer: cleanText(draft.manufacturer),
    customer: cleanText(draft.customer),
    product: cleanText(draft.product),
    measureUnit: normalizeMeasureUnit(draft.measureUnit, DEFAULT_MEASURE_UNIT) || DEFAULT_MEASURE_UNIT,
    quantity: Number(toFiniteNumber(draft.quantity).toFixed(3)),
    unitPrice: toFiniteNumber(draft.unitPrice),
    amount: toFiniteNumber(draft.amount),
    rebate: toFiniteNumber(draft.rebate),
    weight: Number(toFiniteNumber(draft.weight).toFixed(3)),
    cubic: Number(toFiniteNumber(draft.cubic).toFixed(3)),
    transit: cleanText(draft.transit),
    delivery: cleanText(draft.delivery),
    pickup: cleanText(draft.pickup),
    note: cleanText(draft.note),
    isLoaded: false,
    loadedAt: 0,
    dispatchDate: "",
    dispatchTruckNo: "",
    dispatchPlateNo: "",
    dispatchDriver: "",
    dispatchContactName: "",
    dispatchDestination: "",
    destination: "",
    homeSettleStatus: "",
    homeSettleStatusUpdatedAt: 0,
    homeSettleRemindedAt: 0,
    homeSettleSnoozeUntil: 0,
    homeSettleLastAction: "",
    homeSettleLastActionAt: 0,
    isPaid: false,
    paidAt: 0,
    createdAt: Date.now(),
  };
}

function renderShipmentImportPreview(context) {
  if (!shipmentImportPreviewBody) {
    return;
  }
  if (!context || !Array.isArray(context.previewRows) || !context.previewRows.length) {
    shipmentImportPreviewBody.innerHTML = '<tr><td colspan="14" class="empty">暂无可预览数据</td></tr>';
    return;
  }

  shipmentImportPreviewBody.innerHTML = context.previewRows
    .map((item) => {
      const draft = item.draft || {};
      const rowClass = item.valid ? "shipment-import-row-ok" : "shipment-import-row-error";
      return `
      <tr class="${rowClass}">
        <td>${escapeHtml(String(item.rowNo || "-"))}</td>
        <td>${escapeHtml(draft.date || "-")}</td>
        <td>${escapeHtml(draft.serialNo || "-")}</td>
        <td>${escapeHtml(draft.manufacturer || "-")}</td>
        <td>${escapeHtml(draft.customer || "-")}</td>
        <td>${escapeHtml(draft.product || "-")}</td>
        <td>${escapeHtml(formatMeasureValue(draft.quantity, draft.measureUnit))}</td>
        <td>${escapeHtml(formatMeasureUnit(draft.measureUnit))}</td>
        <td>${toFiniteNumber(draft.unitPrice).toFixed(2)}</td>
        <td>${toFiniteNumber(draft.amount).toFixed(2)}</td>
        <td>${formatDecimal(toFiniteNumber(draft.weight))}</td>
        <td>${formatDecimal(toFiniteNumber(draft.cubic))}</td>
        <td>${escapeHtml(draft.note || "-")}</td>
        <td>${escapeHtml(item.message || "-")}</td>
      </tr>`;
    })
    .join("");
}

function openShipmentImportPreviewModal(context) {
  if (!shipmentImportPreviewModal) {
    return;
  }
  pendingShipmentImportContext = context;
  renderShipmentImportPreview(context);
  if (shipmentImportPreviewTip) {
    shipmentImportPreviewTip.textContent = [
      `文件：${context.fileName} / 工作表：${context.sheetName}`,
      `共 ${context.previewRows.length} 行，可导入 ${context.validRows.length} 行，错误 ${context.invalidRows.length} 行。`,
      "请确认预览内容无误后再点击“确认导入”。",
    ].join(" ");
  }
  if (shipmentImportPreviewConfirmBtn) {
    shipmentImportPreviewConfirmBtn.disabled = context.validRows.length === 0;
  }
  shipmentImportPreviewModal.classList.add("is-open");
  syncModalBodyState();
}

function closeShipmentImportPreviewModal() {
  if (!shipmentImportPreviewModal) {
    return;
  }
  shipmentImportPreviewModal.classList.remove("is-open");
  pendingShipmentImportContext = null;
  if (shipmentImportPreviewConfirmBtn) {
    shipmentImportPreviewConfirmBtn.disabled = false;
  }
  syncModalBodyState();
}

function handleShipmentImportPreviewConfirm() {
  const context = pendingShipmentImportContext;
  if (!context) {
    return;
  }
  if (!context.validRows.length) {
    alert("当前没有可导入的有效数据。");
    return;
  }

  context.validRows.forEach((item) => {
    const record = buildShipmentImportRecordFromDraft(item.draft || {});
    shipments.push(record);
    upsertCustomCatalogProfile(record);
  });

  writeStorage(STORAGE_KEYS.shipments, shipments);
  resetShipmentFormAfterSubmit();
  renderShipmentSection();
  closeShipmentImportPreviewModal();

  const failedCount = context.invalidRows.length;
  const message = failedCount
    ? `批量导入完成：成功 ${context.validRows.length} 条，失败 ${failedCount} 条。`
    : `批量导入完成：成功 ${context.validRows.length} 条。`;
  alert(message);
}

function buildSelectedShipmentExportContext() {
  const filteredRows = sortShipmentRows(getFilteredShipments());
  if (!filteredRows.length) {
    alert("当前筛选条件下没有可导出的记录。");
    return null;
  }

  pruneSelectedShipmentExportIds(filteredRows);
  const rows = filteredRows.filter((item) => selectedShipmentIdsForExport.has(cleanText(item.id)));
  if (!rows.length) {
    alert("请先在汇总区勾选需要导出的记录。");
    return null;
  }

  const { from, to } = getDateRangeFromInputs(shipmentFromDateInput, shipmentToDateInput);
  const typeValue = cleanText(shipmentFilter?.value || "all");
  const typeLabelMap = {
    all: "全部",
    receive: "收货",
    ship: "发货",
  };
  const typePart = sanitizeExportPart(typeLabelMap[typeValue] || "全部");
  const manufacturerPart = sanitizeExportPart(shipmentManufacturerFilter?.value || "", "");
  const datePart = sanitizeExportPart(formatDateRangePart(from, to), "全部日期");
  const todayPart = dateValue(new Date());
  const filenameParts = ["收发记录", "勾选", typePart];

  if (manufacturerPart) {
    filenameParts.push(manufacturerPart);
  }
  filenameParts.push(datePart, `已选${rows.length}条`, todayPart);

  return {
    rows,
    filename: `${filenameParts.join("_")}.csv`,
    headers: SHIPMENT_EXPORT_HEADERS,
  };
}

function renderShipmentExportPreview(context) {
  if (!shipmentExportPreviewBody) {
    return;
  }
  if (!context || !Array.isArray(context.rows) || !context.rows.length) {
    shipmentExportPreviewBody.innerHTML = '<tr><td colspan="12" class="empty">暂无可预览数据</td></tr>';
    return;
  }

  shipmentExportPreviewBody.innerHTML = context.rows
    .map((item) => {
      const quantityText = formatMeasureValue(item.quantity, item.measureUnit);
      return `
      <tr>
        <td>${escapeHtml(item.date || "-")}</td>
        <td class="mono">${escapeHtml(item.serialNo || "-")}</td>
        <td>${escapeHtml(item.manufacturer || "-")}</td>
        <td>${escapeHtml(item.customer || "-")}</td>
        <td>${escapeHtml(item.product || "-")}</td>
        <td>${escapeHtml(quantityText || "-")}</td>
        <td>${escapeHtml(formatMeasureUnit(item.measureUnit))}</td>
        <td>${toFiniteNumber(item.unitPrice).toFixed(2)}</td>
        <td>${toFiniteNumber(item.amount).toFixed(2)}</td>
        <td>${formatDecimal(getShipmentTotalWeight(item))}</td>
        <td>${formatDecimal(getShipmentTotalCubic(item))}</td>
        <td>${escapeHtml(item.note || "-")}</td>
      </tr>`;
    })
    .join("");
}

function openShipmentExportPreviewModal(context) {
  if (!shipmentExportPreviewModal) {
    return;
  }
  pendingShipmentExportContext = context;
  renderShipmentExportPreview(context);
  if (shipmentExportPreviewTip) {
    shipmentExportPreviewTip.textContent = `将导出 ${context.rows.length} 条记录，文件名：${context.filename}`;
  }
  shipmentExportPreviewModal.classList.add("is-open");
  syncModalBodyState();
}

function closeShipmentExportPreviewModal() {
  if (!shipmentExportPreviewModal) {
    return;
  }
  shipmentExportPreviewModal.classList.remove("is-open");
  pendingShipmentExportContext = null;
  syncModalBodyState();
}

function handleShipmentExportPreviewClick() {
  const context = buildSelectedShipmentExportContext();
  if (!context) {
    return;
  }
  openShipmentExportPreviewModal(context);
}

function handleShipmentExportPreviewConfirm() {
  if (!pendingShipmentExportContext) {
    return;
  }
  const context = pendingShipmentExportContext;
  downloadCsv(context.filename, context.headers, buildShipmentExportRows(context.rows));
  closeShipmentExportPreviewModal();
}

function exportShipmentsCsv() {
  const rows = sortShipmentRows(getFilteredShipments());
  if (!rows.length) {
    alert("当前筛选条件下没有可导出的记录。");
    return;
  }

  const { from, to } = getDateRangeFromInputs(shipmentFromDateInput, shipmentToDateInput);
  const typeValue = cleanText(shipmentFilter?.value || "all");
  const typeLabelMap = {
    all: "全部",
    receive: "收货",
    ship: "发货",
  };
  const typePart = sanitizeExportPart(typeLabelMap[typeValue] || "全部");
  const manufacturerPart = sanitizeExportPart(shipmentManufacturerFilter?.value || "", "");
  const datePart = sanitizeExportPart(formatDateRangePart(from, to), "全部日期");
  const todayPart = dateValue(new Date());
  const filenameParts = ["收发记录", typePart];

  if (manufacturerPart) {
    filenameParts.push(manufacturerPart);
  }
  filenameParts.push(datePart, todayPart);

  downloadCsv(`${filenameParts.join("_")}.csv`, SHIPMENT_EXPORT_HEADERS, buildShipmentExportRows(rows));
}

function exportSelectedShipmentsCsv() {
  const context = buildSelectedShipmentExportContext();
  if (!context) {
    return;
  }
  downloadCsv(context.filename, context.headers, buildShipmentExportRows(context.rows));
}

function exportShipmentsByManufacturerCsv() {
  const manufacturer = cleanText(shipmentManufacturerFilter?.value);
  if (!manufacturer) {
    alert("请先输入厂家名称，再导出厂家发货记录。");
    shipmentManufacturerFilter?.focus();
    return;
  }

  const keyword = manufacturer.toLowerCase();
  const { from, to } = getDateRangeFromInputs(shipmentFromDateInput, shipmentToDateInput);
  const rows = sortShipmentRows(
    shipments.filter((item) => {
      if (item.type !== "ship") {
        return false;
      }
      if (!isDateInRange(item.date, from, to)) {
        return false;
      }
      return cleanText(item.manufacturer).toLowerCase().includes(keyword);
    })
  );

  if (!rows.length) {
    alert(`未找到厂家“${manufacturer}”的发货记录。`);
    return;
  }

  const datePart = sanitizeExportPart(formatDateRangePart(from, to), "全部日期");
  const manufacturerPart = sanitizeExportPart(manufacturer, "厂家");
  const todayPart = dateValue(new Date());
  const filename = `厂家发货_${manufacturerPart}_${datePart}_${todayPart}.csv`;
  downloadCsv(filename, SHIPMENT_EXPORT_HEADERS, buildShipmentExportRows(rows));
}

function downloadCsv(filename, headers, rows) {
  const lines = [headers, ...rows].map((row) => row.map(csvCell).join(","));
  const content = `\uFEFF${lines.join("\n")}`;
  const blob = new Blob([content], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);

  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();

  URL.revokeObjectURL(url);
}

function csvCell(value) {
  const text = String(value ?? "");
  if (/[",\n]/.test(text)) {
    return `"${text.replaceAll('"', '""')}"`;
  }
  return text;
}

function sumShipment(type) {
  return shipments
    .filter((item) => item.type === type)
    .reduce(
      (acc, item) => {
        acc.quantity += Number(item.quantity) || 0;
        acc.weight += getShipmentTotalWeight(item);
        acc.cubic += getShipmentTotalCubic(item);
        return acc;
      },
      { quantity: 0, weight: 0, cubic: 0 }
    );
}

function normalizeShipmentType(value) {
  const text = cleanText(value).toLowerCase();

  if (!text) {
    return "";
  }

  if (["receive", "in", "r"].includes(text) || text.includes("收")) {
    return "receive";
  }

  if (["ship", "out", "s"].includes(text) || text.includes("发") || text.includes("出")) {
    return "ship";
  }

  return "";
}

function normalizeMeasureUnit(value, fallback = "") {
  const text = cleanText(value).toLowerCase();

  if (!text) {
    return fallback;
  }

  if (["piece", "pieces", "pc", "pcs", "件", "件数"].includes(text) || text.includes("件")) {
    return "piece";
  }

  if (
    ["ton", "tons", "t", "吨", "吨位", "kg", "公斤", "千克"].includes(text) ||
    text.includes("吨") ||
    text.includes("kg") ||
    text.includes("公斤") ||
    text.includes("千克")
  ) {
    return "ton";
  }

  if (
    ["cubic", "cube", "cbm", "m3", "立方", "立方米", "方", "方数", "square", "sq", "m2", "㎡", "平方", "平米", "平方米"].includes(text) ||
    text.includes("立方") ||
    text.includes("平方") ||
    text.includes("平米") ||
    text.includes("㎡")
  ) {
    return "cubic";
  }

  return fallback;
}

function normalizeShipmentLoadedFlag(value) {
  if (typeof value === "boolean") {
    return value;
  }

  const text = cleanText(value).toLowerCase();
  if (!text) {
    return false;
  }

  if (["1", "true", "yes", "y", "loaded", "done"].includes(text)) {
    return true;
  }

  if (["0", "false", "no", "n", "pending"].includes(text)) {
    return false;
  }

  if (text.includes("未装") || text.includes("待装") || text.includes("未配")) {
    return false;
  }

  if (text.includes("已装") || text.includes("装车") || text.includes("配货")) {
    return true;
  }

  return false;
}

function normalizeShipmentPaidFlag(value, fallback = false) {
  if (typeof value === "boolean") {
    return value;
  }

  const text = cleanText(value).toLowerCase();
  if (!text) {
    return fallback;
  }

  if (["1", "true", "yes", "y", "paid", "done", "settled", "received"].includes(text)) {
    return true;
  }

  if (["0", "false", "no", "n", "unpaid", "pending"].includes(text)) {
    return false;
  }

  if (text.includes("未收") || text.includes("欠款") || text.includes("未结")) {
    return false;
  }

  if (text.includes("已收") || text.includes("收款") || text.includes("结清") || text.includes("已结")) {
    return true;
  }

  return fallback;
}

function formatMeasureUnit(value) {
  const unit = normalizeMeasureUnit(value, DEFAULT_MEASURE_UNIT) || DEFAULT_MEASURE_UNIT;
  return MEASURE_UNIT_LABELS[unit] || MEASURE_UNIT_LABELS[DEFAULT_MEASURE_UNIT];
}

function formatMeasureValue(value, measureUnit) {
  const amount = toFiniteNumber(value);
  return new Intl.NumberFormat("zh-CN", {
    minimumFractionDigits: 0,
    maximumFractionDigits: 3,
  }).format(amount);
}

function formatMeasureNumber(value) {
  const amount = toFiniteNumber(value);
  return new Intl.NumberFormat("zh-CN", {
    minimumFractionDigits: 0,
    maximumFractionDigits: 3,
  }).format(amount);
}

function normalizeFinanceType(value) {
  const text = cleanText(value).toLowerCase();

  if (!text) {
    return "";
  }

  if (["income", "in", "i"].includes(text) || text.includes("收") || text.includes("入")) {
    return "income";
  }

  if (
    ["expense", "out", "e"].includes(text) ||
    text.includes("支") ||
    text.includes("付") ||
    text.includes("出")
  ) {
    return "expense";
  }

  return "";
}

function parseDateFromAny(value) {
  if (value instanceof Date) {
    return Number.isNaN(value.getTime()) ? null : dateValue(value);
  }

  if (typeof value === "number") {
    if (!Number.isFinite(value)) {
      return null;
    }

    if (value > 10_000_000_000) {
      return dateValue(new Date(value));
    }

    if (value > 1_000_000_000) {
      return dateValue(new Date(value * 1000));
    }

    if (value > 10000 && Number.isInteger(value)) {
      const asText = String(value);
      if (/^\d{8}$/.test(asText)) {
        return normalizeYmdText(asText);
      }
    }

    const date = excelSerialToDate(value);
    return date ? dateValue(date) : null;
  }

  const text = cleanText(value);
  if (!text) {
    return null;
  }

  if (/^\d{4}[-/.]\d{1,2}[-/.]\d{1,2}$/.test(text)) {
    return normalizeYmdText(text);
  }

  if (/^\d{8}$/.test(text)) {
    return normalizeYmdText(text);
  }

  if (/^\d+(\.\d+)?$/.test(text)) {
    return parseDateFromAny(Number(text));
  }

  const parsed = new Date(text);
  if (!Number.isNaN(parsed.getTime())) {
    return dateValue(parsed);
  }

  return null;
}

function normalizeYmdText(text) {
  const normalized = text.replaceAll(".", "-").replaceAll("/", "-");

  if (/^\d{8}$/.test(normalized)) {
    const year = normalized.slice(0, 4);
    const month = normalized.slice(4, 6);
    const day = normalized.slice(6, 8);
    return `${year}-${month}-${day}`;
  }

  const [year, month, day] = normalized.split("-");
  const date = new Date(Number(year), Number(month) - 1, Number(day));

  if (
    Number.isNaN(date.getTime()) ||
    date.getFullYear() !== Number(year) ||
    date.getMonth() + 1 !== Number(month) ||
    date.getDate() !== Number(day)
  ) {
    return null;
  }

  return `${year.padStart(4, "0")}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
}

function excelSerialToDate(serial) {
  const days = Number(serial);
  if (!Number.isFinite(days)) {
    return null;
  }

  const utcMs = Math.round((days - 25569) * 86400 * 1000);
  const date = new Date(utcMs);
  return Number.isNaN(date.getTime()) ? null : date;
}

function parseNumericValue(value) {
  if (typeof value === "number") {
    return value;
  }

  const text = cleanText(value)
    .replaceAll(",", "")
    .replaceAll("¥", "")
    .replaceAll("￥", "");

  if (!text) {
    return NaN;
  }

  return Number(text);
}

function parseOptionalNonNegative(value) {
  const text = cleanText(value);
  if (!text) {
    return 0;
  }

  const num = parseNumericValue(text);
  if (!Number.isFinite(num) || num < 0) {
    return NaN;
  }
  return num;
}

function normalizeInvoiceFlag(value) {
  if (typeof value === "boolean") {
    return value;
  }

  const text = cleanText(value).toLowerCase();
  if (!text) {
    return false;
  }

  if (["yes", "y", "true", "1"].includes(text)) {
    return true;
  }

  if (["no", "n", "false", "0"].includes(text)) {
    return false;
  }

  if (text.includes("否") || text.includes("无") || text.includes("不开")) {
    return false;
  }

  if (text.includes("是") || text.includes("开") || text.includes("有")) {
    return true;
  }

  return false;
}

function normalizePaymentMethod(value) {
  const text = cleanText(value).toLowerCase();
  if (!text) {
    return "";
  }
  if (text === "signback" || text.includes("签回")) {
    return "signback";
  }
  if (text === "pickup" || text.includes("自提")) {
    return "pickup";
  }
  if (text === "delivery" || text.includes("送货")) {
    return "delivery";
  }
  if (text === "notify" || text.includes("通知")) {
    return "notify";
  }
  return "";
}

function formatPaymentMethod(value) {
  const key = normalizePaymentMethod(value);
  return key ? PAYMENT_METHOD_LABELS[key] : "未填付款方式";
}

function computeShipmentTotalFee(item) {
  return toFiniteNumber(item.freightFee) + toFiniteNumber(item.deliveryFee) + toFiniteNumber(item.insuranceFee);
}

function toFiniteNumber(value, fallback = 0) {
  const num = Number(value);
  return Number.isFinite(num) ? num : fallback;
}

function sortByDateDesc(a, b) {
  if (a.date !== b.date) {
    return b.date.localeCompare(a.date);
  }
  return (b.createdAt || 0) - (a.createdAt || 0);
}

function tickClock() {
  const now = new Date();
  currentTimeEl.textContent = now.toLocaleString("zh-CN", {
    weekday: "long",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    hour12: false,
  });

  // 在家结提醒页停留时自动刷新，确保“稍后处理(1小时)”到期后可实时再次提醒
  if (activePanelId === "status-panel") {
    renderStatusSection();
  }
}

function dateValue(date) {
  const tzOffset = date.getTimezoneOffset() * 60000;
  return new Date(date.getTime() - tzOffset).toISOString().slice(0, 10);
}

function formatDateTime(value) {
  const ts = Number(value);
  if (!Number.isFinite(ts) || ts <= 0) {
    return "-";
  }
  const date = new Date(ts);
  if (Number.isNaN(date.getTime())) {
    return "-";
  }
  return date.toLocaleString("zh-CN", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hour12: false,
  });
}

function formatNumber(value) {
  return new Intl.NumberFormat("zh-CN", {
    minimumFractionDigits: 0,
    maximumFractionDigits: 2,
  }).format(value);
}

function formatDecimal(value) {
  return Number(value).toFixed(2);
}

function formatCurrency(value) {
  return new Intl.NumberFormat("zh-CN", {
    style: "currency",
    currency: "CNY",
    minimumFractionDigits: 2,
  }).format(value);
}

function cleanText(value) {
  return String(value ?? "").trim();
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

async function hydrateFromFileStorageIfAvailable() {
  fileStorageEnabled = await detectFileStorageAvailability();
  fileStorageReady = true;
  if (!fileStorageEnabled) {
    return;
  }

  const keys = Array.from(FILE_STORAGE_KEYS);
  for (const key of keys) {
    const payload = await fetchFileStorageValue(key);
    if (!payload.exists) {
      continue;
    }
    try {
      localStorage.setItem(key, JSON.stringify(payload.value));
    } catch {
      // Ignore storage quota failures and keep runtime state.
    }
  }

  shipments = readStorage(STORAGE_KEYS.shipments);
  financeRecords = readStorage(STORAGE_KEYS.finance);
  dispatchRecords = readStorage(STORAGE_KEYS.dispatchRecords);
  homeSettleActionRecords = sanitizeHomeSettleActionRecords(readStorage(STORAGE_KEYS.homeSettleActionRecords));
  customCatalogProfiles = sanitizeCustomCatalogProfiles(readStorage(STORAGE_KEYS.catalogProfiles));
}

function updateStorageModeBadge() {
  if (!storageModeBadge) {
    return;
  }

  storageModeBadge.classList.remove("is-file", "is-browser");
  storageModeBadge.classList.add("is-manual");

  const saveState = manualSaveDirty
    ? "手动保存模式：有未保存更改"
    : manualSaveLastAt > 0
      ? `手动保存模式：已保存 ${formatDateTime(manualSaveLastAt)}`
      : "手动保存模式：尚未保存";
  storageModeBadge.textContent = saveState;
  storageModeBadge.title = "点击查看跨电脑接续说明（手动保存数据包/手动读取数据包）";
}

function handleStorageModeBadgeClick() {
  const lines = [
    "当前为【手动保存模式】。",
    "",
    "建议操作：",
    "1）每次录入后点击“手动保存数据包”（建议保存到项目文件夹内）。",
    "2）换电脑后点击“手动读取数据包”，继续工作。",
    "3）若顶部显示“有未保存更改”，请先手动保存再关闭页面。",
  ];
  alert(lines.join("\n"));
}

function applyStorageGuard() {
  closeStorageRequiredModal();
}

function openStorageRequiredModal() {
  return;
}

function closeStorageRequiredModal() {
  if (!storageRequiredModal) {
    return;
  }
  storageRequiredModal.classList.remove("is-open");
  syncModalBodyState();
}

async function handleStorageGuardRetryClick() {
  return;
}

function handleStorageGuardContinueClick() {
  return;
}

async function detectFileStorageAvailability() {
  fileStorageDataDir = "";
  return false;
}

async function fetchFileStorageValue(key) {
  return { exists: false };
}

function shouldPersistToFileStorage(key) {
  return false;
}

function cloneForStorage(value) {
  try {
    return JSON.parse(JSON.stringify(value));
  } catch {
    return value;
  }
}

function scheduleFileStorageWrite(key, value) {
  return;
}

async function flushFileStorageWrites() {
  return;
}

async function persistFileStorageValue(key, value) {
  return;
}

function readObjectStorage(key, fallback = {}) {
  try {
    const raw = localStorage.getItem(key);
    if (!raw) {
      return fallback;
    }
    const parsed = JSON.parse(raw);
    if (!parsed || Array.isArray(parsed) || typeof parsed !== "object") {
      return fallback;
    }
    return parsed;
  } catch {
    return fallback;
  }
}

function writeObjectStorage(key, value) {
  try {
    localStorage.setItem(key, JSON.stringify(value));
  } catch {
    // Ignore storage quota failures and keep runtime state.
  }
  scheduleFileStorageWrite(key, value);
}

function readSimpleStorage(key, fallback = []) {
  try {
    const raw = localStorage.getItem(key);
    if (!raw) {
      return fallback;
    }
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) {
      return fallback;
    }
    return parsed;
  } catch {
    return fallback;
  }
}

function writeSimpleStorage(key, value) {
  try {
    localStorage.setItem(key, JSON.stringify(value));
  } catch {
    // Ignore storage quota failures and keep runtime state.
  }
  scheduleFileStorageWrite(key, value);
}

function readStorage(key) {
  try {
    const raw = localStorage.getItem(key);
    if (!raw) {
      return [];
    }

    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) {
      return [];
    }

    if (key === STORAGE_KEYS.shipments) {
      return parsed
        .map((item, idx) => {
          let type = normalizeShipmentType(item.type) || "receive";
          const date = parseDateFromAny(item.date);
          const measureUnit =
            normalizeMeasureUnit(item.measureUnit || item.unit || item.uom, DEFAULT_MEASURE_UNIT) || DEFAULT_MEASURE_UNIT;
          const quantity = toFiniteNumber(item.quantity);
          const unitWeightCandidate = toFiniteNumber(item.weight, NaN);
          const explicitUnitWeight = toFiniteNumber(item.unitWeight, NaN);
          const totalWeightCandidate = toFiniteNumber(item.totalWeight, NaN);
          let weight = Number.isFinite(explicitUnitWeight) ? explicitUnitWeight : unitWeightCandidate;
          if (!Number.isFinite(weight) && Number.isFinite(totalWeightCandidate) && quantity > 0) {
            weight = totalWeightCandidate / quantity;
          }
          if (!Number.isFinite(weight)) {
            weight = 0;
          }
          const unitPrice = Math.max(0, toFiniteNumber(item.unitPrice));
          const rebate = Math.max(0, toFiniteNumber(item.rebate));
          const legacyAmount = Math.max(0, toFiniteNumber(item.freightFee) + toFiniteNumber(item.deliveryFee) + toFiniteNumber(item.insuranceFee));
          const amountCandidate = toFiniteNumber(item.amount, NaN);
          const amount = Number.isFinite(amountCandidate)
            ? Math.max(0, amountCandidate)
            : legacyAmount > 0
              ? legacyAmount
              : Math.max(0, quantity * unitPrice);
          const manufacturer = cleanText(item.manufacturer) || cleanText(item.consignor);
          const customer = cleanText(item.customer) || cleanText(item.consignee) || cleanText(item.partner);
          const product = cleanText(item.product) || cleanText(item.cargoName);
          const departStation = cleanText(item.departStation);
          const arriveStation = cleanText(item.arriveStation);
          const routeText = departStation || arriveStation ? `${departStation || "-"}→${arriveStation || "-"}` : "";
          const legacyPaymentMethod = normalizePaymentMethod(item.paymentMethod);
          const delivery = cleanText(item.delivery) || (legacyPaymentMethod === "delivery" ? "送货" : "");
          const pickup = cleanText(item.pickup) || (legacyPaymentMethod === "pickup" ? "自提" : "");
          const isLoaded = normalizeShipmentLoadedFlag(item.isLoaded ?? item.loaded ?? item.loadStatus ?? item.status);
          if (isLoaded && type === "receive") {
            type = "ship";
          }
          const isPaid = normalizeShipmentPaidFlag(
            item.isPaid ?? item.paid ?? item.paymentReceived ?? item.received ?? item.collectionStatus,
            false
          );
          const loadedAtCandidate = toFiniteNumber(item.loadedAt, NaN);
          const paidAtCandidate = toFiniteNumber(item.paidAt, NaN);
          const loadedAt =
            isLoaded && Number.isFinite(loadedAtCandidate) && loadedAtCandidate > 0
              ? loadedAtCandidate
              : isLoaded
                ? Date.now() + idx
                : 0;
          const dispatchDate =
            parseDateFromAny(item.dispatchDate || item.dispatch_date || item.loadedDate) ||
            (isLoaded ? parseDateFromAny(loadedAt) || "" : "");
          const dispatchTruckNo = cleanText(item.dispatchTruckNo || item.truckNo || item.carNo || item.carIndex);
          const dispatchPlateNo = cleanText(item.dispatchPlateNo || item.plateNo || item.plate).toUpperCase();
          const dispatchDriver = cleanText(item.dispatchDriver || item.driver);
          const dispatchContactName = cleanText(item.dispatchContactName || item.contactName || item.name);
          const dispatchDestinationInput = cleanText(item.dispatchDestination || item.destination);
          const dispatchDestination =
            dispatchDestinationInput ||
            (isLoaded
              ? buildDispatchDestinationText({
                  dispatchDate,
                  dispatchTruckNo,
                  dispatchPlateNo,
                  dispatchDriver,
                  dispatchContactName,
                })
              : "");
          const paidAt =
            isPaid && Number.isFinite(paidAtCandidate) && paidAtCandidate > 0
              ? paidAtCandidate
              : isPaid
                ? Date.now() + idx
                : 0;
          const homeSettleStatus = normalizeHomeSettleStatusId(
            item.homeSettleStatus || item.settleStatus || item.reminderStatus
          );
          const homeSettleStatusUpdatedAt = Math.max(0, toFiniteNumber(item.homeSettleStatusUpdatedAt));
          const homeSettleRemindedAt = Math.max(
            0,
            toFiniteNumber(item.homeSettleRemindedAt ?? item.homeSettleReminderAt ?? item.remindedAt)
          );
          const homeSettleSnoozeUntil = Math.max(
            0,
            toFiniteNumber(item.homeSettleSnoozeUntil ?? item.homeSettleNextReminderAt ?? item.snoozeUntil)
          );
          const homeSettleLastAction = cleanText(item.homeSettleLastAction || item.lastReminderAction);
          const homeSettleLastActionAt = Math.max(
            0,
            toFiniteNumber(item.homeSettleLastActionAt ?? item.lastReminderActionAt)
          );

          if (!date || quantity <= 0 || weight < 0) {
            return null;
          }

          return {
            ...item,
            type,
            date,
            serialNo: cleanText(item.serialNo) || cleanText(item.orderNo),
            manufacturer,
            customer,
            product,
            measureUnit,
            quantity,
            unitPrice,
            amount,
            rebate,
            weight,
            transit: cleanText(item.transit) || routeText,
            delivery,
            pickup,
            note: cleanText(item.note),
            isLoaded,
            loadedAt,
            dispatchDate,
            dispatchTruckNo,
            dispatchPlateNo,
            dispatchDriver,
            dispatchContactName,
            dispatchDestination,
            destination: dispatchDestination,
            homeSettleStatus,
            homeSettleStatusUpdatedAt,
            homeSettleRemindedAt,
            homeSettleSnoozeUntil,
            homeSettleLastAction,
            homeSettleLastActionAt,
            isPaid,
            paidAt,
            createdAt: Number(item.createdAt) || Date.now() + idx,
          };
        })
        .filter(Boolean);
    }

    if (key === STORAGE_KEYS.finance) {
      return parsed
        .map((item, idx) => {
          const parsedAmount = toFiniteNumber(item.amount);
          let type = normalizeFinanceType(item.type);
          let amount = parsedAmount;

          if (!type && parsedAmount < 0) {
            type = "expense";
            amount = Math.abs(parsedAmount);
          }

          if (!type && parsedAmount > 0) {
            type = "income";
          }

          const date = parseDateFromAny(item.date);
          amount = toFiniteNumber(amount);

          if (!type || !date || amount <= 0) {
            return null;
          }

          return {
            ...item,
            type,
            date,
            amount,
            createdAt: Number(item.createdAt) || Date.now() + idx,
          };
        })
        .filter(Boolean);
    }

    if (key === STORAGE_KEYS.dispatchRecords) {
      return parsed
        .map((record, idx) => {
          const items = Array.isArray(record.items)
            ? record.items
                .map((item, itemIdx) => {
                  const measureUnit =
                    normalizeMeasureUnit(item.measureUnit || item.unit || item.uom, DEFAULT_MEASURE_UNIT) || DEFAULT_MEASURE_UNIT;
                  const quantity = Math.max(0, toFiniteNumber(item.quantity));
                  const unitPrice = Math.max(0, toFiniteNumber(item.unitPrice));
                  const amountCandidate = toFiniteNumber(item.amount, NaN);
                  const amount = Number.isFinite(amountCandidate) ? Math.max(0, amountCandidate) : Math.max(0, quantity * unitPrice);
                  const unitWeightCandidate = toFiniteNumber(item.unitWeight ?? item.weight, NaN);
                  const totalWeightCandidate = toFiniteNumber(item.totalWeight, NaN);
                  const unitCubicCandidate = toFiniteNumber(item.unitCubic ?? item.cubic, NaN);
                  const totalCubicCandidate = toFiniteNumber(item.totalCubic, NaN);
                  const unitWeight = Number.isFinite(unitWeightCandidate) ? Math.max(0, unitWeightCandidate) : 0;
                  const totalWeight =
                    Number.isFinite(totalWeightCandidate) && totalWeightCandidate >= 0
                      ? totalWeightCandidate
                      : Math.max(0, unitWeight * quantity);
                  const unitCubic =
                    Number.isFinite(unitCubicCandidate) && unitCubicCandidate >= 0
                      ? unitCubicCandidate
                      : measureUnit === "cubic"
                        ? 1
                        : 0;
                  const totalCubic =
                    Number.isFinite(totalCubicCandidate) && totalCubicCandidate >= 0
                      ? totalCubicCandidate
                      : Math.max(0, unitCubic * quantity);
                  const dispatchDate = parseDateFromAny(item.dispatchDate) || "";
                  const dispatchTruckNo = cleanText(item.dispatchTruckNo || item.truckNo);
                  const dispatchPlateNo = cleanText(item.dispatchPlateNo || item.plateNo).toUpperCase();
                  const dispatchDriver = cleanText(item.dispatchDriver || item.driver);
                  const dispatchContactName = cleanText(item.dispatchContactName || item.contactName || item.name);
                  const dispatchDestinationInput = cleanText(item.dispatchDestination || item.destination);
                  const dispatchDestination =
                    dispatchDestinationInput ||
                    buildDispatchDestinationText({
                      dispatchDate,
                      dispatchTruckNo,
                      dispatchPlateNo,
                      dispatchDriver,
                      dispatchContactName,
                    });

                  return {
                    shipmentId: cleanText(item.shipmentId || item.id),
                    date: parseDateFromAny(item.date) || cleanText(item.date),
                    serialNo: cleanText(item.serialNo),
                    manufacturer: cleanText(item.manufacturer),
                    customer: cleanText(item.customer),
                    product: cleanText(item.product),
                    measureUnit,
                    quantity,
                    unitPrice,
                    amount,
                    unitWeight,
                    unitCubic,
                    totalWeight,
                    totalCubic,
                    note: cleanText(item.note),
                    dispatchDate,
                    dispatchTruckNo,
                    dispatchPlateNo,
                    dispatchDriver,
                    dispatchContactName,
                    dispatchDestination,
                    createdAt: Number(item.createdAt) || Date.now() + itemIdx,
                  };
                })
                .filter(Boolean)
            : [];

          if (items.length === 0) {
            return null;
          }

          const loadedAtCandidate = toFiniteNumber(record.loadedAt, NaN);
          const createdAtCandidate = toFiniteNumber(record.createdAt, NaN);
          const loadedAt =
            Number.isFinite(loadedAtCandidate) && loadedAtCandidate > 0
              ? loadedAtCandidate
              : Number.isFinite(createdAtCandidate) && createdAtCandidate > 0
                ? createdAtCandidate
                : Date.now() + idx;
          const createdAt =
            Number.isFinite(createdAtCandidate) && createdAtCandidate > 0 ? createdAtCandidate : loadedAt;
          const totals = computeDispatchRecordTotals(items);
          const totalAmountCandidate = toFiniteNumber(record.totalAmount, NaN);
          const totalWeightCandidate = toFiniteNumber(record.totalWeight, NaN);
          const totalCubicCandidate = toFiniteNumber(record.totalCubic, NaN);
          const totalPiecesCandidate = toFiniteNumber(record.totalPieces, NaN);
          const itemCountCandidate = toFiniteNumber(record.itemCount, NaN);
          const recordMeta = getDispatchMetaFromRecord({
            ...record,
            items,
            loadedAt,
            createdAt,
          });

          return {
            id: cleanText(record.id) || `dispatch_${loadedAt}_${idx}`,
            loadedAt,
            createdAt,
            itemCount:
              Number.isFinite(itemCountCandidate) && itemCountCandidate >= 0
                ? Math.round(itemCountCandidate)
                : items.length,
            totalAmount:
              Number.isFinite(totalAmountCandidate) && totalAmountCandidate >= 0
                ? totalAmountCandidate
                : totals.totalAmount,
            totalWeight:
              Number.isFinite(totalWeightCandidate) && totalWeightCandidate >= 0
                ? totalWeightCandidate
                : totals.totalWeight,
            totalCubic:
              Number.isFinite(totalCubicCandidate) && totalCubicCandidate >= 0
                ? totalCubicCandidate
                : totals.totalCubic,
            totalPieces:
              Number.isFinite(totalPiecesCandidate) && totalPiecesCandidate >= 0
                ? totalPiecesCandidate
                : totals.totalPieces,
            dispatchDate: recordMeta.dispatchDate,
            dispatchTruckNo: recordMeta.dispatchTruckNo,
            dispatchPlateNo: recordMeta.dispatchPlateNo,
            dispatchDriver: recordMeta.dispatchDriver,
            dispatchContactName: recordMeta.dispatchContactName,
            dispatchDestination: recordMeta.dispatchDestination,
            productsSummary: cleanText(record.productsSummary) || buildDispatchProductsSummary(items),
            items,
          };
        })
        .filter(Boolean);
    }

    return parsed;
  } catch {
    return [];
  }
}

function writeStorage(key, value) {
  try {
    localStorage.setItem(key, JSON.stringify(value));
  } catch {
    // Ignore storage quota failures and keep runtime state.
  }
  scheduleFileStorageWrite(key, value);
  if (FILE_STORAGE_KEYS.has(key)) {
    markManualSaveDirty();
  }
}

function makeId() {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") {
    return crypto.randomUUID();
  }
  return `id_${Date.now()}_${Math.floor(Math.random() * 1e6)}`;
}
