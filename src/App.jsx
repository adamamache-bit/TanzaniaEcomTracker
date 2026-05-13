import React, { useCallback, useDeferredValue, useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import {
  BarChart3,
  Boxes,
  TrendingUp,
  Wallet,
  ClipboardList,
  Calculator,
  LayoutGrid,
  Rocket,
  AlertTriangle,
  Archive,
  Users,
  ShoppingBag,
  MessageSquare,
  Settings,
  Phone,
  MapPin,
  CalendarDays,
  ChevronLeft,
  ChevronRight,
  ChevronDown,
  XCircle,
} from "lucide-react";
import {
  ResponsiveContainer,
  BarChart,
  Bar,
  CartesianGrid,
  XAxis,
  YAxis,
  Tooltip,
  Legend,
  PieChart,
  Pie,
  Cell,
  LineChart,
  Line,
} from "recharts";
import {
  getCloudSession,
  listCloudWorkspaceBackups,
  loadCloudWorkspace,
  onCloudAuthStateChange,
  restoreCloudWorkspaceBackup,
  saveCloudWorkspace,
  saveCloudWorkspaceAnon,
  signInCloud,
  signOutCloud,
  signUpCloud,
  subscribeToCloudWorkspace,
} from "./lib/cloudWorkspace";
import {
  addDaysToDateString,
  addMonths,
  appendCustomerHistory,
  buildCalendarMatrix,
  calculateCodFunnelMetrics,
  calculateProductAlerts,
  calculateProductPerformance,
  buildHistoryEntry,
  buildMappedMetaRows,
  buildNextId,
  doesDateRangeOverlap,
  DEFAULT_CONFIRMATION_STATUSES,
  DEFAULT_POST_CONFIRMATION_STATUSES,
  ensureShippingStatusForConfirmed,
  excelDateToInput,
  formatDateInput,
  formatInteger,
  formatLongDate,
  formatMetaLeadSourceLabel,
  formatOffersSummary,
  formatStatusLabel,
  formatTZS,
  formatUSD,
  formatUsdFromTzs,
  getConfirmationBucket,
  getCustomerConfirmationStatus,
  getCustomerEffectiveStatus,
  getCustomerOrderTotalTzs,
  getCustomerShippingStatus,
  getDateRangeFromPreset,
  getDayBucket,
  getDefaultCloudWorkspaceState,
  getDefaultImportMeta,
  getDefaultMetaAdsState,
  getDefaultServiceForm,
  getDefaultSituationData,
  getEmptyCustomerForm,
  getEmptyExpeditionForm,
  getMetaApiBase,
  getProductPricing,
  getSharedApiBase,
  getStatusBadgeStyle,
  getShippingBucket,
  getStatusColor,
  getTodayString,
  getUnitProductCostUSD,
  getWeekLabel,
  getWeekStartString,
  hasMeaningfulWorkspaceData,
  INITIAL_CUSTOMERS,
  initialProducts,
  initialTracking,
  isConfirmationCancelled,
  isConfirmationConfirmed,
  isConfirmationNew,
  isCountableLeadForService,
  isDateWithinRange,
  isShippingDelivered,
  isShippingInProgress,
  isShippingReturned,
  matchProductIdFromText,
  META_RANGE_PRESETS,
  PAGE_DATE_FILTER_PRESETS,
  normalizeHeaderName,
  normalizeOrderStatus,
  normalizePhoneValue,
  normalizeProductOffers,
  parseDateInput,
  parseLooseNumber,
  resolveDateRangeFilter,
  sanitizeCustomerRecord,
  sanitizeMetaAdsState,
  sanitizeProductRecord,
  sanitizeServiceForm,
  sanitizeSituationData,
  serviceCountryData,
  startOfMonth,
  USD_TO_TZS,
} from "./lib/appLogic";
import { supabaseEnabled, supabaseWorkspaceId } from "./lib/supabaseClient";
import { checkNormalizedTablesEmpty, clearNormalizedProducts, deleteExtraChargeFromSupabase, deleteManualAdsSpendFromSupabase, deleteOwnerInjectionFromSupabase, loadAdsCampaignsFromSupabase, loadAdsSpendByProductFromSupabase, loadExtraChargesFromSupabase, loadManualAdsSpendFromSupabase, loadOwnerInjectionsFromSupabase, loadProfitOverviewFromSupabase, loadRevenueImportFromSupabase, loadRevenueImportRowsFromSupabase, loadWorkspaceFromNormalizedTables, migrateWorkspaceToNormalizedTables, saveAdsCampaignsToSupabase, saveExtraChargeToSupabase, saveManualAdsSpendToSupabase, saveOwnerInjectionToSupabase, saveRevenueImportRowsToSupabase, saveRevenueImportToSupabase, syncNormalizedTables } from "./lib/supabaseSync";
import { parseImportedExcelRows } from "./utils/importMapping";
import {
  calculateAvailableStock,
  calculateAvailableStockValue,
  calculateDamagedStock,
  calculateDeliveredStock,
  calculateReceivedStock,
  calculateReservedStock,
  calculateReturnedStock,
  calculateServiceFeeForOrder,
} from "./utils/stockCalculations";
import { getOrderRevenueAmounts } from "./utils/calculations";
import { calculateTotalAdsSpend, deduplicateCampaigns } from "./utils/adsMapping";
import {
  normalizeStatus,
  isConfirmedStatus,
  isNoReplyStatus,
  isCancelledStatus,
  isInvalidLeadStatus,
  isStockHoldStatus,
  isDeliveredStatus,
  isReturnedStatus,
  isPendingShippingStatus,
  isBlankShippingStatus,
} from "./config/statusMapping";

const STORAGE_KEY = "tanzania-ecom-tracker-v16";
const AUTO_BACKUP_KEY = "tanzania-ecom-tracker-auto-backup-v1";
const AUTO_BACKUP_META_KEY = "tanzania-ecom-tracker-auto-backup-meta-v1";
const IMPORT_META_KEY = "tanzania-ecom-tracker-import-meta-v1";
const DEFAULT_PAGE_DATE_PRESET = "last30Days";

// Profit Center localStorage keys (fallback when Supabase unavailable or tables not yet created)
const LS_MANUAL_ADS = "profit_manual_ads_v1";
const LS_EXTRA_CHARGES = "profit_extra_charges_v1";
const LS_OWNER_INJECTIONS = "profit_owner_injections_v1";
const LS_REVENUE_ROWS = "profit_revenue_rows_v1";
const LS_REVENUE_IMPORT = "profit_revenue_import_v1";

function readLS(key) {
  try { return JSON.parse(localStorage.getItem(key) || "null"); } catch { return null; }
}
function writeLS(key, val) {
  try { localStorage.setItem(key, JSON.stringify(val)); } catch { void 0; }
}

function createPageDateFilterState(preset = DEFAULT_PAGE_DATE_PRESET) {
  const range = getDateRangeFromPreset(preset);
  return {
    preset,
    startDate: range.start,
    endDate: range.end,
  };
}

function buildOrderStatusMovement(prevOrder, nextOrder) {
  const productId = String(nextOrder?.productId || "");
  if (!productId) return null;

  const orderId = String(nextOrder?.sourceOrderId || nextOrder?.order_id || nextOrder?.id || "");
  const qty = Math.max(1, Math.round(Number(nextOrder?.quantity || 1)));
  const nextConf = nextOrder?.confirmationStatus || "";
  const nextShip = nextOrder?.shippingStatus || "";
  const isNowConfirmed = isConfirmedStatus(nextConf);
  const isNowDelivered = isDeliveredStatus(nextShip);
  const isNowReturned = isReturnedStatus(nextShip);

  if (prevOrder) {
    const prevConf = prevOrder?.confirmationStatus || "";
    const prevShip = prevOrder?.shippingStatus || "";
    const wasConfirmed = isConfirmedStatus(prevConf);
    const wasDelivered = isDeliveredStatus(prevShip);
    const wasReturned = isReturnedStatus(prevShip);
    const wasCancelled = isCancelledStatus(prevConf);
    const isNowCancelled = isCancelledStatus(nextConf);

    if (!wasConfirmed && isNowConfirmed && !isNowDelivered && !isNowReturned) {
      return { product_id: productId, type: "stock_reserved", quantity_change: -qty, source_reference: orderId, note: `Order ${orderId} confirmed — ${qty} unit(s) reserved` };
    }
    if (!wasDelivered && isNowDelivered) {
      return { product_id: productId, type: "stock_delivered", quantity_change: -qty, source_reference: orderId, note: `Order ${orderId} delivered — ${qty} unit(s)` };
    }
    if (!wasReturned && isNowReturned) {
      return { product_id: productId, type: "stock_returned", quantity_change: qty, source_reference: orderId, note: `Order ${orderId} returned — ${qty} unit(s) back` };
    }
    if (wasConfirmed && isNowCancelled && !wasCancelled) {
      return { product_id: productId, type: "stock_released", quantity_change: qty, source_reference: orderId, note: `Order ${orderId} cancelled — ${qty} unit(s) released` };
    }
  } else {
    if (isNowConfirmed && !isNowDelivered && !isNowReturned) {
      return { product_id: productId, type: "stock_reserved", quantity_change: -qty, source_reference: orderId, note: `New confirmed order ${orderId} — ${qty} unit(s) reserved` };
    }
    if (isNowDelivered) {
      return { product_id: productId, type: "stock_delivered", quantity_change: -qty, source_reference: orderId, note: `Order ${orderId} imported as delivered — ${qty} unit(s)` };
    }
    if (isNowReturned) {
      return { product_id: productId, type: "stock_returned", quantity_change: qty, source_reference: orderId, note: `Order ${orderId} imported as returned — ${qty} unit(s) back` };
    }
  }
  return null;
}

function applyOrderMovementsToState(movements, prev) {
  if (!movements.length) return prev;
  const existingKeys = new Set(prev.map((m) => `${m.source_reference}::${m.type}`));
  const now = new Date();
  const newEntries = movements
    .filter((m) => m.source_reference && !existingKeys.has(`${m.source_reference}::${m.type}`))
    .map((m, i) => ({
      movement_id: `mv-${now.getTime() + i}-${Math.random().toString(36).slice(2, 6)}`,
      date: now.toISOString().slice(0, 10),
      created_at: now.toISOString(),
      ...m,
    }));
  return newEntries.length > 0 ? [...prev, ...newEntries] : prev;
}

function isQuotaExceededError(error) {
  if (!error) return false;
  const message = String(error?.message || error || "").toLowerCase();
  return (
    message.includes("quota") ||
    message.includes("storage full") ||
    error?.name === "QuotaExceededError" ||
    error?.code === 22 ||
    error?.code === 1014
  );
}

function buildCompactCustomerForBrowser(customer, { slim = false } = {}) {
  const sanitized = sanitizeCustomerRecord(customer);
  if (!slim) {
    return {
      ...sanitized,
      history: [],
    };
  }

  return {
    id: sanitized.id,
    customerName: sanitized.customerName,
    phone: sanitized.phone,
    city: sanitized.city,
    productId: sanitized.productId,
    quantity: sanitized.quantity,
    orderDate: sanitized.orderDate,
    paymentMethod: sanitized.paymentMethod,
    status: sanitized.status,
    confirmationStatus: sanitized.confirmationStatus,
    shippingStatus: sanitized.shippingStatus,
    orderTotalTzs: sanitized.orderTotalTzs,
    sourceOrderId: sanitized.sourceOrderId,
    importSource: sanitized.importSource,
    lastImportedAt: sanitized.lastImportedAt,
    lastShippingImportedAt: sanitized.lastShippingImportedAt,
    assignedTo: sanitized.assignedTo,
    history: [],
  };
}

function buildBrowserPersistedSnapshot(source = {}, { slim = false } = {}) {
  return {
    products: Array.isArray(source.products) ? source.products.map(sanitizeProductRecord) : [],
    tracking: slim ? [] : Array.isArray(source.tracking) ? source.tracking : [],
    customers: Array.isArray(source.customers)
      ? source.customers.map((customer) => buildCompactCustomerForBrowser(customer, { slim }))
      : [],
    serviceForm: sanitizeServiceForm(source.serviceForm || getDefaultServiceForm()),
    situationData: sanitizeSituationData(source.situationData || getDefaultSituationData()),
    metaAdsState: sanitizeMetaAdsState({
      ...(source.metaAdsState || getDefaultMetaAdsState()),
      accessToken: "",
    }),
    importMeta: {
      lastOrdersImportAt: source.importMeta?.lastOrdersImportAt || null,
      lastShippingImportAt: source.importMeta?.lastShippingImportAt || null,
    },
    stockPurchases: Array.isArray(source.stockPurchases) ? source.stockPurchases : [],
    stockMovements: slim ? [] : Array.isArray(source.stockMovements) ? source.stockMovements : [],
  };
}

function persistBrowserSnapshotSafely(snapshot, { exportedAt, onSaved } = {}) {
  if (typeof window === "undefined") return false;

  const candidates = [
    buildBrowserPersistedSnapshot(snapshot, { slim: false }),
    buildBrowserPersistedSnapshot(snapshot, { slim: true }),
  ];

  for (const candidate of candidates) {
    try {
      const payload = {
        ...candidate,
        exportedAt: exportedAt || new Date().toISOString(),
        version: 1,
      };
      localStorage.setItem(STORAGE_KEY, JSON.stringify(candidate));
      localStorage.setItem(AUTO_BACKUP_KEY, JSON.stringify(payload));
      const now = new Date().toISOString();
      localStorage.setItem(AUTO_BACKUP_META_KEY, JSON.stringify({ lastAutoBackupAt: now }));
      onSaved?.(now);
      return true;
    } catch (error) {
      if (!isQuotaExceededError(error)) {
        return false;
      }
      try {
        localStorage.removeItem(STORAGE_KEY);
        localStorage.removeItem(AUTO_BACKUP_KEY);
      } catch {
        // ignore cleanup issues
      }
    }
  }

  return false;
}

function readLocalWorkspaceSnapshotFromStorage() {
  if (typeof window === "undefined") return null;

  const buildSnapshot = (source = {}, importMetaSource = null) => ({
    products: Array.isArray(source.products) ? source.products.map(sanitizeProductRecord) : [],
    tracking: Array.isArray(source.tracking) ? source.tracking : [],
    customers: Array.isArray(source.customers) ? source.customers.map(sanitizeCustomerRecord) : [],
    serviceForm: sanitizeServiceForm(source.serviceForm || getDefaultServiceForm()),
    situationData: sanitizeSituationData(source.situationData || getDefaultSituationData()),
    metaAdsState: sanitizeMetaAdsState(source.metaAdsState || getDefaultMetaAdsState()),
    importMeta: {
      lastOrdersImportAt: source.importMeta?.lastOrdersImportAt || importMetaSource?.lastOrdersImportAt || null,
      lastShippingImportAt: source.importMeta?.lastShippingImportAt || importMetaSource?.lastShippingImportAt || null,
    },
  });

  try {
    const importMetaRaw = localStorage.getItem(IMPORT_META_KEY);
    const storedImportMeta = importMetaRaw ? JSON.parse(importMetaRaw) : null;

    const autoBackupRaw = localStorage.getItem(AUTO_BACKUP_KEY);
    if (autoBackupRaw) {
      const autoBackup = buildSnapshot(JSON.parse(autoBackupRaw), storedImportMeta);
      if (hasMeaningfulWorkspaceData(autoBackup)) return autoBackup;
    }

    const storageRaw = localStorage.getItem(STORAGE_KEY);
    if (storageRaw) {
      const localSnapshot = buildSnapshot(JSON.parse(storageRaw), storedImportMeta);
      if (hasMeaningfulWorkspaceData(localSnapshot)) return localSnapshot;
    }
  } catch {
    // ignore browser backup parsing issue
  }

  return null;
}

const pageBg = "#f5f7fb";
const cardBg = "rgba(255, 255, 255, 0.94)";
const cardBorder = "#d9e1ec";
const textMain = "#172033";
const textSoft = "#667085";
const inputBg = "rgba(250, 252, 255, 0.98)";
const accent = "#2358d5";
const green = "#158f63";
const red = "#d9485f";
const amber = "#c78322";

const styles = {
  shell: {
    minHeight: "100vh",
    background: `radial-gradient(circle at top left, rgba(35, 88, 213, 0.11), transparent 24%), radial-gradient(circle at top right, rgba(199, 131, 34, 0.08), transparent 22%), linear-gradient(180deg, #f8fafc 0%, ${pageBg} 100%)`,
    color: textMain,
    fontFamily: "\"Segoe UI Variable Text\", \"Segoe UI\", Arial, sans-serif",
    overflowX: "hidden",
  },
  layout: { display: "grid", gridTemplateColumns: "236px 1fr", minHeight: "100vh", maxWidth: "100%" },
  sidebar: {
    background: "linear-gradient(180deg, rgba(255,255,255,0.98), rgba(246,249,252,0.96))",
    borderRight: `1px solid ${cardBorder}`,
    padding: 18,
    backdropFilter: "blur(16px)",
    position: "sticky",
    top: 0,
    alignSelf: "start",
    height: "100dvh",
    maxHeight: "100vh",
    overflowY: "auto",
    overflowX: "hidden",
    minWidth: 0,
  },
  main: { padding: 22, minWidth: 0 },
  topbar: {
    display: "grid",
    gap: 16,
    marginBottom: 8,
    padding: 20,
    borderRadius: 24,
    border: `1px solid rgba(217, 225, 236, 0.95)`,
    background: "linear-gradient(135deg, rgba(255,255,255,0.98), rgba(244,248,253,0.96))",
    boxShadow: "0 18px 42px rgba(23, 32, 51, 0.08)",
    backdropFilter: "blur(18px)",
  },
  card: {
    background: cardBg,
    border: `1px solid ${cardBorder}`,
    borderRadius: 18,
    boxShadow: "0 14px 34px rgba(23, 32, 51, 0.06)",
    backdropFilter: "blur(14px)",
  },
  kpiCard: {
    background: "linear-gradient(180deg, rgba(255,255,255,0.99), rgba(246,249,253,0.94))",
    border: `1px solid ${cardBorder}`,
    borderRadius: 18,
    padding: 16,
    boxShadow: "0 12px 28px rgba(23, 32, 51, 0.06)",
  },
  brandPanel: {
    borderRadius: 20,
    padding: 18,
    border: `1px solid rgba(217, 225, 236, 0.95)`,
    background: "linear-gradient(160deg, rgba(255,255,255,0.99), rgba(241,246,255,0.92))",
    boxShadow: "0 14px 30px rgba(23, 32, 51, 0.07)",
  },
  brandMark: {
    width: 42,
    height: 42,
    borderRadius: 14,
    display: "grid",
    placeItems: "center",
    background: "linear-gradient(135deg, #172033, #2358d5)",
    color: "white",
    boxShadow: "0 12px 22px rgba(35, 88, 213, 0.24)",
    flexShrink: 0,
  },
  heroGrid: {
    display: "grid",
    gap: 16,
    alignItems: "start",
  },
  heroAside: {
    borderRadius: 20,
    padding: 18,
    border: "1px solid rgba(29, 95, 208, 0.12)",
    background: "linear-gradient(160deg, rgba(23,32,51,0.98), rgba(29,95,208,0.96))",
    color: "white",
    boxShadow: "0 18px 36px rgba(23, 32, 51, 0.16)",
    alignSelf: "start",
  },
  heroStat: {
    padding: "12px 14px",
    borderRadius: 14,
    border: "1px solid rgba(255,255,255,0.12)",
    background: "rgba(255,255,255,0.08)",
    backdropFilter: "blur(10px)",
  },
  softStat: {
    padding: "10px 12px",
    borderRadius: 14,
    border: `1px solid ${cardBorder}`,
    background: "linear-gradient(180deg, rgba(255,255,255,0.97), rgba(246,249,253,0.92))",
  },
  input: {
    width: "100%",
    padding: "10px 12px",
    borderRadius: 12,
    border: `1px solid ${cardBorder}`,
    background: inputBg,
    color: textMain,
    fontWeight: 600,
    fontSize: 14,
    boxSizing: "border-box",
    outline: "none",
    boxShadow: "inset 0 1px 0 rgba(255,255,255,0.82)",
  },
  btnPrimary: {
    background: "linear-gradient(135deg, #172033, #2358d5)",
    color: "white",
    border: "none",
    borderRadius: 12,
    padding: "10px 14px",
    fontWeight: 800,
    fontSize: 14,
    cursor: "pointer",
    boxShadow: "0 12px 22px rgba(35, 88, 213, 0.22)",
    transition: "transform 0.16s ease, box-shadow 0.16s ease, opacity 0.16s ease",
  },
  btnSecondary: {
    background: "rgba(255,255,255,0.88)",
    color: textMain,
    border: `1px solid ${cardBorder}`,
    borderRadius: 12,
    padding: "10px 14px",
    fontWeight: 800,
    fontSize: 14,
    cursor: "pointer",
    boxShadow: "0 8px 18px rgba(23, 32, 51, 0.04)",
    transition: "transform 0.16s ease, box-shadow 0.16s ease, border-color 0.16s ease",
  },
  topbarActions: {
    display: "flex",
    gap: 8,
    flexWrap: "wrap",
  },
  sectionHeader: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, marginBottom: 14 },
  badge: { display: "inline-flex", alignItems: "center", gap: 6, padding: "7px 11px", borderRadius: 999, fontWeight: 800, fontSize: 11, letterSpacing: 0.22 },
  fieldBlock: { display: "flex", flexDirection: "column", gap: 8 },
  fieldLabel: { fontSize: 11, fontWeight: 800, color: textSoft, letterSpacing: 0.42, textTransform: "uppercase" },
  sectionEyebrow: { fontSize: 11, color: accent, fontWeight: 800, letterSpacing: 0.6, textTransform: "uppercase" },
};

function getDecisionStyle(decision) {
  if (["SCALE", "GOOD PRODUCT", "In Stock", "Arrived", "OK"].includes(decision)) {
    return { ...styles.badge, background: "#ecfdf5", color: green, border: "1px solid #bbf7d0" };
  }
  if (["WATCH", "TEST", "SOON", "Low Stock", "Pending"].includes(decision)) {
    return { ...styles.badge, background: "#fffbeb", color: amber, border: "1px solid #fde68a" };
  }
  return { ...styles.badge, background: "#fef2f2", color: red, border: "1px solid #fecaca" };
}

function getAlertBadgeStyle(tone = "warning") {
  if (tone === "danger") {
    return { ...styles.badge, background: "#fef2f2", color: red, border: "1px solid #fecaca" };
  }
  if (tone === "success") {
    return { ...styles.badge, background: "#ecfdf5", color: green, border: "1px solid #bbf7d0" };
  }
  return { ...styles.badge, background: "#fffbeb", color: amber, border: "1px solid #fde68a" };
}

function getProductPerformanceStatus(profit, winnerThresholdTzs) {
  if (Number(profit || 0) < 0) return "LOSS";
  if (Number(profit || 0) > Number(winnerThresholdTzs || 0)) return "WINNER";
  return "TESTING";
}

function SidebarItem({ active, icon, label, onClick }) {
  const [isHovered, setIsHovered] = useState(false);

  return (
    <button
      onClick={onClick}
      onMouseEnter={() => setIsHovered(true)}
      onMouseLeave={() => setIsHovered(false)}
      style={{
        display: "flex",
        alignItems: "center",
        gap: 12,
        padding: "10px 12px",
        borderRadius: 14,
        background: active
          ? "linear-gradient(135deg, rgba(35, 88, 213, 0.14), rgba(35, 88, 213, 0.05))"
          : isHovered
            ? "rgba(255,255,255,0.78)"
            : "transparent",
        color: active ? accent : isHovered ? textMain : textSoft,
        fontWeight: active ? 800 : 700,
        fontSize: 14,
        marginBottom: 6,
        width: "100%",
        border: active ? `1px solid rgba(35, 88, 213, 0.16)` : "1px solid transparent",
        cursor: "pointer",
        textAlign: "left",
        transform: isHovered ? "translateX(2px)" : "translateX(0)",
        boxShadow: active
          ? "0 10px 18px rgba(35, 88, 213, 0.12)"
          : isHovered
            ? "0 8px 18px rgba(23, 32, 51, 0.05)"
            : "none",
        transition: "all 0.18s ease",
      }}
    >
      <span
        style={{
          display: "inline-flex",
          width: 32,
          height: 32,
          alignItems: "center",
          justifyContent: "center",
          borderRadius: 10,
          background: active ? "rgba(35, 88, 213, 0.12)" : "rgba(255,255,255,0.74)",
          border: active ? "1px solid rgba(35, 88, 213, 0.12)" : `1px solid ${cardBorder}`,
          transform: isHovered ? "scale(1.04)" : "scale(1)",
          transition: "transform 0.18s ease",
          flexShrink: 0,
        }}
      >
        {icon}
      </span>
      <span className="nav-label">{label}</span>
    </button>
  );
}

function KpiCard({ icon, title, value, sub, valueColor = textMain }) {
  return (
    <div style={{ ...styles.kpiCard, position: "relative", overflow: "hidden" }}>
      <div
        style={{
          position: "absolute",
          inset: "0 auto auto 0",
          width: 84,
          height: 84,
          borderRadius: 28,
          background: "radial-gradient(circle, rgba(29,95,208,0.14), transparent 68%)",
          pointerEvents: "none",
        }}
      />
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 12 }}>
        <div style={{ color: textSoft, fontSize: 11, fontWeight: 800, letterSpacing: 0.55, textTransform: "uppercase" }}>{title}</div>
        <div
          style={{
            color: accent,
            width: 36,
            height: 36,
            borderRadius: 12,
            display: "grid",
            placeItems: "center",
            background: "linear-gradient(135deg, rgba(35, 88, 213, 0.14), rgba(35, 88, 213, 0.05))",
          }}
        >
          {icon}
        </div>
      </div>
      <div className="kpi-value-text" style={{ fontSize: 22, fontWeight: 900, color: valueColor, lineHeight: "1.05" }}>{value}</div>
      <div style={{ marginTop: 7, color: textSoft, fontSize: 12, lineHeight: 1.45, overflowWrap: "break-word" }}>{sub}</div>
    </div>
  );
}

function MiniStat({ label, value, tone = "blue", sub = null, dark = false }) {
  const palettes = {
    blue: {
      background: dark ? "rgba(255,255,255,0.08)" : "linear-gradient(135deg, rgba(29,95,208,0.12), rgba(29,95,208,0.04))",
      border: dark ? "1px solid rgba(255,255,255,0.12)" : "1px solid rgba(29,95,208,0.1)",
      valueColor: dark ? "white" : accent,
      labelColor: dark ? "rgba(255,255,255,0.72)" : textSoft,
    },
    green: {
      background: dark ? "rgba(255,255,255,0.08)" : "linear-gradient(135deg, rgba(31,143,95,0.14), rgba(31,143,95,0.04))",
      border: dark ? "1px solid rgba(255,255,255,0.12)" : "1px solid rgba(31,143,95,0.12)",
      valueColor: dark ? "white" : green,
      labelColor: dark ? "rgba(255,255,255,0.72)" : textSoft,
    },
    amber: {
      background: dark ? "rgba(255,255,255,0.08)" : "linear-gradient(135deg, rgba(199,131,34,0.16), rgba(199,131,34,0.05))",
      border: dark ? "1px solid rgba(255,255,255,0.12)" : "1px solid rgba(199,131,34,0.14)",
      valueColor: dark ? "white" : amber,
      labelColor: dark ? "rgba(255,255,255,0.72)" : textSoft,
    },
  };
  const palette = palettes[tone] || palettes.blue;

  return (
    <div style={{ padding: "12px 14px", borderRadius: 14, background: palette.background, border: palette.border }}>
      <div style={{ fontSize: 10, fontWeight: 800, letterSpacing: 0.48, textTransform: "uppercase", color: palette.labelColor }}>{label}</div>
      <div style={{ marginTop: 7, fontSize: 21, fontWeight: 900, color: palette.valueColor, lineHeight: 1.05 }}>{value}</div>
      {sub ? <div style={{ marginTop: 5, color: palette.labelColor, fontSize: 11.5, lineHeight: 1.4 }}>{sub}</div> : null}
    </div>
  );
}

function MetaDateRangePicker({ value, onApply, responsiveColumns }) {
  const [isOpen, setIsOpen] = useState(false);
  const [draftRange, setDraftRange] = useState(value);
  const [leftMonth, setLeftMonth] = useState(() => startOfMonth(parseDateInput(value.start) || new Date()));
  const triggerRef = useRef(null);
  const showSecondMonth = responsiveColumns("show", "hide", "hide") === "show";

  const rightMonth = useMemo(() => addMonths(leftMonth, 1), [leftMonth]);

  const applyPreset = (preset) => {
    const nextRange = preset.getRange();
    setDraftRange(nextRange);
    setLeftMonth(startOfMonth(parseDateInput(nextRange.start) || new Date()));
  };

  const selectDate = (dateString) => {
    if (!draftRange.start || (draftRange.start && draftRange.end)) {
      setDraftRange({ start: dateString, end: "" });
      return;
    }

    if (dateString < draftRange.start) {
      setDraftRange({ start: dateString, end: draftRange.start });
      return;
    }

    setDraftRange({ start: draftRange.start, end: dateString });
  };

  const updateDraftBoundary = (field, nextValue) => {
    const safeValue = String(nextValue || "").trim();
    if (!safeValue) {
      setDraftRange((prev) => ({ ...prev, [field]: "" }));
      return;
    }

    setDraftRange((prev) => {
      const nextRange = { ...prev, [field]: safeValue };

      if (field === "start" && nextRange.end && nextRange.end < safeValue) {
        nextRange.end = safeValue;
      }

      if (field === "end" && nextRange.start && safeValue < nextRange.start) {
        nextRange.start = safeValue;
      }

      return nextRange;
    });

    const parsed = parseDateInput(safeValue);
    if (parsed) {
      setLeftMonth(startOfMonth(parsed));
    }
  };

  const renderMonth = (monthDate) => {
    const days = buildCalendarMatrix(monthDate);
    const currentMonth = monthDate.getMonth();

    return (
      <div style={{ display: "grid", gap: 10 }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 10, fontWeight: 800, color: textMain }}>
          <span>{monthDate.toLocaleString("en-GB", { month: "short" })}</span>
          <ChevronDown size={14} />
          <span>{monthDate.getFullYear()}</span>
          <ChevronDown size={14} />
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(7, minmax(28px, 1fr))", gap: 6, color: textSoft, fontSize: 12, textAlign: "center" }}>
          {["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"].map((day) => (
            <div key={day}>{day}</div>
          ))}
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(7, minmax(28px, 1fr))", gap: 6 }}>
          {days.map((day) => {
            const dateString = formatDateInput(day);
            const isCurrentMonth = day.getMonth() === currentMonth;
            const isStart = draftRange.start === dateString;
            const isEnd = draftRange.end === dateString;
            const isBetween = draftRange.start && draftRange.end && dateString > draftRange.start && dateString < draftRange.end;

            return (
              <button
                key={dateString}
                type="button"
                onClick={() => selectDate(dateString)}
                style={{
                  border: "none",
                  borderRadius: 10,
                  height: 34,
                  cursor: "pointer",
                  fontWeight: isStart || isEnd ? 800 : 600,
                  color: isStart || isEnd ? "white" : isCurrentMonth ? textMain : "#9ca3af",
                  background: isStart || isEnd ? accent : isBetween ? "rgba(29,95,208,0.12)" : "transparent",
                }}
              >
                {day.getDate()}
              </button>
            );
          })}
        </div>
      </div>
    );
  };

  const canApply = Boolean(draftRange.start && draftRange.end);
  const popupWidth =
    typeof window === "undefined"
      ? 980
      : Math.min(980, Math.max(320, window.innerWidth - 32));
  const popupMaxHeight =
    typeof window !== "undefined"
      ? Math.max(320, window.innerHeight - 32)
      : 760;

  return (
    <div style={{ position: "relative" }}>
      <button
        type="button"
        ref={triggerRef}
        style={{
          ...styles.input,
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          gap: 12,
          textAlign: "left",
          cursor: "pointer",
          padding: "14px 16px",
          borderRadius: 16,
          background: "linear-gradient(180deg, rgba(255,255,255,0.97), rgba(248,244,238,0.9))",
        }}
        onClick={() => setIsOpen((prev) => !prev)}
      >
        <span style={{ display: "inline-flex", alignItems: "center", gap: 10 }}>
          <CalendarDays size={16} color={textSoft} />
          <span style={{ fontWeight: 700, color: textMain }}>
            {draftRange.start && draftRange.end ? `${formatLongDate(draftRange.start)} - ${formatLongDate(draftRange.end)}` : "Select date range"}
          </span>
        </span>
        <ChevronDown size={16} color={textSoft} />
      </button>

      {isOpen ? (
        <div
          style={{
            position: "fixed",
            inset: 0,
            zIndex: 1200,
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            padding: 16,
            background: "rgba(23,32,51,0.16)",
            backdropFilter: "blur(6px)",
          }}
          onClick={() => setIsOpen(false)}
        >
          <div
            style={{
              width: popupWidth,
              maxHeight: popupMaxHeight,
              overflowY: "auto",
              padding: 18,
              borderRadius: 24,
              border: `1px solid ${cardBorder}`,
              background: "linear-gradient(180deg, rgba(255,255,255,0.99), rgba(247,243,237,0.97))",
              boxShadow: "0 30px 60px rgba(23,32,51,0.16)",
            }}
            onClick={(e) => e.stopPropagation()}
          >
            <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("220px 1fr", "1fr", "1fr"), gap: 18 }}>
              <div style={{ paddingRight: 12, borderRight: showSecondMonth ? `1px solid ${cardBorder}` : "none" }}>
                <div style={{ fontWeight: 800, marginBottom: 14 }}>Recently used</div>
                <div style={{ display: "grid", gap: 6 }}>
                  {META_RANGE_PRESETS.map((preset) => (
                    <button
                      key={preset.label}
                      type="button"
                      onClick={() => applyPreset(preset)}
                      style={{
                        textAlign: "left",
                        padding: "10px 12px",
                        borderRadius: 12,
                        border: "none",
                        background: "transparent",
                        cursor: "pointer",
                        color: textMain,
                        fontWeight: 600,
                      }}
                    >
                      {preset.label}
                    </button>
                  ))}
                </div>
              </div>

              <div style={{ display: "grid", gap: 16 }}>
                <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12 }}>
                  <button type="button" style={{ ...styles.btnSecondary, padding: "10px 12px", borderRadius: 14 }} onClick={() => setLeftMonth((prev) => addMonths(prev, -1))}>
                    <ChevronLeft size={16} />
                  </button>
                  <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("1fr 1fr", "1fr", "1fr"), gap: 18, width: "100%" }}>
                    {renderMonth(leftMonth)}
                    {showSecondMonth ? renderMonth(rightMonth) : null}
                  </div>
                  <button type="button" style={{ ...styles.btnSecondary, padding: "10px 12px", borderRadius: 14 }} onClick={() => setLeftMonth((prev) => addMonths(prev, 1))}>
                    <ChevronRight size={16} />
                  </button>
                </div>

                <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("1fr auto 1fr", "1fr auto 1fr", "1fr"), gap: 10, alignItems: "center" }}>
                  <input
                    style={styles.input}
                    type="date"
                    value={draftRange.start || ""}
                    onChange={(e) => updateDraftBoundary("start", e.target.value)}
                    placeholder="Start date"
                  />
                  <div style={{ color: textSoft, fontWeight: 800, textAlign: "center" }}>-</div>
                  <input
                    style={styles.input}
                    type="date"
                    value={draftRange.end || ""}
                    onChange={(e) => updateDraftBoundary("end", e.target.value)}
                    placeholder="End date"
                  />
                </div>

                <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap" }}>
                  <div style={{ color: textSoft, fontSize: 12 }}>Dates are shown in Casablanca time</div>
                  <div style={{ display: "flex", gap: 10 }}>
                    <button type="button" style={{ ...styles.btnSecondary, borderRadius: 16 }} onClick={() => { setDraftRange(value); setIsOpen(false); }}>
                      Cancel
                    </button>
                    <button
                      type="button"
                      style={{ ...styles.btnPrimary, borderRadius: 16, opacity: canApply ? 1 : 0.5 }}
                      disabled={!canApply}
                      onClick={() => {
                        onApply(draftRange);
                        setIsOpen(false);
                      }}
                    >
                      Update
                    </button>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      ) : null}
    </div>
  );
}

function generateMappingCode(name, existingProducts, excludeId) {
  const words = String(name || "").trim().split(/\s+/).filter(Boolean);
  const seenInitials = new Set();
  let code = "";

  for (const word of words) {
    const parts = word.split("-").filter(Boolean);
    for (const part of parts) {
      const letterMatches = part.match(/[a-zA-Z]+/g);
      const digitMatches = part.match(/\d+/g);
      const letters = letterMatches ? letterMatches.join("") : "";
      const digits = digitMatches ? digitMatches.join("") : "";

      if (letters) {
        const isAcronym = letters === letters.toUpperCase() && letters.length > 1;
        const isShort = letters.length <= 2;
        const token = (isAcronym || isShort) ? letters.toUpperCase() : letters[0].toUpperCase();
        if (token.length === 1 && seenInitials.has(token)) continue;
        token.split("").forEach((ch) => seenInitials.add(ch));
        code += token + digits;
      } else if (digits) {
        code += digits;
      }
    }
  }

  if (!code) return "";

  const taken = (Array.isArray(existingProducts) ? existingProducts : [])
    .filter((p) => (excludeId ? p.id !== excludeId : true))
    .map((p) => (p.mappingCode || "").toUpperCase());

  if (!taken.includes(code)) return code;
  let n = 2;
  while (taken.includes(`${code}${n}`)) n++;
  return `${code}${n}`;
}

export default function App() {
  const ordersImportInputRef = useRef(null);
  const shippingImportInputRef = useRef(null);
  const restoreJsonInputRef = useRef(null);
  const selectAllCustomersRef = useRef(null);
  const selectAllShippingRef = useRef(null);
  const metaAutoSyncLockRef = useRef(false);
  const metaSpendBootstrapRef = useRef("");
  const sharedSyncLockRef = useRef(false);
  const sharedHydratingRef = useRef(false);
  const sharedVersionRef = useRef(0);
  const lastSharedPayloadRef = useRef("");
  const normalizedInitRef = useRef(false);
  const latestSharedStateRef = useRef({});
  const queuedSharedSnapshotRef = useRef(null);
  const initialBrowserSnapshotRef = useRef(null);
  if (initialBrowserSnapshotRef.current === null) {
    initialBrowserSnapshotRef.current = readLocalWorkspaceSnapshotFromStorage();
  }
  const initialBrowserSnapshot = initialBrowserSnapshotRef.current;
  const [viewportWidth, setViewportWidth] = useState(() =>
    typeof window === "undefined" ? 1280 : window.innerWidth
  );
  const [activePage, setActivePage] = useState("executive");
  const showExecutiveAdminTools = activePage === "settingsAudit" && supabaseEnabled;
  const [ordersTab, setOrdersTab] = useState("pipeline");
  const [shippingTab, setShippingTab] = useState("queue");

  const [trackingSubTab, setTrackingSubTab] = useState("meta");
  const [adsCampaignsData, setAdsCampaignsData] = useState({ available: false, campaigns: [], lastLoaded: null });
  const [supabaseAdsSpendByProduct, setSupabaseAdsSpendByProduct] = useState({});
  const [profitOverviewDirect, setProfitOverviewDirect] = useState(null);
  const [revenueImport, setRevenueImport] = useState(null);
  const [revenueImportRows, setRevenueImportRows] = useState({});
  const [revenueImportNotice, setRevenueImportNotice] = useState("");
  const [manualAdsSpend, setManualAdsSpend] = useState([]);
  const [manualAdsForm, setManualAdsForm] = useState({ weekStart: "", weekEnd: "", productId: "", productName: "", amountUsd: "", notes: "" });
  const [manualAdsNotice, setManualAdsNotice] = useState("");
  const [extraCharges, setExtraCharges] = useState([]);
  const [extraChargesForm, setExtraChargesForm] = useState({ date: "", category: "other", description: "", amountUsd: "", amountTsh: "" });
  const [extraChargesNotice, setExtraChargesNotice] = useState("");
  const [profitTab, setProfitTab] = useState("overview");
  const [ownerInjections, setOwnerInjections] = useState([]);
  const [ownerInjectionForm, setOwnerInjectionForm] = useState({ date: "", amountUsd: "", amountTsh: "", notes: "" });
  const [ownerInjectionNotice, setOwnerInjectionNotice] = useState("");
  const [simProductName, setSimProductName] = useState("");
  const [simInputs, setSimInputs] = useState({ totalLeads: "", confirmationRate: "", deliveryRate: "", cpl: "", sellingPriceTsh: "", productCostUsd: "", serviceFeePerUnit: "9" });
  const [migrateNotice, setMigrateNotice] = useState("");
  const [migrating, setMigrating] = useState(false);
  const [syncNotice, setSyncNotice] = useState("");
  const [clearProductsConfirm, setClearProductsConfirm] = useState(false);
  const [settingsAuditTab, setSettingsAuditTab] = useState("workspace");
  const [showCloudBackups, setShowCloudBackups] = useState(false);
  const [_aiBriefExpanded, _setAiBriefExpanded] = useState(false);
  const [_aiAssistantPrompt, _setAiAssistantPrompt] = useState("what-happened");
  const [_aiAssistantQuestion, _setAiAssistantQuestion] = useState("");
  const [selectedService, setSelectedService] = useState("standard");
  const [selectedCountry, setSelectedCountry] = useState("tanzania");
  const [serviceForm, setServiceForm] = useState(() =>
    sanitizeServiceForm(initialBrowserSnapshot?.serviceForm || getDefaultServiceForm())
  );
  const [situationData, setSituationData] = useState(() => {
    if (supabaseEnabled) return sanitizeSituationData(initialBrowserSnapshot?.situationData || getDefaultSituationData());
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      return raw ? sanitizeSituationData(JSON.parse(raw).situationData) : getDefaultSituationData();
    } catch {
      return getDefaultSituationData();
    }
  });
  const [adInputDrafts, setAdInputDrafts] = useState({});
  const [expeditionForm, setExpeditionForm] = useState(getEmptyExpeditionForm);
  const [editingProductId, setEditingProductId] = useState(null);
  const [showAddProductForm, setShowAddProductForm] = useState(false);
  const [customerForm, setCustomerForm] = useState(getEmptyCustomerForm(initialProducts[0]?.id || "P001"));
  const [overviewFilters, setOverviewFilters] = useState({
    productId: "all",
    periodMode: "all",
    startDate: "",
    endDate: "",
  });
  const [confirmationSummaryFilters, setConfirmationSummaryFilters] = useState({
    period: "thisWeek",
    productId: "all",
    startDate: "",
    endDate: "",
  });
  const [productDetailsFilters, setProductDetailsFilters] = useState({
    period: "last7Days",
    productId: "all",
    startDate: "",
    endDate: "",
    rowLimit: 10,
  });
  const [dashboardDateFilter, setDashboardDateFilter] = useState(() =>
    createPageDateFilterState()
  );
  const [productPerformanceDateFilter, setProductPerformanceDateFilter] = useState(() =>
    createPageDateFilterState()
  );
  const [lastAutoBackupAt, setLastAutoBackupAt] = useState(() => {
    try {
      const raw = localStorage.getItem(AUTO_BACKUP_META_KEY);
      return raw ? JSON.parse(raw).lastAutoBackupAt || null : null;
    } catch {
      return null;
    }
  });
  const [ordersImportNotice, setOrdersImportNotice] = useState("");
  const [ordersImportDetails, setOrdersImportDetails] = useState(null);
  const [ordersImportHistory, setOrdersImportHistory] = useState([]);
  const [selectedLeadId, setSelectedLeadId] = useState("");
  const [shippingImportNotice, setShippingImportNotice] = useState("");
  const [shippingImportDetails, setShippingImportDetails] = useState(null);
  const [importMeta, setImportMeta] = useState(() => {
    if (supabaseEnabled) return initialBrowserSnapshot?.importMeta || getDefaultImportMeta();
    try {
      const raw = localStorage.getItem(IMPORT_META_KEY);
      return raw ? JSON.parse(raw) : getDefaultImportMeta();
    } catch {
      return getDefaultImportMeta();
    }
  });

  useEffect(() => {
    if (activePage === "dashboard") {
      setActivePage("executive");
      return;
    }
    if (activePage === "ordersHub") {
      setActivePage("customersOrders");
      return;
    }
    if (activePage === "catalogHub") {
      setActivePage("products");
      return;
    }
    if (activePage === "financeHub") {
      setActivePage("tracking");
      return;
    }
    if (activePage === "performanceHub") {
      setActivePage("taskCenter");
      return;
    }
    if (activePage === "operationsHub") {
      setActivePage("settingsAudit");
      return;
    }
    if (activePage === "catalogHub") {
      setActivePage("products");
    }
  }, [activePage]);
  const [metaAdsState, setMetaAdsState] = useState(() => {
    if (supabaseEnabled) return sanitizeMetaAdsState(initialBrowserSnapshot?.metaAdsState || getDefaultMetaAdsState());
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      return raw ? sanitizeMetaAdsState(JSON.parse(raw).metaAdsState) : getDefaultMetaAdsState();
    } catch {
      return getDefaultMetaAdsState();
    }
  });
  const [metaAdsAccounts, setMetaAdsAccounts] = useState([]);
  const [metaAdsInsights, setMetaAdsInsights] = useState(null);
  const [metaAdsNotice, setMetaAdsNotice] = useState("");
  const [metaAdsLoading, setMetaAdsLoading] = useState({ accounts: false, insights: false, apply: false });
  const [cloudAuth, setCloudAuth] = useState({
    loading: supabaseEnabled,
    ready: !supabaseEnabled,
    user: null,
    session: null,
    email: "",
    password: "",
    mode: "signin",
    notice: supabaseEnabled ? "Online mode ready. Sign in to open the cloud workspace." : "Supabase is not configured yet.",
  });
  const [sharedWorkspace, setSharedWorkspace] = useState({
    mode: "local",
    available: false,
    loading: false,
    saving: false,
    initialized: false,
    version: 0,
    updatedAt: null,
    notice: "Local workspace mode",
  });
  const [cloudBackupState, setCloudBackupState] = useState({
    loading: false,
    restoringId: null,
    available: true,
    items: [],
    notice: "",
  });
  const [cloudBackupOpen, setCloudBackupOpen] = useState(false);
  const [currentTime, setCurrentTime] = useState(() => Date.now());
  const [customerListFilters, setCustomerListFilters] = useState({
    search: "",
    status: "all",
    productId: "all",
    city: "all",
    pageSize: 25,
  });
  const [customerListPage, setCustomerListPage] = useState(1);
  const [selectedCustomerIds, setSelectedCustomerIds] = useState([]);
  const [selectedShippingIds, setSelectedShippingIds] = useState([]);
  const [bulkCustomerStatus, setBulkCustomerStatus] = useState("confirmed");
  const [bulkCustomerOwner, setBulkCustomerOwner] = useState("");
  const [bulkShippingStatus, setBulkShippingStatus] = useState("shipped");
  const [customerHistoryTargetId, setCustomerHistoryTargetId] = useState("");
  const [shippingListFilters, setShippingListFilters] = useState({
    search: "",
    status: "all",
    pageSize: 25,
  });
  const [shippingListPage, setShippingListPage] = useState(1);
  const [auditSearch, setAuditSearch] = useState("");
  const [stockTab, setStockTab] = useState("overview");
  const [stockPurchases, setStockPurchases] = useState(() => {
    if (supabaseEnabled) return Array.isArray(initialBrowserSnapshot?.stockPurchases) ? initialBrowserSnapshot.stockPurchases : [];
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      return raw ? (JSON.parse(raw).stockPurchases || []) : [];
    } catch { return []; }
  });
  const [stockMovements, setStockMovements] = useState(() => {
    if (supabaseEnabled) return Array.isArray(initialBrowserSnapshot?.stockMovements) ? initialBrowserSnapshot.stockMovements : [];
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      return raw ? (JSON.parse(raw).stockMovements || []) : [];
    } catch { return []; }
  });
  const [purchaseForm, setPurchaseForm] = useState({
    product_id: "", quantity_ordered: "", source_country: "dubai", supplier_name: "",
    purchase_date: "", expected_arrival_date: "", usable_stock_date: "",
    buy_price_per_unit_usd: "", shipping_cost_usd: "", sourcing_cost_usd: "",
    other_charges_tsh: "", quantity_received: "", status: "ordered", notes: "",
  });
  const [editingPurchaseId, setEditingPurchaseId] = useState(null);
  const [manualAdjForm, setManualAdjForm] = useState({ product_id: "", quantity_change: "", reason: "stock_count_correction", note: "" });
  const [receiveStockInput, setReceiveStockInput] = useState({ purchaseId: null, qty: "", notes: "" });

  const [products, setProducts] = useState(() => {
    if (supabaseEnabled) return Array.isArray(initialBrowserSnapshot?.products) ? initialBrowserSnapshot.products.map(sanitizeProductRecord) : [];
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      return raw ? (JSON.parse(raw).products || initialProducts).map(sanitizeProductRecord) : initialProducts.map(sanitizeProductRecord);
    } catch {
      return initialProducts.map(sanitizeProductRecord);
    }
  });

  const [tracking, setTracking] = useState(() => {
    if (supabaseEnabled) return Array.isArray(initialBrowserSnapshot?.tracking) ? initialBrowserSnapshot.tracking : [];
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      return raw ? JSON.parse(raw).tracking || initialTracking : initialTracking;
    } catch {
      return initialTracking;
    }
  });

  const [customers, setCustomers] = useState(() => {
    if (supabaseEnabled) return Array.isArray(initialBrowserSnapshot?.customers) ? initialBrowserSnapshot.customers.map(sanitizeCustomerRecord) : [];
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      return raw ? (JSON.parse(raw).customers || INITIAL_CUSTOMERS).map(sanitizeCustomerRecord) : INITIAL_CUSTOMERS;
    } catch {
      return INITIAL_CUSTOMERS;
    }
  });

  const buildSharedStateSnapshot = useCallback(
    () => ({
      products,
      tracking,
      customers,
      serviceForm,
      situationData,
      metaAdsState,
      importMeta,
      stockPurchases,
      stockMovements,
    }),
    [customers, importMeta, metaAdsState, products, serviceForm, situationData, tracking, stockPurchases, stockMovements]
  );

  useEffect(() => {
    latestSharedStateRef.current = {
      products,
      tracking,
      customers,
      serviceForm,
      situationData,
      metaAdsState,
      importMeta,
      stockPurchases,
      stockMovements,
    };
  }, [customers, importMeta, metaAdsState, products, serviceForm, situationData, tracking, stockPurchases, stockMovements]);

  const readBrowserBackupSnapshot = useCallback(() => readLocalWorkspaceSnapshotFromStorage(), []);

  const applySharedStateSnapshot = useCallback((snapshot = {}) => {
    const cloudDefaults = getDefaultCloudWorkspaceState();
    const localBaseSnapshot =
      (hasMeaningfulWorkspaceData(latestSharedStateRef.current) && latestSharedStateRef.current) ||
      readBrowserBackupSnapshot() ||
      null;
    const fallbackSnapshot = localBaseSnapshot
      ? {
          ...cloudDefaults,
          ...localBaseSnapshot,
        }
      : supabaseEnabled
        ? cloudDefaults
        : {
            products: initialProducts.map(sanitizeProductRecord),
            tracking: [...initialTracking],
            customers: INITIAL_CUSTOMERS.map(sanitizeCustomerRecord),
            serviceForm: getDefaultServiceForm(),
            situationData: getDefaultSituationData(),
            metaAdsState: getDefaultMetaAdsState(),
            importMeta: getDefaultImportMeta(),
          };
    const normalizedSnapshot = {
      products: Array.isArray(snapshot.products) ? snapshot.products.map(sanitizeProductRecord) : fallbackSnapshot.products,
      tracking: Array.isArray(snapshot.tracking) ? snapshot.tracking : fallbackSnapshot.tracking,
      customers: Array.isArray(snapshot.customers) ? snapshot.customers.map(sanitizeCustomerRecord) : fallbackSnapshot.customers,
      serviceForm: sanitizeServiceForm(snapshot.serviceForm || fallbackSnapshot.serviceForm),
      situationData: sanitizeSituationData(snapshot.situationData || fallbackSnapshot.situationData),
      metaAdsState: sanitizeMetaAdsState(snapshot.metaAdsState || fallbackSnapshot.metaAdsState),
      importMeta: {
        lastOrdersImportAt: snapshot.importMeta?.lastOrdersImportAt || null,
        lastShippingImportAt: snapshot.importMeta?.lastShippingImportAt || null,
      },
      stockPurchases: Array.isArray(snapshot.stockPurchases) ? snapshot.stockPurchases : [],
      stockMovements: Array.isArray(snapshot.stockMovements) ? snapshot.stockMovements : [],
    };

    sharedHydratingRef.current = true;
    try {
      setProducts(normalizedSnapshot.products);
      setTracking(normalizedSnapshot.tracking);
      setCustomers(normalizedSnapshot.customers);
      setServiceForm(normalizedSnapshot.serviceForm);
      setSituationData(normalizedSnapshot.situationData);
      setMetaAdsState(normalizedSnapshot.metaAdsState);
      setImportMeta(normalizedSnapshot.importMeta);
      setStockPurchases(normalizedSnapshot.stockPurchases);
      setStockMovements(normalizedSnapshot.stockMovements);
    } finally {
      window.setTimeout(() => {
        sharedHydratingRef.current = false;
      }, 0);
    }
  }, [readBrowserBackupSnapshot]);

  const handleMigrateLocalToCloud = useCallback(async () => {
    if (!cloudAuth.user) {
      setMigrateNotice("Sign in to your cloud account first.");
      return;
    }
    const localSnapshot = readLocalWorkspaceSnapshotFromStorage();
    if (!localSnapshot || !hasMeaningfulWorkspaceData(localSnapshot)) {
      setMigrateNotice("No local data found to migrate.");
      return;
    }
    setMigrating(true);
    setMigrateNotice("Migrating data to cloud...");
    try {
      await saveCloudWorkspace(localSnapshot, {
        workspaceId: supabaseWorkspaceId,
        userId: cloudAuth.user.id,
        backupReason: "localStorage-migration",
      });
      const pCount = localSnapshot.products?.length || 0;
      const oCount = localSnapshot.customers?.length || 0;
      const tCount = localSnapshot.tracking?.length || 0;
      setMigrateNotice(`Migration complete — ${pCount} products, ${oCount} orders, ${tCount} tracking rows pushed to cloud.`);
      lastSharedPayloadRef.current = "";
      applySharedStateSnapshot(localSnapshot);
    } catch (error) {
      setMigrateNotice(`Migration failed: ${error instanceof Error ? error.message : "Unknown error"}`);
    } finally {
      setMigrating(false);
    }
  }, [cloudAuth.user, applySharedStateSnapshot]);

  useEffect(() => {
    if (supabaseEnabled) return;
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return;
      const parsed = JSON.parse(raw);
      if (parsed.serviceForm) setServiceForm(sanitizeServiceForm(parsed.serviceForm));
      if (Array.isArray(parsed.customers)) setCustomers(parsed.customers.map(sanitizeCustomerRecord));
      if (parsed.situationData) setSituationData(sanitizeSituationData(parsed.situationData));
      if (parsed.metaAdsState) setMetaAdsState(sanitizeMetaAdsState(parsed.metaAdsState));
    } catch {
      // ignore restore issue
    }
  }, []);

  useEffect(() => {
    if (!supabaseEnabled) return undefined;

    let cancelled = false;
    setCloudAuth((prev) => ({ ...prev, loading: true, ready: false }));

    getCloudSession()
      .then(({ session, user }) => {
        if (cancelled) return;
        setCloudAuth((prev) => ({
          ...prev,
          loading: false,
          ready: true,
          session,
          user,
          notice: user ? `Cloud access connected as ${user.email || "user"}` : "Sign in to use the online shared workspace.",
        }));
      })
      .catch((error) => {
        if (cancelled) return;
        setCloudAuth((prev) => ({
          ...prev,
          loading: false,
          ready: true,
          notice: error instanceof Error ? error.message : "Unable to initialize cloud access.",
        }));
      });

    const unsubscribe = onCloudAuthStateChange(({ session, user }) => {
      if (cancelled) return;
      setCloudAuth((prev) => ({
        ...prev,
        loading: false,
        ready: true,
        session,
        user,
        notice: user ? `Cloud access connected as ${user.email || "user"}` : "Sign in to use the online shared workspace.",
      }));
    });

    return () => {
      cancelled = true;
      unsubscribe();
    };
  }, []);

  useEffect(() => {
    let cancelled = false;

    const loadSharedWorkspace = async () => {
      if (supabaseEnabled) {
        if (!cloudAuth.ready) return;
        if (!cloudAuth.user) {
          setSharedWorkspace({
            mode: "cloud",
            available: false,
            loading: false,
            saving: false,
            initialized: true,
            version: 0,
            updatedAt: null,
            notice: "Cloud login required",
          });
          return;
        }
      }

      setSharedWorkspace((prev) => ({ ...prev, loading: true }));
      try {
        let payload;
        if (supabaseEnabled && cloudAuth.user) {
          const cloudPayload = await loadCloudWorkspace(supabaseWorkspaceId);
          payload = { ok: true, version: cloudPayload.version, updatedAt: cloudPayload.updatedAt, state: cloudPayload.state || {} };
        } else {
          const response = await fetch(getSharedApiBase(), { headers: { Accept: "application/json" } });
          const remotePayload = await response.json().catch(() => ({}));
          if (!response.ok || !remotePayload?.ok) throw new Error(remotePayload?.error || "Unable to load shared workspace.");
          payload = remotePayload;
        }
        if (cancelled) return;
        const remoteState = payload.state || {};
        const cloudMode = supabaseEnabled && cloudAuth.user;
        const remoteHasData = hasMeaningfulWorkspaceData(remoteState);
        const localSnapshot = latestSharedStateRef.current || {};
        const localHasData = hasMeaningfulWorkspaceData(localSnapshot);
        const browserBackupSnapshot = readBrowserBackupSnapshot();
        const recoveredFromBrowserBackup = !remoteHasData && !localHasData && Boolean(browserBackupSnapshot);
        const shouldKeepLocal = !remoteHasData && localHasData;

        if (cloudMode) {
          let resolvedState = remoteState;
          let recoveredFromNormalized = false;

          if (remoteHasData) {
            console.log("[Supabase] workspace blob loaded:", remoteState.products?.length, "products,", remoteState.customers?.length, "orders");
            lastSharedPayloadRef.current = JSON.stringify(remoteState);
          } else {
            // workspace blob empty — try normalized tables as fallback
            console.log("[Supabase] workspace blob empty, checking normalized tables...");
            let normalizedSnapshot = null;
            try {
              normalizedSnapshot = await loadWorkspaceFromNormalizedTables(supabaseWorkspaceId);
              console.log("[Supabase] normalized tables:", normalizedSnapshot?.products?.length ?? 0, "products,", normalizedSnapshot?.customers?.length ?? 0, "orders,", normalizedSnapshot?.tracking?.length ?? 0, "tracking rows");
            } catch (e) {
              console.warn("[Supabase] normalized tables load failed:", e.message);
            }

            if (normalizedSnapshot && hasMeaningfulWorkspaceData(normalizedSnapshot)) {
              resolvedState = normalizedSnapshot;
              recoveredFromNormalized = true;
              lastSharedPayloadRef.current = "";
              // Write back to workspace blob so future loads are fast
              try {
                await saveCloudWorkspace(resolvedState, {
                  workspaceId: supabaseWorkspaceId,
                  userId: cloudAuth.user.id,
                  backupReason: "table-recovery",
                });
                console.log("[Supabase] workspace blob restored from normalized tables");
              } catch (e) {
                console.warn("[Supabase] blob restore failed:", e.message);
              }
            } else if (shouldKeepLocal) {
              resolvedState = localSnapshot;
              lastSharedPayloadRef.current = "";
            } else if (browserBackupSnapshot) {
              resolvedState = browserBackupSnapshot;
              lastSharedPayloadRef.current = "";
            } else {
              lastSharedPayloadRef.current = JSON.stringify(remoteState);
            }
          }

          // Merge locally-saved products not yet in remote (race condition: user adds product before workspace loads)
          if (browserBackupSnapshot?.products?.length > 0) {
            const resolvedProductIds = new Set((resolvedState.products || []).map((p) => p.id));
            const localOnlyProducts = browserBackupSnapshot.products.filter((p) => p.id && !resolvedProductIds.has(p.id));
            if (localOnlyProducts.length > 0) {
              resolvedState = { ...resolvedState, products: [...(resolvedState.products || []), ...localOnlyProducts] };
              lastSharedPayloadRef.current = "";
            }
          }
          if (browserBackupSnapshot?.stockPurchases?.length > 0) {
            const resolvedPurchaseIds = new Set((resolvedState.stockPurchases || []).map((p) => p.id));
            const localOnlyPurchases = browserBackupSnapshot.stockPurchases.filter((p) => p.id && !resolvedPurchaseIds.has(p.id));
            if (localOnlyPurchases.length > 0) {
              resolvedState = { ...resolvedState, stockPurchases: [...(resolvedState.stockPurchases || []), ...localOnlyPurchases] };
              lastSharedPayloadRef.current = "";
            }
          }

          applySharedStateSnapshot(resolvedState);
          sharedVersionRef.current = Number(payload.version || 0);
          setSharedWorkspace({
            mode: "cloud",
            available: true,
            loading: false,
            saving: false,
            initialized: true,
            version: Number(payload.version || 0),
            updatedAt: payload.updatedAt || null,
            notice: recoveredFromNormalized
              ? "Cloud workspace restored from tables"
              : recoveredFromBrowserBackup
                ? "Recovered cloud data from browser backup"
                : shouldKeepLocal
                  ? "Cloud workspace was empty - keeping local data"
                  : "Cloud workspace connected",
          });
          if (!normalizedInitRef.current && supabaseEnabled) {
            normalizedInitRef.current = true;
            const migSource = hasMeaningfulWorkspaceData(resolvedState) ? resolvedState : null;
            if (migSource && !recoveredFromNormalized) {
              checkNormalizedTablesEmpty(supabaseWorkspaceId)
                .then((isEmpty) => {
                  if (!isEmpty) return null;
                  return migrateWorkspaceToNormalizedTables(migSource, supabaseWorkspaceId);
                })
                .then((res) => {
                  if (res?.success) setSyncNotice(`Data synced to cloud ✓ — ${res.counts.products} products, ${res.counts.orders} orders`);
                })
                .catch(() => null);
            }
          }
          return;
        }

        const remoteLooksFresh =
          !remoteHasData &&
          Number(payload.version || 0) <= 0 &&
          !payload.updatedAt;
        let recoveredFromAutoBackup = false;

        if (!cloudMode && remoteLooksFresh && !localHasData) {
          if (browserBackupSnapshot) {
            applySharedStateSnapshot(browserBackupSnapshot);
            recoveredFromAutoBackup = true;
          }
        }

        const shouldKeepLocalShared = remoteLooksFresh && localHasData;

        if (!shouldKeepLocalShared && !recoveredFromAutoBackup) {
          applySharedStateSnapshot(remoteState);
        }
        sharedVersionRef.current = Number(payload.version || 0);
        lastSharedPayloadRef.current = shouldKeepLocalShared || recoveredFromAutoBackup ? "" : JSON.stringify(remoteState);
        setSharedWorkspace({
          mode: supabaseEnabled ? "cloud" : "shared",
          available: true,
          loading: false,
          saving: false,
          initialized: true,
          version: Number(payload.version || 0),
          updatedAt: payload.updatedAt || null,
          notice: recoveredFromAutoBackup
            ? "Recovered from browser auto backup"
            : shouldKeepLocalShared
              ? `${supabaseEnabled ? "Cloud" : "Shared"} workspace was empty - keeping local data`
              : `${supabaseEnabled ? "Cloud" : "Shared"} workspace connected`,
        });
      } catch {
        if (cancelled) return;
        setSharedWorkspace({
          mode: supabaseEnabled ? "cloud" : "local",
          available: false,
          loading: false,
          saving: false,
          initialized: true,
          version: 0,
          updatedAt: null,
          notice: supabaseEnabled ? "Cloud workspace unavailable" : "Local workspace mode",
        });
      }
    };

    loadSharedWorkspace();
    return () => {
      cancelled = true;
    };
  }, [applySharedStateSnapshot, cloudAuth.ready, cloudAuth.user, readBrowserBackupSnapshot]);

  useEffect(() => {
    try {
      localStorage.setItem(IMPORT_META_KEY, JSON.stringify(importMeta));
    } catch {
      // ignore browser quota issues for import metadata
    }
  }, [importMeta]);

  useEffect(() => {
    if (typeof window === "undefined") return undefined;

    const handleResize = () => setViewportWidth(window.innerWidth);
    window.addEventListener("resize", handleResize);
    return () => window.removeEventListener("resize", handleResize);
  }, []);

  useEffect(() => {
    const interval = window.setInterval(() => setCurrentTime(Date.now()), 300000);
    return () => window.clearInterval(interval);
  }, []);

  useEffect(() => {
    const nextSnapshot = {
      products,
      tracking,
      customers,
      serviceForm,
      situationData,
      metaAdsState,
      importMeta,
      stockPurchases,
      stockMovements,
    };
    const nextSnapshotHasData = hasMeaningfulWorkspaceData(nextSnapshot);
    const existingBrowserBackup = readLocalWorkspaceSnapshotFromStorage();
    if (!nextSnapshotHasData && existingBrowserBackup) {
      return;
    }

    persistBrowserSnapshotSafely(nextSnapshot, {
      exportedAt: new Date().toISOString(),
      onSaved: (now) => setLastAutoBackupAt(now),
    });
  }, [customers, importMeta, metaAdsState, products, readBrowserBackupSnapshot, serviceForm, situationData, tracking, stockPurchases, stockMovements]);

  useEffect(() => {
    const bucket = getDayBucket(currentTime);
    if (!bucket) return;

    const totalSpendTzs = tracking.reduce((sum, row) => sum + Number(row.adSpend || 0), 0);
    const observedByProduct = tracking.reduce((acc, row) => {
      const productId = String(row.productId || "");
      if (!productId) return acc;
      acc[productId] = (acc[productId] || 0) + Number(row.adSpend || 0);
      return acc;
    }, {});

    setSituationData((prev) => {
      const existing = Array.isArray(prev.hourlyAdsSnapshots) ? prev.hourlyAdsSnapshots : [];
      const hasBucket = existing.some((entry) => entry.bucket === bucket);
      if (hasBucket) return prev;

      const lastObservedTotal = Number(prev.lastObservedAdsSpendTzs || 0);
      const deltaTotal = Math.max(0, totalSpendTzs - lastObservedTotal);
      const nextCumulativeByProduct = { ...(prev.cumulativeAdsByProduct || {}) };
      Object.entries(observedByProduct).forEach(([productId, amount]) => {
        const previousObserved = Number(prev.lastObservedAdsByProduct?.[productId] || 0);
        const delta = Math.max(0, Number(amount || 0) - previousObserved);
        if (delta > 0) {
          nextCumulativeByProduct[productId] = Number(nextCumulativeByProduct[productId] || 0) + delta;
        }
      });

      const nextHourlyAdsSnapshots = [
        {
          id: `daily-ads-${Date.now()}`,
          bucket,
          totalSpendTzs,
          capturedAt: new Date().toISOString(),
          source: "tracking",
        },
        ...existing,
      ].slice(0, 168);

      return {
        ...prev,
        hourlyAdsSnapshots: nextHourlyAdsSnapshots,
        cumulativeAdsTotalTzs: Number(prev.cumulativeAdsTotalTzs || 0) + deltaTotal,
        cumulativeAdsByProduct: nextCumulativeByProduct,
        lastObservedAdsSpendTzs: totalSpendTzs,
        lastObservedAdsByProduct: observedByProduct,
        lastAdsAccumulatedAt: new Date().toISOString(),
      };
    });
  }, [currentTime, tracking]);

  const persistSharedSnapshot = useCallback(
    async function persistSharedSnapshotInner(
      nextSnapshot,
      {
        progressNotice = supabaseEnabled ? "Saving cloud changes..." : "Saving shared workspace...",
        successNotice = supabaseEnabled ? "Cloud workspace synced" : "Shared workspace synced",
        failurePrefix = supabaseEnabled ? "Cloud workspace sync failed" : "Shared workspace sync failed",
      } = {}
    ) {
      if (!sharedWorkspace.initialized) return true;

      const serialized = JSON.stringify(nextSnapshot || {});
      if (serialized === lastSharedPayloadRef.current) return true;

      if (sharedSyncLockRef.current) {
        queuedSharedSnapshotRef.current = {
          snapshot: nextSnapshot,
          options: { progressNotice, successNotice, failurePrefix },
        };
        return false;
      }

      sharedSyncLockRef.current = true;
      latestSharedStateRef.current = nextSnapshot;
      lastSharedPayloadRef.current = serialized;
      setSharedWorkspace((prev) => ({
        ...prev,
        mode: supabaseEnabled ? "cloud" : "shared",
        available: true,
        loading: false,
        saving: true,
        notice: progressNotice,
      }));

      try {
        let payload;
        if (supabaseEnabled && cloudAuth.user) {
          const saved = await saveCloudWorkspace(nextSnapshot, {
            workspaceId: supabaseWorkspaceId,
            userId: cloudAuth.user.id,
          });
          payload = { ok: true, version: saved.version, updatedAt: saved.updatedAt, backup: saved.backup || null };
        } else if (supabaseEnabled) {
          const saved = await saveCloudWorkspaceAnon(nextSnapshot, { workspaceId: supabaseWorkspaceId });
          payload = { ok: true, version: saved.version, updatedAt: saved.updatedAt, backup: null };
        } else {
          const response = await fetch(getSharedApiBase(), {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ state: nextSnapshot }),
          });
          const remotePayload = await response.json().catch(() => ({}));
          if (!response.ok || !remotePayload?.ok) {
            throw new Error(remotePayload?.error || "Unable to save shared workspace.");
          }
          payload = remotePayload;
        }

        sharedVersionRef.current = Number(payload.version || sharedVersionRef.current || 0);
        lastSharedPayloadRef.current = serialized;
        setSharedWorkspace((prev) => ({
          ...prev,
          mode: supabaseEnabled ? "cloud" : "shared",
          available: true,
          loading: false,
          saving: false,
          initialized: true,
          version: Number(payload.version || 0),
          updatedAt: payload.updatedAt || prev.updatedAt || null,
          notice: successNotice,
        }));
        if (supabaseEnabled && payload.backup) {
          setCloudBackupState((prev) => {
            if (!payload.backup.available) {
              return {
                ...prev,
                available: false,
                notice: payload.backup.notice || "Run the latest Supabase schema to enable restore history.",
              };
            }

            if (!payload.backup.saved || !payload.backup.entry) {
              return payload.backup.notice
                ? {
                    ...prev,
                    available: true,
                    notice: payload.backup.notice,
                  }
                : prev;
            }

            const nextItems = [payload.backup.entry, ...prev.items.filter((item) => item.id !== payload.backup.entry.id)].slice(0, 8);
            return {
              ...prev,
              available: true,
              items: nextItems,
              notice: "",
            };
          });
        }
        if (supabaseEnabled && cloudAuth.user) {
          void syncNormalizedTables(nextSnapshot, supabaseWorkspaceId);
        }
        return true;
      } catch (error) {
        lastSharedPayloadRef.current = "";
        setSharedWorkspace((prev) => ({
          ...prev,
          mode: prev.available ? (supabaseEnabled ? "cloud" : "shared") : "local",
          available: prev.available,
          loading: false,
          saving: false,
          notice: `${failurePrefix}${error instanceof Error && error.message ? `: ${error.message}` : ""}`,
        }));
        return false;
      } finally {
        sharedSyncLockRef.current = false;
        const queued = queuedSharedSnapshotRef.current;
        queuedSharedSnapshotRef.current = null;
        if (queued) {
          window.setTimeout(() => {
            void persistSharedSnapshotInner(queued.snapshot, queued.options);
          }, 0);
        }
      }
    },
    [cloudAuth.user, sharedWorkspace.initialized]
  );

  useEffect(() => {
    if (!sharedWorkspace.initialized) return;
    if (sharedHydratingRef.current) return;

    const snapshot = buildSharedStateSnapshot();
    latestSharedStateRef.current = snapshot;

    const timeout = window.setTimeout(() => {
      void persistSharedSnapshot(snapshot);
    }, 180);

    return () => {
      window.clearTimeout(timeout);
    };
  }, [buildSharedStateSnapshot, cloudAuth.user, persistSharedSnapshot, sharedWorkspace.initialized]);

  useEffect(() => {
    let cancelled = false;

    const pollSharedWorkspace = async () => {
      try {
        if (supabaseEnabled) {
          if (!cloudAuth.user) return;

          const payload = await loadCloudWorkspace(supabaseWorkspaceId);
          const remoteState = payload.state || {};
          const remoteVersion = Number(payload.version || 0);
          const remoteSerialized = JSON.stringify(remoteState);
          const localSnapshot = latestSharedStateRef.current || {};
          const remoteHasData = hasMeaningfulWorkspaceData(remoteState);
          const localHasData = hasMeaningfulWorkspaceData(localSnapshot);

          if (!cancelled) {
            setSharedWorkspace((prev) => ({
              ...prev,
              mode: "cloud",
              available: true,
              version: remoteVersion,
              updatedAt: payload.updatedAt || prev.updatedAt,
              notice:
                prev.notice === "Cloud workspace unavailable" || prev.notice === "Cloud workspace sync delayed"
                  ? "Cloud workspace connected"
                  : prev.notice,
            }));
          }

          if (!remoteHasData && localHasData && !sharedSyncLockRef.current) {
            lastSharedPayloadRef.current = "";
            void persistSharedSnapshot(localSnapshot, {
              progressNotice: "Restoring cloud workspace data...",
              successNotice: "Cloud workspace restored",
              failurePrefix: "Cloud workspace restore failed",
            });
            return;
          }

          if (
            !sharedSyncLockRef.current &&
            (remoteVersion > Number(sharedVersionRef.current || 0) || remoteSerialized !== lastSharedPayloadRef.current)
          ) {
            if (cancelled) return;
            applySharedStateSnapshot(remoteState);
            sharedVersionRef.current = remoteVersion;
            lastSharedPayloadRef.current = remoteSerialized;
            setSharedWorkspace((prev) => ({
              ...prev,
              mode: "cloud",
              available: true,
              version: remoteVersion,
              updatedAt: payload.updatedAt || prev.updatedAt,
              notice: "Cloud workspace updated",
            }));
          }

          return;
        }

        const metaResponse = await fetch(`${getSharedApiBase()}/meta`, { headers: { Accept: "application/json" } });
        const metaPayload = await metaResponse.json().catch(() => ({}));
        if (!metaResponse.ok || !metaPayload?.ok) throw new Error(metaPayload?.error || "Unable to check shared workspace.");

        const remoteVersion = Number(metaPayload.version || 0);
        if (!cancelled) {
          setSharedWorkspace((prev) => ({
            ...prev,
            mode: "shared",
            available: true,
            version: remoteVersion,
            updatedAt: metaPayload.updatedAt || prev.updatedAt,
            notice: prev.notice === "Local workspace mode" ? "Shared workspace connected" : prev.notice,
          }));
        }

          if (remoteVersion > Number(sharedVersionRef.current || 0) && !sharedSyncLockRef.current) {
            const response = await fetch(getSharedApiBase(), { headers: { Accept: "application/json" } });
            const payload = await response.json().catch(() => ({}));
            if (!response.ok || !payload?.ok) throw new Error(payload?.error || "Unable to refresh shared workspace.");
            if (cancelled) return;
            const remoteState = payload.state || {};
            const localSnapshot = latestSharedStateRef.current || {};
            if (!hasMeaningfulWorkspaceData(remoteState) && hasMeaningfulWorkspaceData(localSnapshot)) {
              lastSharedPayloadRef.current = "";
              void persistSharedSnapshot(localSnapshot, {
                progressNotice: "Restoring shared workspace data...",
                successNotice: "Shared workspace restored",
                failurePrefix: "Shared workspace restore failed",
              });
              return;
            }
            applySharedStateSnapshot(remoteState);
            sharedVersionRef.current = Number(payload.version || remoteVersion);
            lastSharedPayloadRef.current = JSON.stringify(remoteState);
            setSharedWorkspace((prev) => ({
              ...prev,
              mode: "shared",
            available: true,
            version: Number(payload.version || remoteVersion),
            updatedAt: payload.updatedAt || prev.updatedAt,
            notice: "Shared workspace updated",
          }));
        }
      } catch {
        if (cancelled) return;
        setSharedWorkspace((prev) => ({
          ...prev,
          mode: prev.available ? prev.mode : supabaseEnabled ? "cloud" : "local",
          available: prev.available,
          notice: supabaseEnabled ? "Cloud workspace sync delayed" : "Local workspace mode",
        }));
      }
    };

    pollSharedWorkspace();
    const interval = window.setInterval(pollSharedWorkspace, supabaseEnabled ? 5000 : 15000);
    return () => {
      cancelled = true;
      window.clearInterval(interval);
    };
  }, [applySharedStateSnapshot, cloudAuth.user, persistSharedSnapshot]);

  const refreshCloudBackups = useCallback(
    async ({ silent = false } = {}) => {
      if (!supabaseEnabled || !cloudAuth.user) {
        setCloudBackupState({
          loading: false,
          restoringId: null,
          available: true,
          items: [],
          notice: "",
        });
        return;
      }

      if (!silent) {
        setCloudBackupState((prev) => ({
          ...prev,
          loading: true,
          notice: prev.notice && !prev.available ? prev.notice : "",
        }));
      }

      try {
        const response = await listCloudWorkspaceBackups(supabaseWorkspaceId, 8);
        setCloudBackupState((prev) => ({
          ...prev,
          loading: false,
          available: response.available,
          items: response.items || [],
          notice: response.notice || "",
        }));
      } catch (error) {
        setCloudBackupState((prev) => ({
          ...prev,
          loading: false,
          available: false,
          notice: error instanceof Error ? error.message : "Unable to load cloud restore history.",
        }));
      }
    },
    [cloudAuth.user]
  );

  const restoreCloudBackup = useCallback(
    async (backupId) => {
      if (!cloudAuth.user) return;
      const targetBackup = cloudBackupState.items.find((item) => item.id === backupId);
      const backupLabel = targetBackup?.created_at ? new Date(targetBackup.created_at).toLocaleString() : `#${backupId}`;
      if (!window.confirm(`Restore the workspace from backup ${backupLabel}? Current live data will be replaced by that saved version.`)) {
        return;
      }

      setCloudBackupState((prev) => ({
        ...prev,
        restoringId: backupId,
        notice: "",
      }));
      setSharedWorkspace((prev) => ({
        ...prev,
        notice: "Restoring cloud backup...",
        saving: true,
      }));

      try {
        await restoreCloudWorkspaceBackup(backupId, {
          workspaceId: supabaseWorkspaceId,
          userId: cloudAuth.user.id,
        });

        const restoredWorkspace = await loadCloudWorkspace(supabaseWorkspaceId);
        if (restoredWorkspace?.state) {
          applySharedStateSnapshot(restoredWorkspace.state);
          latestSharedStateRef.current = restoredWorkspace.state;
          lastSharedPayloadRef.current = JSON.stringify(restoredWorkspace.state || {});
          sharedVersionRef.current = Number(restoredWorkspace.version || sharedVersionRef.current || 0);
        }

        setSharedWorkspace((prev) => ({
          ...prev,
          saving: false,
          initialized: true,
          available: true,
          mode: "cloud",
          version: Number(restoredWorkspace?.version || prev.version || 0),
          updatedAt: restoredWorkspace?.updatedAt || prev.updatedAt || null,
          notice: "Cloud backup restored",
        }));
        await refreshCloudBackups({ silent: true });
      } catch (error) {
        setSharedWorkspace((prev) => ({
          ...prev,
          saving: false,
          notice: `Cloud backup restore failed${error instanceof Error && error.message ? `: ${error.message}` : ""}`,
        }));
        setCloudBackupState((prev) => ({
          ...prev,
          notice: error instanceof Error ? error.message : "Unable to restore cloud backup.",
        }));
      } finally {
        setCloudBackupState((prev) => ({
          ...prev,
          restoringId: null,
        }));
      }
    },
    [applySharedStateSnapshot, cloudAuth.user, cloudBackupState.items, refreshCloudBackups]
  );

  useEffect(() => {
    if (!supabaseEnabled || !cloudAuth.user) {
      setCloudBackupState({
        loading: false,
        restoringId: null,
        available: true,
        items: [],
        notice: "",
      });
      return;
    }

    void refreshCloudBackups({ silent: false });
  }, [cloudAuth.user, refreshCloudBackups]);

  useEffect(() => {
    if (!supabaseEnabled || !cloudAuth.user) return undefined;

    const unsubscribe = subscribeToCloudWorkspace(supabaseWorkspaceId, (payload) => {
      const remoteState = payload?.state || {};
      const remoteVersion = Number(payload?.version || 0);
      const remoteSerialized = JSON.stringify(remoteState);
      if (
        sharedSyncLockRef.current ||
        (remoteVersion <= Number(sharedVersionRef.current || 0) && remoteSerialized === lastSharedPayloadRef.current)
      ) {
        return;
      }
      const localSnapshot = latestSharedStateRef.current || {};
      if (!hasMeaningfulWorkspaceData(remoteState) && hasMeaningfulWorkspaceData(localSnapshot)) {
        lastSharedPayloadRef.current = "";
        void persistSharedSnapshot(localSnapshot, {
          progressNotice: "Restoring cloud workspace data...",
          successNotice: "Cloud workspace restored",
          failurePrefix: "Cloud workspace restore failed",
        });
        return;
      }
      applySharedStateSnapshot(remoteState);
      sharedVersionRef.current = remoteVersion;
      lastSharedPayloadRef.current = remoteSerialized;
      setSharedWorkspace((prev) => ({
        ...prev,
        mode: "cloud",
        available: true,
        version: remoteVersion,
        updatedAt: payload.updatedAt || prev.updatedAt,
        notice: "Cloud workspace updated live",
      }));
    });

    return () => {
      unsubscribe();
    };
  }, [applySharedStateSnapshot, cloudAuth.user, persistSharedSnapshot]);

  const persistStockSnapshot = useCallback(
    async (nextPurchases, nextMovements) => {
      const nextSnapshot = {
        ...(latestSharedStateRef.current || getDefaultCloudWorkspaceState()),
        stockPurchases: nextPurchases,
        stockMovements: nextMovements,
      };
      latestSharedStateRef.current = nextSnapshot;
      return persistSharedSnapshot(nextSnapshot, {
        progressNotice: "Saving stock changes...",
        successNotice: "Stock data synced",
        failurePrefix: "Stock sync failed",
      });
    },
    [persistSharedSnapshot]
  );

  const getManualSeedPurchaseId = useCallback((productId) => `manual-stock-${productId}`, []);

  const upsertManualSeedStockPurchase = useCallback(
    (currentPurchases, productRecord) => {
      if (!productRecord?.id) return currentPurchases;

      const productId = productRecord.id;
      const manualSeedId = getManualSeedPurchaseId(productId);
      const exchangeRate = Number(serviceForm?.exchangeRate || USD_TO_TZS) || USD_TO_TZS;
      const quantity = Math.max(0, Number(productRecord.totalQty || 0));
      const sourceCountry = String(productRecord.source || "china").toLowerCase();
      const nowIso = new Date().toISOString();
      const today = getTodayString();
      const existingSeed = currentPurchases.find((purchase) => purchase.id === manualSeedId) || null;
      const hasRealPurchases = currentPurchases.some(
        (purchase) => purchase.product_id === productId && purchase.id !== manualSeedId
      );

      if (hasRealPurchases) {
        return existingSeed
          ? currentPurchases.filter((purchase) => purchase.id !== manualSeedId)
          : currentPurchases;
      }

      if (quantity <= 0) {
        return existingSeed
          ? currentPurchases.filter((purchase) => purchase.id !== manualSeedId)
          : currentPurchases;
      }

      const shippingCostUsd = Math.max(0, Number(productRecord.shippingTotal || 0) / exchangeRate);
      const otherChargesTsh = Math.max(0, Number(productRecord.otherCharges || 0));
      const otherChargesUsd = otherChargesTsh > 0 ? otherChargesTsh / exchangeRate : 0;
      const buyPricePerUnitUsd = Math.max(0, Number(productRecord.purchaseUnitPrice || 0));
      const totalBuyCostUsd = buyPricePerUnitUsd * quantity;
      const totalLandedCostUsd =
        totalBuyCostUsd +
        shippingCostUsd +
        Math.max(0, Number(productRecord.sourcingCostUsd || 0)) +
        otherChargesUsd;

      const seedPurchase = {
        id: manualSeedId,
        product_id: productId,
        quantity_ordered: quantity,
        quantity_received: quantity,
        source_country: sourceCountry,
        supplier_name: productRecord.supplierName || "",
        purchase_date: existingSeed?.purchase_date || productRecord.stockOrderedAt || today,
        expected_arrival_date: "",
        usable_stock_date: existingSeed?.usable_stock_date || productRecord.stockArrivedAt || today,
        buy_price_per_unit_usd: buyPricePerUnitUsd,
        shipping_cost_usd: shippingCostUsd,
        sourcing_cost_usd: Math.max(0, Number(productRecord.sourcingCostUsd || 0)),
        other_charges_tsh: otherChargesTsh,
        other_charges_usd: otherChargesUsd,
        total_buy_cost_usd: totalBuyCostUsd,
        total_landed_cost_usd: totalLandedCostUsd,
        landed_cost_per_unit_usd: quantity > 0 ? totalLandedCostUsd / quantity : 0,
        status: "received",
        notes: "Auto-generated from product stock quantity",
        created_at: existingSeed?.created_at || nowIso,
        updated_at: nowIso,
      };

      if (existingSeed) {
        return currentPurchases.map((purchase) => (purchase.id === manualSeedId ? seedPurchase : purchase));
      }

      return [...currentPurchases, seedPurchase];
    },
    [getManualSeedPurchaseId, serviceForm]
  );

  const addStockMovement = useCallback((movement, currentMovements) => {
    const entry = {
      movement_id: `mv-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
      date: new Date().toISOString().slice(0, 10),
      created_at: new Date().toISOString(),
      ...movement,
    };
    return [...currentMovements, entry];
  }, []);

  const savePurchase = useCallback(() => {
    const f = purchaseForm;
    if (!f.product_id) return;
    const qty = Math.max(0, Number(f.quantity_ordered) || 0);
    const buyPriceUsd = Math.max(0, Number(f.buy_price_per_unit_usd) || 0);
    const shippingUsd = Math.max(0, Number(f.shipping_cost_usd) || 0);
    const sourcingUsd = Math.max(0, Number(f.sourcing_cost_usd) || 0);
    const otherTsh = Math.max(0, Number(f.other_charges_tsh) || 0);
    const exchangeRate = Number(serviceForm?.exchangeRate || USD_TO_TZS);
    const otherUsd = otherTsh > 0 ? otherTsh / exchangeRate : 0;
    const totalBuyCostUsd = qty * buyPriceUsd;
    const totalLandedCostUsd = totalBuyCostUsd + shippingUsd + sourcingUsd + otherUsd;
    const landedCostPerUnitUsd = qty > 0 ? totalLandedCostUsd / qty : 0;

    const isEdit = Boolean(editingPurchaseId);
    const purchaseId = editingPurchaseId || `pur-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`;
    const now = new Date().toISOString();

    const record = {
      id: purchaseId,
      product_id: f.product_id,
      quantity_ordered: qty,
      quantity_received: Number(f.quantity_received) || 0,
      source_country: f.source_country || "dubai",
      supplier_name: f.supplier_name || "",
      purchase_date: f.purchase_date || now.slice(0, 10),
      expected_arrival_date: f.expected_arrival_date || "",
      usable_stock_date: f.usable_stock_date || "",
      buy_price_per_unit_usd: buyPriceUsd,
      shipping_cost_usd: shippingUsd,
      sourcing_cost_usd: sourcingUsd,
      other_charges_tsh: otherTsh,
      other_charges_usd: otherUsd,
      total_buy_cost_usd: totalBuyCostUsd,
      total_landed_cost_usd: totalLandedCostUsd,
      landed_cost_per_unit_usd: landedCostPerUnitUsd,
      status: f.status || "ordered",
      notes: f.notes || "",
      created_at: isEdit ? (stockPurchases.find((p) => p.id === purchaseId)?.created_at || now) : now,
      updated_at: now,
    };

    setStockPurchases((prev) => {
      const next = isEdit ? prev.map((p) => p.id === purchaseId ? record : p) : [...prev, record];
      setStockMovements((prevMov) => {
        const movement = {
          product_id: f.product_id,
          type: isEdit ? "stock_correction" : "purchase_ordered",
          quantity_change: qty,
          before_quantity: 0,
          after_quantity: qty,
          source_reference: purchaseId,
          note: `Purchase ${isEdit ? "updated" : "created"}: ${f.supplier_name || "supplier"}`,
        };
        const nextMov = addStockMovement(movement, prevMov);
        void persistStockSnapshot(next, nextMov);
        return nextMov;
      });
      return next;
    });
    setEditingPurchaseId(null);
    setPurchaseForm({ product_id: "", quantity_ordered: "", source_country: "dubai", supplier_name: "", purchase_date: "", expected_arrival_date: "", usable_stock_date: "", buy_price_per_unit_usd: "", shipping_cost_usd: "", sourcing_cost_usd: "", other_charges_tsh: "", quantity_received: "", status: "ordered", notes: "" });
  }, [purchaseForm, editingPurchaseId, stockPurchases, serviceForm, addStockMovement, persistStockSnapshot]);

  const deletePurchase = useCallback((purchaseId) => {
    setStockPurchases((prev) => {
      const next = prev.filter((p) => p.id !== purchaseId);
      setStockMovements((prevMov) => {
        void persistStockSnapshot(next, prevMov);
        return prevMov;
      });
      return next;
    });
  }, [persistStockSnapshot]);

  const updatePurchaseStatus = useCallback((purchaseId, newStatus, quantityReceived) => {
    setStockPurchases((prev) => {
      const purchase = prev.find((p) => p.id === purchaseId);
      if (!purchase) return prev;
      const qty = quantityReceived != null ? Math.max(0, Number(quantityReceived) || 0) : purchase.quantity_received;
      const updated = { ...purchase, status: newStatus, quantity_received: qty, updated_at: new Date().toISOString() };
      const next = prev.map((p) => p.id === purchaseId ? updated : p);
      setStockMovements((prevMov) => {
        const movement = {
          product_id: purchase.product_id,
          type: newStatus === "received" ? "purchase_received" : "stock_correction",
          quantity_change: newStatus === "received" ? qty : 0,
          before_quantity: 0,
          after_quantity: qty,
          source_reference: purchaseId,
          note: `Purchase marked as ${newStatus}${quantityReceived != null ? ` (qty: ${qty})` : ""}`,
        };
        const nextMov = addStockMovement(movement, prevMov);
        void persistStockSnapshot(next, nextMov);
        return nextMov;
      });
      return next;
    });
  }, [addStockMovement, persistStockSnapshot]);

  const markArrivedEarly = useCallback((purchaseId) => {
    setStockPurchases((prev) => {
      const purchase = prev.find((p) => p.id === purchaseId);
      if (!purchase) return prev;
      const today = getTodayString();
      const updated = { ...purchase, status: "arrived", actual_arrival_date: today, is_early_arrival: true, updated_at: new Date().toISOString() };
      const next = prev.map((p) => p.id === purchaseId ? updated : p);
      setStockMovements((prevMov) => {
        const movement = { product_id: purchase.product_id, type: "early_arrival", quantity_change: 0, before_quantity: 0, after_quantity: 0, source_reference: purchaseId, note: `Arrived early on ${today}` };
        const nextMov = addStockMovement(movement, prevMov);
        void persistStockSnapshot(next, nextMov);
        return nextMov;
      });
      return next;
    });
  }, [addStockMovement, persistStockSnapshot]);

  const receiveStockNow = useCallback((purchaseId, qtyJustReceived, receivingNotes) => {
    setStockPurchases((prev) => {
      const purchase = prev.find((p) => p.id === purchaseId);
      if (!purchase) return prev;
      const alreadyReceived = Math.max(0, Number(purchase.quantity_received) || 0);
      const ordered = Math.max(0, Number(purchase.quantity_ordered) || 0);
      const remaining = ordered - alreadyReceived;
      const qty = Math.max(0, Number(qtyJustReceived) || 0);
      if (qty <= 0 || qty > remaining) return prev;
      const newQtyReceived = alreadyReceived + qty;
      const newRemaining = ordered - newQtyReceived;
      const today = getTodayString();
      const newStatus = newRemaining === 0 ? "received" : "partially_received";
      const updated = { ...purchase, status: newStatus, quantity_received: newQtyReceived, remaining_quantity: newRemaining, received_date: today, receiving_notes: receivingNotes || purchase.receiving_notes || "", updated_at: new Date().toISOString() };
      const next = prev.map((p) => p.id === purchaseId ? updated : p);
      setStockMovements((prevMov) => {
        const movement = { product_id: purchase.product_id, type: "purchase_received", quantity_change: qty, before_quantity: alreadyReceived, after_quantity: newQtyReceived, source_reference: purchaseId, note: receivingNotes || `Received ${qty} units (${newStatus})` };
        const nextMov = addStockMovement(movement, prevMov);
        void persistStockSnapshot(next, nextMov);
        return nextMov;
      });
      return next;
    });
  }, [addStockMovement, persistStockSnapshot]);

  const saveManualAdjustment = useCallback(() => {
    const { product_id, quantity_change, reason, note } = manualAdjForm;
    if (!product_id || !quantity_change) return;
    const delta = Number(quantity_change);
    if (!Number.isFinite(delta) || delta === 0) return;
    const adjId = `adj-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`;
    setStockMovements((prev) => {
      const movement = {
        product_id,
        type: "manual_adjustment",
        quantity_change: delta,
        before_quantity: 0,
        after_quantity: 0,
        source_reference: adjId,
        note: `${reason.replace(/_/g, " ")}: ${note || "no note"}`,
      };
      const next = addStockMovement(movement, prev);
      setStockPurchases((prevPur) => { void persistStockSnapshot(prevPur, next); return prevPur; });
      return next;
    });
    setManualAdjForm({ product_id: "", quantity_change: "", reason: "stock_count_correction", note: "" });
  }, [manualAdjForm, addStockMovement, persistStockSnapshot]);

  const updateProductAdInput = useCallback((productId, field, value) => {
    setSituationData((prev) => ({
      ...prev,
      adInputs: {
        ...(prev.adInputs || {}),
        [productId]: {
          averageLeadCostTzs: Math.max(0, parseLooseNumber(prev.adInputs?.[productId]?.averageLeadCostTzs)),
          incomingLeads: Math.max(0, Math.round(parseLooseNumber(prev.adInputs?.[productId]?.incomingLeads))),
          manualAdsSpendTzs: Math.max(0, parseLooseNumber(prev.adInputs?.[productId]?.manualAdsSpendTzs)),
          leads: Math.max(0, Math.round(parseLooseNumber(prev.adInputs?.[productId]?.leads))),
          confirmedOrders: Math.max(0, Math.round(parseLooseNumber(prev.adInputs?.[productId]?.confirmedOrders))),
          deliveredOrders: Math.max(0, Math.round(parseLooseNumber(prev.adInputs?.[productId]?.deliveredOrders))),
          [field]:
            field === "manualAdsSpendTzs"
              ? Math.max(0, parseLooseNumber(value))
              : Math.max(0, Math.round(parseLooseNumber(value))),
        },
      },
    }));
  }, []);

  const updateProductAlertThreshold = useCallback((field, value) => {
    setSituationData((prev) => ({
      ...prev,
      productAlertThresholds: {
        ...(prev.productAlertThresholds || getDefaultSituationData().productAlertThresholds),
        [field]:
          field === "minDeliveryRatePct"
            ? Math.max(0, Math.min(100, parseLooseNumber(value)))
            : field === "minStockQuantity" || field === "lowDeliveredOrders"
              ? Math.max(0, Math.round(parseLooseNumber(value)))
              : Math.max(0, parseLooseNumber(value)),
      },
    }));
  }, []);

  const getProduct = useCallback((id) => products.find((p) => p.id === id), [products]);
  const responsiveColumns = useCallback(
    (desktop, tablet = "1fr 1fr", mobile = "1fr") => {
      if (viewportWidth <= 640) return mobile;
      if (viewportWidth <= 1100) return tablet;
      return desktop;
    },
    [viewportWidth]
  );
  const isCompact = viewportWidth <= 1100;

  const buildOperationalCustomerKeys = useCallback((customer) => {
    if (!customer) return [];

    const importKey = String(customer.import_key || "").trim();
    if (importKey) {
      return [`import:${importKey}`];
    }

    const sourceOrderId = String(customer.sourceOrderId || customer.order_id || "").trim();
    if (sourceOrderId) {
      const productId = String(customer.productId || "").trim();
      const lineItemIndex = Math.max(0, Number(customer.line_item_index || 0));
      return [`source:${sourceOrderId}::${productId}::${lineItemIndex}`];
    }

    const keys = [];
    const phone = normalizePhoneValue(customer.phone);
    const productId = String(customer.productId || "").trim();
    const orderDate = String(customer.orderDate || "").trim();
    const quantity = Math.max(1, Number(customer.quantity || 1));
    const customerName = normalizeHeaderName(customer.customerName);

    if (phone && productId && orderDate) {
      keys.push(`order:${phone}::${productId}::${orderDate}::${quantity}`);
    }

    if (phone && productId && customerName) {
      keys.push(`customer:${phone}::${productId}::${customerName}::${quantity}`);
    }

    return Array.from(new Set(keys.filter(Boolean)));
  }, []);

  const getOperationalCustomerFreshness = useCallback((customer) => {
    const shippingImportedAt = Date.parse(customer?.lastShippingImportedAt || "") || 0;
    const importedAt = Date.parse(customer?.lastImportedAt || "") || 0;
    const updatedAt = Date.parse(customer?.updatedAt || "") || 0;
    const actualDeliveryDate = parseDateInput(customer?.actualDeliveryDate)?.getTime() || 0;
    const orderDate = parseDateInput(customer?.orderDate)?.getTime() || 0;
    return Math.max(shippingImportedAt, importedAt, updatedAt, actualDeliveryDate, orderDate);
  }, []);

  const compareOperationalCustomers = useCallback(
    (candidate, existing) => {
      const freshnessGap =
        getOperationalCustomerFreshness(candidate) - getOperationalCustomerFreshness(existing);
      if (freshnessGap !== 0) return freshnessGap;

      const candidateSignals =
        (candidate?.sourceOrderId ? 4 : 0) +
        (candidate?.lastShippingImportedAt ? 3 : 0) +
        (candidate?.lastImportedAt ? 2 : 0) +
        (Number(candidate?.orderTotalTzs || 0) > 0 ? 1 : 0);
      const existingSignals =
        (existing?.sourceOrderId ? 4 : 0) +
        (existing?.lastShippingImportedAt ? 3 : 0) +
        (existing?.lastImportedAt ? 2 : 0) +
        (Number(existing?.orderTotalTzs || 0) > 0 ? 1 : 0);
      if (candidateSignals !== existingSignals) return candidateSignals - existingSignals;

      return String(candidate?.id || "").localeCompare(String(existing?.id || ""));
    },
    [getOperationalCustomerFreshness]
  );

  const operationalCustomers = useMemo(() => {
    const mergedCustomers = [];
    const keyToIndex = new Map();
    const keysByIndex = [];

    customers.filter(Boolean).forEach((customer) => {
      const keys = buildOperationalCustomerKeys(customer);

      if (!keys.length) {
        mergedCustomers.push(customer);
        keysByIndex.push(new Set());
        return;
      }

      const matchingIndexes = Array.from(
        new Set(keys.map((key) => keyToIndex.get(key)).filter((value) => Number.isInteger(value)))
      );

      if (!matchingIndexes.length) {
        const index = mergedCustomers.push(customer) - 1;
        const keySet = new Set(keys);
        keysByIndex[index] = keySet;
        keySet.forEach((key) => keyToIndex.set(key, index));
        return;
      }

      const targetIndex = matchingIndexes[0];
      const existing = mergedCustomers[targetIndex];
      const targetKeys = keysByIndex[targetIndex] || new Set();

      keys.forEach((key) => {
        targetKeys.add(key);
        keyToIndex.set(key, targetIndex);
      });
      keysByIndex[targetIndex] = targetKeys;

      if (compareOperationalCustomers(customer, existing) > 0) {
        mergedCustomers[targetIndex] = customer;
      }
    });

    return mergedCustomers;
  }, [buildOperationalCustomerKeys, compareOperationalCustomers, customers]);

  const resolvedOperationalCustomers = useMemo(() => {
    if (!products.length) return operationalCustomers;
    return operationalCustomers.map((customer) => {
      if (customer.productId) return customer;
      const rawName = String(customer.product_name_raw || customer.product_ref || "").trim();
      if (!rawName) return customer;
      const resolvedId = matchProductIdFromText(rawName, products);
      if (!resolvedId) return customer;
      return { ...customer, productId: resolvedId };
    });
  }, [operationalCustomers, products]);

  const serviceLeadCustomers = useMemo(
    () =>
      operationalCustomers.filter((customer) =>
        isCountableLeadForService(getCustomerConfirmationStatus(customer))
      ),
    [operationalCustomers]
  );

  const buildCustomerMetricsByProduct = useCallback(
    (customerRows = []) =>
      customerRows.reduce((acc, customer) => {
        const product = getProduct(customer.productId);
        if (!product) return acc;

        const productId = product.id;
        const confirmationStatus = getCustomerConfirmationStatus(customer);
        const shippingStatus = getCustomerShippingStatus(customer);
        const quantity = Math.max(1, Number(customer.quantity || 1));

        if (!acc[productId]) {
          acc[productId] = {
            orders: 0,
            orderedUnits: 0,
            confirmed: 0,
            confirmedUnits: 0,
            toPrepare: 0,
            toPrepareUnits: 0,
            outDelivered: 0,
            outDeliveredUnits: 0,
            shipping: 0,
            shippingUnits: 0,
            delivered: 0,
            deliveredUnits: 0,
            returned: 0,
            returnedUnits: 0,
            cancelled: 0,
            cancelledUnits: 0,
            revenue: 0,
            revenueUsd: 0,
            serviceFeeTsh: 0,
            serviceFeeUsd: 0,
            missingAmountDeliveredOrders: 0,
            estimatedRevenueOrders: 0,
            missingRegionOrders: 0,
            statusCounts: {},
          };
        }

        acc[productId].orders += 1;
        acc[productId].orderedUnits += quantity;
        const statusKey = shippingStatus || confirmationStatus;
        const shippingBucket = shippingStatus ? getShippingBucket(shippingStatus) : "";
        acc[productId].statusCounts[statusKey] = (acc[productId].statusCounts[statusKey] || 0) + 1;

        if (isConfirmationConfirmed(confirmationStatus)) {
          acc[productId].confirmed += 1;
          acc[productId].confirmedUnits += quantity;
        }
        if (shippingBucket === "to_prepare") {
          acc[productId].toPrepare += 1;
          acc[productId].toPrepareUnits += quantity;
        }
        if (shippingBucket === "shipped") {
          acc[productId].outDelivered += 1;
          acc[productId].outDeliveredUnits += quantity;
        }
        if (shippingStatus && isShippingInProgress(shippingStatus)) {
          acc[productId].shipping += 1;
          acc[productId].shippingUnits += quantity;
        }
        if (isShippingDelivered(shippingStatus)) {
          const revenueAmounts = getOrderRevenueAmounts(customer, product, USD_TO_TZS);
          const serviceFee = calculateServiceFeeForOrder(customer, { exchangeRate: USD_TO_TZS });
          acc[productId].delivered += 1;
          acc[productId].deliveredUnits += quantity;
          acc[productId].revenue += revenueAmounts.revenueTsh;
          acc[productId].revenueUsd += revenueAmounts.revenueUsd;
          acc[productId].serviceFeeTsh += serviceFee.tsh;
          acc[productId].serviceFeeUsd += serviceFee.usd;
          if (revenueAmounts.estimatedRevenueUsed) acc[productId].estimatedRevenueOrders += 1;
          if (revenueAmounts.auditFlags.includes("missing_amount_delivered_order")) acc[productId].missingAmountDeliveredOrders += 1;
          if (serviceFee.auditFlags.includes("missing_region_fee_source")) acc[productId].missingRegionOrders += 1;
        }
        if (isShippingReturned(shippingStatus)) {
          acc[productId].returned += 1;
          acc[productId].returnedUnits += quantity;
        }
        if (isConfirmationCancelled(confirmationStatus) || isShippingReturned(shippingStatus)) {
          acc[productId].cancelled += 1;
          acc[productId].cancelledUnits += quantity;
        }

        return acc;
      }, {}),
    [getProduct]
  );

  const customerMetricsByProduct = useMemo(() => {
    return buildCustomerMetricsByProduct(resolvedOperationalCustomers);
  }, [buildCustomerMetricsByProduct, resolvedOperationalCustomers]);

  const buildProductDashboardRows = useCallback(
    (metricsByProduct = {}, trackingRows = [], resolvedCustomers = [], spendByProduct = {}) =>
      products
        .map((product) => {
          const rows = trackingRows.filter((t) => t.productId === product.id);
          const customerMetrics = metricsByProduct[product.id] || {
          orders: 0,
          orderedUnits: 0,
          confirmed: 0,
          confirmedUnits: 0,
          toPrepare: 0,
          toPrepareUnits: 0,
          outDelivered: 0,
          outDeliveredUnits: 0,
          shipping: 0,
          shippingUnits: 0,
          delivered: 0,
          deliveredUnits: 0,
          returned: 0,
          returnedUnits: 0,
          cancelled: 0,
          cancelledUnits: 0,
          revenue: 0,
          revenueUsd: 0,
          serviceFeeTsh: 0,
          serviceFeeUsd: 0,
          missingAmountDeliveredOrders: 0,
          estimatedRevenueOrders: 0,
          missingRegionOrders: 0,
          statusCounts: {},
        };
          let spend = 0;
          rows.forEach((row) => {
            spend += Number(row.adSpend || 0);
          });
          // Override with cumulative Meta ads spend when available (persisted across sessions)
          const cumulativeProductSpend = spendByProduct[product.id];
          if (cumulativeProductSpend?.spendTsh > 0) {
            spend = cumulativeProductSpend.spendTsh;
          }

          const deliveredUnits = Number(customerMetrics.deliveredUnits || 0);
        const shippingUnits = Number(customerMetrics.shippingUnits || 0);
        const toPrepareUnits = Number(customerMetrics.toPrepareUnits || 0);
        const outDeliveredUnits = Number(customerMetrics.outDeliveredUnits || 0);
        const returnedUnits = Number(customerMetrics.returnedUnits || 0);
        const confirmedUnits = Number(customerMetrics.confirmedUnits || 0);
        const orderedUnits = Number(customerMetrics.orderedUnits || 0);
        const delivered = Number(customerMetrics.delivered || 0);
        const confirmed = Number(customerMetrics.confirmed || 0);
        const toPrepare = Number(customerMetrics.toPrepare || 0);
        const outDelivered = Number(customerMetrics.outDelivered || 0);
        const shipping = Number(customerMetrics.shipping || 0);
        const orders = Number(customerMetrics.orders || 0);
        const revenue = Number(customerMetrics.revenue || 0);
        const productPerformance = calculateProductPerformance(product, {
          unitsSold: deliveredUnits,
          totalRevenue: revenue,
          totalAdsSpend: spend,
          totalServiceFeeTzs: Number(customerMetrics.serviceFeeTsh || 0),
        });
        const unitProductCost = productPerformance.costPerUnitUsd;
        const unitProductCostTzs = productPerformance.costPerUnitTzs;
        const deliveryCostPerUnitTzs = productPerformance.deliveryCostPerUnitTzs;
        const profit = productPerformance.profit;
        const cpa = deliveredUnits > 0 ? spend / deliveredUnits : 0;
        const costPerLead = orders > 0 ? spend / orders : 0;
        const roas = spend > 0 ? revenue / spend : 0;
        const confirmRate = orders > 0 ? confirmed / orders : 0;
        const deliveryRate = confirmed > 0 ? delivered / confirmed : 0;
        const margin = productPerformance.profitMargin;
        const initialStock = Number(product.totalQty || 0);
        const realPurchasesForProduct = stockPurchases.filter((p) => p.product_id === product.id);
        // Synthetic purchase is used ONLY for cost/value estimation when no formal purchases exist.
        // Stock quantity calculations always use realPurchasesForProduct so that
        // available=0 when no purchases are recorded (prevents phantom stock).
        const productStockPurchasesForCost = realPurchasesForProduct.length > 0 ? realPurchasesForProduct : [{
          product_id: product.id,
          quantity_received: product.totalQty,
          quantity_ordered: product.totalQty,
          buy_price_per_unit_usd: product.purchaseUnitPrice,
          shipping_cost_usd: Number(product.shippingTotal || 0) / USD_TO_TZS,
          other_charges_tsh: product.otherCharges,
          sourcing_cost_usd: product.sourcingCostUsd,
          status: ["arrived", "received"].includes(String(product.stockArrivalStatus || "").toLowerCase()) ? "received" : "in_transit",
        }];
        const productOrders = resolvedCustomers.filter((customer) => customer.productId === product.id);
        const acceptedStock = calculateReceivedStock(product.id, realPurchasesForProduct);
        const reservedStock = calculateReservedStock(product.id, productOrders);
        const deliveredStock = calculateDeliveredStock(product.id, productOrders);
        const returnedStock = calculateReturnedStock(product.id, productOrders);
        const damagedStock = calculateDamagedStock(product.id, productOrders);
        const availableStock = calculateAvailableStock(product.id, realPurchasesForProduct, productOrders);
        const currentStock = Math.max(0, availableStock);
        const stockValue = calculateAvailableStockValue(product.id, productStockPurchasesForCost, productOrders, USD_TO_TZS);
        const incomingStock = realPurchasesForProduct.filter((p) => ["ordered", "in_transit"].includes(p.status)).reduce((s, p) => s + Math.max(0, Number(p.quantity_ordered || 0) - Number(p.quantity_received || 0)), 0);
        const salesPerDay = deliveredUnits > 0 ? deliveredUnits / 30 : 0;
        const arrivalDays = Number(product.estimatedArrivalDays || 0);
        const safetyFactor = 1.3;
        const reorderPoint = Math.ceil(salesPerDay * arrivalDays * safetyFactor);
        const reorderSoonPoint = Math.ceil(reorderPoint * 1.2);

        let decision = "WATCH";
        if (profit > 0 && deliveryRate >= 0.6) decision = "SCALE";
        if (profit < 0) decision = "KILL";

        const score = Math.max(
          0,
          Math.min(
            100,
            Math.round(
              (profit > 0 ? 40 : 0) +
                (roas >= 2 ? 25 : 0) +
                (deliveryRate >= 0.5 ? 20 : 0) +
                (confirmRate >= 0.5 ? 15 : 0)
            )
          )
        );

          let reorderStatus = "OK";
          if (availableStock <= reorderPoint) reorderStatus = "ORDER NOW";
          else if (availableStock <= reorderSoonPoint) reorderStatus = "SOON";

          return {
          ...product,
          stockQuantity: currentStock,
          costPerUnit: unitProductCostTzs,
          unitProductCost,
          unitProductCostTzs,
          deliveryCostPerUnit: deliveryCostPerUnitTzs,
          deliveryCostPerUnitTzs,
          totalUnitsSold: deliveredUnits,
          totalRevenue: revenue,
          totalAdsSpend: spend,
          totalProductCost: productPerformance.totalProductCost,
          totalProductCostTzs: productPerformance.totalProductCost,
          totalDeliveryCost: productPerformance.totalDeliveryCost,
          totalDeliveryCostTzs: productPerformance.totalDeliveryCost,
          serviceFeeTsh: Number(customerMetrics.serviceFeeTsh || 0),
          serviceFeeUsd: Number(customerMetrics.serviceFeeUsd || 0),
          totalImportCost:
            Number(product.purchaseUnitPrice || 0) * Number(product.totalQty || 0) * USD_TO_TZS +
            Number(product.shippingTotal || 0) +
            Number(product.otherCharges || 0),
          spend,
          orders,
          orderedUnits,
          confirmed,
          toPrepare,
          outDelivered,
          shipping,
          delivered,
          confirmedUnits,
          toPrepareUnits,
          outDeliveredUnits,
          shippingUnits,
          deliveredUnits,
          returnedOrders: customerMetrics.returned,
          returnedUnits,
          cancelledOrders: customerMetrics.cancelled,
          cancelledUnits: customerMetrics.cancelledUnits,
          revenue,
          profit,
          profitMargin: margin,
          cpa,
          costPerLead,
          roas,
          confirmRate,
          deliveryRate,
          margin,
          decision,
          score,
          initialStock,
          acceptedStock,
          currentStock,
          reservedStock,
          availableStock,
          deliveredStock,
          returnedStock,
          damagedStock,
          stockValueUsd: stockValue.valueUsd,
          stockValueTsh: stockValue.valueTsh,
          salesPerDay,
          reorderPoint,
          reorderSoonPoint,
          reorderStatus,
          incomingStock,
          missingAmountDeliveredOrders: Number(customerMetrics.missingAmountDeliveredOrders || 0),
          estimatedRevenueOrders: Number(customerMetrics.estimatedRevenueOrders || 0),
          missingRegionOrders: Number(customerMetrics.missingRegionOrders || 0),
          statusCounts: customerMetrics.statusCounts,
          automatedFromOrders: true,
          };
        })
        .sort((a, b) => b.score - a.score),
    [products, stockPurchases]
  );

  // Merge saved campaigns from state (always available) and Supabase table (when migrated).
  // Deduplication key: campaignId + dateStart + dateEnd — prevents double-counting same period.
  const cumulativeCampaigns = useMemo(() => {
    const all = [
      ...(adsCampaignsData.campaigns || []),
      ...(metaAdsState.savedCampaigns || []),
    ];
    const seen = new Set();
    return all.filter((c) => {
      const key = `${c.campaignId}::${c.dateStart}::${c.dateEnd}`;
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  }, [adsCampaignsData.campaigns, metaAdsState.savedCampaigns]);

  const cumulativeSpendByProduct = useMemo(() => {
    const map = {};
    for (const c of cumulativeCampaigns) {
      if (!c.productId || c.isUnmapped) continue;
      if (!map[c.productId]) map[c.productId] = { spendTsh: 0, spendUsd: 0, leads: 0, campaignCount: 0 };
      map[c.productId].spendTsh += Number(c.spendTsh || 0);
      map[c.productId].spendUsd += Number(c.spendUsd || 0);
      map[c.productId].leads += Number(c.leads || 0);
      map[c.productId].campaignCount += 1;
    }
    return map;
  }, [cumulativeCampaigns]);

  const cumulativeUnmappedSpendTsh = useMemo(
    () => cumulativeCampaigns.filter((c) => c.isUnmapped || !c.productId).reduce((s, c) => s + Number(c.spendTsh || 0), 0),
    [cumulativeCampaigns]
  );

  const totalCumulativeSpendTsh = useMemo(
    () => cumulativeCampaigns.reduce((s, c) => s + Number(c.spendTsh || 0), 0),
    [cumulativeCampaigns]
  );

  const productDashboard = useMemo(() => {
    return buildProductDashboardRows(customerMetricsByProduct, tracking, resolvedOperationalCustomers, cumulativeSpendByProduct);
  }, [buildProductDashboardRows, customerMetricsByProduct, tracking, resolvedOperationalCustomers, cumulativeSpendByProduct]);

  const bestProduct = productDashboard[0];
  const productDashboardMap = useMemo(
    () => Object.fromEntries(productDashboard.map((product) => [product.id, product])),
    [productDashboard]
  );

  const cplTrackerRows = useMemo(() => {
    return products
      .map((product) => {
        const cumulative = cumulativeSpendByProduct[product.id] || { spendTsh: 0, spendUsd: 0, leads: 0, campaignCount: 0 };
        const dashRow = productDashboardMap[product.id] || {};
        const orders = Number(dashRow.orders || 0);
        const confirmed = Number(dashRow.confirmed || 0);
        const delivered = Number(dashRow.deliveredUnits || 0);
        const leads = cumulative.leads;
        const spendUsd = cumulative.spendUsd;
        const spendTsh = cumulative.spendTsh;
        return {
          id: product.id,
          name: product.name,
          mappingCode: product.mappingCode || product.id,
          spendTsh,
          spendUsd,
          leads,
          campaignCount: cumulative.campaignCount,
          orders,
          confirmed,
          delivered,
          cpl: leads > 0 ? spendUsd / leads : 0,
          cplConfirmed: confirmed > 0 ? spendUsd / confirmed : 0,
          cplDelivered: delivered > 0 ? spendUsd / delivered : 0,
          confirmRate: orders > 0 ? confirmed / orders : 0,
          deliveryRate: confirmed > 0 ? delivered / confirmed : 0,
        };
      })
      .filter((r) => r.spendUsd > 0 || r.leads > 0);
  }, [products, cumulativeSpendByProduct, productDashboardMap]);

  const cplTrackerGlobal = useMemo(() => {
    const totalSpendUsd = cplTrackerRows.reduce((s, r) => s + r.spendUsd, 0);
    const totalSpendTsh = cplTrackerRows.reduce((s, r) => s + r.spendTsh, 0);
    const totalLeads = cplTrackerRows.reduce((s, r) => s + r.leads, 0);
    const totalOrders = cplTrackerRows.reduce((s, r) => s + r.orders, 0);
    const totalConfirmed = cplTrackerRows.reduce((s, r) => s + r.confirmed, 0);
    const totalDelivered = cplTrackerRows.reduce((s, r) => s + r.delivered, 0);
    return {
      totalSpendUsd,
      totalSpendTsh,
      totalLeads,
      totalOrders,
      totalConfirmed,
      totalDelivered,
      cpl: totalLeads > 0 ? totalSpendUsd / totalLeads : 0,
      cplConfirmed: totalConfirmed > 0 ? totalSpendUsd / totalConfirmed : 0,
      cplDelivered: totalDelivered > 0 ? totalSpendUsd / totalDelivered : 0,
      confirmRate: totalOrders > 0 ? totalConfirmed / totalOrders : 0,
      deliveryRate: totalConfirmed > 0 ? totalDelivered / totalConfirmed : 0,
    };
  }, [cplTrackerRows]);

  const normalizedTrackingRows = useMemo(() => {
    const today = getTodayString();
    return tracking.map((row) => {
      const fallbackDate =
        row.dateStart ||
        row.metaSince ||
        (row.metaImportedAt ? String(row.metaImportedAt).slice(0, 10) : today);
      const dateStart = row.dateStart || row.metaSince || fallbackDate;
      const dateEnd = row.dateEnd || row.metaUntil || dateStart;
      return {
        ...row,
        dateStart,
        dateEnd,
      };
    });
  }, [tracking]);

  const resolvedDashboardDateRange = useMemo(
    () => resolveDateRangeFilter(dashboardDateFilter),
    [dashboardDateFilter]
  );
  const resolvedProductPerformanceDateRange = useMemo(
    () => resolveDateRangeFilter(productPerformanceDateFilter),
    [productPerformanceDateFilter]
  );

  const dashboardFilteredServiceLeadCustomers = useMemo(
    () =>
      serviceLeadCustomers.filter((customer) =>
        isDateWithinRange(customer.orderDate, resolvedDashboardDateRange.start, resolvedDashboardDateRange.end)
      ),
    [resolvedDashboardDateRange.end, resolvedDashboardDateRange.start, serviceLeadCustomers]
  );
  const productPerformanceFilteredOperationalCustomers = useMemo(
    () =>
      operationalCustomers.filter((customer) =>
        isDateWithinRange(customer.orderDate, resolvedProductPerformanceDateRange.start, resolvedProductPerformanceDateRange.end)
      ),
    [operationalCustomers, resolvedProductPerformanceDateRange.end, resolvedProductPerformanceDateRange.start]
  );
  const dashboardFilteredTrackingRows = useMemo(
    () =>
      normalizedTrackingRows.filter((row) =>
        doesDateRangeOverlap(row.dateStart, row.dateEnd, resolvedDashboardDateRange.start, resolvedDashboardDateRange.end)
      ),
    [normalizedTrackingRows, resolvedDashboardDateRange.end, resolvedDashboardDateRange.start]
  );
  const productPerformanceFilteredTrackingRows = useMemo(
    () =>
      normalizedTrackingRows.filter((row) =>
        doesDateRangeOverlap(
          row.dateStart,
          row.dateEnd,
          resolvedProductPerformanceDateRange.start,
          resolvedProductPerformanceDateRange.end
        )
      ),
    [normalizedTrackingRows, resolvedProductPerformanceDateRange.end, resolvedProductPerformanceDateRange.start]
  );

  const dashboardFilteredCustomerMetricsByProduct = useMemo(
    () => buildCustomerMetricsByProduct(dashboardFilteredServiceLeadCustomers),
    [buildCustomerMetricsByProduct, dashboardFilteredServiceLeadCustomers]
  );
  const productPerformanceFilteredCustomerMetricsByProduct = useMemo(
    () => buildCustomerMetricsByProduct(productPerformanceFilteredOperationalCustomers),
    [buildCustomerMetricsByProduct, productPerformanceFilteredOperationalCustomers]
  );

  const dashboardFilteredProductDashboard = useMemo(
    () => buildProductDashboardRows(dashboardFilteredCustomerMetricsByProduct, dashboardFilteredTrackingRows),
    [buildProductDashboardRows, dashboardFilteredCustomerMetricsByProduct, dashboardFilteredTrackingRows]
  );
  const filteredProductPerformanceProductDashboard = useMemo(
    () =>
      buildProductDashboardRows(
        productPerformanceFilteredCustomerMetricsByProduct,
        productPerformanceFilteredTrackingRows
      ),
    [
      buildProductDashboardRows,
      productPerformanceFilteredCustomerMetricsByProduct,
      productPerformanceFilteredTrackingRows,
    ]
  );

  const productsCatalogSummary = useMemo(() => {
    const totalProducts = products.length;
    const totalUnits = products.reduce((sum, product) => sum + Number(product.totalQty || 0), 0);
    const totalImportBudgetTzs = products.reduce(
      (sum, product) =>
        sum +
        (Number(product.purchaseUnitPrice || 0) * Number(product.totalQty || 0) * USD_TO_TZS) +
        Number(product.shippingTotal || 0) +
        Number(product.otherCharges || 0),
      0
    );

    return {
      totalProducts,
      totalUnits,
      totalImportBudgetTzs,
      topScore: bestProduct?.score ?? 0,
    };
  }, [products, bestProduct]);

  const productPerformanceRows = useMemo(() => {
    return productDashboard.map((product) => {
      const manualInput = situationData.adInputs?.[product.id] || {};
      const rawFunnel = calculateCodFunnelMetrics({
        adsSpend: manualInput.manualAdsSpendTzs,
        leads: manualInput.leads,
        confirmedOrders: manualInput.confirmedOrders,
        deliveredOrders: manualInput.deliveredOrders,
      });
      const effectiveFunnel = calculateCodFunnelMetrics({
        adsSpend: rawFunnel.adsSpend > 0 ? rawFunnel.adsSpend : Number(product.totalAdsSpend || 0),
        leads: rawFunnel.leads > 0 ? rawFunnel.leads : Number(product.orders || 0),
        confirmedOrders:
          rawFunnel.confirmedOrders > 0 ? rawFunnel.confirmedOrders : Number(product.confirmed || 0),
        deliveredOrders:
          rawFunnel.deliveredOrders > 0
            ? rawFunnel.deliveredOrders
            : Math.max(0, Number(product.delivered || product.deliveredUnits || 0)),
      });
      const effectiveAdsSpendTzs = effectiveFunnel.adsSpend;
      const performance = calculateProductPerformance(product, {
        unitsSold: Number(product.totalUnitsSold || 0),
        totalRevenue: Number(product.totalRevenue || 0),
        totalAdsSpend: effectiveAdsSpendTzs,
      });
      const alertsData = calculateProductAlerts(
        {
          profit: performance.profit,
          stockQuantity: Number(product.stockQuantity || 0),
          adsSpend: effectiveAdsSpendTzs,
          leads: effectiveFunnel.leads,
          confirmedOrders: effectiveFunnel.confirmedOrders,
          deliveredOrders: effectiveFunnel.deliveredOrders,
          deliveryRate: effectiveFunnel.deliveryRate,
        },
        situationData.productAlertThresholds
      );
      const performanceStatus = getProductPerformanceStatus(
        performance.profit,
        situationData.productWinnerThresholdTzs
      );

      return {
        ...product,
        manualAdsSpendTzs: rawFunnel.adsSpend,
        funnelLeads: rawFunnel.leads,
        funnelConfirmedOrders: rawFunnel.confirmedOrders,
        funnelDeliveredOrders: rawFunnel.deliveredOrders,
        effectiveLeads: effectiveFunnel.leads,
        effectiveConfirmedOrders: effectiveFunnel.confirmedOrders,
        effectiveDeliveredOrders: effectiveFunnel.deliveredOrders,
        dashboardAdsSpendTzs: effectiveAdsSpendTzs,
        dashboardTotalProductCostTzs: performance.totalProductCost,
        dashboardTotalDeliveryCostTzs: performance.totalDeliveryCost,
        dashboardProfitTzs: performance.profit,
        dashboardProfitMargin: performance.profitMargin,
        dashboardCpaTzs: effectiveFunnel.cpa,
        dashboardConfirmationRate: effectiveFunnel.confirmationRate,
        dashboardDeliveryRate: effectiveFunnel.deliveryRate,
        performanceStatus,
        productAlerts: alertsData.alerts,
        productAlertCount: alertsData.alertCount,
      };
    });
  }, [productDashboard, situationData.adInputs, situationData.productAlertThresholds, situationData.productWinnerThresholdTzs]);

  const filteredProductPerformanceRows = useMemo(() => {
    return filteredProductPerformanceProductDashboard.map((product) => {
      const filteredFunnel = calculateCodFunnelMetrics({
        adsSpend: Number(product.totalAdsSpend || 0),
        leads: Number(product.orders || 0),
        confirmedOrders: Number(product.confirmed || 0),
        deliveredOrders: Number(product.delivered || 0),
      });
      const performance = calculateProductPerformance(product, {
        unitsSold: Number(product.totalUnitsSold || 0),
        totalRevenue: Number(product.totalRevenue || 0),
        totalAdsSpend: filteredFunnel.adsSpend,
      });
      const alertsData = calculateProductAlerts(
        {
          profit: performance.profit,
          stockQuantity: Number(product.stockQuantity || 0),
          adsSpend: filteredFunnel.adsSpend,
          leads: filteredFunnel.leads,
          confirmedOrders: filteredFunnel.confirmedOrders,
          deliveredOrders: filteredFunnel.deliveredOrders,
          deliveryRate: filteredFunnel.deliveryRate,
        },
        situationData.productAlertThresholds
      );
      const performanceStatus = getProductPerformanceStatus(
        performance.profit,
        situationData.productWinnerThresholdTzs
      );

      return {
        ...product,
        manualAdsSpendTzs: 0,
        funnelLeads: filteredFunnel.leads,
        funnelConfirmedOrders: filteredFunnel.confirmedOrders,
        funnelDeliveredOrders: filteredFunnel.deliveredOrders,
        effectiveLeads: filteredFunnel.leads,
        effectiveConfirmedOrders: filteredFunnel.confirmedOrders,
        effectiveDeliveredOrders: filteredFunnel.deliveredOrders,
        dashboardAdsSpendTzs: filteredFunnel.adsSpend,
        dashboardTotalProductCostTzs: performance.totalProductCost,
        dashboardTotalDeliveryCostTzs: performance.totalDeliveryCost,
        dashboardProfitTzs: performance.profit,
        dashboardProfitMargin: performance.profitMargin,
        dashboardCpaTzs: filteredFunnel.cpa,
        dashboardConfirmationRate: filteredFunnel.confirmationRate,
        dashboardDeliveryRate: filteredFunnel.deliveryRate,
        performanceStatus,
        productAlerts: alertsData.alerts,
        productAlertCount: alertsData.alertCount,
      };
    });
  }, [
    filteredProductPerformanceProductDashboard,
    situationData.productAlertThresholds,
    situationData.productWinnerThresholdTzs,
  ]);

  const dashboardDateSummary = useMemo(() => {
    const deliveredLeadCustomers = dashboardFilteredServiceLeadCustomers.filter(
      (customer) =>
        isConfirmationConfirmed(getCustomerConfirmationStatus(customer)) &&
        isShippingDelivered(getCustomerShippingStatus(customer))
    );
    const totalRevenue = deliveredLeadCustomers.reduce((sum, customer) => {
      const product = products.find((item) => item.id === customer.productId);
      return sum + getOrderRevenueAmounts(customer, product, USD_TO_TZS).revenueTsh;
    }, 0);
    const totalAdsSpend = calculateTotalAdsSpend(dashboardFilteredTrackingRows, products, USD_TO_TZS);
    const totalProductCost = dashboardFilteredProductDashboard.reduce(
      (sum, row) => sum + Number(row.totalProductCostTzs || 0),
      0
    );
    const totalDeliveryCost = dashboardFilteredProductDashboard.reduce(
      (sum, row) => sum + Number(row.totalDeliveryCostTzs || 0),
      0
    );
    const totalProfit = totalRevenue - totalAdsSpend - totalProductCost - totalDeliveryCost;
    const totalLeads = dashboardFilteredServiceLeadCustomers.length;
    const totalConfirmedOrders = dashboardFilteredServiceLeadCustomers.filter((customer) =>
      isConfirmationConfirmed(getCustomerConfirmationStatus(customer))
    ).length;
    const totalDeliveredOrders = deliveredLeadCustomers.length;
    const globalConfirmationRate = totalLeads > 0 ? (totalConfirmedOrders / totalLeads) * 100 : 0;
    const globalDeliveryRate =
      totalConfirmedOrders > 0 ? (totalDeliveredOrders / totalConfirmedOrders) * 100 : 0;
    const averageProfitMargin =
      totalRevenue > 0 ? (totalProfit / totalRevenue) * 100 : 0;

    return {
      totalRevenue,
      totalAdsSpend,
      totalProductCost,
      totalDeliveryCost,
      totalProfit,
      averageProfitMargin,
      totalLeads,
      totalConfirmedOrders,
      totalDeliveredOrders,
      globalConfirmationRate,
      globalDeliveryRate,
    };
  }, [dashboardFilteredProductDashboard, dashboardFilteredServiceLeadCustomers, dashboardFilteredTrackingRows, products]);

  const productPerformanceDateSummary = useMemo(() => {
    const totals = filteredProductPerformanceRows.reduce(
      (acc, row) => {
        acc.totalRevenue += Number(row.totalRevenue || 0);
        acc.totalAdsSpend += Number(row.dashboardAdsSpendTzs || 0);
        acc.totalProductCost += Number(row.dashboardTotalProductCostTzs || 0);
        acc.totalDeliveryCost += Number(row.dashboardTotalDeliveryCostTzs || 0);
        acc.totalProfit += Number(row.dashboardProfitTzs || 0);
        acc.totalLeads += Number(row.effectiveLeads || 0);
        acc.totalConfirmedOrders += Number(row.effectiveConfirmedOrders || 0);
        acc.totalDeliveredOrders += Number(row.effectiveDeliveredOrders || 0);
        return acc;
      },
      {
        totalRevenue: 0,
        totalAdsSpend: 0,
        totalProductCost: 0,
        totalDeliveryCost: 0,
        totalProfit: 0,
        totalLeads: 0,
        totalConfirmedOrders: 0,
        totalDeliveredOrders: 0,
      }
    );

    return {
      ...totals,
      averageProfitMargin: totals.totalRevenue > 0 ? (totals.totalProfit / totals.totalRevenue) * 100 : 0,
      globalConfirmationRate: totals.totalLeads > 0 ? (totals.totalConfirmedOrders / totals.totalLeads) * 100 : 0,
      globalDeliveryRate:
        totals.totalConfirmedOrders > 0
          ? (totals.totalDeliveredOrders / totals.totalConfirmedOrders) * 100
          : 0,
    };
  }, [filteredProductPerformanceRows]);

  const productAlertsSummary = useMemo(() => {
    const rows = productPerformanceRows.filter((row) => row.productAlerts.length > 0);
    return {
      totalProducts: rows.length,
      totalAlerts: rows.reduce((sum, row) => sum + row.productAlertCount, 0),
      topRows: rows.slice(0, 4),
    };
  }, [productPerformanceRows]);

  const controlPanelSummary = useMemo(() => {
    const totals = productPerformanceRows.reduce(
      (acc, row) => {
        acc.totalRevenueTzs += Number(row.totalRevenue || 0);
        acc.totalAdsSpendTzs += Number(row.dashboardAdsSpendTzs || 0);
        acc.totalProductCostTzs += Number(row.dashboardTotalProductCostTzs || 0);
        acc.totalDeliveryCostTzs += Number(row.dashboardTotalDeliveryCostTzs || 0);
        acc.totalProfitTzs += Number(row.dashboardProfitTzs || 0);
        acc.totalMarginPct += Number(row.dashboardProfitMargin || 0);
        if (Number(row.totalRevenue || 0) > 0) acc.marginRows += 1;
        return acc;
      },
      {
        totalRevenueTzs: 0,
        totalAdsSpendTzs: 0,
        totalProductCostTzs: 0,
        totalDeliveryCostTzs: 0,
        totalProfitTzs: 0,
        totalMarginPct: 0,
        marginRows: 0,
      }
    );

    const topWinningProducts = [...productPerformanceRows]
      .filter((row) => Number(row.dashboardProfitTzs || 0) > 0)
      .sort((a, b) => Number(b.dashboardProfitTzs || 0) - Number(a.dashboardProfitTzs || 0))
      .slice(0, 5);

    const losingProducts = [...productPerformanceRows]
      .filter((row) => Number(row.dashboardProfitTzs || 0) < 0)
      .sort((a, b) => Number(a.dashboardProfitTzs || 0) - Number(b.dashboardProfitTzs || 0))
      .slice(0, 5);

    const lowStockProducts = [...productPerformanceRows]
      .filter(
        (row) =>
          Number(row.stockQuantity || 0) <
          Number(situationData.productAlertThresholds?.minStockQuantity || 0)
      )
      .sort((a, b) => Number(a.stockQuantity || 0) - Number(b.stockQuantity || 0))
      .slice(0, 5);

    const needsAttentionProducts = [...productPerformanceRows]
      .filter((row) => row.productAlerts.length > 0)
      .sort((a, b) => Number(b.productAlertCount || 0) - Number(a.productAlertCount || 0))
      .slice(0, 6);

    return {
      totalRevenueTzs: totals.totalRevenueTzs,
      totalAdsSpendTzs: totals.totalAdsSpendTzs,
      totalProductCostTzs: totals.totalProductCostTzs,
      totalDeliveryCostTzs: totals.totalDeliveryCostTzs,
      totalProfitTzs: totals.totalProfitTzs,
      averageProfitMarginPct:
        totals.marginRows > 0 ? totals.totalMarginPct / totals.marginRows : 0,
      totalLeads: dashboardDateSummary.totalLeads,
      totalConfirmedOrders: dashboardDateSummary.totalConfirmedOrders,
      totalDeliveredOrders: dashboardDateSummary.totalDeliveredOrders,
      globalConfirmationRate: dashboardDateSummary.globalConfirmationRate,
      globalDeliveryRate: dashboardDateSummary.globalDeliveryRate,
      topWinningProducts,
      losingProducts,
      lowStockProducts,
      needsAttentionProducts,
    };
  }, [dashboardDateSummary, productPerformanceRows, situationData.productAlertThresholds?.minStockQuantity]);

  const downloadBlobFile = useCallback((content, fileName, mimeType) => {
    const blob = content instanceof Blob ? content : new Blob([content], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = fileName;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    URL.revokeObjectURL(url);
  }, []);

  const serializeCsvValue = useCallback((value) => {
    if (value === null || value === undefined) return "";
    const stringValue =
      typeof value === "string"
        ? value
        : typeof value === "number" || typeof value === "boolean"
          ? String(value)
          : JSON.stringify(value);
    if (/[",\n]/.test(stringValue)) {
      return `"${stringValue.replace(/"/g, "\"\"")}"`;
    }
    return stringValue;
  }, []);

  const buildCsvContent = useCallback(
    (headers, rows) => {
      const csvLines = [
        headers.map((header) => serializeCsvValue(header)).join(","),
        ...rows.map((row) =>
          headers.map((header) => serializeCsvValue(row?.[header])).join(",")
        ),
      ];
      return csvLines.join("\n");
    },
    [serializeCsvValue]
  );

  const exportAllDataToCsv = useCallback(() => {
    const headers = [
      "dataset",
      "id",
      "record_key",
      "name",
      "product_id",
      "product_name",
      "code",
      "status",
      "order_date",
      "quantity",
      "revenue_tzs",
      "ad_spend_tzs",
      "cost_per_unit_usd",
      "delivery_cost_unit_usd",
      "phone",
      "source",
      "created_at",
      "updated_at",
      "payload_json",
    ];

    const productRows = products.map((product) => ({
      dataset: "product",
      id: product.id,
      record_key: product.code || product.id,
      name: product.name,
      product_id: product.id,
      product_name: product.name,
      code: product.code || "",
      status: product.stockArrivalStatus || "",
      quantity: Number(product.qty || 0),
      revenue_tzs: Number(product.sellPrice || 0),
      cost_per_unit_usd: Number(product.buyPrice || 0),
      delivery_cost_unit_usd: Number(product.deliveryCostUsd || 8.5),
      source: product.source || "",
      created_at: product.createdAt || "",
      updated_at: product.updatedAt || "",
      payload_json: JSON.stringify(product),
    }));

    const customerRows = customers.map((customer) => {
      const product = products.find((item) => item.id === customer.productId);
      return {
        dataset: "customer_order",
        id: customer.id,
        record_key: customer.sourceOrderId || customer.id,
        name: customer.customerName || "",
        product_id: customer.productId || "",
        product_name: product?.name || "",
        code: customer.sourceOrderId || "",
        status: `${customer.status || ""}|${customer.shippingStatus || ""}`,
        order_date: customer.orderDate || "",
        quantity: Number(customer.quantity || 0),
        revenue_tzs: Number(customer.orderTotalTzs || 0),
        phone: customer.phone || "",
        source: customer.importSource || customer.source || "",
        created_at: customer.createdAt || "",
        updated_at: customer.updatedAt || "",
        payload_json: JSON.stringify(customer),
      };
    });

    const trackingRowsCsv = tracking.map((row) => ({
      dataset: "tracking_row",
      id: row.id,
      record_key: row.id,
      name: row.name || "",
      product_id: row.productId || "",
      product_name: products.find((item) => item.id === row.productId)?.name || "",
      status: row.status || "",
      order_date: row.dateStart || row.startDate || row.date || "",
      quantity: Number(row.deliveredOrders || row.orders || 0),
      revenue_tzs: Number(row.revenue || 0),
      ad_spend_tzs: Number(row.adSpend || 0),
      source: row.source || "tracking",
      created_at: row.createdAt || "",
      updated_at: row.updatedAt || "",
      payload_json: JSON.stringify(row),
    }));

    const metadataRows = [
      { dataset: "service_form", id: "service_form", record_key: "service_form", payload_json: JSON.stringify(serviceForm) },
      { dataset: "situation_data", id: "situation_data", record_key: "situation_data", payload_json: JSON.stringify(situationData) },
      { dataset: "meta_ads_state", id: "meta_ads_state", record_key: "meta_ads_state", payload_json: JSON.stringify(metaAdsState) },
      { dataset: "import_meta", id: "import_meta", record_key: "import_meta", payload_json: JSON.stringify(importMeta) },
    ];

    const csvContent = buildCsvContent(headers, [
      ...productRows,
      ...customerRows,
      ...trackingRowsCsv,
      ...metadataRows,
    ]);

    downloadBlobFile(
      csvContent,
      `tanzania-ecom-all-data-${new Date().toISOString().slice(0, 10)}.csv`,
      "text/csv;charset=utf-8"
    );
    setSharedWorkspace((prev) => ({ ...prev, notice: "All app data exported to CSV" }));
  }, [buildCsvContent, customers, downloadBlobFile, importMeta, metaAdsState, products, serviceForm, situationData, tracking]);

  const exportProductPerformanceToCsv = useCallback(() => {
    const headers = [
      "product_name",
      "stock_quantity",
      "units_sold",
      "total_revenue_tzs",
      "total_ads_spend_tzs",
      "total_product_cost_tzs",
      "total_delivery_cost_tzs",
      "profit_tzs",
      "profit_margin_pct",
      "leads",
      "confirmed_orders",
      "delivered_orders",
      "cpa_tzs",
      "confirmation_rate_pct",
      "delivery_rate_pct",
      "status",
      "alerts",
    ];

    const csvContent = buildCsvContent(
      headers,
      productPerformanceRows.map((row) => ({
        product_name: row.name,
        stock_quantity: Number(row.stockQuantity || 0),
        units_sold: Number(row.totalUnitsSold || 0),
        total_revenue_tzs: Number(row.totalRevenue || 0),
        total_ads_spend_tzs: Number(row.dashboardAdsSpendTzs || 0),
        total_product_cost_tzs: Number(row.dashboardTotalProductCostTzs || 0),
        total_delivery_cost_tzs: Number(row.dashboardTotalDeliveryCostTzs || 0),
        profit_tzs: Number(row.dashboardProfitTzs || 0),
        profit_margin_pct: Number(row.dashboardProfitMargin || 0).toFixed(2),
        leads: Number(row.effectiveLeads || 0),
        confirmed_orders: Number(row.effectiveConfirmedOrders || 0),
        delivered_orders: Number(row.effectiveDeliveredOrders || 0),
        cpa_tzs: row.dashboardCpaTzs === null || row.dashboardCpaTzs === undefined ? "N/A" : Number(row.dashboardCpaTzs || 0),
        confirmation_rate_pct: Number(row.dashboardConfirmationRate || 0).toFixed(2),
        delivery_rate_pct: Number(row.dashboardDeliveryRate || 0).toFixed(2),
        status: row.performanceStatus,
        alerts: row.productAlerts.map((alert) => alert.message).join(" | "),
      }))
    );

    downloadBlobFile(
      csvContent,
      `tanzania-ecom-product-performance-${new Date().toISOString().slice(0, 10)}.csv`,
      "text/csv;charset=utf-8"
    );
    setSharedWorkspace((prev) => ({ ...prev, notice: "Product performance report exported to CSV" }));
  }, [buildCsvContent, downloadBlobFile, productPerformanceRows]);

  const backupAllAppDataToJson = useCallback(() => {
    const payload = {
      snapshotVersion: 1,
      exportedAt: new Date().toISOString(),
      app: "Tanzania Ecom Tracker",
      state: buildSharedStateSnapshot(),
    };

    downloadBlobFile(
      JSON.stringify(payload, null, 2),
      `tanzania-ecom-backup-${new Date().toISOString().slice(0, 10)}.json`,
      "application/json;charset=utf-8"
    );
    setSharedWorkspace((prev) => ({ ...prev, notice: "Backup JSON downloaded" }));
  }, [buildSharedStateSnapshot, downloadBlobFile]);

  const restoreAppDataFromJson = useCallback(
    async (event) => {
      const file = event.target.files?.[0];
      if (!file) return;

      try {
        const rawText = await file.text();
        const parsed = JSON.parse(rawText);
        const snapshot =
          parsed && typeof parsed === "object" && parsed.state && typeof parsed.state === "object"
            ? parsed.state
            : parsed;

        if (!snapshot || typeof snapshot !== "object") {
          throw new Error("Invalid JSON backup format.");
        }

        applySharedStateSnapshot(snapshot);
        setSharedWorkspace((prev) => ({
          ...prev,
          notice: `JSON backup restored from ${file.name}`,
        }));
        if (supabaseEnabled) {
          migrateWorkspaceToNormalizedTables(snapshot, supabaseWorkspaceId)
            .then((res) => {
              if (res?.success) {
                setSyncNotice(`Restore synced to cloud ✓ — ${res.counts.products} products, ${res.counts.orders} orders`);
              }
            })
            .catch(() => null);
        }
      } catch (error) {
        setSharedWorkspace((prev) => ({
          ...prev,
          notice: `Restore JSON failed${error instanceof Error && error.message ? `: ${error.message}` : ""}`,
        }));
      } finally {
        event.target.value = "";
      }
    },
    [applySharedStateSnapshot]
  );

  const trackingSummary = useMemo(() => {
    const spend = tracking.reduce((sum, row) => sum + Number(row.adSpend || 0), 0);

    return productDashboard.reduce(
      (acc, product) => {
        acc.orders += Number(product.orders || 0);
        acc.confirmed += Number(product.confirmed || 0);
        acc.delivered += Number(product.deliveredUnits || 0);
        acc.revenue += Number(product.revenue || 0);
        acc.profit += Number(product.profit || 0);
        return acc;
      },
      { rows: tracking.length, spend, orders: 0, confirmed: 0, delivered: 0, revenue: 0, profit: 0 }
    );
  }, [productDashboard, tracking]);

  const selectedCustomerProduct = useMemo(
    () => products.find((product) => product.id === customerForm.productId),
    [products, customerForm.productId]
  );

  const customerFormPricing = useMemo(
    () => getProductPricing(selectedCustomerProduct, customerForm.quantity),
    [selectedCustomerProduct, customerForm.quantity]
  );

  const _customerFormOrderValue = useMemo(
    () => Number(customerFormPricing.totalPrice || 0),
    [customerFormPricing]
  );

  const confirmationStatusCatalog = useMemo(() => {
    const seen = new Set(DEFAULT_CONFIRMATION_STATUSES);
    serviceLeadCustomers.forEach((customer) => {
      const key = normalizeOrderStatus(customer.confirmationStatus || customer.status);
      if (key) seen.add(key);
    });

    return Array.from(seen)
      .map((status) => {
        const count = serviceLeadCustomers.filter(
          (customer) => normalizeOrderStatus(customer.confirmationStatus || customer.status) === status
        ).length;
        return {
          value: status,
          label: formatStatusLabel(status),
          bucket: getConfirmationBucket(status),
          count,
          color: getStatusColor(status),
        };
      })
      .filter((status) => status.count > 0 || DEFAULT_CONFIRMATION_STATUSES.includes(status.value))
      .sort((a, b) => {
        const order = { confirmed: 0, cancelled: 1, new: 2, pending: 3 };
        const gap = (order[a.bucket] ?? 9) - (order[b.bucket] ?? 9);
        if (gap !== 0) return gap;
        if (b.count !== a.count) return b.count - a.count;
        return a.label.localeCompare(b.label);
      });
  }, [serviceLeadCustomers]);

  const confirmationStatusMap = useMemo(
    () => Object.fromEntries(confirmationStatusCatalog.map((status) => [status.value, status])),
    [confirmationStatusCatalog]
  );

  const shippingStatusCatalog = useMemo(() => {
    const seen = new Set(DEFAULT_POST_CONFIRMATION_STATUSES);
    operationalCustomers.forEach((customer) => {
      const key = normalizeOrderStatus(
        customer.shippingStatus || (isConfirmationConfirmed(getCustomerConfirmationStatus(customer)) ? "to-prepare" : "")
      );
      if (key) seen.add(key);
    });

    return Array.from(seen)
      .map((status) => {
        const count = operationalCustomers.filter((customer) => {
          const effectiveShippingStatus = normalizeOrderStatus(
            customer.shippingStatus || (isConfirmationConfirmed(getCustomerConfirmationStatus(customer)) ? "to-prepare" : "")
          );
          return effectiveShippingStatus === status;
        }).length;
        return {
          value: status,
          label: formatStatusLabel(status),
          bucket: getShippingBucket(status),
          count,
          color: getStatusColor(status),
        };
      })
      .filter((status) => status.count > 0 || DEFAULT_POST_CONFIRMATION_STATUSES.includes(status.value))
      .sort((a, b) => {
        const order = { to_prepare: 0, shipped: 1, delivered: 2, returned: 3 };
        const gap = (order[a.bucket] ?? 9) - (order[b.bucket] ?? 9);
        if (gap !== 0) return gap;
        if (b.count !== a.count) return b.count - a.count;
        return a.label.localeCompare(b.label);
      });
  }, [operationalCustomers]);

  const shippingStatusMap = useMemo(
    () => Object.fromEntries(shippingStatusCatalog.map((status) => [status.value, status])),
    [shippingStatusCatalog]
  );

  const teamRoster = useMemo(() => {
    const seeded = ["Call Center", "Shipping Team", "Stock Team"];
    const salaryNames = situationData.salaries
      .map((entry) => String(entry.name || "").trim())
      .filter(Boolean);
    return Array.from(new Set([...seeded, ...salaryNames])).sort((a, b) => a.localeCompare(b));
  }, [situationData.salaries]);

  const selectedServiceDataset = useMemo(() => {
    const config = serviceCountryData[selectedService]?.[selectedCountry];
    if (!config) return null;

    const totalLeads = Number(serviceForm.totalLeads || 0);
    const confirmationRate = Number(serviceForm.confirmationRate || 0) / 100;
    const deliveryRate = Number(serviceForm.deliveryRate || 0) / 100;
    const sellingPriceTzs = Number(serviceForm.sellingPriceTzs || 0);
    const productCostTzs = Number(serviceForm.productCostTzs || 0);
    const costPerLeadUsd = Number(serviceForm.cplUsd || 0);
    const adSpendUsd = totalLeads * costPerLeadUsd;

    const confirmed = Math.round(totalLeads * confirmationRate);
    const delivered = Math.round(confirmed * deliveryRate);
    const sellingPriceUsd = sellingPriceTzs / config.usdToTzs;
    const productCostUsd = productCostTzs / config.usdToTzs;
    const revenueUsd = delivered * sellingPriceUsd;
    const deliveryFeesUsd = delivered * config.deliveryFeeUsdPerDelivered;
    const productCostTotalUsd = delivered * productCostUsd;
    const serviceFeeUsd = revenueUsd * (config.serviceFeePercent / 100);
    const totalServiceChargeUsd = serviceFeeUsd + deliveryFeesUsd;
    const adCostPerDeliveredUsd = delivered > 0 ? adSpendUsd / delivered : 0;
    const totalProfitUsd = revenueUsd - productCostTotalUsd - totalServiceChargeUsd - adSpendUsd;
    const profitPerOrderUsd = delivered > 0 ? totalProfitUsd / delivered : 0;
    const profitPerPieceUsd = delivered > 0 ? totalProfitUsd / delivered : 0;
    const totalProfitTzs = totalProfitUsd * config.usdToTzs;
    const profitPerPieceTzs = profitPerPieceUsd * config.usdToTzs;
    const revenueTzs = revenueUsd * config.usdToTzs;
    const grossMarginPerDeliveredUsd = sellingPriceUsd - productCostUsd - config.deliveryFeeUsdPerDelivered;
    const breakEvenCplUsd = confirmationRate > 0 && deliveryRate > 0 ? grossMarginPerDeliveredUsd * confirmationRate * deliveryRate : 0;
    const breakEvenPriceUsd = productCostUsd + config.deliveryFeeUsdPerDelivered + adCostPerDeliveredUsd;
    const marginPercent = revenueUsd > 0 ? (totalProfitUsd / revenueUsd) * 100 : 0;

    let decision = "TEST";
    if (totalProfitUsd > 0 && costPerLeadUsd <= breakEvenCplUsd) decision = "GOOD PRODUCT";
    if (totalProfitUsd < 0) decision = "BAD PRODUCT";

    const score = Math.max(0, Math.min(100, Math.round((marginPercent > 0 ? 40 : 0) + (costPerLeadUsd <= breakEvenCplUsd ? 30 : 0) + (deliveryRate >= 0.5 ? 15 : 0) + (confirmationRate >= 0.5 ? 15 : 0))));

    const serviceFeePerOrderUsd = sellingPriceUsd * (config.serviceFeePercent / 100) + config.deliveryFeeUsdPerDelivered;
    const profitFor100LeadsUsd = (() => {
      const d100 = 100 * confirmationRate * deliveryRate;
      const rev100 = d100 * sellingPriceUsd;
      const cost100 = d100 * productCostUsd;
      const svc100 = rev100 * (config.serviceFeePercent / 100) + d100 * config.deliveryFeeUsdPerDelivered;
      const ads100 = 100 * costPerLeadUsd;
      return rev100 - cost100 - svc100 - ads100;
    })();

    return {
      confirmed,
      delivered,
      sellingPriceUsd,
      productCostUsd,
      revenueUsd,
      revenueTzs,
      deliveryFeesUsd,
      serviceFeeUsd,
      totalServiceChargeUsd,
      productCostTotalUsd,
      costPerLeadUsd,
      adCostPerDeliveredUsd,
      profitPerOrderUsd,
      profitPerPieceUsd,
      profitPerPieceTzs,
      totalProfitUsd,
      totalProfitTzs,
      breakEvenCplUsd,
      breakEvenPriceUsd,
      marginPercent,
      decision,
      score,
      serviceFeePerOrderUsd,
      profitFor100LeadsUsd,
    };
  }, [selectedService, selectedCountry, serviceForm]);

  const pendingDubaiNotifications = useMemo(() => {
    const today = getTodayString();
    return products.filter(
      (p) =>
        (p.source || "") === "dubai" &&
        p.stockArrivalStatus !== "arrived" &&
        p.nextArrivalCheckDate &&
        p.nextArrivalCheckDate <= today
    );
  }, [products]);

  const reorderNotifications = useMemo(() => {
    return productDashboard.filter((product) => product.reorderStatus === "ORDER NOW" || product.reorderStatus === "SOON");
  }, [productDashboard]);

  const shippingImportReminder = useMemo(() => {
    const now = new Date(currentTime);
    const todayLabel = formatDateInput(now);
    const cutoffReached = now.getHours() >= 18;
    const lastShippingImportAt = importMeta?.lastShippingImportAt || null;
    const lastShippingImportDay = lastShippingImportAt ? formatDateInput(new Date(lastShippingImportAt)) : null;
    const confirmedPipelineCount = operationalCustomers.filter((customer) =>
      isConfirmationConfirmed(getCustomerConfirmationStatus(customer))
    ).length;

    return {
      isVisible: confirmedPipelineCount > 0 && cutoffReached && lastShippingImportDay !== todayLabel,
      confirmedPipelineCount,
      lastShippingImportAt,
      lastShippingImportLabel: lastShippingImportAt ? new Date(lastShippingImportAt).toLocaleString() : "No shipping import yet",
    };
  }, [currentTime, importMeta, operationalCustomers]);

  const selectedMetaAccount = useMemo(
    () => metaAdsAccounts.find((account) => String(account.id) === String(metaAdsState.accountId)),
    [metaAdsAccounts, metaAdsState.accountId]
  );

  const metaCampaignRows = useMemo(() => {
    return buildMappedMetaRows(metaAdsInsights?.rows || [], products, metaAdsState.campaignMappings);
  }, [metaAdsInsights, metaAdsState.campaignMappings, products]);

  const unmappedMetaCampaignRows = useMemo(
    () => metaCampaignRows.filter((row) => !row.mappedProductId),
    [metaCampaignRows]
  );

  const mappedCampaignsByProduct = useMemo(() => {
    const map = {};
    for (const row of metaCampaignRows) {
      if (!row.mappedProductId) continue;
      if (!map[row.mappedProductId]) map[row.mappedProductId] = [];
      map[row.mappedProductId].push(row);
    }
    return map;
  }, [metaCampaignRows]);

  const metaInsightsSummary = useMemo(() => {
    if (!metaAdsInsights?.summary) {
      return {
        spend: 0,
        impressions: 0,
        reach: 0,
        clicks: 0,
        inlineLinkClicks: 0,
        uniqueInlineLinkClicks: 0,
        landingPageViews: 0,
        leads: 0,
        actualLeads: 0,
        trackedLeads: 0,
        trackedLeadSource: "no_signal",
        ctr: 0,
        cpc: 0,
        costPerLead: 0,
      };
    }

    return metaAdsInsights.summary;
  }, [metaAdsInsights]);

  const metaCurrencyIsTzs = String(selectedMetaAccount?.currency || "").toUpperCase() === "TZS";
  const formatMetaMoney = useCallback(
    (value) => (metaCurrencyIsTzs ? formatTZS(value) : formatUSD(value)),
    [metaCurrencyIsTzs]
  );

  const metaDashboardMetrics = useMemo(() => {
    const spend = Number(metaInsightsSummary.spend || 0);
    const impressions = Number(metaInsightsSummary.impressions || 0);
    const reach = Number(metaInsightsSummary.reach || 0);
    const clicks = Number(metaInsightsSummary.clicks || 0);
    const inlineLinkClicks = Number(metaInsightsSummary.inlineLinkClicks || 0);
    const uniqueInlineLinkClicks = Number(metaInsightsSummary.uniqueInlineLinkClicks || 0);
    const landingPageViews = Number(metaInsightsSummary.landingPageViews || 0);
    const actualLeads = Number(metaInsightsSummary.actualLeads || metaInsightsSummary.leads || 0);
    const leads = Number(metaInsightsSummary.trackedLeads || metaInsightsSummary.leads || 0);
    const ctr = Number(metaInsightsSummary.ctr || 0);
    const cpc = Number(metaInsightsSummary.cpc || 0);
    const cpl = Number(metaInsightsSummary.costPerLead || 0);
    const cpm = Number(metaInsightsSummary.cpm || 0);
    const cpp = Number(metaInsightsSummary.cpp || 0);
    const frequency = Number(metaInsightsSummary.frequency || 0);
    return {
      spend,
      impressions,
      reach,
      clicks,
      inlineLinkClicks,
      uniqueInlineLinkClicks,
      landingPageViews,
      leads,
      actualLeads,
      trackedLeadSource: String(metaInsightsSummary.trackedLeadSource || "no_signal"),
      ctr,
      cpc,
      cpl,
      cpm,
      cpp,
      frequency,
      campaigns: metaCampaignRows.length,
    };
  }, [metaCampaignRows.length, metaInsightsSummary]);

  const isMetaTokenExpiredError = (error) => {
    const message = error instanceof Error ? error.message : String(error || "");
    return /session has expired|access token|oauth/i.test(message);
  };

  const handleMetaRequestError = useCallback((error, fallbackMessage = "Unable to sync Meta Ads.") => {
    const message = error instanceof Error ? error.message : String(error || fallbackMessage);
    if (isMetaTokenExpiredError(error)) {
      setMetaAdsState((prev) => ({ ...prev, autoSync: false }));
      const cleanMessage =
        "Meta access token expired. Auto-sync has been paused. Generate a new Meta access token, paste it here, then click Refresh insights to reconnect.";
      setMetaAdsNotice(cleanMessage);
      return cleanMessage;
    }

    const cleanMessage = message || fallbackMessage;
    setMetaAdsNotice(cleanMessage);
    return cleanMessage;
  }, []);

  const loadMetaAdAccounts = async () => {
    if (!metaAdsState.accessToken.trim()) {
      setMetaAdsNotice("Meta sync is optional. Paste your Meta access token only if you want to import Ads Manager data. The manual Tracking section below already works without it.");
      return;
    }

    setMetaAdsLoading((prev) => ({ ...prev, accounts: true }));
    setMetaAdsNotice("");

    try {
      const response = await fetch(`${getMetaApiBase()}/ad-accounts`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ accessToken: metaAdsState.accessToken.trim() }),
      });
      const payload = await response.json();
      if (!response.ok || !payload?.ok) throw new Error(payload?.error || "Unable to load Meta ad accounts.");

      setMetaAdsAccounts(payload.accounts || []);
      setMetaAdsNotice(payload.accounts?.length ? `Loaded ${payload.accounts.length} ad account(s).` : "No ad account found for this token.");

      if (!metaAdsState.accountId && payload.accounts?.[0]?.id) {
        setMetaAdsState((prev) => ({ ...prev, accountId: String(payload.accounts[0].id) }));
      }
    } catch (error) {
      handleMetaRequestError(error, "Unable to load Meta ad accounts.");
    } finally {
      setMetaAdsLoading((prev) => ({ ...prev, accounts: false }));
    }
  };

  const fetchMetaInsightsPayload = useCallback(async () => {
    if (!metaAdsState.accessToken.trim()) {
      throw new Error("Meta sync is optional. Paste your Meta access token only if you want to import Ads Manager data. The manual Tracking section below already works without it.");
    }
    if (!metaAdsState.accountId) {
      throw new Error("Choose an ad account before syncing insights.");
    }

    const response = await fetch(`${getMetaApiBase()}/insights`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        accessToken: metaAdsState.accessToken.trim(),
        accountId: metaAdsState.accountId,
        since: metaAdsState.dateStart,
        until: metaAdsState.dateEnd,
      }),
    });
    const payload = await response.json();
    if (!response.ok || !payload?.ok) throw new Error(payload?.error || "Unable to load Meta insights.");
    return payload;
  }, [metaAdsState.accessToken, metaAdsState.accountId, metaAdsState.dateEnd, metaAdsState.dateStart]);

  const fetchMetaSpendTotalPayload = useCallback(async () => {
    if (!metaAdsState.accessToken.trim()) {
      throw new Error("Meta access token is required.");
    }
    if (!metaAdsState.accountId) {
      throw new Error("Choose an ad account before syncing Meta total spend.");
    }

    const response = await fetch(`${getMetaApiBase()}/spend-total`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        accessToken: metaAdsState.accessToken.trim(),
        accountId: metaAdsState.accountId,
      }),
    });
    const payload = await response.json();
    if (!response.ok || !payload?.ok) throw new Error(payload?.error || "Unable to load Meta total spend.");
    return payload;
  }, [metaAdsState.accessToken, metaAdsState.accountId]);

  const fetchMetaDailySpendPayload = useCallback(async (date) => {
    if (!metaAdsState.accessToken.trim()) {
      throw new Error("Meta access token is required.");
    }
    if (!metaAdsState.accountId) {
      throw new Error("Choose an ad account before syncing Meta daily spend.");
    }

    const response = await fetch(`${getMetaApiBase()}/spend-daily`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        accessToken: metaAdsState.accessToken.trim(),
        accountId: metaAdsState.accountId,
        date,
      }),
    });
    const payload = await response.json();
    if (!response.ok || !payload?.ok) throw new Error(payload?.error || "Unable to load Meta daily spend.");
    return payload;
  }, [metaAdsState.accessToken, metaAdsState.accountId]);

  const importMetaInsightsPayload = useCallback(
    (payload, options = {}) => {
      const rawRows = deduplicateCampaigns(payload?.rows || []);
      const allMappedRows = buildMappedMetaRows(rawRows, products, metaAdsState.campaignMappings);

      const accountCurrency = String(selectedMetaAccount?.currency || "USD").toUpperCase();
      const convertSpendToTzs = (amount) => (accountCurrency === "TZS" ? amount : amount * USD_TO_TZS);

      const mappedRows = allMappedRows.filter((row) => row.mappedProductId && Number(row.spend || 0) > 0);
      const unmappedRows = allMappedRows.filter((row) => !row.mappedProductId && Number(row.spend || 0) > 0);

      const totalCampaigns = allMappedRows.length;
      const autoMappedCount = mappedRows.filter((r) => r.autoMapped).length;
      const manualMappedCount = mappedRows.filter((r) => r.manuallyMapped).length;
      const unmappedCount = unmappedRows.length;
      const unmappedSpendTzs = unmappedRows.reduce((sum, row) => sum + convertSpendToTzs(Number(row.spend || 0)), 0);
      const importedAt = new Date().toISOString();

      // Build campaign records for ALL rows (mapped + unmapped) — done BEFORE early return
      // so even fully-unmapped imports are persisted for CPL tracking.
      const newCampaignRecords = allMappedRows
        .filter((row) => Number(row.spend || 0) > 0)
        .map((row) => {
          const rawSpend = Number(row.spend || 0);
          // Prefer the actual Meta campaign_id (row.campaignId) over the synthetic row.id fallback
          const campaignId = String(row.campaignId || row.id || "");
          const matchedProduct = row.mappedProductId ? products.find((p) => p.id === row.mappedProductId) : null;
          return {
            campaignId,
            campaignName: String(row.campaignName || ""),
            dateStart: metaAdsState.dateStart,
            dateEnd: metaAdsState.dateEnd,
            // spendUsd: always convert to USD regardless of account currency
            spendUsd: accountCurrency === "TZS" ? rawSpend / USD_TO_TZS : rawSpend,
            spendTsh: Math.round(convertSpendToTzs(rawSpend)),
            productId: row.mappedProductId || "",
            productName: matchedProduct?.name || "",
            isMapped: Boolean(row.mappedProductId),
            isUnmapped: !row.mappedProductId,
            impressions: Number(row.impressions || 0),
            clicks: Number(row.clicks || 0),
            leads: Math.max(0, Number(row.trackedLeads ?? row.leads ?? 0)),
            savedAt: importedAt,
          };
        })
        .filter((r) => r.campaignId); // drop any rows with no usable campaign ID

      // Merge into savedCampaigns; dedup key = campaignId::dateStart::dateEnd
      // Same period re-imported → REPLACE (no double-count). New period → ACCUMULATE.
      const existingSaved = metaAdsState.savedCampaigns || [];
      const savedMap = new Map();
      for (const c of existingSaved) savedMap.set(`${c.campaignId}::${c.dateStart}::${c.dateEnd}`, c);
      for (const c of newCampaignRecords) savedMap.set(`${c.campaignId}::${c.dateStart}::${c.dateEnd}`, c);
      const mergedSavedCampaigns = Array.from(savedMap.values());

      // Persist to Supabase fire-and-forget; also update local adsCampaignsData state
      if (newCampaignRecords.length) {
        if (supabaseEnabled) {
          saveAdsCampaignsToSupabase(newCampaignRecords)
            .then(() => Promise.all([
              loadAdsSpendByProductFromSupabase().then(setSupabaseAdsSpendByProduct).catch(() => {}),
              loadProfitOverviewFromSupabase().then((r) => { if (r) setProfitOverviewDirect(r); }).catch(() => {}),
            ]))
            .catch(() => {});
        }
        setAdsCampaignsData((prev) => {
          const sbMap = new Map();
          for (const c of (prev.campaigns || [])) sbMap.set(`${c.campaignId}::${c.dateStart}::${c.dateEnd}`, c);
          for (const c of newCampaignRecords) sbMap.set(`${c.campaignId}::${c.dateStart}::${c.dateEnd}`, c);
          return { ...prev, campaigns: Array.from(sbMap.values()), lastLoaded: importedAt };
        });
      }

      // Early return AFTER saving — unmapped-only imports still persist campaign records
      if (!mappedRows.length) {
        setMetaAdsState((prev) => ({
          ...prev,
          lastSyncAt: importedAt,
          unmappedImportedSpendTzs: Math.round(unmappedSpendTzs),
          savedCampaigns: mergedSavedCampaigns,
        }));
        if (!options.silent) {
          setMetaAdsNotice(
            unmappedRows.length
              ? `No campaigns matched to products. ${unmappedRows.length} campaign(s) saved — assign them below.`
              : "No campaign rows found to import."
          );
        }
        return false;
      }

      const groupedByProduct = mappedRows.reduce((acc, row) => {
        const productId = row.mappedProductId;
        if (!acc[productId]) acc[productId] = { spendTzs: 0, leads: 0, actualLeads: 0 };
        acc[productId].spendTzs += convertSpendToTzs(Number(row.spend || 0));
        acc[productId].leads += Math.max(0, Number(row.trackedLeads ?? row.leads ?? 0));
        acc[productId].actualLeads += Math.max(0, Number(row.actualLeads ?? row.leads ?? 0));
        return acc;
      }, {});
      const totalMappedSpendTzs = Object.values(groupedByProduct).reduce((s, e) => s + e.spendTzs, 0);

      setTracking((prev) => {
        const next = [...prev];
        Object.entries(groupedByProduct).forEach(([productId, stats]) => {
          const existingIndex = next.findIndex((row) => row.productId === productId && row.metaManaged);
          const nextPayload = {
            productId,
            adSpend: Math.round(stats.spendTzs),
            orders: 0,
            confirmed: 0,
            delivered: 0,
            metaManaged: true,
            metaImportedAt: importedAt,
            metaSince: metaAdsState.dateStart,
            metaUntil: metaAdsState.dateEnd,
            dateStart: metaAdsState.dateStart,
            dateEnd: metaAdsState.dateEnd,
            metaCurrency: accountCurrency,
          };
          if (existingIndex >= 0) {
            next[existingIndex] = { ...next[existingIndex], ...nextPayload };
          } else {
            next.push({ id: buildNextId(next, "T"), ...nextPayload });
          }
        });
        return next;
      });

      setSituationData((prev) => ({
        ...prev,
        adInputs: {
          ...prev.adInputs,
          ...Object.fromEntries(
            Object.entries(groupedByProduct).map(([productId, stats]) => [
              productId,
              {
                averageLeadCostTzs: stats.leads > 0 ? stats.spendTzs / stats.leads : Number(prev.adInputs?.[productId]?.averageLeadCostTzs || 0),
                incomingLeads: Math.round(stats.leads),
              },
            ])
          ),
        },
      }));

      setMetaAdsState((prev) => ({
        ...prev,
        lastSyncAt: importedAt,
        unmappedImportedSpendTzs: Math.round(unmappedSpendTzs),
        savedCampaigns: mergedSavedCampaigns,
        lastSyncSummary: {
          since: metaAdsState.dateStart,
          until: metaAdsState.dateEnd,
          totalCampaigns,
          autoMappedCampaigns: autoMappedCount,
          manualMappedCampaigns: manualMappedCount,
          unmappedCampaigns: unmappedCount,
          matchedProducts: Object.keys(groupedByProduct).length,
          matchedRows: mappedRows.length,
          totalSpendTzs: Math.round(totalMappedSpendTzs),
          unmappedSpendTzs: Math.round(unmappedSpendTzs),
          totalLeads: Object.values(groupedByProduct).reduce((s, e) => s + e.leads, 0),
          totalActualLeads: Object.values(groupedByProduct).reduce((s, e) => s + e.actualLeads, 0),
        },
      }));

      if (!options.silent) {
        const reportParts = [
          `${totalCampaigns} campaigns imported — ${autoMappedCount} auto-mapped, ${manualMappedCount} manual, ${unmappedCount} unmapped.`,
          `Spend to products: ${formatTZS(Math.round(totalMappedSpendTzs))} across ${Object.keys(groupedByProduct).length} product(s).`,
          unmappedSpendTzs > 0
            ? `Unmapped spend: ${formatTZS(Math.round(unmappedSpendTzs))} (${unmappedCount} campaign(s) not assigned).`
            : "All campaigns with spend mapped to products.",
        ];
        setMetaAdsNotice(reportParts.join(" | "));
      }

      return true;
    },
    [metaAdsState.campaignMappings, metaAdsState.dateEnd, metaAdsState.dateStart, metaAdsState.savedCampaigns, products, selectedMetaAccount?.currency]
  );

  const syncMetaTotalSpend = useCallback(async (options = {}) => {
    const todayBucket = getDayBucket(currentTime);
    const hasBaseline = Number(metaAdsState.baselineTotalSpendTzs || 0) > 0 || Boolean(metaAdsState.baselineSpendBucket);
    const hasTodaySnapshot = Array.isArray(metaAdsState.dailySpendSnapshots)
      ? metaAdsState.dailySpendSnapshots.some((entry) => entry.bucket === todayBucket)
      : false;
    const shouldFetchDailySpend =
      hasBaseline &&
      Boolean(metaAdsState.baselineSpendBucket) &&
      todayBucket > String(metaAdsState.baselineSpendBucket) &&
      !hasTodaySnapshot;

    if (
      !options.force &&
      metaAdsState.lastLifetimeSpendSyncDate === todayBucket &&
      Number(metaAdsState.lifetimeSpendTzs || 0) > 0 &&
      (!shouldFetchDailySpend || hasTodaySnapshot)
    ) {
      return;
    }

    const totalPayload = await fetchMetaSpendTotalPayload();
    const totalAmount = Number(totalPayload?.spend || 0);
    const totalSpendTzs = metaCurrencyIsTzs ? totalAmount : totalAmount * USD_TO_TZS;
    const capturedAt = totalPayload?.capturedAt || new Date().toISOString();
    const buildMissingBuckets = (fromBucket, existingSnapshots) => {
      if (!fromBucket || todayBucket <= String(fromBucket)) return [];
      const snapshotBuckets = new Set((existingSnapshots || []).map((entry) => String(entry.bucket || "")));
      const buckets = [];
      let cursor = parseDateInput(String(fromBucket));
      const end = parseDateInput(todayBucket);
      if (!cursor || !end) return [];
      cursor.setDate(cursor.getDate() + 1);
      while (cursor <= end) {
        const bucket = formatDateInput(cursor);
        if (!snapshotBuckets.has(bucket)) {
          buckets.push(bucket);
        }
        cursor.setDate(cursor.getDate() + 1);
      }
      return buckets;
    };

    const missingBuckets = hasBaseline ? buildMissingBuckets(metaAdsState.baselineSpendBucket, metaAdsState.dailySpendSnapshots) : [];
    const dailyPayloads = [];
    for (const bucket of missingBuckets) {
      const payload = await fetchMetaDailySpendPayload(bucket);
      dailyPayloads.push(payload);
    }

    setMetaAdsState((prev) => {
      const prevHasBaseline = Number(prev.baselineTotalSpendTzs || 0) > 0 || Boolean(prev.baselineSpendBucket);
      const prevHasTodaySnapshot = Array.isArray(prev.dailySpendSnapshots)
        ? prev.dailySpendSnapshots.some((entry) => entry.bucket === todayBucket)
        : false;

      if (
        !options.force &&
        prev.lastLifetimeSpendSyncDate === todayBucket &&
        Number(prev.lifetimeSpendTzs || 0) > 0 &&
        (prevHasTodaySnapshot || !prevHasBaseline || todayBucket <= String(prev.baselineSpendBucket || ""))
      ) {
        return prev;
      }

      const baselineTotalSpendTzs = prevHasBaseline ? Number(prev.baselineTotalSpendTzs || 0) : totalSpendTzs;
      const baselineSpendBucket = prevHasBaseline ? prev.baselineSpendBucket || todayBucket : todayBucket;
      let dailySpendSnapshots = Array.isArray(prev.dailySpendSnapshots) ? [...prev.dailySpendSnapshots] : [];
      if (prevHasBaseline) {
        const prevMissingBuckets = buildMissingBuckets(baselineSpendBucket, dailySpendSnapshots);
        const payloadMap = new Map(
          dailyPayloads.map((payload) => {
            const bucket = String(payload?.date || "");
            const amount = Number(payload?.spend || 0);
            const spendTzs = metaCurrencyIsTzs ? amount : amount * USD_TO_TZS;
            return [
              bucket,
              {
                id: `meta-daily-${bucket}`,
                bucket,
                totalSpendTzs: null,
                newSpendTzs: Math.max(0, spendTzs),
                capturedAt: payload?.capturedAt || capturedAt,
                source: "meta_daily",
              },
            ];
          })
        );

        dailySpendSnapshots = [
          ...prevMissingBuckets
            .map((bucket) => payloadMap.get(bucket))
            .filter(Boolean),
          ...dailySpendSnapshots,
        ]
          .filter((entry, index, array) => array.findIndex((candidate) => candidate.bucket === entry.bucket) === index)
          .sort((a, b) => String(b.bucket || "").localeCompare(String(a.bucket || "")))
          .slice(0, 120);
      }

      const sortedAscendingSnapshots = [...dailySpendSnapshots].sort((a, b) => String(a.bucket || "").localeCompare(String(b.bucket || "")));
      let runningTotalTzs = baselineTotalSpendTzs;
      dailySpendSnapshots = sortedAscendingSnapshots.map((entry) => {
        runningTotalTzs += Number(entry.newSpendTzs || 0);
        return {
          ...entry,
          totalSpendTzs: runningTotalTzs,
        };
      }).sort((a, b) => String(b.bucket || "").localeCompare(String(a.bucket || "")));

      const cumulativeTrackedSpendTzs = totalSpendTzs;
      const lastDailySpendTzs = dailyPayloads.length
        ? Number(dailySpendSnapshots.find((entry) => entry.bucket === todayBucket)?.newSpendTzs || 0)
        : 0;

      return {
        ...prev,
        baselineTotalSpendTzs,
        baselineSpendBucket,
        lifetimeSpendTzs: totalSpendTzs,
        lastLifetimeSpendSyncDate: todayBucket,
        lifetimeSpendCapturedAt: capturedAt,
        cumulativeTrackedSpendTzs,
        dailySpendSnapshots,
        lastSyncSummary: prev.lastSyncSummary
          ? {
              ...prev.lastSyncSummary,
              accountTotalSpendTzs: totalSpendTzs,
              trackedCumulativeSpendTzs: cumulativeTrackedSpendTzs,
              lastDailySpendTzs,
            }
          : prev.lastSyncSummary,
      };
    });
  }, [
    currentTime,
    fetchMetaDailySpendPayload,
    fetchMetaSpendTotalPayload,
    metaAdsState.baselineSpendBucket,
    metaAdsState.baselineTotalSpendTzs,
    metaAdsState.dailySpendSnapshots,
    metaAdsState.lastLifetimeSpendSyncDate,
    metaAdsState.lifetimeSpendTzs,
    metaCurrencyIsTzs,
  ]);

  const refreshMetaInsights = useCallback(async (options = {}) => {
    setMetaAdsLoading((prev) => ({ ...prev, insights: !options.silent }));
    if (!options.silent) setMetaAdsNotice("");
    try {
      const payload = await fetchMetaInsightsPayload();
      setMetaAdsInsights(payload);
      const mappedRows = buildMappedMetaRows(payload?.rows || [], products, metaAdsState.campaignMappings);
      setMetaAdsState((prev) => ({
        ...prev,
        campaignMetadata: mappedRows.map((row) => row.campaignMetadata),
      }));
      if (options.applyToApp || metaAdsState.autoSync) {
        importMetaInsightsPayload(payload, { silent: true });
      }
      if (options.syncTotalSpend) {
        await syncMetaTotalSpend({ force: true });
      }
      if (!options.silent) setMetaAdsNotice(`Insights updated for ${metaAdsState.dateStart} -> ${metaAdsState.dateEnd}.`);
    } catch (error) {
      if (!options.silent || isMetaTokenExpiredError(error)) handleMetaRequestError(error, "Unable to load Meta insights.");
    } finally {
      setMetaAdsLoading((prev) => ({ ...prev, insights: false }));
    }
  }, [fetchMetaInsightsPayload, handleMetaRequestError, importMetaInsightsPayload, metaAdsState.autoSync, metaAdsState.campaignMappings, metaAdsState.dateEnd, metaAdsState.dateStart, products, syncMetaTotalSpend]);

  const applyMetaInsightsToApp = () => {
    setMetaAdsLoading((prev) => ({ ...prev, apply: true }));
    try {
      importMetaInsightsPayload(metaAdsInsights, { silent: false });
    } finally {
      setMetaAdsLoading((prev) => ({ ...prev, apply: false }));
    }
  };

  useEffect(() => {
    if (!metaAdsState.accessToken.trim() || !metaAdsState.accountId) return;
    if (metaAdsState.autoSync) return;
    if (metaAdsState.lastSyncAt || metaAdsState.baselineSpendBucket || Number(metaAdsState.cumulativeTrackedSpendTzs || 0) > 0) return;

    setMetaAdsState((prev) => {
      if (!prev.accessToken.trim() || !prev.accountId || prev.autoSync) return prev;
      if (prev.lastSyncAt || prev.baselineSpendBucket || Number(prev.cumulativeTrackedSpendTzs || 0) > 0) return prev;
      return {
        ...prev,
        autoSync: true,
      };
    });
  }, [
    metaAdsState.accessToken,
    metaAdsState.accountId,
    metaAdsState.autoSync,
    metaAdsState.baselineSpendBucket,
    metaAdsState.cumulativeTrackedSpendTzs,
    metaAdsState.lastSyncAt,
  ]);

  useEffect(() => {
    if (!metaAdsState.accessToken.trim() || !metaAdsState.accountId) return undefined;

    const bootstrapKey = `${metaAdsState.accountId}|${metaAdsState.accessToken.slice(0, 16)}|meta-baseline`;
    if (metaSpendBootstrapRef.current === bootstrapKey) return undefined;

    let cancelled = false;
    const bootstrapMetaBaseline = async () => {
      try {
        await syncMetaTotalSpend({ force: true });
        if (!cancelled) {
          metaSpendBootstrapRef.current = bootstrapKey;
        }
      } catch (error) {
        if (!cancelled && isMetaTokenExpiredError(error)) {
          handleMetaRequestError(error, "Unable to load Meta total spend.");
        }
      }
    };

    bootstrapMetaBaseline();
    return () => {
      cancelled = true;
    };
  }, [handleMetaRequestError, metaAdsState.accessToken, metaAdsState.accountId, syncMetaTotalSpend]);

  useEffect(() => {
    if (!metaAdsState.autoSync) return undefined;
    if (!metaAdsState.accessToken.trim() || !metaAdsState.accountId) return undefined;

    const bootstrapKey = `${metaAdsState.accountId}|${metaAdsState.accessToken.slice(0, 16)}`;
    const runSync = async () => {
      if (metaAutoSyncLockRef.current) return;
      metaAutoSyncLockRef.current = true;
      try {
        await refreshMetaInsights({ silent: true, applyToApp: true });
        const shouldBootstrap = metaSpendBootstrapRef.current !== bootstrapKey;
        await syncMetaTotalSpend({ force: shouldBootstrap });
        if (shouldBootstrap) {
          metaSpendBootstrapRef.current = bootstrapKey;
        }
      } catch (error) {
        handleMetaRequestError(error, "Unable to auto-sync Meta Ads.");
      } finally {
        metaAutoSyncLockRef.current = false;
      }
    };

    runSync();
    const interval = window.setInterval(runSync, Math.max(1, Number(metaAdsState.autoSyncIntervalMinutes || 3)) * 60000);
    return () => window.clearInterval(interval);
  }, [
    importMetaInsightsPayload,
    metaAdsState.accessToken,
    metaAdsState.accountId,
    metaAdsState.autoSync,
    metaAdsState.autoSyncIntervalMinutes,
    metaAdsState.campaignMappings,
    metaAdsState.dateEnd,
    metaAdsState.dateStart,
    handleMetaRequestError,
    refreshMetaInsights,
    syncMetaTotalSpend,
  ]);

  // Load campaign history, per-product spend totals, profit overview, and revenue import directly from Supabase on mount.
  useEffect(() => {
    if (!supabaseEnabled) return;
    loadAdsCampaignsFromSupabase()
      .then((result) => setAdsCampaignsData({ ...result, lastLoaded: new Date().toISOString() }))
      .catch(() => {});
    loadAdsSpendByProductFromSupabase()
      .then(setSupabaseAdsSpendByProduct)
      .catch(() => {});
    loadProfitOverviewFromSupabase()
      .then((result) => { if (result) setProfitOverviewDirect(result); })
      .catch(() => {});
    loadRevenueImportFromSupabase()
      .then((result) => {
        if (result) { setRevenueImport(result); writeLS(LS_REVENUE_IMPORT, result); }
        else { const local = readLS(LS_REVENUE_IMPORT); if (local) setRevenueImport(local); }
      })
      .catch(() => { const local = readLS(LS_REVENUE_IMPORT); if (local) setRevenueImport(local); });
    loadRevenueImportRowsFromSupabase()
      .then((rows) => {
        if (rows && Object.keys(rows).length > 0) { setRevenueImportRows(rows); writeLS(LS_REVENUE_ROWS, rows); }
        else { const local = readLS(LS_REVENUE_ROWS); if (local && Object.keys(local).length > 0) setRevenueImportRows(local); }
      })
      .catch(() => { const local = readLS(LS_REVENUE_ROWS); if (local && Object.keys(local).length > 0) setRevenueImportRows(local); });
    loadManualAdsSpendFromSupabase()
      .then((rows) => {
        if (rows.length > 0) { setManualAdsSpend(rows); writeLS(LS_MANUAL_ADS, rows); }
        else { const local = readLS(LS_MANUAL_ADS); if (Array.isArray(local) && local.length > 0) setManualAdsSpend(local); }
      })
      .catch(() => { const local = readLS(LS_MANUAL_ADS); if (Array.isArray(local) && local.length > 0) setManualAdsSpend(local); });
    loadExtraChargesFromSupabase()
      .then((rows) => {
        if (rows.length > 0) { setExtraCharges(rows); writeLS(LS_EXTRA_CHARGES, rows); }
        else { const local = readLS(LS_EXTRA_CHARGES); if (Array.isArray(local) && local.length > 0) setExtraCharges(local); }
      })
      .catch(() => { const local = readLS(LS_EXTRA_CHARGES); if (Array.isArray(local) && local.length > 0) setExtraCharges(local); });
    loadOwnerInjectionsFromSupabase()
      .then((rows) => {
        if (rows.length > 0) { setOwnerInjections(rows); writeLS(LS_OWNER_INJECTIONS, rows); }
        else { const local = readLS(LS_OWNER_INJECTIONS); if (Array.isArray(local) && local.length > 0) setOwnerInjections(local); }
      })
      .catch(() => { const local = readLS(LS_OWNER_INJECTIONS); if (Array.isArray(local) && local.length > 0) setOwnerInjections(local); });
  }, []);

  const submitCloudAuth = async () => {
    const email = cloudAuth.email.trim();
    const password = cloudAuth.password;
    if (!email || !password) {
      setCloudAuth((prev) => ({ ...prev, notice: "Enter email and password first." }));
      return;
    }

    setCloudAuth((prev) => ({ ...prev, loading: true, notice: prev.mode === "signup" ? "Creating cloud access..." : "Signing in..." }));
    try {
      if (cloudAuth.mode === "signup") {
        await signUpCloud({ email, password });
        setCloudAuth((prev) => ({
          ...prev,
          loading: false,
          password: "",
          notice: "Account created. If email confirmation is enabled, confirm it then sign in.",
        }));
      } else {
        await signInCloud({ email, password });
        setCloudAuth((prev) => ({
          ...prev,
          loading: false,
          password: "",
          notice: `Connected to cloud workspace as ${email}`,
        }));
      }
    } catch (error) {
      setCloudAuth((prev) => ({
        ...prev,
        loading: false,
        notice: error instanceof Error ? error.message : "Unable to access the cloud workspace.",
      }));
    }
  };

  const logoutCloudAuth = async () => {
    try {
      await signOutCloud();
      setCloudAuth((prev) => ({
        ...prev,
        password: "",
        notice: "Signed out from cloud workspace.",
      }));
    } catch (error) {
      setCloudAuth((prev) => ({
        ...prev,
        notice: error instanceof Error ? error.message : "Unable to sign out.",
      }));
    }
  };

  const createProductId = () => buildNextId(products, "P");
  const createCustomerId = () => buildNextId(customers, "C");

  const saveExpeditionProduct = async () => {
    if (!expeditionForm.name.trim()) return;

    const source = expeditionForm.source || "china";
    const normalizedProduct = {
      name: expeditionForm.name.trim(),
      mappingCode: expeditionForm.mappingCode
        ? expeditionForm.mappingCode.toUpperCase().replace(/[^A-Z0-9]/g, "")
        : generateMappingCode(expeditionForm.name.trim(), products, editingProductId || undefined),
      source,
      sellingPrice: Number(expeditionForm.sellingPrice || 0),
      purchaseUnitPrice: Number(expeditionForm.purchaseUnitPrice || 0),
      totalQty: Number(expeditionForm.totalQty || 0),
      shippingTotal: Number(expeditionForm.shippingTotal || 0),
      otherCharges: Number(expeditionForm.otherCharges || 0),
      delivery: Number(expeditionForm.delivery || 0),
      estimatedArrivalDays: Number(expeditionForm.estimatedArrivalDays || 0),
      supplierName: expeditionForm.supplierName.trim(),
      supplierContact: expeditionForm.supplierContact.trim(),
      lifecycleStatus: expeditionForm.lifecycleStatus || "test",
      defectRate: Math.max(0, Number(expeditionForm.defectRate || 0)),
      notes: expeditionForm.notes.trim(),
      offers: normalizeProductOffers(expeditionForm.offers),
    };

    let nextProducts;
    let savedProductRecord;

    if (editingProductId) {
      nextProducts = products.map((product) =>
          product.id === editingProductId
            ? {
                ...product,
                ...normalizedProduct,
                stockArrivalStatus: source === "dubai" ? product.stockArrivalStatus || "pending" : "arrived",
                stockOrderedAt: product.stockOrderedAt || getTodayString(),
                nextArrivalCheckDate:
                  source === "dubai"
                      ? product.nextArrivalCheckDate || addDaysToDateString(getTodayString(), Number(expeditionForm.estimatedArrivalDays || 0))
                      : null,
                  stockArrivedAt: source === "dubai" ? product.stockArrivedAt || null : product.stockArrivedAt || getTodayString(),
                }
              : product
        );
      savedProductRecord = nextProducts.find((product) => product.id === editingProductId) || null;
    } else {
      const newProduct = {
        id: createProductId(),
        ...normalizedProduct,
        stockArrivalStatus: source === "dubai" ? "pending" : "arrived",
        stockOrderedAt: getTodayString(),
        nextArrivalCheckDate:
          source === "dubai"
            ? addDaysToDateString(getTodayString(), Number(expeditionForm.estimatedArrivalDays || 0))
            : null,
        stockArrivedAt: source === "dubai" ? null : getTodayString(),
      };

      nextProducts = [...products, newProduct];
      savedProductRecord = newProduct;
    }

    const nextStockPurchases = savedProductRecord
      ? upsertManualSeedStockPurchase(stockPurchases, savedProductRecord)
      : stockPurchases;

    setProducts(nextProducts);
    setStockPurchases(nextStockPurchases);
    const nextSnapshot = {
      ...(latestSharedStateRef.current || getDefaultCloudWorkspaceState()),
      products: nextProducts.map(sanitizeProductRecord),
      stockPurchases: nextStockPurchases,
    };
    latestSharedStateRef.current = nextSnapshot;
    const saveOk = await persistSharedSnapshot(nextSnapshot, {
      progressNotice: "Saving product changes to cloud...",
      successNotice: editingProductId ? "Cloud product updated" : "Cloud product added",
      failurePrefix: "Cloud product sync failed",
    });
    if (!saveOk) {
      setSharedWorkspace((prev) => ({
        ...prev,
        notice: supabaseEnabled
          ? "Cloud sync failed — product saved locally. Sign in or check your connection and try again."
          : "Save failed — product kept in memory only.",
      }));
    }

    setEditingProductId(null);
    setExpeditionForm(getEmptyExpeditionForm());
    setShowAddProductForm(false);
    setActivePage("products");
  };

  const startEditingProduct = (product) => {
    setEditingProductId(product.id);
    setExpeditionForm({
      name: product.name || "",
      mappingCode: product.mappingCode || generateMappingCode(product.name || "", products, product.id),
      source: product.source || "china",
      sellingPrice: Number(product.sellingPrice || 0),
      purchaseUnitPrice: Number(product.purchaseUnitPrice || 0),
      totalQty: Number(product.totalQty || 0),
      shippingTotal: Number(product.shippingTotal || 0),
      otherCharges: Number(product.otherCharges || 0),
      delivery: Number(product.delivery || 0),
      estimatedArrivalDays: Number(product.estimatedArrivalDays || 0),
      supplierName: product.supplierName || "",
      supplierContact: product.supplierContact || "",
      lifecycleStatus: product.lifecycleStatus || "test",
      defectRate: Number(product.defectRate || 0),
      notes: product.notes || "",
      offers: normalizeProductOffers(product.offers),
    });
  };

  const cancelEditingProduct = () => {
    setEditingProductId(null);
    setShowAddProductForm(false);
    setExpeditionForm(getEmptyExpeditionForm());
  };

  const addProductOfferTier = () => {
    setExpeditionForm((prev) => ({
      ...prev,
      offers: [...(prev.offers || []), { minQty: Math.max(2, (prev.offers || []).length + 2), totalPrice: 0 }],
    }));
  };

  const updateProductOfferTier = (index, field, value) => {
    setExpeditionForm((prev) => ({
      ...prev,
      offers: (prev.offers || []).map((offer, offerIndex) =>
        offerIndex === index
          ? {
              ...offer,
              [field]: field === "minQty" ? Math.max(2, Number(value || 0)) : Math.max(0, parseLooseNumber(value)),
            }
          : offer
      ),
    }));
  };

  const removeProductOfferTier = (index) => {
    setExpeditionForm((prev) => ({
      ...prev,
      offers: (prev.offers || []).filter((_, offerIndex) => offerIndex !== index),
    }));
  };

  const deleteProduct = (productId) => {
    const nextProducts = products.filter((p) => p.id !== productId);
    const nextTracking = tracking.filter((t) => t.productId !== productId);
    const nextCustomers = customers.filter((c) => c.productId !== productId);
    const nextStockPurchases = stockPurchases.filter((purchase) => purchase.product_id !== productId);
    const nextStockMovements = stockMovements.filter((movement) => movement.product_id !== productId);

    setProducts(nextProducts);
    setTracking(nextTracking);
    setCustomers(nextCustomers);
    setStockPurchases(nextStockPurchases);
    setStockMovements(nextStockMovements);

    const nextSnapshot = {
      ...(latestSharedStateRef.current || getDefaultCloudWorkspaceState()),
      products: nextProducts.map(sanitizeProductRecord),
      tracking: nextTracking,
      customers: nextCustomers.map(sanitizeCustomerRecord),
      stockPurchases: nextStockPurchases,
      stockMovements: nextStockMovements,
    };
    latestSharedStateRef.current = nextSnapshot;
    void persistSharedSnapshot(nextSnapshot, {
      progressNotice: "Saving product deletion to cloud...",
      successNotice: "Cloud product deleted",
      failurePrefix: "Cloud product delete failed",
    });

    if (editingProductId === productId) {
      setEditingProductId(null);
      setExpeditionForm(getEmptyExpeditionForm());
    }
  };

  const handleClearAllProducts = async () => {
    const nextSnapshot = {
      ...(latestSharedStateRef.current || getDefaultCloudWorkspaceState()),
      products: [],
    };
    setProducts([]);
    setEditingProductId(null);
    setExpeditionForm(getEmptyExpeditionForm());
    setClearProductsConfirm(false);
    latestSharedStateRef.current = nextSnapshot;
    // Clear normalized Supabase table so the fallback loader doesn't restore deleted products
    try { await clearNormalizedProducts(supabaseWorkspaceId); } catch { /* ignore */ }
    // Clear localStorage backups so browser backup doesn't restore them either
    try {
      localStorage.removeItem(STORAGE_KEY);
      localStorage.removeItem(AUTO_BACKUP_KEY);
      localStorage.removeItem(AUTO_BACKUP_META_KEY);
    } catch { /* ignore */ }
    void persistSharedSnapshot(nextSnapshot, {
      progressNotice: "Clearing products...",
      successNotice: "All products deleted",
      failurePrefix: "Clear products failed",
    });
  };

  const resolveImportedProductId = useCallback(
    (rawValue) => matchProductIdFromText(rawValue, products),
    [products]
  );

  const findMatchingCustomerIndex = useCallback((customerList, payload) => {
    const importKey = String(payload.import_key || "").trim();
    if (importKey) {
      return customerList.findIndex((customer) => String(customer?.import_key || "").trim() === importKey);
    }

    const normalizedSourceOrderId = String(payload.sourceOrderId || "").trim();
    const normalizedPhone = normalizePhoneValue(payload.phone);
    const normalizedName = normalizeHeaderName(payload.customerName || payload.client_name);
    const normalizedOrderDate = String(payload.orderDate || "").trim();
    const normalizedQuantity = payload.quantity ? Number(payload.quantity) : null;
    const normalizedProductRef = normalizeHeaderName(payload.product_ref);
    const normalizedLineItemIndex = Math.max(0, Number(payload.line_item_index || 0));

    if (normalizedSourceOrderId) {
      return customerList.findIndex((customer) => {
        return (
          Boolean(customer?.sourceOrderId || customer?.order_id) &&
          String(customer.sourceOrderId || customer.order_id).trim() === normalizedSourceOrderId &&
          String(customer?.productId || "").trim() === String(payload.productId || "").trim() &&
          Math.max(0, Number(customer?.line_item_index || 0)) === normalizedLineItemIndex
        );
      });
    }

    return customerList.findIndex((customer) => {
      const customerPhone = normalizePhoneValue(customer.phone);
      const samePhone = normalizedPhone && customerPhone === normalizedPhone;
      const sameProduct = payload.productId ? customer.productId === payload.productId : true;
      const sameDate = normalizedOrderDate ? String(customer.orderDate || "") === normalizedOrderDate : true;
      const sameQuantity = normalizedQuantity ? Number(customer.quantity || 1) === normalizedQuantity : true;
      const sameName = normalizedName
        ? normalizeHeaderName(customer.customerName) === normalizedName
        : true;
      const sameRef = normalizedProductRef
        ? normalizeHeaderName(customer.product_ref) === normalizedProductRef
        : true;

      if (samePhone && sameProduct && sameDate && sameQuantity && sameRef) return true;
      if (samePhone && sameName && sameProduct && sameRef) return true;
      return false;
    });
  }, []);

  const importOrdersFromExcel = useCallback(
    async (event) => {
      const file = event.target.files?.[0];
      if (!file) return;

      try {
        const buffer = await file.arrayBuffer();
        const workbook = XLSX.read(buffer, { type: "array", cellDates: false });
        const firstSheetName = workbook.SheetNames[0];
        const firstSheet = workbook.Sheets[firstSheetName];
        const rows = XLSX.utils.sheet_to_json(firstSheet, { defval: "" });

        if (!rows.length) {
          setOrdersImportNotice("The Excel file is empty.");
          setOrdersImportDetails(null);
          return;
        }

        const nextCustomers = [...customers];
        const importFinishedAt = new Date().toISOString();
        const { parsedRows, report } = parseImportedExcelRows(rows, {
          exchangeRate: USD_TO_TZS,
          resolveProductId: resolveImportedProductId,
        });
        let createdCount = 0;
        let updatedCount = 0;
        let skippedCount = 0;
        const unmatchedProducts = new Set();
        const detectedHeaders = Object.keys(rows[0] || {});
        const unknownConfirmationStatuses = new Set();
        const unknownShippingStatuses = new Set();
        const orderMovements = [];

        parsedRows.forEach((row) => {
          const productId = row.productId || resolveImportedProductId(row.product_ref || row.product_name);
          if (!row.client_name) {
            skippedCount += 1;
            return;
          }
          if (!row.normalized_phone) {
            skippedCount += 1;
            return;
          }
          if (!String(row.product_name || "").trim()) {
            skippedCount += 1;
            return;
          }
          if (!productId) {
            unmatchedProducts.add(String(row.product_name || row.product_ref || "").trim());
          }

          const normalizedConfirmation = normalizeStatus(row.confirmation_status_raw);
          const normalizedShipping = normalizeStatus(row.shipping_status_raw);
          if (
            normalizedConfirmation &&
            !isConfirmedStatus(normalizedConfirmation) &&
            !isNoReplyStatus(normalizedConfirmation) &&
            !isCancelledStatus(normalizedConfirmation) &&
            !isInvalidLeadStatus(normalizedConfirmation) &&
            !isStockHoldStatus(normalizedConfirmation)
          ) {
            unknownConfirmationStatuses.add(normalizedConfirmation);
          }
          if (
            normalizedShipping &&
            !isDeliveredStatus(normalizedShipping) &&
            !isReturnedStatus(normalizedShipping) &&
            !isPendingShippingStatus(normalizedShipping) &&
            !isBlankShippingStatus(normalizedShipping)
          ) {
            unknownShippingStatuses.add(normalizedShipping);
          }

          const confirmationStatus = normalizedConfirmation ? normalizeOrderStatus(normalizedConfirmation) : "new-order";
          const shippingStatus = normalizedShipping ? normalizeOrderStatus(normalizedShipping) : "";
          const existingIndex = findMatchingCustomerIndex(nextCustomers, {
            import_key: row.import_key,
            sourceOrderId: row.order_id,
            customerName: row.client_name,
            phone: row.phone,
            productId,
            product_ref: row.product_ref,
            quantity: row.quantity,
            orderDate: excelDateToInput(row.created_at),
            line_item_index: row.line_item_index,
          });

          if (existingIndex >= 0) {
            const existing = nextCustomers[existingIndex];
            const previousConfirmation = normalizeOrderStatus(existing.confirmationStatus || existing.status);
            const previousShipping = normalizeOrderStatus(existing.shippingStatus);
            const nextShippingStatus = ensureShippingStatusForConfirmed(
              confirmationStatus || previousConfirmation,
              shippingStatus || existing.shippingStatus
            );
            const statusChanged = previousConfirmation !== confirmationStatus;
            const shippingChanged = previousShipping !== nextShippingStatus;
            const amountChanged = Math.max(0, Number(existing.amount_tsh || existing.orderTotalTzs || 0)) !== Math.max(0, Number(row.amount_tsh || 0));
            report.existingLeadsUpdated += 1;
            if (statusChanged || shippingChanged) report.statusChangesDetected += 1;
            const updatedOrder = sanitizeCustomerRecord({
              ...existing,
              customerName: row.client_name || existing.customerName,
              phone: row.phone || existing.phone,
              city: row.city || existing.city,
              address: row.address || existing.address,
              productId,
              quantity: row.quantity || existing.quantity,
              orderDate: excelDateToInput(row.created_at || existing.orderDate),
              status: nextShippingStatus || confirmationStatus,
              confirmationStatus,
              shippingStatus: nextShippingStatus,
              confirmation_status_raw: row.confirmation_status_raw || existing.confirmation_status_raw,
              shipping_status_raw: row.shipping_status_raw || existing.shipping_status_raw,
              confirmation_updated_at: row.confirmation_updated_at || existing.confirmation_updated_at || null,
              amount_tsh: row.amount_tsh || existing.amount_tsh || existing.orderTotalTzs || 0,
              amount_usd: row.amount_usd || existing.amount_usd || 0,
              orderTotalTzs: row.amount_tsh || existing.orderTotalTzs || 0,
              sourceOrderId: existing.sourceOrderId || row.order_id || null,
              order_id: row.order_id || existing.order_id || existing.sourceOrderId || null,
              import_key: row.import_key || existing.import_key || null,
              normalized_phone: row.normalized_phone || existing.normalized_phone,
              product_ref: row.product_ref || existing.product_ref,
              raw_row_data: row.raw_row_data || existing.raw_row_data,
              extra_fields: row.extra_fields || existing.extra_fields || {},
              line_item_index: row.line_item_index ?? existing.line_item_index ?? 0,
              multi_product_revenue_allocated: Boolean(row.multi_product_revenue_allocated || existing.multi_product_revenue_allocated),
              importSource: "excel",
              lastImportedAt: importFinishedAt,
              lastShippingImportedAt: shippingChanged ? importFinishedAt : existing.lastShippingImportedAt || null,
              updatedAt: row.updated_at || importFinishedAt,
              history: statusChanged || shippingChanged || amountChanged
                ? appendCustomerHistory(
                    existing,
                    buildHistoryEntry({
                      action: "orders_import_updated",
                      source: "excel-orders",
                      details: [
                        statusChanged ? `Confirmation ${formatStatusLabel(previousConfirmation)} -> ${formatStatusLabel(confirmationStatus)}` : "",
                        shippingChanged ? `Shipping ${formatStatusLabel(previousShipping || "blank")} -> ${formatStatusLabel(nextShippingStatus || "blank")}` : "",
                        amountChanged ? `Amount synced to ${formatTZS(row.amount_tsh || 0)}` : "",
                      ].filter(Boolean).join(" | "),
                    })
                  )
                : existing.history,
            });
            nextCustomers[existingIndex] = updatedOrder;
            const orderMov = buildOrderStatusMovement(existing, updatedOrder);
            if (orderMov) orderMovements.push(orderMov);
            updatedCount += 1;
            return;
          }

          const newOrder = sanitizeCustomerRecord({
            id: buildNextId(nextCustomers, "C"),
            customerName: row.client_name,
            phone: row.phone,
            city: row.city,
            address: row.address,
            productId: productId || "",
            product_name_raw: productId ? "" : String(row.product_name || row.product_ref || "").trim(),
            quantity: row.quantity,
            orderDate: excelDateToInput(row.created_at),
            paymentMethod: "COD",
            status: shippingStatus || confirmationStatus,
            confirmationStatus,
            shippingStatus: ensureShippingStatusForConfirmed(confirmationStatus, shippingStatus),
            confirmation_status_raw: row.confirmation_status_raw,
            shipping_status_raw: row.shipping_status_raw,
            confirmation_updated_at: row.confirmation_updated_at || null,
            amount_tsh: row.amount_tsh,
            amount_usd: row.amount_usd,
            orderTotalTzs: row.amount_tsh,
            notes: "",
            sourceOrderId: row.order_id || null,
            order_id: row.order_id || null,
            import_key: row.import_key,
            normalized_phone: row.normalized_phone,
            product_ref: row.product_ref,
            raw_row_data: row.raw_row_data,
            extra_fields: row.extra_fields || {},
            line_item_index: row.line_item_index || 0,
            multi_product_revenue_allocated: Boolean(row.multi_product_revenue_allocated),
            importSource: "excel",
            lastImportedAt: importFinishedAt,
            lastShippingImportedAt: shippingStatus ? importFinishedAt : null,
            updatedAt: row.updated_at || importFinishedAt,
            assignedTo: "Call Center",
            history: [
              buildHistoryEntry({
                action: "orders_import_created",
                source: "excel-orders",
                details: `Imported with ${formatStatusLabel(confirmationStatus)}${shippingStatus ? ` | shipping ${formatStatusLabel(shippingStatus)}` : ""}${!productId ? ` | unmatched product: ${String(row.product_name || "").trim()}` : ""}`,
              }),
            ],
          });
          nextCustomers.unshift(newOrder);
          const newOrderMov = buildOrderStatusMovement(null, newOrder);
          if (newOrderMov) orderMovements.push(newOrderMov);
          report.newLeadsAdded += 1;
          createdCount += 1;
        });

        const sanitizedNextCustomers = nextCustomers.map(sanitizeCustomerRecord).filter(Boolean);
        const nextImportMeta = { ...(importMeta || getDefaultImportMeta()), lastOrdersImportAt: importFinishedAt };
        setCustomers(sanitizedNextCustomers);
        setImportMeta(nextImportMeta);
        if (orderMovements.length > 0) {
          setStockMovements((prev) => applyOrderMovementsToState(orderMovements, prev));
        }
        const nextSnapshot = {
          ...(latestSharedStateRef.current || getDefaultCloudWorkspaceState()),
          customers: sanitizedNextCustomers,
          importMeta: nextImportMeta,
        };
        latestSharedStateRef.current = nextSnapshot;
        void persistSharedSnapshot(nextSnapshot, {
          progressNotice: "Saving imported leads to cloud...",
          successNotice: "Cloud leads import synced",
          failurePrefix: "Cloud leads import sync failed",
        });
        const unmatchedCount = unmatchedProducts.size;
        setOrdersImportNotice(
          `Excel imported: ${createdCount} new, ${updatedCount} updated, ${skippedCount} skipped${unmatchedCount ? `, ${unmatchedCount} unmatched product(s) imported anyway` : ""}.`
        );
        const importRecord = {
          detectedHeaders,
          reasonCounts: {
            missingName: 0,
            missingPhone: report.missingPhoneRows,
            missingProduct: report.missingProductRows,
            unknownProduct: unmatchedCount,
            missingCode: report.missingCodeRows,
            missingAmount: report.missingAmountRows,
            unknownConfirmationStatuses: unknownConfirmationStatuses.size,
            unknownShippingStatuses: unknownShippingStatuses.size,
            statusChangesDetected: report.statusChangesDetected,
          },
          unmatchedProducts: Array.from(unmatchedProducts).slice(0, 10),
          importedAt: new Date().toISOString(),
          summary: `${createdCount} new, ${updatedCount} updated, ${skippedCount} skipped`,
        };
        setOrdersImportDetails(importRecord);
        setOrdersImportHistory((prev) => [importRecord, ...prev].slice(0, 20));
      } catch (error) {
        const message = error instanceof Error ? error.message : "Excel import failed";
        setOrdersImportNotice(`Excel import failed: ${message}`);
        setOrdersImportDetails(null);
      } finally {
        event.target.value = "";
      }
    },
    [customers, findMatchingCustomerIndex, importMeta, persistSharedSnapshot, resolveImportedProductId]
  );

  const importRevenueFromExcel = useCallback(
    async (event) => {
      const file = event.target.files?.[0];
      if (!file) return;
      setRevenueImportNotice("Reading file…");
      try {
        const buffer = await file.arrayBuffer();
        const workbook = XLSX.read(buffer, { type: "array", cellDates: false });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const rawRows = XLSX.utils.sheet_to_json(firstSheet, { defval: "" });
        if (!rawRows.length) { setRevenueImportNotice("File is empty."); return; }

        // Normalize header names once
        const headerMap = {};
        Object.keys(rawRows[0]).forEach((key) => { headerMap[normalizeHeaderName(key)] = key; });
        const getCol = (aliases) => aliases.map((a) => headerMap[a]).find(Boolean);

        const codeCol = getCol(["code", "order id", "orderid", "order no", "reference", "ref"]);
        const amountCol = getCol(["amount", "total", "total amount", "montant", "prix"]);
        const shipCol = getCol(["shipping status", "delivery status", "statut livraison", "statut expedition"]);
        const cityCol = getCol(["city", "ville", "region", "localite"]);
        const qtyCol = getCol(["quantity", "qty", "quantite", "qte"]);

        const exchangeRate = Number(serviceForm?.exchangeRate || USD_TO_TZS);

        // Start with the existing rows map (deduplicate by CODE)
        const mergedRows = { ...revenueImportRows };
        let newCount = 0;
        let updatedCount = 0;
        let _skippedNoCode = 0;
        const importedAt = new Date().toISOString();

        for (const row of rawRows) {
          const code = codeCol ? String(row[codeCol] || "").trim() : "";
          if (!code) { _skippedNoCode++; continue; }

          const amountTsh = parseLooseNumber(amountCol ? String(row[amountCol] || "0") : "0");
          const status = shipCol ? String(row[shipCol] || "") : "";
          const city = cityCol ? String(row[cityCol] || "") : "";
          const qty = Math.max(1, parseLooseNumber(qtyCol ? String(row[qtyCol] || "1") : "1") || 1);

          if (mergedRows[code]) {
            updatedCount++;
          } else {
            newCount++;
          }
          mergedRows[code] = { amount_tsh: amountTsh, status, city, qty, imported_at: importedAt };
        }

        // Recalculate totals from ALL stored rows (delivered only) — prevents any double-counting
        let revenueTsh = 0;
        let deliveredCount = 0;
        let deliveredUnits = 0;
        let serviceChargesUsd = 0;
        for (const r of Object.values(mergedRows)) {
          if (!isShippingDelivered(r.status)) continue;
          revenueTsh += Number(r.amount_tsh || 0);
          const qty = Math.max(1, Number(r.qty || 1));
          deliveredCount++;
          deliveredUnits += qty;
          serviceChargesUsd += qty * (String(r.city || "").toLowerCase().includes("dar") ? 8 : 9);
        }

        const data = {
          revenueTsh,
          revenueUsd: revenueTsh / exchangeRate,
          deliveredCount,
          deliveredUnits,
          serviceChargesUsd,
          importedAt,
          rowCount: rawRows.length,
          totalStoredOrders: Object.keys(mergedRows).length,
          lastImportNewCount: newCount,
          lastImportUpdatedCount: updatedCount,
        };

        setRevenueImportRows(mergedRows);
        setRevenueImport(data);
        writeLS(LS_REVENUE_ROWS, mergedRows);
        writeLS(LS_REVENUE_IMPORT, data);
        setRevenueImportNotice(
          `Last import: ${new Date().toLocaleDateString()} — ${newCount} new orders added, ${updatedCount} orders updated, ${deliveredCount} total delivered`
        );
        if (supabaseEnabled) {
          saveRevenueImportRowsToSupabase(mergedRows).catch(() => {});
          saveRevenueImportToSupabase(data).catch(() => {});
        }
      } catch (err) {
        setRevenueImportNotice(`Import failed: ${err.message}`);
      }
      event.target.value = "";
    },
    [serviceForm, revenueImportRows]
  );

  const importShippingFromExcel = useCallback(
    async (event) => {
      const file = event.target.files?.[0];
      if (!file) return;

      try {
        const buffer = await file.arrayBuffer();
        const workbook = XLSX.read(buffer, { type: "array", cellDates: false });
        const firstSheetName = workbook.SheetNames[0];
        const firstSheet = workbook.Sheets[firstSheetName];
        const rows = XLSX.utils.sheet_to_json(firstSheet, { defval: "" });

        if (!rows.length) {
          setShippingImportNotice("The shipping Excel file is empty.");
          setShippingImportDetails(null);
          return;
        }

        const nextCustomers = [...customers];
        const importFinishedAt = new Date().toISOString();
        const { parsedRows, report } = parseImportedExcelRows(rows, {
          exchangeRate: USD_TO_TZS,
          resolveProductId: resolveImportedProductId,
        });
        let updatedCount = 0;
        let unchangedCount = 0;
        let skippedCount = 0;
        const reasonCounts = {
          missingStatus: 0,
          unmatchedOrder: 0,
        };
        const unmatchedExamples = new Set();
        const detectedHeaders = Object.keys(rows[0] || {});
        const shippingMovements = [];

        parsedRows.forEach((row) => {
          const nextStatus = normalizeOrderStatus(row.shipping_status_raw);
          if (!String(row.shipping_status_raw || "").trim()) {
            reasonCounts.missingStatus += 1;
            skippedCount += 1;
            return;
          }

          const existingIndex = findMatchingCustomerIndex(nextCustomers, {
            import_key: row.import_key,
            sourceOrderId: row.order_id,
            customerName: row.client_name,
            phone: row.phone,
            productId: row.productId,
            product_ref: row.product_ref,
            quantity: row.quantity,
            orderDate: excelDateToInput(row.created_at),
            line_item_index: row.line_item_index,
          });

          if (existingIndex < 0) {
            reasonCounts.unmatchedOrder += 1;
            skippedCount += 1;
            unmatchedExamples.add(row.order_id || row.phone || row.client_name || String(row.product_name || "").trim() || "Unknown row");
            return;
          }

          const existing = nextCustomers[existingIndex];
          const currentStatus = normalizeOrderStatus(getCustomerShippingStatus(existing));

          if (currentStatus === nextStatus) {
            unchangedCount += 1;
            return;
          }

          const updatedShippingOrder = sanitizeCustomerRecord({
            ...existing,
            shippingStatus: nextStatus,
            status: nextStatus,
            confirmationStatus: isConfirmationConfirmed(existing.confirmationStatus) ? existing.confirmationStatus : "confirmed",
            orderTotalTzs: row.amount_tsh || Number(existing.orderTotalTzs || 0),
            amount_tsh: row.amount_tsh || existing.amount_tsh || Number(existing.orderTotalTzs || 0),
            amount_usd: row.amount_usd || existing.amount_usd || 0,
            sourceOrderId: existing.sourceOrderId || row.order_id || null,
            order_id: row.order_id || existing.order_id || existing.sourceOrderId || null,
            import_key: row.import_key || existing.import_key || null,
            product_ref: row.product_ref || existing.product_ref,
            raw_row_data: row.raw_row_data || existing.raw_row_data,
            extra_fields: row.extra_fields || existing.extra_fields || {},
            lastShippingImportedAt: importFinishedAt,
            actualDeliveryDate: isShippingDelivered(nextStatus) ? existing.actualDeliveryDate || getTodayString() : existing.actualDeliveryDate || "",
            importSource: existing.importSource || "excel",
            assignedTo: existing.assignedTo || "Shipping Team",
            updatedAt: row.updated_at || importFinishedAt,
            history: appendCustomerHistory(
              existing,
              buildHistoryEntry({
                action: "shipping_import_updated",
                source: "excel-shipping",
                details: `Shipping ${formatStatusLabel(currentStatus)} -> ${formatStatusLabel(nextStatus)}`,
              })
            ),
          });
          nextCustomers[existingIndex] = updatedShippingOrder;
          const shipMov = buildOrderStatusMovement(existing, updatedShippingOrder);
          if (shipMov) shippingMovements.push(shipMov);
          updatedCount += 1;
        });

        const sanitizedNextCustomers = nextCustomers.map(sanitizeCustomerRecord).filter(Boolean);
        const nextImportMeta = { ...(importMeta || getDefaultImportMeta()), lastShippingImportAt: importFinishedAt };
        setCustomers(sanitizedNextCustomers);
        if (shippingMovements.length > 0) {
          setStockMovements((prev) => applyOrderMovementsToState(shippingMovements, prev));
        }
        setImportMeta(nextImportMeta);
        const nextSnapshot = {
          ...(latestSharedStateRef.current || getDefaultCloudWorkspaceState()),
          customers: sanitizedNextCustomers,
          importMeta: nextImportMeta,
        };
        latestSharedStateRef.current = nextSnapshot;
        void persistSharedSnapshot(nextSnapshot, {
          progressNotice: "Saving shipping import to cloud...",
          successNotice: "Cloud shipping import synced",
          failurePrefix: "Cloud shipping import sync failed",
        });
        setShippingImportNotice(
          `Shipping Excel imported: ${updatedCount} updated, ${unchangedCount} unchanged, ${skippedCount} skipped.`
        );
        setShippingImportDetails({
          detectedHeaders,
          reasonCounts: {
            ...reasonCounts,
            missingCode: report.missingCodeRows,
            missingPhone: report.missingPhoneRows,
            missingAmount: report.missingAmountRows,
            missingProduct: report.missingProductRows,
          },
          unmatchedExamples: Array.from(unmatchedExamples).slice(0, 6),
        });
      } catch (error) {
        const message = error instanceof Error ? error.message : "Shipping Excel import failed";
        setShippingImportNotice(`Shipping Excel import failed: ${message}`);
        setShippingImportDetails(null);
      } finally {
        event.target.value = "";
      }
    },
    [customers, findMatchingCustomerIndex, importMeta, persistSharedSnapshot, resolveImportedProductId]
  );

  const saveCustomerOrder = async () => {
    if (!customerForm.customerName.trim() || !customerForm.phone.trim()) return;
    if (!products.length || !getProduct(customerForm.productId)) {
      alert("Add a product first before saving a customer order.");
      return;
    }

    const newCustomer = sanitizeCustomerRecord({
      id: createCustomerId(),
      customerName: customerForm.customerName.trim(),
      phone: customerForm.phone.trim(),
      city: customerForm.city.trim(),
      address: customerForm.address.trim(),
      productId: customerForm.productId,
      quantity: Math.max(1, Number(customerForm.quantity || 1)),
      orderDate: customerForm.orderDate || getTodayString(),
      paymentMethod: customerForm.paymentMethod,
      status: customerForm.status,
      confirmationStatus: customerForm.status,
      shippingStatus: ensureShippingStatusForConfirmed(customerForm.status, ""),
      orderTotalTzs: customerFormPricing.totalPrice,
      notes: customerForm.notes.trim(),
      leadSource: customerForm.leadSource,
      campaignName: customerForm.campaignName.trim(),
      adsetName: customerForm.adsetName.trim(),
      creativeName: customerForm.creativeName.trim(),
      priority: customerForm.priority,
      customerType: customerForm.customerType,
      callAttempts: Math.max(0, Number(customerForm.callAttempts || 0)),
      cancelReason: customerForm.cancelReason.trim(),
      unreachedReason: customerForm.unreachedReason.trim(),
      carrierName: customerForm.carrierName.trim(),
      trackingNumber: customerForm.trackingNumber.trim(),
      expectedDeliveryDate: customerForm.expectedDeliveryDate || "",
      returnReason: customerForm.returnReason.trim(),
      sourceOrderId: null,
      importSource: "manual",
      lastImportedAt: null,
      assignedTo: "Call Center",
      history: [
        buildHistoryEntry({
          action: "manual_order_created",
          source: "manual",
          details: `Created with ${formatStatusLabel(customerForm.status)} | ${formatTZS(customerFormPricing.totalPrice)}`,
        }),
      ],
    });

    setCustomers((prev) => [newCustomer, ...prev]);
    setCustomerForm(getEmptyCustomerForm(products[0]?.id || "P001"));
    setActivePage("customersOrders");
  };

  const deleteCustomerOrder = (customerId) => {
    setCustomers((prev) => prev.filter((c) => c.id !== customerId));
    if (customerHistoryTargetId === customerId) setCustomerHistoryTargetId("");
  };

  const deleteSelectedCustomerOrders = () => {
    if (selectedCustomerIds.length === 0) return;
    setCustomers((prev) => prev.filter((customer) => !selectedCustomerIds.includes(customer.id)));
    setSelectedCustomerIds([]);
    if (selectedCustomerIds.includes(customerHistoryTargetId)) setCustomerHistoryTargetId("");
  };

  const updateCustomerStatus = (customerId, nextStatus) => {
    const existingForMovement = customers.find((c) => c.id === customerId);
    setCustomers((prev) =>
      prev.map((c) =>
        c.id === customerId
          ? sanitizeCustomerRecord({
              ...c,
              status: nextStatus,
              confirmationStatus: nextStatus,
              shippingStatus: ensureShippingStatusForConfirmed(nextStatus, c.shippingStatus),
              history: appendCustomerHistory(
                c,
                buildHistoryEntry({
                  action: "confirmation_status_updated",
                  source: "manual",
                  details: `Confirmation ${formatStatusLabel(getCustomerConfirmationStatus(c))} -> ${formatStatusLabel(nextStatus)}`,
                })
              ),
            })
          : c
      )
    );
    if (existingForMovement) {
      const previewOrder = {
        ...existingForMovement,
        confirmationStatus: nextStatus,
        shippingStatus: ensureShippingStatusForConfirmed(nextStatus, existingForMovement.shippingStatus),
      };
      const mov = buildOrderStatusMovement(existingForMovement, previewOrder);
      if (mov) setStockMovements((prev) => applyOrderMovementsToState([mov], prev));
    }
  };

  const updateCustomerShippingStatus = (customerId, nextStatus) => {
    const existingForShipMovement = customers.find((c) => c.id === customerId);
    setCustomers((prev) =>
      prev.map((c) =>
        c.id === customerId
          ? sanitizeCustomerRecord({
              ...c,
              shippingStatus: nextStatus,
              status: nextStatus,
              confirmationStatus: isConfirmationConfirmed(c.confirmationStatus) ? c.confirmationStatus : "confirmed",
              lastShippingImportedAt: new Date().toISOString(),
              actualDeliveryDate: isShippingDelivered(nextStatus) ? c.actualDeliveryDate || getTodayString() : c.actualDeliveryDate || "",
              assignedTo: c.assignedTo || "Shipping Team",
              history: appendCustomerHistory(
                c,
                buildHistoryEntry({
                  action: "shipping_status_updated",
                  source: "manual",
                  details: `Shipping ${formatStatusLabel(getCustomerShippingStatus(c) || "to-prepare")} -> ${formatStatusLabel(nextStatus)}`,
                })
              ),
            })
          : c
      )
    );
    if (existingForShipMovement) {
      const previewShipOrder = {
        ...existingForShipMovement,
        shippingStatus: nextStatus,
        confirmationStatus: isConfirmationConfirmed(existingForShipMovement.confirmationStatus) ? existingForShipMovement.confirmationStatus : "confirmed",
      };
      const shipMov = buildOrderStatusMovement(existingForShipMovement, previewShipOrder);
      if (shipMov) setStockMovements((prev) => applyOrderMovementsToState([shipMov], prev));
    }
  };

  const assignCustomerOwner = (customerId, nextOwner) => {
    setCustomers((prev) =>
      prev.map((customer) =>
        customer.id === customerId
          ? sanitizeCustomerRecord({
              ...customer,
              assignedTo: nextOwner,
              history: appendCustomerHistory(
                customer,
                buildHistoryEntry({
                  action: "owner_assigned",
                  source: "manual",
                  details: nextOwner ? `Assigned to ${nextOwner}` : "Owner cleared",
                })
              ),
            })
          : customer
      )
    );
  };

  const updateCustomersBulkConfirmationStatus = () => {
    if (!selectedCustomerIds.length || !bulkCustomerStatus) return;
    const targetIds = new Set(selectedCustomerIds);
    setCustomers((prev) =>
      prev.map((customer) =>
        targetIds.has(customer.id)
          ? sanitizeCustomerRecord({
              ...customer,
              status: bulkCustomerStatus,
              confirmationStatus: bulkCustomerStatus,
              shippingStatus: ensureShippingStatusForConfirmed(bulkCustomerStatus, customer.shippingStatus),
              history: appendCustomerHistory(
                customer,
                buildHistoryEntry({
                  action: "bulk_confirmation_update",
                  source: "bulk",
                  details: `Confirmation set to ${formatStatusLabel(bulkCustomerStatus)}`,
                })
              ),
            })
          : customer
      )
    );
  };

  const assignCustomersBulkOwner = () => {
    if (!selectedCustomerIds.length) return;
    const targetIds = new Set(selectedCustomerIds);
    setCustomers((prev) =>
      prev.map((customer) =>
        targetIds.has(customer.id)
          ? sanitizeCustomerRecord({
              ...customer,
              assignedTo: bulkCustomerOwner,
              history: appendCustomerHistory(
                customer,
                buildHistoryEntry({
                  action: "bulk_owner_assignment",
                  source: "bulk",
                  details: bulkCustomerOwner ? `Assigned to ${bulkCustomerOwner}` : "Owner cleared",
                })
              ),
            })
          : customer
      )
    );
  };

  const deleteSelectedShippingOrders = () => {
    if (selectedShippingIds.length === 0) return;
    setCustomers((prev) => prev.filter((customer) => !selectedShippingIds.includes(customer.id)));
    setSelectedShippingIds([]);
    if (selectedShippingIds.includes(customerHistoryTargetId)) setCustomerHistoryTargetId("");
  };

  const updateShippingBulkStatus = () => {
    if (!selectedShippingIds.length || !bulkShippingStatus) return;
    const targetIds = new Set(selectedShippingIds);
    setCustomers((prev) =>
      prev.map((customer) =>
        targetIds.has(customer.id)
          ? sanitizeCustomerRecord({
              ...customer,
              shippingStatus: bulkShippingStatus,
              status: bulkShippingStatus,
              confirmationStatus: isConfirmationConfirmed(customer.confirmationStatus) ? customer.confirmationStatus : "confirmed",
              lastShippingImportedAt: new Date().toISOString(),
              assignedTo: customer.assignedTo || "Shipping Team",
              history: appendCustomerHistory(
                customer,
                buildHistoryEntry({
                  action: "bulk_shipping_update",
                  source: "bulk",
                  details: `Shipping set to ${formatStatusLabel(bulkShippingStatus)}`,
                })
              ),
            })
          : customer
      )
    );
  };

  const markDubaiStockArrived = (productId) => {
    setProducts((prev) => {
      const nextProducts = prev.map((p) =>
        p.id === productId
          ? {
              ...p,
              stockArrivalStatus: "arrived",
              stockArrivedAt: getTodayString(),
              nextArrivalCheckDate: null,
            }
          : p
      );
      const nextSnapshot = {
        ...(latestSharedStateRef.current || getDefaultCloudWorkspaceState()),
        products: nextProducts.map(sanitizeProductRecord),
      };
      latestSharedStateRef.current = nextSnapshot;
      void persistSharedSnapshot(nextSnapshot, {
        progressNotice: "Saving stock arrival update...",
        successNotice: "Stock arrival updated",
        failurePrefix: "Stock arrival update failed",
      });
      return nextProducts;
    });
  };

  const markDubaiStockNotYet = (productId) => {
    setProducts((prev) => {
      const nextProducts = prev.map((p) =>
        p.id === productId
          ? {
              ...p,
              stockArrivalStatus: "pending",
              nextArrivalCheckDate: addDaysToDateString(getTodayString(), 1),
            }
          : p
      );
      const nextSnapshot = {
        ...(latestSharedStateRef.current || getDefaultCloudWorkspaceState()),
        products: nextProducts.map(sanitizeProductRecord),
      };
      latestSharedStateRef.current = nextSnapshot;
      void persistSharedSnapshot(nextSnapshot, {
        progressNotice: "Saving stock arrival update...",
        successNotice: "Stock arrival updated",
        failurePrefix: "Stock arrival update failed",
      });
      return nextProducts;
    });
  };

  const exportReport = () => {
    const reportLines = [
      "Tanzania Ecom Tracker Report",
      `Generated at: ${new Date().toLocaleString()}`,
      "",
      "Summary",
      `- Total products: ${products.length}`,
      `- Total tracking rows: ${tracking.length}`,
      `- Total customer orders: ${dashboardDateSummary.totalLeads}`,
      `- Confirmed orders: ${dashboardDateSummary.totalConfirmedOrders}`,
      `- Delivered orders: ${dashboardDateSummary.totalDeliveredOrders}`,
      `- Revenue: ${formatTZS(dashboardDateSummary.totalRevenue)}`,
      "",
      "Products",
      ...productDashboard.map((product) =>
        `- ${product.name} | stock=${product.availableStock} | delivered=${product.deliveredUnits} units | profit=${formatTZS(product.profit)} | decision=${product.decision}`
      ),
      "",
      "Recent Orders",
      ...operationalCustomers.slice(0, 10).map((customer) => {
        const product = getProduct(customer.productId);
        return `- ${customer.customerName} | ${product?.name || customer.productId} | qty=${customer.quantity} | status=${customer.status} | date=${customer.orderDate}`;
      }),
    ];

    const blob = new Blob([reportLines.join("\n")], { type: "text/plain;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `tanzania-ecom-report-${new Date().toISOString().slice(0, 10)}.txt`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  };

  const customersDashboard = useMemo(() => {
    const totalLeads = serviceLeadCustomers.length;
    const confirmedLeads = serviceLeadCustomers.filter((c) => isConfirmationConfirmed(getCustomerConfirmationStatus(c))).length;
    const deliveredLeadCustomers = serviceLeadCustomers.filter(
      (c) =>
        isConfirmationConfirmed(getCustomerConfirmationStatus(c)) &&
        isShippingDelivered(getCustomerShippingStatus(c))
    );
    const deliveredLeads = deliveredLeadCustomers.length;
    const newLeads = serviceLeadCustomers.filter((c) => isConfirmationNew(getCustomerConfirmationStatus(c))).length;
    const cancelledLeads = serviceLeadCustomers.filter((c) => isConfirmationCancelled(getCustomerConfirmationStatus(c))).length;
    const otherLeads = totalLeads - confirmedLeads - cancelledLeads - newLeads;

    const totalRevenue = deliveredLeadCustomers
      .reduce((sum, c) => {
        const product = products.find((p) => p.id === c.productId);
        return sum + getCustomerOrderTotalTzs(c, product);
      }, 0);

    const confirmationRate = totalLeads > 0 ? (confirmedLeads / totalLeads) * 100 : 0;
    const deliveryRate = confirmedLeads > 0 ? (deliveredLeads / confirmedLeads) * 100 : 0;

    return {
      totalOrders: totalLeads,
      totalQty: serviceLeadCustomers.reduce((sum, c) => sum + Number(c.quantity || 0), 0),
      confirmedOrders: confirmedLeads,
      deliveredOrders: deliveredLeads,
      newOrders: newLeads,
      cancelledOrders: cancelledLeads,
      otherOrders: Math.max(0, otherLeads),
      totalRevenue,
      confirmationRate,
      deliveryRate,
    };
  }, [products, serviceLeadCustomers]);

  const confirmationMetrics = useMemo(() => {
    const total = serviceLeadCustomers.length;
    const confirmed = serviceLeadCustomers.filter((c) => isConfirmationConfirmed(getCustomerConfirmationStatus(c))).length;
    const noReply = serviceLeadCustomers.filter((c) => isNoReplyStatus(getCustomerConfirmationStatus(c))).length;
    const cancelled = serviceLeadCustomers.filter((c) => isConfirmationCancelled(getCustomerConfirmationStatus(c))).length;
    const invalid = serviceLeadCustomers.filter((c) => isInvalidLeadStatus(getCustomerConfirmationStatus(c))).length;
    const stockHold = serviceLeadCustomers.filter((c) => isStockHoldStatus(getCustomerConfirmationStatus(c))).length;
    const unknown = Math.max(0, total - confirmed - noReply - cancelled - invalid - stockHold);
    const confirmationRate = total > 0 ? (confirmed / total) * 100 : 0;
    const productMap = {};
    for (const c of serviceLeadCustomers) {
      const pid = c.productId || "__none__";
      if (!productMap[pid]) productMap[pid] = { confirmed: 0, noReply: 0, cancelled: 0, total: 0 };
      productMap[pid].total++;
      const s = getCustomerConfirmationStatus(c);
      if (isConfirmationConfirmed(s)) productMap[pid].confirmed++;
      else if (isNoReplyStatus(s)) productMap[pid].noReply++;
      else if (isConfirmationCancelled(s)) productMap[pid].cancelled++;
    }
    const productRows = Object.entries(productMap).map(([pid, data]) => ({
      productId: pid,
      productName: pid === "__none__" ? "Unknown" : (products.find((p) => p.id === pid)?.name || pid),
      ...data,
      confirmationRate: data.total > 0 ? (data.confirmed / data.total) * 100 : 0,
    })).sort((a, b) => b.total - a.total);
    const cityMap = {};
    for (const c of serviceLeadCustomers) {
      const city = String(c.city || "").trim() || "Unknown";
      if (!cityMap[city]) cityMap[city] = { confirmed: 0, total: 0 };
      cityMap[city].total++;
      if (isConfirmationConfirmed(getCustomerConfirmationStatus(c))) cityMap[city].confirmed++;
    }
    const cityRows = Object.entries(cityMap).map(([city, data]) => ({
      city, ...data,
      confirmationRate: data.total > 0 ? (data.confirmed / data.total) * 100 : 0,
    })).sort((a, b) => b.total - a.total).slice(0, 25);
    return { total, confirmed, noReply, cancelled, invalid, stockHold, unknown, confirmationRate, productRows, cityRows };
  }, [serviceLeadCustomers, products]);

  const ordersPageCityList = useMemo(
    () => [...new Set(operationalCustomers.map((c) => String(c.city || "").trim()).filter(Boolean))].sort(),
    [operationalCustomers]
  );

  const liveAutomationSummary = useMemo(() => {
    const totalAdSpendTzs = tracking.reduce((sum, row) => sum + Number(row.adSpend || 0), 0);
    const deliveredUnits = productDashboard.reduce((sum, product) => sum + Number(product.deliveredUnits || 0), 0);
    const reservedUnits = productDashboard.reduce((sum, product) => sum + Number(product.reservedStock || 0), 0);
    const availableUnits = productDashboard.reduce((sum, product) => sum + Number(product.availableStock || 0), 0);
    const importCostDeliveredTzs = productDashboard.reduce(
      (sum, product) => sum + Number(product.totalProductCostTzs || product.totalProductCost || 0),
      0
    );
    const localDeliveryCostTzs = productDashboard.reduce(
      (sum, product) => sum + Number(product.totalDeliveryCostTzs || product.totalDeliveryCost || 0),
      0
    );

    return {
      totalLeads: customersDashboard.totalOrders,
      confirmedOrders: customersDashboard.confirmedOrders,
      deliveredOrders: customersDashboard.deliveredOrders,
      deliveredUnits,
      totalRevenueTzs: customersDashboard.totalRevenue,
      totalAdSpendTzs,
      importCostDeliveredTzs,
      localDeliveryCostTzs,
      totalOperationalCostTzs: importCostDeliveredTzs + localDeliveryCostTzs + totalAdSpendTzs,
      grossProfitTzs: customersDashboard.totalRevenue - importCostDeliveredTzs - localDeliveryCostTzs - totalAdSpendTzs,
      reservedUnits,
      availableUnits,
    };
  }, [customersDashboard, productDashboard, tracking]);

  const _liveServiceDataset = useMemo(() => {
    const config = serviceCountryData[selectedService]?.[selectedCountry];
    if (!config) return null;

    const totalLeads = Number(liveAutomationSummary.totalLeads || 0);
    const confirmed = Number(liveAutomationSummary.confirmedOrders || 0);
    const delivered = Number(liveAutomationSummary.deliveredOrders || 0);
    const deliveredUnits = Number(liveAutomationSummary.deliveredUnits || 0);
    const revenueTzs = Number(liveAutomationSummary.totalRevenueTzs || 0);
    const revenueUsd = revenueTzs / config.usdToTzs;
    const adSpendUsd = Number(liveAutomationSummary.totalAdSpendTzs || 0) / config.usdToTzs;
    const productCostTotalUsd = Number(liveAutomationSummary.importCostDeliveredTzs || 0) / config.usdToTzs;
    const localDeliveryCostUsd = Number(liveAutomationSummary.localDeliveryCostTzs || 0) / config.usdToTzs;
    const serviceFeeUsd = revenueUsd * (config.serviceFeePercent / 100);
    const deliveryFeesUsd = delivered * config.deliveryFeeUsdPerDelivered;
    const totalServiceChargeUsd = serviceFeeUsd + deliveryFeesUsd;
    const costPerLeadUsd = totalLeads > 0 ? adSpendUsd / totalLeads : 0;
    const adCostPerDeliveredUsd = delivered > 0 ? adSpendUsd / delivered : 0;
    const totalProfitUsd = revenueUsd - productCostTotalUsd - localDeliveryCostUsd - totalServiceChargeUsd - adSpendUsd;
    const totalProfitTzs = totalProfitUsd * config.usdToTzs;
    const profitPerOrderUsd = delivered > 0 ? totalProfitUsd / delivered : 0;
    const profitPerPieceUsd = deliveredUnits > 0 ? totalProfitUsd / deliveredUnits : 0;
    const profitPerPieceTzs = profitPerPieceUsd * config.usdToTzs;
    const confirmationRate = totalLeads > 0 ? confirmed / totalLeads : 0;
    const deliveryRate = confirmed > 0 ? delivered / confirmed : 0;
    const grossMarginPerDeliveredUsd = delivered > 0 ? (revenueUsd / delivered) - ((productCostTotalUsd + localDeliveryCostUsd) / delivered) - config.deliveryFeeUsdPerDelivered : 0;
    const breakEvenCplUsd = confirmationRate > 0 && deliveryRate > 0 ? grossMarginPerDeliveredUsd * confirmationRate * deliveryRate : 0;
    const breakEvenPriceUsd = delivered > 0 ? (productCostTotalUsd + localDeliveryCostUsd + totalServiceChargeUsd + adSpendUsd) / delivered : 0;
    const marginPercent = revenueUsd > 0 ? (totalProfitUsd / revenueUsd) * 100 : 0;

    let decision = "WATCH";
    if (totalProfitUsd > 0 && deliveryRate >= 0.5) decision = "GOOD PRODUCT";
    if (totalProfitUsd < 0) decision = "BAD PRODUCT";

    const score = Math.max(
      0,
      Math.min(
        100,
        Math.round(
          (marginPercent > 0 ? 35 : 0) +
            (deliveryRate >= 0.5 ? 25 : 0) +
            (confirmationRate >= 0.5 ? 20 : 0) +
            (revenueUsd > 0 ? 20 : 0)
        )
      )
    );

    return {
      totalLeads,
      confirmed,
      delivered,
      deliveredUnits,
      adSpendUsd,
      revenueUsd,
      revenueTzs,
      productCostTotalUsd,
      localDeliveryCostUsd,
      deliveryFeesUsd,
      serviceFeeUsd,
      totalServiceChargeUsd,
      costPerLeadUsd,
      adCostPerDeliveredUsd,
      profitPerOrderUsd,
      profitPerPieceUsd,
      profitPerPieceTzs,
      totalProfitUsd,
      totalProfitTzs,
      breakEvenCplUsd,
      breakEvenPriceUsd,
      marginPercent,
      decision,
      score,
      confirmationRate,
      deliveryRate,
    };
  }, [liveAutomationSummary, selectedCountry, selectedService]);

  const deferredCustomerSearch = useDeferredValue(customerListFilters.search);

  const compactCustomerRows = useMemo(() => {
    const searchValue = normalizeHeaderName(deferredCustomerSearch);

    return operationalCustomers
      .map((customer) => {
        const product = products.find((p) => p.id === customer.productId);
        const totalValue = getCustomerOrderTotalTzs(customer, product);
        const normalizedStatus = getCustomerConfirmationStatus(customer);
        return {
          ...customer,
          status: normalizedStatus,
          statusLabel: confirmationStatusMap[normalizedStatus]?.label || formatStatusLabel(normalizedStatus),
          productName: product?.name || customer.productId,
          totalValue,
        };
      })
      .filter((customer) => {
        if (customerListFilters.status !== "all" && customer.status !== customerListFilters.status) return false;
        if (customerListFilters.productId !== "all" && customer.productId !== customerListFilters.productId) return false;
        if (customerListFilters.city !== "all" && normalizeHeaderName(customer.city) !== normalizeHeaderName(customerListFilters.city)) return false;
        if (!searchValue) return true;

        const haystack = normalizeHeaderName(
          [
            customer.id,
            customer.customerName,
            customer.phone,
            customer.city,
            customer.productName,
            customer.orderDate,
            customer.sourceOrderId,
          ]
            .filter(Boolean)
            .join(" ")
        );

        return haystack.includes(searchValue);
      })
      .sort((a, b) => {
        const dateA = parseDateInput(a.orderDate)?.getTime() || 0;
        const dateB = parseDateInput(b.orderDate)?.getTime() || 0;
        if (dateB !== dateA) return dateB - dateA;
        return String(b.id).localeCompare(String(a.id));
      });
  }, [confirmationStatusMap, customerListFilters.status, customerListFilters.productId, customerListFilters.city, deferredCustomerSearch, operationalCustomers, products]);

  const customerListPageCount = Math.max(1, Math.ceil(compactCustomerRows.length / Number(customerListFilters.pageSize || 25)));
  const selectedCustomerIdSet = useMemo(() => new Set(selectedCustomerIds), [selectedCustomerIds]);
  const filteredCustomerIds = useMemo(() => compactCustomerRows.map((customer) => customer.id), [compactCustomerRows]);
  const allFilteredSelected = filteredCustomerIds.length > 0 && filteredCustomerIds.every((id) => selectedCustomerIdSet.has(id));
  const someFilteredSelected = filteredCustomerIds.some((id) => selectedCustomerIdSet.has(id)) && !allFilteredSelected;
  const historyTargetCustomer = useMemo(
    () => operationalCustomers.find((customer) => customer.id === customerHistoryTargetId) || null,
    [customerHistoryTargetId, operationalCustomers]
  );

  const selectedLead = useMemo(
    () => operationalCustomers.find((c) => c.id === selectedLeadId) || null,
    [selectedLeadId, operationalCustomers]
  );

  const paginatedCustomerRows = useMemo(() => {
    const pageSize = Number(customerListFilters.pageSize || 25);
    const safePage = Math.min(customerListPage, customerListPageCount);
    const startIndex = (safePage - 1) * pageSize;
    return compactCustomerRows.slice(startIndex, startIndex + pageSize);
  }, [compactCustomerRows, customerListFilters.pageSize, customerListPage, customerListPageCount]);

  const filteredCustomerSummary = useMemo(() => {
    return compactCustomerRows.reduce(
      (acc, customer) => {
        acc.totalValue += Number(customer.totalValue || 0);
        if (isConfirmationConfirmed(customer.status)) acc.confirmed += 1;
        else if (isConfirmationCancelled(customer.status)) acc.cancelled += 1;
        else acc.pending += 1;
        return acc;
      },
      { totalValue: 0, confirmed: 0, cancelled: 0, pending: 0 }
    );
  }, [compactCustomerRows]);

  const deferredShippingSearch = useDeferredValue(shippingListFilters.search);

  const compactShippingRows = useMemo(() => {
    const searchValue = normalizeHeaderName(deferredShippingSearch);

    return operationalCustomers
      .map((customer) => {
        const product = products.find((p) => p.id === customer.productId);
        const normalizedStatus = getCustomerShippingStatus(customer) || "to-prepare";
        return {
          ...customer,
          status: normalizedStatus,
          statusLabel: shippingStatusMap[normalizedStatus]?.label || formatStatusLabel(normalizedStatus),
          productName: product?.name || customer.productId,
          totalValue: getCustomerOrderTotalTzs(customer, product),
          lastShippingImportLabel: customer.lastShippingImportedAt
            ? new Date(customer.lastShippingImportedAt).toLocaleString()
            : "Not imported yet",
        };
      })
      .filter((customer) => isConfirmationConfirmed(getCustomerConfirmationStatus(customer)))
      .filter((customer) => {
        if (shippingListFilters.status !== "all" && customer.status !== shippingListFilters.status) return false;
        if (!searchValue) return true;

        const haystack = normalizeHeaderName(
          [
            customer.id,
            customer.customerName,
            customer.phone,
            customer.city,
            customer.productName,
            customer.orderDate,
            customer.sourceOrderId,
            customer.statusLabel,
          ]
            .filter(Boolean)
            .join(" ")
        );

        return haystack.includes(searchValue);
      })
      .sort((a, b) => {
        const dateA = parseDateInput(a.orderDate)?.getTime() || 0;
        const dateB = parseDateInput(b.orderDate)?.getTime() || 0;
        if (dateB !== dateA) return dateB - dateA;
        return String(b.id).localeCompare(String(a.id));
      });
  }, [deferredShippingSearch, operationalCustomers, products, shippingListFilters.status, shippingStatusMap]);

  const shippingListPageCount = Math.max(1, Math.ceil(compactShippingRows.length / Number(shippingListFilters.pageSize || 25)));
  const selectedShippingIdSet = useMemo(() => new Set(selectedShippingIds), [selectedShippingIds]);
  const filteredShippingIds = useMemo(() => compactShippingRows.map((customer) => customer.id), [compactShippingRows]);
  const allFilteredShippingSelected = filteredShippingIds.length > 0 && filteredShippingIds.every((id) => selectedShippingIdSet.has(id));
  const someFilteredShippingSelected = filteredShippingIds.some((id) => selectedShippingIdSet.has(id)) && !allFilteredShippingSelected;

  const paginatedShippingRows = useMemo(() => {
    const pageSize = Number(shippingListFilters.pageSize || 25);
    const safePage = Math.min(shippingListPage, shippingListPageCount);
    const startIndex = (safePage - 1) * pageSize;
    return compactShippingRows.slice(startIndex, startIndex + pageSize);
  }, [compactShippingRows, shippingListFilters.pageSize, shippingListPage, shippingListPageCount]);

  const filteredShippingSummary = useMemo(() => {
    return compactShippingRows.reduce(
      (acc, customer) => {
        acc.totalValue += Number(customer.totalValue || 0);
        if (isShippingDelivered(customer.status)) acc.delivered += 1;
        else if (isShippingReturned(customer.status)) acc.returned += 1;
        else acc.inFlow += 1;
        return acc;
      },
      { totalValue: 0, delivered: 0, returned: 0, inFlow: 0 }
    );
  }, [compactShippingRows]);

  const shippingSummary = useMemo(() => {
    const activeShipping = compactShippingRows.filter((customer) => isShippingInProgress(customer.status)).length;
    const deliveredShipping = compactShippingRows.filter((customer) => isShippingDelivered(customer.status)).length;
    const cancelledShipping = compactShippingRows.filter((customer) => isShippingReturned(customer.status)).length;
    const otherShipping = compactShippingRows.filter(
      (customer) => !isShippingInProgress(customer.status) && !isShippingDelivered(customer.status) && !isShippingReturned(customer.status)
    ).length;

    return {
      total: compactShippingRows.length,
      activeShipping,
      deliveredShipping,
      cancelledShipping,
      otherShipping,
    };
  }, [compactShippingRows]);

  const scalingInsights = useMemo(() => {
    return productDashboard
      .map((product) => {
        const checks = {
          volume: Number(product.orders || 0) >= 5,
          profit: Number(product.profit || 0) > 0,
          confirm: Number(product.confirmRate || 0) >= 0.4,
          delivery: Number(product.deliveryRate || 0) >= 0.5,
          roas: Number(product.roas || 0) >= 1.8,
          stock: Number(product.availableStock || 0) > Math.max(Number(product.reorderPoint || 0), 5),
        };

        const strengths = [];
        const blockers = [];

        if (checks.volume) strengths.push("Lead volume is active");
        else blockers.push("Not enough order volume yet");

        if (checks.profit) strengths.push("Product is profitable");
        else blockers.push("Profit is still negative");

        if (checks.confirm) strengths.push("Confirmation rate is healthy");
        else blockers.push("Confirmation rate is too low");

        if (checks.delivery) strengths.push("Delivery rate is strong");
        else blockers.push("Delivery rate needs work");

        if (checks.roas) strengths.push("ROAS supports scaling");
        else blockers.push("ROAS is still too weak");

        if (checks.stock) strengths.push("Stock can support more volume");
        else blockers.push("Stock is too tight to scale safely");

        const shouldScale = Object.values(checks).every(Boolean);
        const scaleReadiness = Math.round((Object.values(checks).filter(Boolean).length / Object.keys(checks).length) * 100);
        const recommendedAction = shouldScale
          ? Number(product.roas || 0) >= 3
            ? "Scale aggressively"
            : "Scale carefully"
          : scaleReadiness >= 60
            ? "Keep testing"
            : "Fix before scaling";

        return {
          ...product,
          shouldScale,
          scaleReadiness,
          recommendedAction,
          strengths,
          blockers,
        };
      })
      .sort((a, b) => {
        if (Number(b.shouldScale) !== Number(a.shouldScale)) return Number(b.shouldScale) - Number(a.shouldScale);
        if (b.scaleReadiness !== a.scaleReadiness) return b.scaleReadiness - a.scaleReadiness;
        return Number(b.profit || 0) - Number(a.profit || 0);
      });
  }, [productDashboard]);

  const scalingSummary = useMemo(() => {
    const ready = scalingInsights.filter((product) => product.shouldScale);
    const watch = scalingInsights.filter((product) => !product.shouldScale && product.scaleReadiness >= 60);
    const blocked = scalingInsights.filter((product) => product.scaleReadiness < 60);

    return {
      ready,
      watch,
      blocked,
      topCandidate: ready[0] || watch[0] || scalingInsights[0] || null,
    };
  }, [scalingInsights]);

  const situationsSummary = useMemo(() => {
    const salariesTotalTzs = situationData.salaries.reduce((sum, entry) => sum + Number(entry.amountTzs || 0), 0);
    const manualFixedChargesTzs = situationData.fixedCharges.reduce((sum, entry) => sum + Number(entry.amountTzs || 0), 0);
    const purchaseBudgetTzs = products.reduce(
      (sum, product) => sum + (Number(product.purchaseUnitPrice || 0) * Number(product.totalQty || 0) * USD_TO_TZS),
      0
    );
    const importChargesTzs = products.reduce(
      (sum, product) => sum + Number(product.shippingTotal || 0) + Number(product.otherCharges || 0),
      0
    );
    const localDeliveryTzs = liveAutomationSummary.localDeliveryCostTzs || 0;
    const fixedChargesTzs = salariesTotalTzs + manualFixedChargesTzs;

    const productEconomics = products
      .map((product) => {
      const adInput = situationData.adInputs?.[product.id] || {};
      const hasManualAdInput = Object.prototype.hasOwnProperty.call(situationData.adInputs || {}, product.id);
      const averageLeadCostTzs = hasManualAdInput ? Number(adInput.averageLeadCostTzs || 0) : 0;
      const incomingLeads = hasManualAdInput ? Number(adInput.incomingLeads || 0) : 0;
      const revenueTzs = Number(product.sellingPrice || 0) * Number(product.totalQty || 0);
      const stockCostTzs =
        (Number(product.purchaseUnitPrice || 0) * Number(product.totalQty || 0) * USD_TO_TZS) +
        Number(product.shippingTotal || 0) +
        Number(product.otherCharges || 0);
      const productFixedChargeBonusTzs = 8.5 * USD_TO_TZS * Number(product.totalQty || 0);
      const fixedChargesProductTzs = stockCostTzs + productFixedChargeBonusTzs;
      const currentAdsCostTzs = averageLeadCostTzs * incomingLeads;
      const maxAdsCostTzs = Math.max(revenueTzs - fixedChargesProductTzs, 0);
      const adsInputSourceLabel = hasManualAdInput ? "Manual ads input" : "No ads input yet";
      const revenuePercent = revenueTzs > 0 ? 100 : 0;
      const marginOnVariableCostTzs = revenueTzs - currentAdsCostTzs;
      const tmcvPercent = revenueTzs > 0 ? (marginOnVariableCostTzs / revenueTzs) * 100 : 0;
      const fixedChargesPercent = revenueTzs > 0 ? (fixedChargesProductTzs / revenueTzs) * 100 : 0;
      const resultTzs = marginOnVariableCostTzs - fixedChargesProductTzs;
      const resultPercent = revenueTzs > 0 ? (resultTzs / revenueTzs) * 100 : 0;
      const srValueTzs = tmcvPercent > 0 ? fixedChargesProductTzs / (tmcvPercent / 100) : null;
      const effectiveSellingPriceTzs = Number(product.sellingPrice || 0);
      const srVolume = srValueTzs && effectiveSellingPriceTzs > 0 ? srValueTzs / effectiveSellingPriceTzs : null;
      const breakEvenTimeMonths = srValueTzs && revenueTzs > 0 ? (srValueTzs * 12) / revenueTzs : null;

      return {
        ...product,
        sourcedQty: Number(product.totalQty || 0),
        revenueTzs,
        revenuePercent,
        averageLeadCostTzs,
        adsInputSourceLabel,
        leadVolume: incomingLeads,
        adsCostTzs: maxAdsCostTzs,
        currentAdsCostTzs,
        adsUsageRatio: maxAdsCostTzs > 0 ? currentAdsCostTzs / maxAdsCostTzs : 0,
        adsCostPercent: revenueTzs > 0 ? (maxAdsCostTzs / revenueTzs) * 100 : 0,
        marginOnVariableCostTzs,
        tmcvPercent,
        allocatedFixedChargesTzs: fixedChargesProductTzs,
        fixedChargesPercent,
        resultTzs,
        resultPercent,
        srValueTzs,
        srVolume,
        breakEvenTimeMonths,
        effectiveSellingPriceTzs,
      };
    })
      .sort((a, b) => {
        if (Number(b.revenueTzs || 0) !== Number(a.revenueTzs || 0)) return Number(b.revenueTzs || 0) - Number(a.revenueTzs || 0);
        return String(a.name).localeCompare(String(b.name));
      });
    const configuredAdsUsedTzs = productEconomics.reduce((sum, product) => sum + Number(product.currentAdsCostTzs || 0), 0);
    const configuredAverageLeadCostTzs =
      productEconomics.filter((product) => Number(product.averageLeadCostTzs || 0) > 0).length > 0
        ? productEconomics.reduce((sum, product) => sum + Number(product.averageLeadCostTzs || 0), 0) /
          productEconomics.filter((product) => Number(product.averageLeadCostTzs || 0) > 0).length
        : 0;
    const metaTrackedAdsTzs = Number(metaAdsState.cumulativeTrackedSpendTzs || 0);
    const effectiveAdsSpendTzs = metaTrackedAdsTzs > 0 ? metaTrackedAdsTzs : configuredAdsUsedTzs;
    const detectedChargesTzs = purchaseBudgetTzs + importChargesTzs + effectiveAdsSpendTzs + localDeliveryTzs + fixedChargesTzs;

    return {
      salariesTotalTzs,
      manualFixedChargesTzs,
      purchaseBudgetTzs,
      importChargesTzs,
      adSpendTzs: effectiveAdsSpendTzs,
      localDeliveryTzs,
      fixedChargesTzs,
      detectedChargesTzs,
      productEconomics,
      configuredAdsUsedTzs,
      metaTrackedAdsTzs,
      effectiveAdsSpendTzs,
      configuredAverageLeadCostTzs,
    };
  }, [liveAutomationSummary.localDeliveryCostTzs, metaAdsState.cumulativeTrackedSpendTzs, products, situationData]);

  const weeklyProductProfitRows = useMemo(() => {
    const grouped = {};

    operationalCustomers.forEach((customer) => {
      const product = getProduct(customer.productId);
      if (!product) return;

      const weekStart = getWeekStartString(
        isShippingDelivered(getCustomerShippingStatus(customer)) ? customer.lastShippingImportedAt || customer.orderDate : customer.orderDate
      );
      const key = `${product.id}::${weekStart}`;
      if (!grouped[key]) {
        grouped[key] = {
          key,
          productId: product.id,
          productName: product.name,
          weekStart,
          weekLabel: getWeekLabel(weekStart),
          orders: 0,
          confirmed: 0,
          deliveredOrders: 0,
          deliveredUnits: 0,
          returnedOrders: 0,
          revenueTzs: 0,
          adSpendTzs: 0,
          localDeliveryTzs: 0,
          importCostTzs: 0,
          profitTzs: 0,
        };
      }

      grouped[key].orders += 1;
      if (isConfirmationConfirmed(getCustomerConfirmationStatus(customer))) grouped[key].confirmed += 1;
      if (isShippingReturned(getCustomerShippingStatus(customer))) grouped[key].returnedOrders += 1;
      if (isShippingDelivered(getCustomerShippingStatus(customer))) {
        const qty = Math.max(1, Number(customer.quantity || 1));
        grouped[key].deliveredOrders += 1;
        grouped[key].deliveredUnits += qty;
        grouped[key].revenueTzs += getCustomerOrderTotalTzs(customer, product);
        grouped[key].localDeliveryTzs += calculateServiceFeeForOrder(customer, { exchangeRate: USD_TO_TZS }).tsh;
        grouped[key].importCostTzs += getUnitProductCostUSD(product) * USD_TO_TZS * qty;
      }
    });

    const rows = Object.values(grouped).map((row) => {
      const productMetrics = productDashboardMap[row.productId] || {};
      const deliveredUnitsBase = Math.max(1, Number(productMetrics.deliveredUnits || 0));
      const adSpendShare = Number(productMetrics.spend || 0) / deliveredUnitsBase;
      const adSpendTzs = adSpendShare * Number(row.deliveredUnits || 0);
      const profitTzs = row.revenueTzs - row.localDeliveryTzs - row.importCostTzs - adSpendTzs;

      return {
        ...row,
        adSpendTzs,
        profitTzs,
        profitPerDeliveredOrderTzs: row.deliveredOrders > 0 ? profitTzs / row.deliveredOrders : 0,
      };
    });

    return rows.sort((a, b) => {
      const dateGap = String(b.weekStart).localeCompare(String(a.weekStart));
      if (dateGap !== 0) return dateGap;
      if (b.profitTzs !== a.profitTzs) return b.profitTzs - a.profitTzs;
      return a.productName.localeCompare(b.productName);
    });
  }, [getProduct, operationalCustomers, productDashboardMap]);

  const stockForecastRows = useMemo(() => {
    return productDashboard
      .map((product) => {
        const dailyDeliveredUnits = Number(product.salesPerDay || 0);
        const availableStock = Number(product.availableStock || 0);
        const reservedStock = Number(product.reservedStock || 0);
        const daysUntilStockout = dailyDeliveredUnits > 0 ? availableStock / dailyDeliveredUnits : null;
        const reorderDeadlineDays = daysUntilStockout != null ? daysUntilStockout - Number(product.estimatedArrivalDays || 0) : null;
        const projectedStockoutDate =
          daysUntilStockout != null ? addDaysToDateString(getTodayString(), Math.max(0, Math.round(daysUntilStockout))) : "N/A";
        let urgency = "Stable";
        if (daysUntilStockout != null && daysUntilStockout <= 7) urgency = "Critical";
        else if (daysUntilStockout != null && daysUntilStockout <= 14) urgency = "Watch";

        return {
          ...product,
          dailyDeliveredUnits,
          daysUntilStockout,
          reorderDeadlineDays,
          projectedStockoutDate,
          urgency,
          reservedStock,
          availableStock,
        };
      })
      .sort((a, b) => {
        const aDays = a.daysUntilStockout == null ? Number.POSITIVE_INFINITY : a.daysUntilStockout;
        const bDays = b.daysUntilStockout == null ? Number.POSITIVE_INFINITY : b.daysUntilStockout;
        return aDays - bDays;
      });
  }, [productDashboard]);

  const stockAlerts = useMemo(() => {
    const today = getTodayString();
    const alerts = [];
    productDashboard.forEach((product) => {
      const minStock = Math.max(0, Number(product.minStockQuantity ?? situationData?.productAlertThresholds?.minStockQuantity ?? 3));
      const avail = Number(product.availableStock || 0);
      const incoming = Number(product.incomingStock || 0);
      const reserved = Number(product.reservedStock || 0);
      const landedCost = Number(product.unitProductCost || 0);

      if (avail <= 0 && incoming <= 0) {
        alerts.push({ productId: product.id, productName: product.name, type: "out_of_stock", severity: "critical", message: "Out of stock — no incoming stock" });
      } else if (avail <= 0) {
        alerts.push({ productId: product.id, productName: product.name, type: "out_of_stock", severity: "critical", message: `Out of stock — ${incoming} units incoming` });
      } else if (avail <= minStock) {
        alerts.push({ productId: product.id, productName: product.name, type: "low_stock", severity: "warning", message: `Low stock: ${avail} units (threshold: ${minStock})` });
      }
      if (avail > 0 && landedCost <= 0) {
        alerts.push({ productId: product.id, productName: product.name, type: "missing_landed_cost", severity: "info", message: "Available stock but no landed cost configured" });
      }
      if (reserved > 0 && reserved > avail + incoming) {
        alerts.push({ productId: product.id, productName: product.name, type: "oversold", severity: "critical", message: `Oversold: ${reserved} reserved vs ${avail} available` });
      }
    });
    stockPurchases.forEach((purchase) => {
      if (["ordered", "in_transit"].includes(purchase.status) && purchase.expected_arrival_date && purchase.expected_arrival_date < today) {
        const prod = products.find((p) => p.id === purchase.product_id);
        alerts.push({ productId: purchase.product_id, productName: prod?.name || purchase.product_id, type: "delayed_incoming", severity: "warning", message: `Incoming shipment overdue since ${purchase.expected_arrival_date}` });
      }
    });
    return alerts;
  }, [productDashboard, stockPurchases, products, situationData]);

  const stockAuditData = useMemo(() => {
    const totalProducts = products.length;
    let totalAccepted = 0, totalAvailable = 0, totalOutDelivered = 0, totalReserved = 0, totalIncoming = 0, totalDelivered = 0, totalReturned = 0, totalDamaged = 0;
    let missingCostCount = 0, negativeStockCount = 0;
    const perProduct = productDashboard.map((product) => {
      const accepted = Number(product.acceptedStock || 0);
      const avail = Number(product.availableStock || 0);
      const outDelivered = Number(product.reservedStock || 0);
      const reserved = Number(product.reservedStock || 0);
      const incoming = Number(product.incomingStock || 0);
      const delivered = Number(product.deliveredStock || 0);
      const returned = Number(product.returnedStock || 0);
      const damaged = Number(product.damagedStock || 0);
      const cost = Number(product.unitProductCost || 0);
      totalAccepted += accepted;
      totalAvailable += avail;
      totalOutDelivered += outDelivered;
      totalReserved += reserved;
      totalIncoming += incoming;
      totalDelivered += delivered;
      totalReturned += returned;
      totalDamaged += damaged;
      if (avail > 0 && cost <= 0) missingCostCount++;
      if (avail < 0) negativeStockCount++;
      const manualAdj = stockMovements.filter((m) => m.product_id === product.id && m.type === "manual_adjustment").reduce((s, m) => s + Number(m.quantity_change || 0), 0);
      return {
        id: product.id,
        name: product.name,
        accepted,
        available: avail,
        outDelivered,
        reserved,
        incoming,
        delivered,
        returned,
        damaged,
        manualAdjustments: manualAdj,
        landedCostUsd: cost,
        stockValueUsd: Number(product.stockValueUsd || 0),
      };
    });
    const totalStockValueUsd = perProduct.reduce((s, r) => s + r.stockValueUsd, 0);
    return { totalProducts, totalAccepted, totalAvailable, totalOutDelivered, totalReserved, totalIncoming, totalDelivered, totalReturned, totalDamaged, missingCostCount, negativeStockCount, totalStockValueUsd, perProduct };
  }, [productDashboard, stockMovements, products]);

  const taskCenterData = useMemo(() => {
    const tasks = [];

    stockForecastRows.forEach((product) => {
      if (product.urgency === "Critical" || product.urgency === "Watch") {
        tasks.push({
          id: `stock-${product.id}`,
          type: "stock",
          priority: product.urgency === "Critical" ? "High" : "Medium",
          title: `${product.name}: reorder stock`,
          owner: "Stock Team",
          page: "stock",
          detail:
            product.daysUntilStockout != null
              ? `${product.availableStock} units available | stockout in about ${Math.max(1, Math.round(product.daysUntilStockout))} day(s)`
              : "Sales rhythm not detected yet",
        });
      }
    });

    if (shippingImportReminder.isVisible) {
      tasks.push({
        id: "shipping-import-reminder",
        type: "shipping",
        priority: "High",
        title: "Import today shipping Excel",
        owner: "Shipping Team",
        page: "shipping",
        detail: `${shippingImportReminder.confirmedPipelineCount} confirmed order(s) still waiting for a shipping update.`,
      });
    }

    scalingSummary.ready.slice(0, 5).forEach((product) => {
      tasks.push({
        id: `scale-${product.id}`,
        type: "scaling",
        priority: "Medium",
        title: `${product.name}: ready to scale`,
        owner: "Marketing",
        page: "scaling",
        detail: `${product.scaleReadiness}% readiness | ROAS ${Number(product.roas || 0).toFixed(2)} | profit ${formatTZS(product.profit)}`,
      });
    });

    productDashboard
      .filter((product) => Number(product.returnedUnits || 0) >= 3 || Number(product.deliveryRate || 0) < 0.35)
      .slice(0, 6)
      .forEach((product) => {
        tasks.push({
          id: `anomaly-${product.id}`,
          type: "anomaly",
          priority: "Medium",
          title: `${product.name}: anomaly detected`,
          owner: "Operations",
          page: "dashboard",
          detail: `${product.returnedUnits || 0} returned units | ${Math.round((product.deliveryRate || 0) * 100)}% delivery rate`,
        });
      });

    if (metaAdsState.autoSync && metaAdsState.lastSyncSummary && Number(metaAdsState.lastSyncSummary.totalLeads || 0) === 0 && Number(metaAdsState.lastSyncSummary.totalSpendTzs || 0) > 0) {
      tasks.push({
        id: "meta-tracking-gap",
        type: "marketing",
        priority: "Medium",
        title: "Meta tracking gap to review",
        owner: "Marketing",
        page: "tracking",
        detail: "Spend is coming in but tracked leads are zero. Check lead source and campaign tracking.",
      });
    }

    return tasks.sort((a, b) => {
      const priorityWeight = { High: 0, Medium: 1, Low: 2 };
      const gap = (priorityWeight[a.priority] ?? 9) - (priorityWeight[b.priority] ?? 9);
      if (gap !== 0) return gap;
      return a.title.localeCompare(b.title);
    });
  }, [metaAdsState.autoSync, metaAdsState.lastSyncSummary, productDashboard, scalingSummary.ready, shippingImportReminder, stockForecastRows]);

  const homeCockpitSummary = useMemo(() => {
    const productToScale = controlPanelSummary.topWinningProducts[0] || null;
    const productToStopOrFix =
      controlPanelSummary.losingProducts[0] || controlPanelSummary.needsAttentionProducts[0] || null;
    const productToRestock =
      controlPanelSummary.lowStockProducts[0] ||
      stockForecastRows.find((product) => product.urgency === "Critical" || product.urgency === "Watch") ||
      null;
    const biggestProblemToday = taskCenterData[0] || null;

    return {
      productToScale,
      productToStopOrFix,
      productToRestock,
      biggestProblemToday,
      topActions: taskCenterData.slice(0, 5),
      topAlerts: productAlertsSummary.topRows.slice(0, 5),
    };
  }, [controlPanelSummary, productAlertsSummary.topRows, stockForecastRows, taskCenterData]);

  const _decisionBoardColumns = useMemo(() => {
    const columns = {
      "Scale Now": [],
      Restock: [],
      "Stop / Pause": [],
      "Fix Delivery": [],
      Watch: [],
    };

    taskCenterData.forEach((task) => {
      let columnName = "Watch";
      let metric = task.priority;
      let action = task.detail;

      if (task.type === "scaling") {
        columnName = "Scale Now";
        metric = "Scale readiness / ROAS";
        action = "Increase budget carefully and keep tracking delivery quality.";
      } else if (task.type === "stock") {
        columnName = "Restock";
        metric = "Available stock / days until stockout";
        action = "Prepare a new sourcing order before stock pressure blocks sales.";
      } else if (task.type === "shipping" || task.type === "anomaly") {
        columnName = "Fix Delivery";
        metric = "Delivery rate / returned units";
        action = "Review delivery statuses, returns and carrier execution today.";
      } else if (task.type === "marketing" && task.priority === "High") {
        columnName = "Stop / Pause";
        metric = "Ads spend vs delivered";
        action = "Pause or inspect the campaign before more spend is wasted.";
      } else if (task.priority === "High") {
        columnName = "Stop / Pause";
      }

      columns[columnName].push({
        id: task.id,
        product: task.title,
        reason: task.detail,
        metric,
        action,
        page: task.page,
        priority: task.priority,
      });
    });

    return columns;
  }, [taskCenterData]);

  const _aiAssistantAnswer = useMemo(() => {
    const trimmedQuestion = _aiAssistantQuestion.trim();
    if (trimmedQuestion) {
      return [
        `Question received: ${trimmedQuestion}`,
        `Revenue: ${formatTZS(controlPanelSummary.totalRevenueTzs)} | Profit: ${formatTZS(controlPanelSummary.totalProfitTzs)} | Delivered: ${formatInteger(controlPanelSummary.totalDeliveredOrders)}`,
        homeCockpitSummary.biggestProblemToday
          ? `Priority issue: ${homeCockpitSummary.biggestProblemToday.title} — ${homeCockpitSummary.biggestProblemToday.detail}`
          : "No critical blocker is detected right now.",
      ].join("\n\n");
    }

    switch (_aiAssistantPrompt) {
      case "what-happened":
        return [
          `Today the app is tracking ${formatInteger(controlPanelSummary.totalLeads)} leads, ${formatInteger(controlPanelSummary.totalConfirmedOrders)} confirmed orders and ${formatInteger(controlPanelSummary.totalDeliveredOrders)} delivered orders.`,
          `Revenue is ${formatTZS(controlPanelSummary.totalRevenueTzs)} with ${formatTZS(controlPanelSummary.totalProfitTzs)} profit after ads, product and delivery costs.`,
          homeCockpitSummary.biggestProblemToday
            ? `Main blocker: ${homeCockpitSummary.biggestProblemToday.title}.`
            : "There is no major blocker flagged right now.",
        ].join(" ");
      case "what-now":
        return homeCockpitSummary.topActions.length
          ? homeCockpitSummary.topActions
              .map((task, index) => `${index + 1}. ${task.title} — ${task.detail}`)
              .join("\n")
          : "No urgent action is open right now.";
      case "which-scale":
        return homeCockpitSummary.productToScale
          ? `${homeCockpitSummary.productToScale.name} is the best candidate to scale now because profit is ${formatTZS(homeCockpitSummary.productToScale.dashboardProfitTzs || 0)} and margin is ${Number(homeCockpitSummary.productToScale.dashboardProfitMargin || 0).toFixed(1)}%.`
          : "No product is clearly ready to scale yet.";
      case "where-losing":
        return controlPanelSummary.losingProducts.length
          ? controlPanelSummary.losingProducts
              .slice(0, 5)
              .map((row) => `${row.name}: ${formatTZS(row.dashboardProfitTzs || 0)} profit`)
              .join("\n")
          : "No product is currently losing money.";
      case "what-restock":
        return homeCockpitSummary.productToRestock
          ? `${homeCockpitSummary.productToRestock.name} needs restock attention. ${homeCockpitSummary.productToRestock.availableStock || 0} units available.`
          : "No restock pressure is critical right now.";
      default:
        return "Ask a question or choose one of the quick prompts.";
    }
  }, [_aiAssistantPrompt, _aiAssistantQuestion, controlPanelSummary, homeCockpitSummary]);

  const calendarEvents = useMemo(() => {
    const events = [];

    pendingDubaiNotifications.forEach((product) => {
      events.push({
        id: `dubai-${product.id}`,
        date: product.nextArrivalCheckDate || getTodayString(),
        type: "arrival",
        title: `${product.name}: Dubai follow-up`,
        detail: `Check stock arrival for ${product.name}.`,
      });
    });

    if (shippingImportReminder.isVisible) {
      events.push({
        id: "shipping-cutoff",
        date: getTodayString(),
        type: "shipping",
        title: "Shipping import reminder",
        detail: `${shippingImportReminder.confirmedPipelineCount} confirmed order(s) need today shipping import.`,
      });
    }

    stockForecastRows
      .filter((product) => product.daysUntilStockout != null && product.daysUntilStockout <= 21)
      .slice(0, 8)
      .forEach((product) => {
        events.push({
          id: `stockout-${product.id}`,
          date: product.projectedStockoutDate,
          type: "stock",
          title: `${product.name}: projected stockout`,
          detail: `${product.availableStock} units available | ${Math.max(1, Math.round(product.daysUntilStockout || 0))} day(s) left.`,
        });
      });

    return events.sort((a, b) => String(a.date).localeCompare(String(b.date)));
  }, [pendingDubaiNotifications, shippingImportReminder, stockForecastRows]);

  const teamWorkloadRows = useMemo(() => {
    const grouped = {};
    operationalCustomers.forEach((customer) => {
      const owner = String(customer.assignedTo || "").trim();
      if (!owner) return;
      if (!grouped[owner]) {
        grouped[owner] = { owner, total: 0, confirmed: 0, delivered: 0, shipping: 0 };
      }
      grouped[owner].total += 1;
      if (isConfirmationConfirmed(getCustomerConfirmationStatus(customer))) grouped[owner].confirmed += 1;
      if (isShippingDelivered(getCustomerShippingStatus(customer))) grouped[owner].delivered += 1;
      if (isShippingInProgress(getCustomerShippingStatus(customer))) grouped[owner].shipping += 1;
    });
    return Object.values(grouped).sort((a, b) => b.total - a.total);
  }, [operationalCustomers]);

  const executiveSummary = useMemo(() => {
    const today = getTodayString();
    const monthStart = formatDateInput(new Date(new Date().getFullYear(), new Date().getMonth(), 1));
    const todayOrders = operationalCustomers.filter((customer) => customer.orderDate === today);
    const todayRevenueTzs = todayOrders.reduce((sum, customer) => {
      if (!isShippingDelivered(getCustomerShippingStatus(customer))) return sum;
      return sum + getCustomerOrderTotalTzs(customer, getProduct(customer.productId));
    }, 0);
    const monthRevenueTzs = operationalCustomers.reduce((sum, customer) => {
      if (String(customer.orderDate || "") < monthStart) return sum;
      if (!isShippingDelivered(getCustomerShippingStatus(customer))) return sum;
      return sum + getCustomerOrderTotalTzs(customer, getProduct(customer.productId));
    }, 0);
    const openTasks = taskCenterData.length;
    const highPriorityTasks = taskCenterData.filter((task) => task.priority === "High").length;
    const stockImmobilizedTzs = products.reduce(
      (sum, product) => sum + (Number(product.availableStock || 0) * getUnitProductCostUSD(product) * USD_TO_TZS),
      0
    );
    const fixedChargesTzs = Number(situationsSummary.fixedChargesTzs || 0);
    const grossProfitTzs = Number(liveAutomationSummary.grossProfitTzs || 0);
    const estimatedNetAfterFixedTzs = grossProfitTzs - fixedChargesTzs;
    return {
      todayOrders: todayOrders.length,
      todayRevenueTzs,
      monthRevenueTzs,
      openTasks,
      highPriorityTasks,
      stockImmobilizedTzs,
      grossProfitTzs,
      estimatedNetAfterFixedTzs,
    };
  }, [getProduct, liveAutomationSummary.grossProfitTzs, operationalCustomers, products, situationsSummary.fixedChargesTzs, taskCenterData]);

  const cashflowSummary = useMemo(() => {
    const cashInTzs = Number(liveAutomationSummary.totalRevenueTzs || 0);
    const variableOutTzs = Number(liveAutomationSummary.totalOperationalCostTzs || 0);
    const fixedOutTzs = Number(situationsSummary.fixedChargesTzs || 0);
    return {
      cashInTzs,
      variableOutTzs,
      fixedOutTzs,
      netCashTzs: cashInTzs - variableOutTzs - fixedOutTzs,
    };
  }, [liveAutomationSummary.totalOperationalCostTzs, liveAutomationSummary.totalRevenueTzs, situationsSummary.fixedChargesTzs]);


  const profitCenterRows = useMemo(() => {
    const baseRows = productDashboard.map((product) => {
      const adInput = situationData.adInputs?.[product.id] || {};
      const manualAdsUsedTzs = Number(adInput.averageLeadCostTzs || 0) * Number(adInput.incomingLeads || 0);
      // Primary: direct Supabase query (is_mapped=true, summed by product_id)
      const sbData = supabaseAdsSpendByProduct[product.id];
      // Fallback: in-memory cumulative from adsCampaignsData + metaAdsState.savedCampaigns
      const cumulativeData = sbData || cumulativeSpendByProduct[product.id];
      const liveObservedAdsTzs = sbData?.spendTsh > 0
        ? sbData.spendTsh
        : Number(product.spend || product.totalAdsSpend || 0);
      const stockPurchaseTzs = Number(product.totalProductCostTzs || product.totalProductCost || 0);
      const importChargesTzs = Math.max(0, Number(product.totalImportCost || 0) - stockPurchaseTzs);
      const deliveryChargesTzs = Number(product.serviceFeeTsh || product.totalDeliveryCostTzs || 0);
      const productChargesTzs = stockPurchaseTzs + importChargesTzs + deliveryChargesTzs;
      const adsChargesTzs = liveObservedAdsTzs > 0 ? liveObservedAdsTzs : manualAdsUsedTzs;
      const totalChargesTzs = productChargesTzs + adsChargesTzs;
      const balanceTzs = Number(product.revenue || 0) - totalChargesTzs;
      const deliveredCount = Number(product.delivered || product.deliveredUnits || 0);
      const adsSourceLabel = cumulativeData
        ? `Meta ads — ${cumulativeData.campaignCount} campaign(s)`
        : liveObservedAdsTzs > 0 ? "Tracking row spend" : manualAdsUsedTzs > 0 ? "Manual ads input" : "No ads input yet";

      return {
        ...product,
        manualAdsUsedTzs,
        liveObservedAdsTzs,
        stockPurchaseTzs,
        importChargesTzs,
        deliveryChargesTzs,
        productChargesTzs,
        adsChargesTzs,
        totalChargesTzs,
        balanceTzs,
        adsSourceLabel,
        balancePerOrderTzs: deliveredCount > 0 ? balanceTzs / deliveredCount : 0,
        balanceMarginPercent: Number(product.revenue || 0) > 0 ? (balanceTzs / Number(product.revenue || 0)) * 100 : 0,
      };
    });
    return baseRows.sort((a, b) => Number(b.balanceTzs || 0) - Number(a.balanceTzs || 0));
  }, [productDashboard, situationData.adInputs, cumulativeSpendByProduct, supabaseAdsSpendByProduct]);

  const auditRows = useMemo(() => {
    return operationalCustomers
      .flatMap((customer) =>
        (customer.history || []).map((entry) => ({
          ...entry,
          customerId: customer.id,
          customerName: customer.customerName,
          productName: getProduct(customer.productId)?.name || customer.productId,
        }))
      )
      .sort((a, b) => String(b.at || "").localeCompare(String(a.at || "")));
  }, [getProduct, operationalCustomers]);

  const teamScorecardRows = useMemo(() => {
    const salaryLookup = {};
    (situationData.salaries || []).forEach((entry) => {
      const amountTzs = Number(entry.amountTzs || 0);
      const keys = [entry.name, entry.role].map((value) => normalizeHeaderName(value)).filter(Boolean);
      keys.forEach((key) => {
        salaryLookup[key] = Math.max(Number(salaryLookup[key] || 0), amountTzs);
      });
    });

    const grouped = {};
    operationalCustomers.forEach((customer) => {
      const owner = customer.assignedTo || "Unassigned";
      const product = getProduct(customer.productId);
      const shippingStatus = getCustomerShippingStatus(customer);
      const delivered = isShippingDelivered(shippingStatus);
      const confirmed = isConfirmationConfirmed(getCustomerConfirmationStatus(customer));
      const inShipping = isShippingInProgress(shippingStatus) || delivered;
      if (!grouped[owner]) {
        grouped[owner] = {
          owner,
          totalOrders: 0,
          confirmedOrders: 0,
          shippingOrders: 0,
          deliveredOrders: 0,
          revenueTzs: 0,
          profitTzs: 0,
        };
      }

      grouped[owner].totalOrders += 1;
      if (confirmed) grouped[owner].confirmedOrders += 1;
      if (inShipping) grouped[owner].shippingOrders += 1;

      if (delivered) {
        const quantity = Math.max(1, Number(customer.quantity || 1));
        const revenueTzs = getCustomerOrderTotalTzs(customer, product);
        const importCostTzs = getUnitProductCostUSD(product) * USD_TO_TZS * quantity;
        const localDeliveryTzs = calculateServiceFeeForOrder(customer, { exchangeRate: USD_TO_TZS }).tsh;
        grouped[owner].deliveredOrders += 1;
        grouped[owner].revenueTzs += revenueTzs;
        grouped[owner].profitTzs += revenueTzs - importCostTzs - localDeliveryTzs;
      }
    });

    return Object.values(grouped)
      .map((row) => {
        const salaryTzs = Number(salaryLookup[normalizeHeaderName(row.owner)] || 0);
        const confirmationRate = row.totalOrders > 0 ? (row.confirmedOrders / row.totalOrders) * 100 : 0;
        const deliveryRate = row.confirmedOrders > 0 ? (row.deliveredOrders / row.confirmedOrders) * 100 : 0;
        return {
          ...row,
          confirmationRate,
          deliveryRate,
          salaryTzs,
          netAfterSalaryTzs: row.profitTzs - salaryTzs,
        };
      })
      .sort((a, b) => Number(b.revenueTzs || 0) - Number(a.revenueTzs || 0));
  }, [getProduct, operationalCustomers, situationData.salaries]);

  const deferredAuditSearch = useDeferredValue(auditSearch);
  const filteredAuditRows = useMemo(() => {
    const searchValue = normalizeHeaderName(deferredAuditSearch);
    if (!searchValue) return auditRows;
    return auditRows.filter((row) =>
      [
        row.customerName,
        row.customerId,
        row.productName,
        row.action,
        row.source,
        row.details,
        row.from,
        row.to,
      ]
        .map((value) => normalizeHeaderName(value))
        .join(" ")
        .includes(searchValue)
    );
  }, [auditRows, deferredAuditSearch]);

  // Rebuilt profit metrics — single source of truth for the Profit Center.
  const profitOverviewMetrics = useMemo(() => {
    const exchangeRate = Number(serviceForm?.exchangeRate || USD_TO_TZS);

    // Revenue: Excel import → Supabase orders → fallback
    const revenueTsh = revenueImport?.revenueTsh ?? profitOverviewDirect?.revenueTsh ?? profitCenterRows.reduce((s, r) => s + Number(r.revenue || 0), 0);
    const revenueUsd = revenueTsh / exchangeRate;
    const deliveredUnits = revenueImport?.deliveredUnits ?? profitOverviewDirect?.deliveredUnits ?? 0;

    // Stock Charges: received stock purchases only
    const stockChargesUsd = stockPurchases
      .filter((p) => String(p.status || "").toLowerCase() === "received")
      .reduce((sum, p) => sum + Number(p.total_landed_cost_usd || 0), 0);
    const stockChargesTzs = stockChargesUsd * exchangeRate;

    // Service Charges: Excel import → Supabase orders (Dar=8, other=9)
    const serviceChargesUsd = revenueImport?.serviceChargesUsd ?? profitOverviewDirect?.serviceChargesUsd ?? 0;
    const serviceChargesTzs = serviceChargesUsd * exchangeRate;

    // Ads: Meta (from ads_campaigns table) + Manual weekly entries
    const metaAdsUsd = profitOverviewDirect?.adsSpendUsd ?? 0;
    const manualAdsUsd = manualAdsSpend.reduce((s, e) => s + Number(e.amountUsd || 0), 0);
    const totalAdsUsd = metaAdsUsd + manualAdsUsd;
    const totalAdsTzs = totalAdsUsd * exchangeRate;

    // Extra Charges
    const extraChargesUsd = extraCharges.reduce((s, e) => s + Number(e.amountUsd || 0), 0);
    const extraChargesTzs = extraChargesUsd * exchangeRate;

    // Business Profit
    const businessProfitUsd = revenueUsd - stockChargesUsd - serviceChargesUsd - totalAdsUsd - extraChargesUsd;
    const businessProfitTzs = businessProfitUsd * exchangeRate;
    const profitMarginPct = revenueUsd > 0 ? (businessProfitUsd / revenueUsd) * 100 : 0;

    return {
      revenueTsh, revenueUsd, deliveredUnits,
      stockChargesUsd, stockChargesTzs,
      serviceChargesUsd, serviceChargesTzs,
      metaAdsUsd, metaAdsTzs: metaAdsUsd * exchangeRate,
      manualAdsUsd, manualAdsTzs: manualAdsUsd * exchangeRate,
      totalAdsUsd, totalAdsTzs,
      extraChargesUsd, extraChargesTzs,
      businessProfitUsd, businessProfitTzs, profitMarginPct,
      revenueImportedAt: revenueImport?.importedAt || null,
      revenueSource: revenueImport ? "excel" : profitOverviewDirect ? "supabase" : "memory",
      // backward compat aliases used by profitCenterRows
      productChargesUsd: stockChargesUsd, productChargesTzs: stockChargesTzs,
      adsSpendUsd: totalAdsUsd, adsSpendTzs: totalAdsTzs,
      globalExpensesTzs: 0,
    };
  }, [profitOverviewDirect, revenueImport, stockPurchases, serviceForm, manualAdsSpend, extraCharges, profitCenterRows]);

  // Per-product manual ads lookup
  const manualAdsByProduct = useMemo(() => {
    const map = {};
    for (const e of manualAdsSpend) {
      const pid = e.productId || "__unmapped__";
      map[pid] = (map[pid] || 0) + Number(e.amountUsd || 0);
    }
    return map;
  }, [manualAdsSpend]);

  // Per-product profit rows (used in Product Profit tab)
  const productProfitRows = useMemo(() => {
    const exchangeRate = Number(serviceForm?.exchangeRate || USD_TO_TZS);
    return profitCenterRows.map((row) => {
      const revenueUsd = Number(row.revenue || 0) / exchangeRate;
      const stockCostUsd = Number(row.stockPurchaseTzs || 0) / exchangeRate;
      const serviceFeesUsd = Number(row.deliveryChargesTzs || 0) / exchangeRate;
      const metaAdsUsd = Number(row.adsChargesTzs || 0) / exchangeRate;
      const manualAdsForProduct = manualAdsByProduct[row.id] || 0;
      const totalAdsUsd = metaAdsUsd + manualAdsForProduct;
      const profitUsd = revenueUsd - stockCostUsd - serviceFeesUsd - totalAdsUsd;
      const marginPct = revenueUsd > 0 ? (profitUsd / revenueUsd) * 100 : 0;
      return {
        ...row,
        revenueUsd, stockCostUsd, serviceFeesUsd,
        metaAdsUsd, manualAdsUsd: manualAdsForProduct, totalAdsUsd,
        profitUsd, profitTzs: profitUsd * exchangeRate, marginPct,
        status: revenueUsd === 0 ? "no-data" : profitUsd >= 0 ? "positive" : "negative",
      };
    });
  }, [profitCenterRows, manualAdsByProduct, serviceForm]);

  const profitCenterSummary = useMemo(() => {
    const totals = profitCenterRows.reduce(
      (acc, row) => {
        acc.revenueTzs += Number(row.revenue || 0);
        acc.stockPurchaseTzs += Number(row.stockPurchaseTzs || 0);
        acc.importChargesTzs += Number(row.importChargesTzs || 0);
        acc.deliveryChargesTzs += Number(row.deliveryChargesTzs || 0);
        acc.productChargesTzs += Number(row.productChargesTzs || 0);
        acc.adsSpendTzs += Number(row.adsChargesTzs || 0);
        acc.totalChargesTzs += Number(row.totalChargesTzs || 0);
        acc.balanceTzs += Number(row.balanceTzs || 0);
        return acc;
      },
      { revenueTzs: 0, stockPurchaseTzs: 0, importChargesTzs: 0, deliveryChargesTzs: 0, productChargesTzs: 0, adsSpendTzs: 0, totalChargesTzs: 0, balanceTzs: 0 }
    );

    return {
      ...totals,
      adsSourceLabel: profitCenterRows[0]?.adsSourceLabel || "No ads input yet",
      profitableProducts: profitCenterRows.filter((row) => Number(row.balanceTzs || 0) > 0).length,
      topProduct: profitCenterRows[0] || null,
      lastHourlyAdsSnapshot: metaAdsState.dailySpendSnapshots?.[0] || null,
    };
  }, [metaAdsState.dailySpendSnapshots, profitCenterRows]);

  const auditSummary = useMemo(() => {
    return {
      totalEntries: auditRows.length,
      imports: auditRows.filter((row) => String(row.action || "").includes("import")).length,
      manualChanges: auditRows.filter((row) => String(row.source || "").includes("manual")).length,
      latestEntryAt: auditRows[0]?.at || null,
    };
  }, [auditRows]);

  const _addSituationSalary = () => {
    setSituationData((prev) => ({
      ...prev,
      salaries: [...prev.salaries, { id: `salary-${Date.now()}`, name: "", role: "", amountTzs: 0 }],
    }));
  };

  const _updateSituationSalary = (salaryId, field, value) => {
    setSituationData((prev) => ({
      ...prev,
      salaries: prev.salaries.map((entry) =>
        entry.id === salaryId
          ? { ...entry, [field]: field === "amountTzs" ? Math.max(0, parseLooseNumber(value) * USD_TO_TZS) : value }
          : entry
      ),
    }));
  };

  const _removeSituationSalary = (salaryId) => {
    setSituationData((prev) => ({
      ...prev,
      salaries: prev.salaries.filter((entry) => entry.id !== salaryId),
    }));
  };

  const _addSituationFixedCharge = () => {
    setSituationData((prev) => ({
      ...prev,
      fixedCharges: [...prev.fixedCharges, { id: `fixed-${Date.now()}`, label: "", amountTzs: 0 }],
    }));
  };

  const _updateSituationFixedCharge = (chargeId, field, value) => {
    setSituationData((prev) => ({
      ...prev,
      fixedCharges: prev.fixedCharges.map((entry) =>
        entry.id === chargeId
          ? { ...entry, [field]: field === "amountTzs" ? Math.max(0, parseLooseNumber(value) * USD_TO_TZS) : value }
          : entry
      ),
    }));
  };

  const _removeSituationFixedCharge = (chargeId) => {
    setSituationData((prev) => ({
      ...prev,
      fixedCharges: prev.fixedCharges.filter((entry) => entry.id !== chargeId),
    }));
  };

  const updateSituationAdInput = (productId, field, value) => {
    setAdInputDrafts((prev) => ({
      ...prev,
      [productId]: {
        ...prev[productId],
        [field]: value,
      },
    }));

    setSituationData((prev) => {
      const current = prev.adInputs?.[productId] || { averageLeadCostTzs: 0, incomingLeads: 0 };
      const nextEntry = {
        ...current,
        [field]:
          field === "averageLeadCostTzs"
            ? Math.max(0, parseLooseNumber(value) * USD_TO_TZS)
            : Math.max(0, Math.round(parseLooseNumber(value))),
      };

      return {
        ...prev,
        adInputs: {
          ...prev.adInputs,
          [productId]: nextEntry,
        },
      };
    });
  };

  const getSituationAdInputDisplayValue = (productId, field, fallbackValue) => {
    const draftValue = adInputDrafts?.[productId]?.[field];
    if (draftValue != null) return draftValue;
    return fallbackValue;
  };

  useEffect(() => {
    setCustomerListPage(1);
  }, [customerListFilters.search, customerListFilters.status, customerListFilters.pageSize]);

  useEffect(() => {
    setCustomerListPage((prev) => Math.min(prev, customerListPageCount));
  }, [customerListPageCount]);

  useEffect(() => {
    setSelectedCustomerIds((prev) => prev.filter((id) => operationalCustomers.some((customer) => customer.id === id)));
  }, [operationalCustomers]);

  useEffect(() => {
    if (selectAllCustomersRef.current) {
      selectAllCustomersRef.current.indeterminate = someFilteredSelected;
    }
  }, [someFilteredSelected]);

  useEffect(() => {
    setShippingListPage(1);
  }, [shippingListFilters.search, shippingListFilters.status, shippingListFilters.pageSize]);

  useEffect(() => {
    setShippingListPage((prev) => Math.min(prev, shippingListPageCount));
  }, [shippingListPageCount]);

  useEffect(() => {
    setSelectedShippingIds((prev) => prev.filter((id) => operationalCustomers.some((customer) => customer.id === id)));
  }, [operationalCustomers]);

  useEffect(() => {
    if (selectAllShippingRef.current) {
      selectAllShippingRef.current.indeterminate = someFilteredShippingSelected;
    }
  }, [someFilteredShippingSelected]);

  const ordersChartData = useMemo(() => {
    const grouped = operationalCustomers.reduce((acc, customer) => {
      const dateKey = customer.orderDate || "No date";
      const confirmationStatus = getCustomerConfirmationStatus(customer);
      const shippingStatus = getCustomerShippingStatus(customer) || getCustomerEffectiveStatus(customer);
      if (!acc[dateKey]) {
        acc[dateKey] = {
          date: dateKey,
          incoming: 0,
          confirmed: 0,
          delivered: 0,
        };
      }

      acc[dateKey].incoming += 1;
      if (isConfirmationConfirmed(confirmationStatus)) {
        acc[dateKey].confirmed += 1;
      }
      if (isShippingDelivered(shippingStatus)) {
        acc[dateKey].delivered += 1;
      }

      return acc;
    }, {});

    return Object.values(grouped)
      .sort((a, b) => String(a.date).localeCompare(String(b.date)))
      .slice(-10);
  }, [operationalCustomers]);

  const filteredCustomersForOverview = useMemo(() => {
    return serviceLeadCustomers.filter((customer) => {
      const matchesProduct = overviewFilters.productId === "all" || customer.productId === overviewFilters.productId;

      let matchesDate = true;
      if (overviewFilters.periodMode === "custom") {
        if (overviewFilters.startDate && customer.orderDate < overviewFilters.startDate) matchesDate = false;
        if (overviewFilters.endDate && customer.orderDate > overviewFilters.endDate) matchesDate = false;
      }

      return matchesProduct && matchesDate;
    });
  }, [overviewFilters, serviceLeadCustomers]);

  const overviewSummary = useMemo(() => {
    const incoming = filteredCustomersForOverview.length;
    const newCount = filteredCustomersForOverview.filter((c) => isConfirmationNew(getCustomerConfirmationStatus(c))).length;
    const pending = filteredCustomersForOverview.filter((c) => getConfirmationBucket(getCustomerConfirmationStatus(c)) === "pending").length;
    const awaitingDelivery = filteredCustomersForOverview.filter((c) => {
      const confirmationStatus = getCustomerConfirmationStatus(c);
      const shippingStatus = getCustomerShippingStatus(c) || "to-prepare";
      return isConfirmationConfirmed(confirmationStatus) && isShippingInProgress(shippingStatus);
    }).length;
    const delivered = filteredCustomersForOverview.filter((c) => isShippingDelivered(getCustomerShippingStatus(c))).length;
    const cancelled = filteredCustomersForOverview.filter(
      (c) => isConfirmationCancelled(getCustomerConfirmationStatus(c)) || isShippingReturned(getCustomerShippingStatus(c))
    ).length;
    const confirmed = filteredCustomersForOverview.filter((c) => isConfirmationConfirmed(getCustomerConfirmationStatus(c))).length;

    const revenue = filteredCustomersForOverview.reduce((sum, customer) => {
      if (!isShippingDelivered(getCustomerShippingStatus(customer))) return sum;
      const product = products.find((p) => p.id === customer.productId);
      return sum + getCustomerOrderTotalTzs(customer, product);
    }, 0);

    const statusBreakdown = [
      { label: "New Order", count: newCount, color: getStatusColor("new-order") },
      { label: "Pending", count: pending, color: getStatusColor("pending") },
      { label: "Confirmed", count: awaitingDelivery, color: getStatusColor("confirmed") },
      { label: "Delivered", count: delivered, color: getStatusColor("delivered") },
      { label: "Cancelled / Returned", count: cancelled, color: getStatusColor("cancelled") },
    ]
      .filter((status) => status.count > 0)
      .map((status) => ({
        ...status,
        pct: incoming > 0 ? Number(((status.count / incoming) * 100).toFixed(1)) : 0,
      }));

    return {
      incoming,
      newCount,
      pending,
      confirmed,
      awaitingDelivery,
      delivered,
      cancelled,
      revenue,
      statusBreakdown,
    };
  }, [filteredCustomersForOverview, products]);

  const overviewPieData = useMemo(() => {
    const visibleStatuses = overviewSummary.statusBreakdown.slice(0, 6).map((status) => ({
      name: status.label,
      value: status.pct,
      color: status.color,
    }));
    const hiddenStatuses = overviewSummary.statusBreakdown.slice(6);

    if (hiddenStatuses.length > 0) {
      visibleStatuses.push({
        name: `Other (${hiddenStatuses.length})`,
        value: Number(hiddenStatuses.reduce((sum, status) => sum + status.pct, 0).toFixed(1)),
        color: "#7c3aed",
      });
    }

    return visibleStatuses;
  }, [overviewSummary]);

  const confirmationDetails = useMemo(() => {
    const total = customersDashboard.totalOrders || 0;
    const items = confirmationStatusCatalog
      .filter((status) => status.count > 0)
      .slice(0, 8)
      .map((status) => ({
        label: status.label,
        count: status.count,
        color: status.color,
      }));

    return {
      total,
      items: items.map((item) => ({
        ...item,
        pct: total > 0 ? Math.round((item.count / total) * 100) : 0,
      })),
    };
  }, [confirmationStatusCatalog, customersDashboard.totalOrders]);

  const deliveryDetails = useMemo(() => {
    const confirmedCustomers = serviceLeadCustomers.filter((c) =>
      isConfirmationConfirmed(getCustomerConfirmationStatus(c))
    );
    const confirmedBase = confirmedCustomers.length;
    const counts = confirmedCustomers.reduce((acc, customer) => {
      const statusKey = normalizeOrderStatus(getCustomerShippingStatus(customer) || "to-prepare");
      acc[statusKey] = (acc[statusKey] || 0) + 1;
      return acc;
    }, {});

    const items = Object.entries(counts)
      .map(([status, count]) => ({
        label: shippingStatusMap[status]?.label || formatStatusLabel(status),
        count,
        color: shippingStatusMap[status]?.color || getStatusColor(status),
      }))
      .sort((a, b) => {
        if (b.count !== a.count) return b.count - a.count;
        return a.label.localeCompare(b.label);
      })
      .slice(0, 8);

    return {
      total: confirmedBase,
      items: items.map((item) => ({
        ...item,
        pct: confirmedBase > 0 ? Math.round((item.count / confirmedBase) * 100) : 0,
      })),
    };
  }, [serviceLeadCustomers, shippingStatusMap]);

  const getPeriodStartDate = (period) => {
    const today = new Date();
    const current = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    if (period === "today") return current;
    if (period === "yesterday") {
      const d = new Date(current);
      d.setDate(d.getDate() - 1);
      return d;
    }
    if (period === "thisWeek") {
      const d = new Date(current);
      const day = d.getDay() || 7;
      d.setDate(d.getDate() - (day - 1));
      return d;
    }
    if (period === "thisMonth") {
      return new Date(current.getFullYear(), current.getMonth(), 1);
    }
    if (period === "last7Days") {
      const d = new Date(current);
      d.setDate(d.getDate() - 6);
      return d;
    }
    return null;
  };

  const filteredCustomersForConfirmationSummary = useMemo(() => {
    const startDatePreset = getPeriodStartDate(confirmationSummaryFilters.period);

    return serviceLeadCustomers.filter((customer) => {
      const matchesProduct =
        confirmationSummaryFilters.productId === "all" ||
        customer.productId === confirmationSummaryFilters.productId;

      if (!matchesProduct) return false;

      const orderDate = parseDateInput(customer.orderDate);
      if (!orderDate) return false;
      if (Number.isNaN(orderDate.getTime())) return false;

      // CUSTOM DATE FILTER
      if (confirmationSummaryFilters.period === "custom") {
        if (
          confirmationSummaryFilters.startDate &&
          orderDate < parseDateInput(confirmationSummaryFilters.startDate)
        )
          return false;
        if (
          confirmationSummaryFilters.endDate &&
          orderDate > parseDateInput(confirmationSummaryFilters.endDate)
        )
          return false;
        return true;
      }

      // PRESET FILTER
      if (!startDatePreset) return true;
      return orderDate >= startDatePreset;
    });
  }, [confirmationSummaryFilters, serviceLeadCustomers]);

  const confirmationSummary = useMemo(() => {
    const totalLeads = filteredCustomersForConfirmationSummary.length;
    const confirmed = filteredCustomersForConfirmationSummary.filter((c) => isConfirmationConfirmed(getCustomerConfirmationStatus(c))).length;
    const cancelled = filteredCustomersForConfirmationSummary.filter((c) => isConfirmationCancelled(getCustomerConfirmationStatus(c))).length;
    const newOrder = filteredCustomersForConfirmationSummary.filter((c) => isConfirmationNew(getCustomerConfirmationStatus(c))).length;
    const pending = totalLeads - confirmed - cancelled - newOrder;
    const confirmationRate = totalLeads > 0 ? (confirmed / totalLeads) * 100 : 0;

    const grouped = filteredCustomersForConfirmationSummary.reduce((acc, customer) => {
      const key = customer.orderDate || "No date";
      if (!acc[key]) {
        acc[key] = { date: key, cancelled: 0, confirmed: 0, newOrder: 0, pending: 0 };
      }
      if (isConfirmationCancelled(getCustomerConfirmationStatus(customer))) acc[key].cancelled += 1;
      else if (isConfirmationConfirmed(getCustomerConfirmationStatus(customer))) acc[key].confirmed += 1;
      else if (isConfirmationNew(getCustomerConfirmationStatus(customer))) acc[key].newOrder += 1;
      else acc[key].pending += 1;
      return acc;
    }, {});

    const chartData = Object.values(grouped).sort((a, b) => String(a.date).localeCompare(String(b.date))).slice(-10);

    const breakdown = [
      { label: "Cancelled", count: cancelled, color: "#ef4444" },
      { label: "Confirmed", count: confirmed, color: "#84cc16" },
      { label: "New Order", count: newOrder, color: "#6366f1" },
      { label: "Pending", count: Math.max(0, pending), color: "#67e8f9" },
    ].map((item) => ({
      ...item,
      pct: totalLeads > 0 ? Math.round((item.count / totalLeads) * 100) : 0,
    }));

    return {
      totalLeads,
      confirmed,
      confirmationRate,
      chartData,
      breakdown,
    };
  }, [filteredCustomersForConfirmationSummary]);

  const filteredCustomersForProductDetails = useMemo(() => {
    const startDatePreset = getPeriodStartDate(productDetailsFilters.period);

    return serviceLeadCustomers.filter((customer) => {
      const matchesProduct = productDetailsFilters.productId === "all" || customer.productId === productDetailsFilters.productId;
      if (!matchesProduct) return false;

      if (productDetailsFilters.period === "all") return true;

      const orderDate = parseDateInput(customer.orderDate);
      if (!orderDate || Number.isNaN(orderDate.getTime())) return false;

      if (productDetailsFilters.period === "custom") {
        if (productDetailsFilters.startDate && orderDate < parseDateInput(productDetailsFilters.startDate)) return false;
        if (productDetailsFilters.endDate && orderDate > parseDateInput(productDetailsFilters.endDate)) return false;
        return true;
      }

      if (!startDatePreset) return true;
      return orderDate >= startDatePreset;
    });
  }, [productDetailsFilters, serviceLeadCustomers]);

  const productDetailsRows = useMemo(() => {
    return products
      .filter((product) => productDetailsFilters.productId === "all" || product.id === productDetailsFilters.productId)
      .map((product) => {
        const productOrders = filteredCustomersForProductDetails.filter((customer) => customer.productId === product.id);
        const leads = productOrders.length;
        const confirmedOrders = productOrders.filter((customer) => isConfirmationConfirmed(getCustomerConfirmationStatus(customer))).length;
        const deliveredOrders = productOrders.filter((customer) => isShippingDelivered(getCustomerShippingStatus(customer))).length;
        const totalRevenue = productOrders
          .filter((customer) => isShippingDelivered(getCustomerShippingStatus(customer)))
          .reduce((sum, customer) => sum + getCustomerOrderTotalTzs(customer, product), 0);
        const totalDeliveredUnits = productOrders
          .filter((customer) => isShippingDelivered(getCustomerShippingStatus(customer)))
          .reduce((sum, customer) => sum + Number(customer.quantity || 0), 0);
        const confirmationRate = leads > 0 ? (confirmedOrders / leads) * 100 : 0;
        const deliveryRate = confirmedOrders > 0 ? (deliveredOrders / confirmedOrders) * 100 : 0;
        const leadToDeliveryRate = leads > 0 ? (deliveredOrders / leads) * 100 : 0;
        const aov = deliveredOrders > 0 ? totalRevenue / deliveredOrders : 0;
        const nameParts = String(product.name || "").trim().split(/\s+/).filter(Boolean);
        const initials = (nameParts[0]?.[0] || "") + (nameParts[1]?.[0] || nameParts[0]?.[1] || "");

        return {
          id: product.id,
          name: product.name,
          source: product.source,
          initials: initials.toUpperCase() || "PR",
          leads,
          confirmedOrders,
          deliveredOrders,
          totalDeliveredUnits,
          confirmationRate,
          deliveryRate,
          leadToDeliveryRate,
          totalRevenue,
          aov,
        };
      })
      .sort((a, b) => {
        if (b.leads !== a.leads) return b.leads - a.leads;
        if (b.totalRevenue !== a.totalRevenue) return b.totalRevenue - a.totalRevenue;
        return a.name.localeCompare(b.name);
      });
  }, [filteredCustomersForProductDetails, productDetailsFilters.productId, products]);

  const visibleProductDetailsRows = useMemo(
    () => productDetailsRows.slice(0, Number(productDetailsFilters.rowLimit || 10)),
    [productDetailsRows, productDetailsFilters.rowLimit]
  );

  const showCloudLoginGate = supabaseEnabled && cloudAuth.ready && !cloudAuth.user;
  const showWorkspaceSyncNotice =
    Boolean(sharedWorkspace.notice) &&
    /(failed|unavailable|offline|delayed|error)/i.test(sharedWorkspace.notice);
  const showCloudAuthNotice =
    Boolean(cloudAuth.notice) &&
    /(failed|error|unable|invalid|denied|required)/i.test(cloudAuth.notice);

  return (
    <div style={styles.shell}>
      <div
        style={{
          ...styles.layout,
          gridTemplateColumns: isCompact ? "1fr" : "260px 1fr",
          filter: showCloudLoginGate ? "blur(10px)" : "none",
          pointerEvents: showCloudLoginGate ? "none" : "auto",
          userSelect: showCloudLoginGate ? "none" : "auto",
          transition: "filter 160ms ease",
        }}
      >
        <aside style={{ ...styles.sidebar, borderRight: isCompact ? "none" : `1px solid ${cardBorder}`, borderBottom: isCompact ? `1px solid ${cardBorder}` : "none" }}>
          <div style={{ ...styles.brandPanel, marginBottom: 28 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
              <div style={styles.brandMark}>
                <TrendingUp size={20} />
              </div>
              <div>
                <div style={{ fontSize: 12, color: accent, fontWeight: 800, letterSpacing: 0.6, textTransform: "uppercase" }}>Tanzania OS</div>
                <div style={{ fontSize: 24, fontWeight: 900, lineHeight: 1.05 }}>Ecom Tracker</div>
              </div>
            </div>
            <div style={{ marginTop: 14, color: textSoft, lineHeight: 1.55 }}>
              Premium control tower for products, leads, stock flow and delivery performance.
            </div>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginTop: 16 }}>
              <span style={{ ...styles.badge, background: "rgba(29,95,208,0.08)", color: accent, border: "1px solid rgba(29,95,208,0.12)" }}>Live operations</span>
              <span style={{ ...styles.badge, background: "rgba(31,143,95,0.08)", color: green, border: "1px solid rgba(31,143,95,0.12)" }}>{products.length} products</span>
            </div>
          </div>

          <SidebarItem active={activePage === "executive"} onClick={() => setActivePage("executive")} icon={<BarChart3 size={18} />} label="Home" />
          <SidebarItem active={activePage === "customersOrders"} onClick={() => { setActivePage("customersOrders"); setOrdersTab("pipeline"); }} icon={<Users size={18} />} label="Orders" />
          <SidebarItem active={activePage === "shipping"} onClick={() => { setActivePage("shipping"); setShippingTab("queue"); }} icon={<ShoppingBag size={18} />} label="Shipping" />
          <SidebarItem active={["products", "stock", "multiDashboard"].includes(activePage)} onClick={() => { setActivePage("products"); setStockTab("overview"); }} icon={<Archive size={18} />} label="Products & Stock" />
          <SidebarItem active={activePage === "tracking"} onClick={() => setActivePage("tracking")} icon={<Calculator size={18} />} label="Ads & Tracking" />
          <SidebarItem active={["serviceSum", "situations", "profitCenter"].includes(activePage)} onClick={() => { setActivePage("profitCenter"); setProfitTab("overview"); }} icon={<Wallet size={18} />} label="Profit" />
          <SidebarItem active={["taskCenter", "calendar", "team", "alerts"].includes(activePage)} onClick={() => setActivePage("taskCenter")} icon={<ClipboardList size={18} />} label="Decisions" />
          <SidebarItem active={activePage === "aiAssistant"} onClick={() => setActivePage("aiAssistant")} icon={<MessageSquare size={18} />} label="AI Assistant" />
          <SidebarItem active={["settingsAudit", "audit"].includes(activePage)} onClick={() => { setActivePage("settingsAudit"); setSettingsAuditTab("workspace"); }} icon={<Settings size={18} />} label="Settings & Audit" />

        </aside>

        <main style={{ ...styles.main, padding: isCompact ? 18 : 24 }}>
          <input ref={ordersImportInputRef} type="file" accept=".xlsx,.xls,.csv" onChange={importOrdersFromExcel} style={{ display: "none" }} />
          <input ref={restoreJsonInputRef} type="file" accept=".json,application/json" onChange={restoreAppDataFromJson} style={{ display: "none" }} />
          <div style={{ width: "100%", minWidth: 0, display: "grid", gap: 18 }}>
          {activePage === "executive" ? (
          <div style={styles.topbar}>
            <div style={{ ...styles.heroGrid, gridTemplateColumns: responsiveColumns("minmax(0, 1.2fr) minmax(320px, 0.8fr)", "1fr", "1fr") }}>
              <div>
                <div style={styles.sectionEyebrow}>Operations cockpit</div>
                <div style={{ fontSize: isCompact ? 28 : 36, fontWeight: 900, marginTop: 8, lineHeight: 1.02, maxWidth: 680 }}>
                  Tanzania Ecom Tracker
                </div>
                <div style={{ color: textSoft, marginTop: 10, maxWidth: 620, lineHeight: 1.65 }}>
                  A sharper command center for ecommerce execution, with product performance, lead quality, stock flow and delivery health in one premium workspace.
                </div>
                <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginTop: 16 }}>
                  <span style={{ ...styles.badge, background: "rgba(29,95,208,0.08)", color: accent, border: "1px solid rgba(29,95,208,0.12)" }}>Multi-view analytics</span>
                  <span style={{ ...styles.badge, background: "rgba(31,143,95,0.08)", color: green, border: "1px solid rgba(31,143,95,0.12)" }}>Live stock tracking</span>
                  <span
                    style={{
                      ...styles.badge,
                      background: sharedWorkspace.available ? "rgba(31,143,95,0.08)" : "rgba(199,131,34,0.12)",
                      color: sharedWorkspace.available ? green : amber,
                      border: sharedWorkspace.available ? "1px solid rgba(31,143,95,0.12)" : "1px solid rgba(199,131,34,0.18)",
                    }}
                  >
                    {sharedWorkspace.available ? (supabaseEnabled ? "Cloud workspace live" : "Shared workspace live") : (supabaseEnabled ? "Cloud workspace offline" : "Local workspace")}
                  </span>
                </div>
                <div style={{ color: textSoft, marginTop: 10, fontSize: 13 }}>
                  Auto-save: {lastAutoBackupAt ? `Last saved ${new Date(lastAutoBackupAt).toLocaleString()}` : "Changes save automatically as you work"}
                </div>
                <div style={{ color: textSoft, marginTop: 6, fontSize: 13 }}>
                  Last shipping import: {shippingImportReminder.lastShippingImportLabel}
                </div>
                {showWorkspaceSyncNotice ? (
                  <div style={{ color: textSoft, marginTop: 6, fontSize: 13 }}>
                    Workspace sync: {sharedWorkspace.notice}{sharedWorkspace.updatedAt ? ` | ${new Date(sharedWorkspace.updatedAt).toLocaleString()}` : ""}
                  </div>
                ) : null}
                {showExecutiveAdminTools ? (
                  <div style={{ ...styles.softStat, marginTop: 16, background: "linear-gradient(180deg, rgba(255,255,255,0.98), rgba(244,248,255,0.9))" }}>
                    <div style={{ fontSize: 11, fontWeight: 800, letterSpacing: 0.45, textTransform: "uppercase", color: accent }}>Cloud access</div>
                    <div style={{ marginTop: 8, fontWeight: 800, fontSize: 18 }}>
                      {cloudAuth.user ? cloudAuth.user.email || "Authenticated user" : "Sign in to share the live app"}
                    </div>
                    {showCloudAuthNotice ? (
                      <div style={{ color: textSoft, marginTop: 6, lineHeight: 1.5 }}>{cloudAuth.notice}</div>
                    ) : null}
                    {!cloudAuth.user ? (
                      <div style={{ display: "grid", gap: 10, marginTop: 14 }}>
                        <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("1fr 1fr 120px", "1fr 1fr", "1fr"), gap: 10 }}>
                          <input
                            style={styles.input}
                            type="email"
                            placeholder="Email"
                            value={cloudAuth.email}
                            onChange={(e) => setCloudAuth((prev) => ({ ...prev, email: e.target.value }))}
                          />
                          <input
                            style={styles.input}
                            type="password"
                            placeholder="Password"
                            value={cloudAuth.password}
                            onChange={(e) => setCloudAuth((prev) => ({ ...prev, password: e.target.value }))}
                          />
                          <select
                            style={styles.input}
                            value={cloudAuth.mode}
                            onChange={(e) => setCloudAuth((prev) => ({ ...prev, mode: e.target.value }))}
                          >
                            <option value="signin">Sign in</option>
                            <option value="signup">Create access</option>
                          </select>
                        </div>
                        <button style={styles.btnPrimary} onClick={submitCloudAuth} disabled={cloudAuth.loading}>
                          {cloudAuth.loading ? "Connecting..." : cloudAuth.mode === "signup" ? "Create cloud access" : "Open cloud workspace"}
                        </button>
                      </div>
                    ) : (
                      <div style={{ display: "grid", gap: 12, marginTop: 14 }}>
                        <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("1fr 160px", "1fr", "1fr"), gap: 10 }}>
                          <div style={{ ...styles.badge, justifyContent: "flex-start", padding: "12px 14px", height: "100%", background: "rgba(31,143,95,0.1)", color: green, border: "1px solid rgba(31,143,95,0.16)" }}>
                            Workspace: {supabaseWorkspaceId}
                          </div>
                          <button style={styles.btnSecondary} onClick={logoutCloudAuth}>
                            Sign out
                          </button>
                        </div>
                        <div
                          style={{
                            borderRadius: 14,
                            border: `1px solid ${cardBorder}`,
                            background: "rgba(248,250,255,0.76)",
                            padding: "10px 12px",
                          }}
                        >
                          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, flexWrap: "wrap" }}>
                            <div style={{ minWidth: 0, display: "grid", gap: 2 }}>
                              <div style={{ fontSize: 11, fontWeight: 800, letterSpacing: 0.42, textTransform: "uppercase", color: accent }}>
                                Restore points
                              </div>
                              <div style={{ color: textSoft, fontSize: 13, lineHeight: 1.45 }}>
                                {cloudBackupState.available
                                  ? `${formatInteger(cloudBackupState.items.length)} backup${cloudBackupState.items.length > 1 ? "s" : ""}${
                                      cloudBackupState.items[0]?.created_at ? ` | latest ${new Date(cloudBackupState.items[0].created_at).toLocaleString()}` : ""
                                    }`
                                  : "Restore history not enabled yet"}
                              </div>
                            </div>
                            <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
                              <button
                                style={{ ...styles.btnSecondary, padding: "10px 12px", minHeight: 0, borderRadius: 12, fontSize: 13 }}
                                onClick={() => void refreshCloudBackups()}
                                disabled={cloudBackupState.loading || Boolean(cloudBackupState.restoringId)}
                              >
                                {cloudBackupState.loading ? "Loading..." : "Refresh"}
                              </button>
                              {cloudBackupState.available && cloudBackupState.items.length ? (
                                <button
                                  style={{ ...styles.btnSecondary, padding: "10px 12px", minHeight: 0, borderRadius: 12, fontSize: 13 }}
                                  onClick={() => setCloudBackupOpen((prev) => !prev)}
                                  disabled={Boolean(cloudBackupState.restoringId)}
                                >
                                  {cloudBackupOpen ? "Hide" : "Show"}
                                </button>
                              ) : null}
                            </div>
                          </div>
                          {cloudBackupState.notice ? (
                            <div style={{ color: cloudBackupState.available ? textSoft : amber, marginTop: 8, fontSize: 12, lineHeight: 1.45 }}>
                              {cloudBackupState.notice}
                            </div>
                          ) : null}
                          {cloudBackupOpen && cloudBackupState.available && cloudBackupState.items.length ? (
                            <div style={{ display: "grid", gap: 8, marginTop: 10 }}>
                              {cloudBackupState.items.slice(0, 5).map((backup) => {
                                const summary = backup.summary || {};
                                return (
                                  <div
                                    key={backup.id}
                                    style={{
                                      display: "grid",
                                      gridTemplateColumns: responsiveColumns("minmax(0, 1fr) 112px", "1fr 110px", "1fr"),
                                      gap: 8,
                                      alignItems: "center",
                                      padding: "8px 10px",
                                      borderRadius: 12,
                                      border: `1px solid ${cardBorder}`,
                                      background: "rgba(255,255,255,0.82)",
                                    }}
                                  >
                                    <div style={{ minWidth: 0 }}>
                                      <div style={{ fontWeight: 800, fontSize: 13 }}>
                                        {backup.created_at ? new Date(backup.created_at).toLocaleString() : `Backup #${backup.id}`}
                                      </div>
                                      <div style={{ color: textSoft, marginTop: 2, fontSize: 12, lineHeight: 1.4 }}>
                                        V{formatInteger(backup.workspace_version || 0)} | {formatInteger(summary.products || 0)} p | {formatInteger(summary.customers || 0)} cmd | {formatInteger(summary.tracking || 0)} tr
                                      </div>
                                    </div>
                                    <button
                                      style={{ ...styles.btnSecondary, padding: "9px 12px", minHeight: 0, borderRadius: 12, fontSize: 13 }}
                                      onClick={() => void restoreCloudBackup(backup.id)}
                                      disabled={cloudBackupState.loading || cloudBackupState.restoringId === backup.id}
                                    >
                                      {cloudBackupState.restoringId === backup.id ? "..." : "Restore"}
                                    </button>
                                  </div>
                                );
                              })}
                            </div>
                          ) : null}
                        </div>
                      </div>
                    )}
                  </div>
                ) : null}
                <div style={{ ...styles.topbarActions, marginTop: 18, display: "none" }}>
                  <button style={styles.btnPrimary} onClick={exportReport}>Export Report</button>
                  <button style={styles.btnSecondary} onClick={exportAllDataToCsv}>Export All CSV</button>
                  <button style={styles.btnSecondary} onClick={exportProductPerformanceToCsv}>Export Product CSV</button>
                  <button style={styles.btnSecondary} onClick={backupAllAppDataToJson}>Backup JSON</button>
                  <button style={styles.btnSecondary} onClick={() => restoreJsonInputRef.current?.click()}>Restore JSON</button>
                </div>
              </div>

              <div style={styles.heroAside}>
                <div style={{ fontSize: 12, fontWeight: 800, letterSpacing: 0.55, textTransform: "uppercase", color: "rgba(255,255,255,0.72)" }}>
                  Weekly pulse
                </div>
                <div style={{ marginTop: 10, fontSize: 24, fontWeight: 900, lineHeight: 1.08 }}>
                  {bestProduct?.name || "No product highlighted yet"}
                </div>
                <div style={{ marginTop: 10, color: "rgba(255,255,255,0.78)", lineHeight: 1.55 }}>
                  Best performer based on delivery health, ROAS, margin and operational readiness.
                </div>
                <div style={{ display: "grid", gap: 12, marginTop: 18 }}>
                    <MiniStat label="Profit" value={bestProduct ? formatTZS(bestProduct.profit) : "N/A"} tone="green" dark sub={bestProduct ? `${bestProduct.deliveredUnits} delivered units | ${bestProduct.availableStock} available` : "Add performance data to unlock insights"} />
                  <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("1fr 1fr", "1fr 1fr", "1fr"), gap: 12 }}>
                    <MiniStat label="Decision" value={bestProduct?.decision || "WATCH"} dark />
                    <MiniStat label="ROAS" value={bestProduct ? bestProduct.roas.toFixed(2) : "0.00"} tone="amber" dark />
                  </div>
                </div>
              </div>
            </div>
          </div>
          ) : null}

          {activePage === "executive" && shippingImportReminder.isVisible ? (
            <div
              style={{
                ...styles.card,
                marginBottom: 20,
                padding: 20,
                border: "1px solid #fde68a",
                background: "linear-gradient(135deg, rgba(255,251,235,0.98), rgba(255,244,214,0.94))",
              }}
            >
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 16, flexWrap: "wrap" }}>
                <div>
                  <div style={{ ...styles.sectionEyebrow, color: amber }}>Daily reminder</div>
                  <div style={{ fontSize: 22, fontWeight: 900, marginTop: 8 }}>Import the shipping status Excel before closing the day</div>
                  <div style={{ color: textSoft, marginTop: 8, lineHeight: 1.6 }}>
                    {shippingImportReminder.confirmedPipelineCount} confirmed order(s) still depend on today&apos;s shipping update. This reminder stays visible until you import the shipping Excel file.
                  </div>
                </div>
                <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                  <button style={styles.btnSecondary} onClick={() => setActivePage("shipping")}>
                    Open Shipping
                  </button>
                  <button
                    style={styles.btnPrimary}
                    onClick={() => {
                      setActivePage("shipping");
                      setTimeout(() => shippingImportInputRef.current?.click(), 50);
                    }}
                  >
                    Import Shipping Excel
                  </button>
                </div>
              </div>
            </div>
          ) : null}

          {activePage === "executive" ? (
            <div style={{ display: "grid", gap: 20, marginBottom: 20 }}>
              <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(6, minmax(0, 1fr))", "repeat(3, minmax(0, 1fr))", "1fr 1fr"), gap: 16 }}>
                <KpiCard icon={<Wallet size={18} />} title="Revenue" value={formatTZS(controlPanelSummary.totalRevenueTzs)} sub="Delivered orders only" valueColor={green} />
                <KpiCard icon={<TrendingUp size={18} />} title="Profit" value={formatTZS(controlPanelSummary.totalProfitTzs)} sub="After ads, product and delivery costs" valueColor={controlPanelSummary.totalProfitTzs >= 0 ? green : red} />
                <KpiCard icon={<Calculator size={18} />} title="Ads Spend" value={formatTZS(controlPanelSummary.totalAdsSpendTzs)} sub="Connected ads + manual tracking" valueColor={amber} />
                <KpiCard icon={<Rocket size={18} />} title="Delivered Orders" value={formatInteger(controlPanelSummary.totalDeliveredOrders)} sub="Completed deliveries" valueColor={green} />
                <KpiCard icon={<Phone size={18} />} title="Confirmation Rate" value={`${Math.round(controlPanelSummary.globalConfirmationRate)}%`} sub={`${formatInteger(controlPanelSummary.totalConfirmedOrders)} confirmed`} />
                <KpiCard icon={<ShoppingBag size={18} />} title="Delivery Rate" value={`${Math.round(controlPanelSummary.globalDeliveryRate)}%`} sub={`${formatInteger(controlPanelSummary.totalDeliveredOrders)} delivered`} valueColor={green} />
              </div>

              <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("1.1fr 1.1fr 1.1fr 1.3fr", "1fr 1fr", "1fr"), gap: 16 }}>
                <div style={{ ...styles.card, padding: 18 }}>
                  <div style={styles.sectionEyebrow}>Product to scale</div>
                  <div style={{ fontSize: 22, fontWeight: 900, marginTop: 8 }}>{homeCockpitSummary.productToScale?.name || "No scale candidate"}</div>
                  <div style={{ color: textSoft, marginTop: 8, lineHeight: 1.5 }}>
                    {homeCockpitSummary.productToScale
                      ? `${formatTZS(homeCockpitSummary.productToScale.dashboardProfitTzs || 0)} profit | ${Number(homeCockpitSummary.productToScale.dashboardProfitMargin || 0).toFixed(1)}% margin`
                      : "No product is clearly outperforming the rest yet."}
                  </div>
                </div>
                <div style={{ ...styles.card, padding: 18 }}>
                  <div style={styles.sectionEyebrow}>Product to stop / fix</div>
                  <div style={{ fontSize: 22, fontWeight: 900, marginTop: 8 }}>{homeCockpitSummary.productToStopOrFix?.name || "Nothing critical"}</div>
                  <div style={{ color: textSoft, marginTop: 8, lineHeight: 1.5 }}>
                    {homeCockpitSummary.productToStopOrFix
                      ? homeCockpitSummary.productToStopOrFix.productAlerts?.[0]?.message || `${formatTZS(homeCockpitSummary.productToStopOrFix.dashboardProfitTzs || 0)} profit`
                      : "No product currently needs a stop/fix call."}
                  </div>
                </div>
                <div style={{ ...styles.card, padding: 18 }}>
                  <div style={styles.sectionEyebrow}>Product to restock</div>
                  <div style={{ fontSize: 22, fontWeight: 900, marginTop: 8 }}>{homeCockpitSummary.productToRestock?.name || "Stock healthy"}</div>
                  <div style={{ color: textSoft, marginTop: 8, lineHeight: 1.5 }}>
                    {homeCockpitSummary.productToRestock
                      ? `${homeCockpitSummary.productToRestock.availableStock || 0} units available`
                      : "No urgent restock pressure right now."}
                  </div>
                </div>
                <div style={{ ...styles.card, padding: 18 }}>
                  <div style={styles.sectionEyebrow}>Biggest problem today</div>
                  <div style={{ fontSize: 22, fontWeight: 900, marginTop: 8 }}>{homeCockpitSummary.biggestProblemToday?.title || "No blocker"}</div>
                  <div style={{ color: textSoft, marginTop: 8, lineHeight: 1.5 }}>
                    {homeCockpitSummary.biggestProblemToday?.detail || "Nothing critical is blocking execution right now."}
                  </div>
                </div>
              </div>

              <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("1fr 1fr", "1fr", "1fr"), gap: 16 }}>
                <div style={{ ...styles.card, padding: 18 }}>
                  <div style={styles.sectionHeader}>
                    <div>
                      <div style={styles.sectionEyebrow}>Top 5 critical alerts</div>
                      <div style={{ fontSize: 22, fontWeight: 900, marginTop: 8 }}>What needs attention now</div>
                    </div>
                    <button style={styles.btnSecondary} onClick={() => setActivePage("taskCenter")}>Open Decisions</button>
                  </div>
                  <div style={{ display: "grid", gap: 10 }}>
                    {homeCockpitSummary.topAlerts.length ? homeCockpitSummary.topAlerts.map((row) => (
                      <div key={`home-alert-${row.id}`} style={{ ...styles.softStat, display: "grid", gap: 6 }}>
                        <div style={{ display: "flex", justifyContent: "space-between", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
                          <div style={{ fontWeight: 800 }}>{row.name}</div>
                          <span style={getDecisionStyle(row.performanceStatus === "WINNER" ? "OK" : row.performanceStatus === "LOSS" ? "KILL" : "WATCH")}>
                            {row.performanceStatus}
                          </span>
                        </div>
                        <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                          {row.productAlerts.slice(0, 2).map((alert) => (
                            <span key={`home-alert-badge-${row.id}-${alert.key}`} style={getAlertBadgeStyle(alert.tone)}>
                              {alert.message}
                            </span>
                          ))}
                        </div>
                      </div>
                    )) : <div style={{ color: textSoft }}>No active critical alerts.</div>}
                  </div>
                </div>
                <div style={{ ...styles.card, padding: 18 }}>
                  <div style={styles.sectionHeader}>
                    <div>
                      <div style={styles.sectionEyebrow}>Today action list</div>
                      <div style={{ fontSize: 22, fontWeight: 900, marginTop: 8 }}>Top 5 actions only</div>
                    </div>
                    <button style={styles.btnSecondary} onClick={() => setActivePage("taskCenter")}>Action board</button>
                  </div>
                  <div style={{ display: "grid", gap: 10 }}>
                    {homeCockpitSummary.topActions.length ? homeCockpitSummary.topActions.map((task, index) => (
                      <div key={`home-task-${task.id}`} style={{ ...styles.softStat, display: "grid", gap: 6 }}>
                        <div style={{ display: "flex", justifyContent: "space-between", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
                          <div style={{ fontWeight: 800 }}>{index + 1}. {task.title}</div>
                          <span style={getDecisionStyle(task.priority === "High" ? "KILL" : task.priority === "Medium" ? "WATCH" : "OK")}>
                            {task.priority}
                          </span>
                        </div>
                        <div style={{ color: textSoft, lineHeight: 1.5 }}>{task.detail}</div>
                      </div>
                    )) : <div style={{ color: textSoft }}>No action to push right now.</div>}
                  </div>
                </div>
              </div>
            </div>
          ) : null}

{activePage === "dashboard" && (
            <>
              {pendingDubaiNotifications.length > 0 && (
                <div
                  style={{
                    ...styles.card,
                    padding: 22,
                    marginBottom: 20,
                    border: "1px solid #fde68a",
                    background: "linear-gradient(135deg, rgba(255,251,235,0.98), rgba(255,245,222,0.94))",
                  }}
                >
                  <div style={{ ...styles.sectionEyebrow, color: amber, marginBottom: 6 }}>Attention required</div>
                  <div style={{ fontSize: 24, fontWeight: 900, marginBottom: 8 }}>Dubai Stock Arrival Notifications</div>
                  <div style={{ color: textSoft, marginBottom: 16, lineHeight: 1.6 }}>
                    If the stock has not arrived yet, click <strong>Not Yet</strong>. The dashboard will remind you again tomorrow until you confirm arrival.
                  </div>
                  <div style={{ display: "grid", gap: 12 }}>
                    {pendingDubaiNotifications.map((product) => (
                      <div
                        key={product.id}
                        style={{
                          background: "rgba(255,255,255,0.88)",
                          border: "1px solid #fde68a",
                          borderRadius: 18,
                          padding: 18,
                          display: "flex",
                          justifyContent: "space-between",
                          alignItems: "center",
                          gap: 16,
                          flexWrap: "wrap",
                          boxShadow: "0 12px 24px rgba(199, 131, 34, 0.08)",
                        }}
                      >
                        <div>
                          <div style={{ fontWeight: 800, fontSize: 16 }}>{product.name}</div>
                          <div style={{ color: textSoft, marginTop: 6, fontSize: 14 }}>
                            Source: Dubai | Ordered: {product.stockOrderedAt || "N/A"} | Estimated: {product.stockOrderedAt ? addDaysToDateString(product.stockOrderedAt, Number(product.estimatedArrivalDays || 0)) : "N/A"} | Next check: {product.nextArrivalCheckDate || "N/A"}
                          </div>
                        </div>
                        <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                          <button style={{ ...styles.btnSecondary, border: "1px solid #fecaca", color: red, background: "#fef2f2" }} onClick={() => markDubaiStockNotYet(product.id)}>
                            Not Yet
                          </button>
                          <button style={{ ...styles.btnPrimary, background: green }} onClick={() => markDubaiStockArrived(product.id)}>
                            Arrived
                          </button>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              )}

              {reorderNotifications.length > 0 && (
                <div
                  style={{
                    ...styles.card,
                    padding: 22,
                    marginBottom: 20,
                    border: "1px solid #fecaca",
                    background: "linear-gradient(135deg, rgba(255,247,237,0.98), rgba(255,237,237,0.9))",
                  }}
                >
                  <div style={{ ...styles.sectionEyebrow, color: red, marginBottom: 6 }}>Inventory risk</div>
                  <div style={{ fontSize: 24, fontWeight: 900, marginBottom: 8 }}>Stock Reorder Alerts</div>
                  <div style={{ color: textSoft, marginBottom: 16, lineHeight: 1.6 }}>These products are reaching their minimum stock level. Check them now to avoid stockout.</div>
                  <div style={{ display: "grid", gap: 12 }}>
                    {reorderNotifications.map((product) => (
                      <div
                        key={product.id}
                        style={{
                          background: "rgba(255,255,255,0.9)",
                          border: "1px solid #fed7aa",
                          borderRadius: 18,
                          padding: 18,
                          display: "flex",
                          justifyContent: "space-between",
                          alignItems: "center",
                          gap: 16,
                          flexWrap: "wrap",
                          boxShadow: "0 12px 24px rgba(217, 72, 95, 0.08)",
                        }}
                      >
                        <div>
                          <div style={{ fontWeight: 800, fontSize: 16 }}>{product.name}</div>
                          <div style={{ color: textSoft, marginTop: 6, fontSize: 14 }}>
                            Available: {product.availableStock} | Min stock: {product.reorderPoint} | Sales/day: {product.salesPerDay.toFixed(1)} | Source: {product.source || "N/A"}
                          </div>
                        </div>
                        <div style={getDecisionStyle(product.reorderStatus)}>{product.reorderStatus}</div>
                      </div>
                    ))}
                  </div>
                </div>
              )}

              <PageDateFilterBar
                title="Dashboard metrics range"
                value={dashboardDateFilter}
                onChange={setDashboardDateFilter}
              />

              <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(6, minmax(0, 1fr))", "repeat(2, minmax(0, 1fr))", "1fr"), gap: 16, marginBottom: 24 }}>
                <KpiCard
                  icon={<Users size={18} />}
                  title="Total Leads"
                  value={dashboardDateSummary.totalLeads}
                  sub="All incoming orders"
                  valueColor="#94a3b8"
                />
                <KpiCard
                  icon={<Phone size={18} />}
                  title="Confirmed Leads"
                  value={dashboardDateSummary.totalConfirmedOrders}
                  sub={`${dashboardDateSummary.globalConfirmationRate.toFixed(1)}% confirmation rate`}
                  valueColor="#f59e0b"
                />
                <KpiCard
                  icon={<Rocket size={18} />}
                  title="Delivered Leads"
                  value={dashboardDateSummary.totalDeliveredOrders}
                  sub={`${dashboardDateSummary.globalDeliveryRate.toFixed(1)}% delivery rate`}
                  valueColor="#16a34a"
                />
                <KpiCard
                  icon={<Wallet size={18} />}
                  title="Total Revenue"
                  value={formatTZS(dashboardDateSummary.totalRevenue)}
                  sub="From delivered leads only"
                  valueColor="#16a34a"
                />
                <KpiCard
                  icon={<ClipboardList size={18} />}
                  title="Ads Spend"
                  value={formatTZS(dashboardDateSummary.totalAdsSpend)}
                  sub="Tracking rows in range"
                  valueColor={amber}
                />
                <KpiCard
                  icon={<TrendingUp size={18} />}
                  title="Profit"
                  value={formatTZS(dashboardDateSummary.totalProfit)}
                  sub={`${dashboardDateSummary.averageProfitMargin.toFixed(1)}% average margin`}
                  valueColor={dashboardDateSummary.totalProfit >= 0 ? green : red}
                />
              </div>

              <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("1fr 1fr", "1fr", "1fr"), gap: 16, marginBottom: 16 }}>
                <div style={{ ...styles.card, padding: 22 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 10 }}>
                    <div style={{ width: 34, height: 34, borderRadius: 12, background: "linear-gradient(135deg, rgba(29,95,208,0.14), rgba(29,95,208,0.04))", display: "flex", alignItems: "center", justifyContent: "center", color: accent }}>
                      <Phone size={15} />
                    </div>
                    <div>
                      <div style={styles.sectionEyebrow}>Sales funnel</div>
                      <div style={{ fontSize: 18, fontWeight: 800, marginTop: 2 }}>CONFIRMATION DETAILS</div>
                    </div>
                  </div>

                  <div style={{ display: "grid", placeItems: "center", padding: "12px 0 18px", borderBottom: `1px solid ${cardBorder}` }}>
                    <div style={{ fontSize: 44, fontWeight: 800, color: accent }}>{confirmationDetails.total}</div>
                    <div style={{ color: textSoft, fontSize: 13 }}>Total Leads</div>
                  </div>

                  <div style={{ display: "grid", gap: 12, paddingTop: 16 }}>
                    {confirmationDetails.items.map((item) => (
                      <div key={item.label} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12 }}>
                        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                          <div style={{ width: 28, height: 28, borderRadius: 8, background: `${item.color}22`, display: "flex", alignItems: "center", justifyContent: "center" }}>
                            <span style={{ width: 10, height: 10, borderRadius: 999, background: item.color, display: "inline-block" }} />
                          </div>
                          <div>
                            <div style={{ fontWeight: 700, fontSize: 14 }}>{item.label}</div>
                            <div style={{ color: textSoft, fontSize: 12 }}>{item.count} orders</div>
                          </div>
                        </div>
                        <div style={{ fontWeight: 800, color: textMain }}>{item.pct}%</div>
                      </div>
                    ))}
                  </div>
                </div>

                <div style={{ ...styles.card, padding: 22 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 10 }}>
                    <div style={{ width: 34, height: 34, borderRadius: 12, background: "linear-gradient(135deg, rgba(31,143,95,0.14), rgba(31,143,95,0.04))", display: "flex", alignItems: "center", justifyContent: "center", color: green }}>
                      <Rocket size={15} />
                    </div>
                    <div>
                      <div style={{ ...styles.sectionEyebrow, color: green }}>Fulfillment</div>
                      <div style={{ fontSize: 18, fontWeight: 800, marginTop: 2 }}>DELIVERY DETAILS</div>
                    </div>
                  </div>

                  <div style={{ display: "grid", placeItems: "center", padding: "12px 0 18px", borderBottom: `1px solid ${cardBorder}` }}>
                    <div style={{ fontSize: 44, fontWeight: 800, color: accent }}>{deliveryDetails.total}</div>
                    <div style={{ color: textSoft, fontSize: 13 }}>Tracked Orders</div>
                  </div>

                  <div style={{ display: "grid", gap: 12, paddingTop: 16 }}>
                    {deliveryDetails.items.map((item) => (
                      <div key={item.label} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12 }}>
                        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                          <div style={{ width: 28, height: 28, borderRadius: 8, background: `${item.color}22`, display: "flex", alignItems: "center", justifyContent: "center" }}>
                            <span style={{ width: 10, height: 10, borderRadius: 999, background: item.color, display: "inline-block" }} />
                          </div>
                          <div>
                            <div style={{ fontWeight: 700, fontSize: 14 }}>{item.label}</div>
                            <div style={{ color: textSoft, fontSize: 12 }}>{item.count} orders</div>
                          </div>
                        </div>
                        <div style={{ fontWeight: 800, color: textMain }}>{item.pct}%</div>
                      </div>
                    ))}
                  </div>
                </div>

              </div>

              <div style={{ display: "grid", gridTemplateColumns: "1fr", gap: 16, marginBottom: 16 }}>
                <div style={{ ...styles.card, padding: 22 }}>
                  <div style={styles.sectionEyebrow}>Executive analytics</div>
                  <div style={{ fontSize: 24, fontWeight: 900, marginTop: 8, marginBottom: 8 }}>Global Overview</div>
                  <div style={{ color: textSoft, marginBottom: 18, lineHeight: 1.6 }}>Filtrable overview by product and custom date range, with conversion and delivery distribution.</div>

                  <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("1fr 1fr 1fr 1fr", "1fr 1fr", "1fr"), gap: 12, marginBottom: 18 }}>
                    <div style={styles.fieldBlock}>
                      <label style={styles.fieldLabel}>Produit</label>
                      <select
                        style={styles.input}
                        value={overviewFilters.productId}
                        onChange={(e) => setOverviewFilters({ ...overviewFilters, productId: e.target.value })}
                      >
                        <option value="all">Tous les produits</option>
                        {products.map((product) => (
                          <option key={product.id} value={product.id}>{product.name}</option>
                        ))}
                      </select>
                    </div>

                    <div style={styles.fieldBlock}>
                      <label style={styles.fieldLabel}>Periode</label>
                      <select
                        style={styles.input}
                        value={overviewFilters.periodMode}
                        onChange={(e) => setOverviewFilters({ ...overviewFilters, periodMode: e.target.value })}
                      >
                        <option value="all">Toutes les periodes</option>
                        <option value="custom">Periode personnalisee</option>
                      </select>
                    </div>

                    <div style={styles.fieldBlock}>
                      <label style={styles.fieldLabel}>Date debut</label>
                      <input
                        style={styles.input}
                        type="date"
                        value={overviewFilters.startDate}
                        disabled={overviewFilters.periodMode !== "custom"}
                        onChange={(e) => setOverviewFilters({ ...overviewFilters, startDate: e.target.value })}
                      />
                    </div>

                    <div style={styles.fieldBlock}>
                      <label style={styles.fieldLabel}>Date fin</label>
                      <input
                        style={styles.input}
                        type="date"
                        value={overviewFilters.endDate}
                        disabled={overviewFilters.periodMode !== "custom"}
                        onChange={(e) => setOverviewFilters({ ...overviewFilters, endDate: e.target.value })}
                      />
                    </div>
                  </div>
                  <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("0.95fr 1.05fr", "1fr", "1fr"), alignItems: "center", gap: 8 }}>
                    <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(2, minmax(0, 1fr))", "repeat(2, minmax(0, 1fr))", "1fr"), gap: 12, marginBottom: 14, gridColumn: "1 / -1" }}>
                      <div style={{ padding: "10px 12px", borderRadius: 14, background: "#f8fafc", border: `1px solid ${cardBorder}` }}>
                        <div style={{ color: textSoft, fontSize: 12, fontWeight: 700 }}>Incoming</div>
                        <div style={{ fontSize: 22, fontWeight: 800 }}>{overviewSummary.incoming}</div>
                      </div>
                      <div style={{ padding: "10px 12px", borderRadius: 14, background: "linear-gradient(135deg, rgba(245,158,11,0.12), rgba(245,158,11,0.04))", border: "1px solid rgba(245,158,11,0.14)" }}>
                        <div style={{ color: textSoft, fontSize: 12, fontWeight: 700 }}>Confirmed</div>
                        <div style={{ fontSize: 22, fontWeight: 800 }}>{overviewSummary.confirmed}</div>
                      </div>
                      <div style={{ padding: "10px 12px", borderRadius: 14, background: "linear-gradient(135deg, rgba(22,163,74,0.12), rgba(22,163,74,0.04))", border: "1px solid rgba(22,163,74,0.14)" }}>
                        <div style={{ color: textSoft, fontSize: 12, fontWeight: 700 }}>Delivered</div>
                        <div style={{ fontSize: 22, fontWeight: 800 }}>{overviewSummary.delivered}</div>
                      </div>
                      <div style={{ padding: "10px 12px", borderRadius: 14, background: "linear-gradient(135deg, rgba(29,95,208,0.12), rgba(29,95,208,0.04))", border: "1px solid rgba(29,95,208,0.12)" }}>
                        <div style={{ color: textSoft, fontSize: 12, fontWeight: 700 }}>Revenue</div>
                        <div style={{ fontSize: 22, fontWeight: 800 }}>{formatTZS(overviewSummary.revenue)}</div>
                      </div>
                    </div>
                    <div style={{ width: "100%", height: 260 }}>
                      <ResponsiveContainer>
                        <PieChart>
                          <Pie
                            data={overviewPieData}
                            dataKey="value"
                            nameKey="name"
                            innerRadius={58}
                            outerRadius={92}
                            paddingAngle={3}
                          >
                            {overviewPieData.map((entry) => (
                              <Cell key={entry.name} fill={entry.color} />
                            ))}
                          </Pie>
                          <Tooltip formatter={(value) => `${value}%`} />
                        </PieChart>
                      </ResponsiveContainer>
                    </div>
                    <div style={{ display: "grid", gap: 12 }}>
                      {overviewPieData.map((item) => (
                        <div
                          key={item.name}
                          style={{
                            display: "flex",
                            alignItems: "center",
                            justifyContent: "space-between",
                            padding: "10px 12px",
                            borderRadius: 14,
                            background: "#f8fafc",
                            border: `1px solid ${cardBorder}`,
                          }}
                        >
                          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                            <span style={{ width: 12, height: 12, borderRadius: 999, background: item.color, display: "inline-block" }} />
                            <span style={{ color: textSoft, fontWeight: 600 }}>{item.name}</span>
                          </div>
                          <strong style={{ color: item.color }}>{item.value}%</strong>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>

              </div>

              <div style={{ ...styles.card, padding: 22, marginBottom: 16 }}>
                <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 10, marginBottom: 18 }}>
                  <Phone size={18} color="#ef4444" />
                  <div style={{ fontSize: 18, fontWeight: 900 }}>CONFIRMATION SUMMARY</div>
                </div>

                <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("1fr 1fr 1fr 1fr 0.3fr", "1fr 1fr", "1fr"), gap: 12, marginBottom: 14 }}>
                  <select
                    style={styles.input}
                    value={confirmationSummaryFilters.period}
                    onChange={(e) =>
                      setConfirmationSummaryFilters((prev) => ({
                        ...prev,
                        period: e.target.value,
                      }))
                    }
                  >
                    <option value="today">Today</option>
                    <option value="yesterday">Yesterday</option>
                    <option value="thisWeek">This Week</option>
                    <option value="thisMonth">This Month</option>
                    <option value="all">All Time</option>
                    <option value="custom">Custom Date</option>
                  </select>

                  <input
                    type="date"
                    style={styles.input}
                    disabled={confirmationSummaryFilters.period !== "custom"}
                    value={confirmationSummaryFilters.startDate}
                    onChange={(e) =>
                      setConfirmationSummaryFilters((prev) => ({
                        ...prev,
                        startDate: e.target.value,
                      }))
                    }
                  />

                  <input
                    type="date"
                    style={styles.input}
                    disabled={confirmationSummaryFilters.period !== "custom"}
                    value={confirmationSummaryFilters.endDate}
                    onChange={(e) =>
                      setConfirmationSummaryFilters((prev) => ({
                        ...prev,
                        endDate: e.target.value,
                      }))
                    }
                  />

                  <select
                    style={styles.input}
                    value={confirmationSummaryFilters.productId}
                    onChange={(e) =>
                      setConfirmationSummaryFilters((prev) => ({
                        ...prev,
                        productId: e.target.value,
                      }))
                    }
                  >
                    <option value="all">Filter by product</option>
                    {products.map((product) => (
                      <option key={product.id} value={product.id}>
                        {product.name}
                      </option>
                    ))}
                  </select>

                  <button
                    style={{ ...styles.btnPrimary, background: "#dc2626", padding: "12px 0" }}
                    onClick={() =>
                      setConfirmationSummaryFilters({
                        period: "thisWeek",
                        productId: "all",
                        startDate: "",
                        endDate: "",
                      })
                    }
                  >
                    Reset
                  </button>
                </div>

                <div style={{ fontWeight: 800, marginBottom: 14 }}>
                  Confirmation Rate: {confirmationSummary.confirmationRate.toFixed(0)}% ({confirmationSummary.confirmed} Orders confirmed)
                </div>

                <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("0.27fr 0.43fr 0.30fr", "1fr", "1fr"), gap: 12, alignItems: "start" }}>
                  <div style={{ display: "grid", gap: 10, paddingTop: 12 }}>
                    {confirmationSummary.breakdown.map((item) => (
                      <div key={item.label} style={{ display: "flex", alignItems: "center", gap: 8 }}>
                        <span style={{ width: 10, height: 10, borderRadius: 999, background: item.color, display: "inline-block" }} />
                        <span style={{ fontSize: 13 }}>{item.label.toLowerCase()}</span>
                      </div>
                    ))}
                  </div>

                  <div>
                    <div style={{ fontWeight: 800, textAlign: "center", marginBottom: 8 }}>Confirmation per day</div>
                    <div style={{ width: "100%", height: 260 }}>
                      <ResponsiveContainer>
                        <LineChart data={confirmationSummary.chartData}>
                          <CartesianGrid strokeDasharray="3 3" />
                          <XAxis dataKey="date" />
                          <YAxis allowDecimals={false} />
                          <Tooltip />
                          <Line type="monotone" dataKey="cancelled" stroke="#ef4444" strokeWidth={2} />
                          <Line type="monotone" dataKey="confirmed" stroke="#84cc16" strokeWidth={2} />
                          <Line type="monotone" dataKey="newOrder" stroke="#6366f1" strokeWidth={2} />
                          <Line type="monotone" dataKey="pending" stroke="#67e8f9" strokeWidth={2} />
                        </LineChart>
                      </ResponsiveContainer>
                    </div>
                  </div>

                  <div style={{ borderLeft: `1px solid ${cardBorder}`, paddingLeft: 18 }}>
                    {confirmationSummary.breakdown.map((item) => (
                      <div key={item.label} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "8px 0", borderBottom: `1px solid ${cardBorder}` }}>
                        <div style={{ fontWeight: 700 }}>{item.label}</div>
                        <div style={{ display: "inline-flex", alignItems: "center", justifyContent: "center", minWidth: 28, height: 28, borderRadius: 999, background: item.color, color: "white", fontWeight: 800, fontSize: 12 }}>{item.count}</div>
                      </div>
                    ))}
                  </div>
                </div>
              </div>

              <div style={{ ...styles.card, padding: 22, position: "relative", zIndex: 20, overflow: "visible" }}>
                <div style={styles.sectionHeader}>
                  <div>
                    <div style={{ fontSize: 22, fontWeight: 800 }}>Evolution des commandes</div>
                    <div style={{ color: textSoft, marginTop: 6 }}>
                      Comparaison entre les commandes entrantes, confirmees et livrees selon la date de commande.
                    </div>
                  </div>
                </div>

                <div style={{ width: "100%", height: 360 }}>
                  <ResponsiveContainer>
                    <BarChart data={ordersChartData}>
                      <CartesianGrid strokeDasharray="3 3" />
                      <XAxis dataKey="date" />
                      <YAxis allowDecimals={false} />
                      <Tooltip />
                      <Legend />
                      <Bar dataKey="incoming" name="Leads" fill="#94a3b8" radius={[6, 6, 0, 0]} />
                      <Bar dataKey="confirmed" name="Confirmees" fill="#f59e0b" radius={[6, 6, 0, 0]} />
                      <Bar dataKey="delivered" name="Livrees" fill="#16a34a" radius={[6, 6, 0, 0]} />
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>

              <div style={{ ...styles.card, padding: 22, marginTop: 16 }}>
                <div style={{ ...styles.sectionHeader, alignItems: "flex-start", flexDirection: isCompact ? "column" : "row" }}>
                  <div>
                    <div style={styles.sectionEyebrow}>Products intelligence</div>
                    <div style={{ fontSize: 24, fontWeight: 900, marginTop: 8 }}>
                      Products Details
                      <span style={{ color: amber, marginLeft: 10, fontSize: 16, fontWeight: 700 }}>
                        ({visibleProductDetailsRows.length > 0 ? 1 : 0} - {visibleProductDetailsRows.length} of {productDetailsRows.length})
                      </span>
                    </div>
                    <div style={{ color: textSoft, marginTop: 6, lineHeight: 1.6 }}>
                      Detailed leaderboard of your products with lead quality, delivery efficiency, revenue and AOV.
                    </div>
                  </div>

                  <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("120px 160px 160px 110px", "1fr 1fr", "1fr"), gap: 12, width: isCompact ? "100%" : "auto" }}>
                    <select
                      style={styles.input}
                      value={productDetailsFilters.rowLimit}
                      onChange={(e) =>
                        setProductDetailsFilters((prev) => ({
                          ...prev,
                          rowLimit: Number(e.target.value),
                        }))
                      }
                    >
                      <option value={5}>5</option>
                      <option value={10}>10</option>
                      <option value={25}>25</option>
                      <option value={50}>50</option>
                    </select>

                    <select
                      style={styles.input}
                      value={productDetailsFilters.period}
                      onChange={(e) =>
                        setProductDetailsFilters((prev) => ({
                          ...prev,
                          period: e.target.value,
                          startDate: e.target.value === "custom" ? prev.startDate : "",
                          endDate: e.target.value === "custom" ? prev.endDate : "",
                        }))
                      }
                    >
                      <option value="last7Days">Last 7 days</option>
                      <option value="today">Today</option>
                      <option value="yesterday">Yesterday</option>
                      <option value="thisWeek">This week</option>
                      <option value="thisMonth">This month</option>
                      <option value="all">All time</option>
                      <option value="custom">Custom date</option>
                    </select>

                    <select
                      style={styles.input}
                      value={productDetailsFilters.productId}
                      onChange={(e) =>
                        setProductDetailsFilters((prev) => ({
                          ...prev,
                          productId: e.target.value,
                        }))
                      }
                    >
                      <option value="all">All products</option>
                      {products.map((product) => (
                        <option key={product.id} value={product.id}>
                          {product.name}
                        </option>
                      ))}
                    </select>

                    <button
                      style={styles.btnSecondary}
                      onClick={() =>
                        setProductDetailsFilters({
                          period: "last7Days",
                          productId: "all",
                          startDate: "",
                          endDate: "",
                          rowLimit: 10,
                        })
                      }
                    >
                      Reset view
                    </button>
                  </div>
                </div>

                {productDetailsFilters.period === "custom" && (
                  <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("180px 180px", "1fr 1fr", "1fr"), gap: 12, marginBottom: 18 }}>
                    <input
                      type="date"
                      style={styles.input}
                      value={productDetailsFilters.startDate}
                      onChange={(e) =>
                        setProductDetailsFilters((prev) => ({
                          ...prev,
                          startDate: e.target.value,
                        }))
                      }
                    />
                    <input
                      type="date"
                      style={styles.input}
                      value={productDetailsFilters.endDate}
                      onChange={(e) =>
                        setProductDetailsFilters((prev) => ({
                          ...prev,
                          endDate: e.target.value,
                        }))
                      }
                    />
                  </div>
                )}

                <div style={{ overflowX: "auto", border: `1px solid ${cardBorder}`, borderRadius: 22, background: "linear-gradient(180deg, rgba(255,255,255,0.95), rgba(248,244,238,0.86))" }}>
                  <table style={{ width: "100%", borderCollapse: "separate", borderSpacing: 0 }}>
                    <thead>
                      <tr>
                        {["#", "Product", "Leads", "Confirmation (%)", "Delivery (%)", "Rate From Lead", "Delivered Leads", "Total Revenue", "AOV"].map((head) => (
                          <th
                            key={head}
                            style={{
                              textAlign: head === "Product" ? "left" : "center",
                              padding: "15px 14px",
                              color: textSoft,
                              fontSize: 12,
                              fontWeight: 800,
                              letterSpacing: 0.45,
                              textTransform: "uppercase",
                              borderBottom: `1px solid ${cardBorder}`,
                              background: "rgba(247, 243, 237, 0.92)",
                              whiteSpace: "nowrap",
                            }}
                          >
                            {head}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {visibleProductDetailsRows.map((row, index) => (
                        <tr key={row.id} style={{ background: index % 2 === 0 ? "rgba(255,255,255,0.76)" : "rgba(243,246,251,0.62)" }}>
                          <td style={{ padding: "14px", textAlign: "center", borderBottom: `1px solid ${cardBorder}`, fontWeight: 800, color: accent }}>
                            #{index + 1}
                          </td>
                          <td style={{ padding: "14px", borderBottom: `1px solid ${cardBorder}`, minWidth: 260 }}>
                            <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
                              <div
                                style={{
                                  width: 38,
                                  height: 38,
                                  borderRadius: 12,
                                  display: "grid",
                                  placeItems: "center",
                                  background: row.source === "dubai"
                                    ? "linear-gradient(135deg, rgba(199,131,34,0.18), rgba(199,131,34,0.06))"
                                    : "linear-gradient(135deg, rgba(29,95,208,0.14), rgba(29,95,208,0.04))",
                                  color: row.source === "dubai" ? amber : accent,
                                  fontWeight: 900,
                                  flexShrink: 0,
                                }}
                              >
                                {row.initials}
                              </div>
                              <div>
                                <div style={{ fontWeight: 800 }}>{row.name}</div>
                                <div style={{ color: textSoft, fontSize: 12, marginTop: 4 }}>{row.id} | {row.source || "N/A"}</div>
                              </div>
                            </div>
                          </td>
                          <td style={{ padding: "14px", textAlign: "center", borderBottom: `1px solid ${cardBorder}`, fontWeight: 700 }}>{row.leads}</td>
                          <td style={{ padding: "14px", textAlign: "center", borderBottom: `1px solid ${cardBorder}` }}>{row.confirmationRate.toFixed(0)}%</td>
                          <td style={{ padding: "14px", textAlign: "center", borderBottom: `1px solid ${cardBorder}` }}>{row.deliveryRate.toFixed(0)}%</td>
                          <td style={{ padding: "14px", textAlign: "center", borderBottom: `1px solid ${cardBorder}` }}>{row.leadToDeliveryRate.toFixed(0)}%</td>
                          <td style={{ padding: "14px", textAlign: "center", borderBottom: `1px solid ${cardBorder}`, fontWeight: 700, color: green }}>{row.deliveredOrders}</td>
                          <td style={{ padding: "14px", textAlign: "center", borderBottom: `1px solid ${cardBorder}`, fontWeight: 700 }}>{formatTZS(row.totalRevenue)}</td>
                          <td style={{ padding: "14px", textAlign: "center", borderBottom: `1px solid ${cardBorder}` }}>{formatTZS(row.aov)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  {visibleProductDetailsRows.length === 0 ? (
                    <div style={{ padding: 24, color: textSoft }}>No product data available for the selected filters.</div>
                  ) : null}
                </div>
              </div>
            </>
          )}

{activePage === "multiDashboard" && (
            <div style={{ ...styles.card, padding: 22 }}>
              <div style={styles.sectionHeader}>
                <div>
                  <div style={{ fontSize: 24, fontWeight: 900 }}>Fichier stock</div>
                  <div style={{ color: textSoft, marginTop: 6 }}>
                    {editingProductId
                      ? "Modifiez les informations du lot fournisseur puis enregistrez les changements."
                      : "Ajoutez ici les produits commandes chez le fournisseur. Une fois sauvegardes, ils rejoignent automatiquement le stock produit."}
                  </div>
                </div>
                <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                  {editingProductId ? (
                    <button style={styles.btnSecondary} onClick={cancelEditingProduct}>Cancel Edit</button>
                  ) : null}
                  <button style={styles.btnSecondary} onClick={() => setActivePage("products")}>
                    Voir le stock
                  </button>
                  <button style={styles.btnPrimary} onClick={saveExpeditionProduct}>
                    {editingProductId ? "Update Product" : "Save Product"}
                  </button>
                </div>
              </div>

                <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(2, minmax(0, 1fr))", "1fr", "1fr"), gap: 16 }}>
                <div style={styles.fieldBlock}>
                  <label style={styles.fieldLabel}>Product Name</label>
                  <input
                    style={styles.input}
                    placeholder="Ex: Electric Callus Remover"
                    value={expeditionForm.name}
                    onChange={(e) => {
                      const name = e.target.value;
                      setExpeditionForm((prev) => ({
                        ...prev,
                        name,
                        mappingCode: generateMappingCode(name, products, editingProductId || undefined),
                      }));
                    }}
                  />
                </div>
                <div style={styles.fieldBlock}>
                  <label style={styles.fieldLabel}>Mapping Code</label>
                  <input
                    style={styles.input}
                    placeholder="Auto-generated"
                    value={expeditionForm.mappingCode || ""}
                    onChange={(e) =>
                      setExpeditionForm((prev) => ({
                        ...prev,
                        mappingCode: e.target.value.toUpperCase().replace(/[^A-Z0-9]/g, ""),
                      }))
                    }
                  />
                  <div style={{ fontSize: 11, color: textSoft, marginTop: 5 }}>
                    Auto-generated · uppercase + numbers only · editable
                  </div>
                </div>
                <div style={styles.fieldBlock}>
                  <label style={styles.fieldLabel}>Source (China / Dubai)</label>
                  <select style={styles.input} value={expeditionForm.source} onChange={(e) => setExpeditionForm({ ...expeditionForm, source: e.target.value })}>
                    <option value="china">China</option>
                    <option value="dubai">Dubai</option>
                  </select>
                </div>
                <div style={styles.fieldBlock}>
                  <label style={styles.fieldLabel}>Selling Price (TZS)</label>
                  <input style={styles.input} type="number" placeholder="Ex: 39000" value={expeditionForm.sellingPrice} onChange={(e) => setExpeditionForm({ ...expeditionForm, sellingPrice: e.target.value })} />
                </div>
                <div style={styles.fieldBlock}>
                  <label style={styles.fieldLabel}>Purchase Unit Price (USD)</label>
                  <input style={styles.input} type="number" placeholder="Ex: 5" value={expeditionForm.purchaseUnitPrice} onChange={(e) => setExpeditionForm({ ...expeditionForm, purchaseUnitPrice: e.target.value })} />
                </div>
                <div style={styles.fieldBlock}>
                  <label style={styles.fieldLabel}>Total Quantity</label>
                  <input style={styles.input} type="number" placeholder="Ex: 100" value={expeditionForm.totalQty} onChange={(e) => setExpeditionForm({ ...expeditionForm, totalQty: e.target.value })} />
                </div>
                <div style={styles.fieldBlock}>
                  <label style={styles.fieldLabel}>Shipping Total (TZS)</label>
                  <input style={styles.input} type="number" placeholder="Ex: 220000" value={expeditionForm.shippingTotal} onChange={(e) => setExpeditionForm({ ...expeditionForm, shippingTotal: e.target.value })} />
                </div>
                <div style={styles.fieldBlock}>
                  <label style={styles.fieldLabel}>Other Charges (TZS)</label>
                  <input style={styles.input} type="number" placeholder="Ex: 90000" value={expeditionForm.otherCharges} onChange={(e) => setExpeditionForm({ ...expeditionForm, otherCharges: e.target.value })} />
                </div>
                <div style={styles.fieldBlock}>
                  <label style={styles.fieldLabel}>Local Delivery (TZS)</label>
                  <input style={styles.input} type="number" placeholder="Ex: 7000" value={expeditionForm.delivery} onChange={(e) => setExpeditionForm({ ...expeditionForm, delivery: e.target.value })} />
                </div>
                  <div style={styles.fieldBlock}>
                    <label style={styles.fieldLabel}>Estimated Arrival Days</label>
                    <input style={styles.input} type="number" placeholder="Ex: 3 for Dubai, 15 for China" value={expeditionForm.estimatedArrivalDays} onChange={(e) => setExpeditionForm({ ...expeditionForm, estimatedArrivalDays: e.target.value })} />
                  </div>
                  <div style={styles.fieldBlock}>
                    <label style={styles.fieldLabel}>Supplier Name</label>
                    <input style={styles.input} value={expeditionForm.supplierName} onChange={(e) => setExpeditionForm({ ...expeditionForm, supplierName: e.target.value })} placeholder="Supplier or sourcing agent" />
                  </div>
                  <div style={styles.fieldBlock}>
                    <label style={styles.fieldLabel}>Supplier Contact</label>
                    <input style={styles.input} value={expeditionForm.supplierContact} onChange={(e) => setExpeditionForm({ ...expeditionForm, supplierContact: e.target.value })} placeholder="Phone, WhatsApp, email..." />
                  </div>
                  <div style={styles.fieldBlock}>
                    <label style={styles.fieldLabel}>Lifecycle</label>
                    <select style={styles.input} value={expeditionForm.lifecycleStatus} onChange={(e) => setExpeditionForm({ ...expeditionForm, lifecycleStatus: e.target.value })}>
                      <option value="test">Test</option>
                      <option value="winner">Winner</option>
                      <option value="scaling">Scaling</option>
                      <option value="mature">Mature</option>
                      <option value="declining">Declining</option>
                      <option value="kill">Kill</option>
                    </select>
                  </div>
                  <div style={styles.fieldBlock}>
                    <label style={styles.fieldLabel}>Defect Rate %</label>
                    <input style={styles.input} type="number" min="0" step="0.1" value={expeditionForm.defectRate} onChange={(e) => setExpeditionForm({ ...expeditionForm, defectRate: e.target.value })} />
                  </div>
                  <div style={{ ...styles.fieldBlock, gridColumn: isCompact ? "auto" : "1 / -1" }}>
                    <label style={styles.fieldLabel}>Product Notes</label>
                    <textarea style={{ ...styles.input, minHeight: 84, resize: "vertical" }} value={expeditionForm.notes} onChange={(e) => setExpeditionForm({ ...expeditionForm, notes: e.target.value })} />
                  </div>
                  <div style={{ ...styles.kpiCard }}>
                  <div style={{ color: textSoft, fontSize: 13, fontWeight: 600 }}>Auto Cost / Piece (USD)</div>
                  <div style={{ fontSize: 28, fontWeight: 800, marginTop: 10 }}>
                    {formatUSD(
                      Number(expeditionForm.totalQty || 0) > 0
                        ? ((Number(expeditionForm.purchaseUnitPrice || 0) * Number(expeditionForm.totalQty || 0)) + (Number(expeditionForm.shippingTotal || 0) / USD_TO_TZS) + (Number(expeditionForm.otherCharges || 0) / USD_TO_TZS)) / Number(expeditionForm.totalQty || 0)
                        : 0
                    )}
                  </div>
                  <div style={{ marginTop: 8, color: textSoft, fontSize: 13 }}>Calculated automatically before saving.</div>
                </div>
              </div>

              <div style={{ ...styles.softStat, marginTop: 16 }}>
                <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap", marginBottom: 14 }}>
                  <div>
                    <div style={{ fontSize: 11, fontWeight: 800, letterSpacing: 0.45, textTransform: "uppercase", color: textSoft }}>Product offers</div>
                    <div style={{ marginTop: 8, fontSize: 18, fontWeight: 900 }}>Configure quantity bundles</div>
                    <div style={{ marginTop: 6, color: textSoft, fontSize: 13, lineHeight: 1.5 }}>Example: `2 pcs = 150,000 TZS`, `3 pcs = 210,000 TZS`. The app will use these offers in order value and revenue metrics.</div>
                  </div>
                  <button style={styles.btnSecondary} onClick={addProductOfferTier}>Add Offer</button>
                </div>

                <div style={{ display: "grid", gap: 10 }}>
                  {(expeditionForm.offers || []).length ? (
                    expeditionForm.offers.map((offer, index) => (
                      <div key={`${offer.minQty}-${index}`} style={{ display: "grid", gridTemplateColumns: responsiveColumns("120px 1fr auto", "120px 1fr auto", "1fr"), gap: 10, alignItems: "end" }}>
                        <div style={styles.fieldBlock}>
                          <label style={styles.fieldLabel}>Min Qty</label>
                          <input style={styles.input} type="number" min="2" value={offer.minQty} onChange={(e) => updateProductOfferTier(index, "minQty", e.target.value)} />
                        </div>
                        <div style={styles.fieldBlock}>
                          <label style={styles.fieldLabel}>Offer Total Price (TZS)</label>
                          <input style={styles.input} type="number" min="0" value={offer.totalPrice} onChange={(e) => updateProductOfferTier(index, "totalPrice", e.target.value)} />
                        </div>
                        <button style={{ ...styles.btnSecondary, background: "#fef2f2", color: red, border: "1px solid #fecaca" }} onClick={() => removeProductOfferTier(index)}>
                          Remove
                        </button>
                      </div>
                    ))
                  ) : (
                    <div style={{ color: textSoft, fontSize: 14 }}>No bundle offer yet. The base selling price will be used for all quantities.</div>
                  )}
                </div>
              </div>
            </div>
          )}

{activePage === "products" && (
            <div style={{ display: "grid", gap: 20 }}>
              <PageHeader
                eyebrow="Products & Stock"
                title="Catalog and stock management"
                description="Manage product catalog, stock purchases, availability and movements from one page."
                action={(
                  <>
                    {clearProductsConfirm ? (
                      <>
                        <span style={{ fontSize: 13, color: red, fontWeight: 700 }}>Delete all products?</span>
                        <button style={{ ...styles.btnSecondary, background: "#fef2f2", color: red, border: "1px solid #fecaca" }} onClick={handleClearAllProducts}>Yes, delete all</button>
                        <button style={styles.btnSecondary} onClick={() => setClearProductsConfirm(false)}>Cancel</button>
                      </>
                    ) : (
                      <button style={{ ...styles.btnSecondary, color: red, borderColor: "#fecaca" }} onClick={() => setClearProductsConfirm(true)}>
                        Clear all products
                      </button>
                    )}
                    <button style={styles.btnPrimary} onClick={() => { setStockTab("catalog"); setShowAddProductForm(true); setEditingProductId(null); setExpeditionForm(getEmptyExpeditionForm()); }}>
                      New product
                    </button>
                  </>
                )}
              />
              <InlineTabs
                items={[
                  { value: "catalog", label: "Product Catalog" },
                  { value: "purchases", label: "Stock Purchases" },
                  { value: "incoming", label: "Incoming Shipments" },
                  { value: "overview", label: "Stock Overview" },
                  { value: "movements", label: "Stock Movements" },
                  { value: "alerts", label: "Stock Alerts" },
                  { value: "audit", label: "Stock Audit" },
                ]}
                value={stockTab}
                onChange={setStockTab}
              />
              <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(4, minmax(0, 1fr))", "repeat(2, minmax(0, 1fr))", "1fr"), gap: 16 }}>
                <KpiCard icon={<Boxes size={18} />} title="Catalog" value={productsCatalogSummary.totalProducts} sub="Products in catalog" />
                <KpiCard icon={<Archive size={18} />} title="Available stock" value={stockAuditData.totalAvailable} sub={`${stockAuditData.totalIncoming} incoming`} />
                <KpiCard icon={<AlertTriangle size={18} />} title="Alerts" value={stockAlerts.length} sub={`${stockAlerts.filter((a) => a.severity === "critical").length} critical`} valueColor={stockAlerts.some((a) => a.severity === "critical") ? red : amber} />
                <KpiCard icon={<Wallet size={18} />} title="Stock value" value={formatUSD(stockAuditData.totalStockValueUsd)} sub="Available stock valuation" valueColor={accent} />
              </div>

              {/* ===== TAB: PRODUCT CATALOG ===== */}
              <div style={{ display: stockTab === "catalog" ? "grid" : "none", gap: 16 }}>
              <div style={{ ...styles.card, padding: 22 }}>
                <div style={styles.sectionHeader}>
                  <div>
                    <div style={styles.sectionEyebrow}>Product catalog</div>
                    <div style={{ fontSize: 24, fontWeight: 900, marginTop: 8 }}>All products</div>
                    <div style={{ color: textSoft, marginTop: 6, lineHeight: 1.6 }}>Define product catalog, prices and mapping codes. Mapping codes link products to ad campaigns.</div>
                  </div>
                  <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
                    {editingProductId ? (
                      <button style={styles.btnSecondary} onClick={cancelEditingProduct}>Cancel Edit</button>
                    ) : null}
                    <button style={styles.btnPrimary} onClick={() => { setShowAddProductForm(true); setEditingProductId(null); setExpeditionForm(getEmptyExpeditionForm()); }}>Add Product</button>
                  </div>
                </div>

                {showAddProductForm && !editingProductId ? (
                  <div style={{ ...styles.softStat, marginBottom: 18 }}>
                    <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap", marginBottom: 16 }}>
                      <div>
                        <div style={{ fontSize: 11, fontWeight: 800, letterSpacing: 0.45, textTransform: "uppercase", color: textSoft }}>Add new product</div>
                        <div style={{ marginTop: 8, fontSize: 22, fontWeight: 900 }}>New product</div>
                      </div>
                      <div style={{ display: "flex", gap: 10 }}>
                        <button style={styles.btnSecondary} onClick={cancelEditingProduct}>Cancel</button>
                        <button style={styles.btnPrimary} onClick={saveExpeditionProduct}>Save Product</button>
                      </div>
                    </div>
                    <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(2, minmax(0, 1fr))", "1fr", "1fr"), gap: 16 }}>
                      <div style={styles.fieldBlock}>
                        <label style={styles.fieldLabel}>Product Name</label>
                        <input style={styles.input} placeholder="Ex: Electric Callus Remover" value={expeditionForm.name} onChange={(e) => { const name = e.target.value; setExpeditionForm((prev) => ({ ...prev, name, mappingCode: generateMappingCode(name, products, undefined) })); }} />
                      </div>
                      <div style={styles.fieldBlock}>
                        <label style={styles.fieldLabel}>Mapping Code</label>
                        <input style={styles.input} placeholder="Auto-generated" value={expeditionForm.mappingCode || ""} onChange={(e) => setExpeditionForm((prev) => ({ ...prev, mappingCode: e.target.value.toUpperCase().replace(/[^A-Z0-9]/g, "") }))} />
                        <div style={{ fontSize: 11, color: textSoft, marginTop: 5 }}>Auto-generated · uppercase + numbers only · editable</div>
                      </div>
                      <div style={styles.fieldBlock}>
                        <label style={styles.fieldLabel}>Source</label>
                        <select style={styles.input} value={expeditionForm.source} onChange={(e) => setExpeditionForm({ ...expeditionForm, source: e.target.value })}>
                          <option value="china">China</option>
                          <option value="dubai">Dubai</option>
                        </select>
                      </div>
                      <div style={styles.fieldBlock}>
                        <label style={styles.fieldLabel}>Selling Price (TZS)</label>
                        <input style={styles.input} type="number" placeholder="Ex: 39000" value={expeditionForm.sellingPrice} onChange={(e) => setExpeditionForm({ ...expeditionForm, sellingPrice: e.target.value })} />
                      </div>
                      <div style={styles.fieldBlock}>
                        <label style={styles.fieldLabel}>Purchase Unit Price (USD)</label>
                        <input style={styles.input} type="number" placeholder="Ex: 5" value={expeditionForm.purchaseUnitPrice} onChange={(e) => setExpeditionForm({ ...expeditionForm, purchaseUnitPrice: e.target.value })} />
                      </div>
                      <div style={styles.fieldBlock}>
                        <label style={styles.fieldLabel}>Total Quantity</label>
                        <input style={styles.input} type="number" placeholder="Ex: 100" value={expeditionForm.totalQty} onChange={(e) => setExpeditionForm({ ...expeditionForm, totalQty: e.target.value })} />
                      </div>
                      <div style={styles.fieldBlock}>
                        <label style={styles.fieldLabel}>Supplier Name</label>
                        <input style={styles.input} placeholder="Supplier or agent" value={expeditionForm.supplierName || ""} onChange={(e) => setExpeditionForm({ ...expeditionForm, supplierName: e.target.value })} />
                      </div>
                      <div style={styles.fieldBlock}>
                        <label style={styles.fieldLabel}>Lifecycle</label>
                        <select style={styles.input} value={expeditionForm.lifecycleStatus || "test"} onChange={(e) => setExpeditionForm({ ...expeditionForm, lifecycleStatus: e.target.value })}>
                          <option value="test">Test</option>
                          <option value="winner">Winner</option>
                          <option value="scaling">Scaling</option>
                          <option value="mature">Mature</option>
                          <option value="discontinued">Discontinued</option>
                        </select>
                      </div>
                    </div>
                  </div>
                ) : null}

                {editingProductId ? (
                  <div style={{ ...styles.softStat, marginBottom: 18 }}>
                    <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap", marginBottom: 16 }}>
                      <div>
                        <div style={{ fontSize: 11, fontWeight: 800, letterSpacing: 0.45, textTransform: "uppercase", color: textSoft }}>Product editor</div>
                        <div style={{ marginTop: 8, fontSize: 22, fontWeight: 900 }}>
                          Edit {expeditionForm.name || "product"}
                        </div>
                      </div>
                      <button style={styles.btnPrimary} onClick={saveExpeditionProduct}>
                        Update Product
                      </button>
                    </div>
                    <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(2, minmax(0, 1fr))", "1fr", "1fr"), gap: 16 }}>
                      <div style={styles.fieldBlock}>
                        <label style={styles.fieldLabel}>Product Name</label>
                        <input style={styles.input} value={expeditionForm.name} onChange={(e) => setExpeditionForm({ ...expeditionForm, name: e.target.value })} />
                      </div>
                      <div style={styles.fieldBlock}>
                    <label style={styles.fieldLabel}>Source</label>
                        <select style={styles.input} value={expeditionForm.source} onChange={(e) => setExpeditionForm({ ...expeditionForm, source: e.target.value })}>
                          <option value="china">China</option>
                          <option value="dubai">Dubai</option>
                        </select>
                      </div>
                      <div style={styles.fieldBlock}>
                        <label style={styles.fieldLabel}>Selling Price (TZS)</label>
                        <input style={styles.input} type="number" value={expeditionForm.sellingPrice} onChange={(e) => setExpeditionForm({ ...expeditionForm, sellingPrice: e.target.value })} />
                      </div>
                      <div style={styles.fieldBlock}>
                        <label style={styles.fieldLabel}>Purchase Unit Price (USD)</label>
                        <input style={styles.input} type="number" value={expeditionForm.purchaseUnitPrice} onChange={(e) => setExpeditionForm({ ...expeditionForm, purchaseUnitPrice: e.target.value })} />
                      </div>
                      <div style={styles.fieldBlock}>
                        <label style={styles.fieldLabel}>Total Quantity</label>
                        <input style={styles.input} type="number" value={expeditionForm.totalQty} onChange={(e) => setExpeditionForm({ ...expeditionForm, totalQty: e.target.value })} />
                      </div>
                      <div style={styles.fieldBlock}>
                        <label style={styles.fieldLabel}>Shipping Total (TZS)</label>
                        <input style={styles.input} type="number" value={expeditionForm.shippingTotal} onChange={(e) => setExpeditionForm({ ...expeditionForm, shippingTotal: e.target.value })} />
                      </div>
                      <div style={styles.fieldBlock}>
                        <label style={styles.fieldLabel}>Other Charges (TZS)</label>
                        <input style={styles.input} type="number" value={expeditionForm.otherCharges} onChange={(e) => setExpeditionForm({ ...expeditionForm, otherCharges: e.target.value })} />
                      </div>
                      <div style={styles.fieldBlock}>
                        <label style={styles.fieldLabel}>Local Delivery (TZS)</label>
                        <input style={styles.input} type="number" value={expeditionForm.delivery} onChange={(e) => setExpeditionForm({ ...expeditionForm, delivery: e.target.value })} />
                      </div>
                      <div style={styles.fieldBlock}>
                        <label style={styles.fieldLabel}>Estimated Arrival Days</label>
                        <input style={styles.input} type="number" value={expeditionForm.estimatedArrivalDays} onChange={(e) => setExpeditionForm({ ...expeditionForm, estimatedArrivalDays: e.target.value })} />
                      </div>
                      <div style={styles.fieldBlock}>
                        <label style={styles.fieldLabel}>Supplier Name</label>
                        <input style={styles.input} value={expeditionForm.supplierName} onChange={(e) => setExpeditionForm({ ...expeditionForm, supplierName: e.target.value })} />
                      </div>
                      <div style={styles.fieldBlock}>
                        <label style={styles.fieldLabel}>Supplier Contact</label>
                        <input style={styles.input} value={expeditionForm.supplierContact} onChange={(e) => setExpeditionForm({ ...expeditionForm, supplierContact: e.target.value })} />
                      </div>
                      <div style={styles.fieldBlock}>
                        <label style={styles.fieldLabel}>Lifecycle</label>
                        <select style={styles.input} value={expeditionForm.lifecycleStatus} onChange={(e) => setExpeditionForm({ ...expeditionForm, lifecycleStatus: e.target.value })}>
                          <option value="test">Test</option>
                          <option value="winner">Winner</option>
                          <option value="scaling">Scaling</option>
                          <option value="mature">Mature</option>
                          <option value="declining">Declining</option>
                          <option value="kill">Kill</option>
                        </select>
                      </div>
                      <div style={styles.fieldBlock}>
                        <label style={styles.fieldLabel}>Defect Rate %</label>
                        <input style={styles.input} type="number" min="0" step="0.1" value={expeditionForm.defectRate} onChange={(e) => setExpeditionForm({ ...expeditionForm, defectRate: e.target.value })} />
                      </div>
                      <div style={{ ...styles.fieldBlock, gridColumn: isCompact ? "auto" : "1 / -1" }}>
                        <label style={styles.fieldLabel}>Product Notes</label>
                        <textarea style={{ ...styles.input, minHeight: 84, resize: "vertical" }} value={expeditionForm.notes} onChange={(e) => setExpeditionForm({ ...expeditionForm, notes: e.target.value })} />
                      </div>
                    </div>

                    <div style={{ ...styles.softStat, marginTop: 16 }}>
                      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap", marginBottom: 14 }}>
                        <div>
                          <div style={{ fontSize: 11, fontWeight: 800, letterSpacing: 0.45, textTransform: "uppercase", color: textSoft }}>Product offers</div>
                          <div style={{ marginTop: 8, fontSize: 18, fontWeight: 900 }}>Quantity pricing</div>
                        </div>
                        <button style={styles.btnSecondary} onClick={addProductOfferTier}>Add Offer</button>
                      </div>
                      <div style={{ display: "grid", gap: 10 }}>
                        {(expeditionForm.offers || []).length ? (
                          expeditionForm.offers.map((offer, index) => (
                            <div key={`${offer.minQty}-${index}`} style={{ display: "grid", gridTemplateColumns: responsiveColumns("120px 1fr auto", "120px 1fr auto", "1fr"), gap: 10, alignItems: "end" }}>
                              <div style={styles.fieldBlock}>
                                <label style={styles.fieldLabel}>Min Qty</label>
                                <input style={styles.input} type="number" min="2" value={offer.minQty} onChange={(e) => updateProductOfferTier(index, "minQty", e.target.value)} />
                              </div>
                              <div style={styles.fieldBlock}>
                                <label style={styles.fieldLabel}>Offer Total Price (TZS)</label>
                                <input style={styles.input} type="number" min="0" value={offer.totalPrice} onChange={(e) => updateProductOfferTier(index, "totalPrice", e.target.value)} />
                              </div>
                              <button style={{ ...styles.btnSecondary, background: "#fef2f2", color: red, border: "1px solid #fecaca" }} onClick={() => removeProductOfferTier(index)}>
                                Remove
                              </button>
                            </div>
                          ))
                        ) : (
                          <div style={{ color: textSoft, fontSize: 14 }}>No quantity offer yet for this product.</div>
                        )}
                      </div>
                    </div>
                  </div>
                ) : null}

                <div style={{ overflowX: "auto", border: `1px solid ${cardBorder}`, borderRadius: 22, background: "linear-gradient(180deg, rgba(255,255,255,0.92), rgba(248,244,238,0.82))", boxShadow: "inset 0 1px 0 rgba(255,255,255,0.9)" }}>
                  {products.length > 0 ? (
                    <table style={{ width: "100%", borderCollapse: "separate", borderSpacing: 0 }}>
                      <thead>
                        <tr>
                          {["Product", "Source", "Sell TZS", "Offers", "Buy USD", "Qty", "Shipping TZS", "Other TZS", "Delivery TZS", "Auto Cost/Piece USD", "Total Import TZS", "Action"].map((head) => (
                            <th key={head} style={{ textAlign: "left", padding: "16px 14px", color: textSoft, fontSize: 12, fontWeight: 800, letterSpacing: 0.4, textTransform: "uppercase", borderBottom: `1px solid ${cardBorder}`, background: "rgba(247, 243, 237, 0.92)" }}>{head}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {products.map((p, index) => {
                          const qty = Number(p.totalQty || 0);
                          const totalImportCostTzs = (Number(p.purchaseUnitPrice || 0) * qty * USD_TO_TZS) + Number(p.shippingTotal || 0) + Number(p.otherCharges || 0);
                          const unitProductCostUsd = qty > 0 ? totalImportCostTzs / USD_TO_TZS / qty : 0;
                          return (
                            <tr key={p.id} style={{ background: index % 2 === 0 ? "rgba(255,255,255,0.72)" : "rgba(250,247,242,0.8)" }}>
                              <td style={{ padding: "16px 14px", borderBottom: `1px solid ${cardBorder}` }}>
                                <div style={{ fontWeight: 800 }}>{p.name}</div>
                                <div style={{ color: textSoft, fontSize: 12, marginTop: 4 }}>{p.id}</div>
                                {p.mappingCode ? (
                                  <div style={{ marginTop: 8 }}>
                                    <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
                                      <span style={{ ...styles.badge, background: "rgba(199,131,34,0.1)", color: amber, border: "1px solid rgba(199,131,34,0.2)", fontFamily: "monospace", fontSize: 13, fontWeight: 900, letterSpacing: 1 }}>
                                        {p.mappingCode}
                                      </span>
                                      <button
                                        style={{ ...styles.btnSecondary, padding: "4px 10px", fontSize: 11, borderRadius: 8 }}
                                        onClick={() => navigator.clipboard?.writeText(p.mappingCode)}
                                      >
                                        Copy
                                      </button>
                                    </div>
                                    <div style={{ marginTop: 6, fontFamily: "monospace", fontSize: 11, color: textSoft, background: "rgba(23,32,51,0.04)", padding: "4px 8px", borderRadius: 6, display: "inline-block" }}>
                                      TZ | {p.mappingCode} | TEST | COLD
                                    </div>
                                  </div>
                                ) : null}
                                <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginTop: 8 }}>
                                  <span style={{ ...styles.badge, background: "rgba(29,95,208,0.08)", color: accent, border: "1px solid rgba(29,95,208,0.12)" }}>
                                    {formatStatusLabel(p.lifecycleStatus || "test")}
                                  </span>
                                  {p.supplierName ? (
                                    <span style={{ ...styles.badge, background: "rgba(31,143,95,0.08)", color: green, border: "1px solid rgba(31,143,95,0.12)" }}>
                                      {p.supplierName}
                                    </span>
                                  ) : null}
                                </div>
                                {(p.supplierContact || Number(p.defectRate || 0) > 0 || p.notes) ? (
                                  <div style={{ color: textSoft, fontSize: 12, lineHeight: 1.5, marginTop: 8 }}>
                                    {p.supplierContact ? `Contact: ${p.supplierContact}` : ""}
                                    {p.supplierContact && Number(p.defectRate || 0) > 0 ? " | " : ""}
                                    {Number(p.defectRate || 0) > 0 ? `Defect ${Number(p.defectRate || 0).toFixed(1)}%` : ""}
                                    {p.notes ? `${p.supplierContact || Number(p.defectRate || 0) > 0 ? " | " : ""}${p.notes}` : ""}
                                  </div>
                                ) : null}
                              </td>
                              <td style={{ padding: "16px 14px", borderBottom: `1px solid ${cardBorder}` }}>
                                <span style={{ ...styles.badge, background: p.source === "dubai" ? "rgba(199,131,34,0.1)" : "rgba(29,95,208,0.08)", color: p.source === "dubai" ? amber : accent, border: p.source === "dubai" ? "1px solid rgba(199,131,34,0.14)" : "1px solid rgba(29,95,208,0.12)" }}>
                                  {p.source || "N/A"}
                                </span>
                              </td>
                              <td style={{ padding: "16px 14px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 700 }}>{formatTZS(p.sellingPrice)}</td>
                              <td style={{ padding: "16px 14px", borderBottom: `1px solid ${cardBorder}`, minWidth: 220 }}>
                                <div style={{ color: textSoft, fontSize: 13, lineHeight: 1.5 }}>{formatOffersSummary(p.offers)}</div>
                              </td>
                              <td style={{ padding: "16px 14px", borderBottom: `1px solid ${cardBorder}` }}>{formatUSD(p.purchaseUnitPrice)}</td>
                              <td style={{ padding: "16px 14px", borderBottom: `1px solid ${cardBorder}` }}>{p.totalQty}</td>
                              <td style={{ padding: "16px 14px", borderBottom: `1px solid ${cardBorder}` }}>{formatTZS(p.shippingTotal)}</td>
                              <td style={{ padding: "16px 14px", borderBottom: `1px solid ${cardBorder}` }}>{formatTZS(p.otherCharges)}</td>
                              <td style={{ padding: "16px 14px", borderBottom: `1px solid ${cardBorder}` }}>{formatTZS(p.delivery)}</td>
                              <td style={{ padding: "16px 14px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 800, color: accent }}>{formatUSD(unitProductCostUsd)}</td>
                              <td style={{ padding: "16px 14px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 700 }}>{formatTZS(totalImportCostTzs)}</td>
                              <td style={{ padding: "16px 14px", borderBottom: `1px solid ${cardBorder}` }}>
                                <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                                  <button style={{ ...styles.btnSecondary, padding: "10px 12px" }} onClick={() => startEditingProduct(p)}>
                                    Edit
                                  </button>
                                  <button style={{ ...styles.btnSecondary, background: "#fef2f2", color: red, border: "1px solid #fecaca", padding: "10px 12px" }} onClick={() => deleteProduct(p.id)}>
                                    Delete
                                  </button>
                                </div>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  ) : (
                    <div style={{ padding: 28, color: textSoft }}>No products saved yet.</div>
                  )}
                </div>

                {/* Performance dashboard stays in catalog tab, below the product table */}
                <div style={{ ...styles.softStat, marginTop: 18, padding: 18 }}>
                  <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap", marginBottom: 14 }}>
                    <div>
                      <div style={{ fontSize: 11, fontWeight: 800, letterSpacing: 0.45, textTransform: "uppercase", color: textSoft }}>Funnel stats by product</div>
                      <div style={{ marginTop: 8, fontSize: 20, fontWeight: 900 }}>Confirmation funnel by product</div>
                      <div style={{ color: textSoft, marginTop: 6, lineHeight: 1.5 }}>
                        Confirmation and delivery rates per product. For profit analysis, use the Profit Center page.
                      </div>
                    </div>
                  </div>

                  <PageDateFilterBar
                    title="Product performance range"
                    value={productPerformanceDateFilter}
                    onChange={setProductPerformanceDateFilter}
                  />

                  <div
                    style={{
                      display: "grid",
                      gridTemplateColumns: responsiveColumns("repeat(8, minmax(0, 1fr))", "repeat(2, minmax(0, 1fr))", "1fr"),
                      gap: 12,
                      margin: "16px 0 18px",
                    }}
                  >
                    <MiniStat label="Revenue" value={formatTZS(productPerformanceDateSummary.totalRevenue)} tone="green" />
                    <MiniStat label="Ads Spend" value={formatTZS(productPerformanceDateSummary.totalAdsSpend)} tone="amber" />
                    <MiniStat label="Profit" value={formatTZS(productPerformanceDateSummary.totalProfit)} tone={productPerformanceDateSummary.totalProfit >= 0 ? "green" : "red"} />
                    <MiniStat label="Leads" value={formatInteger(productPerformanceDateSummary.totalLeads)} tone="blue" />
                    <MiniStat label="Confirmed" value={formatInteger(productPerformanceDateSummary.totalConfirmedOrders)} tone="amber" />
                    <MiniStat label="Delivered" value={formatInteger(productPerformanceDateSummary.totalDeliveredOrders)} tone="green" />
                    <MiniStat label="Confirmation Rate" value={`${productPerformanceDateSummary.globalConfirmationRate.toFixed(1)}%`} tone="blue" />
                    <MiniStat label="Delivery Rate" value={`${productPerformanceDateSummary.globalDeliveryRate.toFixed(1)}%`} tone="green" />
                  </div>

                  <div
                    style={{
                      display: "grid",
                      gridTemplateColumns: responsiveColumns("repeat(5, minmax(0, 1fr))", "repeat(2, minmax(0, 1fr))", "1fr"),
                      gap: 12,
                      marginBottom: 16,
                    }}
                  >
                    <div style={styles.fieldBlock}>
                      <div style={styles.fieldLabel}>Winner Profit Threshold</div>
                      <input
                        style={styles.input}
                        type="number"
                        min="0"
                        value={situationData.productWinnerThresholdTzs || ""}
                        onChange={(e) => setSituationData((prev) => ({ ...prev, productWinnerThresholdTzs: Math.max(0, parseLooseNumber(e.target.value)) }))}
                      />
                    </div>
                    <div style={styles.fieldBlock}>
                      <div style={styles.fieldLabel}>Minimum Stock Threshold</div>
                      <input
                        style={styles.input}
                        type="number"
                        min="0"
                        value={situationData.productAlertThresholds?.minStockQuantity || ""}
                        onChange={(e) => updateProductAlertThreshold("minStockQuantity", e.target.value)}
                      />
                    </div>
                    <div style={styles.fieldBlock}>
                      <div style={styles.fieldLabel}>Minimum Delivery Rate %</div>
                      <input
                        style={styles.input}
                        type="number"
                        min="0"
                        max="100"
                        value={situationData.productAlertThresholds?.minDeliveryRatePct || ""}
                        onChange={(e) => updateProductAlertThreshold("minDeliveryRatePct", e.target.value)}
                      />
                    </div>
                    <div style={styles.fieldBlock}>
                      <div style={styles.fieldLabel}>High Ads Spend Threshold</div>
                      <input
                        style={styles.input}
                        type="number"
                        min="0"
                        value={situationData.productAlertThresholds?.highAdsSpendTzs || ""}
                        onChange={(e) => updateProductAlertThreshold("highAdsSpendTzs", e.target.value)}
                      />
                    </div>
                    <div style={styles.fieldBlock}>
                      <div style={styles.fieldLabel}>Low Delivered Orders Limit</div>
                      <input
                        style={styles.input}
                        type="number"
                        min="0"
                        value={situationData.productAlertThresholds?.lowDeliveredOrders || ""}
                        onChange={(e) => updateProductAlertThreshold("lowDeliveredOrders", e.target.value)}
                      />
                    </div>
                  </div>

                  <div style={{ overflowX: "auto", border: `1px solid ${cardBorder}`, borderRadius: 20, background: "rgba(255,255,255,0.84)" }}>
                    <table style={{ width: "100%", borderCollapse: "separate", borderSpacing: 0 }}>
                      <thead>
                        <tr>
                          {["Product", "Status", "Alerts", "Stock Qty", "Units Sold", "Total Revenue", "Ads Spend", "Leads", "Confirmed Orders", "Delivered Orders", "CPA", "Confirmation Rate", "Delivery Rate", "Total Product Cost", "Total Delivery Cost", "Profit", "Profit Margin"].map((head) => (
                            <th
                              key={head}
                              style={{
                                textAlign: "left",
                                padding: "14px 12px",
                                color: textSoft,
                                fontSize: 12,
                                fontWeight: 800,
                                letterSpacing: 0.4,
                                textTransform: "uppercase",
                                borderBottom: `1px solid ${cardBorder}`,
                                background: "rgba(247, 243, 237, 0.92)",
                                whiteSpace: "nowrap",
                              }}
                            >
                              {head}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {filteredProductPerformanceRows.map((row, index) => (
                          <tr key={`performance-${row.id}`} style={{ background: index % 2 === 0 ? "rgba(255,255,255,0.72)" : "rgba(250,247,242,0.8)" }}>
                            <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}`, minWidth: 220 }}>
                              <div style={{ fontWeight: 800 }}>{row.name}</div>
                              <div style={{ color: textSoft, fontSize: 12, marginTop: 4 }}>{row.id}</div>
                            </td>
                            <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}`, whiteSpace: "nowrap" }}>
                              <span style={getDecisionStyle(row.performanceStatus === "WINNER" ? "OK" : row.performanceStatus === "TESTING" ? "WATCH" : "KILL")}>
                                {row.performanceStatus}
                              </span>
                            </td>
                            <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}`, minWidth: 260 }}>
                              <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                                {row.productAlerts.length ? row.productAlerts.map((alert) => (
                                  <span key={`${row.id}-${alert.key}`} style={getAlertBadgeStyle(alert.tone)}>
                                    {alert.message}
                                  </span>
                                )) : <span style={getAlertBadgeStyle("success")}>No active alerts</span>}
                              </div>
                            </td>
                            <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 700 }}>{row.stockQuantity}</td>
                            <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>{row.totalUnitsSold}</td>
                            <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 700 }}>{formatTZS(row.totalRevenue)}</td>
                            <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}`, minWidth: 150 }}>
                              <input
                                style={styles.input}
                                type="number"
                                min="0"
                                value={row.manualAdsSpendTzs || ""}
                                placeholder={String(Number(row.dashboardAdsSpendTzs || 0))}
                                onChange={(e) => updateProductAdInput(row.id, "manualAdsSpendTzs", e.target.value)}
                              />
                            </td>
                            <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}`, minWidth: 120 }}>
                              <input
                                style={styles.input}
                                type="number"
                                min="0"
                                value={row.funnelLeads || ""}
                                placeholder={String(Number(row.effectiveLeads || 0))}
                                onChange={(e) => updateProductAdInput(row.id, "leads", e.target.value)}
                              />
                            </td>
                            <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}`, minWidth: 140 }}>
                              <input
                                style={styles.input}
                                type="number"
                                min="0"
                                value={row.funnelConfirmedOrders || ""}
                                placeholder={String(Number(row.effectiveConfirmedOrders || 0))}
                                onChange={(e) => updateProductAdInput(row.id, "confirmedOrders", e.target.value)}
                              />
                            </td>
                            <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}`, minWidth: 140 }}>
                              <input
                                style={styles.input}
                                type="number"
                                min="0"
                                value={row.funnelDeliveredOrders || ""}
                                placeholder={String(Number(row.effectiveDeliveredOrders || 0))}
                                onChange={(e) => updateProductAdInput(row.id, "deliveredOrders", e.target.value)}
                              />
                            </td>
                            <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 700 }}>
                              {row.effectiveDeliveredOrders > 0 ? formatTZS(row.dashboardCpaTzs) : "N/A"}
                            </td>
                            <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>{row.dashboardConfirmationRate.toFixed(1)}%</td>
                            <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>{row.dashboardDeliveryRate.toFixed(1)}%</td>
                            <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>{formatTZS(row.dashboardTotalProductCostTzs)}</td>
                            <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>{formatTZS(row.dashboardTotalDeliveryCostTzs)}</td>
                            <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 800, color: row.dashboardProfitTzs >= 0 ? green : red }}>
                              {formatTZS(row.dashboardProfitTzs)}
                            </td>
                            <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 700 }}>{row.dashboardProfitMargin.toFixed(1)}%</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
              </div>
              {/* end catalog tab */}

              {/* ===== TAB: STOCK PURCHASES ===== */}
              <div style={{ display: stockTab === "purchases" ? "grid" : "none", gap: 16 }}>
              <div style={{ ...styles.card, padding: 22 }}>
                <div style={styles.sectionHeader}>
                  <div>
                    <div style={styles.sectionEyebrow}>Stock purchases</div>
                    <div style={{ fontSize: 24, fontWeight: 900, marginTop: 8 }}>{editingPurchaseId ? "Edit purchase" : "New stock purchase"}</div>
                    <div style={{ color: textSoft, marginTop: 6, lineHeight: 1.6 }}>Record a new stock order. Only received purchases count toward available stock.</div>
                  </div>
                  {editingPurchaseId ? (
                    <button style={styles.btnSecondary} onClick={() => { setEditingPurchaseId(null); setPurchaseForm({ product_id: "", quantity_ordered: "", source_country: "dubai", supplier_name: "", purchase_date: "", expected_arrival_date: "", usable_stock_date: "", buy_price_per_unit_usd: "", shipping_cost_usd: "", sourcing_cost_usd: "", other_charges_tsh: "", quantity_received: "", status: "ordered", notes: "" }); }}>Cancel</button>
                  ) : null}
                </div>
                <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(3, minmax(0, 1fr))", "repeat(2, minmax(0, 1fr))", "1fr"), gap: 14 }}>
                  <div style={styles.fieldBlock}>
                    <label style={styles.fieldLabel}>Product</label>
                    <select style={styles.input} value={purchaseForm.product_id} onChange={(e) => setPurchaseForm((f) => ({ ...f, product_id: e.target.value }))}>
                      <option value="">— Select product —</option>
                      {products.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
                    </select>
                  </div>
                  <div style={styles.fieldBlock}>
                    <label style={styles.fieldLabel}>Qty ordered</label>
                    <input style={styles.input} type="number" min="0" value={purchaseForm.quantity_ordered} onChange={(e) => setPurchaseForm((f) => ({ ...f, quantity_ordered: e.target.value }))} placeholder="100" />
                  </div>
                  <div style={styles.fieldBlock}>
                    <label style={styles.fieldLabel}>Source country</label>
                    <select style={styles.input} value={purchaseForm.source_country} onChange={(e) => setPurchaseForm((f) => ({ ...f, source_country: e.target.value }))}>
                      <option value="dubai">Dubai</option>
                      <option value="china">China</option>
                      <option value="other">Other</option>
                    </select>
                  </div>
                  <div style={styles.fieldBlock}>
                    <label style={styles.fieldLabel}>Buy price / unit (USD)</label>
                    <input style={styles.input} type="number" min="0" step="0.01" value={purchaseForm.buy_price_per_unit_usd} onChange={(e) => setPurchaseForm((f) => ({ ...f, buy_price_per_unit_usd: e.target.value }))} placeholder="5.00" />
                  </div>
                  <div style={styles.fieldBlock}>
                    <label style={styles.fieldLabel}>Shipping cost (USD)</label>
                    <input style={styles.input} type="number" min="0" step="0.01" value={purchaseForm.shipping_cost_usd} onChange={(e) => setPurchaseForm((f) => ({ ...f, shipping_cost_usd: e.target.value }))} placeholder="0" />
                  </div>
                  <div style={styles.fieldBlock}>
                    <label style={styles.fieldLabel}>Sourcing cost (USD)</label>
                    <input style={styles.input} type="number" min="0" step="0.01" value={purchaseForm.sourcing_cost_usd} onChange={(e) => setPurchaseForm((f) => ({ ...f, sourcing_cost_usd: e.target.value }))} placeholder="0" />
                  </div>
                  <div style={styles.fieldBlock}>
                    <label style={styles.fieldLabel}>Other charges (TZS)</label>
                    <input style={styles.input} type="number" min="0" value={purchaseForm.other_charges_tsh} onChange={(e) => setPurchaseForm((f) => ({ ...f, other_charges_tsh: e.target.value }))} placeholder="0" />
                  </div>
                  <div style={styles.fieldBlock}>
                    <label style={styles.fieldLabel}>Supplier name</label>
                    <input style={styles.input} value={purchaseForm.supplier_name} onChange={(e) => setPurchaseForm((f) => ({ ...f, supplier_name: e.target.value }))} placeholder="Supplier or agent" />
                  </div>
                  <div style={styles.fieldBlock}>
                    <label style={styles.fieldLabel}>Purchase date</label>
                    <input style={styles.input} type="date" value={purchaseForm.purchase_date} onChange={(e) => setPurchaseForm((f) => ({ ...f, purchase_date: e.target.value }))} />
                  </div>
                  <div style={styles.fieldBlock}>
                    <label style={styles.fieldLabel}>Expected arrival</label>
                    <input style={styles.input} type="date" value={purchaseForm.expected_arrival_date} onChange={(e) => setPurchaseForm((f) => ({ ...f, expected_arrival_date: e.target.value }))} />
                  </div>
                  <div style={styles.fieldBlock}>
                    <label style={styles.fieldLabel}>Usable stock date</label>
                    <input style={styles.input} type="date" value={purchaseForm.usable_stock_date} onChange={(e) => setPurchaseForm((f) => ({ ...f, usable_stock_date: e.target.value }))} />
                  </div>
                  <div style={styles.fieldBlock}>
                    <label style={styles.fieldLabel}>Status</label>
                    <select style={styles.input} value={purchaseForm.status} onChange={(e) => setPurchaseForm((f) => ({ ...f, status: e.target.value }))}>
                      <option value="ordered">Ordered</option>
                      <option value="in_transit">In Transit</option>
                      <option value="arrived">Arrived</option>
                      <option value="received">Received</option>
                      <option value="delayed">Delayed</option>
                      <option value="cancelled">Cancelled</option>
                    </select>
                  </div>
                  <div style={styles.fieldBlock}>
                    <label style={styles.fieldLabel}>Qty received</label>
                    <input style={styles.input} type="number" min="0" value={purchaseForm.quantity_received} onChange={(e) => setPurchaseForm((f) => ({ ...f, quantity_received: e.target.value }))} placeholder="0" />
                  </div>
                  <div style={styles.fieldBlock}>
                    <label style={styles.fieldLabel}>Notes</label>
                    <input style={styles.input} value={purchaseForm.notes} onChange={(e) => setPurchaseForm((f) => ({ ...f, notes: e.target.value }))} placeholder="Optional notes" />
                  </div>
                </div>
                {purchaseForm.product_id && Number(purchaseForm.quantity_ordered) > 0 ? (
                  <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(3, minmax(0, 1fr))", "1fr 1fr", "1fr"), gap: 12, marginTop: 14, padding: 16, borderRadius: 16, background: "rgba(247,243,237,0.85)", border: `1px solid ${cardBorder}` }}>
                    {(() => {
                      const qty = Number(purchaseForm.quantity_ordered) || 0;
                      const buy = Number(purchaseForm.buy_price_per_unit_usd) || 0;
                      const ship = Number(purchaseForm.shipping_cost_usd) || 0;
                      const src = Number(purchaseForm.sourcing_cost_usd) || 0;
                      const otherUsd = (Number(purchaseForm.other_charges_tsh) || 0) / Number(serviceForm?.exchangeRate || USD_TO_TZS);
                      const totalBuy = qty * buy;
                      const totalLanded = totalBuy + ship + src + otherUsd;
                      const perUnit = qty > 0 ? totalLanded / qty : 0;
                      return (
                        <>
                          <MiniStat label="Total buy cost" value={formatUSD(totalBuy)} tone="blue" />
                          <MiniStat label="Total landed cost" value={formatUSD(totalLanded)} tone="amber" />
                          <MiniStat label="Landed cost / unit" value={perUnit > 0 ? formatUSD(perUnit) : "N/A"} tone="green" />
                        </>
                      );
                    })()}
                  </div>
                ) : null}
                <div style={{ marginTop: 16, display: "flex", gap: 10 }}>
                  <button style={styles.btnPrimary} onClick={savePurchase}>{editingPurchaseId ? "Update Purchase" : "Save Purchase"}</button>
                </div>
              </div>
              <div style={{ ...styles.card, padding: 22 }}>
                <div style={{ fontSize: 18, fontWeight: 800, marginBottom: 14 }}>All stock purchases</div>
                {stockPurchases.length === 0 ? (
                  <div style={{ color: textSoft }}>No stock purchases recorded yet.</div>
                ) : (
                  <div style={{ overflowX: "auto" }}>
                    <table style={{ width: "100%", borderCollapse: "separate", borderSpacing: 0 }}>
                      <thead>
                        <tr>{["Product", "Source", "Qty Ordered", "Qty Received", "Buy/Unit (USD)", "Landed/Unit (USD)", "Total Landed (USD)", "Purchase Date", "Expected Arrival", "Status", "Supplier", "Actions"].map((h) => (
                          <th key={h} style={{ textAlign: "left", padding: "12px 10px", color: textSoft, fontSize: 12, fontWeight: 800, letterSpacing: 0.4, textTransform: "uppercase", borderBottom: `1px solid ${cardBorder}`, background: "rgba(247,243,237,0.92)", whiteSpace: "nowrap" }}>{h}</th>
                        ))}</tr>
                      </thead>
                      <tbody>
                        {stockPurchases.map((pur, idx) => {
                          const prod = products.find((p) => p.id === pur.product_id);
                          return (
                            <tr key={pur.id} style={{ background: idx % 2 === 0 ? "rgba(255,255,255,0.72)" : "rgba(250,247,242,0.8)" }}>
                              <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 700 }}>{prod?.name || pur.product_id}</td>
                              <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}` }}>{pur.source_country || "—"}</td>
                              <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}` }}>{pur.quantity_ordered}</td>
                              <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}` }}>{pur.quantity_received || 0}</td>
                              <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}` }}>{formatUSD(pur.buy_price_per_unit_usd)}</td>
                              <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 700, color: accent }}>{formatUSD(pur.landed_cost_per_unit_usd)}</td>
                              <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}` }}>{formatUSD(pur.total_landed_cost_usd)}</td>
                              <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}` }}>{pur.purchase_date || "—"}</td>
                              <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}` }}>{pur.expected_arrival_date || "—"}</td>
                              <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}` }}><span style={getDecisionStyle(pur.status)}>{pur.status}</span></td>
                              <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}` }}>{pur.supplier_name || "—"}</td>
                              <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}`, whiteSpace: "nowrap" }}>
                                <div style={{ display: "inline-flex", gap: 6 }}>
                                  <button style={{ ...styles.btnSecondary, padding: "6px 10px", fontSize: 12 }} onClick={() => { setEditingPurchaseId(pur.id); setPurchaseForm({ product_id: pur.product_id, quantity_ordered: String(pur.quantity_ordered || ""), source_country: pur.source_country || "dubai", supplier_name: pur.supplier_name || "", purchase_date: pur.purchase_date || "", expected_arrival_date: pur.expected_arrival_date || "", usable_stock_date: pur.usable_stock_date || "", buy_price_per_unit_usd: String(pur.buy_price_per_unit_usd || ""), shipping_cost_usd: String(pur.shipping_cost_usd || ""), sourcing_cost_usd: String(pur.sourcing_cost_usd || ""), other_charges_tsh: String(pur.other_charges_tsh || ""), quantity_received: String(pur.quantity_received || ""), status: pur.status || "ordered", notes: pur.notes || "" }); }}>Edit</button>
                                  <button style={{ ...styles.btnSecondary, background: "#fef2f2", color: red, border: "1px solid #fecaca", padding: "6px 10px", fontSize: 12 }} onClick={() => deletePurchase(pur.id)}>Delete</button>
                                </div>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
              </div>
              {/* end purchases tab */}

              {/* ===== TAB: INCOMING SHIPMENTS ===== */}
              <div style={{ display: stockTab === "incoming" ? "grid" : "none", gap: 16 }}>
              <div style={{ ...styles.card, padding: 22 }}>
                <div style={styles.sectionHeader}>
                  <div>
                    <div style={styles.sectionEyebrow}>Incoming shipments</div>
                    <div style={{ fontSize: 24, fontWeight: 900, marginTop: 8 }}>Purchases not yet received</div>
                    <div style={{ color: textSoft, marginTop: 6, lineHeight: 1.6 }}>Ordered and in-transit stock does not count toward available stock until marked as Received.</div>
                  </div>
                </div>
                {(() => {
                  const today = getTodayString();
                  const incoming = stockPurchases.filter((p) => !["received", "cancelled"].includes(p.status));
                  if (incoming.length === 0) return <div style={{ color: textSoft, padding: "20px 0" }}>No incoming shipments.</div>;
                  return (
                    <div style={{ overflowX: "auto" }}>
                      <table style={{ width: "100%", borderCollapse: "separate", borderSpacing: 0 }}>
                        <thead>
                          <tr>{["Product", "Source", "Ordered", "Received", "Remaining", "Landed/Unit", "Expected Arrival", "Actual Arrival", "Received Date", "Delay", "Status", "Actions"].map((h) => (
                            <th key={h} style={{ textAlign: "left", padding: "12px 10px", color: textSoft, fontSize: 12, fontWeight: 800, letterSpacing: 0.4, textTransform: "uppercase", borderBottom: `1px solid ${cardBorder}`, background: "rgba(247,243,237,0.92)", whiteSpace: "nowrap" }}>{h}</th>
                          ))}</tr>
                        </thead>
                        <tbody>
                          {incoming.map((pur, idx) => {
                            const prod = products.find((p) => p.id === pur.product_id);
                            const qtyReceived = Math.max(0, Number(pur.quantity_received) || 0);
                            const qtyOrdered = Math.max(0, Number(pur.quantity_ordered) || 0);
                            const remaining = qtyOrdered - qtyReceived;
                            const isDelayed = pur.expected_arrival_date && pur.expected_arrival_date < today && !["arrived", "partially_received"].includes(pur.status);
                            const isReceiving = receiveStockInput.purchaseId === pur.id;
                            return (
                              <tr key={pur.id} style={{ background: idx % 2 === 0 ? "rgba(255,255,255,0.72)" : "rgba(250,247,242,0.8)" }}>
                                <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 700 }}>{prod?.name || pur.product_id}</td>
                                <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}` }}>{pur.source_country || "—"}</td>
                                <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}` }}>{qtyOrdered}</td>
                                <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}` }}>{qtyReceived}</td>
                                <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 700, color: remaining <= 0 ? textSoft : textMain }}>{remaining}</td>
                                <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 700 }}>{formatUSD(pur.landed_cost_per_unit_usd)}</td>
                                <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}` }}>{pur.expected_arrival_date || "—"}</td>
                                <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}` }}>{pur.actual_arrival_date ? <span style={{ display: "inline-flex", alignItems: "center", gap: 4 }}>{pur.actual_arrival_date}{pur.is_early_arrival ? <span style={{ fontSize: 10, background: "#d1fae5", color: "#065f46", borderRadius: 4, padding: "1px 5px", fontWeight: 700 }}>EARLY</span> : null}</span> : "—"}</td>
                                <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}` }}>{pur.received_date || "—"}</td>
                                <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}` }}>{isDelayed ? <span style={getDecisionStyle("KILL")}>Delayed</span> : <span style={getDecisionStyle("OK")}>On time</span>}</td>
                                <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}` }}>
                                  <span style={getDecisionStyle(pur.status)}>{pur.status.replace(/_/g, " ")}</span>
                                </td>
                                <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}`, minWidth: 240 }}>
                                  {isReceiving ? (
                                    <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                                      <div style={{ fontSize: 11, color: textSoft }}>Remaining: <strong>{remaining}</strong></div>
                                      <input
                                        type="number"
                                        min={1}
                                        max={remaining}
                                        placeholder={`Qty (max ${remaining})`}
                                        value={receiveStockInput.qty}
                                        onChange={(e) => setReceiveStockInput((s) => ({ ...s, qty: e.target.value }))}
                                        style={{ ...styles.input, width: 90, padding: "4px 8px", fontSize: 12 }}
                                      />
                                      <input
                                        type="text"
                                        placeholder="Notes (optional)"
                                        value={receiveStockInput.notes}
                                        onChange={(e) => setReceiveStockInput((s) => ({ ...s, notes: e.target.value }))}
                                        style={{ ...styles.input, padding: "4px 8px", fontSize: 12 }}
                                      />
                                      {(() => {
                                        const enteredQty = Number(receiveStockInput.qty) || 0;
                                        const overReceive = enteredQty > remaining;
                                        return overReceive ? <div style={{ fontSize: 11, color: red, fontWeight: 700 }}>Cannot receive more than remaining quantity ({remaining})</div> : null;
                                      })()}
                                      <div style={{ display: "flex", gap: 6 }}>
                                        <button style={{ ...styles.btnPrimary, padding: "4px 10px", fontSize: 11 }} onClick={() => {
                                          const enteredQty = Number(receiveStockInput.qty) || 0;
                                          if (enteredQty <= 0 || enteredQty > remaining) return;
                                          receiveStockNow(pur.id, enteredQty, receiveStockInput.notes);
                                          setReceiveStockInput({ purchaseId: null, qty: "", notes: "" });
                                        }}>Confirm</button>
                                        <button style={{ ...styles.btnSecondary, padding: "4px 10px", fontSize: 11 }} onClick={() => setReceiveStockInput({ purchaseId: null, qty: "", notes: "" })}>Cancel</button>
                                      </div>
                                    </div>
                                  ) : (
                                    <div style={{ display: "inline-flex", gap: 6, flexWrap: "wrap" }}>
                                      {!["arrived", "partially_received"].includes(pur.status) ? (
                                        <button style={{ ...styles.btnSecondary, padding: "5px 8px", fontSize: 11, background: "#d1fae5", color: "#065f46", border: "1px solid #6ee7b7" }} onClick={() => markArrivedEarly(pur.id)}>Arrived Early</button>
                                      ) : null}
                                      {remaining > 0 ? (
                                        <button style={{ ...styles.btnSecondary, padding: "5px 8px", fontSize: 11 }} onClick={() => setReceiveStockInput({ purchaseId: pur.id, qty: String(remaining), notes: "" })}>Receive Stock</button>
                                      ) : null}
                                      <button style={{ ...styles.btnSecondary, padding: "5px 8px", fontSize: 11 }} onClick={() => updatePurchaseStatus(pur.id, "delayed", pur.quantity_received)}>Mark Delayed</button>
                                      <button style={{ ...styles.btnSecondary, background: "#fef2f2", color: red, border: "1px solid #fecaca", padding: "5px 8px", fontSize: 11 }} onClick={() => updatePurchaseStatus(pur.id, "cancelled", 0)}>Cancel</button>
                                    </div>
                                  )}
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  );
                })()}
              </div>
              </div>
              {/* end incoming tab */}

              {/* ===== TAB: STOCK OVERVIEW ===== */}
              <div style={{ display: stockTab === "overview" ? "grid" : "none", gap: 16 }}>
              <div style={{ ...styles.card, padding: 22 }}>
                <div style={styles.sectionHeader}>
                  <div>
                    <div style={styles.sectionEyebrow}>Stock overview</div>
                    <div style={{ fontSize: 24, fontWeight: 900, marginTop: 8 }}>Real-time stock by product</div>
                    <div style={{ color: textSoft, marginTop: 6, lineHeight: 1.6 }}>In Stock = Accepted − Out Delivered − Delivered − Damaged. Incoming = ordered / in-transit only.</div>
                  </div>
                </div>
                <div style={{ overflowX: "auto" }}>
                  <table style={{ width: "100%", borderCollapse: "separate", borderSpacing: 0 }}>
                    <thead>
                      <tr>{["Product", "Accepted", "In Stock", "Out Delivered", "Delivered", "Damaged", "Incoming", "Stock Value USD", "Landed Cost/Unit", "Sales/Day", "Days to Stockout", "Reorder Status", "Actions"].map((h) => (
                        <th key={h} style={{ textAlign: "left", padding: "12px 10px", color: textSoft, fontSize: 12, fontWeight: 800, letterSpacing: 0.4, textTransform: "uppercase", borderBottom: `1px solid ${cardBorder}`, background: "rgba(247,243,237,0.92)", whiteSpace: "nowrap" }}>{h}</th>
                      ))}</tr>
                    </thead>
                    <tbody>
                      {stockForecastRows.map((row, idx) => {
                        const minStock = Number(situationData?.productAlertThresholds?.minStockQuantity ?? 3);
                        return (
                          <tr key={row.id} style={{ background: idx % 2 === 0 ? "rgba(255,255,255,0.72)" : "rgba(250,247,242,0.8)" }}>
                            <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 700 }}>{row.name}</td>
                            <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}` }}>{row.acceptedStock || 0}</td>
                            <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 800, color: row.availableStock <= 0 ? red : row.availableStock <= minStock ? amber : textMain }}>{row.availableStock}</td>
                            <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}` }}>{row.reservedStock || 0}</td>
                            <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}` }}>{row.deliveredStock || 0}</td>
                            <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}`, color: (row.damagedStock || 0) > 0 ? amber : textMain }}>{row.damagedStock || 0}</td>
                            <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}` }}>{row.incomingStock || 0}</td>
                            <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 700 }}>{formatUSD(row.stockValueUsd)}</td>
                            <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}` }}>{row.unitProductCost > 0 ? formatUSD(row.unitProductCost) : "—"}</td>
                            <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}` }}>{row.salesPerDay.toFixed(1)}</td>
                            <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}` }}>{row.daysUntilStockout != null ? `${Math.max(1, Math.round(row.daysUntilStockout))}d` : "—"}</td>
                            <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}` }}><span style={getDecisionStyle(row.reorderStatus)}>{row.reorderStatus}</span></td>
                            <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}`, whiteSpace: "nowrap" }}>
                              <div style={{ display: "inline-flex", gap: 6 }}>
                                <button style={{ ...styles.btnSecondary, padding: "5px 8px", fontSize: 11 }} onClick={() => { setStockTab("purchases"); setPurchaseForm((f) => ({ ...f, product_id: row.id })); }}>Restock</button>
                                <button style={{ ...styles.btnSecondary, padding: "5px 8px", fontSize: 11 }} onClick={() => setStockTab("movements")}>Movements</button>
                              </div>
                            </td>
                          </tr>
                        );
                      })}
                      {stockForecastRows.length === 0 ? <tr><td colSpan={13} style={{ padding: 20, color: textSoft }}>No products yet.</td></tr> : null}
                    </tbody>
                  </table>
                </div>
              </div>
              </div>
              {/* end overview tab */}

              {/* ===== TAB: STOCK MOVEMENTS ===== */}
              <div style={{ display: stockTab === "movements" ? "grid" : "none", gap: 16 }}>
              <div style={{ ...styles.card, padding: 22 }}>
                <div style={styles.sectionHeader}>
                  <div>
                    <div style={styles.sectionEyebrow}>Stock movements</div>
                    <div style={{ fontSize: 24, fontWeight: 900, marginTop: 8 }}>Movement log</div>
                    <div style={{ color: textSoft, marginTop: 6, lineHeight: 1.6 }}>Every stock change is recorded here. Manual adjustments require a reason.</div>
                  </div>
                </div>
                <div style={{ ...styles.softStat, marginBottom: 20, padding: 16 }}>
                  <div style={{ fontWeight: 800, marginBottom: 12 }}>Manual stock adjustment</div>
                  <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("1fr 1fr 1fr 1fr auto", "1fr 1fr", "1fr"), gap: 12, alignItems: "end" }}>
                    <div style={styles.fieldBlock}>
                      <label style={styles.fieldLabel}>Product</label>
                      <select style={styles.input} value={manualAdjForm.product_id} onChange={(e) => setManualAdjForm((f) => ({ ...f, product_id: e.target.value }))}>
                        <option value="">— Select product —</option>
                        {products.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
                      </select>
                    </div>
                    <div style={styles.fieldBlock}>
                      <label style={styles.fieldLabel}>Qty change (+/-)</label>
                      <input style={styles.input} type="number" value={manualAdjForm.quantity_change} onChange={(e) => setManualAdjForm((f) => ({ ...f, quantity_change: e.target.value }))} placeholder="e.g. -5 or +10" />
                    </div>
                    <div style={styles.fieldBlock}>
                      <label style={styles.fieldLabel}>Reason</label>
                      <select style={styles.input} value={manualAdjForm.reason} onChange={(e) => setManualAdjForm((f) => ({ ...f, reason: e.target.value }))}>
                        <option value="stock_count_correction">Stock count correction</option>
                        <option value="damaged_item">Damaged item</option>
                        <option value="lost_item">Lost item</option>
                        <option value="warehouse_correction">Warehouse correction</option>
                        <option value="other">Other</option>
                      </select>
                    </div>
                    <div style={styles.fieldBlock}>
                      <label style={styles.fieldLabel}>Note</label>
                      <input style={styles.input} value={manualAdjForm.note} onChange={(e) => setManualAdjForm((f) => ({ ...f, note: e.target.value }))} placeholder="Required note" />
                    </div>
                    <button style={styles.btnPrimary} onClick={saveManualAdjustment}>Apply</button>
                  </div>
                </div>
                {stockMovements.length === 0 ? (
                  <div style={{ color: textSoft }}>No movements recorded yet.</div>
                ) : (
                  <div style={{ overflowX: "auto" }}>
                    <table style={{ width: "100%", borderCollapse: "separate", borderSpacing: 0 }}>
                      <thead>
                        <tr>{["Date", "Product", "Type", "Qty Change", "Reference", "Note"].map((h) => (
                          <th key={h} style={{ textAlign: "left", padding: "12px 10px", color: textSoft, fontSize: 12, fontWeight: 800, letterSpacing: 0.4, textTransform: "uppercase", borderBottom: `1px solid ${cardBorder}`, background: "rgba(247,243,237,0.92)", whiteSpace: "nowrap" }}>{h}</th>
                        ))}</tr>
                      </thead>
                      <tbody>
                        {[...stockMovements].reverse().map((mov, idx) => {
                          const prod = products.find((p) => p.id === mov.product_id);
                          return (
                            <tr key={mov.movement_id} style={{ background: idx % 2 === 0 ? "rgba(255,255,255,0.72)" : "rgba(250,247,242,0.8)" }}>
                              <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}` }}>{mov.date}</td>
                              <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 700 }}>{prod?.name || mov.product_id}</td>
                              <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}` }}><span style={{ ...styles.badge, background: "rgba(29,95,208,0.08)", color: accent }}>{mov.type?.replace(/_/g, " ")}</span></td>
                              <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 700, color: Number(mov.quantity_change) >= 0 ? green : red }}>{Number(mov.quantity_change) >= 0 ? `+${mov.quantity_change}` : mov.quantity_change}</td>
                              <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}`, fontSize: 12, color: textSoft }}>{mov.source_reference || "—"}</td>
                              <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}`, fontSize: 13 }}>{mov.note || "—"}</td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
              </div>
              {/* end movements tab */}

              {/* ===== TAB: STOCK ALERTS ===== */}
              <div style={{ display: stockTab === "alerts" ? "grid" : "none", gap: 16 }}>
              <div style={{ ...styles.card, padding: 22 }}>
                <div style={styles.sectionHeader}>
                  <div>
                    <div style={styles.sectionEyebrow}>Stock alerts</div>
                    <div style={{ fontSize: 24, fontWeight: 900, marginTop: 8 }}>Active stock issues</div>
                  </div>
                </div>
                {stockAlerts.length === 0 ? (
                  <div style={{ color: textSoft, padding: "16px 0" }}>No stock alerts. Everything looks good.</div>
                ) : (
                  <div style={{ display: "grid", gap: 10 }}>
                    {stockAlerts.map((alert, idx) => (
                      <div key={idx} style={{ padding: "14px 18px", borderRadius: 16, border: `1px solid ${alert.severity === "critical" ? "#fecaca" : alert.severity === "warning" ? "rgba(199,131,34,0.25)" : cardBorder}`, background: alert.severity === "critical" ? "#fef2f2" : alert.severity === "warning" ? "rgba(199,131,34,0.06)" : "rgba(255,255,255,0.85)", display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 12, flexWrap: "wrap" }}>
                        <div>
                          <div style={{ fontWeight: 800, color: alert.severity === "critical" ? red : alert.severity === "warning" ? amber : textSoft }}>{alert.productName}</div>
                          <div style={{ fontSize: 13, color: textMain, marginTop: 4 }}>{alert.message}</div>
                        </div>
                        <span style={{ ...styles.badge, background: alert.severity === "critical" ? "rgba(220,38,38,0.1)" : "rgba(199,131,34,0.1)", color: alert.severity === "critical" ? red : amber }}>{alert.type?.replace(/_/g, " ")}</span>
                      </div>
                    ))}
                  </div>
                )}
              </div>
              </div>
              {/* end alerts tab */}

              {/* ===== TAB: STOCK AUDIT ===== */}
              <div style={{ display: stockTab === "audit" ? "grid" : "none", gap: 16 }}>
              <div style={{ ...styles.card, padding: 22 }}>
                <div style={styles.sectionHeader}>
                  <div>
                    <div style={styles.sectionEyebrow}>Stock audit</div>
                    <div style={{ fontSize: 24, fontWeight: 900, marginTop: 8 }}>Full inventory audit</div>
                    <div style={{ color: textSoft, marginTop: 6, lineHeight: 1.6 }}>Cross-check all stock figures. Identify missing costs, negative stock and inconsistencies.</div>
                  </div>
                </div>
                <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(4, minmax(0, 1fr))", "1fr 1fr", "1fr"), gap: 12, marginBottom: 20 }}>
                  <MiniStat label="Products" value={stockAuditData.totalProducts} tone="blue" />
                  <MiniStat label="Total accepted" value={stockAuditData.totalAccepted} tone="blue" />
                  <MiniStat label="Total in stock" value={stockAuditData.totalAvailable} tone="green" />
                  <MiniStat label="Total out delivered" value={stockAuditData.totalOutDelivered} tone="amber" />
                  <MiniStat label="Total delivered" value={stockAuditData.totalDelivered} tone="green" />
                  <MiniStat label="Total damaged" value={stockAuditData.totalDamaged} tone={stockAuditData.totalDamaged > 0 ? "amber" : "green"} />
                  <MiniStat label="Stock value USD" value={formatUSD(stockAuditData.totalStockValueUsd)} tone="green" />
                  <MiniStat label="Missing cost" value={stockAuditData.missingCostCount} tone={stockAuditData.missingCostCount > 0 ? "amber" : "green"} sub={stockAuditData.missingCostCount > 0 ? "products need cost" : "all OK"} />
                </div>
                <div style={{ overflowX: "auto" }}>
                  <table style={{ width: "100%", borderCollapse: "separate", borderSpacing: 0 }}>
                    <thead>
                      <tr>{["Product", "Accepted", "In Stock", "Out Delivered", "Delivered", "Damaged", "Manual Adj.", "Landed Cost/Unit", "Stock Value USD"].map((h) => (
                        <th key={h} style={{ textAlign: "left", padding: "12px 10px", color: textSoft, fontSize: 12, fontWeight: 800, letterSpacing: 0.4, textTransform: "uppercase", borderBottom: `1px solid ${cardBorder}`, background: "rgba(247,243,237,0.92)", whiteSpace: "nowrap" }}>{h}</th>
                      ))}</tr>
                    </thead>
                    <tbody>
                      {stockAuditData.perProduct.map((row, idx) => (
                        <tr key={row.id} style={{ background: idx % 2 === 0 ? "rgba(255,255,255,0.72)" : "rgba(250,247,242,0.8)" }}>
                          <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 700 }}>{row.name}</td>
                          <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}` }}>{row.accepted}</td>
                          <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 800, color: row.available < 0 ? red : textMain }}>{row.available}</td>
                          <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}` }}>{row.outDelivered}</td>
                          <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}` }}>{row.delivered}</td>
                          <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}`, color: row.damaged > 0 ? amber : textMain }}>{row.damaged}</td>
                          <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}`, color: row.manualAdjustments !== 0 ? amber : textSoft }}>{row.manualAdjustments !== 0 ? `${row.manualAdjustments > 0 ? "+" : ""}${row.manualAdjustments}` : "—"}</td>
                          <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}`, color: row.landedCostUsd <= 0 ? red : textMain }}>{row.landedCostUsd > 0 ? formatUSD(row.landedCostUsd) : "Missing"}</td>
                          <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 700 }}>{formatUSD(row.stockValueUsd)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
              </div>
              {/* end audit tab */}

            </div>
          )}

{activePage === "stock" && (() => { setActivePage("products"); setStockTab("overview"); return null; })()}

{activePage === "customersOrders" && (
            <div style={{ display: "grid", gap: 20 }}>
              <PageHeader
                eyebrow="Orders"
                title="Lead pipeline and confirmation"
                description="Import leads, track confirmation status and analyse performance. Shipping and finance are handled in their own menus."
                action={<button style={styles.btnPrimary} onClick={() => ordersImportInputRef.current?.click()}>Import Excel</button>}
              />
              <InlineTabs
                items={[
                  { value: "import", label: "Import Leads" },
                  { value: "pipeline", label: "Lead Pipeline" },
                  { value: "lead-details", label: "Lead Details" },
                  { value: "analytics", label: "Confirmation Analytics" },
                  { value: "audit", label: "Import History" },
                ]}
                value={ordersTab}
                onChange={setOrdersTab}
              />
              <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(4, minmax(0, 1fr))", "repeat(2, minmax(0, 1fr))", "1fr"), gap: 16 }}>
                <KpiCard icon={<Users size={18} />} title="Total Leads" value={confirmationMetrics.total} sub="All leads imported" />
                <KpiCard icon={<ShoppingBag size={18} />} title="Confirmed" value={confirmationMetrics.confirmed} sub={`${confirmationMetrics.confirmationRate.toFixed(1)}% confirmation rate`} valueColor={green} />
                <KpiCard icon={<Phone size={18} />} title="No Reply" value={confirmationMetrics.noReply} sub="Unreached leads" valueColor={amber} />
                <KpiCard icon={<XCircle size={18} />} title="Cancelled" value={confirmationMetrics.cancelled} sub="Cancelled or rejected" valueColor={red} />
              </div>

              {/* Import Leads tab */}
              <div style={{ display: ordersTab === "import" ? "grid" : "none", gap: 16 }}>
                <div style={{ ...styles.card, padding: 22 }}>
                  <div style={styles.sectionHeader}>
                    <div>
                      <div style={styles.sectionEyebrow}>Import leads</div>
                      <div style={{ fontSize: 24, fontWeight: 900, marginTop: 8 }}>Import Excel Leads</div>
                      <div style={{ color: textSoft, marginTop: 6, lineHeight: 1.6 }}>
                        Import your leads Excel file. Each row becomes a lead with confirmation status. Required columns: <strong>Code</strong>, <strong>Phone</strong>, <strong>Product Name</strong>. Optional: Customer, City, Address, Amount, Quantity, Conf Status, Shipping Status, Date.
                      </div>
                    </div>
                    <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                      <button style={styles.btnSecondary} onClick={() => { setOrdersTab("pipeline"); }}>
                        View Pipeline
                      </button>
                      <button style={styles.btnPrimary} onClick={() => ordersImportInputRef.current?.click()}>
                        Import Excel
                      </button>
                    </div>
                  </div>
                  {ordersImportNotice ? (
                    <div style={{ ...styles.softStat, marginTop: 16 }}>
                      <div style={{ fontSize: 11, fontWeight: 800, letterSpacing: 0.45, textTransform: "uppercase", color: textSoft }}>Last import result</div>
                      <div style={{ marginTop: 8, color: textMain, fontWeight: 700 }}>{ordersImportNotice}</div>
                      {ordersImportDetails ? (
                        <div style={{ marginTop: 12, display: "grid", gap: 6, color: textSoft, fontSize: 13, lineHeight: 1.5 }}>
                          <div>Detected headers: {ordersImportDetails.detectedHeaders?.length ? ordersImportDetails.detectedHeaders.join(" · ") : "N/A"}</div>
                          <div>
                            Issues: missing code {ordersImportDetails.reasonCounts?.missingCode || 0} · missing phone {ordersImportDetails.reasonCounts?.missingPhone || 0} · missing product {ordersImportDetails.reasonCounts?.missingProduct || 0} · unmatched product {ordersImportDetails.reasonCounts?.unknownProduct || 0} (imported anyway)
                          </div>
                          <div>
                            Status: {ordersImportDetails.reasonCounts?.statusChangesDetected || 0} status changes · unknown conf {ordersImportDetails.reasonCounts?.unknownConfirmationStatuses || 0} · unknown shipping {ordersImportDetails.reasonCounts?.unknownShippingStatuses || 0}
                          </div>
                          {ordersImportDetails.unmatchedProducts?.length ? (
                            <div>Unmatched products (still imported): {ordersImportDetails.unmatchedProducts.join(" · ")}</div>
                          ) : null}
                        </div>
                      ) : null}
                    </div>
                  ) : null}
                </div>

                <div style={{ ...styles.card, padding: 22 }}>
                  <div style={styles.sectionHeader}>
                    <div>
                      <div style={styles.sectionEyebrow}>Add lead manually</div>
                      <div style={{ fontSize: 22, fontWeight: 800, marginTop: 6 }}>New lead form</div>
                    </div>
                    <button style={styles.btnPrimary} onClick={saveCustomerOrder}>Save Lead</button>
                  </div>
                  <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(3, minmax(0, 1fr))", "repeat(2, minmax(0, 1fr))", "1fr"), gap: 14 }}>
                    <div style={styles.fieldBlock}>
                      <label style={styles.fieldLabel}>Customer name</label>
                      <input style={styles.input} value={customerForm.customerName} onChange={(e) => setCustomerForm({ ...customerForm, customerName: e.target.value })} placeholder="Ex: Amina Yusuf" />
                    </div>
                    <div style={styles.fieldBlock}>
                      <label style={styles.fieldLabel}>Phone</label>
                      <input style={styles.input} value={customerForm.phone} onChange={(e) => setCustomerForm({ ...customerForm, phone: e.target.value })} placeholder="Ex: +255712345678" />
                    </div>
                    <div style={styles.fieldBlock}>
                      <label style={styles.fieldLabel}>City</label>
                      <input style={styles.input} value={customerForm.city} onChange={(e) => setCustomerForm({ ...customerForm, city: e.target.value })} placeholder="Ex: Dar es Salaam" />
                    </div>
                    <div style={styles.fieldBlock}>
                      <label style={styles.fieldLabel}>Address</label>
                      <input style={styles.input} value={customerForm.address} onChange={(e) => setCustomerForm({ ...customerForm, address: e.target.value })} placeholder="Ex: Mikocheni, Block A" />
                    </div>
                    <div style={styles.fieldBlock}>
                      <label style={styles.fieldLabel}>Product</label>
                      <select style={styles.input} value={customerForm.productId} onChange={(e) => setCustomerForm({ ...customerForm, productId: e.target.value })}>
                        {products.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
                      </select>
                    </div>
                    <div style={styles.fieldBlock}>
                      <label style={styles.fieldLabel}>Quantity</label>
                      <input style={styles.input} type="number" min="1" value={customerForm.quantity} onChange={(e) => setCustomerForm({ ...customerForm, quantity: e.target.value })} />
                    </div>
                    <div style={styles.fieldBlock}>
                      <label style={styles.fieldLabel}>Order date</label>
                      <input style={styles.input} type="date" value={customerForm.orderDate} onChange={(e) => setCustomerForm({ ...customerForm, orderDate: e.target.value })} />
                    </div>
                    <div style={styles.fieldBlock}>
                      <label style={styles.fieldLabel}>Payment method</label>
                      <select style={styles.input} value={customerForm.paymentMethod} onChange={(e) => setCustomerForm({ ...customerForm, paymentMethod: e.target.value })}>
                        <option value="COD">COD</option>
                        <option value="Card">Card</option>
                        <option value="M-Pesa">M-Pesa</option>
                        <option value="Cash">Cash</option>
                      </select>
                    </div>
                    <div style={styles.fieldBlock}>
                      <label style={styles.fieldLabel}>Confirmation status</label>
                      <select style={styles.input} value={normalizeOrderStatus(customerForm.status)} onChange={(e) => setCustomerForm({ ...customerForm, status: e.target.value })}>
                        {confirmationStatusCatalog.map((status) => (
                          <option key={status.value} value={status.value}>{status.label}</option>
                        ))}
                      </select>
                    </div>
                    <div style={styles.fieldBlock}>
                      <label style={styles.fieldLabel}>Lead source</label>
                      <select style={styles.input} value={customerForm.leadSource} onChange={(e) => setCustomerForm({ ...customerForm, leadSource: e.target.value })}>
                        <option value="manual">Manual</option>
                        <option value="meta">Meta Ads</option>
                        <option value="tiktok">TikTok Ads</option>
                        <option value="whatsapp">WhatsApp</option>
                        <option value="sheet">Sheet</option>
                        <option value="marketplace">Marketplace</option>
                        <option value="other">Other</option>
                      </select>
                    </div>
                    <div style={styles.fieldBlock}>
                      <label style={styles.fieldLabel}>Notes</label>
                      <input style={styles.input} value={customerForm.notes} onChange={(e) => setCustomerForm({ ...customerForm, notes: e.target.value })} placeholder="Ex: Call in the afternoon" />
                    </div>
                  </div>
                </div>
              </div>

              {/* Lead Pipeline tab */}
              <div style={{ display: ordersTab === "pipeline" ? "grid" : "none", gap: 16 }}>
              <div style={{ ...styles.card, padding: 22 }}>
                <div style={styles.sectionHeader}>
                  <div>
                    <div style={styles.sectionEyebrow}>Lead pipeline</div>
                    <div style={{ fontSize: 24, fontWeight: 900, marginTop: 8 }}>All leads</div>
                    <div style={{ color: textSoft, marginTop: 6, lineHeight: 1.6 }}>All imported and manually added leads with confirmation status. Click View to open a lead detail.</div>
                  </div>
                  <button style={styles.btnPrimary} onClick={() => setOrdersTab("import")}>
                    Add Lead
                  </button>
                </div>

                <div style={{ display: "grid", gap: 16 }}>
                  <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(3, minmax(0, 1fr))", "1fr 1fr", "1fr"), gap: 12 }}>
                    <MiniStat label="Filtered leads" value={compactCustomerRows.length} tone="blue" sub="Matching current filters" />
                    <MiniStat label="Confirmed" value={filteredCustomerSummary.confirmed} tone="green" sub="Active confirmation" />
                    <MiniStat label="Pending / Cancelled" value={filteredCustomerSummary.pending + filteredCustomerSummary.cancelled} tone="amber" sub="In flow or lost" />
                  </div>

                  <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("minmax(200px, 1fr) 160px 160px 140px 110px", "1fr 1fr", "1fr"), gap: 12 }}>
                    <input
                      style={styles.input}
                      value={customerListFilters.search}
                      onChange={(e) => setCustomerListFilters((prev) => ({ ...prev, search: e.target.value }))}
                      placeholder="Search name, phone, city, product, id..."
                    />
                    <select
                      style={styles.input}
                      value={customerListFilters.status}
                      onChange={(e) => setCustomerListFilters((prev) => ({ ...prev, status: e.target.value }))}
                    >
                      <option value="all">All statuses</option>
                      {confirmationStatusCatalog.map((status) => (
                        <option key={status.value} value={status.value}>
                          {status.label} ({status.count})
                        </option>
                      ))}
                    </select>
                    <select
                      style={styles.input}
                      value={customerListFilters.productId}
                      onChange={(e) => setCustomerListFilters((prev) => ({ ...prev, productId: e.target.value }))}
                    >
                      <option value="all">All products</option>
                      {products.map((p) => (
                        <option key={p.id} value={p.id}>{p.name}</option>
                      ))}
                    </select>
                    <select
                      style={styles.input}
                      value={customerListFilters.city}
                      onChange={(e) => setCustomerListFilters((prev) => ({ ...prev, city: e.target.value }))}
                    >
                      <option value="all">All cities</option>
                      {ordersPageCityList.map((city) => (
                        <option key={city} value={city}>{city}</option>
                      ))}
                    </select>
                    <select
                      style={styles.input}
                      value={customerListFilters.pageSize}
                      onChange={(e) => setCustomerListFilters((prev) => ({ ...prev, pageSize: Number(e.target.value) }))}
                    >
                      <option value={10}>10 / page</option>
                      <option value={25}>25 / page</option>
                      <option value={50}>50 / page</option>
                      <option value={100}>100 / page</option>
                    </select>
                  </div>

                  <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("minmax(220px, 1fr) minmax(220px, 1fr) auto auto auto", "1fr 1fr", "1fr"), gap: 10, alignItems: "end" }}>
                    <div style={styles.fieldBlock}>
                      <label style={styles.fieldLabel}>Bulk confirmation status</label>
                      <select style={styles.input} value={bulkCustomerStatus} onChange={(e) => setBulkCustomerStatus(e.target.value)}>
                        {confirmationStatusCatalog.map((status) => (
                          <option key={status.value} value={status.value}>
                            {status.label}
                          </option>
                        ))}
                      </select>
                    </div>
                    <div style={styles.fieldBlock}>
                      <label style={styles.fieldLabel}>Bulk owner</label>
                      <select style={styles.input} value={bulkCustomerOwner} onChange={(e) => setBulkCustomerOwner(e.target.value)}>
                        <option value="">No owner</option>
                        {teamRoster.map((member) => (
                          <option key={member} value={member}>
                            {member}
                          </option>
                        ))}
                      </select>
                    </div>
                    <button style={styles.btnSecondary} disabled={selectedCustomerIds.length === 0} onClick={updateCustomersBulkConfirmationStatus}>
                      Apply status
                    </button>
                    <button style={styles.btnSecondary} disabled={selectedCustomerIds.length === 0} onClick={assignCustomersBulkOwner}>
                      Assign owner
                    </button>
                    <button style={{ ...styles.btnSecondary, background: "#fef2f2", color: red, border: "1px solid #fecaca" }} disabled={selectedCustomerIds.length === 0} onClick={deleteSelectedCustomerOrders}>
                      Delete selected
                    </button>
                  </div>

                  {historyTargetCustomer ? (
                    <div style={{ ...styles.softStat, display: "grid", gap: 10 }}>
                      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12 }}>
                        <div>
                          <div style={{ fontWeight: 800 }}>Order history: {historyTargetCustomer.customerName}</div>
                          <div style={{ color: textSoft, fontSize: 12, marginTop: 4 }}>{historyTargetCustomer.id} | {historyTargetCustomer.sourceOrderId || "Manual reference"}</div>
                        </div>
                        <button style={styles.btnSecondary} onClick={() => setCustomerHistoryTargetId("")}>Close history</button>
                      </div>
                      <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(4, minmax(0, 1fr))", "1fr 1fr", "1fr"), gap: 10 }}>
                        <MiniStat label="Source lead" value={historyTargetCustomer.leadSource || "manual"} tone="blue" sub={historyTargetCustomer.campaignName || "No campaign"} />
                        <MiniStat label="Priority" value={formatStatusLabel(historyTargetCustomer.priority || "normal")} tone="amber" sub={`${formatStatusLabel(historyTargetCustomer.customerType || "new")} customer`} />
                        <MiniStat label="Shipping" value={historyTargetCustomer.carrierName || "Not assigned"} tone="green" sub={historyTargetCustomer.trackingNumber || "No tracking"} />
                        <MiniStat label="Business reason" value={historyTargetCustomer.cancelReason || historyTargetCustomer.unreachedReason || historyTargetCustomer.returnReason || "None"} dark sub="Cancel / unreached / return" />
                      </div>
                      <div style={{ display: "grid", gap: 8 }}>
                        {historyTargetCustomer.history?.length ? historyTargetCustomer.history.slice(0, 8).map((entry) => (
                          <div key={entry.id} style={{ padding: "10px 12px", borderRadius: 12, background: "rgba(255,255,255,0.82)", border: `1px solid ${cardBorder}` }}>
                            <div style={{ fontWeight: 700, fontSize: 13 }}>{formatStatusLabel(entry.action)}</div>
                            <div style={{ color: textSoft, fontSize: 12, marginTop: 4 }}>{entry.at ? new Date(entry.at).toLocaleString() : "No date"} | {entry.source}</div>
                            <div style={{ color: textMain, fontSize: 13, marginTop: 6 }}>{entry.details || "No details"}</div>
                          </div>
                        )) : <div style={{ color: textSoft }}>No history saved yet.</div>}
                      </div>
                    </div>
                  ) : null}

                  <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap" }}>
                    <div style={{ color: textSoft, fontSize: 13 }}>
                      Showing {paginatedCustomerRows.length} of {compactCustomerRows.length} filtered orders | Selected {selectedCustomerIds.length}
                    </div>
                    <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
                      <button
                        style={styles.btnSecondary}
                        disabled={selectedCustomerIds.length === 0}
                        onClick={() => setSelectedCustomerIds([])}
                      >
                        Clear selection
                      </button>
                      <button
                        style={{ ...styles.btnSecondary, background: "#fef2f2", color: red, border: "1px solid #fecaca" }}
                        disabled={selectedCustomerIds.length === 0}
                        onClick={deleteSelectedCustomerOrders}
                      >
                        Delete selected
                      </button>
                      <button
                        style={styles.btnSecondary}
                        disabled={customerListPage <= 1}
                        onClick={() => setCustomerListPage((prev) => Math.max(1, prev - 1))}
                      >
                        Previous
                      </button>
                      <span style={{ color: textSoft, fontSize: 13, fontWeight: 700 }}>
                        Page {customerListPage} / {customerListPageCount}
                      </span>
                      <button
                        style={styles.btnSecondary}
                        disabled={customerListPage >= customerListPageCount}
                        onClick={() => setCustomerListPage((prev) => Math.min(customerListPageCount, prev + 1))}
                      >
                        Next
                      </button>
                    </div>
                  </div>

                  <div style={{ overflowX: "auto", border: `1px solid ${cardBorder}`, borderRadius: 22, background: "linear-gradient(180deg, rgba(255,255,255,0.96), rgba(248,244,238,0.9))" }}>
                    <table style={{ width: "100%", borderCollapse: "separate", borderSpacing: 0 }}>
                      <thead>
                        <tr>
                          {["Select", "Customer", "Product", "Order", "Owner", "Amount", "Status", "Details", "Actions"].map((head) => (
                            <th
                              key={head}
                              style={{
                                textAlign: head === "Actions" ? "right" : "left",
                                padding: "12px 10px",
                                color: textSoft,
                                fontSize: 12,
                                fontWeight: 800,
                                letterSpacing: 0.45,
                                textTransform: "uppercase",
                                borderBottom: `1px solid ${cardBorder}`,
                                background: "rgba(247, 243, 237, 0.92)",
                                whiteSpace: "nowrap",
                              }}
                            >
                              {head === "Select" ? (
                                <input
                                  ref={selectAllCustomersRef}
                                  type="checkbox"
                                  checked={allFilteredSelected}
                                  onChange={() =>
                                    setSelectedCustomerIds((prev) => {
                                      const next = new Set(prev);
                                      if (allFilteredSelected) {
                                        filteredCustomerIds.forEach((id) => next.delete(id));
                                      } else {
                                        filteredCustomerIds.forEach((id) => next.add(id));
                                      }
                                      return Array.from(next);
                                    })
                                  }
                                />
                              ) : (
                                head
                              )}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {paginatedCustomerRows.map((customer, index) => {
                          return (
                            <tr key={customer.id} style={{ background: index % 2 === 0 ? "rgba(255,255,255,0.7)" : "rgba(250,247,242,0.82)" }}>
                              <td style={{ padding: "10px 10px", borderBottom: `1px solid ${cardBorder}`, width: 54 }}>
                                <input
                                  type="checkbox"
                                  checked={selectedCustomerIdSet.has(customer.id)}
                                  onChange={() =>
                                    setSelectedCustomerIds((prev) =>
                                      prev.includes(customer.id) ? prev.filter((id) => id !== customer.id) : [...prev, customer.id]
                                    )
                                  }
                                />
                              </td>
                              <td style={{ padding: "10px 10px", borderBottom: `1px solid ${cardBorder}`, minWidth: 210 }}>
                                <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                                  <div
                                    style={{
                                      width: 30,
                                      height: 30,
                                      borderRadius: 10,
                                      display: "grid",
                                      placeItems: "center",
                                      background: "linear-gradient(135deg, rgba(29,95,208,0.14), rgba(29,95,208,0.04))",
                                      color: accent,
                                      fontWeight: 900,
                                      flexShrink: 0,
                                    }}
                                  >
                                    {String(customer.customerName || "C").trim().slice(0, 2).toUpperCase()}
                                  </div>
                                  <div style={{ minWidth: 0 }}>
                                    <div style={{ fontWeight: 800, fontSize: 14, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{customer.customerName}</div>
                                    <div style={{ color: textSoft, fontSize: 11, marginTop: 2 }}>{customer.id} | {customer.phone}</div>
                                    <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginTop: 6 }}>
                                      <span style={{ ...styles.badge, background: "rgba(29,95,208,0.08)", color: accent, border: "1px solid rgba(29,95,208,0.12)" }}>
                                        {formatStatusLabel(customer.leadSource || "manual")}
                                      </span>
                                      <span style={{ ...styles.badge, background: "rgba(199,131,34,0.12)", color: amber, border: "1px solid rgba(199,131,34,0.18)" }}>
                                        {formatStatusLabel(customer.priority || "normal")}
                                      </span>
                                    </div>
                                  </div>
                                </div>
                              </td>
                              <td style={{ padding: "10px 10px", borderBottom: `1px solid ${cardBorder}`, minWidth: 180 }}>
                                <div style={{ fontWeight: 700, fontSize: 13 }}>{customer.productName}</div>
                                <div style={{ color: textSoft, fontSize: 11, marginTop: 2 }}>{customer.city || "N/A"}</div>
                                <div style={{ color: textSoft, fontSize: 11, marginTop: 4, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
                                  {customer.campaignName || customer.creativeName || "No campaign detail"}
                                </div>
                              </td>
                              <td style={{ padding: "10px 10px", borderBottom: `1px solid ${cardBorder}`, minWidth: 160 }}>
                                <div style={{ color: textMain, fontWeight: 700, fontSize: 13 }}>{customer.orderDate}</div>
                                <div style={{ color: textSoft, fontSize: 11, marginTop: 2 }}>
                                  Qty {customer.quantity} | {customer.paymentMethod}
                                </div>
                              </td>
                              <td style={{ padding: "10px 10px", borderBottom: `1px solid ${cardBorder}`, minWidth: 180 }}>
                                <select
                                  style={{ ...styles.input, padding: "8px 10px", minWidth: 0, fontSize: 12 }}
                                  value={customer.assignedTo || ""}
                                  onChange={(e) => assignCustomerOwner(customer.id, e.target.value)}
                                >
                                  <option value="">No owner</option>
                                  {teamRoster.map((member) => (
                                    <option key={member} value={member}>
                                      {member}
                                    </option>
                                  ))}
                                </select>
                              </td>
                              <td style={{ padding: "10px 10px", borderBottom: `1px solid ${cardBorder}`, whiteSpace: "nowrap", fontWeight: 800, fontSize: 13 }}>
                                {formatTZS(customer.totalValue)}
                              </td>
                              <td style={{ padding: "10px 10px", borderBottom: `1px solid ${cardBorder}`, minWidth: 180 }}>
                                <div style={{ display: "grid", gap: 8 }}>
                                  <span style={getStatusBadgeStyle(customer.status)}>
                                    {customer.statusLabel}
                                  </span>
                                  <select
                                    style={{ ...styles.input, padding: "8px 10px", minWidth: 0, fontSize: 12 }}
                                    value={customer.confirmationStatus || customer.status}
                                    onChange={(e) => updateCustomerStatus(customer.id, e.target.value)}
                                  >
                                    {confirmationStatusCatalog.map((status) => (
                                      <option key={status.value} value={status.value}>
                                        {status.label}
                                      </option>
                                    ))}
                                  </select>
                                </div>
                              </td>
                              <td style={{ padding: "10px 10px", borderBottom: `1px solid ${cardBorder}`, minWidth: 110 }}>
                                <button style={styles.btnSecondary} onClick={() => { setSelectedLeadId(customer.id); setOrdersTab("lead-details"); }}>
                                  View
                                </button>
                              </td>
                              <td style={{ padding: "10px 10px", borderBottom: `1px solid ${cardBorder}`, textAlign: "right", minWidth: 190 }}>
                                <div style={{ display: "inline-flex", gap: 8, flexWrap: "wrap", justifyContent: "flex-end" }}>
                                  <button
                                    style={{ ...styles.btnSecondary, background: "#fef2f2", color: red, border: "1px solid #fecaca", padding: "8px 10px", fontSize: 12 }}
                                    onClick={() => deleteCustomerOrder(customer.id)}
                                  >
                                    Delete
                                  </button>
                                </div>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>

                    {compactCustomerRows.length === 0 ? (
                      <div style={{ padding: 24, color: textSoft }}>No leads match the current filters.</div>
                    ) : null}
                  </div>
                </div>
              </div>
              </div>

              {/* Lead Details tab */}
              <div style={{ display: ordersTab === "lead-details" ? "grid" : "none", gap: 16 }}>
                {selectedLead ? (
                  <div style={{ ...styles.card, padding: 22 }}>
                    <div style={styles.sectionHeader}>
                      <div>
                        <div style={styles.sectionEyebrow}>Lead detail</div>
                        <div style={{ fontSize: 24, fontWeight: 900, marginTop: 8 }}>{selectedLead.customerName || "Unnamed lead"}</div>
                        <div style={{ color: textSoft, marginTop: 4, fontSize: 13 }}>{selectedLead.id} · {selectedLead.phone} · {selectedLead.city || "No city"}</div>
                      </div>
                      <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                        <button style={styles.btnSecondary} onClick={() => { setCustomerHistoryTargetId(selectedLead.id); setOrdersTab("pipeline"); }}>
                          Order history
                        </button>
                        <button style={styles.btnSecondary} onClick={() => setOrdersTab("pipeline")}>
                          Back to pipeline
                        </button>
                      </div>
                    </div>
                    <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(3, minmax(0, 1fr))", "1fr 1fr", "1fr"), gap: 14, marginTop: 8 }}>
                      <MiniStat label="Product" value={products.find((p) => p.id === selectedLead.productId)?.name || selectedLead.productId || "—"} tone="blue" sub={`Qty ${selectedLead.quantity || 1}`} />
                      <MiniStat label="Confirmation status" value={confirmationStatusMap[getCustomerConfirmationStatus(selectedLead)]?.label || getCustomerConfirmationStatus(selectedLead)} tone="green" sub={selectedLead.orderDate || "No date"} />
                      <MiniStat label="Payment method" value={selectedLead.paymentMethod || "COD"} tone="amber" sub={selectedLead.leadSource || "manual"} />
                      <MiniStat label="Owner" value={selectedLead.assignedTo || "Unassigned"} sub={selectedLead.priority || "normal"} />
                      <MiniStat label="Campaign" value={selectedLead.campaignName || "—"} tone="blue" sub={selectedLead.adsetName || selectedLead.creativeName || "No ad detail"} />
                      <MiniStat label="Source order ID" value={selectedLead.sourceOrderId || "—"} sub={`Customer type: ${selectedLead.customerType || "new"}`} />
                    </div>
                    {(selectedLead.notes || selectedLead.cancelReason || selectedLead.unreachedReason || selectedLead.returnReason) ? (
                      <div style={{ marginTop: 16, padding: 16, borderRadius: 16, background: "rgba(247,243,237,0.85)", border: `1px solid ${cardBorder}` }}>
                        {selectedLead.notes ? <div style={{ fontSize: 13, color: textMain }}><strong>Notes:</strong> {selectedLead.notes}</div> : null}
                        {selectedLead.cancelReason ? <div style={{ fontSize: 13, color: textMain, marginTop: 6 }}><strong>Cancel reason:</strong> {selectedLead.cancelReason}</div> : null}
                        {selectedLead.unreachedReason ? <div style={{ fontSize: 13, color: textMain, marginTop: 6 }}><strong>Unreached reason:</strong> {selectedLead.unreachedReason}</div> : null}
                        {selectedLead.returnReason ? <div style={{ fontSize: 13, color: textMain, marginTop: 6 }}><strong>Return reason:</strong> {selectedLead.returnReason}</div> : null}
                      </div>
                    ) : null}
                    {selectedLead.history?.length ? (
                      <div style={{ marginTop: 16, display: "grid", gap: 8 }}>
                        <div style={{ fontWeight: 800, fontSize: 13, color: textSoft }}>History ({selectedLead.history.length} events)</div>
                        {selectedLead.history.slice(0, 10).map((entry) => (
                          <div key={entry.id} style={{ padding: "10px 14px", borderRadius: 12, background: "rgba(255,255,255,0.88)", border: `1px solid ${cardBorder}` }}>
                            <div style={{ fontWeight: 700, fontSize: 13 }}>{formatStatusLabel(entry.action)}</div>
                            <div style={{ color: textSoft, fontSize: 12, marginTop: 4 }}>{entry.at ? new Date(entry.at).toLocaleString() : "No date"} · {entry.source}</div>
                            {entry.details ? <div style={{ color: textMain, fontSize: 13, marginTop: 6 }}>{entry.details}</div> : null}
                          </div>
                        ))}
                      </div>
                    ) : null}
                  </div>
                ) : (
                  <div style={{ ...styles.card, padding: 36, textAlign: "center" }}>
                    <div style={{ color: textSoft, fontSize: 15 }}>No lead selected. Go to <strong>Lead Pipeline</strong> and click <strong>View</strong> on a lead.</div>
                  </div>
                )}
              </div>

              {/* Confirmation Analytics tab */}
              <div style={{ display: ordersTab === "analytics" ? "grid" : "none", gap: 16 }}>
                <div style={{ ...styles.card, padding: 22 }}>
                  <div style={styles.sectionHeader}>
                    <div>
                      <div style={styles.sectionEyebrow}>Confirmation analytics</div>
                      <div style={{ fontSize: 24, fontWeight: 900, marginTop: 8 }}>Confirmation performance</div>
                      <div style={{ color: textSoft, marginTop: 6, lineHeight: 1.6 }}>Lead volume, confirmation rate, owner workload and city breakdown.</div>
                    </div>
                  </div>

                  <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(4, minmax(0, 1fr))", "1fr 1fr", "1fr"), gap: 12, marginBottom: 20 }}>
                    <MiniStat label="Total leads" value={formatInteger(confirmationMetrics.total)} tone="blue" />
                    <MiniStat label="Confirmed" value={formatInteger(confirmationMetrics.confirmed)} tone="green" sub={`${confirmationMetrics.confirmationRate.toFixed(1)}% rate`} />
                    <MiniStat label="No reply" value={formatInteger(confirmationMetrics.noReply)} tone="amber" />
                    <MiniStat label="Cancelled" value={formatInteger(confirmationMetrics.cancelled)} tone="amber" />
                  </div>

                  <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("1fr 1fr", "1fr", "1fr"), gap: 16, marginBottom: 16 }}>
                    <div style={{ overflowX: "auto", border: `1px solid ${cardBorder}`, borderRadius: 20 }}>
                      <table style={{ width: "100%", borderCollapse: "separate", borderSpacing: 0 }}>
                        <thead>
                          <tr>
                            {["Status", "Count", "Share"].map((head) => (
                              <th key={head} style={{ textAlign: "left", padding: "12px 12px", color: textSoft, fontSize: 12, fontWeight: 800, letterSpacing: 0.4, textTransform: "uppercase", borderBottom: `1px solid ${cardBorder}`, background: "rgba(247, 243, 237, 0.92)" }}>{head}</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {confirmationStatusCatalog.map((status, index) => {
                            const share = confirmationMetrics.total > 0 ? (Number(status.count || 0) / confirmationMetrics.total) * 100 : 0;
                            return (
                              <tr key={`anl-status-${status.value}`} style={{ background: index % 2 === 0 ? "rgba(255,255,255,0.72)" : "rgba(250,247,242,0.8)" }}>
                                <td style={{ padding: "12px 12px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 700 }}>{status.label}</td>
                                <td style={{ padding: "12px 12px", borderBottom: `1px solid ${cardBorder}` }}>{formatInteger(status.count || 0)}</td>
                                <td style={{ padding: "12px 12px", borderBottom: `1px solid ${cardBorder}` }}>{share.toFixed(1)}%</td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>

                    <div style={{ overflowX: "auto", border: `1px solid ${cardBorder}`, borderRadius: 20 }}>
                      <table style={{ width: "100%", borderCollapse: "separate", borderSpacing: 0 }}>
                        <thead>
                          <tr>
                            {["Owner", "Leads", "Confirmed", "Conf. Rate"].map((head) => (
                              <th key={head} style={{ textAlign: "left", padding: "12px 12px", color: textSoft, fontSize: 12, fontWeight: 800, letterSpacing: 0.4, textTransform: "uppercase", borderBottom: `1px solid ${cardBorder}`, background: "rgba(247, 243, 237, 0.92)" }}>{head}</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {teamWorkloadRows.map((row, index) => (
                            <tr key={`anl-owner-${row.owner}`} style={{ background: index % 2 === 0 ? "rgba(255,255,255,0.72)" : "rgba(250,247,242,0.8)" }}>
                              <td style={{ padding: "12px 12px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 700 }}>{row.owner}</td>
                              <td style={{ padding: "12px 12px", borderBottom: `1px solid ${cardBorder}` }}>{row.total}</td>
                              <td style={{ padding: "12px 12px", borderBottom: `1px solid ${cardBorder}` }}>{row.confirmed}</td>
                              <td style={{ padding: "12px 12px", borderBottom: `1px solid ${cardBorder}` }}>{row.total > 0 ? `${((row.confirmed / row.total) * 100).toFixed(1)}%` : "—"}</td>
                            </tr>
                          ))}
                          {teamWorkloadRows.length === 0 ? (
                            <tr><td colSpan={4} style={{ padding: 24, color: textSoft }}>No owner assignments yet.</td></tr>
                          ) : null}
                        </tbody>
                      </table>
                    </div>
                  </div>

                  <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("1fr 1fr", "1fr", "1fr"), gap: 16 }}>
                    <div style={{ overflowX: "auto", border: `1px solid ${cardBorder}`, borderRadius: 20 }}>
                      <div style={{ padding: "12px 16px", fontWeight: 800, fontSize: 13, borderBottom: `1px solid ${cardBorder}`, background: "rgba(247,243,237,0.92)", borderRadius: "20px 20px 0 0" }}>By product</div>
                      <table style={{ width: "100%", borderCollapse: "separate", borderSpacing: 0 }}>
                        <thead>
                          <tr>
                            {["Product", "Leads", "Confirmed", "Conf. Rate"].map((head) => (
                              <th key={head} style={{ textAlign: "left", padding: "12px 12px", color: textSoft, fontSize: 12, fontWeight: 800, letterSpacing: 0.4, textTransform: "uppercase", borderBottom: `1px solid ${cardBorder}`, background: "rgba(247, 243, 237, 0.92)" }}>{head}</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {confirmationMetrics.productRows.map((row, index) => (
                            <tr key={`anl-prod-${row.productId}`} style={{ background: index % 2 === 0 ? "rgba(255,255,255,0.72)" : "rgba(250,247,242,0.8)" }}>
                              <td style={{ padding: "12px 12px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 700 }}>{row.productName}</td>
                              <td style={{ padding: "12px 12px", borderBottom: `1px solid ${cardBorder}` }}>{row.total}</td>
                              <td style={{ padding: "12px 12px", borderBottom: `1px solid ${cardBorder}` }}>{row.confirmed}</td>
                              <td style={{ padding: "12px 12px", borderBottom: `1px solid ${cardBorder}` }}>{row.confirmationRate.toFixed(1)}%</td>
                            </tr>
                          ))}
                          {confirmationMetrics.productRows.length === 0 ? (
                            <tr><td colSpan={4} style={{ padding: 24, color: textSoft }}>No data yet.</td></tr>
                          ) : null}
                        </tbody>
                      </table>
                    </div>

                    <div style={{ overflowX: "auto", border: `1px solid ${cardBorder}`, borderRadius: 20 }}>
                      <div style={{ padding: "12px 16px", fontWeight: 800, fontSize: 13, borderBottom: `1px solid ${cardBorder}`, background: "rgba(247,243,237,0.92)", borderRadius: "20px 20px 0 0" }}>By city</div>
                      <table style={{ width: "100%", borderCollapse: "separate", borderSpacing: 0 }}>
                        <thead>
                          <tr>
                            {["City", "Leads", "Confirmed", "Conf. Rate"].map((head) => (
                              <th key={head} style={{ textAlign: "left", padding: "12px 12px", color: textSoft, fontSize: 12, fontWeight: 800, letterSpacing: 0.4, textTransform: "uppercase", borderBottom: `1px solid ${cardBorder}`, background: "rgba(247, 243, 237, 0.92)" }}>{head}</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {confirmationMetrics.cityRows.map((row, index) => (
                            <tr key={`anl-city-${row.city}`} style={{ background: index % 2 === 0 ? "rgba(255,255,255,0.72)" : "rgba(250,247,242,0.8)" }}>
                              <td style={{ padding: "12px 12px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 700 }}>{row.city}</td>
                              <td style={{ padding: "12px 12px", borderBottom: `1px solid ${cardBorder}` }}>{row.total}</td>
                              <td style={{ padding: "12px 12px", borderBottom: `1px solid ${cardBorder}` }}>{row.confirmed}</td>
                              <td style={{ padding: "12px 12px", borderBottom: `1px solid ${cardBorder}` }}>{row.confirmationRate.toFixed(1)}%</td>
                            </tr>
                          ))}
                          {confirmationMetrics.cityRows.length === 0 ? (
                            <tr><td colSpan={4} style={{ padding: 24, color: textSoft }}>No city data yet.</td></tr>
                          ) : null}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </div>
              </div>

              {/* Import History / Audit tab */}
              <div style={{ display: ordersTab === "audit" ? "grid" : "none", gap: 16 }}>
                <div style={{ ...styles.card, padding: 22 }}>
                  <div style={styles.sectionHeader}>
                    <div>
                      <div style={styles.sectionEyebrow}>Import history</div>
                      <div style={{ fontSize: 24, fontWeight: 900, marginTop: 8 }}>Excel import log</div>
                      <div style={{ color: textSoft, marginTop: 6, lineHeight: 1.6 }}>Last {ordersImportHistory.length} imports. Up to 20 entries are kept per session.</div>
                    </div>
                  </div>
                  {ordersImportHistory.length === 0 ? (
                    <div style={{ color: textSoft, padding: "20px 0" }}>No imports yet this session.</div>
                  ) : (
                    <div style={{ display: "grid", gap: 12 }}>
                      {ordersImportHistory.map((record, index) => (
                        <div key={index} style={{ padding: "16px 18px", borderRadius: 16, border: `1px solid ${cardBorder}`, background: "rgba(255,255,255,0.85)" }}>
                          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 12, flexWrap: "wrap" }}>
                            <div>
                              <div style={{ fontWeight: 800 }}>Import #{ordersImportHistory.length - index}</div>
                              <div style={{ color: textSoft, fontSize: 12, marginTop: 3 }}>{record.importedAt ? new Date(record.importedAt).toLocaleString() : "Unknown time"}</div>
                            </div>
                            <span style={{ ...styles.badge, background: "rgba(22,163,74,0.1)", color: green, border: "1px solid rgba(22,163,74,0.2)" }}>{record.summary}</span>
                          </div>
                          {record.reasonCounts ? (
                            <div style={{ marginTop: 10, color: textSoft, fontSize: 13, lineHeight: 1.6 }}>
                              Issues: missing code {record.reasonCounts.missingCode || 0} · missing phone {record.reasonCounts.missingPhone || 0} · unmatched product {record.reasonCounts.unknownProduct || 0} · status changes {record.reasonCounts.statusChangesDetected || 0}
                            </div>
                          ) : null}
                          {record.detectedHeaders?.length ? (
                            <div style={{ marginTop: 6, color: textSoft, fontSize: 12 }}>Headers: {record.detectedHeaders.join(" · ")}</div>
                          ) : null}
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              </div>
            </div>
          )}

{activePage === "shipping" && (
            <div style={{ display: "grid", gap: 20 }}>
              <input
                ref={shippingImportInputRef}
                type="file"
                accept=".xlsx,.xls,.csv"
                onChange={importShippingFromExcel}
                style={{ display: "none" }}
              />
              <PageHeader
                eyebrow="Shipping"
                title="Delivery flow and shipping control"
                description="Update shipping statuses, monitor delivery queue health and isolate exceptions without mixing confirmation or finance work."
                action={<button style={styles.btnPrimary} onClick={() => shippingImportInputRef.current?.click()}>Import Shipping Excel</button>}
              />
              <InlineTabs
                items={[
                  { value: "import", label: "Import Shipping" },
                  { value: "queue", label: "Delivery Queue" },
                  { value: "analytics", label: "Delivery Analytics" },
                  { value: "exceptions", label: "Exceptions" },
                ]}
                value={shippingTab}
                onChange={setShippingTab}
              />

              <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(4, minmax(0, 1fr))", "repeat(2, minmax(0, 1fr))", "1fr"), gap: 16 }}>
                <KpiCard icon={<ShoppingBag size={18} />} title="Shipping Queue" value={shippingSummary.total} sub="Orders already out of new lead stage" />
                <KpiCard icon={<Rocket size={18} />} title="In Delivery Flow" value={shippingSummary.activeShipping} sub="Confirmed, transit or shipping statuses" valueColor={amber} />
                <KpiCard icon={<Wallet size={18} />} title="Delivered" value={shippingSummary.deliveredShipping} sub="Completed delivery orders" valueColor={green} />
                <KpiCard icon={<AlertTriangle size={18} />} title="Exceptions" value={shippingSummary.cancelledShipping + shippingSummary.otherShipping} sub="Cancelled or custom shipping statuses" valueColor={red} />
              </div>

              <div style={{ ...styles.card, padding: 22 }}>
                <div style={styles.sectionHeader}>
                  <div>
                    <div style={styles.sectionEyebrow}>Shipping control</div>
                    <div style={{ fontSize: 24, fontWeight: 900, marginTop: 8 }}>Suivi des commandes confirmees</div>
                    <div style={{ color: textSoft, marginTop: 6, lineHeight: 1.6 }}>
                      Importez le fichier Excel de shipping pour mettre a jour les statuts des commandes deja existantes. Si une ligne est deja connue et que le statut change, l'app la met a jour. Sinon elle l'ignore automatiquement.
                    </div>
                  </div>
                  <button style={styles.btnPrimary} onClick={() => shippingImportInputRef.current?.click()}>
                    Import Shipping Excel
                  </button>
                </div>

                <div style={{ display: "grid", gap: 16 }}>
                  <div style={{ padding: 18, borderRadius: 20, border: `1px solid ${cardBorder}`, background: "linear-gradient(180deg, rgba(255,255,255,0.94), rgba(248,244,238,0.88))" }}>
                    <div style={{ color: textSoft, fontSize: 13, lineHeight: 1.6 }}>
                      Colonnes conseillees : <strong>Order ID</strong>, <strong>Phone</strong>, <strong>Product name</strong>, <strong>Order date</strong>, <strong>Quantity</strong>, <strong>Shipping status</strong>.
                    </div>
                    <div style={{ marginTop: 8, color: textSoft, fontSize: 13, lineHeight: 1.6 }}>
                      L'import shipping ne cree pas de nouvelles commandes : il met seulement a jour les commandes deja presentes dans l'app.
                    </div>

                    {shippingImportNotice ? (
                      <div style={{ marginTop: 14, paddingTop: 14, borderTop: `1px solid ${cardBorder}` }}>
                        <div style={{ color: textMain, fontWeight: 700 }}>{shippingImportNotice}</div>

                        {shippingImportDetails ? (
                          <div style={{ marginTop: 10, color: textSoft, fontSize: 13, lineHeight: 1.6 }}>
                            <div>
                              Detected headers: {shippingImportDetails.detectedHeaders.length ? shippingImportDetails.detectedHeaders.join(" | ") : "N/A"}
                            </div>
                            <div>
                              Skip reasons: missing status {shippingImportDetails.reasonCounts.missingStatus}, unmatched order {shippingImportDetails.reasonCounts.unmatchedOrder}
                            </div>
                            {shippingImportDetails.unmatchedExamples.length ? (
                              <div>
                                Unmatched examples: {shippingImportDetails.unmatchedExamples.join(" | ")}
                              </div>
                            ) : null}
                          </div>
                        ) : null}
                      </div>
                    ) : null}
                  </div>

                  <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(4, minmax(0, 1fr))", "1fr 1fr", "1fr"), gap: 12 }}>
                    <MiniStat label="Shipping value" value={formatTZS(filteredShippingSummary.totalValue)} tone="blue" sub="Filtered shipping queue value" />
                    <MiniStat label="In flow" value={filteredShippingSummary.inFlow} tone="amber" sub="Preparing or in transit" />
                    <MiniStat label="Delivered" value={filteredShippingSummary.delivered} tone="green" sub="Completed shipping orders" />
                    <MiniStat label="Returned" value={filteredShippingSummary.returned} tone="amber" sub="Exceptions to review" />
                  </div>

                  <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("minmax(260px, 1fr) 180px 120px", "1fr 1fr", "1fr"), gap: 12 }}>
                    <input
                      style={styles.input}
                      value={shippingListFilters.search}
                      onChange={(e) => setShippingListFilters((prev) => ({ ...prev, search: e.target.value }))}
                      placeholder="Search order id, phone, customer, product..."
                    />
                    <select
                      style={styles.input}
                      value={shippingListFilters.status}
                      onChange={(e) => setShippingListFilters((prev) => ({ ...prev, status: e.target.value }))}
                    >
                      <option value="all">All shipping statuses</option>
                      {shippingStatusCatalog
                        .map((status) => (
                          <option key={status.value} value={status.value}>
                            {status.label}
                          </option>
                        ))}
                    </select>
                    <select
                      style={styles.input}
                      value={shippingListFilters.pageSize}
                      onChange={(e) => setShippingListFilters((prev) => ({ ...prev, pageSize: Number(e.target.value) }))}
                    >
                      <option value={10}>10 / page</option>
                      <option value={25}>25 / page</option>
                      <option value={50}>50 / page</option>
                      <option value={100}>100 / page</option>
                    </select>
                  </div>

                  <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("minmax(240px, 1fr) auto auto", "1fr 1fr", "1fr"), gap: 10, alignItems: "end" }}>
                    <div style={styles.fieldBlock}>
                      <label style={styles.fieldLabel}>Bulk shipping status</label>
                      <select style={styles.input} value={bulkShippingStatus} onChange={(e) => setBulkShippingStatus(e.target.value)}>
                        {shippingStatusCatalog.map((status) => (
                          <option key={status.value} value={status.value}>
                            {status.label}
                          </option>
                        ))}
                      </select>
                    </div>
                    <button style={styles.btnSecondary} disabled={selectedShippingIds.length === 0} onClick={updateShippingBulkStatus}>
                      Apply shipping status
                    </button>
                    <button style={{ ...styles.btnSecondary, background: "#fef2f2", color: red, border: "1px solid #fecaca" }} disabled={selectedShippingIds.length === 0} onClick={deleteSelectedShippingOrders}>
                      Delete selected
                    </button>
                  </div>

                  {historyTargetCustomer ? (
                    <div style={{ ...styles.softStat, display: "grid", gap: 10 }}>
                      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12 }}>
                        <div>
                          <div style={{ fontWeight: 800 }}>Order history: {historyTargetCustomer.customerName}</div>
                          <div style={{ color: textSoft, fontSize: 12, marginTop: 4 }}>{historyTargetCustomer.id} | {historyTargetCustomer.sourceOrderId || "Manual reference"}</div>
                        </div>
                        <button style={styles.btnSecondary} onClick={() => setCustomerHistoryTargetId("")}>Close history</button>
                      </div>
                      <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(4, minmax(0, 1fr))", "1fr 1fr", "1fr"), gap: 10 }}>
                        <MiniStat label="Source lead" value={historyTargetCustomer.leadSource || "manual"} tone="blue" sub={historyTargetCustomer.campaignName || "No campaign"} />
                        <MiniStat label="Priority" value={formatStatusLabel(historyTargetCustomer.priority || "normal")} tone="amber" sub={`${formatStatusLabel(historyTargetCustomer.customerType || "new")} customer`} />
                        <MiniStat label="Shipping" value={historyTargetCustomer.carrierName || "Not assigned"} tone="green" sub={historyTargetCustomer.trackingNumber || "No tracking"} />
                        <MiniStat label="Business reason" value={historyTargetCustomer.cancelReason || historyTargetCustomer.unreachedReason || historyTargetCustomer.returnReason || "None"} dark sub="Cancel / unreached / return" />
                      </div>
                      <div style={{ display: "grid", gap: 8 }}>
                        {historyTargetCustomer.history?.length ? historyTargetCustomer.history.slice(0, 8).map((entry) => (
                          <div key={entry.id} style={{ padding: "10px 12px", borderRadius: 12, background: "rgba(255,255,255,0.82)", border: `1px solid ${cardBorder}` }}>
                            <div style={{ fontWeight: 700, fontSize: 13 }}>{formatStatusLabel(entry.action)}</div>
                            <div style={{ color: textSoft, fontSize: 12, marginTop: 4 }}>{entry.at ? new Date(entry.at).toLocaleString() : "No date"} | {entry.source}</div>
                            <div style={{ color: textMain, fontSize: 13, marginTop: 6 }}>{entry.details || "No details"}</div>
                          </div>
                        )) : <div style={{ color: textSoft }}>No history saved yet.</div>}
                      </div>
                    </div>
                  ) : null}

                  <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap" }}>
                    <div style={{ color: textSoft, fontSize: 13 }}>
                      Showing {paginatedShippingRows.length} of {compactShippingRows.length} shipping orders | Selected {selectedShippingIds.length}
                    </div>
                    <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
                      <button
                        style={styles.btnSecondary}
                        disabled={shippingListPage <= 1}
                        onClick={() => setShippingListPage((prev) => Math.max(1, prev - 1))}
                      >
                        Previous
                      </button>
                      <span style={{ color: textSoft, fontSize: 13, fontWeight: 700 }}>
                        Page {shippingListPage} / {shippingListPageCount}
                      </span>
                      <button
                        style={styles.btnSecondary}
                        disabled={shippingListPage >= shippingListPageCount}
                        onClick={() => setShippingListPage((prev) => Math.min(shippingListPageCount, prev + 1))}
                      >
                        Next
                      </button>
                    </div>
                  </div>

                  <div style={{ overflowX: "auto", border: `1px solid ${cardBorder}`, borderRadius: 22, background: "linear-gradient(180deg, rgba(255,255,255,0.96), rgba(248,244,238,0.9))" }}>
                    <table style={{ width: "100%", borderCollapse: "separate", borderSpacing: 0 }}>
                      <thead>
                        <tr>
                          {["Select", "Customer", "Product", "Reference", "Order", "Owner", "Shipping Status", "History", "Last Shipping Import", "Actions"].map((head) => (
                            <th
                              key={head}
                              style={{
                                textAlign: head === "Actions" ? "right" : "left",
                                padding: "12px 10px",
                                color: textSoft,
                                fontSize: 12,
                                fontWeight: 800,
                                letterSpacing: 0.45,
                                textTransform: "uppercase",
                                borderBottom: `1px solid ${cardBorder}`,
                                background: "rgba(247, 243, 237, 0.92)",
                                whiteSpace: "nowrap",
                              }}
                            >
                              {head === "Select" ? (
                                <input
                                  ref={selectAllShippingRef}
                                  type="checkbox"
                                  checked={allFilteredShippingSelected}
                                  onChange={() =>
                                    setSelectedShippingIds((prev) => {
                                      const next = new Set(prev);
                                      if (allFilteredShippingSelected) {
                                        filteredShippingIds.forEach((id) => next.delete(id));
                                      } else {
                                        filteredShippingIds.forEach((id) => next.add(id));
                                      }
                                      return Array.from(next);
                                    })
                                  }
                                />
                              ) : (
                                head
                              )}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {paginatedShippingRows.map((customer, index) => (
                          <tr key={customer.id} style={{ background: index % 2 === 0 ? "rgba(255,255,255,0.7)" : "rgba(250,247,242,0.82)" }}>
                            <td style={{ padding: "10px 10px", borderBottom: `1px solid ${cardBorder}`, width: 54 }}>
                              <input
                                type="checkbox"
                                checked={selectedShippingIdSet.has(customer.id)}
                                onChange={() =>
                                  setSelectedShippingIds((prev) =>
                                    prev.includes(customer.id) ? prev.filter((id) => id !== customer.id) : [...prev, customer.id]
                                  )
                                }
                              />
                            </td>
                            <td style={{ padding: "10px 10px", borderBottom: `1px solid ${cardBorder}`, minWidth: 220 }}>
                              <div style={{ fontWeight: 800, fontSize: 14 }}>{customer.customerName}</div>
                              <div style={{ color: textSoft, fontSize: 11, marginTop: 2 }}>{customer.id} | {customer.phone}</div>
                              <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginTop: 6 }}>
                                {customer.carrierName ? (
                                  <span style={{ ...styles.badge, background: "rgba(31,143,95,0.08)", color: green, border: "1px solid rgba(31,143,95,0.12)" }}>
                                    {customer.carrierName}
                                  </span>
                                ) : null}
                                {customer.trackingNumber ? (
                                  <span style={{ ...styles.badge, background: "rgba(29,95,208,0.08)", color: accent, border: "1px solid rgba(29,95,208,0.12)" }}>
                                    {customer.trackingNumber}
                                  </span>
                                ) : null}
                              </div>
                            </td>
                            <td style={{ padding: "10px 10px", borderBottom: `1px solid ${cardBorder}`, minWidth: 180 }}>
                              <div style={{ fontWeight: 700, fontSize: 13 }}>{customer.productName}</div>
                              <div style={{ color: textSoft, fontSize: 11, marginTop: 2 }}>{customer.city || "N/A"}</div>
                            </td>
                            <td style={{ padding: "10px 10px", borderBottom: `1px solid ${cardBorder}`, minWidth: 180 }}>
                              <div style={{ fontWeight: 700, fontSize: 13 }}>{customer.sourceOrderId || customer.id}</div>
                              <div style={{ color: textSoft, fontSize: 11, marginTop: 2 }}>
                                {customer.importSource || "manual"}
                              </div>
                            </td>
                            <td style={{ padding: "10px 10px", borderBottom: `1px solid ${cardBorder}`, minWidth: 170 }}>
                              <div style={{ color: textMain, fontWeight: 700, fontSize: 13 }}>{customer.orderDate}</div>
                              <div style={{ color: textSoft, fontSize: 11, marginTop: 2 }}>
                                Qty {customer.quantity} | {formatTZS(customer.totalValue)}
                              </div>
                              <div style={{ color: textSoft, fontSize: 11, marginTop: 4 }}>
                                ETA {customer.expectedDeliveryDate || "N/A"} {customer.actualDeliveryDate ? `| Delivered ${customer.actualDeliveryDate}` : ""}
                              </div>
                            </td>
                            <td style={{ padding: "10px 10px", borderBottom: `1px solid ${cardBorder}`, minWidth: 180 }}>
                              <select
                                style={{ ...styles.input, padding: "8px 10px", minWidth: 0, fontSize: 12 }}
                                value={customer.assignedTo || ""}
                                onChange={(e) => assignCustomerOwner(customer.id, e.target.value)}
                              >
                                <option value="">No owner</option>
                                {teamRoster.map((member) => (
                                  <option key={member} value={member}>
                                    {member}
                                  </option>
                                ))}
                              </select>
                            </td>
                            <td style={{ padding: "10px 10px", borderBottom: `1px solid ${cardBorder}`, minWidth: 190 }}>
                              <div style={{ display: "grid", gap: 8 }}>
                                <span style={getStatusBadgeStyle(customer.status)}>
                                  {customer.statusLabel}
                                </span>
                                <select
                                  style={{ ...styles.input, padding: "8px 10px", minWidth: 0, fontSize: 12 }}
                                  value={customer.shippingStatus || "to-prepare"}
                                  onChange={(e) => updateCustomerShippingStatus(customer.id, e.target.value)}
                                >
                                  {shippingStatusCatalog.map((status) => (
                                    <option key={status.value} value={status.value}>
                                      {status.label}
                                    </option>
                                  ))}
                                </select>
                                <button
                                  style={{ ...styles.btnSecondary, background: "#fef2f2", color: red, border: "1px solid #fecaca", padding: "8px 10px", fontSize: 12 }}
                                  onClick={() => deleteCustomerOrder(customer.id)}
                                >
                                  Delete Order
                                </button>
                              </div>
                            </td>
                            <td style={{ padding: "10px 10px", borderBottom: `1px solid ${cardBorder}`, minWidth: 110 }}>
                              <button style={styles.btnSecondary} onClick={() => setCustomerHistoryTargetId(customer.id)}>
                                View
                              </button>
                            </td>
                            <td style={{ padding: "10px 10px", borderBottom: `1px solid ${cardBorder}`, minWidth: 180 }}>
                              <div style={{ color: textMain, fontWeight: 700, fontSize: 13 }}>{customer.lastShippingImportLabel}</div>
                              <div style={{ color: textSoft, fontSize: 11, marginTop: 2 }}>
                                {customer.importSource === "excel" ? "Imported from Excel" : "Shipping sync from Excel"}
                              </div>
                            </td>
                            <td style={{ padding: "10px 10px", borderBottom: `1px solid ${cardBorder}`, minWidth: 140, textAlign: "right" }}>
                              <button
                                style={{ ...styles.btnSecondary, background: "#fef2f2", color: red, border: "1px solid #fecaca", padding: "8px 10px", fontSize: 12 }}
                                onClick={() => deleteCustomerOrder(customer.id)}
                              >
                                Delete Order
                              </button>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>

                    {compactShippingRows.length === 0 ? (
                      <div style={{ padding: 24, color: textSoft }}>No shipping orders available yet. Confirmed or post-confirmation orders will appear here.</div>
                    ) : null}
                  </div>
                </div>
              </div>
            </div>
          )}

{["tracking", "financeHub"].includes(activePage) && (
            <div style={{ display: "grid", gap: 20 }}>
              <PageHeader
                eyebrow="Ads & Tracking"
                title="Campaign mapping and product tracking"
                description="Track ads spend, Meta mapping and live COD funnel metrics at product level without mixing shipping or finance controls."
              />
              <div style={{ display: "none", gap: 10, flexWrap: "wrap" }}>
                <button style={activePage === "tracking" ? styles.btnPrimary : styles.btnSecondary} onClick={() => setActivePage("tracking")}>Tracking</button>
                <button style={activePage === "serviceSum" ? styles.btnPrimary : styles.btnSecondary} onClick={() => setActivePage("serviceSum")}>Service Sum</button>
                <button style={activePage === "situations" ? styles.btnPrimary : styles.btnSecondary} onClick={() => setActivePage("situations")}>Rentabilité</button>
                <button style={activePage === "profitCenter" ? styles.btnPrimary : styles.btnSecondary} onClick={() => setActivePage("profitCenter")}>Profit Center</button>
              </div>
              <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(4, minmax(0, 1fr))", "repeat(2, minmax(0, 1fr))", "1fr"), gap: 16 }}>
                <KpiCard icon={<ClipboardList size={18} />} title="Tracking rows" value={trackingSummary.rows} sub="Manual and Meta-synced spend rows" />
                <KpiCard icon={<Wallet size={18} />} title="Cumulative ad spend" value={formatTZS(totalCumulativeSpendTsh || trackingSummary.spend)} sub={totalCumulativeSpendTsh > 0 ? `${cumulativeCampaigns.length} campaigns tracked across all imports` : `${trackingSummary.orders} customer orders synced`} valueColor={accent} />
                <KpiCard icon={<TrendingUp size={18} />} title="Revenue" value={formatTZS(trackingSummary.revenue)} sub={`${trackingSummary.delivered} delivered units from orders`} valueColor={green} />
                <KpiCard icon={<Rocket size={18} />} title="Profit" value={formatTZS(trackingSummary.profit)} sub="Orders automate revenue and stock impact" valueColor={trackingSummary.profit >= 0 ? green : red} />
              </div>

              <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                {[["meta", "Meta Ads"], ["tracking", "Tracking"], ["cpl", "CPL Tracker"]].map(([id, label]) => (
                  <button
                    key={id}
                    style={{ ...(trackingSubTab === id ? styles.btnPrimary : styles.btnSecondary), borderRadius: 18, padding: "10px 20px" }}
                    onClick={() => setTrackingSubTab(id)}
                  >
                    {label}
                  </button>
                ))}
              </div>

              {trackingSubTab === "meta" && (
              <div style={{ ...styles.card, padding: 22 }}>
                <div style={styles.sectionHeader}>
                  <div>
                    <div style={styles.sectionEyebrow}>Meta Ads bridge</div>
                    <div style={{ fontSize: 24, fontWeight: 900, marginTop: 8 }}>Optional Meta Ads sync</div>
                    <div style={{ color: textSoft, marginTop: 6, lineHeight: 1.6 }}>
                      This block is only for importing data from Meta Ads Manager. If you do not want Meta sync, you can ignore this card and use the manual `Tracking` section just below.
                    </div>
                  </div>
                  <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                    <button
                      style={{
                        ...styles.btnSecondary,
                        padding: "13px 18px",
                        borderRadius: 18,
                        background: "linear-gradient(180deg, rgba(255,255,255,0.98), rgba(242,246,255,0.92))",
                        border: "1px solid rgba(29,95,208,0.16)",
                        color: accent,
                        boxShadow: "0 14px 28px rgba(29,95,208,0.08)",
                      }}
                      onClick={loadMetaAdAccounts}
                      disabled={metaAdsLoading.accounts}
                    >
                      {metaAdsLoading.accounts ? "Loading accounts..." : "Load accounts"}
                    </button>
                    <button
                      style={{
                        ...styles.btnSecondary,
                        padding: "13px 18px",
                        borderRadius: 18,
                        background: "linear-gradient(180deg, rgba(255,255,255,0.98), rgba(245,250,244,0.92))",
                        border: "1px solid rgba(31,143,95,0.16)",
                        color: green,
                        boxShadow: "0 14px 28px rgba(31,143,95,0.08)",
                      }}
                      onClick={() => refreshMetaInsights({ syncTotalSpend: true })}
                      disabled={metaAdsLoading.insights}
                    >
                      {metaAdsLoading.insights ? "Refreshing..." : "Refresh insights"}
                    </button>
                    <button
                      style={{
                        ...styles.btnPrimary,
                        padding: "13px 20px",
                        borderRadius: 18,
                        background: "linear-gradient(135deg, #0f172a, #1d5fd0, #2c7be5)",
                        boxShadow: "0 18px 34px rgba(29, 95, 208, 0.28)",
                      }}
                      onClick={applyMetaInsightsToApp}
                      disabled={metaAdsLoading.apply || !metaCampaignRows.length}
                    >
                      {metaAdsLoading.apply ? "Importing..." : "Import to app"}
                    </button>
                  </div>
                </div>

                <div style={{ ...styles.softStat, marginBottom: 16, border: "1px solid rgba(29,95,208,0.14)", background: "linear-gradient(180deg, rgba(239,245,255,0.9), rgba(255,255,255,0.92))" }}>
                  <div style={{ fontSize: 12, fontWeight: 800, letterSpacing: 0.45, textTransform: "uppercase", color: accent }}>Important</div>
                  <div style={{ marginTop: 8, color: textMain, lineHeight: 1.6 }}>
                    `Tracking` works in two ways:
                    manual mode below where you type `Ad spend` yourself, and optional `Meta Ads` mode here if you want automatic import from Facebook Ads Manager.
                  </div>
                </div>

                <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("1fr 180px", "1fr", "1fr"), gap: 12, marginBottom: 16 }}>
                  <div style={{ ...styles.softStat, border: "1px solid rgba(31,143,95,0.18)", background: "linear-gradient(180deg, rgba(236,253,245,0.88), rgba(255,255,255,0.94))" }}>
                    <div style={{ fontSize: 12, fontWeight: 800, letterSpacing: 0.45, textTransform: "uppercase", color: green }}>Auto sync</div>
                    <div style={{ marginTop: 8, fontWeight: 800, fontSize: 18, color: textMain }}>
                      {metaAdsState.autoSync ? "Live sync active" : "Manual sync only"}
                    </div>
                    <div style={{ marginTop: 6, color: textSoft, lineHeight: 1.5 }}>
                      While this page stays open, the app will import Meta changes automatically every {metaAdsState.autoSyncIntervalMinutes} minute(s) without page refresh.
                    </div>
                  </div>
                  <button
                    style={{
                      ...(metaAdsState.autoSync ? styles.btnPrimary : styles.btnSecondary),
                      borderRadius: 18,
                      minHeight: 88,
                    }}
                    onClick={() => setMetaAdsState((prev) => ({ ...prev, autoSync: !prev.autoSync }))}
                  >
                    {metaAdsState.autoSync ? "Pause auto sync" : "Activate auto sync"}
                  </button>
                </div>

                <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("1.25fr 1fr 1.1fr", "1fr 1fr", "1fr"), gap: 12, marginBottom: 16 }}>
                  <div style={styles.fieldBlock}>
                    <label style={styles.fieldLabel}>Meta access token</label>
                    <input
                      style={styles.input}
                      type="password"
                      placeholder="EAAB..."
                      value={metaAdsState.accessToken}
                      onChange={(e) => setMetaAdsState((prev) => ({ ...prev, accessToken: e.target.value }))}
                    />
                  </div>
                  <div style={styles.fieldBlock}>
                    <label style={styles.fieldLabel}>Ad account ID</label>
                    <input
                      style={styles.input}
                      list="meta-ad-accounts"
                      placeholder="act_123456789 or 123456789"
                      value={metaAdsState.accountId}
                      onChange={(e) => setMetaAdsState((prev) => ({ ...prev, accountId: e.target.value }))}
                    />
                    <datalist id="meta-ad-accounts">
                      {metaAdsAccounts.map((account) => (
                        <option key={account.id} value={account.id}>
                          {account.name} ({account.currency || "USD"})
                        </option>
                      ))}
                    </datalist>
                    <div style={{ color: textSoft, fontSize: 12, marginTop: 6 }}>
                      You can load accounts automatically, or paste your ad account ID manually if you already know it.
                    </div>
                  </div>
                  <div style={styles.fieldBlock}>
                    <label style={styles.fieldLabel}>Date range</label>
                    <MetaDateRangePicker
                      key={`${metaAdsState.dateStart}-${metaAdsState.dateEnd}`}
                      value={{ start: metaAdsState.dateStart, end: metaAdsState.dateEnd }}
                      onApply={(range) => setMetaAdsState((prev) => ({ ...prev, dateStart: range.start, dateEnd: range.end }))}
                      responsiveColumns={responsiveColumns}
                    />
                    <div style={{ color: textSoft, fontSize: 12, marginTop: 6 }}>
                      Click this single block to set both the start and end dates for the import window.
                    </div>
                  </div>
                </div>

                <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(4, minmax(0, 1fr))", "repeat(2, minmax(0, 1fr))", "1fr"), gap: 12 }}>
                  <MiniStat label="Selected account" value={selectedMetaAccount?.name || metaAdsState.accountId || "Not loaded"} tone="blue" sub={selectedMetaAccount ? `${selectedMetaAccount.currency || "USD"} | ${selectedMetaAccount.timezoneName || "Meta account"}` : metaAdsState.accountId ? "Manual ad account ID" : "Load accounts or paste account ID manually"} />
                  <MiniStat label="Meta spend" value={formatMetaMoney(metaDashboardMetrics.spend)} tone="amber" sub={`${metaDashboardMetrics.campaigns} campaign rows in range`} />
                  <MiniStat label="Tracked leads" value={metaDashboardMetrics.leads} tone="green" sub={`${formatMetaLeadSourceLabel(metaDashboardMetrics.trackedLeadSource)} | Actual leads ${formatInteger(metaDashboardMetrics.actualLeads)}`} />
                  <MiniStat label="Last import" value={metaAdsState.lastSyncAt ? new Date(metaAdsState.lastSyncAt).toLocaleString() : "Not imported"} tone="blue" sub={metaAdsState.lastSyncSummary ? `${metaAdsState.lastSyncSummary.totalCampaigns || metaAdsState.lastSyncSummary.matchedRows || 0} campaigns | ${metaAdsState.lastSyncSummary.matchedProducts || 0} products | Unmapped: ${formatTZS(metaAdsState.lastSyncSummary.unmappedSpendTzs || 0)}` : "No Meta import yet"} />
                </div>

                <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(7, minmax(0, 1fr))", "1fr 1fr", "1fr"), gap: 12, marginTop: 12 }}>
                  <MiniStat label="Impressions" value={formatInteger(metaDashboardMetrics.impressions)} tone="blue" sub="Total ad views in range" />
                  <MiniStat label="Reach" value={formatInteger(metaDashboardMetrics.reach)} tone="green" sub="Unique people reached" />
                  <MiniStat label="Clicks (all)" value={formatInteger(metaDashboardMetrics.clicks)} tone="amber" sub="All click types from Meta" />
                  <MiniStat label="Link clicks" value={formatInteger(metaDashboardMetrics.inlineLinkClicks)} tone="blue" sub={`${formatInteger(metaDashboardMetrics.uniqueInlineLinkClicks)} unique link clicks`} />
                  <MiniStat label="Landing page views" value={formatInteger(metaDashboardMetrics.landingPageViews)} tone="green" sub={`CTR ${metaDashboardMetrics.ctr.toFixed(2)}%`} />
                  <MiniStat label="CPC" value={formatMetaMoney(metaDashboardMetrics.cpc)} tone="blue" sub="Cost per link click" />
                  <MiniStat label="CPM" value={formatMetaMoney(metaDashboardMetrics.cpm)} tone="amber" sub={`CPP ${formatMetaMoney(metaDashboardMetrics.cpp)} | Freq ${metaDashboardMetrics.frequency.toFixed(2)}`} />
                </div>

                {/* Per-product campaign breakdown */}
                {Object.keys(mappedCampaignsByProduct).length > 0 && (
                  <div style={{ ...styles.softStat, marginTop: 16, border: "1px solid rgba(31,143,95,0.16)", background: "linear-gradient(180deg, rgba(236,253,245,0.6), rgba(255,255,255,0.9))" }}>
                    <div style={{ display: "flex", justifyContent: "space-between", gap: 12, flexWrap: "wrap", marginBottom: 14 }}>
                      <div>
                        <div style={{ fontSize: 11, fontWeight: 800, letterSpacing: 0.45, textTransform: "uppercase", color: green }}>Auto-mapped</div>
                        <div style={{ marginTop: 6, fontSize: 20, fontWeight: 900 }}>Campaigns by product</div>
                        <div style={{ marginTop: 4, color: textSoft, fontSize: 13 }}>
                          {metaCampaignRows.filter((r) => r.mappedProductId).length} campaign(s) mapped to {Object.keys(mappedCampaignsByProduct).length} product(s)
                        </div>
                      </div>
                    </div>
                    <div style={{ display: "grid", gap: 10 }}>
                      {Object.entries(mappedCampaignsByProduct).map(([productId, campaigns]) => {
                        const product = products.find((p) => p.id === productId);
                        const totalSpendTzs = campaigns.reduce((sum, r) => sum + (r.spendTzs || 0), 0);
                        return (
                          <div key={productId} style={{ padding: 14, borderRadius: 14, border: `1px solid ${cardBorder}`, background: "rgba(255,255,255,0.84)" }}>
                            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                              <div style={{ fontWeight: 800, fontSize: 15 }}>{product?.name || productId}</div>
                              <div style={{ fontWeight: 800, color: accent }}>{formatTZS(Math.round(totalSpendTzs))}</div>
                            </div>
                            <div style={{ display: "grid", gap: 5 }}>
                              {campaigns.map((row) => (
                                <div key={row.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", fontSize: 12, color: textSoft }}>
                                  <div style={{ flex: 1, minWidth: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", marginRight: 8 }}>{row.campaignName || "—"}</div>
                                  <div style={{ display: "flex", gap: 8, alignItems: "center", flexShrink: 0 }}>
                                    {row.autoMapped && <span style={{ ...styles.badge, background: "rgba(31,143,95,0.1)", color: green, fontSize: 10 }}>auto</span>}
                                    {row.manuallyMapped && <span style={{ ...styles.badge, background: "rgba(29,95,208,0.1)", color: accent, fontSize: 10 }}>manual</span>}
                                    <span style={{ fontWeight: 700, color: textMain }}>{formatTZS(Math.round(row.spendTzs || 0))}</span>
                                  </div>
                                </div>
                              ))}
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                )}

                {unmappedMetaCampaignRows.length ? (
                  <div style={{ ...styles.softStat, marginTop: 16 }}>
                    <div style={{ display: "flex", justifyContent: "space-between", gap: 12, flexWrap: "wrap", marginBottom: 14 }}>
                      <div>
                        <div style={{ fontSize: 11, fontWeight: 800, letterSpacing: 0.45, textTransform: "uppercase", color: amber }}>Needs attention</div>
                        <div style={{ marginTop: 8, fontSize: 20, fontWeight: 900 }}>Unmapped campaigns</div>
                        <div style={{ marginTop: 6, color: textSoft, lineHeight: 1.55 }}>
                          Campaigns are auto-mapped when their name contains a product mapping code (e.g. <strong>TZ | DSP4 | COLD</strong> maps to the product with code <strong>DSP4</strong>). Assign manually below or skip — skipped spend is tracked as global unmapped cost.
                        </div>
                      </div>
                      <div style={{ display: "flex", flexDirection: "column", gap: 6, alignItems: "flex-end", flexShrink: 0 }}>
                        <div style={{ ...styles.badge, background: "rgba(199,131,34,0.12)", color: amber, border: "1px solid rgba(199,131,34,0.18)" }}>
                          {unmappedMetaCampaignRows.length} campaigns unassigned
                        </div>
                        <div style={{ fontSize: 12, color: textSoft, fontWeight: 600 }}>
                          Unmapped spend: {formatTZS(Math.round(unmappedMetaCampaignRows.reduce((s, r) => s + (r.spendTzs || 0), 0)))}
                        </div>
                      </div>
                    </div>

                    <div style={{ display: "grid", gap: 10 }}>
                      {unmappedMetaCampaignRows.slice(0, 10).map((row) => (
                        <div
                          key={row.id}
                          style={{
                            display: "grid",
                            gridTemplateColumns: responsiveColumns("minmax(0, 1.4fr) minmax(220px, 0.8fr)", "1fr", "1fr"),
                            gap: 12,
                            alignItems: "center",
                            padding: 14,
                            borderRadius: 16,
                            border: `1px solid ${cardBorder}`,
                            background: "rgba(255,255,255,0.84)",
                          }}
                        >
                          <div style={{ minWidth: 0 }}>
                            <div style={{ fontWeight: 800, color: textMain }}>{row.campaignName || "Unnamed campaign"}</div>
                            <div style={{ marginTop: 6, color: textSoft, fontSize: 12, lineHeight: 1.5 }}>
                              Spend: <strong>{formatTZS(Math.round(row.spendTzs || 0))}</strong> | Code detected: <strong>{row.productCode || "none"}</strong> | Phase: <strong>{row.phase || "N/A"}</strong> | Version: <strong>{row.version || "N/A"}</strong>
                            </div>
                          </div>
                          <select
                            style={styles.input}
                            value={metaAdsState.campaignMappings[row.id] || ""}
                            onChange={(e) =>
                              setMetaAdsState((prev) => ({
                                ...prev,
                                campaignMappings: {
                                  ...prev.campaignMappings,
                                  [row.id]: e.target.value,
                                },
                              }))
                            }
                          >
                            <option value="">Skip (add to unmapped spend)</option>
                            {products.map((product) => (
                              <option key={product.id} value={product.id}>
                                [{product.mappingCode || product.id}] {product.name}
                              </option>
                            ))}
                          </select>
                        </div>
                      ))}
                    </div>
                  </div>
                ) : null}

                {metaAdsNotice ? (
                  <div style={{ marginTop: 16, padding: "14px 16px", borderRadius: 18, border: `1px solid ${cardBorder}`, background: "linear-gradient(180deg, rgba(255,255,255,0.92), rgba(250,247,242,0.88))", color: textMain, boxShadow: "0 10px 22px rgba(23,32,51,0.05)" }}>
                    {metaAdsNotice}
                  </div>
                ) : null}

              </div>
              )}

              {trackingSubTab === "tracking" && (
              <div style={{ ...styles.card, padding: 22 }}>
              <div style={styles.sectionHeader}>
                <div>
                  <div style={styles.sectionEyebrow}>Performance engine</div>
                  <div style={{ fontSize: 24, fontWeight: 900, marginTop: 8 }}>Tracking</div>
                  <div style={{ color: textSoft, marginTop: 6, lineHeight: 1.6 }}>Ad spend stays manual here, while confirmations, deliveries, revenue and stock now sync automatically from customer orders.</div>
                </div>
                <button
                  style={styles.btnSecondary}
                  onClick={() =>
                    setTracking((prev) => {
                      if (!products.length) {
                        alert("Add a product first before adding tracking.");
                        return prev;
                      }

                      return [
                        ...prev,
                        {
                          id: buildNextId(prev, "T"),
                          productId: products[0].id,
                          adSpend: 0,
                          orders: 0,
                          confirmed: 0,
                          delivered: 0,
                          dateStart: getTodayString(),
                          dateEnd: getTodayString(),
                        },
                      ];
                    })
                  }
                >
                  <ClipboardList size={16} style={{ marginRight: 8, verticalAlign: "middle" }} />
                  Add tracking
                </button>
              </div>

              <div style={{ display: "grid", gap: 16 }}>
                {tracking.map((t, i) => {
                  const automatedProduct = productDashboardMap[t.productId];
                  const automatedCalc = automatedProduct
                    ? {
                        decision: automatedProduct.decision,
                        profit: automatedProduct.profit,
                        cpa: automatedProduct.cpa,
                        revenue: automatedProduct.revenue,
                        confirmRate: automatedProduct.confirmRate,
                        deliveryRate: automatedProduct.deliveryRate,
                        roas: automatedProduct.roas,
                        orders: automatedProduct.orders,
                        confirmedOrders: automatedProduct.confirmed,
                        deliveredUnits: automatedProduct.deliveredUnits,
                        reservedUnits: automatedProduct.reservedStock,
                        availableUnits: automatedProduct.availableStock,
                      }
                    : {
                        decision: "WATCH",
                        profit: 0,
                        cpa: 0,
                        revenue: 0,
                        confirmRate: 0,
                        deliveryRate: 0,
                        roas: 0,
                        orders: 0,
                        confirmedOrders: 0,
                        deliveredUnits: 0,
                        reservedUnits: 0,
                        availableUnits: 0,
                      };
                  const linkedProduct = products.find((p) => p.id === t.productId);
                  return (
                    <div key={t.id} style={{ border: `1px solid ${cardBorder}`, borderRadius: 22, padding: 18, background: "linear-gradient(180deg, rgba(255,255,255,0.96), rgba(249,246,241,0.9))", boxShadow: "0 14px 28px rgba(23, 32, 51, 0.06)" }}>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 16, flexWrap: "wrap", marginBottom: 14 }}>
                        <div>
                          <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
                            <div style={{ fontSize: 18, fontWeight: 800 }}>{linkedProduct?.name || "Tracking row"}</div>
                            <span style={{ ...styles.badge, background: "rgba(29,95,208,0.08)", color: accent, border: "1px solid rgba(29,95,208,0.12)" }}>{t.id}</span>
                            {t.metaManaged ? (
                              <span style={{ ...styles.badge, background: "rgba(31,143,95,0.12)", color: green, border: "1px solid rgba(31,143,95,0.18)" }}>
                                Meta Sync
                              </span>
                            ) : null}
                          </div>
                          <div style={{ color: textSoft, marginTop: 6 }}>
                            {t.metaManaged
                              ? `Meta imported spend for ${t.metaSince || "selected range"} -> ${t.metaUntil || "selected range"}.`
                              : "Ad spend is manual here. Orders, confirmations, delivery, revenue and stock are auto-synced from the app."}
                          </div>
                        </div>
                        <div style={getDecisionStyle(automatedCalc.decision || "WATCH")}>{automatedCalc.decision || "WATCH"}</div>
                      </div>

                      <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("1.4fr 1fr 1fr", "1fr 1fr", "1fr"), gap: 12 }}>
                        <select style={styles.input} value={t.productId} onChange={(e) => {
                          const next = [...tracking];
                          next[i].productId = e.target.value;
                          setTracking(next);
                        }}>
                          {products.map((p) => <option key={p.id} value={p.id}>{p.name || p.id}</option>)}
                        </select>
                        <input style={styles.input} type="number" placeholder="Ad spend" value={t.adSpend} onChange={(e) => {
                          const next = [...tracking];
                          next[i].adSpend = Number(e.target.value || 0);
                          setTracking(next);
                        }} />
                        <div style={{ ...styles.softStat, display: "grid", gap: 6 }}>
                          <div style={{ fontSize: 11, fontWeight: 800, letterSpacing: 0.45, textTransform: "uppercase", color: textSoft }}>Auto pipeline</div>
                          <div style={{ fontSize: 14, fontWeight: 800, color: textMain }}>
                            {automatedCalc.orders} orders | {automatedCalc.confirmedOrders} confirmed
                          </div>
                          <div style={{ fontSize: 12, color: textSoft }}>
                            {automatedCalc.deliveredUnits} delivered units | {automatedCalc.availableUnits} stock available
                          </div>
                        </div>
                      </div>

                      <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(6, minmax(0, 1fr))", "1fr 1fr", "1fr"), gap: 12, marginTop: 16 }}>
                        <MiniStat label="Profit" value={formatTZS(automatedCalc.profit || 0)} tone={(automatedCalc.profit || 0) >= 0 ? "green" : "amber"} sub="Auto net result" />
                        <MiniStat label="CPA" value={formatTZS(Math.round(automatedCalc.cpa || 0))} sub="Ad spend / delivered unit" />
                        <MiniStat label="Revenue" value={formatTZS(automatedCalc.revenue || 0)} tone="green" sub="Auto from delivered orders" />
                        <MiniStat label="Confirm rate" value={`${Math.round((automatedCalc.confirmRate || 0) * 100)}%`} tone="amber" sub={`${automatedCalc.confirmedOrders} confirmed orders`} />
                        <MiniStat label="Delivery rate" value={`${Math.round((automatedCalc.deliveryRate || 0) * 100)}%`} tone="blue" sub={`ROAS ${Number(automatedCalc.roas || 0).toFixed(2)}`} />
                        <MiniStat label="Reserved stock" value={automatedCalc.reservedUnits} tone="amber" sub={`${automatedCalc.availableUnits} available units`} />
                      </div>
                    </div>
                  );
                })}
                {tracking.length === 0 ? <div style={{ color: textSoft }}>No tracking rows yet.</div> : null}
              </div>
              </div>
              )}

              {trackingSubTab === "cpl" && (
              <div style={{ ...styles.card, padding: 22 }}>
                <div style={styles.sectionHeader}>
                  <div>
                    <div style={styles.sectionEyebrow}>Cumulative performance</div>
                    <div style={{ fontSize: 24, fontWeight: 900, marginTop: 8 }}>CPL Tracker</div>
                    <div style={{ color: textSoft, marginTop: 6, lineHeight: 1.6 }}>
                      Per-product cost per lead from all imported Meta campaigns. Data accumulates across imports — reimporting the same period updates rather than duplicates.
                    </div>
                  </div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 6, alignItems: "flex-end" }}>
                    <div style={{ ...styles.badge, background: "rgba(29,95,208,0.1)", color: accent }}>
                      {cumulativeCampaigns.length} campaigns total
                    </div>
                    <div style={{ fontSize: 12, color: textSoft, fontWeight: 600 }}>
                      Total: {formatTZS(totalCumulativeSpendTsh)}
                    </div>
                  </div>
                </div>

                {cplTrackerRows.length === 0 ? (
                  <div style={{ padding: "24px 0", color: textSoft }}>No campaign data yet. Import Meta Ads data first.</div>
                ) : (
                  <div style={{ overflowX: "auto", marginTop: 16 }}>
                    <table style={{ width: "100%", borderCollapse: "separate", borderSpacing: 0 }}>
                      <thead>
                        <tr>
                          {["Product", "Code", "Spend (USD)", "Spend (TZS)", "Leads", "Confirmed", "Delivered", "CPL (USD)", "CPL Confirmed", "CPL Delivered", "Confirm %", "Delivery %"].map((h) => (
                            <th key={h} style={{ textAlign: "left", padding: "12px 10px", color: textSoft, fontSize: 11, fontWeight: 800, letterSpacing: 0.4, textTransform: "uppercase", borderBottom: `1px solid ${cardBorder}`, background: "rgba(247,243,237,0.92)", whiteSpace: "nowrap" }}>
                              {h}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {cplTrackerRows.map((row, i) => (
                          <tr key={row.id} style={{ background: i % 2 === 0 ? "rgba(255,255,255,0.72)" : "rgba(250,247,242,0.8)" }}>
                            <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 800 }}>{row.name}</td>
                            <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}` }}>
                              <span style={{ ...styles.badge, background: "rgba(29,95,208,0.08)", color: accent, fontFamily: "monospace" }}>{row.mappingCode}</span>
                            </td>
                            <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 700 }}>{formatUSD(row.spendUsd)}</td>
                            <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}` }}>{formatTZS(Math.round(row.spendTsh))}</td>
                            <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}` }}>{row.leads}</td>
                            <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}` }}>{row.confirmed}</td>
                            <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}` }}>{row.delivered}</td>
                            <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}`, color: accent, fontWeight: 700 }}>{row.cpl > 0 ? formatUSD(row.cpl) : "—"}</td>
                            <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}` }}>{row.cplConfirmed > 0 ? formatUSD(row.cplConfirmed) : "—"}</td>
                            <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}` }}>{row.cplDelivered > 0 ? formatUSD(row.cplDelivered) : "—"}</td>
                            <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}` }}>{row.confirmRate > 0 ? `${Math.round(row.confirmRate * 100)}%` : "—"}</td>
                            <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}` }}>{row.deliveryRate > 0 ? `${Math.round(row.deliveryRate * 100)}%` : "—"}</td>
                          </tr>
                        ))}
                        <tr style={{ background: "rgba(247,243,237,0.95)", fontWeight: 900 }}>
                          <td style={{ padding: "14px 10px", borderTop: `2px solid ${cardBorder}`, fontWeight: 900 }}>TOTAL</td>
                          <td style={{ padding: "14px 10px", borderTop: `2px solid ${cardBorder}` }}>—</td>
                          <td style={{ padding: "14px 10px", borderTop: `2px solid ${cardBorder}`, color: accent, fontWeight: 900 }}>{formatUSD(cplTrackerGlobal.totalSpendUsd)}</td>
                          <td style={{ padding: "14px 10px", borderTop: `2px solid ${cardBorder}` }}>{formatTZS(Math.round(cplTrackerGlobal.totalSpendTsh))}</td>
                          <td style={{ padding: "14px 10px", borderTop: `2px solid ${cardBorder}` }}>{cplTrackerGlobal.totalLeads}</td>
                          <td style={{ padding: "14px 10px", borderTop: `2px solid ${cardBorder}` }}>{cplTrackerGlobal.totalConfirmed}</td>
                          <td style={{ padding: "14px 10px", borderTop: `2px solid ${cardBorder}` }}>{cplTrackerGlobal.totalDelivered}</td>
                          <td style={{ padding: "14px 10px", borderTop: `2px solid ${cardBorder}`, color: accent }}>{cplTrackerGlobal.cpl > 0 ? formatUSD(cplTrackerGlobal.cpl) : "—"}</td>
                          <td style={{ padding: "14px 10px", borderTop: `2px solid ${cardBorder}` }}>{cplTrackerGlobal.cplConfirmed > 0 ? formatUSD(cplTrackerGlobal.cplConfirmed) : "—"}</td>
                          <td style={{ padding: "14px 10px", borderTop: `2px solid ${cardBorder}` }}>{cplTrackerGlobal.cplDelivered > 0 ? formatUSD(cplTrackerGlobal.cplDelivered) : "—"}</td>
                          <td style={{ padding: "14px 10px", borderTop: `2px solid ${cardBorder}` }}>{cplTrackerGlobal.confirmRate > 0 ? `${Math.round(cplTrackerGlobal.confirmRate * 100)}%` : "—"}</td>
                          <td style={{ padding: "14px 10px", borderTop: `2px solid ${cardBorder}` }}>{cplTrackerGlobal.deliveryRate > 0 ? `${Math.round(cplTrackerGlobal.deliveryRate * 100)}%` : "—"}</td>
                        </tr>
                      </tbody>
                    </table>
                  </div>
                )}

                {cumulativeUnmappedSpendTsh > 0 && (
                  <div style={{ marginTop: 14, padding: "12px 16px", borderRadius: 14, background: "rgba(199,131,34,0.06)", border: "1px solid rgba(199,131,34,0.18)" }}>
                    <span style={{ fontWeight: 800, color: amber }}>Unmapped spend: {formatTZS(Math.round(cumulativeUnmappedSpendTsh))}</span>
                    <span style={{ color: textSoft, fontSize: 13, marginLeft: 10 }}>— campaigns not assigned to any product</span>
                  </div>
                )}
              </div>
              )}
            </div>
          )}

{activePage === "serviceSum" && (
            <div style={{ display: "grid", gap: 20 }}>
              <PageHeader
                eyebrow="Service Simulation"
                title="Test a product before launching it"
                description="Simulation only — does not affect real calculations. Use this to project profitability before committing ad spend."
              />

              <div style={{ ...styles.softStat, border: "1px solid rgba(199,131,34,0.22)", background: "linear-gradient(180deg, rgba(255,251,235,0.9), rgba(255,255,255,0.94))" }}>
                <div style={{ fontSize: 12, fontWeight: 800, letterSpacing: 0.45, textTransform: "uppercase", color: amber }}>Simulation only — does not affect real calculations</div>
                <div style={{ marginTop: 8, color: textMain, lineHeight: 1.6 }}>
                  These numbers are estimates only. Your real business results are in the Profit Center.
                </div>
              </div>

              <div style={{ ...styles.card, padding: 22 }}>
                <div style={styles.sectionHeader}>
                  <div>
                    <div style={styles.sectionEyebrow}>Simulation inputs</div>
                    <div style={{ fontSize: 22, fontWeight: 900, marginTop: 8 }}>Enter your product scenario</div>
                  </div>
                  <div style={{ display: "grid", gridTemplateColumns: "minmax(0, 1fr) minmax(0, 1fr)", gap: 10 }}>
                    <select style={styles.input} value={selectedService} onChange={(e) => setSelectedService(e.target.value)}>
                      <option value="standard">Standard</option>
                      <option value="codzoss">CODZOSS</option>
                    </select>
                    <select style={styles.input} value={selectedCountry} onChange={(e) => setSelectedCountry(e.target.value)}>
                      <option value="tanzania">Tanzania</option>
                      <option value="kenya">Kenya</option>
                    </select>
                  </div>
                </div>

                <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(3, minmax(0, 1fr))", "1fr 1fr", "1fr"), gap: 12, marginTop: 16 }}>
                  <div style={styles.fieldBlock}>
                    <label style={styles.fieldLabel}>Product Name</label>
                    <input style={styles.input} placeholder="e.g. Product A" value={simProductName} onChange={(e) => setSimProductName(e.target.value)} />
                  </div>
                  <div style={styles.fieldBlock}>
                    <label style={styles.fieldLabel}>Selling Price TZS</label>
                    <input style={styles.input} type="number" min="0" value={serviceForm.sellingPriceTzs} onChange={(e) => setServiceForm({ ...serviceForm, sellingPriceTzs: e.target.value })} />
                  </div>
                  <div style={styles.fieldBlock}>
                    <label style={styles.fieldLabel}>Cost per Unit TZS</label>
                    <input style={styles.input} type="number" min="0" value={serviceForm.productCostTzs} onChange={(e) => setServiceForm({ ...serviceForm, productCostTzs: e.target.value })} />
                  </div>
                  <div style={styles.fieldBlock}>
                    <label style={styles.fieldLabel}>Confirmation Rate %</label>
                    <input style={styles.input} type="number" min="0" max="100" value={serviceForm.confirmationRate} onChange={(e) => setServiceForm({ ...serviceForm, confirmationRate: e.target.value })} />
                  </div>
                  <div style={styles.fieldBlock}>
                    <label style={styles.fieldLabel}>Delivery Rate %</label>
                    <input style={styles.input} type="number" min="0" max="100" value={serviceForm.deliveryRate} onChange={(e) => setServiceForm({ ...serviceForm, deliveryRate: e.target.value })} />
                  </div>
                  <div style={styles.fieldBlock}>
                    <label style={styles.fieldLabel}>Ads Spend / Lead USD</label>
                    <input style={styles.input} type="number" min="0" step="0.01" value={serviceForm.cplUsd} onChange={(e) => setServiceForm({ ...serviceForm, cplUsd: e.target.value })} />
                  </div>
                </div>

                {selectedServiceDataset && (
                  <div style={{ marginTop: 16, padding: "12px 16px", borderRadius: 14, background: "rgba(247,243,237,0.7)", border: `1px solid ${cardBorder}`, display: "flex", alignItems: "center", gap: 12 }}>
                    <div style={{ fontSize: 12, fontWeight: 800, letterSpacing: 0.4, textTransform: "uppercase", color: textSoft }}>Auto service fee</div>
                    <div style={{ fontWeight: 900, color: textMain }}>{formatUSD(selectedServiceDataset.serviceFeePerOrderUsd)} per delivered order</div>
                    <div style={{ color: textSoft, fontSize: 12 }}>Auto-computed from service and country selection</div>
                  </div>
                )}
              </div>

              {selectedServiceDataset ? (
                <div style={{ display: "grid", gap: 16 }}>
                  <div style={{ ...styles.card, padding: 22 }}>
                    <div style={styles.sectionEyebrow}>Simulation results</div>
                    <div style={{ fontSize: 22, fontWeight: 900, marginTop: 8, marginBottom: 16 }}>
                      {simProductName ? simProductName : "Product"} — estimated profitability
                    </div>
                    <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(3, minmax(0, 1fr))", "1fr 1fr", "1fr"), gap: 16 }}>
                      <KpiCard
                        title="Estimated Profit / Order"
                        value={formatUSD(selectedServiceDataset.profitPerOrderUsd)}
                        sub="Net profit per delivered order"
                        valueColor={selectedServiceDataset.profitPerOrderUsd >= 0 ? green : red}
                      />
                      <KpiCard
                        title="Estimated Profit / 100 Leads"
                        value={formatUSD(selectedServiceDataset.profitFor100LeadsUsd)}
                        sub={`Based on ${serviceForm.confirmationRate}% confirm rate, ${serviceForm.deliveryRate}% delivery rate`}
                        valueColor={selectedServiceDataset.profitFor100LeadsUsd >= 0 ? green : red}
                      />
                      <KpiCard
                        title="Break-even CPA"
                        value={formatUSD(selectedServiceDataset.breakEvenCplUsd)}
                        sub="Max ad spend per lead before profit turns negative"
                        valueColor={selectedServiceDataset.breakEvenCplUsd >= selectedServiceDataset.costPerLeadUsd ? green : red}
                      />
                    </div>
                  </div>
                </div>
              ) : (
                <div style={{ ...styles.card, padding: 18, background: "#fff7ed", border: "1px solid #fed7aa" }}>
                  <div style={{ fontWeight: 800, color: amber, marginBottom: 6 }}>No service rules configured</div>
                  <div style={{ color: textSoft }}>No rules found for this service and country combination.</div>
                </div>
              )}
            </div>
          )}

{activePage === "situations" && (
            <div style={{ display: "grid", gap: 20 }}>
              <PageHeader
                eyebrow="Break-even Analysis"
                title="Cost structure and product economics"
                description="Analyse fixed charges, ads efficiency and break-even thresholds per product. Manage salaries and fixed charges in Profit Center > Global Expenses."
              />
              <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(4, minmax(0, 1fr))", "repeat(2, minmax(0, 1fr))", "1fr"), gap: 16 }}>
                <KpiCard icon={<Wallet size={18} />} title="Detected Charges" value={formatUsdFromTzs(situationsSummary.detectedChargesTzs)} sub="Products, import, ads, salaries and fixed charges" valueColor={red} />
                <KpiCard icon={<Users size={18} />} title="Salaries" value={formatUsdFromTzs(situationsSummary.salariesTotalTzs)} sub="Employee payroll included in fixed charges" />
                <KpiCard icon={<Archive size={18} />} title="Fixed Charges" value={formatUsdFromTzs(situationsSummary.fixedChargesTzs)} sub="Salaries + manual fixed charges" valueColor={amber} />
                <KpiCard
                  icon={<TrendingUp size={18} />}
                  title="Ads Used"
                  value={formatUsdFromTzs(situationsSummary.adSpendTzs)}
                  sub={
                    situationsSummary.metaTrackedAdsTzs > 0
                      ? "Meta cumulative total automatically included in charges"
                      : situationsSummary.configuredAverageLeadCostTzs > 0
                        ? `Average ad cost ${formatUsdFromTzs(situationsSummary.configuredAverageLeadCostTzs)} per lead`
                        : "Configure average ad cost and incoming leads per product"
                  }
                  valueColor={accent}
                />
              </div>

              <div style={{ ...styles.card, padding: 22 }}>
                <div style={styles.sectionHeader}>
                  <div>
                    <div style={styles.sectionEyebrow}>Cost center</div>
                    <div style={{ fontSize: 24, fontWeight: 900, marginTop: 8 }}>Situations</div>
                    <div style={{ color: textSoft, marginTop: 6, lineHeight: 1.6 }}>
                      Cette page centralise toutes les charges detectees, les salaires, les charges fixes et le calcul manuel des ads par produit.
                    </div>
                  </div>
                </div>

                <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(3, minmax(0, 1fr))", "1fr 1fr", "1fr"), gap: 14 }}>
                  <MiniStat label="Product Purchase" value={formatUsdFromTzs(situationsSummary.purchaseBudgetTzs)} tone="amber" sub="Detected from product buy price x imported stock" />
                  <MiniStat label="Import Charges" value={formatUsdFromTzs(situationsSummary.importChargesTzs)} tone="blue" sub="Shipping total + other charges" />
                  <MiniStat
                    label="Ad Spend"
                    value={formatUsdFromTzs(situationsSummary.adSpendTzs)}
                    tone="green"
                    sub={situationsSummary.metaTrackedAdsTzs > 0 ? "Meta daily cumulative total included automatically" : "Configured from ad cost x incoming leads"}
                  />
                  <MiniStat label="Local Delivery" value={formatUsdFromTzs(situationsSummary.localDeliveryTzs)} tone="blue" sub="Detected delivered orders cost" />
                  <MiniStat label="Manual Fixed" value={formatUsdFromTzs(situationsSummary.manualFixedChargesTzs)} tone="amber" sub="Rent, tools, subscriptions, utilities..." />
                  <MiniStat label="Payroll" value={formatUsdFromTzs(situationsSummary.salariesTotalTzs)} tone="green" sub="Team salaries included in fixed charges" />
                </div>
              </div>

              <div style={{ ...styles.card, padding: 22 }}>
                <div style={styles.sectionHeader}>
                  <div>
                    <div style={styles.sectionEyebrow}>Weekly profit</div>
                    <div style={{ fontSize: 22, fontWeight: 900, marginTop: 8 }}>Net profit per product / week</div>
                    <div style={{ color: textSoft, marginTop: 6, lineHeight: 1.6 }}>
                      This view helps you decide what to push, pause or fix every week based on delivered revenue, product cost, local delivery and allocated ads.
                    </div>
                  </div>
                </div>

                <div style={{ overflowX: "auto", border: `1px solid ${cardBorder}`, borderRadius: 20 }}>
                  <table style={{ width: "100%", borderCollapse: "separate", borderSpacing: 0 }}>
                    <thead>
                      <tr>
                        {["Week", "Product", "Orders", "Delivered", "Revenue", "Ads", "Import Cost", "Delivery Cost", "Net Profit", "Profit / Order"].map((head) => (
                          <th key={head} style={{ textAlign: "left", padding: "14px 12px", color: textSoft, fontSize: 12, fontWeight: 800, letterSpacing: 0.4, textTransform: "uppercase", borderBottom: `1px solid ${cardBorder}`, background: "rgba(247, 243, 237, 0.92)" }}>
                            {head}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {weeklyProductProfitRows.slice(0, 16).map((row, index) => (
                        <tr key={row.key} style={{ background: index % 2 === 0 ? "rgba(255,255,255,0.72)" : "rgba(250,247,242,0.8)" }}>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 700 }}>{row.weekLabel}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 800 }}>{row.productName}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>{row.orders}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>{row.deliveredOrders}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>{formatUsdFromTzs(row.revenueTzs)}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>{formatUsdFromTzs(row.adSpendTzs)}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>{formatUsdFromTzs(row.importCostTzs)}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>{formatUsdFromTzs(row.localDeliveryTzs)}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 800, color: row.profitTzs >= 0 ? green : red }}>{formatUsdFromTzs(row.profitTzs)}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>{formatUsdFromTzs(row.profitPerDeliveredOrderTzs)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  {weeklyProductProfitRows.length === 0 ? <div style={{ padding: 24, color: textSoft }}>No weekly product profit data yet.</div> : null}
                </div>
              </div>

              <div style={{ ...styles.card, padding: 22 }}>
                <div style={styles.sectionHeader}>
                  <div>
                    <div style={styles.sectionEyebrow}>Break-even details</div>
                    <div style={{ fontSize: 22, fontWeight: 900, marginTop: 8 }}>Product profitability threshold</div>
                    <div style={{ color: textSoft, marginTop: 6, lineHeight: 1.6 }}>
                      Formula used here: `CA = prix de vente x quantite sourcee au debut`. Then `charges fixes = cout total du stock + (8.5 USD x nombre de pieces sourcees)`. For ads, you enter `average ad cost` and `incoming leads`, and the app calculates `Ads Used = average ad cost x incoming leads`. The `PM` metric is calculated with `PM = (SR valeur x 12) / CA`.
                    </div>
                  </div>
                </div>

                <div style={{ overflowX: "auto", border: `1px solid ${cardBorder}`, borderRadius: 20 }}>
                  <table style={{ width: "100%", borderCollapse: "separate", borderSpacing: 0 }}>
                    <thead>
                      <tr>
                        {["Product", "Sourced Qty", "Incoming Leads", "CA", "Ads Used", "MCV", "Max Ads Cost", "Fixed Charges", "Result", "SR Value", "SR Volume", "PM", "Action"].map((head) => (
                          <th key={head} style={{ textAlign: "left", padding: "14px 12px", color: textSoft, fontSize: 12, fontWeight: 800, letterSpacing: 0.4, textTransform: "uppercase", borderBottom: `1px solid ${cardBorder}`, background: "rgba(247, 243, 237, 0.92)" }}>
                            {head}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {situationsSummary.productEconomics.map((product, index) => (
                        <tr key={product.id} style={{ background: index % 2 === 0 ? "rgba(255,255,255,0.72)" : "rgba(250,247,242,0.8)" }}>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 800 }}>{product.name}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 800 }}>{product.sourcedQty}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 800 }}>{product.leadVolume}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>
                            <div>{formatUsdFromTzs(product.revenueTzs)}</div>
                            <div style={{ color: textSoft, fontSize: 12 }}>{Number(product.revenuePercent || 0).toFixed(0)}%</div>
                          </td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>
                            <div>{formatUsdFromTzs(product.currentAdsCostTzs)}</div>
                            <div style={{ display: "grid", gap: 8, marginTop: 10 }}>
                              <div>
                                <div style={{ color: textSoft, fontSize: 11, fontWeight: 700, marginBottom: 4 }}>Average ads cost USD</div>
                                <input
                                  style={{ ...styles.input, minWidth: 140 }}
                                  type="number"
                                  min="0"
                                  step="0.01"
                                  value={getSituationAdInputDisplayValue(
                                    product.id,
                                    "averageLeadCostTzs",
                                    Number(product.averageLeadCostTzs || 0) > 0 ? String(Math.round((Number(product.averageLeadCostTzs || 0) / USD_TO_TZS) * 100) / 100) : ""
                                  )}
                                  onChange={(e) => updateSituationAdInput(product.id, "averageLeadCostTzs", e.target.value)}
                                />
                              </div>
                              <div>
                                <div style={{ color: textSoft, fontSize: 11, fontWeight: 700, marginBottom: 4 }}>Incoming leads</div>
                                <input
                                  style={{ ...styles.input, minWidth: 140 }}
                                  type="number"
                                  min="0"
                                  step="1"
                                  value={getSituationAdInputDisplayValue(
                                    product.id,
                                    "incomingLeads",
                                    Number(product.leadVolume || 0) > 0 ? String(product.leadVolume) : ""
                                  )}
                                  onChange={(e) => updateSituationAdInput(product.id, "incomingLeads", e.target.value)}
                                />
                              </div>
                            </div>
                            <div style={{ color: textSoft, fontSize: 12, marginTop: 8 }}>{product.adsInputSourceLabel}</div>
                            <div style={{ marginTop: 10, height: 8, borderRadius: 999, background: "rgba(23,32,51,0.08)", overflow: "hidden" }}>
                              <div
                                style={{
                                  width: `${Math.min(100, Math.max(0, product.adsUsageRatio * 100))}%`,
                                  height: "100%",
                                  borderRadius: 999,
                                  background: product.adsUsageRatio > 1 ? "linear-gradient(90deg, #d9485f, #f97316)" : "linear-gradient(90deg, #1d5fd0, #1f8f5f)",
                                }}
                              />
                            </div>
                          </td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>
                            <div>{formatUsdFromTzs(product.marginOnVariableCostTzs)}</div>
                            <div style={{ color: textSoft, fontSize: 12 }}>{Number(product.tmcvPercent || 0).toFixed(2)}% of CA</div>
                          </td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>
                            <div>{formatUsdFromTzs(product.adsCostTzs)}</div>
                            <div style={{ color: textSoft, fontSize: 12 }}>Maximum ads budget supportable</div>
                            <div style={{ color: textSoft, fontSize: 12 }}>{Number(product.adsCostPercent || 0).toFixed(2)}% of CA</div>
                          </td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>
                            <div>{formatUsdFromTzs(product.allocatedFixedChargesTzs)}</div>
                            <div style={{ color: textSoft, fontSize: 12 }}>{Number(product.fixedChargesPercent || 0).toFixed(2)}% of CA</div>
                          </td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>
                            <div style={{ color: Number(product.resultTzs || 0) >= 0 ? green : red }}>{formatUsdFromTzs(product.resultTzs)}</div>
                            <div style={{ color: textSoft, fontSize: 12 }}>{Number(product.resultPercent || 0).toFixed(2)}%</div>
                          </td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 800 }}>
                            {product.srValueTzs && Number.isFinite(product.srValueTzs) ? formatUsdFromTzs(product.srValueTzs) : "N/A"}
                          </td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 800 }}>
                            {product.srVolume && Number.isFinite(product.srVolume) ? `${product.srVolume.toFixed(2)} pcs` : "N/A"}
                          </td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 800 }}>
                            {product.breakEvenTimeMonths && Number.isFinite(product.breakEvenTimeMonths)
                              ? `${product.breakEvenTimeMonths.toFixed(1)} mois`
                              : "N/A"}
                          </td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>
                            <div style={getDecisionStyle(product.currentAdsCostTzs > product.adsCostTzs ? "BAD PRODUCT" : product.srValueTzs && product.resultTzs > 0 ? "GOOD PRODUCT" : "WATCH")}>
                              {product.currentAdsCostTzs > product.adsCostTzs ? "Ads too high" : product.srValueTzs && product.resultTzs > 0 ? "Healthy" : "Watch"}
                            </div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

{activePage === "performanceHub" && (
            <div style={{ display: "grid", gap: 20 }}>
              <div style={{ ...styles.card, padding: 22, display: ordersTab === "import" ? "block" : "none" }}>
                <div style={styles.sectionHeader}>
                  <div>
                    <div style={styles.sectionEyebrow}>Main control panel</div>
                    <div style={{ fontSize: 24, fontWeight: 900, marginTop: 8 }}>Scale, stop and protect cash</div>
                    <div style={{ color: textSoft, marginTop: 6, lineHeight: 1.6 }}>
                      One board to see revenue, ads, product costs, delivery costs, funnel health and the products that need immediate action.
                    </div>
                  </div>
                </div>

                <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(6, minmax(0, 1fr))", "repeat(2, minmax(0, 1fr))", "1fr"), gap: 16 }}>
                  <KpiCard icon={<Wallet size={18} />} title="Total Revenue" value={formatTZS(controlPanelSummary.totalRevenueTzs)} sub="Delivered orders only" valueColor={green} />
                  <KpiCard icon={<ClipboardList size={18} />} title="Total Ads Spend" value={formatTZS(controlPanelSummary.totalAdsSpendTzs)} sub="Manual + synced product spend" valueColor={amber} />
                  <KpiCard icon={<Archive size={18} />} title="Total Product Cost" value={formatTZS(controlPanelSummary.totalProductCostTzs)} sub="Units sold x cost per unit" valueColor={amber} />
                  <KpiCard icon={<ShoppingBag size={18} />} title="Total Delivery Cost" value={formatTZS(controlPanelSummary.totalDeliveryCostTzs)} sub="Units sold x delivery cost" valueColor={amber} />
                  <KpiCard icon={<TrendingUp size={18} />} title="Total Profit" value={formatTZS(controlPanelSummary.totalProfitTzs)} sub="Revenue - ads - product - delivery" valueColor={Number(controlPanelSummary.totalProfitTzs || 0) >= 0 ? green : red} />
                  <KpiCard icon={<Calculator size={18} />} title="Average Profit Margin" value={`${controlPanelSummary.averageProfitMarginPct.toFixed(1)}%`} sub="Average across revenue-generating products" valueColor={Number(controlPanelSummary.averageProfitMarginPct || 0) >= 0 ? accent : red} />
                  <KpiCard icon={<Users size={18} />} title="Total Leads" value={controlPanelSummary.totalLeads} sub="Countable customer leads" />
                  <KpiCard icon={<Phone size={18} />} title="Total Confirmed Orders" value={controlPanelSummary.totalConfirmedOrders} sub="Confirmed pipeline base" valueColor={amber} />
                  <KpiCard icon={<Rocket size={18} />} title="Total Delivered Orders" value={controlPanelSummary.totalDeliveredOrders} sub="Successfully completed deliveries" valueColor={green} />
                  <KpiCard icon={<BarChart3 size={18} />} title="Global Confirmation Rate" value={`${Math.round(controlPanelSummary.globalConfirmationRate)}%`} sub="Confirmed / total leads" />
                  <KpiCard icon={<BarChart3 size={18} />} title="Global Delivery Rate" value={`${Math.round(controlPanelSummary.globalDeliveryRate)}%`} sub="Delivered / confirmed orders" valueColor={green} />
                  <KpiCard icon={<AlertTriangle size={18} />} title="Products Needing Attention" value={controlPanelSummary.needsAttentionProducts.length} sub="Active alerts on product performance" valueColor={controlPanelSummary.needsAttentionProducts.length > 0 ? red : green} />
                </div>
              </div>

              <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(4, minmax(0, 1fr))", "repeat(2, minmax(0, 1fr))", "1fr"), gap: 16 }}>
                <KpiCard icon={<ClipboardList size={18} />} title="Total Leads" value={customersDashboard.totalOrders} sub={`${Math.round(customersDashboard.confirmationRate)}% confirmation rate`} />
                <KpiCard icon={<ShoppingBag size={18} />} title="Delivered Orders" value={customersDashboard.deliveredOrders} sub={`${Math.round(customersDashboard.deliveryRate)}% delivery rate`} valueColor={green} />
                <KpiCard icon={<Wallet size={18} />} title="Revenue" value={formatUsdFromTzs(liveAutomationSummary.totalRevenueTzs)} sub="Delivered orders revenue" valueColor={green} />
                <KpiCard icon={<TrendingUp size={18} />} title="Gross Profit" value={formatUsdFromTzs(executiveSummary.grossProfitTzs)} sub="Before fixed charges" valueColor={Number(executiveSummary.grossProfitTzs || 0) >= 0 ? green : red} />
                <KpiCard icon={<Calculator size={18} />} title="Net After Fixed" value={formatUsdFromTzs(executiveSummary.estimatedNetAfterFixedTzs)} sub="Estimated profit after fixed charges" valueColor={Number(executiveSummary.estimatedNetAfterFixedTzs || 0) >= 0 ? green : red} />
                <KpiCard icon={<Archive size={18} />} title="Stock Value" value={formatUsdFromTzs(executiveSummary.stockImmobilizedTzs)} sub="Capital locked in available stock" valueColor={amber} />
                <KpiCard icon={<AlertTriangle size={18} />} title="Open Tasks" value={executiveSummary.openTasks} sub={`${executiveSummary.highPriorityTasks} high priority`} valueColor={executiveSummary.highPriorityTasks > 0 ? red : accent} />
                <KpiCard icon={<Rocket size={18} />} title="Top Product" value={profitCenterSummary.topProduct?.name || "N/A"} sub={profitCenterSummary.topProduct ? `${formatUsdFromTzs(profitCenterSummary.topProduct.balanceTzs || 0)} balance` : "No product data yet"} />
              </div>

              <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("1.2fr 1fr", "1fr", "1fr"), gap: 20 }}>
                <div style={{ ...styles.card, padding: 22 }}>
                  <div style={styles.sectionHeader}>
                    <div>
                      <div style={styles.sectionEyebrow}>Executive pulse</div>
                      <div style={{ fontSize: 24, fontWeight: 900, marginTop: 8 }}>What needs attention now</div>
                      <div style={{ color: textSoft, marginTop: 6, lineHeight: 1.6 }}>
                        A single page to decide what to scale, what to fix, and what is slowing down the business today.
                      </div>
                    </div>
                  </div>

                  <div style={{ display: "grid", gap: 12 }}>
                    {taskCenterData.slice(0, 5).map((task) => (
                      <div key={task.id} style={{ ...styles.softStat, display: "flex", alignItems: "center", justifyContent: "space-between", gap: 14, flexWrap: "wrap" }}>
                        <div style={{ minWidth: 260, flex: "1 1 320px" }}>
                          <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
                            <span style={getDecisionStyle(task.priority === "High" ? "KILL" : task.priority === "Medium" ? "WATCH" : "OK")}>{task.priority}</span>
                            <span style={{ ...styles.badge, background: "rgba(35,88,213,0.08)", color: accent, border: "1px solid rgba(35,88,213,0.12)" }}>{formatStatusLabel(task.type)}</span>
                          </div>
                          <div style={{ fontWeight: 800, marginTop: 8 }}>{task.title}</div>
                          <div style={{ color: textSoft, marginTop: 6 }}>{task.detail}</div>
                        </div>
                        <button style={styles.btnPrimary} onClick={() => setActivePage(task.page)}>Open</button>
                      </div>
                    ))}
                    {taskCenterData.length === 0 ? <div style={{ color: textSoft }}>No urgent blocker detected right now.</div> : null}
                  </div>
                </div>

                <div style={{ display: "grid", gap: 20 }}>
                  <div style={{ ...styles.card, padding: 22 }}>
                    <div style={styles.sectionEyebrow}>Growth</div>
                    <div style={{ fontSize: 22, fontWeight: 900, marginTop: 8 }}>Products ready to scale</div>
                    <div style={{ display: "grid", gap: 10, marginTop: 16 }}>
                      {scalingSummary.ready.slice(0, 4).map((product) => (
                        <div key={product.id} style={{ ...styles.softStat, display: "grid", gap: 6 }}>
                          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10 }}>
                            <div style={{ fontWeight: 800 }}>{product.name}</div>
                            <div style={getDecisionStyle("SCALE")}>Scale</div>
                          </div>
                          <div style={{ color: textSoft, fontSize: 13 }}>
                            ROAS {Number(product.roas || 0).toFixed(2)} | {Math.round((product.deliveryRate || 0) * 100)}% delivery | {product.availableStock} units available
                          </div>
                        </div>
                      ))}
                      {scalingSummary.ready.length === 0 ? <div style={{ color: textSoft }}>No product is fully ready to scale yet.</div> : null}
                    </div>
                  </div>

                  <div style={{ ...styles.card, padding: 22 }}>
                    <div style={styles.sectionEyebrow}>Risk watch</div>
                    <div style={{ fontSize: 22, fontWeight: 900, marginTop: 8 }}>Stock and delivery pressure</div>
                    <div style={{ display: "grid", gap: 10, marginTop: 16 }}>
                      {stockForecastRows.filter((product) => product.urgency !== "Healthy").slice(0, 4).map((product) => (
                        <div key={product.id} style={{ ...styles.softStat, display: "grid", gap: 6 }}>
                          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10 }}>
                            <div style={{ fontWeight: 800 }}>{product.name}</div>
                            <div style={getDecisionStyle(product.urgency === "Critical" ? "KILL" : "WATCH")}>{product.urgency}</div>
                          </div>
                          <div style={{ color: textSoft, fontSize: 13 }}>
                            {product.daysUntilStockout != null ? `${product.daysUntilStockout} days left` : "No stockout projection yet"} | Projected {product.projectedStockoutDate || "N/A"}
                          </div>
                        </div>
                      ))}
                      {stockForecastRows.filter((product) => product.urgency !== "Healthy").length === 0 ? <div style={{ color: textSoft }}>No stock risk detected right now.</div> : null}
                    </div>
                  </div>

                  <div style={{ ...styles.card, padding: 22 }}>
                    <div style={styles.sectionEyebrow}>Cashflow</div>
                    <div style={{ fontSize: 22, fontWeight: 900, marginTop: 8 }}>Business cash snapshot</div>
                    <div style={{ display: "grid", gap: 12, marginTop: 16 }}>
                      <MiniStat label="Cash in" value={formatUsdFromTzs(cashflowSummary.cashInTzs)} tone="green" sub="Delivered revenue" />
                      <MiniStat label="Variable out" value={formatUsdFromTzs(cashflowSummary.variableOutTzs)} tone="amber" sub="Ads + delivered cost" />
                      <MiniStat label="Fixed out" value={formatUsdFromTzs(cashflowSummary.fixedOutTzs)} tone="blue" sub="Payroll + fixed charges" />
                      <MiniStat label="Net cash" value={formatUsdFromTzs(cashflowSummary.netCashTzs)} tone={cashflowSummary.netCashTzs >= 0 ? "green" : "amber"} sub="Estimated after variable + fixed costs" />
                    </div>
                  </div>
                </div>
              </div>

              <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("1fr 1fr", "1fr", "1fr"), gap: 20 }}>
                <div style={{ ...styles.card, padding: 22 }}>
                  <div style={styles.sectionEyebrow}>Top winners</div>
                  <div style={{ fontSize: 22, fontWeight: 900, marginTop: 8 }}>Products to scale</div>
                  <div style={{ display: "grid", gap: 10, marginTop: 16 }}>
                    {controlPanelSummary.topWinningProducts.map((product) => (
                      <div key={`winner-${product.id}`} style={{ ...styles.softStat, display: "grid", gap: 6 }}>
                        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10 }}>
                          <div style={{ fontWeight: 800 }}>{product.name}</div>
                          <div style={getDecisionStyle("SCALE")}>WINNER</div>
                        </div>
                        <div style={{ color: textSoft, fontSize: 13 }}>
                          Profit {formatTZS(product.dashboardProfitTzs)} | Margin {product.dashboardProfitMargin.toFixed(1)}% | Delivered {product.effectiveDeliveredOrders}
                        </div>
                      </div>
                    ))}
                    {controlPanelSummary.topWinningProducts.length === 0 ? <div style={{ color: textSoft }}>No winning product yet.</div> : null}
                  </div>
                </div>

                <div style={{ ...styles.card, padding: 22 }}>
                  <div style={styles.sectionEyebrow}>Loss watch</div>
                  <div style={{ fontSize: 22, fontWeight: 900, marginTop: 8 }}>Products to stop or fix</div>
                  <div style={{ display: "grid", gap: 10, marginTop: 16 }}>
                    {controlPanelSummary.losingProducts.map((product) => (
                      <div key={`loser-${product.id}`} style={{ ...styles.softStat, display: "grid", gap: 6 }}>
                        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10 }}>
                          <div style={{ fontWeight: 800 }}>{product.name}</div>
                          <div style={getDecisionStyle("KILL")}>LOSS</div>
                        </div>
                        <div style={{ color: textSoft, fontSize: 13 }}>
                          Profit {formatTZS(product.dashboardProfitTzs)} | Ads {formatTZS(product.dashboardAdsSpendTzs)} | Revenue {formatTZS(product.totalRevenue)}
                        </div>
                      </div>
                    ))}
                    {controlPanelSummary.losingProducts.length === 0 ? <div style={{ color: textSoft }}>No losing product right now.</div> : null}
                  </div>
                </div>

                <div style={{ ...styles.card, padding: 22 }}>
                  <div style={styles.sectionEyebrow}>Low stock</div>
                  <div style={{ fontSize: 22, fontWeight: 900, marginTop: 8 }}>Products to restock</div>
                  <div style={{ display: "grid", gap: 10, marginTop: 16 }}>
                    {controlPanelSummary.lowStockProducts.map((product) => (
                      <div key={`stock-${product.id}`} style={{ ...styles.softStat, display: "grid", gap: 6 }}>
                        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10 }}>
                          <div style={{ fontWeight: 800 }}>{product.name}</div>
                          <div style={getDecisionStyle("Low Stock")}>LOW STOCK</div>
                        </div>
                        <div style={{ color: textSoft, fontSize: 13 }}>
                          Stock {product.stockQuantity} | Minimum {situationData.productAlertThresholds?.minStockQuantity || 0}
                        </div>
                      </div>
                    ))}
                    {controlPanelSummary.lowStockProducts.length === 0 ? <div style={{ color: textSoft }}>No low stock product at the moment.</div> : null}
                  </div>
                </div>

                <div style={{ ...styles.card, padding: 22 }}>
                  <div style={styles.sectionEyebrow}>Attention list</div>
                  <div style={{ fontSize: 22, fontWeight: 900, marginTop: 8 }}>Products that need attention</div>
                  <div style={{ display: "grid", gap: 10, marginTop: 16 }}>
                    {controlPanelSummary.needsAttentionProducts.map((product) => (
                      <div key={`attention-${product.id}`} style={{ ...styles.softStat, display: "grid", gap: 8 }}>
                        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10 }}>
                          <div style={{ fontWeight: 800 }}>{product.name}</div>
                          <div style={getDecisionStyle(product.performanceStatus === "WINNER" ? "OK" : product.performanceStatus === "TESTING" ? "WATCH" : "KILL")}>
                            {product.performanceStatus}
                          </div>
                        </div>
                        <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                          {product.productAlerts.map((alert) => (
                            <span key={`cp-${product.id}-${alert.key}`} style={getAlertBadgeStyle(alert.tone)}>
                              {alert.message}
                            </span>
                          ))}
                        </div>
                      </div>
                    ))}
                    {controlPanelSummary.needsAttentionProducts.length === 0 ? <div style={{ color: textSoft }}>No product currently needs extra attention.</div> : null}
                  </div>
                </div>
              </div>
            </div>
          )}

{["profitCenter", "financeHub"].includes(activePage) && (
            <div style={{ display: "grid", gap: 20 }}>
              <PageHeader
                eyebrow="Profit Center"
                title="Financial results"
                description="Revenue from orders, stock charges from purchases, manual ads and extra charges — all combined into one profit view."
              />
              <InlineTabs
                items={[
                  { value: "overview", label: "Overview" },
                  { value: "product-profit", label: "Product Profit" },
                  { value: "ads-spend", label: "Ads Spend" },
                  { value: "extra-charges", label: "Extra Charges" },
                  { value: "cash-balance", label: "Cash Balance" },
                  { value: "simulator", label: "Simulator" },
                ]}
                value={profitTab}
                onChange={setProfitTab}
              />

              {/* ── OVERVIEW TAB ── */}
              {profitTab === "overview" && (() => {
                const m = profitOverviewMetrics;
                const rows = [
                  { label: "Revenue", value: m.revenueUsd, tsh: m.revenueTsh, color: green, sign: "+" },
                  { label: "Stock Charges", value: m.stockChargesUsd, tsh: m.stockChargesTzs, color: amber, sign: "-" },
                  { label: "Service Fees", value: m.serviceChargesUsd, tsh: m.serviceChargesTzs, color: amber, sign: "-" },
                  { label: "Ads Spend (Meta)", value: m.metaAdsUsd, tsh: m.metaAdsTzs, color: amber, sign: "-" },
                  { label: "Ads Spend (Manual)", value: m.manualAdsUsd, tsh: m.manualAdsTzs, color: amber, sign: "-" },
                  { label: "Extra Charges", value: m.extraChargesUsd, tsh: m.extraChargesTzs, color: amber, sign: "-" },
                ];
                return (
                  <div style={{ display: "grid", gap: 16 }}>
                    {/* KPI cards */}
                    <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(3, minmax(0, 1fr))", "repeat(2, minmax(0, 1fr))", "1fr"), gap: 14 }}>
                      <KpiCard icon={<Wallet size={18} />} title="Revenue" value={formatUSD(m.revenueUsd)} sub={formatTZS(m.revenueTsh) + " — delivered orders"} valueColor={green} />
                      <KpiCard icon={<Archive size={18} />} title="Stock Charges" value={formatUSD(m.stockChargesUsd)} sub={formatTZS(m.stockChargesTzs) + " — received purchases"} valueColor={amber} />
                      <KpiCard icon={<TrendingUp size={18} />} title="Service Fees" value={formatUSD(m.serviceChargesUsd)} sub={formatTZS(m.serviceChargesTzs) + " — Dar $8 · Other $9"} valueColor={amber} />
                      <KpiCard icon={<ClipboardList size={18} />} title="Total Ads" value={formatUSD(m.totalAdsUsd)} sub={`Meta ${formatUSD(m.metaAdsUsd)} + Manual ${formatUSD(m.manualAdsUsd)}`} valueColor={amber} />
                      <KpiCard icon={<Boxes size={18} />} title="Extra Charges" value={formatUSD(m.extraChargesUsd)} sub={formatTZS(m.extraChargesTzs) + " — transport / tools / other"} valueColor={amber} />
                      <KpiCard icon={<TrendingUp size={18} />} title="Delivered Units" value={m.deliveredUnits} sub="From delivered orders" />
                    </div>

                    {/* Profit formula */}
                    <div style={{ ...styles.card, padding: 22 }}>
                      <div style={styles.sectionEyebrow}>Profit Formula</div>
                      <div style={{ fontSize: 20, fontWeight: 900, marginTop: 6, marginBottom: 16 }}>Step-by-step calculation</div>
                      <div style={{ display: "grid", gap: 0 }}>
                        {rows.map((r, i) => (
                          <div key={r.label} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "12px 0", borderBottom: i < rows.length - 1 ? `1px solid ${cardBorder}` : "none" }}>
                            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                              <span style={{ fontSize: 16, fontWeight: 900, color: r.sign === "+" ? green : amber, width: 16 }}>{r.sign}</span>
                              <span style={{ fontWeight: 700 }}>{r.label}</span>
                            </div>
                            <div style={{ textAlign: "right" }}>
                              <div style={{ fontWeight: 900, color: r.color }}>{formatUSD(r.value)}</div>
                              <div style={{ fontSize: 12, color: textSoft }}>{formatTZS(r.tsh)}</div>
                            </div>
                          </div>
                        ))}
                        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "16px 0", marginTop: 4, borderTop: `2px solid ${cardBorder}` }}>
                          <div style={{ fontSize: 18, fontWeight: 900 }}>= Business Profit</div>
                          <div style={{ textAlign: "right" }}>
                            <div style={{ fontSize: 22, fontWeight: 900, color: m.businessProfitUsd >= 0 ? green : red }}>{formatUSD(m.businessProfitUsd)}</div>
                            <div style={{ fontSize: 13, color: textSoft }}>{formatTZS(m.businessProfitTzs)} · {m.profitMarginPct.toFixed(1)}% margin</div>
                          </div>
                        </div>
                      </div>
                    </div>

                    {/* Revenue import */}
                    <div style={{ ...styles.card, padding: 22 }}>
                      <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 16, flexWrap: "wrap" }}>
                        <div>
                          <div style={styles.sectionEyebrow}>Revenue Source</div>
                          <div style={{ fontSize: 18, fontWeight: 900, marginTop: 6 }}>Import Excel for Revenue</div>
                          <div style={{ color: textSoft, fontSize: 13, marginTop: 4, lineHeight: 1.5 }}>Upload shipping / leads Excel. Uses CODE column as unique key — re-importing the same file never duplicates orders.</div>
                          {m.revenueImportedAt && (
                            <div style={{ color: textSoft, fontSize: 12, marginTop: 6 }}>
                              Last import: {new Date(m.revenueImportedAt).toLocaleDateString()} — {revenueImport?.lastImportNewCount ?? 0} new orders added, {revenueImport?.lastImportUpdatedCount ?? 0} orders updated, {revenueImport?.deliveredCount ?? 0} total delivered
                              {revenueImport?.totalStoredOrders ? ` · ${revenueImport.totalStoredOrders} orders tracked total` : ""}
                            </div>
                          )}
                        </div>
                        <label style={{ ...styles.btnPrimary, cursor: "pointer", whiteSpace: "nowrap" }}>
                          Import Excel
                          <input type="file" accept=".xlsx,.xls" onChange={importRevenueFromExcel} style={{ display: "none" }} />
                        </label>
                      </div>
                      {revenueImportNotice && <div style={{ marginTop: 12, padding: "10px 14px", borderRadius: 10, background: revenueImportNotice.startsWith("Import failed") ? "rgba(239,68,68,0.08)" : "rgba(34,197,94,0.08)", color: revenueImportNotice.startsWith("Import failed") ? red : green, fontSize: 13, fontWeight: 700 }}>{revenueImportNotice}</div>}
                    </div>
                  </div>
                );
              })()}

              {/* ── PRODUCT PROFIT TAB ── */}
              {profitTab === "product-profit" && (
                <div style={{ ...styles.card, padding: 22 }}>
                  <div style={styles.sectionEyebrow}>Per-product breakdown</div>
                  <div style={{ fontSize: 22, fontWeight: 900, marginTop: 6, marginBottom: 16 }}>Revenue · Stock · Service · Ads · Profit</div>
                  <div style={{ overflowX: "auto", border: `1px solid ${cardBorder}`, borderRadius: 16 }}>
                    <table style={{ width: "100%", borderCollapse: "separate", borderSpacing: 0 }}>
                      <thead>
                        <tr>
                          {["Product", "Delivered", "Revenue", "Stock Cost", "Service Fees", "Meta Ads", "Manual Ads", "Profit", "Margin", "Status"].map((h) => (
                            <th key={h} style={{ textAlign: "left", padding: "12px 10px", color: textSoft, fontSize: 11, fontWeight: 800, letterSpacing: 0.4, textTransform: "uppercase", borderBottom: `1px solid ${cardBorder}`, background: "rgba(247,243,237,0.92)", whiteSpace: "nowrap" }}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {productProfitRows.map((row, idx) => (
                          <tr key={row.id} style={{ background: idx % 2 === 0 ? "rgba(255,255,255,0.72)" : "rgba(250,247,242,0.8)" }}>
                            <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 800 }}>{row.name}</td>
                            <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}` }}>{row.deliveredUnits}</td>
                            <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}`, color: green, fontWeight: 700 }}>{formatUSD(row.revenueUsd)}<div style={{ fontSize: 11, color: textSoft }}>{formatTZS(row.revenue)}</div></td>
                            <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}`, color: amber }}>{formatUSD(row.stockCostUsd)}</td>
                            <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}`, color: amber }}>{formatUSD(row.serviceFeesUsd)}</td>
                            <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}`, color: amber }}>{formatUSD(row.metaAdsUsd)}</td>
                            <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}`, color: amber }}>{formatUSD(row.manualAdsUsd)}</td>
                            <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 800, color: row.profitUsd >= 0 ? green : red }}>{formatUSD(row.profitUsd)}<div style={{ fontSize: 11, color: textSoft }}>{formatTZS(row.profitTzs)}</div></td>
                            <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}` }}>{row.marginPct.toFixed(1)}%</td>
                            <td style={{ padding: "12px 10px", borderBottom: `1px solid ${cardBorder}` }}>
                              <span style={{ padding: "3px 10px", borderRadius: 20, fontSize: 11, fontWeight: 800, background: row.status === "positive" ? "rgba(34,197,94,0.12)" : row.status === "negative" ? "rgba(239,68,68,0.12)" : "rgba(100,116,139,0.1)", color: row.status === "positive" ? green : row.status === "negative" ? red : textSoft }}>
                                {row.status === "positive" ? "Positive" : row.status === "negative" ? "Negative" : "No Data"}
                              </span>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                    {productProfitRows.length === 0 && <div style={{ padding: 24, color: textSoft }}>No product data yet.</div>}
                  </div>
                </div>
              )}

              {/* ── ADS SPEND TAB ── */}
              {profitTab === "ads-spend" && (() => {
                const xr = Number(serviceForm?.exchangeRate || USD_TO_TZS);
                const addEntry = async () => {
                  if (!manualAdsForm.weekStart || !manualAdsForm.weekEnd || !manualAdsForm.amountUsd) {
                    setManualAdsNotice("Week start, week end and amount are required.");
                    setTimeout(() => setManualAdsNotice(""), 3000);
                    return;
                  }
                  const amtUsd = Number(manualAdsForm.amountUsd) || 0;
                  const product = products.find((p) => p.id === manualAdsForm.productId);
                  const entry = {
                    weekStart: manualAdsForm.weekStart, weekEnd: manualAdsForm.weekEnd,
                    productId: manualAdsForm.productId || "", productName: product?.name || manualAdsForm.productName || "",
                    amountUsd: amtUsd, amountTsh: amtUsd * xr,
                    notes: manualAdsForm.notes || "",
                  };
                  // Optimistic: add immediately so data is never lost
                  const localId = `local-${Date.now()}`;
                  const localEntry = { ...entry, id: localId, createdAt: new Date().toISOString() };
                  setManualAdsSpend((prev) => { const next = [localEntry, ...prev]; writeLS(LS_MANUAL_ADS, next); return next; });
                  setManualAdsForm({ weekStart: "", weekEnd: "", productId: "", productName: "", amountUsd: "", notes: "" });
                  setManualAdsNotice("Saving…");
                  // Persist to Supabase in background
                  try {
                    const saved = await saveManualAdsSpendToSupabase(entry);
                    if (saved?.id) {
                      setManualAdsSpend((prev) => { const next = prev.map((e) => e.id === localId ? { ...localEntry, id: saved.id, createdAt: saved.created_at || localEntry.createdAt } : e); writeLS(LS_MANUAL_ADS, next); return next; });
                    }
                    setManualAdsNotice("Saved.");
                  } catch (err) {
                    setManualAdsNotice(err?.code === "42P01" ? "Saved locally — run SQL migration 004 in Supabase to enable cloud sync." : "Saved locally (Supabase unavailable).");
                  }
                  setTimeout(() => setManualAdsNotice(""), 5000);
                };
                const removeEntry = async (id) => {
                  setManualAdsSpend((prev) => { const next = prev.filter((e) => e.id !== id); writeLS(LS_MANUAL_ADS, next); return next; });
                  if (!String(id).startsWith("local-")) await deleteManualAdsSpendFromSupabase(id).catch(() => {});
                };
                const totalManualUsd = manualAdsSpend.reduce((s, e) => s + Number(e.amountUsd || 0), 0);
                const metaTotal = profitOverviewMetrics.metaAdsUsd;
                return (
                  <div style={{ display: "grid", gap: 16 }}>
                    <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(3, minmax(0, 1fr))", "1fr 1fr", "1fr"), gap: 14 }}>
                      <KpiCard icon={<ClipboardList size={18} />} title="Meta Ads Total" value={formatUSD(metaTotal)} sub={formatTZS(metaTotal * xr) + " — from ads_campaigns"} valueColor={amber} />
                      <KpiCard icon={<Archive size={18} />} title="Manual Ads Total" value={formatUSD(totalManualUsd)} sub={formatTZS(totalManualUsd * xr) + " — weekly entries"} valueColor={amber} />
                      <KpiCard icon={<TrendingUp size={18} />} title="Combined Ads" value={formatUSD(metaTotal + totalManualUsd)} sub={formatTZS((metaTotal + totalManualUsd) * xr) + " — Meta + Manual"} valueColor={amber} />
                    </div>
                    <div style={{ ...styles.card, padding: 22 }}>
                      <div style={styles.sectionEyebrow}>Add Weekly Entry</div>
                      <div style={{ fontSize: 18, fontWeight: 900, marginTop: 6, marginBottom: 14 }}>Manual Ads Spend</div>
                      <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("1fr 1fr 1fr 1fr 1fr auto", "1fr 1fr 1fr", "1fr"), gap: 10, alignItems: "end" }}>
                        <div style={styles.fieldBlock}><label style={styles.fieldLabel}>Week Start</label><input style={styles.input} type="date" value={manualAdsForm.weekStart} onChange={(e) => setManualAdsForm((f) => ({ ...f, weekStart: e.target.value }))} /></div>
                        <div style={styles.fieldBlock}><label style={styles.fieldLabel}>Week End</label><input style={styles.input} type="date" value={manualAdsForm.weekEnd} onChange={(e) => setManualAdsForm((f) => ({ ...f, weekEnd: e.target.value }))} /></div>
                        <div style={styles.fieldBlock}>
                          <label style={styles.fieldLabel}>Product (optional)</label>
                          <select style={styles.input} value={manualAdsForm.productId} onChange={(e) => setManualAdsForm((f) => ({ ...f, productId: e.target.value }))}>
                            <option value="">Global / Unmapped</option>
                            {products.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
                          </select>
                        </div>
                        <div style={styles.fieldBlock}><label style={styles.fieldLabel}>Amount USD</label><input style={styles.input} type="number" min="0" step="0.01" placeholder="0.00" value={manualAdsForm.amountUsd} onChange={(e) => setManualAdsForm((f) => ({ ...f, amountUsd: e.target.value }))} /></div>
                        <div style={styles.fieldBlock}><label style={styles.fieldLabel}>Notes</label><input style={styles.input} placeholder="Optional" value={manualAdsForm.notes} onChange={(e) => setManualAdsForm((f) => ({ ...f, notes: e.target.value }))} /></div>
                        <button style={styles.btnPrimary} onClick={addEntry}>Add</button>
                      </div>
                      {manualAdsNotice && <div style={{ marginTop: 10, color: manualAdsNotice.startsWith("Failed") ? red : green, fontSize: 13, fontWeight: 700 }}>{manualAdsNotice}</div>}
                    </div>
                    <div style={{ ...styles.card, padding: 22 }}>
                      <div style={{ fontSize: 16, fontWeight: 900, marginBottom: 14 }}>All Manual Entries ({manualAdsSpend.length})</div>
                      {manualAdsSpend.length === 0 ? <div style={{ color: textSoft }}>No manual ads entries yet.</div> : (
                        <div style={{ overflowX: "auto", border: `1px solid ${cardBorder}`, borderRadius: 14 }}>
                          <table style={{ width: "100%", borderCollapse: "separate", borderSpacing: 0 }}>
                            <thead><tr>{["Week", "Product", "Amount USD", "Amount TSh", "Notes", ""].map((h) => <th key={h} style={{ textAlign: "left", padding: "10px 10px", color: textSoft, fontSize: 11, fontWeight: 800, textTransform: "uppercase", borderBottom: `1px solid ${cardBorder}`, background: "rgba(247,243,237,0.92)" }}>{h}</th>)}</tr></thead>
                            <tbody>
                              {manualAdsSpend.map((e, i) => (
                                <tr key={e.id} style={{ background: i % 2 === 0 ? "rgba(255,255,255,0.72)" : "rgba(250,247,242,0.8)" }}>
                                  <td style={{ padding: "10px 10px", borderBottom: `1px solid ${cardBorder}`, fontSize: 13 }}>{e.weekStart} → {e.weekEnd}</td>
                                  <td style={{ padding: "10px 10px", borderBottom: `1px solid ${cardBorder}`, fontSize: 13 }}>{e.productName || "Global"}</td>
                                  <td style={{ padding: "10px 10px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 700, color: amber }}>{formatUSD(e.amountUsd)}</td>
                                  <td style={{ padding: "10px 10px", borderBottom: `1px solid ${cardBorder}`, fontSize: 12, color: textSoft }}>{formatTZS(e.amountTsh || e.amountUsd * xr)}</td>
                                  <td style={{ padding: "10px 10px", borderBottom: `1px solid ${cardBorder}`, fontSize: 12, color: textSoft }}>{e.notes}</td>
                                  <td style={{ padding: "10px 10px", borderBottom: `1px solid ${cardBorder}` }}><button style={{ ...styles.btnSecondary, background: "#fef2f2", color: red, border: "1px solid #fecaca", padding: "4px 10px", fontSize: 12 }} onClick={() => removeEntry(e.id)}>Remove</button></td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      )}
                    </div>
                  </div>
                );
              })()}

              {/* ── EXTRA CHARGES TAB ── */}
              {profitTab === "extra-charges" && (() => {
                const xr = Number(serviceForm?.exchangeRate || USD_TO_TZS);
                const CATEGORIES = ["transport", "tools", "software", "salary", "other"];
                const addEntry = async () => {
                  if (!extraChargesForm.date || (!extraChargesForm.amountUsd && !extraChargesForm.amountTsh)) {
                    setExtraChargesNotice("Date and amount are required.");
                    setTimeout(() => setExtraChargesNotice(""), 3000);
                    return;
                  }
                  const amtUsd = Number(extraChargesForm.amountUsd) || (Number(extraChargesForm.amountTsh) / xr);
                  const amtTsh = Number(extraChargesForm.amountTsh) || (amtUsd * xr);
                  const entry = {
                    date: extraChargesForm.date,
                    category: extraChargesForm.category || "other",
                    description: extraChargesForm.description || "",
                    amountUsd: amtUsd, amountTsh: amtTsh,
                  };
                  // Optimistic: add immediately so data is never lost
                  const localId = `local-${Date.now()}`;
                  const localEntry = { ...entry, id: localId, createdAt: new Date().toISOString() };
                  setExtraCharges((prev) => { const next = [localEntry, ...prev]; writeLS(LS_EXTRA_CHARGES, next); return next; });
                  setExtraChargesForm({ date: "", category: "other", description: "", amountUsd: "", amountTsh: "" });
                  setExtraChargesNotice("Saving…");
                  // Persist to Supabase in background
                  try {
                    const saved = await saveExtraChargeToSupabase(entry);
                    if (saved?.id) {
                      setExtraCharges((prev) => { const next = prev.map((e) => e.id === localId ? { ...localEntry, id: saved.id, createdAt: saved.created_at || localEntry.createdAt } : e); writeLS(LS_EXTRA_CHARGES, next); return next; });
                    }
                    setExtraChargesNotice("Saved.");
                  } catch (err) {
                    setExtraChargesNotice(err?.code === "42P01" ? "Saved locally — run SQL migration 004 in Supabase to enable cloud sync." : "Saved locally (Supabase unavailable).");
                  }
                  setTimeout(() => setExtraChargesNotice(""), 5000);
                };
                const removeEntry = async (id) => {
                  setExtraCharges((prev) => { const next = prev.filter((e) => e.id !== id); writeLS(LS_EXTRA_CHARGES, next); return next; });
                  if (!String(id).startsWith("local-")) await deleteExtraChargeFromSupabase(id).catch(() => {});
                };
                const totalUsd = extraCharges.reduce((s, e) => s + Number(e.amountUsd || 0), 0);
                const byCategory = CATEGORIES.map((cat) => ({ cat, total: extraCharges.filter((e) => e.category === cat).reduce((s, e) => s + Number(e.amountUsd || 0), 0) })).filter((x) => x.total > 0);
                return (
                  <div style={{ display: "grid", gap: 16 }}>
                    <div style={{ display: "grid", gridTemplateColumns: responsiveColumns(`repeat(${Math.min(byCategory.length + 1, 4)}, minmax(0, 1fr))`, "repeat(2, 1fr)", "1fr"), gap: 14 }}>
                      <KpiCard icon={<Archive size={18} />} title="Total Extra Charges" value={formatUSD(totalUsd)} sub={formatTZS(totalUsd * xr)} valueColor={amber} />
                      {byCategory.map(({ cat, total }) => <KpiCard key={cat} title={cat.charAt(0).toUpperCase() + cat.slice(1)} value={formatUSD(total)} sub={formatTZS(total * xr)} valueColor={amber} />)}
                    </div>
                    <div style={{ ...styles.card, padding: 22 }}>
                      <div style={styles.sectionEyebrow}>Add Charge</div>
                      <div style={{ fontSize: 18, fontWeight: 900, marginTop: 6, marginBottom: 14 }}>Extra Business Charges</div>
                      <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("1fr 1fr 1fr 1fr 1fr auto", "1fr 1fr 1fr", "1fr"), gap: 10, alignItems: "end" }}>
                        <div style={styles.fieldBlock}><label style={styles.fieldLabel}>Date</label><input style={styles.input} type="date" value={extraChargesForm.date} onChange={(e) => setExtraChargesForm((f) => ({ ...f, date: e.target.value }))} /></div>
                        <div style={styles.fieldBlock}><label style={styles.fieldLabel}>Category</label><select style={styles.input} value={extraChargesForm.category} onChange={(e) => setExtraChargesForm((f) => ({ ...f, category: e.target.value }))}>{CATEGORIES.map((c) => <option key={c} value={c}>{c.charAt(0).toUpperCase() + c.slice(1)}</option>)}</select></div>
                        <div style={styles.fieldBlock}><label style={styles.fieldLabel}>Description</label><input style={styles.input} placeholder="Details" value={extraChargesForm.description} onChange={(e) => setExtraChargesForm((f) => ({ ...f, description: e.target.value }))} /></div>
                        <div style={styles.fieldBlock}><label style={styles.fieldLabel}>Amount USD</label><input style={styles.input} type="number" min="0" step="0.01" placeholder="0.00" value={extraChargesForm.amountUsd} onChange={(e) => setExtraChargesForm((f) => ({ ...f, amountUsd: e.target.value, amountTsh: "" }))} /></div>
                        <div style={styles.fieldBlock}><label style={styles.fieldLabel}>Amount TSh (alt)</label><input style={styles.input} type="number" min="0" placeholder="0" value={extraChargesForm.amountTsh} onChange={(e) => setExtraChargesForm((f) => ({ ...f, amountTsh: e.target.value, amountUsd: "" }))} /></div>
                        <button style={styles.btnPrimary} onClick={addEntry}>Add</button>
                      </div>
                      {extraChargesNotice && <div style={{ marginTop: 10, color: extraChargesNotice.startsWith("Failed") ? red : green, fontSize: 13, fontWeight: 700 }}>{extraChargesNotice}</div>}
                    </div>
                    <div style={{ ...styles.card, padding: 22 }}>
                      <div style={{ fontSize: 16, fontWeight: 900, marginBottom: 14 }}>All Charges ({extraCharges.length})</div>
                      {extraCharges.length === 0 ? <div style={{ color: textSoft }}>No extra charges yet.</div> : (
                        <div style={{ overflowX: "auto", border: `1px solid ${cardBorder}`, borderRadius: 14 }}>
                          <table style={{ width: "100%", borderCollapse: "separate", borderSpacing: 0 }}>
                            <thead><tr>{["Date", "Category", "Description", "USD", "TSh", ""].map((h) => <th key={h} style={{ textAlign: "left", padding: "10px 10px", color: textSoft, fontSize: 11, fontWeight: 800, textTransform: "uppercase", borderBottom: `1px solid ${cardBorder}`, background: "rgba(247,243,237,0.92)" }}>{h}</th>)}</tr></thead>
                            <tbody>
                              {extraCharges.map((e, i) => (
                                <tr key={e.id} style={{ background: i % 2 === 0 ? "rgba(255,255,255,0.72)" : "rgba(250,247,242,0.8)" }}>
                                  <td style={{ padding: "10px 10px", borderBottom: `1px solid ${cardBorder}`, fontSize: 13 }}>{e.date}</td>
                                  <td style={{ padding: "10px 10px", borderBottom: `1px solid ${cardBorder}`, fontSize: 13 }}>{e.category}</td>
                                  <td style={{ padding: "10px 10px", borderBottom: `1px solid ${cardBorder}`, fontSize: 13 }}>{e.description}</td>
                                  <td style={{ padding: "10px 10px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 700, color: amber }}>{formatUSD(e.amountUsd)}</td>
                                  <td style={{ padding: "10px 10px", borderBottom: `1px solid ${cardBorder}`, fontSize: 12, color: textSoft }}>{formatTZS(e.amountTsh || e.amountUsd * xr)}</td>
                                  <td style={{ padding: "10px 10px", borderBottom: `1px solid ${cardBorder}` }}><button style={{ ...styles.btnSecondary, background: "#fef2f2", color: red, border: "1px solid #fecaca", padding: "4px 10px", fontSize: 12 }} onClick={() => removeEntry(e.id)}>Remove</button></td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      )}
                    </div>
                  </div>
                );
              })()}

              {/* ── CASH BALANCE TAB ── */}
              {profitTab === "cash-balance" && (() => {
                const xrCb = Number(serviceForm?.exchangeRate || USD_TO_TZS);
                const m = profitOverviewMetrics;
                const businessProfitUsd = m.businessProfitUsd;
                const businessProfitTsh = m.businessProfitTzs;
                const totalInjectionUsd = ownerInjections.reduce((s, e) => s + Number(e.amountUsd || 0), 0);
                const cashBalanceUsd = businessProfitUsd + totalInjectionUsd;
                const cashBalanceTsh = cashBalanceUsd * xrCb;
                return (
                  <div style={{ display: "grid", gap: 20 }}>
                    {/* Section A — Business Profit */}
                    <div style={{ ...styles.card, padding: 22 }}>
                      <div style={{ fontSize: 11, fontWeight: 800, letterSpacing: 0.5, textTransform: "uppercase", color: green, marginBottom: 6 }}>Section A — Business Profit</div>
                      <div style={{ fontSize: 20, fontWeight: 900, marginBottom: 14 }}>Revenue − All Costs</div>
                      <div style={{ display: "grid", gap: 0 }}>
                        {[
                          { label: "Revenue", v: m.revenueUsd, tsh: m.revenueTsh, c: green },
                          { label: "Stock Charges", v: -m.stockChargesUsd, tsh: -m.stockChargesTzs, c: amber },
                          { label: "Service Fees", v: -m.serviceChargesUsd, tsh: -m.serviceChargesTzs, c: amber },
                          { label: "Total Ads", v: -m.totalAdsUsd, tsh: -m.totalAdsTzs, c: amber },
                          { label: "Extra Charges", v: -m.extraChargesUsd, tsh: -m.extraChargesTzs, c: amber },
                        ].map((r, i, arr) => (
                          <div key={r.label} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "10px 0", borderBottom: i < arr.length - 1 ? `1px solid ${cardBorder}` : "none" }}>
                            <span style={{ fontWeight: 600 }}>{r.v >= 0 ? "+" : "−"} {r.label}</span>
                            <div style={{ textAlign: "right" }}>
                              <div style={{ fontWeight: 800, color: r.c }}>{formatUSD(Math.abs(r.v))}</div>
                              <div style={{ fontSize: 11, color: textSoft }}>{formatTZS(Math.abs(r.tsh))}</div>
                            </div>
                          </div>
                        ))}
                        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "14px 0", marginTop: 4, borderTop: `2px solid ${cardBorder}` }}>
                          <span style={{ fontSize: 16, fontWeight: 900 }}>= Business Profit</span>
                          <div style={{ textAlign: "right" }}>
                            <div style={{ fontSize: 20, fontWeight: 900, color: businessProfitUsd >= 0 ? green : red }}>{formatUSD(businessProfitUsd)}</div>
                            <div style={{ fontSize: 12, color: textSoft }}>{formatTZS(businessProfitTsh)} · {m.profitMarginPct.toFixed(1)}% margin</div>
                          </div>
                        </div>
                      </div>
                      <div style={{ marginTop: 10, padding: "8px 12px", borderRadius: 10, background: "rgba(31,143,95,0.06)", border: "1px solid rgba(31,143,95,0.12)", fontSize: 12, color: textSoft }}>Owner injection is excluded from business profit. It does NOT count as revenue.</div>
                    </div>

                    {/* Section B — Owner Injections */}
                    <div style={{ ...styles.card, padding: 22 }}>
                      <div style={{ fontSize: 11, fontWeight: 800, letterSpacing: 0.5, textTransform: "uppercase", color: accent, marginBottom: 6 }}>Section B — Owner Injections</div>
                      <div style={{ fontSize: 20, fontWeight: 900, marginBottom: 4 }}>Capital Added by Owner</div>
                      <div style={{ fontSize: 13, color: textSoft, marginBottom: 18 }}>Each injection is logged separately with a date and reason. Never affects Business Profit.</div>

                      {/* Add injection form */}
                      <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("1fr 1fr 1fr 1fr auto", "1fr 1fr 1fr", "1fr"), gap: 12, marginBottom: 18, alignItems: "end" }}>
                        <div style={styles.fieldBlock}>
                          <label style={styles.fieldLabel}>Date</label>
                          <input style={styles.input} type="date" value={ownerInjectionForm.date}
                            onChange={(e) => setOwnerInjectionForm((p) => ({ ...p, date: e.target.value }))} />
                        </div>
                        <div style={styles.fieldBlock}>
                          <label style={styles.fieldLabel}>Amount USD</label>
                          <input style={styles.input} type="number" min="0" step="0.01" placeholder="0.00"
                            value={ownerInjectionForm.amountUsd}
                            onChange={(e) => {
                              const usd = e.target.value;
                              setOwnerInjectionForm((p) => ({ ...p, amountUsd: usd, amountTsh: usd !== "" ? (parseLooseNumber(usd) * xrCb).toFixed(0) : "" }));
                            }} />
                        </div>
                        <div style={styles.fieldBlock}>
                          <label style={styles.fieldLabel}>Amount TSh (auto)</label>
                          <input style={styles.input} type="number" min="0" step="1" placeholder="0"
                            value={ownerInjectionForm.amountTsh}
                            onChange={(e) => {
                              const tsh = e.target.value;
                              setOwnerInjectionForm((p) => ({ ...p, amountTsh: tsh, amountUsd: tsh !== "" ? (parseLooseNumber(tsh) / xrCb).toFixed(2) : "" }));
                            }} />
                        </div>
                        <div style={styles.fieldBlock}>
                          <label style={styles.fieldLabel}>Notes (optional)</label>
                          <input style={styles.input} type="text" placeholder="e.g. restocking capital"
                            value={ownerInjectionForm.notes}
                            onChange={(e) => setOwnerInjectionForm((p) => ({ ...p, notes: e.target.value }))} />
                        </div>
                        <button style={{ ...styles.btnPrimary, whiteSpace: "nowrap" }} onClick={async () => {
                          if (!ownerInjectionForm.date || !ownerInjectionForm.amountUsd) {
                            setOwnerInjectionNotice("Date and Amount USD are required.");
                            setTimeout(() => setOwnerInjectionNotice(""), 3000);
                            return;
                          }
                          const entry = {
                            id: ownerInjectionForm.editId || undefined,
                            date: ownerInjectionForm.date,
                            amountUsd: parseLooseNumber(ownerInjectionForm.amountUsd),
                            amountTsh: parseLooseNumber(ownerInjectionForm.amountTsh) || parseLooseNumber(ownerInjectionForm.amountUsd) * xrCb,
                            notes: ownerInjectionForm.notes.trim(),
                          };
                          // Optimistic: add immediately so data is never lost
                          const localId = `local-${Date.now()}`;
                          const localEntry = { ...entry, id: localId, createdAt: new Date().toISOString() };
                          setOwnerInjections((prev) => { const next = [localEntry, ...prev.filter((e) => e.id !== localId)].sort((a, b) => b.date.localeCompare(a.date)); writeLS(LS_OWNER_INJECTIONS, next); return next; });
                          setOwnerInjectionForm({ date: "", amountUsd: "", amountTsh: "", notes: "" });
                          setOwnerInjectionNotice("Saving…");
                          // Persist to Supabase in background
                          try {
                            const saved = await saveOwnerInjectionToSupabase(entry);
                            if (saved?.id) {
                              setOwnerInjections((prev) => {
                                const next = prev.map((e) => e.id === localId ? { id: saved.id, date: saved.date, amountUsd: Number(saved.amount_usd), amountTsh: Number(saved.amount_tsh), notes: saved.notes || "", createdAt: saved.created_at || localEntry.createdAt } : e).sort((a, b) => b.date.localeCompare(a.date));
                                writeLS(LS_OWNER_INJECTIONS, next);
                                return next;
                              });
                            }
                            setOwnerInjectionNotice("Saved.");
                          } catch (err) {
                            setOwnerInjectionNotice(err?.code === "42P01" ? "Saved locally — run SQL migration 005 in Supabase to enable cloud sync." : "Saved locally (Supabase unavailable).");
                          }
                          setTimeout(() => setOwnerInjectionNotice(""), 5000);
                        }}>+ Add Injection</button>
                      </div>
                      {ownerInjectionNotice && <div style={{ fontSize: 12, color: accent, marginBottom: 12 }}>{ownerInjectionNotice}</div>}

                      {/* Entries table */}
                      {ownerInjections.length === 0 ? (
                        <div style={{ fontSize: 13, color: textSoft, padding: "16px 0" }}>No injections recorded yet.</div>
                      ) : (
                        <div style={{ overflowX: "auto" }}>
                          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                            <thead>
                              <tr style={{ borderBottom: `2px solid ${cardBorder}` }}>
                                {["Date", "Amount USD", "Amount TSh", "Notes", ""].map((h) => (
                                  <th key={h} style={{ textAlign: "left", padding: "6px 10px", fontWeight: 700, color: textSoft, fontSize: 11, textTransform: "uppercase" }}>{h}</th>
                                ))}
                              </tr>
                            </thead>
                            <tbody>
                              {ownerInjections.map((inj, i) => (
                                <tr key={inj.id} style={{ borderBottom: i < ownerInjections.length - 1 ? `1px solid ${cardBorder}` : "none", background: i % 2 === 0 ? "transparent" : "rgba(255,255,255,0.02)" }}>
                                  <td style={{ padding: "9px 10px", fontWeight: 600 }}>{inj.date}</td>
                                  <td style={{ padding: "9px 10px", fontWeight: 700, color: green }}>{formatUSD(inj.amountUsd)}</td>
                                  <td style={{ padding: "9px 10px", color: textSoft }}>{formatTZS(inj.amountTsh)}</td>
                                  <td style={{ padding: "9px 10px", color: textSoft, maxWidth: 200 }}>{inj.notes || "—"}</td>
                                  <td style={{ padding: "9px 10px" }}>
                                    <button style={{ ...styles.btnSecondary, background: "#fef2f2", color: red, border: "1px solid #fecaca", padding: "4px 10px", fontSize: 12 }} onClick={async () => {
                                      setOwnerInjections((prev) => { const next = prev.filter((e) => e.id !== inj.id); writeLS(LS_OWNER_INJECTIONS, next); return next; });
                                      if (!String(inj.id).startsWith("local-")) await deleteOwnerInjectionFromSupabase(inj.id).catch(() => {});
                                    }}>Remove</button>
                                  </td>
                                </tr>
                              ))}
                            </tbody>
                            <tfoot>
                              <tr style={{ borderTop: `2px solid ${cardBorder}` }}>
                                <td style={{ padding: "10px 10px", fontWeight: 900 }}>Total ({ownerInjections.length})</td>
                                <td style={{ padding: "10px 10px", fontWeight: 900, color: green }}>{formatUSD(totalInjectionUsd)}</td>
                                <td style={{ padding: "10px 10px", color: textSoft }}>{formatTZS(totalInjectionUsd * xrCb)}</td>
                                <td colSpan={2} />
                              </tr>
                            </tfoot>
                          </table>
                        </div>
                      )}
                    </div>

                    {/* Section C — Cash Balance */}
                    <div style={{ ...styles.card, padding: 22 }}>
                      <div style={{ fontSize: 11, fontWeight: 800, letterSpacing: 0.5, textTransform: "uppercase", color: accent, marginBottom: 6 }}>Section C — Cash Balance</div>
                      <div style={{ fontSize: 20, fontWeight: 900, marginBottom: 18 }}>Business Profit + Owner Injections</div>
                      <div style={{ display: "grid", gap: 0 }}>
                        {[
                          { label: "Business Profit", v: businessProfitUsd, tsh: businessProfitTsh, c: businessProfitUsd >= 0 ? green : red, sign: businessProfitUsd >= 0 ? "+" : "−" },
                          { label: `Owner Injections (${ownerInjections.length})`, v: totalInjectionUsd, tsh: totalInjectionUsd * xrCb, c: accent, sign: "+" },
                        ].map((r, i, arr) => (
                          <div key={r.label} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "10px 0", borderBottom: i < arr.length - 1 ? `1px solid ${cardBorder}` : "none" }}>
                            <span style={{ fontWeight: 600 }}>{r.sign} {r.label}</span>
                            <div style={{ textAlign: "right" }}>
                              <div style={{ fontWeight: 800, color: r.c }}>{formatUSD(Math.abs(r.v))}</div>
                              <div style={{ fontSize: 11, color: textSoft }}>{formatTZS(Math.abs(r.tsh))}</div>
                            </div>
                          </div>
                        ))}
                        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "14px 0", marginTop: 4, borderTop: `2px solid ${cardBorder}` }}>
                          <span style={{ fontSize: 18, fontWeight: 900 }}>= Cash Balance</span>
                          <div style={{ textAlign: "right" }}>
                            <div style={{ fontSize: 24, fontWeight: 900, color: cashBalanceUsd >= 0 ? green : red }}>{formatUSD(cashBalanceUsd)}</div>
                            <div style={{ fontSize: 12, color: textSoft }}>{formatTZS(cashBalanceTsh)}</div>
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>
                );
              })()}

              {/* ── PRODUCT SIMULATOR TAB ── */}
              {profitTab === "simulator" && (() => {
                const xrSim = 2850;
                const n = (key) => parseLooseNumber(simInputs[key] || "0");
                const totalLeads = n("totalLeads");
                const confirmRate = n("confirmationRate") / 100;
                const delivRate = n("deliveryRate") / 100;
                const cpl = n("cpl");
                const sellingPriceTsh = n("sellingPriceTsh");
                const productCostUsd = n("productCostUsd");
                const serviceFee = n("serviceFeePerUnit") || 9;

                const confirmedLeads = totalLeads * confirmRate;
                const deliveredLeads = confirmedLeads * delivRate;
                const totalAdsSpend = totalLeads * cpl;
                const revenueTsh = deliveredLeads * sellingPriceTsh;
                const revenueUsd = revenueTsh / xrSim;
                const productCostTotal = deliveredLeads * productCostUsd;
                const serviceFeeTotal = deliveredLeads * serviceFee;
                const sellingPriceUsd = sellingPriceTsh / xrSim;
                const adsCostPerDelivered = deliveredLeads > 0 ? totalAdsSpend / deliveredLeads : 0;
                const profitPerUnit = sellingPriceUsd - productCostUsd - adsCostPerDelivered - serviceFee;
                const totalProfitUsd = deliveredLeads * profitPerUnit;
                const totalProfitTsh = totalProfitUsd * xrSim;
                const maxCpl = deliveredLeads > 0 && totalLeads > 0
                  ? (sellingPriceUsd - productCostUsd - serviceFee) * (deliveredLeads / totalLeads)
                  : 0;
                const roi = totalAdsSpend > 0 ? (totalProfitUsd / totalAdsSpend) * 100 : 0;
                const hasInputs = totalLeads > 0 && deliveredLeads > 0;
                const isProfit = totalProfitUsd >= 0;
                const isCplGood = maxCpl >= cpl;

                const simField = (label, key, opts = {}) => (
                  <div style={styles.fieldBlock}>
                    <label style={styles.fieldLabel}>{label}</label>
                    <input
                      style={styles.input}
                      type="number"
                      min="0"
                      step={opts.step || "any"}
                      placeholder={opts.placeholder || "0"}
                      value={simInputs[key]}
                      onChange={(e) => setSimInputs((prev) => ({ ...prev, [key]: e.target.value }))}
                    />
                  </div>
                );

                return (
                  <div style={{ display: "grid", gap: 16 }}>
                    {/* Inputs */}
                    <div style={{ ...styles.card, padding: 22 }}>
                      <div style={styles.sectionEyebrow}>Simulator</div>
                      <div style={{ fontSize: 20, fontWeight: 900, marginTop: 6, marginBottom: 18 }}>Enter your numbers</div>
                      <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(3, minmax(0, 1fr))", "repeat(2, minmax(0, 1fr))", "1fr"), gap: 14 }}>
                        {simField("Total Leads", "totalLeads")}
                        {simField("Confirmation Rate %", "confirmationRate", { placeholder: "45" })}
                        {simField("Delivery Rate %", "deliveryRate", { placeholder: "30" })}
                        {simField("CPL — Cost Per Lead (USD)", "cpl", { step: "0.01", placeholder: "0.50" })}
                        {simField("Selling Price (TSh)", "sellingPriceTsh", { placeholder: "50000" })}
                        {simField("Product Cost (USD)", "productCostUsd", { step: "0.01", placeholder: "5.00" })}
                        {simField("Service Fee / Unit (USD)", "serviceFeePerUnit", { step: "0.01", placeholder: "9" })}
                      </div>
                    </div>

                    {/* Funnel + Financial side by side */}
                    <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("1fr 1fr", "1fr", "1fr"), gap: 16 }}>
                      {/* Funnel */}
                      <div style={{ ...styles.card, padding: 22 }}>
                        <div style={styles.sectionEyebrow}>Funnel</div>
                        <div style={{ fontSize: 18, fontWeight: 900, marginTop: 6, marginBottom: 18 }}>Lead → Confirmed → Delivered</div>
                        {[
                          { label: "Total Leads", value: hasInputs ? Math.round(totalLeads).toLocaleString() : "—", color: textMain },
                          { label: `Confirmed (${simInputs.confirmationRate || 0}%)`, value: hasInputs ? Math.round(confirmedLeads).toLocaleString() : "—", color: accent },
                          { label: `Delivered (${simInputs.deliveryRate || 0}%)`, value: hasInputs ? Math.round(deliveredLeads).toLocaleString() : "—", color: green },
                        ].map((row, i, arr) => (
                          <div key={row.label} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "12px 0", borderBottom: i < arr.length - 1 ? `1px solid ${cardBorder}` : "none" }}>
                            <span style={{ fontWeight: 600 }}>{row.label}</span>
                            <span style={{ fontSize: 22, fontWeight: 900, color: row.color }}>{row.value}</span>
                          </div>
                        ))}
                      </div>

                      {/* Financial results */}
                      <div style={{ ...styles.card, padding: 22 }}>
                        <div style={styles.sectionEyebrow}>Financial Results</div>
                        <div style={{ fontSize: 18, fontWeight: 900, marginTop: 6, marginBottom: 18 }}>Revenue → Costs → Profit</div>
                        {[
                          { label: "Revenue", value: hasInputs ? formatUSD(revenueUsd) : "—", sub: hasInputs ? formatTZS(revenueTsh) : "", color: green },
                          { label: "− Product Cost", value: hasInputs ? formatUSD(productCostTotal) : "—", sub: hasInputs ? `$${productCostUsd.toFixed(2)} × ${Math.round(deliveredLeads)} units` : "", color: amber },
                          { label: "− Service Fees", value: hasInputs ? formatUSD(serviceFeeTotal) : "—", sub: hasInputs ? `$${serviceFee} × ${Math.round(deliveredLeads)} units` : "", color: amber },
                          { label: "− Ads Spend", value: hasInputs ? formatUSD(totalAdsSpend) : "—", sub: hasInputs ? `$${cpl.toFixed(2)} CPL × ${Math.round(totalLeads)} leads` : "", color: amber },
                        ].map((row, i, arr) => (
                          <div key={row.label} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "10px 0", borderBottom: i < arr.length - 1 ? `1px solid ${cardBorder}` : "none" }}>
                            <div>
                              <span style={{ fontWeight: 600 }}>{row.label}</span>
                              {row.sub && <div style={{ fontSize: 11, color: textSoft, marginTop: 2 }}>{row.sub}</div>}
                            </div>
                            <span style={{ fontWeight: 800, color: row.color }}>{row.value}</span>
                          </div>
                        ))}
                        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "14px 0", marginTop: 4, borderTop: `2px solid ${cardBorder}` }}>
                          <span style={{ fontSize: 16, fontWeight: 900 }}>= Total Profit</span>
                          <div style={{ textAlign: "right" }}>
                            <div style={{ fontSize: 22, fontWeight: 900, color: isProfit ? green : red }}>{hasInputs ? formatUSD(totalProfitUsd) : "—"}</div>
                            {hasInputs && <div style={{ fontSize: 11, color: textSoft }}>{formatTZS(totalProfitTsh)}</div>}
                          </div>
                        </div>
                      </div>
                    </div>

                    {/* Key metrics */}
                    <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(3, minmax(0, 1fr))", "repeat(2, minmax(0, 1fr))", "1fr"), gap: 14 }}>
                      <KpiCard
                        icon={<Wallet size={18} />}
                        title="Profit / Unit"
                        value={hasInputs ? formatUSD(profitPerUnit) : "—"}
                        sub="Per delivered unit after all costs"
                        valueColor={profitPerUnit >= 0 ? green : red}
                      />
                      <KpiCard
                        icon={<Calculator size={18} />}
                        title="Max CPL (Break-even)"
                        value={hasInputs ? formatUSD(maxCpl) : "—"}
                        sub={
                          hasInputs && cpl > 0
                            ? isCplGood
                              ? `Your CPL $${cpl.toFixed(2)} — profitable`
                              : `Your CPL $${cpl.toFixed(2)} — losing money`
                            : "Enter CPL to compare"
                        }
                        valueColor={hasInputs && cpl > 0 ? (isCplGood ? green : red) : textMain}
                      />
                      <KpiCard
                        icon={<TrendingUp size={18} />}
                        title="ROI"
                        value={hasInputs && totalAdsSpend > 0 ? `${roi.toFixed(1)}%` : "—"}
                        sub="Total Profit / Total Ads Spend"
                        valueColor={roi >= 0 ? green : red}
                      />
                    </div>
                  </div>
                );
              })()}

            </div>
          )}

{["scaling", "performanceHub"].includes(activePage) && (
            <div style={{ display: "grid", gap: 20 }}>
              <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(4, minmax(0, 1fr))", "repeat(2, minmax(0, 1fr))", "1fr"), gap: 16 }}>
                <KpiCard icon={<Rocket size={18} />} title="Ready To Scale" value={scalingSummary.ready.length} sub="Products with strong profit, ROAS, delivery and stock" valueColor={green} />
                <KpiCard icon={<TrendingUp size={18} />} title="Watchlist" value={scalingSummary.watch.length} sub="Close to scale but still need optimization" valueColor={amber} />
                <KpiCard icon={<AlertTriangle size={18} />} title="Blocked" value={scalingSummary.blocked.length} sub="Fix the blockers before increasing spend" valueColor={red} />
                <KpiCard icon={<Boxes size={18} />} title="Top Candidate" value={scalingSummary.topCandidate?.name || "N/A"} sub={scalingSummary.topCandidate ? `${scalingSummary.topCandidate.scaleReadiness}% readiness` : "No scaling data yet"} />
              </div>

              <div style={{ ...styles.card, padding: 22 }}>
                <div style={styles.sectionHeader}>
                  <div>
                    <div style={styles.sectionEyebrow}>Scaling engine</div>
                    <div style={{ fontSize: 24, fontWeight: 900, marginTop: 8 }}>Products to scale</div>
                    <div style={{ color: textSoft, marginTop: 6, lineHeight: 1.6 }}>
                      Cette page analyse automatiquement chaque produit selon le profit, la confirmation, la livraison, le ROAS et le stock disponible.
                    </div>
                  </div>
                </div>

                <div style={{ display: "grid", gap: 14 }}>
                  {scalingInsights.length ? (
                    scalingInsights.map((product) => (
                      <div
                        key={product.id}
                        style={{
                          ...styles.softStat,
                          border: product.shouldScale ? "1px solid rgba(31,143,95,0.22)" : product.scaleReadiness >= 60 ? "1px solid rgba(199,131,34,0.22)" : `1px solid ${cardBorder}`,
                          background: product.shouldScale
                            ? "linear-gradient(180deg, rgba(236,253,245,0.92), rgba(255,255,255,0.9))"
                            : product.scaleReadiness >= 60
                              ? "linear-gradient(180deg, rgba(255,251,235,0.92), rgba(255,255,255,0.9))"
                              : "linear-gradient(180deg, rgba(255,255,255,0.94), rgba(248,244,238,0.88))",
                        }}
                      >
                        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap" }}>
                          <div>
                            <div style={{ fontSize: 20, fontWeight: 900 }}>{product.name}</div>
                            <div style={{ color: textSoft, marginTop: 6 }}>
                              {product.orders} orders | {Math.round(product.confirmRate * 100)}% confirm | {Math.round(product.deliveryRate * 100)}% deliver | ROAS {Number(product.roas || 0).toFixed(2)}
                            </div>
                          </div>
                          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                            <div style={getDecisionStyle(product.shouldScale ? "SCALE" : product.scaleReadiness >= 60 ? "WATCH" : "KILL")}>
                              {product.recommendedAction}
                            </div>
                            <div style={{ ...styles.badge, background: "rgba(29,95,208,0.08)", color: accent, border: "1px solid rgba(29,95,208,0.12)" }}>
                              {product.scaleReadiness}% readiness
                            </div>
                          </div>
                        </div>

                        <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(5, minmax(0, 1fr))", "1fr 1fr", "1fr"), gap: 12, marginTop: 16 }}>
                          <MiniStat label="Profit" value={formatTZS(product.profit)} tone={product.profit >= 0 ? "green" : "amber"} sub="Net product result" />
                          <MiniStat label="Revenue" value={formatTZS(product.revenue)} tone="blue" sub={`${product.deliveredUnits} delivered units`} />
                          <MiniStat label="Available Stock" value={product.availableStock} tone="amber" sub={`Reorder point ${product.reorderPoint}`} />
                          <MiniStat label="Reserved" value={product.reservedStock} tone="blue" sub={`${product.returnedUnits || 0} returned to stock`} />
                          <MiniStat label="Score" value={`${product.score}/100`} tone="green" sub={product.decision} />
                        </div>

                        <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("1fr 1fr", "1fr", "1fr"), gap: 12, marginTop: 16 }}>
                          <div style={{ ...styles.card, padding: 14, background: "rgba(255,255,255,0.7)" }}>
                            <div style={{ fontSize: 11, fontWeight: 800, letterSpacing: 0.45, textTransform: "uppercase", color: textSoft }}>Strengths</div>
                            <div style={{ display: "grid", gap: 6, marginTop: 10, color: textMain, fontSize: 14 }}>
                              {product.strengths.length ? product.strengths.map((item) => <div key={item}>- {item}</div>) : <div>No clear strength yet.</div>}
                            </div>
                          </div>
                          <div style={{ ...styles.card, padding: 14, background: "rgba(255,255,255,0.7)" }}>
                            <div style={{ fontSize: 11, fontWeight: 800, letterSpacing: 0.45, textTransform: "uppercase", color: textSoft }}>Blockers</div>
                            <div style={{ display: "grid", gap: 6, marginTop: 10, color: textMain, fontSize: 14 }}>
                              {product.blockers.length ? product.blockers.map((item) => <div key={item}>- {item}</div>) : <div>No blocker detected.</div>}
                            </div>
                          </div>
                        </div>
                      </div>
                    ))
                  ) : (
                    <div style={{ color: textSoft }}>No product data yet. Add products, tracking rows and orders to activate scaling suggestions.</div>
                  )}
                </div>
              </div>
            </div>
          )}

{["taskCenter", "operationsHub"].includes(activePage) && (
            <div style={{ display: "grid", gap: 20 }}>
              <PageHeader
                eyebrow="Decisions"
                title="Action board"
                description="See what to scale, pause, restock or fix based on the live business signals already calculated by the app."
              />
              <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(4, minmax(0, 1fr))", "1fr 1fr", "1fr"), gap: 16 }}>
                <KpiCard icon={<ClipboardList size={18} />} title="Open tasks" value={taskCenterData.length} sub="Business tasks generated from live data" />
                <KpiCard icon={<AlertTriangle size={18} />} title="High priority" value={taskCenterData.filter((task) => task.priority === "High").length} sub="Needs action today" valueColor={red} />
                <KpiCard icon={<Archive size={18} />} title="Stock tasks" value={taskCenterData.filter((task) => task.type === "stock").length} sub="Reorder or forecast issues" valueColor={amber} />
                <KpiCard icon={<Rocket size={18} />} title="Scaling tasks" value={taskCenterData.filter((task) => task.type === "scaling").length} sub="Products ready for budget increase" valueColor={green} />
              </div>

              <div style={{ ...styles.card, padding: 22 }}>
                <div style={styles.sectionHeader}>
                  <div>
                    <div style={styles.sectionEyebrow}>Task center</div>
                    <div style={{ fontSize: 24, fontWeight: 900, marginTop: 8 }}>Priority actions for the business</div>
                    <div style={{ color: textSoft, marginTop: 6, lineHeight: 1.6 }}>
                      A single inbox for stock, shipping, scaling and anomaly actions generated automatically from the live app data.
                    </div>
                  </div>
                </div>

                <div style={{ display: "grid", gap: 12 }}>
                  {taskCenterData.length ? taskCenterData.map((task) => (
                    <div key={task.id} style={{ ...styles.softStat, display: "flex", alignItems: "center", justifyContent: "space-between", gap: 16, flexWrap: "wrap" }}>
                      <div style={{ minWidth: 260, flex: "1 1 320px" }}>
                        <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
                          <span style={getDecisionStyle(task.priority === "High" ? "KILL" : task.priority === "Medium" ? "WATCH" : "OK")}>{task.priority}</span>
                          <span style={{ ...styles.badge, background: "rgba(35,88,213,0.08)", color: accent, border: "1px solid rgba(35,88,213,0.12)" }}>{formatStatusLabel(task.type)}</span>
                          <span style={{ color: textSoft, fontSize: 12 }}>{task.owner}</span>
                        </div>
                        <div style={{ fontWeight: 800, fontSize: 16, marginTop: 8 }}>{task.title}</div>
                        <div style={{ color: textSoft, marginTop: 6, lineHeight: 1.5 }}>{task.detail}</div>
                      </div>
                      <button style={styles.btnPrimary} onClick={() => setActivePage(task.page)}>
                        Open
                      </button>
                    </div>
                  )) : <div style={{ color: textSoft }}>No business task detected right now.</div>}
                </div>
              </div>

              <div style={{ ...styles.card, padding: 22 }}>
                <div style={styles.sectionHeader}>
                  <div>
                    <div style={styles.sectionEyebrow}>Team mode</div>
                    <div style={{ fontSize: 24, fontWeight: 900, marginTop: 8 }}>Workload by owner</div>
                    <div style={{ color: textSoft, marginTop: 6, lineHeight: 1.6 }}>
                      Orders and shipping can now be assigned to team members. This view helps you see who is carrying what.
                    </div>
                  </div>
                </div>
                <div style={{ overflowX: "auto", border: `1px solid ${cardBorder}`, borderRadius: 20 }}>
                  <table style={{ width: "100%", borderCollapse: "separate", borderSpacing: 0 }}>
                    <thead>
                      <tr>
                        {["Owner", "Total Orders", "Confirmed", "In Shipping", "Delivered"].map((head) => (
                          <th key={head} style={{ textAlign: "left", padding: "14px 12px", color: textSoft, fontSize: 12, fontWeight: 800, letterSpacing: 0.4, textTransform: "uppercase", borderBottom: `1px solid ${cardBorder}`, background: "rgba(247, 243, 237, 0.92)" }}>
                            {head}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {teamWorkloadRows.map((row, index) => (
                        <tr key={row.owner} style={{ background: index % 2 === 0 ? "rgba(255,255,255,0.72)" : "rgba(250,247,242,0.8)" }}>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 800 }}>{row.owner}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>{row.total}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>{row.confirmed}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>{row.shipping}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>{row.delivered}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  {teamWorkloadRows.length === 0 ? <div style={{ padding: 24, color: textSoft }}>No owner assigned yet. Use the owner dropdowns in Orders and Shipping.</div> : null}
                </div>
              </div>
            </div>
          )}

{["calendar", "operationsHub"].includes(activePage) && (
            <div style={{ display: "grid", gap: 20 }}>
              <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(3, minmax(0, 1fr))", "1fr 1fr", "1fr"), gap: 16 }}>
                <KpiCard icon={<CalendarDays size={18} />} title="Upcoming events" value={calendarEvents.length} sub="Operational reminders and projections" />
                <KpiCard icon={<Archive size={18} />} title="Stock events" value={calendarEvents.filter((event) => event.type === "stock").length} sub="Projected stockout dates" valueColor={amber} />
                <KpiCard icon={<ShoppingBag size={18} />} title="Shipping events" value={calendarEvents.filter((event) => event.type === "shipping").length} sub="Daily shipping reminders" valueColor={accent} />
              </div>

              <div style={{ ...styles.card, padding: 22 }}>
                <div style={styles.sectionHeader}>
                  <div>
                    <div style={styles.sectionEyebrow}>Business calendar</div>
                    <div style={{ fontSize: 24, fontWeight: 900, marginTop: 8 }}>Upcoming operational timeline</div>
                    <div style={{ color: textSoft, marginTop: 6, lineHeight: 1.6 }}>
                      All the dates that matter to keep the business moving: stockouts, shipping reminders and follow-ups.
                    </div>
                  </div>
                </div>

                <div style={{ display: "grid", gap: 12 }}>
                  {calendarEvents.length ? calendarEvents.map((event) => (
                    <div key={event.id} style={{ ...styles.softStat, display: "grid", gap: 6 }}>
                      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap" }}>
                        <div style={{ fontWeight: 800 }}>{event.title}</div>
                        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                          <span style={{ ...styles.badge, background: "rgba(35,88,213,0.08)", color: accent, border: "1px solid rgba(35,88,213,0.12)" }}>{formatStatusLabel(event.type)}</span>
                          <span style={{ fontWeight: 800, color: textMain }}>{event.date}</span>
                        </div>
                      </div>
                      <div style={{ color: textSoft }}>{event.detail}</div>
                    </div>
                  )) : <div style={{ color: textSoft }}>No calendar event generated yet.</div>}
                </div>
              </div>
            </div>
          )}

{["team", "operationsHub"].includes(activePage) && (
            <div style={{ display: "grid", gap: 20 }}>
              <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(4, minmax(0, 1fr))", "repeat(2, minmax(0, 1fr))", "1fr"), gap: 16 }}>
                <KpiCard icon={<Users size={18} />} title="Owners Active" value={teamScorecardRows.length} sub="People with assigned orders" />
                <KpiCard icon={<ShoppingBag size={18} />} title="Delivered Orders" value={teamScorecardRows.reduce((sum, row) => sum + Number(row.deliveredOrders || 0), 0)} sub="Delivered orders across all owners" valueColor={green} />
                <KpiCard icon={<Wallet size={18} />} title="Payroll" value={formatUsdFromTzs(situationsSummary.salariesTotalTzs)} sub="Registered salary base" valueColor={amber} />
                <KpiCard icon={<TrendingUp size={18} />} title="Top Owner" value={teamScorecardRows[0]?.owner || "N/A"} sub={teamScorecardRows[0] ? `${formatUsdFromTzs(teamScorecardRows[0].revenueTzs)} delivered revenue` : "Assign orders to activate"} />
              </div>

              <div style={{ ...styles.card, padding: 22 }}>
                <div style={styles.sectionHeader}>
                  <div>
                    <div style={styles.sectionEyebrow}>Team scorecards</div>
                    <div style={{ fontSize: 24, fontWeight: 900, marginTop: 8 }}>Performance by owner</div>
                    <div style={{ color: textSoft, marginTop: 6, lineHeight: 1.6 }}>
                      Follow confirmations, shipping load, delivered revenue and the margin generated by each owner from the live order pipeline.
                    </div>
                  </div>
                </div>

                <div style={{ overflowX: "auto", border: `1px solid ${cardBorder}`, borderRadius: 20 }}>
                  <table style={{ width: "100%", borderCollapse: "separate", borderSpacing: 0 }}>
                    <thead>
                      <tr>
                        {["Owner", "Orders", "Confirmed", "Shipping", "Delivered", "Confirm %", "Deliver %", "Revenue", "Estimated Margin", "Salary", "Net After Salary"].map((head) => (
                          <th key={head} style={{ textAlign: "left", padding: "14px 12px", color: textSoft, fontSize: 12, fontWeight: 800, letterSpacing: 0.4, textTransform: "uppercase", borderBottom: `1px solid ${cardBorder}`, background: "rgba(247, 243, 237, 0.92)" }}>
                            {head}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {teamScorecardRows.map((row, index) => (
                        <tr key={row.owner} style={{ background: index % 2 === 0 ? "rgba(255,255,255,0.72)" : "rgba(250,247,242,0.8)" }}>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 800 }}>{row.owner}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>{row.totalOrders}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>{row.confirmedOrders}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>{row.shippingOrders}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>{row.deliveredOrders}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>{Number(row.confirmationRate || 0).toFixed(0)}%</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>{Number(row.deliveryRate || 0).toFixed(0)}%</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>{formatUsdFromTzs(row.revenueTzs)}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}`, color: Number(row.profitTzs || 0) >= 0 ? green : red }}>{formatUsdFromTzs(row.profitTzs)}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>{formatUsdFromTzs(row.salaryTzs)}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}`, color: Number(row.netAfterSalaryTzs || 0) >= 0 ? green : red, fontWeight: 800 }}>{formatUsdFromTzs(row.netAfterSalaryTzs)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  {teamScorecardRows.length === 0 ? <div style={{ padding: 24, color: textSoft }}>No team scorecard yet. Assign owners inside Orders and Shipping first.</div> : null}
                </div>
              </div>
            </div>
          )}

{activePage === "aiAssistant" && (
            <div style={{ display: "grid", gap: 20 }}>
              <PageHeader
                eyebrow="AI Assistant"
                title="Business analyst workspace"
                description="Review real app outputs, alerts and decision signals before asking the assistant what to scale, stop or fix."
              />
              <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(4, minmax(0, 1fr))", "repeat(2, minmax(0, 1fr))", "1fr"), gap: 16 }}>
                <KpiCard icon={<TrendingUp size={18} />} title="Best Product" value={bestProduct?.name || "N/A"} sub={bestProduct ? formatTZS(bestProduct.dashboardProfitTzs || 0) : "No winning product yet"} valueColor={green} />
                <KpiCard icon={<AlertTriangle size={18} />} title="Biggest Problem" value={taskCenterData[0]?.title || "No blocker"} sub={taskCenterData[0]?.detail || "Nothing critical right now"} valueColor={red} />
                <KpiCard icon={<ClipboardList size={18} />} title="Top Alerts" value={productAlertsSummary.topRows.length} sub="Critical product warnings" valueColor={amber} />
                <KpiCard icon={<Rocket size={18} />} title="Decisions Open" value={controlPanelSummary.needsAttentionProducts.length} sub="Products under review" />
              </div>
              <div style={{ ...styles.card, padding: 22, display: ordersTab === "pipeline" ? "block" : "none" }}>
                <div style={styles.sectionHeader}>
                  <div>
                    <div style={styles.sectionEyebrow}>AI brief</div>
                    <div style={{ fontSize: 24, fontWeight: 900, marginTop: 8 }}>Live data briefing</div>
                    <div style={{ color: textSoft, marginTop: 6, lineHeight: 1.6 }}>
                      This page stays tied to real delivered revenue, ads spend, stock pressure, alert rules and the existing decision engine outputs.
                    </div>
                  </div>
                </div>
                <div style={{ display: "grid", gap: 12 }}>
                  {productAlertsSummary.topRows.slice(0, 5).map((row) => (
                    <div key={`ai-${row.id}`} style={{ ...styles.softStat, display: "grid", gap: 8 }}>
                      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, flexWrap: "wrap" }}>
                        <div style={{ fontWeight: 800 }}>{row.name}</div>
                        <div style={getDecisionStyle(row.performanceStatus === "WINNER" ? "SCALE" : row.performanceStatus === "LOSS" ? "KILL" : "WATCH")}>
                          {row.performanceStatus}
                        </div>
                      </div>
                      <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                        {row.productAlerts.map((alert) => (
                          <span key={`ai-alert-${row.id}-${alert.key}`} style={getAlertBadgeStyle(alert.tone)}>
                            {alert.message}
                          </span>
                        ))}
                      </div>
                    </div>
                  ))}
                  {productAlertsSummary.topRows.length === 0 ? <div style={{ color: textSoft }}>No critical AI briefing generated yet.</div> : null}
                </div>
              </div>
            </div>
          )}

{["audit", "operationsHub", "settingsAudit"].includes(activePage) && (
            <div style={{ display: "grid", gap: 20 }}>
              <PageHeader
                eyebrow="Settings & Audit"
                title="Workspace, backups and technical audit"
                description="Keep cloud access, restore points, exports and technical traceability outside the daily operating workflow."
                action={(
                  <>
                    <button style={styles.btnSecondary} onClick={exportAllDataToCsv}>Export All CSV</button>
                    <button style={styles.btnSecondary} onClick={exportProductPerformanceToCsv}>Export Product CSV</button>
                    <button style={styles.btnSecondary} onClick={backupAllAppDataToJson}>Backup JSON</button>
                    <button style={styles.btnSecondary} onClick={() => restoreJsonInputRef.current?.click()}>Restore JSON</button>
                  </>
                )}
              />
              <InlineTabs
                items={[
                  { value: "workspace", label: "Workspace" },
                  { value: "backups", label: "Backups" },
                  { value: "exports", label: "Exports" },
                  { value: "audit", label: "Audit Trail" },
                ]}
                value={settingsAuditTab}
                onChange={setSettingsAuditTab}
              />
              <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("1.2fr 1fr", "1fr", "1fr"), gap: 16 }}>
                <div style={{ ...styles.card, padding: 18 }}>
                  <div style={styles.sectionEyebrow}>Workspace</div>
                  <div style={{ fontSize: 22, fontWeight: 900, marginTop: 8 }}>{cloudAuth.user?.email || "Cloud access not connected"}</div>
                  <div style={{ color: textSoft, marginTop: 8, lineHeight: 1.6 }}>
                    Workspace ID: {supabaseWorkspaceId || "N/A"} | Sync {cloudBackupState.syncing ? "in progress" : "ready"}
                  </div>
                  <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginTop: 14 }}>
                    <button style={styles.btnSecondary} onClick={() => void refreshCloudBackups()}>Refresh backups</button>
                    <button style={styles.btnSecondary} onClick={() => setShowCloudBackups((prev) => !prev)}>
                      {showCloudBackups ? "Hide restore points" : "Show restore points"}
                    </button>
                  </div>
                  {supabaseEnabled && (
                    <div style={{ marginTop: 18, paddingTop: 16, borderTop: `1px solid ${cardBorder}` }}>
                      <div style={{ fontSize: 11, fontWeight: 800, letterSpacing: 0.45, textTransform: "uppercase", color: amber, marginBottom: 8 }}>One-time migration</div>
                      <div style={{ color: textSoft, fontSize: 13, marginBottom: 12, lineHeight: 1.5 }}>
                        Push your local browser data to the cloud workspace. Run once when migrating to a new device or first cloud setup.
                      </div>
                      <button
                        style={{ ...styles.btnPrimary, opacity: migrating ? 0.6 : 1 }}
                        onClick={() => void handleMigrateLocalToCloud()}
                        disabled={migrating}
                      >
                        {migrating ? "Migrating..." : "Migrate localStorage → Supabase"}
                      </button>
                      {migrateNotice ? (
                        <div style={{ marginTop: 10, padding: "10px 14px", borderRadius: 12, background: migrateNotice.startsWith("Migration complete") ? "rgba(21,143,99,0.08)" : "rgba(217,72,95,0.08)", color: migrateNotice.startsWith("Migration complete") ? green : red, fontSize: 13, fontWeight: 600, lineHeight: 1.5 }}>
                          {migrateNotice}
                        </div>
                      ) : null}
                      {syncNotice ? (
                        <div style={{ marginTop: 10, padding: "10px 14px", borderRadius: 12, background: "rgba(21,143,99,0.08)", color: green, fontSize: 13, fontWeight: 600, lineHeight: 1.5 }}>
                          {syncNotice}
                        </div>
                      ) : null}
                    </div>
                  )}
                </div>
                <div style={{ ...styles.card, padding: 18 }}>
                  <div style={styles.sectionEyebrow}>Exports & restore</div>
                  <div style={{ fontSize: 22, fontWeight: 900, marginTop: 8 }}>Data portability</div>
                  <div style={{ color: textSoft, marginTop: 8, lineHeight: 1.6 }}>
                    Export the full dataset, product performance reports, or restore a JSON backup without breaking the current structure.
                  </div>
                </div>
              </div>
              {/* Business settings — exchange rate */}
              <div style={{ ...styles.card, padding: 22 }}>
                <div style={styles.sectionEyebrow}>Business Settings</div>
                <div style={{ fontSize: 22, fontWeight: 900, marginTop: 6, marginBottom: 4 }}>Exchange Rate</div>
                <div style={{ color: textSoft, fontSize: 13, marginBottom: 16, lineHeight: 1.5 }}>
                  Used everywhere: revenue conversion, stock value, ads spend, profit calculations. Default: 2850 TSh / USD.
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
                  <div style={styles.fieldBlock}>
                    <label style={styles.fieldLabel}>1 USD = ? TSh</label>
                    <input
                      style={{ ...styles.input, width: 160 }}
                      type="number"
                      min="100"
                      step="1"
                      value={serviceForm.exchangeRate}
                      onChange={(e) => {
                        const next = { ...serviceForm, exchangeRate: Number(e.target.value) || USD_TO_TZS };
                        setServiceForm(next);
                        persistSharedSnapshot({ ...latestSharedStateRef.current, serviceForm: next }, { successNotice: "Exchange rate saved" });
                      }}
                    />
                  </div>
                  <div style={{ color: textSoft, fontSize: 13, marginTop: 18 }}>
                    Current: 1 USD = {formatTZS(serviceForm.exchangeRate)} · 1 TSh = {formatUSD(1 / Number(serviceForm.exchangeRate || USD_TO_TZS))}
                  </div>
                  <button
                    style={{ ...styles.btnSecondary, marginTop: 18 }}
                    onClick={() => {
                      const next = { ...serviceForm, exchangeRate: 2850 };
                      setServiceForm(next);
                      persistSharedSnapshot({ ...latestSharedStateRef.current, serviceForm: next }, { successNotice: "Exchange rate reset to 2850" });
                    }}
                  >
                    Reset to 2850
                  </button>
                </div>
              </div>

              <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(4, minmax(0, 1fr))", "repeat(2, minmax(0, 1fr))", "1fr"), gap: 16 }}>
                <KpiCard icon={<LayoutGrid size={18} />} title="Audit Entries" value={auditSummary.totalEntries} sub="Saved history rows" />
                <KpiCard icon={<ClipboardList size={18} />} title="Import Events" value={auditSummary.imports} sub="Orders or shipping imports" valueColor={accent} />
                <KpiCard icon={<Users size={18} />} title="Manual Changes" value={auditSummary.manualChanges} sub="Status and order updates" valueColor={amber} />
                <KpiCard icon={<CalendarDays size={18} />} title="Last Update" value={auditSummary.latestEntryAt ? new Date(auditSummary.latestEntryAt).toLocaleString() : "N/A"} sub="Most recent recorded action" />
              </div>

              <div style={{ ...styles.card, padding: 22 }}>
                <div style={styles.sectionHeader}>
                  <div>
                    <div style={styles.sectionEyebrow}>Audit trail</div>
                    <div style={{ fontSize: 24, fontWeight: 900, marginTop: 8 }}>Everything that changed in the app</div>
                    <div style={{ color: textSoft, marginTop: 6, lineHeight: 1.6 }}>
                      Use this page to trace imports, manual edits, status transitions and owner changes per order.
                    </div>
                  </div>
                </div>

                <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("1fr 220px", "1fr", "1fr"), gap: 12, marginBottom: 16 }}>
                  <input
                    style={styles.input}
                    placeholder="Search customer, order, product, action, source..."
                    value={auditSearch}
                    onChange={(e) => setAuditSearch(e.target.value)}
                  />
                  <div style={{ ...styles.softStat, display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                    <div>
                      <div style={{ fontSize: 11, fontWeight: 800, letterSpacing: 0.45, textTransform: "uppercase", color: textSoft }}>Visible rows</div>
                      <div style={{ fontSize: 20, fontWeight: 900, marginTop: 4 }}>{filteredAuditRows.length}</div>
                    </div>
                    <div style={{ ...styles.badge, background: "rgba(35,88,213,0.08)", color: accent, border: "1px solid rgba(35,88,213,0.12)" }}>
                      Latest first
                    </div>
                  </div>
                </div>

                <div style={{ overflowX: "auto", border: `1px solid ${cardBorder}`, borderRadius: 20 }}>
                  <table style={{ width: "100%", borderCollapse: "separate", borderSpacing: 0 }}>
                    <thead>
                      <tr>
                        {["Date", "Order", "Customer", "Product", "Action", "Source", "Details"].map((head) => (
                          <th key={head} style={{ textAlign: "left", padding: "14px 12px", color: textSoft, fontSize: 12, fontWeight: 800, letterSpacing: 0.4, textTransform: "uppercase", borderBottom: `1px solid ${cardBorder}`, background: "rgba(247, 243, 237, 0.92)" }}>
                            {head}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {filteredAuditRows.slice(0, 200).map((row, index) => (
                        <tr key={`${row.customerId}-${row.at}-${index}`} style={{ background: index % 2 === 0 ? "rgba(255,255,255,0.72)" : "rgba(250,247,242,0.8)" }}>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>{row.at ? new Date(row.at).toLocaleString() : "N/A"}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 700 }}>{row.customerId}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>{row.customerName}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>{row.productName}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>{formatStatusLabel(row.action)}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>{row.source || "system"}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}`, color: textSoft }}>{row.details || "-"}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  {filteredAuditRows.length === 0 ? <div style={{ padding: 24, color: textSoft }}>No audit entry matched this search yet.</div> : null}
                </div>
              </div>
            </div>
          )}

{["alerts", "operationsHub"].includes(activePage) && (
            <div style={{ ...styles.card, padding: 22 }}>
              <div style={{ fontSize: 22, fontWeight: 800, marginBottom: 8 }}>Alerts</div>
              <div style={{ color: textSoft }}>Section reservee pour tes alertes operationnelles et marketing.</div>
            </div>
          )}
          </div>
        </main>
      </div>
      {showCloudLoginGate ? (
        <div
          style={{
            position: "fixed",
            inset: 0,
            zIndex: 80,
            display: "grid",
            placeItems: "center",
            padding: isCompact ? 18 : 32,
            background: "rgba(240, 246, 255, 0.38)",
            backdropFilter: "blur(14px)",
          }}
        >
          <div style={{ ...styles.card, width: "100%", maxWidth: 520, padding: isCompact ? 22 : 28 }}>
            <div style={styles.sectionEyebrow}>Cloud access</div>
            <div style={{ fontSize: 30, fontWeight: 900, marginTop: 10 }}>Sign in to continue</div>
            <div style={{ color: textSoft, marginTop: 8, lineHeight: 1.6 }}>
              {showCloudAuthNotice ? cloudAuth.notice : "Use your email to open the live shared workspace."}
            </div>
            <div style={{ display: "grid", gap: 12, marginTop: 22 }}>
              <input
                style={styles.input}
                type="email"
                placeholder="Email"
                value={cloudAuth.email}
                onChange={(e) => setCloudAuth((prev) => ({ ...prev, email: e.target.value }))}
              />
              <input
                style={styles.input}
                type="password"
                placeholder="Password"
                value={cloudAuth.password}
                onChange={(e) => setCloudAuth((prev) => ({ ...prev, password: e.target.value }))}
              />
              <select
                style={styles.input}
                value={cloudAuth.mode}
                onChange={(e) => setCloudAuth((prev) => ({ ...prev, mode: e.target.value }))}
              >
                <option value="signin">Sign in</option>
                <option value="signup">Create access</option>
              </select>
              <button style={{ ...styles.btnPrimary, minHeight: 56 }} onClick={submitCloudAuth} disabled={cloudAuth.loading}>
                {cloudAuth.loading ? "Connecting..." : cloudAuth.mode === "signup" ? "Create cloud access" : "Open cloud workspace"}
              </button>
            </div>
          </div>
        </div>
      ) : null}
    </div>
  );
}

function PageHeader({ eyebrow = null, title, description, filters = null, action = null }) {
  return (
    <div style={{ ...styles.card, padding: 22 }}>
      <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 16, flexWrap: "wrap" }}>
        <div style={{ minWidth: 0, flex: "1 1 420px" }}>
          {eyebrow ? <div style={styles.sectionEyebrow}>{eyebrow}</div> : null}
          <div style={{ fontSize: 26, fontWeight: 900, marginTop: eyebrow ? 8 : 0 }}>{title}</div>
          {description ? <div style={{ color: textSoft, marginTop: 8, lineHeight: 1.6, maxWidth: 760 }}>{description}</div> : null}
        </div>
        {action ? <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>{action}</div> : null}
      </div>
      {filters ? <div style={{ marginTop: 16 }}>{filters}</div> : null}
    </div>
  );
}

function InlineTabs({ items, value, onChange }) {
  return (
    <div className="tabs-row">
      {items.map((item) => (
        <button
          key={item.value}
          style={value === item.value ? styles.btnPrimary : styles.btnSecondary}
          onClick={() => onChange(item.value)}
        >
          {item.label}
        </button>
      ))}
    </div>
  );
}

function PageDateFilterBar({ title, value, onChange }) {
  const activeRange = value.preset === "custom"
    ? {
        start: value.startDate || "",
        end: value.endDate || value.startDate || "",
      }
    : getDateRangeFromPreset(value.preset || DEFAULT_PAGE_DATE_PRESET);

  return (
    <div style={{ ...styles.card, padding: 18 }}>
      <div style={{ display: "flex", alignItems: "end", justifyContent: "space-between", gap: 12, flexWrap: "wrap" }}>
        <div>
          <div style={styles.sectionEyebrow}>Date filter</div>
          <div style={{ fontSize: 20, fontWeight: 900, marginTop: 8 }}>{title}</div>
          <div style={{ color: textSoft, marginTop: 6, lineHeight: 1.5 }}>
            {activeRange.start && activeRange.end
              ? `${formatLongDate(activeRange.start)} - ${formatLongDate(activeRange.end)}`
              : "Choose a preset or a custom range."}
          </div>
        </div>
      </div>

      <div
        style={{
          display: "grid",
          gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))",
          gap: 12,
          marginTop: 16,
          alignItems: "end",
        }}
      >
        <div style={styles.fieldBlock}>
          <label style={styles.fieldLabel}>Range</label>
          <select
            style={styles.input}
            value={value.preset}
            onChange={(e) => {
              const preset = e.target.value;
              if (preset === "custom") {
                onChange((prev) => ({ ...prev, preset }));
                return;
              }
              const range = getDateRangeFromPreset(preset);
              onChange(() => ({ preset, startDate: range.start, endDate: range.end }));
            }}
          >
            {PAGE_DATE_FILTER_PRESETS.map((option) => (
              <option key={option.value} value={option.value}>
                {option.label}
              </option>
            ))}
          </select>
        </div>

        <div style={styles.fieldBlock}>
          <label style={styles.fieldLabel}>Start date</label>
          <input
            style={styles.input}
            type="date"
            value={value.startDate}
            disabled={value.preset !== "custom"}
            onChange={(e) =>
              onChange((prev) => {
                const startDate = e.target.value;
                const endDate = prev.endDate && prev.endDate < startDate ? startDate : prev.endDate;
                return { ...prev, startDate, endDate };
              })
            }
          />
        </div>

        <div style={styles.fieldBlock}>
          <label style={styles.fieldLabel}>End date</label>
          <input
            style={styles.input}
            type="date"
            value={value.endDate}
            disabled={value.preset !== "custom"}
            onChange={(e) =>
              onChange((prev) => {
                const endDate = e.target.value;
                const startDate = prev.startDate && endDate && endDate < prev.startDate ? endDate : prev.startDate;
                return { ...prev, startDate, endDate };
              })
            }
          />
        </div>
      </div>
    </div>
  );
}

/*
TEST CASES:
1. Save a product from Expedition Product -> it appears in Products and Stock.
2. Delete a product -> product and linked tracking rows disappear.
3. Add tracking row -> dashboard KPIs update.
4. Backup JSON -> file downloads with products, tracking, customers, and service form.
5. Restore JSON -> products, tracking, customers, and service form return.
6. Dashboard shows reorder alerts when available stock is low.
*/
