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
  Phone,
  MapPin,
  CalendarDays,
  ChevronLeft,
  ChevronRight,
  ChevronDown,
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
  buildHistoryEntry,
  buildMappedMetaRows,
  buildNextId,
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
  isShippingDelivered,
  isShippingInProgress,
  isShippingReturned,
  matchProductIdFromText,
  META_RANGE_PRESETS,
  normalizeHeaderName,
  normalizeOrderStatus,
  normalizePhoneValue,
  normalizeProductOffers,
  parseDateInput,
  parseLooseNumber,
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

const STORAGE_KEY = "tanzania-ecom-tracker-v16";
const AUTO_BACKUP_KEY = "tanzania-ecom-tracker-auto-backup-v1";
const AUTO_BACKUP_META_KEY = "tanzania-ecom-tracker-auto-backup-meta-v1";
const IMPORT_META_KEY = "tanzania-ecom-tracker-import-meta-v1";

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
  layout: { display: "grid", gridTemplateColumns: "236px 1fr", minHeight: "100vh", maxWidth: "100%", overflowX: "hidden" },
  sidebar: {
    background: "linear-gradient(180deg, rgba(255,255,255,0.98), rgba(246,249,252,0.96))",
    borderRight: `1px solid ${cardBorder}`,
    padding: 18,
    backdropFilter: "blur(16px)",
    position: "sticky",
    top: 0,
    alignSelf: "start",
    height: "100vh",
    overflowY: "auto",
  },
  main: { padding: 22, minWidth: 0, overflowX: "hidden" },
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
    alignItems: "stretch",
  },
  heroAside: {
    borderRadius: 20,
    padding: 18,
    border: "1px solid rgba(29, 95, 208, 0.12)",
    background: "linear-gradient(160deg, rgba(23,32,51,0.98), rgba(29,95,208,0.96))",
    color: "white",
    boxShadow: "0 18px 36px rgba(23, 32, 51, 0.16)",
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
      <span>{label}</span>
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
      <div style={{ fontSize: 22, fontWeight: 900, color: valueColor, whiteSpace: "normal", wordBreak: "keep-all", lineHeight: "1.05" }}>{value}</div>
      <div style={{ marginTop: 7, color: textSoft, fontSize: 12, lineHeight: 1.45 }}>{sub}</div>
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
        <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)", gap: 6, color: textSoft, fontSize: 12, textAlign: "center" }}>
          {["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"].map((day) => (
            <div key={day}>{day}</div>
          ))}
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)", gap: 6 }}>
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

  return (
    <div style={{ position: "relative" }}>
      <button
        type="button"
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
            position: "absolute",
            top: "calc(100% + 10px)",
            right: 0,
            zIndex: 30,
            width: "min(980px, 92vw)",
            padding: 18,
            borderRadius: 24,
            border: `1px solid ${cardBorder}`,
            background: "linear-gradient(180deg, rgba(255,255,255,0.99), rgba(247,243,237,0.97))",
            boxShadow: "0 30px 60px rgba(23,32,51,0.16)",
          }}
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
                <input style={styles.input} type="text" readOnly value={draftRange.start ? formatLongDate(draftRange.start) : ""} placeholder="Start date" />
                <div style={{ color: textSoft, fontWeight: 800, textAlign: "center" }}>-</div>
                <input style={styles.input} type="text" readOnly value={draftRange.end ? formatLongDate(draftRange.end) : ""} placeholder="End date" />
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
      ) : null}
    </div>
  );
}

export default function App() {
  const ordersImportInputRef = useRef(null);
  const shippingImportInputRef = useRef(null);
  const selectAllCustomersRef = useRef(null);
  const selectAllShippingRef = useRef(null);
  const metaAutoSyncLockRef = useRef(false);
  const metaSpendBootstrapRef = useRef("");
  const sharedSyncLockRef = useRef(false);
  const sharedHydratingRef = useRef(false);
  const sharedVersionRef = useRef(0);
  const lastSharedPayloadRef = useRef("");
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
  const [activePage, setActivePage] = useState("dashboard");
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
      setActivePage("executive");
      return;
    }
    if (activePage === "operationsHub") {
      setActivePage("taskCenter");
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
    }),
    [customers, importMeta, metaAdsState, products, serviceForm, situationData, tracking]
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
    };
  }, [customers, importMeta, metaAdsState, products, serviceForm, situationData, tracking]);

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
    } finally {
      window.setTimeout(() => {
        sharedHydratingRef.current = false;
      }, 0);
    }
  }, [readBrowserBackupSnapshot]);

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
          if (remoteHasData) {
            applySharedStateSnapshot(remoteState);
            lastSharedPayloadRef.current = JSON.stringify(remoteState);
          } else if (shouldKeepLocal) {
            lastSharedPayloadRef.current = "";
          } else if (browserBackupSnapshot) {
            applySharedStateSnapshot(browserBackupSnapshot);
            lastSharedPayloadRef.current = "";
          } else {
            applySharedStateSnapshot(remoteState);
            lastSharedPayloadRef.current = JSON.stringify(remoteState);
          }
          sharedVersionRef.current = Number(payload.version || 0);
          setSharedWorkspace({
            mode: "cloud",
            available: true,
            loading: false,
            saving: false,
            initialized: true,
            version: Number(payload.version || 0),
            updatedAt: payload.updatedAt || null,
            notice: recoveredFromBrowserBackup
              ? "Recovered cloud data from browser backup"
              : shouldKeepLocal
                ? "Cloud workspace was empty - keeping local data"
                : "Cloud workspace connected",
          });
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
    localStorage.setItem(IMPORT_META_KEY, JSON.stringify(importMeta));
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
    };
    const nextSnapshotHasData = hasMeaningfulWorkspaceData(nextSnapshot);
    const existingBrowserBackup = readLocalWorkspaceSnapshotFromStorage();
    if (!nextSnapshotHasData && existingBrowserBackup) {
      return;
    }

    const payload = {
      ...nextSnapshot,
      exportedAt: new Date().toISOString(),
      version: 1,
    };

    localStorage.setItem(STORAGE_KEY, JSON.stringify(nextSnapshot));
    localStorage.setItem(AUTO_BACKUP_KEY, JSON.stringify(payload));

    const now = new Date().toISOString();
    localStorage.setItem(AUTO_BACKUP_META_KEY, JSON.stringify({ lastAutoBackupAt: now }));
    setLastAutoBackupAt(now);
  }, [customers, importMeta, metaAdsState, products, readBrowserBackupSnapshot, serviceForm, situationData, tracking]);

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
      if (supabaseEnabled && !cloudAuth.user) return false;

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
    if (supabaseEnabled && !cloudAuth.user) return;

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

  const persistProductsSnapshot = useCallback(
    async (nextProducts, notice = "Cloud product catalog synced") => {
      const nextSnapshot = {
        ...(latestSharedStateRef.current || getDefaultCloudWorkspaceState()),
        products: nextProducts.map(sanitizeProductRecord),
      };

      latestSharedStateRef.current = nextSnapshot;
      return persistSharedSnapshot(nextSnapshot, {
        progressNotice: "Saving product changes to cloud...",
        successNotice: notice,
        failurePrefix: "Cloud product sync failed",
      });
    },
    [persistSharedSnapshot]
  );

  const getProduct = useCallback((id) => products.find((p) => p.id === id), [products]);
  const responsiveColumns = useCallback(
    (desktop, tablet = "1fr 1fr", mobile = "1fr") => {
      if (viewportWidth <= 640) return mobile;
      if (viewportWidth <= 1024) return tablet;
      return desktop;
    },
    [viewportWidth]
  );
  const isCompact = viewportWidth <= 1024;

  const customerMetricsByProduct = useMemo(() => {
    return customers.reduce((acc, customer) => {
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
          shipping: 0,
          shippingUnits: 0,
          delivered: 0,
          deliveredUnits: 0,
          returned: 0,
          returnedUnits: 0,
          cancelled: 0,
          cancelledUnits: 0,
          revenue: 0,
          statusCounts: {},
        };
      }

      acc[productId].orders += 1;
      acc[productId].orderedUnits += quantity;
      const statusKey = shippingStatus || confirmationStatus;
      acc[productId].statusCounts[statusKey] = (acc[productId].statusCounts[statusKey] || 0) + 1;

      if (isConfirmationConfirmed(confirmationStatus)) {
        acc[productId].confirmed += 1;
        acc[productId].confirmedUnits += quantity;
      }
      if (shippingStatus && isShippingInProgress(shippingStatus)) {
        acc[productId].shipping += 1;
        acc[productId].shippingUnits += quantity;
      }
      if (isShippingDelivered(shippingStatus)) {
        acc[productId].delivered += 1;
        acc[productId].deliveredUnits += quantity;
        acc[productId].revenue += getCustomerOrderTotalTzs(customer, product);
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
    }, {});
  }, [customers, getProduct]);

  const productDashboard = useMemo(() => {
    return products
      .map((product) => {
        const rows = tracking.filter((t) => t.productId === product.id);
        const customerMetrics = customerMetricsByProduct[product.id] || {
          orders: 0,
          orderedUnits: 0,
          confirmed: 0,
          confirmedUnits: 0,
          shipping: 0,
          shippingUnits: 0,
          delivered: 0,
          deliveredUnits: 0,
          returned: 0,
          returnedUnits: 0,
          cancelled: 0,
          cancelledUnits: 0,
          revenue: 0,
          statusCounts: {},
        };
        let spend = 0;

        rows.forEach((row) => {
          spend += Number(row.adSpend || 0);
        });

        const deliveredUnits = Number(customerMetrics.deliveredUnits || 0);
        const shippingUnits = Number(customerMetrics.shippingUnits || 0);
        const returnedUnits = Number(customerMetrics.returnedUnits || 0);
        const confirmedUnits = Number(customerMetrics.confirmedUnits || 0);
        const orderedUnits = Number(customerMetrics.orderedUnits || 0);
        const delivered = Number(customerMetrics.delivered || 0);
        const confirmed = Number(customerMetrics.confirmed || 0);
        const shipping = Number(customerMetrics.shipping || 0);
        const orders = Number(customerMetrics.orders || 0);
        const revenue = Number(customerMetrics.revenue || 0);
        const unitProductCost = getUnitProductCostUSD(product);
        const deliveryTzs = Number(product.delivery || 0);
        const logisticsOutflow = deliveredUnits * ((unitProductCost * USD_TO_TZS) + deliveryTzs);
        const profit = revenue - spend - logisticsOutflow;
        const cpa = deliveredUnits > 0 ? spend / deliveredUnits : 0;
        const costPerLead = orders > 0 ? spend / orders : 0;
        const roas = spend > 0 ? revenue / spend : 0;
        const confirmRate = orders > 0 ? confirmed / orders : 0;
        const deliveryRate = confirmed > 0 ? delivered / confirmed : 0;
        const margin = revenue > 0 ? (profit / revenue) * 100 : 0;
        const initialStock = Number(product.totalQty || 0);
        const reservedStock = Math.max(confirmedUnits - shippingUnits - deliveredUnits - returnedUnits, 0);
        const currentStock = Math.max(initialStock - shippingUnits - deliveredUnits, 0);
        const availableStock = Math.max(currentStock - reservedStock, 0);
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
          unitProductCost,
          totalImportCost:
            Number(product.purchaseUnitPrice || 0) * Number(product.totalQty || 0) +
            Number(product.shippingTotal || 0) +
            Number(product.otherCharges || 0),
          spend,
          orders,
          orderedUnits,
          confirmed,
          shipping,
          delivered,
          confirmedUnits,
          shippingUnits,
          deliveredUnits,
          returnedOrders: customerMetrics.returned,
          returnedUnits,
          cancelledOrders: customerMetrics.cancelled,
          cancelledUnits: customerMetrics.cancelledUnits,
          revenue,
          profit,
          cpa,
          costPerLead,
          roas,
          confirmRate,
          deliveryRate,
          margin,
          decision,
          score,
          initialStock,
          currentStock,
          reservedStock,
          availableStock,
          salesPerDay,
          reorderPoint,
          reorderSoonPoint,
          reorderStatus,
          statusCounts: customerMetrics.statusCounts,
          automatedFromOrders: true,
        };
      })
      .sort((a, b) => b.score - a.score);
  }, [customerMetricsByProduct, products, tracking]);

  const bestProduct = productDashboard[0];
  const productDashboardMap = useMemo(
    () => Object.fromEntries(productDashboard.map((product) => [product.id, product])),
    [productDashboard]
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

  const customerFormOrderValue = useMemo(
    () => Number(customerFormPricing.totalPrice || 0),
    [customerFormPricing]
  );

  const confirmationStatusCatalog = useMemo(() => {
    const seen = new Set(DEFAULT_CONFIRMATION_STATUSES);
    customers.forEach((customer) => {
      const key = normalizeOrderStatus(customer.confirmationStatus || customer.status);
      if (key) seen.add(key);
    });

    return Array.from(seen)
      .map((status) => {
        const count = customers.filter(
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
  }, [customers]);

  const confirmationStatusMap = useMemo(
    () => Object.fromEntries(confirmationStatusCatalog.map((status) => [status.value, status])),
    [confirmationStatusCatalog]
  );

  const shippingStatusCatalog = useMemo(() => {
    const seen = new Set(DEFAULT_POST_CONFIRMATION_STATUSES);
    customers.forEach((customer) => {
      const key = normalizeOrderStatus(
        customer.shippingStatus || (isConfirmationConfirmed(getCustomerConfirmationStatus(customer)) ? "to-prepare" : "")
      );
      if (key) seen.add(key);
    });

    return Array.from(seen)
      .map((status) => {
        const count = customers.filter((customer) => {
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
  }, [customers]);

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
    const confirmedPipelineCount = customers.filter((customer) => isConfirmationConfirmed(getCustomerConfirmationStatus(customer))).length;

    return {
      isVisible: confirmedPipelineCount > 0 && cutoffReached && lastShippingImportDay !== todayLabel,
      confirmedPipelineCount,
      lastShippingImportAt,
      lastShippingImportLabel: lastShippingImportAt ? new Date(lastShippingImportAt).toLocaleString() : "No shipping import yet",
    };
  }, [currentTime, customers, importMeta]);

  const selectedMetaAccount = useMemo(
    () => metaAdsAccounts.find((account) => String(account.id) === String(metaAdsState.accountId)),
    [metaAdsAccounts, metaAdsState.accountId]
  );

  const metaCampaignRows = useMemo(() => {
    return buildMappedMetaRows(metaAdsInsights?.rows || [], products, metaAdsState.campaignMappings);
  }, [metaAdsInsights, metaAdsState.campaignMappings, products]);

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
      setMetaAdsNotice(error instanceof Error ? error.message : "Unable to load Meta ad accounts.");
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

  const importMetaInsightsPayload = useCallback(
    (payload, options = {}) => {
      const mappedRows = buildMappedMetaRows(payload?.rows || [], products, metaAdsState.campaignMappings);
      const matchedRows = mappedRows.filter((row) => row.mappedProductId && Number(row.spend || 0) > 0);
      if (!matchedRows.length) {
        if (!options.silent) setMetaAdsNotice("No matched campaign row is ready to import.");
        return false;
      }

      const accountCurrency = String(selectedMetaAccount?.currency || "USD").toUpperCase();
      const convertSpendToTzs = (amount) => (accountCurrency === "TZS" ? amount : amount * USD_TO_TZS);
      const groupedByProduct = matchedRows.reduce((acc, row) => {
        const productId = row.mappedProductId;
        if (!acc[productId]) {
          acc[productId] = { spendTzs: 0, leads: 0, actualLeads: 0 };
        }
        acc[productId].spendTzs += convertSpendToTzs(Number(row.spend || 0));
        acc[productId].leads += Math.max(0, Number(row.trackedLeads ?? row.leads ?? 0));
        acc[productId].actualLeads += Math.max(0, Number(row.actualLeads ?? row.leads ?? 0));
        return acc;
      }, {});

      const importedAt = new Date().toISOString();

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
            metaCurrency: accountCurrency,
          };

          if (existingIndex >= 0) {
            next[existingIndex] = { ...next[existingIndex], ...nextPayload };
          } else {
            next.push({
              id: buildNextId(next, "T"),
              ...nextPayload,
            });
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
        lastSyncSummary: {
          since: metaAdsState.dateStart,
          until: metaAdsState.dateEnd,
          matchedProducts: Object.keys(groupedByProduct).length,
          matchedRows: matchedRows.length,
          totalSpendTzs: Object.values(groupedByProduct).reduce((sum, entry) => sum + Number(entry.spendTzs || 0), 0),
          totalLeads: Object.values(groupedByProduct).reduce((sum, entry) => sum + Number(entry.leads || 0), 0),
          totalActualLeads: Object.values(groupedByProduct).reduce((sum, entry) => sum + Number(entry.actualLeads || 0), 0),
        },
      }));

      if (!options.silent) {
        setMetaAdsNotice(`Imported Meta spend into ${Object.keys(groupedByProduct).length} product(s).`);
      }

      return true;
    },
    [metaAdsState.campaignMappings, metaAdsState.dateEnd, metaAdsState.dateStart, products, selectedMetaAccount?.currency]
  );

  const syncMetaTotalSpend = useCallback(async (options = {}) => {
    const todayBucket = getDayBucket(currentTime);
    if (!options.force && metaAdsState.lastLifetimeSpendSyncDate === todayBucket && Number(metaAdsState.lifetimeSpendTzs || 0) > 0) {
      return;
    }

    const payload = await fetchMetaSpendTotalPayload();
    const amount = Number(payload?.spend || 0);
    const totalSpendTzs = metaCurrencyIsTzs ? amount : amount * USD_TO_TZS;
    const capturedAt = payload?.capturedAt || new Date().toISOString();

    setMetaAdsState((prev) => {
      if (!options.force && prev.lastLifetimeSpendSyncDate === todayBucket && Number(prev.lifetimeSpendTzs || 0) > 0) {
        return prev;
      }

      const existing = Array.isArray(prev.dailySpendSnapshots) ? prev.dailySpendSnapshots.filter((entry) => entry.bucket !== todayBucket) : [];
      const previousReferenceTotalTzs =
        existing.length > 0 ? Number(existing[0]?.totalSpendTzs || existing[0]?.newSpendTzs || 0) : 0;
      const newSpendTzs = existing.length > 0 ? Math.max(0, totalSpendTzs - previousReferenceTotalTzs) : totalSpendTzs;
      const dailySpendSnapshots = [
        {
          id: `meta-daily-${todayBucket}`,
          bucket: todayBucket,
          totalSpendTzs,
          newSpendTzs,
          capturedAt,
          source: "meta_maximum",
        },
        ...existing,
      ].slice(0, 120);
      const cumulativeTrackedSpendTzs = dailySpendSnapshots.reduce((sum, entry) => sum + Number(entry.newSpendTzs || 0), 0);

      return {
        ...prev,
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
              lastDailySpendTzs: newSpendTzs,
            }
          : prev.lastSyncSummary,
      };
    });
  }, [currentTime, fetchMetaSpendTotalPayload, metaAdsState.lastLifetimeSpendSyncDate, metaAdsState.lifetimeSpendTzs, metaCurrencyIsTzs]);

  const refreshMetaInsights = useCallback(async (options = {}) => {
    setMetaAdsLoading((prev) => ({ ...prev, insights: !options.silent }));
    if (!options.silent) setMetaAdsNotice("");
    try {
      const payload = await fetchMetaInsightsPayload();
      setMetaAdsInsights(payload);
      if (options.applyToApp || metaAdsState.autoSync) {
        importMetaInsightsPayload(payload, { silent: true });
      }
      if (options.syncTotalSpend) {
        await syncMetaTotalSpend({ force: true });
      }
      if (!options.silent) setMetaAdsNotice(`Insights updated for ${metaAdsState.dateStart} -> ${metaAdsState.dateEnd}.`);
    } catch (error) {
      if (!options.silent) setMetaAdsNotice(error instanceof Error ? error.message : "Unable to load Meta insights.");
    } finally {
      setMetaAdsLoading((prev) => ({ ...prev, insights: false }));
    }
  }, [fetchMetaInsightsPayload, importMetaInsightsPayload, metaAdsState.autoSync, metaAdsState.dateEnd, metaAdsState.dateStart, syncMetaTotalSpend]);

  const updateMetaCampaignMapping = (campaignId, productId) => {
    setMetaAdsState((prev) => ({
      ...prev,
      campaignMappings: {
        ...prev.campaignMappings,
        [campaignId]: productId,
      },
    }));
  };

  const applyMetaInsightsToApp = () => {
    setMetaAdsLoading((prev) => ({ ...prev, apply: true }));
    try {
      importMetaInsightsPayload(metaAdsInsights, { silent: false });
    } finally {
      setMetaAdsLoading((prev) => ({ ...prev, apply: false }));
    }
  };

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
    refreshMetaInsights,
    syncMetaTotalSpend,
  ]);

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
    }

    setProducts(nextProducts);
    latestSharedStateRef.current = {
      ...(latestSharedStateRef.current || getDefaultCloudWorkspaceState()),
      products: nextProducts.map(sanitizeProductRecord),
    };
    await persistProductsSnapshot(nextProducts, editingProductId ? "Cloud product updated" : "Cloud product added");

    setEditingProductId(null);
    setExpeditionForm(getEmptyExpeditionForm());
    setActivePage("products");
  };

  const startEditingProduct = (product) => {
    setEditingProductId(product.id);
    setExpeditionForm({
      name: product.name || "",
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

    setProducts(nextProducts);
    setTracking(nextTracking);
    setCustomers(nextCustomers);

    const nextSnapshot = {
      ...(latestSharedStateRef.current || getDefaultCloudWorkspaceState()),
      products: nextProducts.map(sanitizeProductRecord),
      tracking: nextTracking,
      customers: nextCustomers.map(sanitizeCustomerRecord),
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

  const getExcelCellValue = useCallback((row, aliases) => {
    const entries = Object.entries(row || {});
    for (const [key, value] of entries) {
      const normalizedKey = normalizeHeaderName(key);
      if (aliases.includes(normalizedKey)) return value;
    }
    return "";
  }, []);

  const resolveImportedProductId = useCallback(
    (rawValue) => matchProductIdFromText(rawValue, products),
    [products]
  );

  const findMatchingCustomerIndex = useCallback((customerList, payload) => {
    const normalizedSourceOrderId = String(payload.sourceOrderId || "").trim();
    const normalizedPhone = normalizePhoneValue(payload.phone);
    const normalizedName = normalizeHeaderName(payload.customerName);
    const normalizedOrderDate = String(payload.orderDate || "").trim();
    const normalizedQuantity = payload.quantity ? Number(payload.quantity) : null;

    return customerList.findIndex((customer) => {
      if (
        normalizedSourceOrderId &&
        customer.sourceOrderId &&
        String(customer.sourceOrderId).trim() === normalizedSourceOrderId
      ) {
        return true;
      }

      const customerPhone = normalizePhoneValue(customer.phone);
      const samePhone = normalizedPhone && customerPhone === normalizedPhone;
      const sameProduct = payload.productId ? customer.productId === payload.productId : true;
      const sameDate = normalizedOrderDate ? String(customer.orderDate || "") === normalizedOrderDate : true;
      const sameQuantity = normalizedQuantity ? Number(customer.quantity || 1) === normalizedQuantity : true;
      const sameName = normalizedName
        ? normalizeHeaderName(customer.customerName) === normalizedName
        : true;

      if (samePhone && sameProduct && sameDate && sameQuantity) return true;
      if (samePhone && sameName && sameProduct) return true;
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
        let createdCount = 0;
        let updatedCount = 0;
        let skippedCount = 0;
        const reasonCounts = {
          missingName: 0,
          missingPhone: 0,
          missingProduct: 0,
          unknownProduct: 0,
        };
        const unmatchedProducts = new Set();
        const detectedHeaders = Object.keys(rows[0] || {});

        rows.forEach((row) => {
          const sourceOrderId = String(
            getExcelCellValue(row, [
              "code",
              "order id",
              "commande id",
              "id commande",
              "external order id",
              "shipbox order id",
              "reference",
              "reference commande",
              "numero commande",
              "numero de commande",
              "order number",
              "tracking number",
            ])
          ).trim();
          const customerName = String(
            getExcelCellValue(row, [
              "customer",
              "customer name",
              "customer full name",
              "full name",
              "fullname",
              "full_name",
              "client",
              "nom client",
              "name",
              "nom",
              "receiver",
              "receiver name",
              "recipient",
              "recipient name",
              "consignee",
            ])
          ).trim();
          const phone = String(
            getExcelCellValue(row, [
              "phone",
              "phone number",
              "telephone",
              "telephone client",
              "mobile",
              "mobile number",
              "customer phone",
              "receiver phone",
              "recipient phone",
              "tel",
            ])
          ).trim();
          const rawProduct = getExcelCellValue(row, [
            "product",
            "product name",
            "product title",
            "produit",
            "item",
            "item name",
            "article",
            "designation",
            "sku",
            "product id",
          ]);
          const productId = resolveImportedProductId(rawProduct);

          if (!customerName) {
            reasonCounts.missingName += 1;
            skippedCount += 1;
            return;
          }
          if (!phone) {
            reasonCounts.missingPhone += 1;
            skippedCount += 1;
            return;
          }
          if (!String(rawProduct || "").trim()) {
            reasonCounts.missingProduct += 1;
            skippedCount += 1;
            return;
          }
          if (!productId) {
            reasonCounts.unknownProduct += 1;
            unmatchedProducts.add(String(rawProduct || "").trim());
            skippedCount += 1;
            return;
          }

          const quantity = Math.max(
            1,
            Number(
              getExcelCellValue(row, ["quantity", "qty", "quantite", "quantite commande", "quantity ordered", "product qt"]) || 1
            )
          );
          const orderTotalTzs = Math.max(
            0,
            parseLooseNumber(
              getExcelCellValue(row, [
                "total price",
                "total",
                "amount",
                "amount tzs",
                "price total",
                "total amount",
                "montant",
                "montant total",
              ])
            )
          );
          const orderDate = excelDateToInput(
            getExcelCellValue(row, ["order date", "date", "date commande", "created at", "order created at"])
          );
          const rawStatus = getExcelCellValue(row, [
            "conf status",
            "conf.status",
            "status",
            "lead status",
            "lead state",
            "confirmation",
            "confirmation status",
            "confirmation state",
            "call center status",
            "callcenter status",
            "customer status",
            "order status",
            "order state",
            "state",
            "situation",
            "situation lead",
            "etat",
            "etat commande",
            "etat confirmation",
            "status confirmation",
            "status de confirmation",
          ]);
          const status = String(rawStatus || "").trim()
            ? normalizeOrderStatus(rawStatus)
            : "new-order";
          const rawShippingStatus = getExcelCellValue(row, [
            "shipping status",
            "shipping state",
            "shipping stage",
            "shipment status",
            "delivery status",
            "delivery state",
            "shipping",
          ]);
          const shippingStatus = String(rawShippingStatus || "").trim()
            ? normalizeOrderStatus(rawShippingStatus)
            : "";
          const paymentMethod =
            String(
              getExcelCellValue(row, ["payment", "payment method", "paiement", "methode paiement", "payment mode"])
            ).trim() || "COD";
          const city = String(getExcelCellValue(row, ["city", "ville", "destination city"])).trim();
          const address = String(getExcelCellValue(row, ["address", "adresse", "location", "delivery address"])).trim();
          const notes = String(getExcelCellValue(row, ["notes", "note", "comment", "comments", "remark"])).trim();

          const existingIndex = findMatchingCustomerIndex(nextCustomers, {
            sourceOrderId,
            customerName,
            phone,
            productId,
            quantity,
            orderDate,
          });

          if (existingIndex >= 0) {
            const existing = nextCustomers[existingIndex];
            const normalizedExistingStatus = normalizeOrderStatus(existing.confirmationStatus || existing.status);
            const statusChanged = normalizedExistingStatus !== status;
            const existingShippingStatus = normalizeOrderStatus(existing.shippingStatus);
            const nextShippingStatus = ensureShippingStatusForConfirmed(
              statusChanged ? status : existing.confirmationStatus || normalizedExistingStatus,
              shippingStatus || existing.shippingStatus
            );
            const shippingStatusChanged = existingShippingStatus !== nextShippingStatus;
            const historyNotes = [];
            if (statusChanged) historyNotes.push(`Confirmation ${formatStatusLabel(normalizedExistingStatus)} -> ${formatStatusLabel(status)}`);
            if (shippingStatusChanged) historyNotes.push(`Shipping ${formatStatusLabel(existingShippingStatus || "to-prepare")} -> ${formatStatusLabel(nextShippingStatus)}`);
            if (orderTotalTzs) historyNotes.push(`Order value synced to ${formatTZS(orderTotalTzs)}`);

            nextCustomers[existingIndex] = sanitizeCustomerRecord({
              ...existing,
              status: statusChanged ? status : normalizedExistingStatus,
              confirmationStatus: statusChanged ? status : existing.confirmationStatus || normalizedExistingStatus,
              shippingStatus: nextShippingStatus,
              paymentMethod: paymentMethod || existing.paymentMethod,
              city: city || existing.city,
              address: address || existing.address,
              notes: notes || existing.notes,
              orderTotalTzs: orderTotalTzs || Number(existing.orderTotalTzs || 0),
              sourceOrderId: existing.sourceOrderId || sourceOrderId || null,
              importSource: "excel",
              lastImportedAt: importFinishedAt,
              lastShippingImportedAt: shippingStatusChanged ? importFinishedAt : existing.lastShippingImportedAt || null,
              history:
                statusChanged || shippingStatusChanged || orderTotalTzs
                  ? appendCustomerHistory(
                      existing,
                      buildHistoryEntry({
                        action: "orders_import_updated",
                        source: "excel-orders",
                        details: historyNotes.join(" | ") || "Order synced from Excel",
                      })
                    )
                  : existing.history,
            });

            if (statusChanged || shippingStatusChanged || sourceOrderId) updatedCount += 1;
            return;
          }

          nextCustomers.unshift(sanitizeCustomerRecord({
            id: buildNextId(nextCustomers, "C"),
            customerName,
            phone,
            city,
            address,
            productId,
            quantity,
            orderDate,
            paymentMethod,
            status,
            confirmationStatus: status,
            shippingStatus: ensureShippingStatusForConfirmed(status, shippingStatus),
            orderTotalTzs,
            notes,
            sourceOrderId: sourceOrderId || null,
            importSource: "excel",
            lastImportedAt: importFinishedAt,
            lastShippingImportedAt: shippingStatus ? importFinishedAt : null,
            assignedTo: "Call Center",
            history: [
              buildHistoryEntry({
                action: "orders_import_created",
                source: "excel-orders",
                details: `Imported with ${formatStatusLabel(status)}${shippingStatus ? ` | shipping ${formatStatusLabel(shippingStatus)}` : ""}`,
              }),
            ],
          }));
          createdCount += 1;
        });

        setCustomers(nextCustomers);
        setImportMeta((prev) => ({ ...prev, lastOrdersImportAt: importFinishedAt }));
        setOrdersImportNotice(`Excel imported: ${createdCount} new, ${updatedCount} updated, ${skippedCount} skipped.`);
        setOrdersImportDetails({
          detectedHeaders,
          reasonCounts,
          unmatchedProducts: Array.from(unmatchedProducts).slice(0, 6),
        });
      } catch (error) {
        const message = error instanceof Error ? error.message : "Excel import failed";
        setOrdersImportNotice(`Excel import failed: ${message}`);
        setOrdersImportDetails(null);
      } finally {
        event.target.value = "";
      }
    },
    [customers, findMatchingCustomerIndex, getExcelCellValue, resolveImportedProductId]
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
        let updatedCount = 0;
        let unchangedCount = 0;
        let skippedCount = 0;
        const reasonCounts = {
          missingStatus: 0,
          unmatchedOrder: 0,
        };
        const unmatchedExamples = new Set();
        const detectedHeaders = Object.keys(rows[0] || {});

        rows.forEach((row) => {
          const sourceOrderId = String(
            getExcelCellValue(row, [
              "code",
              "order id",
              "commande id",
              "id commande",
              "external order id",
              "shipbox order id",
              "reference",
              "reference commande",
              "numero commande",
              "numero de commande",
              "order number",
              "tracking number",
            ])
          ).trim();
          const customerName = String(
            getExcelCellValue(row, [
              "customer",
              "customer name",
              "customer full name",
              "full name",
              "fullname",
              "full_name",
              "client",
              "nom client",
              "name",
              "nom",
              "receiver",
              "receiver name",
              "recipient",
              "recipient name",
              "consignee",
            ])
          ).trim();
          const phone = String(
            getExcelCellValue(row, [
              "phone",
              "phone number",
              "telephone",
              "telephone client",
              "mobile",
              "mobile number",
              "customer phone",
              "receiver phone",
              "recipient phone",
              "tel",
            ])
          ).trim();
          const rawProduct = getExcelCellValue(row, [
            "product",
            "product name",
            "product title",
            "produit",
            "item",
            "item name",
            "article",
            "designation",
            "sku",
            "product id",
          ]);
          const productId = String(rawProduct || "").trim() ? resolveImportedProductId(rawProduct) : "";
          const quantityRaw = getExcelCellValue(row, ["quantity", "qty", "quantite", "quantite commande", "quantity ordered", "product qt"]);
          const quantity = quantityRaw === "" || quantityRaw == null ? null : Math.max(1, Number(quantityRaw || 1));
          const orderDateRaw = getExcelCellValue(row, ["order date", "date", "date commande", "created at", "order created at"]);
          const orderDate = orderDateRaw === "" || orderDateRaw == null ? "" : excelDateToInput(orderDateRaw);
          const shippingStatusRaw = getExcelCellValue(row, [
            "shipping status",
            "delivery status",
            "shipment status",
            "shipping state",
            "shipping stage",
            "status livraison",
            "delivery state",
            "status",
            "etat",
            "etat commande",
            "order status",
          ]);
          const nextStatus = normalizeOrderStatus(shippingStatusRaw);

          if (!String(shippingStatusRaw || "").trim()) {
            reasonCounts.missingStatus += 1;
            skippedCount += 1;
            return;
          }

          const existingIndex = findMatchingCustomerIndex(nextCustomers, {
            sourceOrderId,
            customerName,
            phone,
            productId,
            quantity,
            orderDate,
          });

          if (existingIndex < 0) {
            reasonCounts.unmatchedOrder += 1;
            skippedCount += 1;
            unmatchedExamples.add(sourceOrderId || phone || customerName || String(rawProduct || "").trim() || "Unknown row");
            return;
          }

          const existing = nextCustomers[existingIndex];
          const currentStatus = normalizeOrderStatus(getCustomerShippingStatus(existing));

          if (currentStatus === nextStatus) {
            unchangedCount += 1;
            return;
          }

          nextCustomers[existingIndex] = sanitizeCustomerRecord({
            ...existing,
            shippingStatus: nextStatus,
            status: nextStatus,
            confirmationStatus: isConfirmationConfirmed(existing.confirmationStatus) ? existing.confirmationStatus : "confirmed",
            sourceOrderId: existing.sourceOrderId || sourceOrderId || null,
            lastShippingImportedAt: importFinishedAt,
            actualDeliveryDate: isShippingDelivered(nextStatus) ? existing.actualDeliveryDate || getTodayString() : existing.actualDeliveryDate || "",
            importSource: existing.importSource || "excel",
            assignedTo: existing.assignedTo || "Shipping Team",
            history: appendCustomerHistory(
              existing,
              buildHistoryEntry({
                action: "shipping_import_updated",
                source: "excel-shipping",
                details: `Shipping ${formatStatusLabel(currentStatus)} -> ${formatStatusLabel(nextStatus)}`,
              })
            ),
          });
          updatedCount += 1;
        });

        setCustomers(nextCustomers);
        setImportMeta((prev) => ({ ...prev, lastShippingImportAt: importFinishedAt }));
        setShippingImportNotice(
          `Shipping Excel imported: ${updatedCount} updated, ${unchangedCount} unchanged, ${skippedCount} skipped.`
        );
        setShippingImportDetails({
          detectedHeaders,
          reasonCounts,
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
    [customers, findMatchingCustomerIndex, getExcelCellValue, resolveImportedProductId]
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
  };

  const updateCustomerShippingStatus = (customerId, nextStatus) => {
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
    setProducts((prev) =>
      prev.map((p) =>
        p.id === productId
          ? {
              ...p,
              stockArrivalStatus: "arrived",
              stockArrivedAt: getTodayString(),
              nextArrivalCheckDate: null,
            }
          : p
      )
    );
  };

  const markDubaiStockNotYet = (productId) => {
    setProducts((prev) =>
      prev.map((p) =>
        p.id === productId
          ? {
              ...p,
              stockArrivalStatus: "pending",
              nextArrivalCheckDate: addDaysToDateString(getTodayString(), 1),
            }
          : p
      )
    );
  };

  const exportReport = () => {
    const reportLines = [
      "Tanzania Ecom Tracker Report",
      `Generated at: ${new Date().toLocaleString()}`,
      "",
      "Summary",
      `- Total products: ${products.length}`,
      `- Total tracking rows: ${tracking.length}`,
      `- Total customer orders: ${customersDashboard.totalOrders}`,
      `- Confirmed orders: ${customersDashboard.confirmedOrders}`,
      `- Delivered orders: ${customersDashboard.deliveredOrders}`,
      `- Revenue: ${formatTZS(customersDashboard.totalRevenue)}`,
      "",
      "Products",
      ...productDashboard.map((product) =>
        `- ${product.name} | stock=${product.availableStock} | delivered=${product.deliveredUnits} units | profit=${formatTZS(product.profit)} | decision=${product.decision}`
      ),
      "",
      "Recent Orders",
      ...customers.slice(0, 10).map((customer) => {
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
    const totalLeads = customers.length;
    const confirmedLeads = customers.filter((c) => isConfirmationConfirmed(getCustomerConfirmationStatus(c))).length;
    const deliveredLeads = customers.filter((c) => isShippingDelivered(getCustomerShippingStatus(c))).length;
    const newLeads = customers.filter((c) => isConfirmationNew(getCustomerConfirmationStatus(c))).length;
    const cancelledLeads = customers.filter((c) => isConfirmationCancelled(getCustomerConfirmationStatus(c))).length;
    const otherLeads = totalLeads - confirmedLeads - cancelledLeads - newLeads;

    const totalRevenue = customers
      .filter((c) => isShippingDelivered(getCustomerShippingStatus(c)))
      .reduce((sum, c) => {
        const product = products.find((p) => p.id === c.productId);
        return sum + getCustomerOrderTotalTzs(c, product);
      }, 0);

    const confirmationRate = totalLeads > 0 ? (confirmedLeads / totalLeads) * 100 : 0;
    const deliveryRate = confirmedLeads > 0 ? (deliveredLeads / confirmedLeads) * 100 : 0;

    return {
      totalOrders: totalLeads,
      totalQty: customers.reduce((sum, c) => sum + Number(c.quantity || 0), 0),
      confirmedOrders: confirmedLeads,
      deliveredOrders: deliveredLeads,
      newOrders: newLeads,
      cancelledOrders: cancelledLeads,
      otherOrders: Math.max(0, otherLeads),
      totalRevenue,
      confirmationRate,
      deliveryRate,
    };
  }, [customers, products]);

  const liveAutomationSummary = useMemo(() => {
    const totalAdSpendTzs = tracking.reduce((sum, row) => sum + Number(row.adSpend || 0), 0);
    const deliveredUnits = productDashboard.reduce((sum, product) => sum + Number(product.deliveredUnits || 0), 0);
    const reservedUnits = productDashboard.reduce((sum, product) => sum + Number(product.reservedStock || 0), 0);
    const availableUnits = productDashboard.reduce((sum, product) => sum + Number(product.availableStock || 0), 0);
    const importCostDeliveredTzs = productDashboard.reduce(
      (sum, product) => sum + (Number(product.deliveredUnits || 0) * Number(product.unitProductCost || 0) * USD_TO_TZS),
      0
    );
    const localDeliveryCostTzs = productDashboard.reduce(
      (sum, product) => sum + (Number(product.deliveredUnits || 0) * Number(product.delivery || 0)),
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

  const liveServiceDataset = useMemo(() => {
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

    return customers
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
  }, [confirmationStatusMap, customerListFilters.status, customers, deferredCustomerSearch, products]);

  const customerListPageCount = Math.max(1, Math.ceil(compactCustomerRows.length / Number(customerListFilters.pageSize || 25)));
  const selectedCustomerIdSet = useMemo(() => new Set(selectedCustomerIds), [selectedCustomerIds]);
  const filteredCustomerIds = useMemo(() => compactCustomerRows.map((customer) => customer.id), [compactCustomerRows]);
  const allFilteredSelected = filteredCustomerIds.length > 0 && filteredCustomerIds.every((id) => selectedCustomerIdSet.has(id));
  const someFilteredSelected = filteredCustomerIds.some((id) => selectedCustomerIdSet.has(id)) && !allFilteredSelected;
  const historyTargetCustomer = useMemo(
    () => customers.find((customer) => customer.id === customerHistoryTargetId) || null,
    [customerHistoryTargetId, customers]
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

    return customers
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
  }, [customers, deferredShippingSearch, products, shippingListFilters.status, shippingStatusMap]);

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

    customers.forEach((customer) => {
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
        grouped[key].localDeliveryTzs += Number(product.delivery || 0) * qty;
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
  }, [customers, getProduct, productDashboardMap]);

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
    customers.forEach((customer) => {
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
  }, [customers]);

  const executiveSummary = useMemo(() => {
    const today = getTodayString();
    const monthStart = formatDateInput(new Date(new Date().getFullYear(), new Date().getMonth(), 1));
    const todayOrders = customers.filter((customer) => customer.orderDate === today);
    const todayRevenueTzs = todayOrders.reduce((sum, customer) => {
      if (!isShippingDelivered(getCustomerShippingStatus(customer))) return sum;
      return sum + getCustomerOrderTotalTzs(customer, getProduct(customer.productId));
    }, 0);
    const monthRevenueTzs = customers.reduce((sum, customer) => {
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
  }, [customers, getProduct, liveAutomationSummary.grossProfitTzs, products, situationsSummary.fixedChargesTzs, taskCenterData]);

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

  const stockValueSummary = useMemo(() => {
    const today = parseDateInput(getTodayString()) || new Date();
    return products.reduce(
      (acc, product) => {
        const dashboardRow = productDashboardMap[product.id];
        const availableStock = Number(dashboardRow?.availableStock ?? product.totalQty ?? 0);
        const stockValueTzs = availableStock * getUnitProductCostUSD(product) * USD_TO_TZS;
        const orderedAt = parseDateInput(product.stockOrderedAt);
        const ageDays = orderedAt && !Number.isNaN(orderedAt.getTime()) ? Math.max(0, Math.round((today - orderedAt) / 86400000)) : 0;
        acc.totalValueTzs += stockValueTzs;
        if (ageDays >= 60) acc.aged60Products += 1;
        if (ageDays >= 90) acc.aged90Products += 1;
        return acc;
      },
      { totalValueTzs: 0, aged60Products: 0, aged90Products: 0 }
    );
  }, [productDashboardMap, products]);

  const profitCenterRows = useMemo(() => {
    return productDashboard
      .map((product) => {
        const adInput = situationData.adInputs?.[product.id] || {};
        const manualAdsUsedTzs = Number(adInput.averageLeadCostTzs || 0) * Number(adInput.incomingLeads || 0);
        const liveObservedAdsTzs = Number(product.spend || 0);
        const cumulativeAdsTzs = manualAdsUsedTzs;
        const deliveredLogisticsTzs = Number(product.revenue || 0) - Number(product.profit || 0) - liveObservedAdsTzs;
        const cumulativeProfitTzs = Number(product.revenue || 0) - cumulativeAdsTzs - deliveredLogisticsTzs;
        const fixedProductChargesTzs =
          (Number(product.purchaseUnitPrice || 0) * Number(product.totalQty || 0) * USD_TO_TZS) +
          Number(product.shippingTotal || 0) +
          Number(product.otherCharges || 0);
        const netAfterFixedTzs = cumulativeProfitTzs - fixedProductChargesTzs;
        const deliveredCount = Number(product.delivered || product.deliveredUnits || 0);
        return {
          ...product,
          manualAdsUsedTzs,
          liveObservedAdsTzs,
          cumulativeAdsTzs,
          deliveredLogisticsTzs,
          cumulativeProfitTzs,
          fixedProductChargesTzs,
          netAfterFixedTzs,
          profitPerOrderTzs: deliveredCount > 0 ? cumulativeProfitTzs / deliveredCount : 0,
          marginPercentLive: Number(product.revenue || 0) > 0 ? (cumulativeProfitTzs / Number(product.revenue || 0)) * 100 : 0,
        };
      })
      .sort((a, b) => Number(b.cumulativeProfitTzs || 0) - Number(a.cumulativeProfitTzs || 0));
  }, [productDashboard, situationData.adInputs]);

  const auditRows = useMemo(() => {
    return customers
      .flatMap((customer) =>
        (customer.history || []).map((entry) => ({
          ...entry,
          customerId: customer.id,
          customerName: customer.customerName,
          productName: getProduct(customer.productId)?.name || customer.productId,
        }))
      )
      .sort((a, b) => String(b.at || "").localeCompare(String(a.at || "")));
  }, [customers, getProduct]);

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
    customers.forEach((customer) => {
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
        const localDeliveryTzs = Number(product?.delivery || 0) * quantity;
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
  }, [customers, getProduct, situationData.salaries]);

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

  const profitCenterSummary = useMemo(() => {
    const totals = profitCenterRows.reduce(
      (acc, row) => {
        acc.revenueTzs += Number(row.revenue || 0);
        acc.deliveredCostTzs += Number(row.deliveredLogisticsTzs || 0);
        acc.productTrackedProfitTzs += Number(row.cumulativeProfitTzs || 0);
        acc.productTrackedAdsTzs += Number(row.cumulativeAdsTzs || 0);
        acc.liveObservedAdsTzs += Number(row.liveObservedAdsTzs || 0);
        acc.fixedChargesTzs += Number(row.fixedProductChargesTzs || 0);
        return acc;
      },
      { revenueTzs: 0, deliveredCostTzs: 0, productTrackedProfitTzs: 0, productTrackedAdsTzs: 0, liveObservedAdsTzs: 0, fixedChargesTzs: 0 }
    );

    const adsSpendTzs =
      Number(metaAdsState.cumulativeTrackedSpendTzs || 0) > 0
        ? Number(metaAdsState.cumulativeTrackedSpendTzs || 0)
        : Number(totals.productTrackedAdsTzs || 0);
    const profitTzs = Number(totals.revenueTzs || 0) - Number(totals.deliveredCostTzs || 0) - adsSpendTzs;
    const netAfterFixedTzs = profitTzs - Number(totals.fixedChargesTzs || 0);

    return {
      ...totals,
      adsSpendTzs,
      profitTzs,
      netAfterFixedTzs,
      profitableProducts: profitCenterRows.filter((row) => Number(row.cumulativeProfitTzs || 0) > 0).length,
      topProduct: profitCenterRows[0] || null,
      lastHourlyAdsSnapshot: metaAdsState.dailySpendSnapshots?.[0] || null,
    };
  }, [metaAdsState.cumulativeTrackedSpendTzs, metaAdsState.dailySpendSnapshots, profitCenterRows]);

  const auditSummary = useMemo(() => {
    return {
      totalEntries: auditRows.length,
      imports: auditRows.filter((row) => String(row.action || "").includes("import")).length,
      manualChanges: auditRows.filter((row) => String(row.source || "").includes("manual")).length,
      latestEntryAt: auditRows[0]?.at || null,
    };
  }, [auditRows]);

  const addSituationSalary = () => {
    setSituationData((prev) => ({
      ...prev,
      salaries: [...prev.salaries, { id: `salary-${Date.now()}`, name: "", role: "", amountTzs: 0 }],
    }));
  };

  const updateSituationSalary = (salaryId, field, value) => {
    setSituationData((prev) => ({
      ...prev,
      salaries: prev.salaries.map((entry) =>
        entry.id === salaryId
          ? { ...entry, [field]: field === "amountTzs" ? Math.max(0, parseLooseNumber(value) * USD_TO_TZS) : value }
          : entry
      ),
    }));
  };

  const removeSituationSalary = (salaryId) => {
    setSituationData((prev) => ({
      ...prev,
      salaries: prev.salaries.filter((entry) => entry.id !== salaryId),
    }));
  };

  const addSituationFixedCharge = () => {
    setSituationData((prev) => ({
      ...prev,
      fixedCharges: [...prev.fixedCharges, { id: `fixed-${Date.now()}`, label: "", amountTzs: 0 }],
    }));
  };

  const updateSituationFixedCharge = (chargeId, field, value) => {
    setSituationData((prev) => ({
      ...prev,
      fixedCharges: prev.fixedCharges.map((entry) =>
        entry.id === chargeId
          ? { ...entry, [field]: field === "amountTzs" ? Math.max(0, parseLooseNumber(value) * USD_TO_TZS) : value }
          : entry
      ),
    }));
  };

  const removeSituationFixedCharge = (chargeId) => {
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
    setSelectedCustomerIds((prev) => prev.filter((id) => customers.some((customer) => customer.id === id)));
  }, [customers]);

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
    setSelectedShippingIds((prev) => prev.filter((id) => customers.some((customer) => customer.id === id)));
  }, [customers]);

  useEffect(() => {
    if (selectAllShippingRef.current) {
      selectAllShippingRef.current.indeterminate = someFilteredShippingSelected;
    }
  }, [someFilteredShippingSelected]);

  const ordersChartData = useMemo(() => {
    const grouped = customers.reduce((acc, customer) => {
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
  }, [customers]);

  const filteredCustomersForOverview = useMemo(() => {
    return customers.filter((customer) => {
      const matchesProduct = overviewFilters.productId === "all" || customer.productId === overviewFilters.productId;

      let matchesDate = true;
      if (overviewFilters.periodMode === "custom") {
        if (overviewFilters.startDate && customer.orderDate < overviewFilters.startDate) matchesDate = false;
        if (overviewFilters.endDate && customer.orderDate > overviewFilters.endDate) matchesDate = false;
      }

      return matchesProduct && matchesDate;
    });
  }, [customers, overviewFilters]);

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
    const confirmedCustomers = customers.filter((c) => isConfirmationConfirmed(getCustomerConfirmationStatus(c)));
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
  }, [customers, shippingStatusMap]);

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

    return customers.filter((customer) => {
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
  }, [customers, confirmationSummaryFilters]);

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

    return customers.filter((customer) => {
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
  }, [customers, productDetailsFilters]);

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

          <SidebarItem active={activePage === "dashboard"} onClick={() => setActivePage("dashboard")} icon={<BarChart3 size={18} />} label="Dashboard" />
          <SidebarItem active={activePage === "customersOrders"} onClick={() => setActivePage("customersOrders")} icon={<Users size={18} />} label="Leads" />
          <SidebarItem active={activePage === "shipping"} onClick={() => setActivePage("shipping")} icon={<ShoppingBag size={18} />} label="Shipping" />
          <SidebarItem active={["products", "stock"].includes(activePage)} onClick={() => setActivePage("products")} icon={<Archive size={18} />} label="Stock" />
          <SidebarItem active={activePage === "multiDashboard"} onClick={() => setActivePage("multiDashboard")} icon={<Boxes size={18} />} label="Fichier" />
          <SidebarItem active={activePage === "tracking"} onClick={() => setActivePage("tracking")} icon={<Calculator size={18} />} label="Finance" />
          <SidebarItem active={activePage === "serviceSum"} onClick={() => setActivePage("serviceSum")} icon={<Calculator size={18} />} label="Simulation" />
          <SidebarItem active={activePage === "situations"} onClick={() => setActivePage("situations")} icon={<Calculator size={18} />} label="Rentabilité" />
          <SidebarItem active={activePage === "profitCenter"} onClick={() => setActivePage("profitCenter")} icon={<Wallet size={18} />} label="Profit Center" />
          <SidebarItem active={["executive", "scaling"].includes(activePage)} onClick={() => setActivePage("executive")} icon={<Rocket size={18} />} label="Performance" />
          <SidebarItem active={["taskCenter", "calendar", "team", "audit", "alerts"].includes(activePage)} onClick={() => setActivePage("taskCenter")} icon={<ClipboardList size={18} />} label="Operations" />
          <SidebarItem active={activePage === "audit"} onClick={() => setActivePage("audit")} icon={<ClipboardList size={18} />} label="Audit" />

          <div style={{ ...styles.card, marginTop: 28, padding: 18, background: "linear-gradient(180deg, rgba(255,255,255,0.98), rgba(240,247,255,0.8))" }}>
            <div style={{ ...styles.sectionEyebrow, color: textSoft }}>Top performer</div>
            <div style={{ marginTop: 10, fontWeight: 900, fontSize: 18, lineHeight: 1.2 }}>{bestProduct?.name || "N/A"}</div>
            <div style={{ marginTop: 8, color: green, fontWeight: 800 }}>{bestProduct ? formatTZS(bestProduct.profit) : "N/A"}</div>
            <div style={{ marginTop: 14, display: "grid", gap: 10 }}>
              <div style={styles.softStat}>
                <div style={{ fontSize: 11, fontWeight: 800, letterSpacing: 0.45, textTransform: "uppercase", color: textSoft }}>Product score</div>
                <div style={{ marginTop: 6, fontSize: 22, fontWeight: 900 }}>{bestProduct?.score ?? 0}/100</div>
              </div>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12 }}>
                <div style={{ fontSize: 13, color: textSoft }}>Decision</div>
                <div style={getDecisionStyle(bestProduct?.decision || "WATCH")}>{bestProduct?.decision || "WATCH"}</div>
              </div>
            </div>
          </div>
        </aside>

        <main style={{ ...styles.main, padding: isCompact ? 18 : 24 }}>
          <input ref={ordersImportInputRef} type="file" accept=".xlsx,.xls,.csv" onChange={importOrdersFromExcel} style={{ display: "none" }} />
          <div style={{ width: "100%", maxWidth: 1560, margin: "0 auto", display: "grid", gap: 18 }}>
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
                {supabaseEnabled ? (
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
                <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(3, minmax(0, 1fr))", "1fr 1fr", "1fr"), gap: 12, marginTop: 18 }}>
                  <MiniStat label="Confirmation rate" value={`${Math.round(customersDashboard.confirmationRate)}%`} sub={`${customersDashboard.confirmedOrders} confirmed leads`} />
                  <MiniStat label="Delivery rate" value={`${Math.round(customersDashboard.deliveryRate)}%`} tone="green" sub={`${customersDashboard.deliveredOrders} delivered orders`} />
                  <MiniStat label="Catalog health" value={`${products.length}`} tone="amber" sub={`${liveAutomationSummary.availableUnits} units available | ${liveAutomationSummary.reservedUnits} reserved`} />
                </div>
                <div style={{ ...styles.topbarActions, marginTop: 18 }}>
                  <button style={styles.btnPrimary} onClick={exportReport}>Export Report</button>
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

          {shippingImportReminder.isVisible ? (
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

              <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(4, minmax(0, 1fr))", "repeat(2, minmax(0, 1fr))", "1fr"), gap: 16, marginBottom: 24 }}>
                <KpiCard
                  icon={<Users size={18} />}
                  title="Total Leads"
                  value={customersDashboard.totalOrders}
                  sub="All incoming orders"
                  valueColor="#94a3b8"
                />
                <KpiCard
                  icon={<Phone size={18} />}
                  title="Confirmed Leads"
                  value={customersDashboard.confirmedOrders}
                  sub="Active pipeline + delivered"
                  valueColor="#f59e0b"
                />
                <KpiCard
                  icon={<Rocket size={18} />}
                  title="Delivered Leads"
                  value={customersDashboard.deliveredOrders}
                  sub="Successfully delivered"
                  valueColor="#16a34a"
                />
                <KpiCard
                  icon={<Wallet size={18} />}
                  title="Total Revenue"
                  value={formatTZS(customersDashboard.totalRevenue)}
                  sub="From delivered leads only"
                  valueColor="#16a34a"
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

              <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("1.2fr 0.8fr", "1fr", "1fr"), gap: 16, marginBottom: 16 }}>
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

                <div style={{ ...styles.card, padding: 22 }}>
                  <div style={{ fontSize: 22, fontWeight: 800, marginBottom: 8 }}>Top Winner</div>
                  <div style={{ color: textSoft, marginBottom: 18 }}>Best product based on current score and profit.</div>
                  <div style={{ fontSize: 20, fontWeight: 800 }}>{bestProduct?.name || "N/A"}</div>
                  <div style={{ marginTop: 8, color: green, fontWeight: 700 }}>{bestProduct ? formatTZS(bestProduct.profit) : "N/A"}</div>
                  <div style={{ marginTop: 12 }}>{bestProduct ? <span style={getDecisionStyle(bestProduct.decision)}>{bestProduct.decision}</span> : <span style={getDecisionStyle("WATCH")}>WATCH</span>}</div>
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

              <div style={{ ...styles.card, padding: 22 }}>
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
                  <input style={styles.input} placeholder="Ex: Electric Callus Remover" value={expeditionForm.name} onChange={(e) => setExpeditionForm({ ...expeditionForm, name: e.target.value })} />
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
              <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(4, minmax(0, 1fr))", "repeat(2, minmax(0, 1fr))", "1fr"), gap: 16 }}>
                <KpiCard icon={<Boxes size={18} />} title="Catalog size" value={productsCatalogSummary.totalProducts} sub="Products ready in the catalog" />
                <KpiCard icon={<Archive size={18} />} title="Total units" value={productsCatalogSummary.totalUnits} sub="Imported stock volume" />
                <KpiCard icon={<Wallet size={18} />} title="Import budget" value={formatTZS(productsCatalogSummary.totalImportBudgetTzs)} sub="Purchase + shipping + charges" valueColor={accent} />
                <KpiCard icon={<TrendingUp size={18} />} title="Top score" value={`${productsCatalogSummary.topScore}/100`} sub={bestProduct ? bestProduct.name : "No highlighted product"} valueColor={green} />
              </div>

              <div style={{ ...styles.card, padding: 22 }}>
                <div style={styles.sectionHeader}>
                  <div>
                    <div style={styles.sectionEyebrow}>Catalog intelligence</div>
                    <div style={{ fontSize: 24, fontWeight: 900, marginTop: 8 }}>Produits en stock</div>
                    <div style={{ color: textSoft, marginTop: 6, lineHeight: 1.6 }}>Tous les produits approvisionnés apparaissent ici automatiquement, avec leur coût réel, leur quantité et leur profil logistique.</div>
                  </div>
                  <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
                    {editingProductId ? (
                      <button style={styles.btnSecondary} onClick={cancelEditingProduct}>
                        Cancel Edit
                      </button>
                    ) : null}
                    <button style={styles.btnSecondary} onClick={() => setActivePage("stock")}>
                      Gestion stock
                    </button>
                    <button style={styles.btnPrimary} onClick={() => setActivePage("multiDashboard")}>
                      Nouveau lot stock
                    </button>
                    <div style={{ ...styles.badge, background: "rgba(29,95,208,0.08)", color: accent, border: "1px solid rgba(29,95,208,0.12)" }}>
                      Live catalog
                    </div>
                  </div>
                </div>

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
              </div>
            </div>
          )}

{activePage === "stock" && (
            <div style={{ ...styles.card, padding: 22 }}>
              <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(6, minmax(0, 1fr))", "1fr 1fr", "1fr"), gap: 16, marginBottom: 16 }}>
                <KpiCard icon={<Archive size={18} />} title="Products tracked" value={stockForecastRows.length} sub="Products with live stock forecast" />
                <KpiCard icon={<AlertTriangle size={18} />} title="Critical stockouts" value={stockForecastRows.filter((product) => product.urgency === "Critical").length} sub="Projected stockout in 7 days or less" valueColor={red} />
                <KpiCard icon={<TrendingUp size={18} />} title="Watchlist" value={stockForecastRows.filter((product) => product.urgency === "Watch").length} sub="Projected stockout in 14 days or less" valueColor={amber} />
                <KpiCard icon={<CalendarDays size={18} />} title="Next stockout" value={stockForecastRows[0]?.projectedStockoutDate || "N/A"} sub={stockForecastRows[0] ? stockForecastRows[0].name : "No projection yet"} />
                <KpiCard icon={<Wallet size={18} />} title="Stock Value" value={formatUsdFromTzs(stockValueSummary.totalValueTzs)} sub="Available stock valuation" valueColor={accent} />
                <KpiCard icon={<Archive size={18} />} title="Aging 60+ days" value={stockValueSummary.aged60Products} sub={`${stockValueSummary.aged90Products} products above 90 days`} valueColor={amber} />
              </div>
              <div style={styles.sectionHeader}>
                <div>
                  <div style={{ fontSize: 22, fontWeight: 800 }}>Gestion de stock</div>
                  <div style={{ color: textSoft, marginTop: 6 }}>Suivez ici le stock réel, les quantités réservées, les sorties en livraison et le moment idéal pour recommander.</div>
                </div>
                <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                  <button style={styles.btnSecondary} onClick={() => setActivePage("products")}>
                    Produits en stock
                  </button>
                  <button style={styles.btnPrimary} onClick={() => setActivePage("multiDashboard")}>
                    Approvisionnement
                  </button>
                </div>
              </div>

              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "separate", borderSpacing: 0 }}>
                  <thead>
                    <tr>
                      {["Product", "Source", "Initial Stock", "Delivered", "Reserved", "Current Stock", "Available Stock", "Sales/Day", "Min Stock", "Reorder", "Forecast", "Stockout Date", "Status", "Estimated Arrival", "Arrival", "Transit Days"].map((head) => (
                        <th key={head} style={{ textAlign: "left", padding: "14px 12px", color: textSoft, fontSize: 13, borderBottom: `1px solid ${cardBorder}`, background: "#f8fafc" }}>{head}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {stockForecastRows.map((product) => {
                      const stockStatus = product.availableStock <= 0 ? "Out of Stock" : product.availableStock <= 10 ? "Low Stock" : "In Stock";
                      return (
                        <tr key={product.id}>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 700 }}>{product.name}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>{product.source || "—"}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>{product.initialStock}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>{product.deliveredUnits}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>{product.reservedStock}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>{product.currentStock}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 700 }}>{product.availableStock}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>{product.salesPerDay.toFixed(1)}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 700 }}>{product.reorderPoint}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}><span style={getDecisionStyle(product.reorderStatus)}>{product.reorderStatus}</span></td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>{product.daysUntilStockout != null ? `${Math.max(1, Math.round(product.daysUntilStockout))} days left` : "No pace yet"}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>{product.projectedStockoutDate}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}><span style={getDecisionStyle(stockStatus)}>{stockStatus}</span></td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>{product.stockOrderedAt ? addDaysToDateString(product.stockOrderedAt, Number(product.estimatedArrivalDays || 0)) : "—"}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>{product.source === "dubai" ? <span style={getDecisionStyle(product.stockArrivalStatus === "arrived" ? "Arrived" : "Pending")}>{product.stockArrivalStatus === "arrived" ? "Arrived" : `Pending (${product.nextArrivalCheckDate || "check"})`}</span> : <span style={getDecisionStyle("Arrived")}>Arrived</span>}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>{product.estimatedArrivalDays ?? "—"}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          )}

{activePage === "customersOrders" && (
            <div style={{ display: "grid", gap: 20 }}>
              <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(4, minmax(0, 1fr))", "repeat(2, minmax(0, 1fr))", "1fr"), gap: 16 }}>
                <KpiCard icon={<Users size={18} />} title="Total Orders" value={customersDashboard.totalOrders} sub="All customer orders" />
                <KpiCard icon={<ShoppingBag size={18} />} title="Confirmed" value={customersDashboard.confirmedOrders} sub="Active pipeline + delivered" />
                <KpiCard icon={<Rocket size={18} />} title="Delivered" value={customersDashboard.deliveredOrders} sub="Successfully delivered" />
                <KpiCard icon={<Wallet size={18} />} title="Revenue" value={formatTZS(customersDashboard.totalRevenue)} sub="Orders revenue estimate" valueColor={green} />
              </div>

              <div style={{ ...styles.card, padding: 22 }}>
                <div style={styles.sectionHeader}>
                  <div>
                    <div style={styles.sectionEyebrow}>Orders import</div>
                    <div style={{ fontSize: 24, fontWeight: 900, marginTop: 8 }}>Import Excel Orders</div>
                    <div style={{ color: textSoft, marginTop: 6, lineHeight: 1.6 }}>
                      Importez ici toutes les commandes clients avec leur statut de confirmation. Les commandes confirmees passent automatiquement dans le menu Shipping.
                    </div>
                  </div>
                  <button style={styles.btnPrimary} onClick={() => ordersImportInputRef.current?.click()}>
                    Import Excel Orders
                  </button>
                </div>

                {ordersImportNotice ? (
                  <div style={{ ...styles.softStat, marginTop: 16 }}>
                    <div style={{ fontSize: 11, fontWeight: 800, letterSpacing: 0.45, textTransform: "uppercase", color: textSoft }}>Excel import</div>
                    <div style={{ marginTop: 8, color: textMain, fontWeight: 700 }}>{ordersImportNotice}</div>
                    <div style={{ marginTop: 6, color: textSoft, fontSize: 13, lineHeight: 1.5 }}>
                      Supported columns include customer name, phone, product/product id, quantity, total price, order date, status, payment method, city, address, notes and external order id.
                    </div>
                    {ordersImportDetails ? (
                      <div style={{ marginTop: 12, display: "grid", gap: 8, color: textSoft, fontSize: 13, lineHeight: 1.5 }}>
                        <div>
                          Detected headers: {ordersImportDetails.detectedHeaders.length ? ordersImportDetails.detectedHeaders.join(" | ") : "N/A"}
                        </div>
                        <div>
                          Skip reasons: missing name {ordersImportDetails.reasonCounts.missingName}, missing phone {ordersImportDetails.reasonCounts.missingPhone}, missing product {ordersImportDetails.reasonCounts.missingProduct}, unknown product {ordersImportDetails.reasonCounts.unknownProduct}
                        </div>
                        {ordersImportDetails.unmatchedProducts.length ? (
                          <div>
                            Unmatched product examples: {ordersImportDetails.unmatchedProducts.join(" | ")}
                          </div>
                        ) : null}
                      </div>
                    ) : null}
                  </div>
                ) : null}
              </div>

              <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("minmax(0, 1.15fr) minmax(300px, 0.85fr)", "1fr", "1fr"), gap: 16 }}>
                <div style={{ ...styles.card, padding: 22 }}>
                  <div style={styles.sectionHeader}>
                    <div>
                      <div style={styles.sectionEyebrow}>Lead intake</div>
                      <div style={{ fontSize: 24, fontWeight: 900, marginTop: 8 }}>Nouveau client / commande</div>
                      <div style={{ color: textSoft, marginTop: 6, lineHeight: 1.6 }}>Entre les informations du client, le produit commandé et les détails de traitement dans une seule fiche opérateur.</div>
                    </div>
                    <button style={styles.btnPrimary} onClick={saveCustomerOrder}>Save Order</button>
                  </div>

                  <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(2, minmax(0, 1fr))", "1fr", "1fr"), gap: 16 }}>
                    <div style={styles.fieldBlock}>
                      <label style={styles.fieldLabel}>Nom du client</label>
                      <input style={styles.input} value={customerForm.customerName} onChange={(e) => setCustomerForm({ ...customerForm, customerName: e.target.value })} placeholder="Ex: Amina Yusuf" />
                    </div>
                    <div style={styles.fieldBlock}>
                      <label style={styles.fieldLabel}>Telephone</label>
                      <input style={styles.input} value={customerForm.phone} onChange={(e) => setCustomerForm({ ...customerForm, phone: e.target.value })} placeholder="Ex: +255712345678" />
                    </div>
                    <div style={styles.fieldBlock}>
                      <label style={styles.fieldLabel}>Ville</label>
                      <input style={styles.input} value={customerForm.city} onChange={(e) => setCustomerForm({ ...customerForm, city: e.target.value })} placeholder="Ex: Dar es Salaam" />
                    </div>
                    <div style={styles.fieldBlock}>
                      <label style={styles.fieldLabel}>Adresse</label>
                      <input style={styles.input} value={customerForm.address} onChange={(e) => setCustomerForm({ ...customerForm, address: e.target.value })} placeholder="Ex: Mikocheni, Block A" />
                    </div>
                    <div style={styles.fieldBlock}>
                      <label style={styles.fieldLabel}>Produit commande</label>
                      <select style={styles.input} value={customerForm.productId} onChange={(e) => setCustomerForm({ ...customerForm, productId: e.target.value })}>
                        {products.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
                      </select>
                    </div>
                    <div style={styles.fieldBlock}>
                      <label style={styles.fieldLabel}>Quantite</label>
                      <input style={styles.input} type="number" min="1" value={customerForm.quantity} onChange={(e) => setCustomerForm({ ...customerForm, quantity: e.target.value })} />
                    </div>
                    <div style={styles.fieldBlock}>
                      <label style={styles.fieldLabel}>Date de commande</label>
                      <input style={styles.input} type="date" value={customerForm.orderDate} onChange={(e) => setCustomerForm({ ...customerForm, orderDate: e.target.value })} />
                    </div>
                    <div style={styles.fieldBlock}>
                      <label style={styles.fieldLabel}>Methode de paiement</label>
                      <select style={styles.input} value={customerForm.paymentMethod} onChange={(e) => setCustomerForm({ ...customerForm, paymentMethod: e.target.value })}>
                        <option value="COD">COD</option>
                        <option value="Card">Card</option>
                        <option value="M-Pesa">M-Pesa</option>
                        <option value="Cash">Cash</option>
                      </select>
                    </div>
                    <div style={styles.fieldBlock}>
                      <label style={styles.fieldLabel}>Statut</label>
                      <select
                        style={styles.input}
                        value={normalizeOrderStatus(customerForm.status)}
                        onChange={(e) => setCustomerForm({ ...customerForm, status: e.target.value })}
                      >
                        {confirmationStatusCatalog.map((status) => (
                          <option key={status.value} value={status.value}>
                            {status.label}
                          </option>
                        ))}
                      </select>
                    </div>
                    <div style={styles.fieldBlock}>
                      <label style={styles.fieldLabel}>Source Lead</label>
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
                      <label style={styles.fieldLabel}>Campaign</label>
                      <input style={styles.input} value={customerForm.campaignName} onChange={(e) => setCustomerForm({ ...customerForm, campaignName: e.target.value })} placeholder="Campaign name" />
                    </div>
                    <div style={styles.fieldBlock}>
                      <label style={styles.fieldLabel}>Ad Set</label>
                      <input style={styles.input} value={customerForm.adsetName} onChange={(e) => setCustomerForm({ ...customerForm, adsetName: e.target.value })} placeholder="Ad set name" />
                    </div>
                    <div style={styles.fieldBlock}>
                      <label style={styles.fieldLabel}>Creative</label>
                      <input style={styles.input} value={customerForm.creativeName} onChange={(e) => setCustomerForm({ ...customerForm, creativeName: e.target.value })} placeholder="Creative or hook" />
                    </div>
                    <div style={styles.fieldBlock}>
                      <label style={styles.fieldLabel}>Priority</label>
                      <select style={styles.input} value={customerForm.priority} onChange={(e) => setCustomerForm({ ...customerForm, priority: e.target.value })}>
                        <option value="low">Low</option>
                        <option value="normal">Normal</option>
                        <option value="high">High</option>
                        <option value="urgent">Urgent</option>
                      </select>
                    </div>
                    <div style={styles.fieldBlock}>
                      <label style={styles.fieldLabel}>Customer Type</label>
                      <select style={styles.input} value={customerForm.customerType} onChange={(e) => setCustomerForm({ ...customerForm, customerType: e.target.value })}>
                        <option value="new">New</option>
                        <option value="repeat">Repeat</option>
                        <option value="vip">VIP</option>
                      </select>
                    </div>
                    <div style={styles.fieldBlock}>
                      <label style={styles.fieldLabel}>Call Attempts</label>
                      <input style={styles.input} type="number" min="0" value={customerForm.callAttempts} onChange={(e) => setCustomerForm({ ...customerForm, callAttempts: e.target.value })} />
                    </div>
                    <div style={styles.fieldBlock}>
                      <label style={styles.fieldLabel}>Cancel Reason</label>
                      <input style={styles.input} value={customerForm.cancelReason} onChange={(e) => setCustomerForm({ ...customerForm, cancelReason: e.target.value })} placeholder="Changed mind, too expensive..." />
                    </div>
                    <div style={styles.fieldBlock}>
                      <label style={styles.fieldLabel}>Unreached Reason</label>
                      <input style={styles.input} value={customerForm.unreachedReason} onChange={(e) => setCustomerForm({ ...customerForm, unreachedReason: e.target.value })} placeholder="No answer, switched off..." />
                    </div>
                    <div style={styles.fieldBlock}>
                      <label style={styles.fieldLabel}>Carrier / Agency</label>
                      <input style={styles.input} value={customerForm.carrierName} onChange={(e) => setCustomerForm({ ...customerForm, carrierName: e.target.value })} placeholder="Carrier or agency" />
                    </div>
                    <div style={styles.fieldBlock}>
                      <label style={styles.fieldLabel}>Tracking Number</label>
                      <input style={styles.input} value={customerForm.trackingNumber} onChange={(e) => setCustomerForm({ ...customerForm, trackingNumber: e.target.value })} placeholder="Waybill / reference" />
                    </div>
                    <div style={styles.fieldBlock}>
                      <label style={styles.fieldLabel}>Expected Delivery</label>
                      <input style={styles.input} type="date" value={customerForm.expectedDeliveryDate} onChange={(e) => setCustomerForm({ ...customerForm, expectedDeliveryDate: e.target.value })} />
                    </div>
                    <div style={styles.fieldBlock}>
                      <label style={styles.fieldLabel}>Return Reason</label>
                      <input style={styles.input} value={customerForm.returnReason} onChange={(e) => setCustomerForm({ ...customerForm, returnReason: e.target.value })} placeholder="Rejected, return stock..." />
                    </div>
                    <div style={styles.fieldBlock}>
                      <label style={styles.fieldLabel}>Notes</label>
                      <input style={styles.input} value={customerForm.notes} onChange={(e) => setCustomerForm({ ...customerForm, notes: e.target.value })} placeholder="Ex: Call in the afternoon" />
                    </div>
                  </div>
                </div>

                <div style={styles.heroAside}>
                  <div style={{ fontSize: 12, fontWeight: 800, letterSpacing: 0.55, textTransform: "uppercase", color: "rgba(255,255,255,0.72)" }}>Order preview</div>
                  <div style={{ marginTop: 10, fontSize: 24, fontWeight: 900, lineHeight: 1.08 }}>{selectedCustomerProduct?.name || "Select a product"}</div>
                  <div style={{ marginTop: 10, color: "rgba(255,255,255,0.78)", lineHeight: 1.55 }}>Visual check before saving the order, so the operator can validate value, status and payment flow at a glance.</div>
                  <div style={{ display: "grid", gap: 12, marginTop: 18 }}>
                    <MiniStat
                      label="Order value"
                      value={formatTZS(customerFormOrderValue)}
                      tone="green"
                      dark
                      sub={
                        customerFormPricing.matchedOffer
                          ? `${Math.max(1, Number(customerForm.quantity || 1))} item(s) | offer ${customerFormPricing.matchedOffer.minQty}+ pcs applied`
                          : `${Math.max(1, Number(customerForm.quantity || 1))} item(s) | base price`
                      }
                    />
                    <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("1fr 1fr", "1fr 1fr", "1fr"), gap: 12 }}>
                      <MiniStat label="Payment" value={customerForm.paymentMethod || "COD"} dark />
                      <MiniStat label="Status" value={formatStatusLabel(customerForm.status || "new-order")} tone="amber" dark />
                    </div>
                    <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("1fr 1fr", "1fr 1fr", "1fr"), gap: 12 }}>
                      <MiniStat label="Source lead" value={formatStatusLabel(customerForm.leadSource || "manual")} tone="blue" dark sub={customerForm.campaignName || "No campaign"} />
                      <MiniStat label="Priority" value={formatStatusLabel(customerForm.priority || "normal")} tone="amber" dark sub={`${formatStatusLabel(customerForm.customerType || "new")} customer`} />
                    </div>
                    <MiniStat label="Destination" value={customerForm.city || "City not set"} tone="blue" dark sub={customerForm.address || "Address not set"} />
                  </div>
                </div>
              </div>

              <div style={{ ...styles.card, padding: 22 }}>
                <div style={styles.sectionHeader}>
                  <div>
                    <div style={styles.sectionEyebrow}>Order pipeline</div>
                    <div style={{ fontSize: 24, fontWeight: 900, marginTop: 8 }}>Liste des commandes clients</div>
                    <div style={{ color: textSoft, marginTop: 6, lineHeight: 1.6 }}>Toutes les commandes saisies apparaissent ici avec leur produit, client, valeur et statut de traitement.</div>
                  </div>
                </div>

                <div style={{ display: "grid", gap: 16 }}>
                  <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(4, minmax(0, 1fr))", "1fr 1fr", "1fr"), gap: 12 }}>
                    <MiniStat label="Filtered value" value={formatTZS(filteredCustomerSummary.totalValue)} tone="blue" sub="Current order selection" />
                    <MiniStat label="Confirmed" value={filteredCustomerSummary.confirmed} tone="green" sub="Active revenue pipeline" />
                    <MiniStat label="Pending" value={filteredCustomerSummary.pending} tone="amber" sub="Still in confirmation flow" />
                    <MiniStat label="Cancelled" value={filteredCustomerSummary.cancelled} tone="amber" sub="Lost or rejected orders" />
                  </div>

                  <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("minmax(260px, 1fr) 180px 120px", "1fr 1fr", "1fr"), gap: 12 }}>
                    <input
                      style={styles.input}
                      value={customerListFilters.search}
                      onChange={(e) => setCustomerListFilters((prev) => ({ ...prev, search: e.target.value }))}
                      placeholder="Search name, phone, city, product, order id..."
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
                          {["Select", "Customer", "Product", "Order", "Owner", "Value", "Status", "History", "Actions"].map((head) => (
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
                                <button style={styles.btnSecondary} onClick={() => setCustomerHistoryTargetId(customer.id)}>
                                  View
                                </button>
                              </td>
                              <td style={{ padding: "10px 10px", borderBottom: `1px solid ${cardBorder}`, textAlign: "right", minWidth: 190 }}>
                                <div style={{ display: "inline-flex", gap: 8, flexWrap: "wrap", justifyContent: "flex-end" }}>
                                  <button
                                    style={{ ...styles.btnSecondary, background: "#fef2f2", color: red, border: "1px solid #fecaca", padding: "8px 10px", fontSize: 12 }}
                                    onClick={() => deleteCustomerOrder(customer.id)}
                                  >
                                    Delete Order
                                  </button>
                                </div>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>

                    {compactCustomerRows.length === 0 ? (
                      <div style={{ padding: 24, color: textSoft }}>No customer orders match the current filters.</div>
                    ) : null}
                  </div>
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
              <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                <button style={activePage === "tracking" ? styles.btnPrimary : styles.btnSecondary} onClick={() => setActivePage("tracking")}>Tracking</button>
                <button style={activePage === "serviceSum" ? styles.btnPrimary : styles.btnSecondary} onClick={() => setActivePage("serviceSum")}>Service Sum</button>
                <button style={activePage === "situations" ? styles.btnPrimary : styles.btnSecondary} onClick={() => setActivePage("situations")}>Rentabilité</button>
                <button style={activePage === "profitCenter" ? styles.btnPrimary : styles.btnSecondary} onClick={() => setActivePage("profitCenter")}>Profit Center</button>
              </div>
              <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(4, minmax(0, 1fr))", "repeat(2, minmax(0, 1fr))", "1fr"), gap: 16 }}>
                <KpiCard icon={<ClipboardList size={18} />} title="Tracking rows" value={trackingSummary.rows} sub="Manual and Meta-synced spend rows" />
                <KpiCard icon={<Wallet size={18} />} title="Ad spend" value={formatTZS(trackingSummary.spend)} sub={`${trackingSummary.orders} customer orders synced`} valueColor={accent} />
                <KpiCard icon={<TrendingUp size={18} />} title="Revenue" value={formatTZS(trackingSummary.revenue)} sub={`${trackingSummary.delivered} delivered units from orders`} valueColor={green} />
                <KpiCard icon={<Rocket size={18} />} title="Profit" value={formatTZS(trackingSummary.profit)} sub="Orders automate revenue and stock impact" valueColor={trackingSummary.profit >= 0 ? green : red} />
              </div>

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
                  <MiniStat label="Last import" value={metaAdsState.lastSyncAt ? new Date(metaAdsState.lastSyncAt).toLocaleString() : "Not imported"} tone="blue" sub={metaAdsState.lastSyncSummary ? `${metaAdsState.lastSyncSummary.matchedProducts} products matched | ${formatInteger(metaAdsState.lastSyncSummary.totalLeads || 0)} tracked leads | Total spend ${formatUsdFromTzs(metaAdsState.lastSyncSummary.accountTotalSpendTzs || metaAdsState.lifetimeSpendTzs || 0)}` : "No Meta import yet"} />
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

                {metaAdsNotice ? (
                  <div style={{ marginTop: 16, padding: "14px 16px", borderRadius: 18, border: `1px solid ${cardBorder}`, background: "linear-gradient(180deg, rgba(255,255,255,0.92), rgba(250,247,242,0.88))", color: textMain, boxShadow: "0 10px 22px rgba(23,32,51,0.05)" }}>
                    {metaAdsNotice}
                  </div>
                ) : null}

                {metaCampaignRows.length ? (
                  <div style={{ marginTop: 18, overflowX: "auto", border: `1px solid ${cardBorder}`, borderRadius: 24, background: "linear-gradient(180deg, rgba(255,255,255,0.95), rgba(246,242,236,0.9))", boxShadow: "0 18px 34px rgba(23,32,51,0.06)" }}>
                    <table style={{ width: "100%", borderCollapse: "separate", borderSpacing: 0 }}>
                      <thead>
                        <tr>
                          {["Campaign", "Match product", "Spend", "Tracked leads", "Tracking source", "CPL", "Clicks (all)", "Link clicks", "LPV", "CTR", "Impressions", "Reach", "CPC", "CPM"].map((head) => (
                            <th key={head} style={{ textAlign: "left", padding: "16px 14px", color: textSoft, fontSize: 12, fontWeight: 900, letterSpacing: 0.5, textTransform: "uppercase", borderBottom: `1px solid ${cardBorder}`, background: "linear-gradient(180deg, rgba(248,244,238,0.98), rgba(243,238,231,0.95))", whiteSpace: "nowrap" }}>
                              {head}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {metaCampaignRows.map((row, index) => {
                          const rowCpm = Number(row.impressions || 0) > 0 ? (Number(row.spend || 0) / Number(row.impressions || 0)) * 1000 : 0;
                          return (
                            <tr key={row.id} style={{ background: index % 2 === 0 ? "rgba(255,255,255,0.78)" : "rgba(248,244,238,0.72)" }}>
                              <td style={{ padding: "16px 14px", borderBottom: `1px solid ${cardBorder}`, minWidth: 240 }}>
                                <div style={{ fontWeight: 800 }}>{row.campaignName}</div>
                                <div style={{ color: textSoft, fontSize: 12, marginTop: 4 }}>{row.adsetName || "No ad set name"}</div>
                              </td>
                              <td style={{ padding: "16px 14px", borderBottom: `1px solid ${cardBorder}`, minWidth: 230 }}>
                                <select style={{ ...styles.input, borderRadius: 12 }} value={row.mappedProductId} onChange={(e) => updateMetaCampaignMapping(row.id, e.target.value)}>
                                  <option value="">Skip this campaign</option>
                                  {products.map((product) => (
                                    <option key={product.id} value={product.id}>
                                      {product.name}
                                    </option>
                                  ))}
                                </select>
                                <div style={{ color: textSoft, fontSize: 12, marginTop: 6 }}>
                                  {row.suggestedProductId ? `Auto-match: ${products.find((product) => product.id === row.suggestedProductId)?.name || "Detected"}` : "No automatic match"}
                                </div>
                              </td>
                              <td style={{ padding: "16px 14px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 800, color: accent }}>{formatMetaMoney(row.spend)}</td>
                              <td style={{ padding: "16px 14px", borderBottom: `1px solid ${cardBorder}`, fontWeight: 800 }}>{formatInteger(row.trackedLeads ?? row.leads)}</td>
                              <td style={{ padding: "16px 14px", borderBottom: `1px solid ${cardBorder}`, color: textSoft, minWidth: 190 }}>
                                <div>{formatMetaLeadSourceLabel(row.trackedLeadType)}</div>
                                <div style={{ color: textSoft, fontSize: 12, marginTop: 4 }}>
                                  {Number(row.actualLeads || row.leads || 0) > 0 ? `${formatInteger(row.actualLeads || row.leads)} actual lead events` : row.leadType ? formatStatusLabel(row.leadType) : "No native lead event"}
                                </div>
                              </td>
                              <td style={{ padding: "16px 14px", borderBottom: `1px solid ${cardBorder}` }}>{formatMetaMoney(row.costPerLead)}</td>
                              <td style={{ padding: "16px 14px", borderBottom: `1px solid ${cardBorder}` }}>{formatInteger(row.clicks)}</td>
                              <td style={{ padding: "16px 14px", borderBottom: `1px solid ${cardBorder}` }}>{formatInteger(row.inlineLinkClicks)}</td>
                              <td style={{ padding: "16px 14px", borderBottom: `1px solid ${cardBorder}` }}>{formatInteger(row.landingPageViews)}</td>
                              <td style={{ padding: "16px 14px", borderBottom: `1px solid ${cardBorder}` }}>{Number(row.ctr || 0).toFixed(2)}%</td>
                              <td style={{ padding: "16px 14px", borderBottom: `1px solid ${cardBorder}` }}>{formatInteger(row.impressions)}</td>
                              <td style={{ padding: "16px 14px", borderBottom: `1px solid ${cardBorder}` }}>{formatInteger(row.reach)}</td>
                              <td style={{ padding: "16px 14px", borderBottom: `1px solid ${cardBorder}` }}>{formatMetaMoney(row.cpc)}</td>
                              <td style={{ padding: "16px 14px", borderBottom: `1px solid ${cardBorder}` }}>{formatMetaMoney(row.cpm || rowCpm)}</td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                ) : null}
              </div>

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
            </div>
          )}

{["serviceSum", "financeHub"].includes(activePage) && (
            <div style={{ ...styles.card, padding: 22 }}>
              <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 18 }}>
                <button style={activePage === "tracking" ? styles.btnPrimary : styles.btnSecondary} onClick={() => setActivePage("tracking")}>Tracking</button>
                <button style={activePage === "serviceSum" ? styles.btnPrimary : styles.btnSecondary} onClick={() => setActivePage("serviceSum")}>Service Sum</button>
                <button style={activePage === "situations" ? styles.btnPrimary : styles.btnSecondary} onClick={() => setActivePage("situations")}>Rentabilité</button>
                <button style={activePage === "profitCenter" ? styles.btnPrimary : styles.btnSecondary} onClick={() => setActivePage("profitCenter")}>Profit Center</button>
              </div>
              <div style={styles.sectionHeader}>
                <div>
                  <div style={{ fontSize: 22, fontWeight: 800 }}>Service Sum</div>
                  <div style={{ color: textSoft, marginTop: 6 }}>Centre de synthese automatique : confirmations, livraisons, revenus, depenses pub et profit sont maintenant relies aux vraies donnees de l'app.</div>
                </div>
              </div>

              <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("1fr 1fr", "1fr", "1fr"), gap: 12, marginBottom: 20 }}>
                <select style={styles.input} value={selectedService} onChange={(e) => setSelectedService(e.target.value)}>
                  <option value="standard">Standard</option>
                  <option value="codzoss">CODZOSS</option>
                </select>
                <select style={styles.input} value={selectedCountry} onChange={(e) => setSelectedCountry(e.target.value)}>
                  <option value="tanzania">Tanzania</option>
                  <option value="kenya">Kenya</option>
                </select>
              </div>

              {liveServiceDataset ? (
                <>
                  <div style={{ ...styles.card, padding: 18, marginBottom: 20, background: "linear-gradient(180deg, rgba(255,255,255,0.96), rgba(238,246,255,0.88))" }}>
                    <div style={styles.sectionEyebrow}>Live app data</div>
                    <div style={{ fontSize: 24, fontWeight: 900, marginTop: 8 }}>Automatic business summary</div>
                    <div style={{ color: textSoft, marginTop: 8, lineHeight: 1.6 }}>
                      Cette partie est entierement alimentee par vos commandes, vos statuts, vos produits et les depenses du menu Tracking. Plus besoin de ressaisir les confirmations et livraisons ici.
                    </div>
                  </div>

                  <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(3, minmax(0, 1fr))", "1fr 1fr", "1fr"), gap: 16 }}>
                    <KpiCard title="Confirmed" value={liveServiceDataset.confirmed} sub={`${Math.round(liveServiceDataset.confirmationRate * 100)}% confirmation rate`} />
                    <KpiCard title="Delivered" value={liveServiceDataset.delivered} sub={`${Math.round(liveServiceDataset.deliveryRate * 100)}% delivery rate`} />
                    <KpiCard title="Delivered Units" value={liveServiceDataset.deliveredUnits} sub="Real delivered quantity" />
                    <KpiCard title="Cost / Lead USD" value={formatUSD(liveServiceDataset.costPerLeadUsd)} sub="Ad spend / total leads" />
                    <KpiCard title="Break-even Ads / Lead" value={formatUSD(liveServiceDataset.breakEvenCplUsd)} sub="Max ad cost per lead to stay profitable" valueColor={liveServiceDataset.breakEvenCplUsd >= liveServiceDataset.costPerLeadUsd ? green : red} />
                    <KpiCard title="Revenue USD" value={formatUSD(liveServiceDataset.revenueUsd)} sub="Live order revenue" valueColor={green} />
                    <KpiCard title="Revenue TZS" value={formatTZS(liveServiceDataset.revenueTzs)} sub="Delivered orders revenue" />
                    <KpiCard title="Product Cost USD" value={formatUSD(liveServiceDataset.productCostTotalUsd)} sub="Delivered units import cost" />
                    <KpiCard title="Local Delivery USD" value={formatUSD(liveServiceDataset.localDeliveryCostUsd)} sub="Last-mile delivery cost" />
                    <KpiCard title="Service Charges USD" value={formatUSD(liveServiceDataset.totalServiceChargeUsd)} sub="Platform fee + per delivery fee" />
                    <KpiCard title="Profit / Commande USD" value={formatUSD(liveServiceDataset.profitPerOrderUsd)} sub="Net profit / delivered order" valueColor={liveServiceDataset.profitPerOrderUsd >= 0 ? green : red} />
                    <KpiCard title="Profit / Piece USD" value={formatUSD(liveServiceDataset.profitPerPieceUsd)} sub="Net profit / delivered unit" valueColor={liveServiceDataset.profitPerPieceUsd >= 0 ? green : red} />
                    <KpiCard title="Total Profit USD" value={formatUSD(liveServiceDataset.totalProfitUsd)} sub="Revenue - costs - service - ads" valueColor={liveServiceDataset.totalProfitUsd >= 0 ? green : red} />
                    <KpiCard title="Score" value={`${liveServiceDataset.score}/100`} sub={liveServiceDataset.decision} valueColor={liveServiceDataset.decision === "GOOD PRODUCT" ? green : liveServiceDataset.decision === "BAD PRODUCT" ? red : amber} />
                  </div>

                  <div style={{ ...styles.card, padding: 18, marginTop: 22, background: "linear-gradient(180deg, rgba(255,255,255,0.98), rgba(248,244,238,0.88))" }}>
                    <div style={styles.sectionEyebrow}>Scenario simulator</div>
                    <div style={{ fontSize: 20, fontWeight: 900, marginTop: 8 }}>Optional projection mode</div>
                    <div style={{ color: textSoft, marginTop: 8, marginBottom: 16, lineHeight: 1.6 }}>
                      Utilisez ce bloc seulement si vous voulez tester un scenario futur. Les donnees live ci-dessus restent la reference principale de l'application.
                    </div>

                    <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(3, minmax(0, 1fr))", "1fr 1fr", "1fr"), gap: 12, marginBottom: 20 }}>
                      <div style={styles.fieldBlock}><label style={styles.fieldLabel}>Total Leads</label><input style={styles.input} value={serviceForm.totalLeads} onChange={(e) => setServiceForm({ ...serviceForm, totalLeads: e.target.value })} /></div>
                      <div style={styles.fieldBlock}><label style={styles.fieldLabel}>Confirmation Rate %</label><input style={styles.input} value={serviceForm.confirmationRate} onChange={(e) => setServiceForm({ ...serviceForm, confirmationRate: e.target.value })} /></div>
                      <div style={styles.fieldBlock}><label style={styles.fieldLabel}>Delivery Rate %</label><input style={styles.input} value={serviceForm.deliveryRate} onChange={(e) => setServiceForm({ ...serviceForm, deliveryRate: e.target.value })} /></div>
                      <div style={styles.fieldBlock}><label style={styles.fieldLabel}>Selling Price TZS</label><input style={styles.input} value={serviceForm.sellingPriceTzs} onChange={(e) => setServiceForm({ ...serviceForm, sellingPriceTzs: e.target.value })} /></div>
                      <div style={styles.fieldBlock}><label style={styles.fieldLabel}>Product Cost TZS</label><input style={styles.input} value={serviceForm.productCostTzs} onChange={(e) => setServiceForm({ ...serviceForm, productCostTzs: e.target.value })} /></div>
                      <div style={styles.fieldBlock}><label style={styles.fieldLabel}>CPL USD</label><input style={styles.input} value={serviceForm.cplUsd} onChange={(e) => setServiceForm({ ...serviceForm, cplUsd: e.target.value })} /></div>
                    </div>

                    <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(3, minmax(0, 1fr))", "1fr 1fr", "1fr"), gap: 16 }}>
                      <KpiCard title="Projected Confirmed" value={selectedServiceDataset.confirmed} sub="Simulator only" />
                      <KpiCard title="Projected Delivered" value={selectedServiceDataset.delivered} sub="Simulator only" />
                      <KpiCard title="Break-even Ads / Lead" value={formatUSD(selectedServiceDataset.breakEvenCplUsd)} sub="Max ad cost per lead before profit turns negative" valueColor={selectedServiceDataset.breakEvenCplUsd >= selectedServiceDataset.costPerLeadUsd ? green : red} />
                      <KpiCard title="Projected Revenue USD" value={formatUSD(selectedServiceDataset.revenueUsd)} sub="Delivered x selling price" valueColor={green} />
                      <KpiCard title="Projected Charges USD" value={formatUSD(selectedServiceDataset.totalServiceChargeUsd)} sub="Service fee + delivery fees" />
                      <KpiCard title="Profit / Commande USD" value={formatUSD(selectedServiceDataset.profitPerOrderUsd)} sub="Net profit / delivered order" valueColor={selectedServiceDataset.profitPerOrderUsd >= 0 ? green : red} />
                      <KpiCard title="Projected Profit USD" value={formatUSD(selectedServiceDataset.totalProfitUsd)} sub="Scenario net result" valueColor={selectedServiceDataset.totalProfitUsd >= 0 ? green : red} />
                      <KpiCard title="Projected Score" value={`${selectedServiceDataset.score}/100`} sub={selectedServiceDataset.decision} valueColor={selectedServiceDataset.decision === "GOOD PRODUCT" ? green : selectedServiceDataset.decision === "BAD PRODUCT" ? red : amber} />
                    </div>
                  </div>
                </>
              ) : (
                <div style={{ ...styles.card, padding: 18, background: "#fff7ed", border: "1px solid #fed7aa" }}>
                  <div style={{ fontWeight: 800, color: amber, marginBottom: 6 }}>No data yet</div>
                  <div style={{ color: textSoft }}>Il n'y a pas encore de regles pour ce service et ce pays.</div>
                </div>
              )}
            </div>
          )}

{["situations", "financeHub"].includes(activePage) && (
            <div style={{ display: "grid", gap: 20 }}>
              <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                <button style={activePage === "tracking" ? styles.btnPrimary : styles.btnSecondary} onClick={() => setActivePage("tracking")}>Tracking</button>
                <button style={activePage === "serviceSum" ? styles.btnPrimary : styles.btnSecondary} onClick={() => setActivePage("serviceSum")}>Service Sum</button>
                <button style={activePage === "situations" ? styles.btnPrimary : styles.btnSecondary} onClick={() => setActivePage("situations")}>Rentabilité</button>
                <button style={activePage === "profitCenter" ? styles.btnPrimary : styles.btnSecondary} onClick={() => setActivePage("profitCenter")}>Profit Center</button>
              </div>
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

              <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("1fr 1fr", "1fr", "1fr"), gap: 20 }}>
                <div style={{ ...styles.card, padding: 22 }}>
                  <div style={styles.sectionHeader}>
                    <div>
                      <div style={styles.sectionEyebrow}>Payroll</div>
                      <div style={{ fontSize: 22, fontWeight: 900, marginTop: 8 }}>Employee salaries</div>
                    </div>
                    <button style={styles.btnSecondary} onClick={addSituationSalary}>Add salary</button>
                  </div>
                  <div style={{ display: "grid", gap: 12 }}>
                    {situationData.salaries.length ? (
                      situationData.salaries.map((entry) => (
                        <div key={entry.id} style={{ display: "grid", gridTemplateColumns: responsiveColumns("1fr 1fr 180px auto", "1fr 1fr", "1fr"), gap: 10, alignItems: "end" }}>
                          <div style={styles.fieldBlock}>
                            <label style={styles.fieldLabel}>Name</label>
                            <input style={styles.input} value={entry.name} onChange={(e) => updateSituationSalary(entry.id, "name", e.target.value)} />
                          </div>
                          <div style={styles.fieldBlock}>
                            <label style={styles.fieldLabel}>Role</label>
                            <input style={styles.input} value={entry.role} onChange={(e) => updateSituationSalary(entry.id, "role", e.target.value)} />
                          </div>
                          <div style={styles.fieldBlock}>
                            <label style={styles.fieldLabel}>Salary USD</label>
                            <input style={styles.input} type="number" min="0" step="0.01" value={(Number(entry.amountTzs || 0) / USD_TO_TZS).toFixed(2)} onChange={(e) => updateSituationSalary(entry.id, "amountTzs", e.target.value)} />
                          </div>
                          <button style={{ ...styles.btnSecondary, background: "#fef2f2", color: red, border: "1px solid #fecaca" }} onClick={() => removeSituationSalary(entry.id)}>
                            Remove
                          </button>
                        </div>
                      ))
                    ) : (
                      <div style={{ color: textSoft }}>No salary added yet.</div>
                    )}
                  </div>
                </div>

                <div style={{ ...styles.card, padding: 22 }}>
                  <div style={styles.sectionHeader}>
                    <div>
                      <div style={styles.sectionEyebrow}>Fixed costs</div>
                      <div style={{ fontSize: 22, fontWeight: 900, marginTop: 8 }}>Other fixed charges</div>
                    </div>
                    <button style={styles.btnSecondary} onClick={addSituationFixedCharge}>Add charge</button>
                  </div>
                  <div style={{ display: "grid", gap: 12 }}>
                    {situationData.fixedCharges.length ? (
                      situationData.fixedCharges.map((entry) => (
                        <div key={entry.id} style={{ display: "grid", gridTemplateColumns: responsiveColumns("1fr 180px auto", "1fr 180px auto", "1fr"), gap: 10, alignItems: "end" }}>
                          <div style={styles.fieldBlock}>
                            <label style={styles.fieldLabel}>Charge label</label>
                            <input style={styles.input} value={entry.label} onChange={(e) => updateSituationFixedCharge(entry.id, "label", e.target.value)} />
                          </div>
                          <div style={styles.fieldBlock}>
                            <label style={styles.fieldLabel}>Amount USD</label>
                            <input style={styles.input} type="number" min="0" step="0.01" value={(Number(entry.amountTzs || 0) / USD_TO_TZS).toFixed(2)} onChange={(e) => updateSituationFixedCharge(entry.id, "amountTzs", e.target.value)} />
                          </div>
                          <button style={{ ...styles.btnSecondary, background: "#fef2f2", color: red, border: "1px solid #fecaca" }} onClick={() => removeSituationFixedCharge(entry.id)}>
                            Remove
                          </button>
                        </div>
                      ))
                    ) : (
                      <div style={{ color: textSoft }}>No fixed charge added yet.</div>
                    )}
                  </div>
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

{["executive", "performanceHub"].includes(activePage) && (
            <div style={{ display: "grid", gap: 20 }}>
              <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(4, minmax(0, 1fr))", "repeat(2, minmax(0, 1fr))", "1fr"), gap: 16 }}>
                <KpiCard icon={<ClipboardList size={18} />} title="Total Leads" value={customersDashboard.totalOrders} sub={`${Math.round(customersDashboard.confirmationRate)}% confirmation rate`} />
                <KpiCard icon={<ShoppingBag size={18} />} title="Delivered Orders" value={customersDashboard.deliveredOrders} sub={`${Math.round(customersDashboard.deliveryRate)}% delivery rate`} valueColor={green} />
                <KpiCard icon={<Wallet size={18} />} title="Revenue" value={formatUsdFromTzs(liveAutomationSummary.totalRevenueTzs)} sub="Delivered orders revenue" valueColor={green} />
                <KpiCard icon={<TrendingUp size={18} />} title="Gross Profit" value={formatUsdFromTzs(executiveSummary.grossProfitTzs)} sub="Before fixed charges" valueColor={Number(executiveSummary.grossProfitTzs || 0) >= 0 ? green : red} />
                <KpiCard icon={<Calculator size={18} />} title="Net After Fixed" value={formatUsdFromTzs(executiveSummary.estimatedNetAfterFixedTzs)} sub="Estimated profit after fixed charges" valueColor={Number(executiveSummary.estimatedNetAfterFixedTzs || 0) >= 0 ? green : red} />
                <KpiCard icon={<Archive size={18} />} title="Stock Value" value={formatUsdFromTzs(executiveSummary.stockImmobilizedTzs)} sub="Capital locked in available stock" valueColor={amber} />
                <KpiCard icon={<AlertTriangle size={18} />} title="Open Tasks" value={executiveSummary.openTasks} sub={`${executiveSummary.highPriorityTasks} high priority`} valueColor={executiveSummary.highPriorityTasks > 0 ? red : accent} />
                <KpiCard icon={<Rocket size={18} />} title="Top Product" value={profitCenterSummary.topProduct?.name || "N/A"} sub={profitCenterSummary.topProduct ? `${formatUsdFromTzs(profitCenterSummary.topProduct.cumulativeProfitTzs || 0)} cumulative profit` : "No product data yet"} />
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
            </div>
          )}

{["profitCenter", "financeHub"].includes(activePage) && (
            <div style={{ display: "grid", gap: 20 }}>
              <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                <button style={activePage === "tracking" ? styles.btnPrimary : styles.btnSecondary} onClick={() => setActivePage("tracking")}>Tracking</button>
                <button style={activePage === "serviceSum" ? styles.btnPrimary : styles.btnSecondary} onClick={() => setActivePage("serviceSum")}>Service Sum</button>
                <button style={activePage === "situations" ? styles.btnPrimary : styles.btnSecondary} onClick={() => setActivePage("situations")}>Rentabilité</button>
                <button style={activePage === "profitCenter" ? styles.btnPrimary : styles.btnSecondary} onClick={() => setActivePage("profitCenter")}>Profit Center</button>
              </div>
              <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(6, minmax(0, 1fr))", "repeat(2, minmax(0, 1fr))", "1fr"), gap: 16 }}>
                <KpiCard icon={<Wallet size={18} />} title="Revenue" value={formatUsdFromTzs(profitCenterSummary.revenueTzs)} sub="All products combined" valueColor={green} />
                <KpiCard icon={<TrendingUp size={18} />} title="Gross Profit" value={formatUsdFromTzs(profitCenterSummary.profitTzs)} sub="Revenue - Meta total ads - delivered cost" valueColor={Number(profitCenterSummary.profitTzs || 0) >= 0 ? green : red} />
                <KpiCard icon={<ClipboardList size={18} />} title="Ads Charges" value={formatUsdFromTzs(profitCenterSummary.adsSpendTzs)} sub={profitCenterSummary.lastHourlyAdsSnapshot ? `Meta maximum total | last ${profitCenterSummary.lastHourlyAdsSnapshot.bucket}` : "Meta maximum total auto check"} valueColor={amber} />
                <KpiCard icon={<Archive size={18} />} title="Product Fixed Charges" value={formatUsdFromTzs(profitCenterSummary.fixedChargesTzs)} sub="Sourcing + import burden by product" valueColor={amber} />
                <KpiCard icon={<Calculator size={18} />} title="Net After Fixed" value={formatUsdFromTzs(profitCenterSummary.netAfterFixedTzs)} sub="Gross profit after product fixed charges" valueColor={Number(profitCenterSummary.netAfterFixedTzs || 0) >= 0 ? green : red} />
                <KpiCard icon={<Boxes size={18} />} title="Profitable Products" value={profitCenterSummary.profitableProducts} sub={`${profitCenterRows.length} products tracked`} />
              </div>

              <div style={{ ...styles.card, padding: 22 }}>
                <div style={styles.sectionHeader}>
                  <div>
                    <div style={styles.sectionEyebrow}>Net profit center</div>
                    <div style={{ fontSize: 24, fontWeight: 900, marginTop: 8 }}>Product economics and profitability</div>
                    <div style={{ color: textSoft, marginTop: 6, lineHeight: 1.6 }}>
                      This table helps you see which products truly create cash, which ones are eating margin, and where fixed product cost is still too heavy.
                    </div>
                  </div>
                </div>

                <div style={{ overflowX: "auto", border: `1px solid ${cardBorder}`, borderRadius: 20 }}>
                  <table style={{ width: "100%", borderCollapse: "separate", borderSpacing: 0 }}>
                    <thead>
                      <tr>
                        {["Product", "Orders", "Delivered Units", "Revenue", "Ads Charges", "Gross Profit", "Fixed Product Charges", "Net After Fixed", "Profit / Order", "Margin", "Action"].map((head) => (
                          <th key={head} style={{ textAlign: "left", padding: "14px 12px", color: textSoft, fontSize: 12, fontWeight: 800, letterSpacing: 0.4, textTransform: "uppercase", borderBottom: `1px solid ${cardBorder}`, background: "rgba(247, 243, 237, 0.92)" }}>
                            {head}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {profitCenterRows.map((row, index) => (
                        <tr key={row.id} style={{ background: index % 2 === 0 ? "rgba(255,255,255,0.72)" : "rgba(250,247,242,0.8)" }}>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>
                            <div style={{ fontWeight: 800 }}>{row.name}</div>
                            <div style={{ color: textSoft, fontSize: 12, marginTop: 4 }}>{row.availableStock} available | {row.reorderStatus}</div>
                          </td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>{row.orders}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>{row.deliveredUnits}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>{formatUsdFromTzs(row.revenue)}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>
                            <div style={{ fontWeight: 800 }}>{formatUsdFromTzs(row.cumulativeAdsTzs)}</div>
                            <div style={{ color: textSoft, fontSize: 12, marginTop: 4 }}>
                              {Number(row.cumulativeAdsTzs || 0) > 0 ? "Manual ads input" : "No ads input yet"}
                            </div>
                          </td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}`, color: Number(row.cumulativeProfitTzs || 0) >= 0 ? green : red, fontWeight: 800 }}>
                            <div>{formatUsdFromTzs(row.cumulativeProfitTzs)}</div>
                            <div style={{ color: textSoft, fontSize: 12, marginTop: 4, fontWeight: 600 }}>Delivered cost {formatUsdFromTzs(row.deliveredLogisticsTzs)}</div>
                          </td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>{formatUsdFromTzs(row.fixedProductChargesTzs)}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}`, color: Number(row.netAfterFixedTzs || 0) >= 0 ? green : red, fontWeight: 800 }}>{formatUsdFromTzs(row.netAfterFixedTzs)}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>{formatUsdFromTzs(row.profitPerOrderTzs)}</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>{Number(row.marginPercentLive || 0).toFixed(1)}%</td>
                          <td style={{ padding: "14px 12px", borderBottom: `1px solid ${cardBorder}` }}>
                            <div style={getDecisionStyle(row.decision || "WATCH")}>{row.decision || "WATCH"}</div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  {profitCenterRows.length === 0 ? <div style={{ padding: 24, color: textSoft }}>No product profitability data yet.</div> : null}
                </div>

                <div style={{ ...styles.softStat, marginTop: 16 }}>
                  <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap" }}>
                    <div>
                      <div style={{ fontSize: 11, fontWeight: 800, letterSpacing: 0.45, textTransform: "uppercase", color: textSoft }}>Meta daily total spend log</div>
                      <div style={{ marginTop: 8, fontSize: 20, fontWeight: 900 }}>The app requests Meta maximum total spend, then checks once per day for the new total automatically</div>
                    </div>
                    <div style={{ ...styles.badge, background: "rgba(199,131,34,0.12)", color: amber, border: "1px solid rgba(199,131,34,0.18)" }}>
                      {profitCenterSummary.lastHourlyAdsSnapshot ? `Last day ${profitCenterSummary.lastHourlyAdsSnapshot.bucket}` : "Waiting first Meta daily check"}
                    </div>
                  </div>
                  <div style={{ display: "grid", gridTemplateColumns: responsiveColumns("repeat(2, minmax(0, 1fr))", "1fr", "1fr"), gap: 12, marginTop: 14 }}>
                    <div style={{ ...styles.softStat, padding: 14 }}>
                      <div style={{ fontSize: 11, fontWeight: 800, letterSpacing: 0.4, textTransform: "uppercase", color: textSoft }}>Meta total ads spent</div>
                      <div style={{ marginTop: 8, fontSize: 24, fontWeight: 900, color: amber }}>{formatUsdFromTzs(profitCenterSummary.adsSpendTzs)}</div>
                    </div>
                    <div style={{ ...styles.softStat, padding: 14 }}>
                      <div style={{ fontSize: 11, fontWeight: 800, letterSpacing: 0.4, textTransform: "uppercase", color: textSoft }}>Current mapped live ads</div>
                      <div style={{ marginTop: 8, fontSize: 24, fontWeight: 900, color: accent }}>{formatUsdFromTzs(profitCenterSummary.liveObservedAdsTzs)}</div>
                    </div>
                  </div>
                  <div style={{ display: "grid", gap: 10, marginTop: 14 }}>
                    {(metaAdsState.dailySpendSnapshots || []).slice(0, 6).map((snapshot) => (
                      <div key={snapshot.id} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, padding: "12px 14px", borderRadius: 14, background: "rgba(255,255,255,0.72)", border: `1px solid ${cardBorder}` }}>
                        <div>
                          <div style={{ fontWeight: 800 }}>{snapshot.bucket}</div>
                          <div style={{ color: textSoft, fontSize: 12, marginTop: 4 }}>
                            {snapshot.capturedAt ? new Date(snapshot.capturedAt).toLocaleString() : "No timestamp"} | New today {formatUsdFromTzs(snapshot.newSpendTzs || 0)}
                          </div>
                        </div>
                        <div style={{ fontWeight: 900, color: accent }}>{formatUsdFromTzs(snapshot.totalSpendTzs)}</div>
                      </div>
                    ))}
                    {!(metaAdsState.dailySpendSnapshots || []).length ? <div style={{ color: textSoft }}>No Meta daily total captured yet. Keep the app open and the app will request Meta maximum total spend automatically once per day.</div> : null}
                  </div>
                </div>
              </div>
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

{["audit", "operationsHub"].includes(activePage) && (
            <div style={{ display: "grid", gap: 20 }}>
              <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                <button style={activePage === "taskCenter" ? styles.btnPrimary : styles.btnSecondary} onClick={() => setActivePage("taskCenter")}>Operations</button>
                <button style={activePage === "audit" ? styles.btnPrimary : styles.btnSecondary} onClick={() => setActivePage("audit")}>Audit</button>
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

/*
TEST CASES:
1. Save a product from Expedition Product -> it appears in Products and Stock.
2. Delete a product -> product and linked tracking rows disappear.
3. Add tracking row -> dashboard KPIs update.
4. Backup JSON -> file downloads with products, tracking, customers, and service form.
5. Restore JSON -> products, tracking, customers, and service form return.
6. Dashboard shows reorder alerts when available stock is low.
*/
