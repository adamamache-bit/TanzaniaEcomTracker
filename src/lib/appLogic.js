import * as XLSX from "xlsx";

export const INITIAL_CUSTOMERS = [
  {
    id: "C001",
    customerName: "Amina Yusuf",
    phone: "+255712345678",
    city: "Dar es Salaam",
    address: "Mikocheni, Block A",
    productId: "P001",
    quantity: 1,
    orderDate: "2026-04-12",
    paymentMethod: "COD",
    status: "new",
    notes: "Call in the afternoon",
  },
];

export const USD_TO_TZS = 2750;
export const META_API_PORT = 4174;
export const SHARED_API_PORT = 4174;

export const initialProducts = [
  {
    id: "P001",
    name: "Electric Callus Remover",
    source: "china",
    sellingPrice: 39000,
    purchaseUnitPrice: 12000,
    totalQty: 100,
    shippingTotal: 180000,
    otherCharges: 70000,
    delivery: 7000,
    estimatedArrivalDays: 15,
    stockArrivalStatus: "arrived",
    stockOrderedAt: "2026-04-01",
    nextArrivalCheckDate: null,
    stockArrivedAt: "2026-04-16",
  },
  {
    id: "P002",
    name: "Hair Dryer 5 in 1",
    source: "dubai",
    sellingPrice: 85000,
    purchaseUnitPrice: 35000,
    totalQty: 60,
    shippingTotal: 220000,
    otherCharges: 90000,
    delivery: 9000,
    estimatedArrivalDays: 3,
    stockArrivalStatus: "pending",
    stockOrderedAt: "2026-04-11",
    nextArrivalCheckDate: "2026-04-14",
    stockArrivedAt: null,
  },
  {
    id: "P003",
    name: "Whitening Toothpaste",
    source: "china",
    sellingPrice: 25000,
    purchaseUnitPrice: 6000,
    totalQty: 200,
    shippingTotal: 120000,
    otherCharges: 40000,
    delivery: 6500,
    estimatedArrivalDays: 15,
    stockArrivalStatus: "arrived",
    stockOrderedAt: "2026-03-15",
    nextArrivalCheckDate: null,
    stockArrivedAt: "2026-03-30",
  },
];

export const initialTracking = [];

export const serviceCountryData = {
  standard: {
    tanzania: {
      label: "Standard Tanzania",
      usdToTzs: 2750,
      serviceFeePercent: 0,
      deliveryFeeUsdPerDelivered: 8.5,
    },
  },
  codzoss: {},
};

export const formatDateInput = (date) => {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
};

export const parseDateInput = (dateString) => {
  if (!dateString) return null;
  const [year, month, day] = String(dateString).split("-").map(Number);
  if (!year || !month || !day) return null;
  return new Date(year, month - 1, day);
};

export const getTodayString = () => formatDateInput(new Date());

export const addDaysToDateString = (dateString, days) => {
  const date = parseDateInput(dateString);
  if (!date) return "N/A";
  if (Number.isNaN(date.getTime())) return "â€”";
  date.setDate(date.getDate() + Number(days || 0));
  return formatDateInput(date);
};

export const formatLongDate = (dateString) => {
  const date = parseDateInput(dateString);
  if (!date || Number.isNaN(date.getTime())) return "N/A";
  return new Intl.DateTimeFormat("en-GB", {
    day: "numeric",
    month: "short",
    year: "numeric",
  }).format(date);
};

export const addMonths = (date, months) => {
  const next = new Date(date);
  next.setMonth(next.getMonth() + months);
  return next;
};

export const startOfMonth = (date) => new Date(date.getFullYear(), date.getMonth(), 1);
export const endOfMonth = (date) => new Date(date.getFullYear(), date.getMonth() + 1, 0);

export const startOfWeek = (date) => {
  const next = new Date(date);
  const diff = next.getDate() - next.getDay();
  return new Date(next.getFullYear(), next.getMonth(), diff);
};

export const buildCalendarMatrix = (baseDate) => {
  const monthStart = startOfMonth(baseDate);
  const monthEnd = endOfMonth(baseDate);
  const startCursor = new Date(monthStart);
  startCursor.setDate(monthStart.getDate() - monthStart.getDay());
  const endCursor = new Date(monthEnd);
  endCursor.setDate(monthEnd.getDate() + (6 - monthEnd.getDay()));

  const days = [];
  const cursor = new Date(startCursor);
  while (cursor <= endCursor) {
    days.push(new Date(cursor));
    cursor.setDate(cursor.getDate() + 1);
  }
  return days;
};

export const META_RANGE_PRESETS = [
  {
    label: "Today",
    getRange: () => {
      const today = getTodayString();
      return { start: today, end: today };
    },
  },
  {
    label: "Yesterday",
    getRange: () => {
      const today = getTodayString();
      const yesterday = addDaysToDateString(today, -1);
      return { start: yesterday, end: yesterday };
    },
  },
  {
    label: "Last 7 days",
    getRange: () => {
      const today = getTodayString();
      return { start: addDaysToDateString(today, -6), end: today };
    },
  },
  {
    label: "Last 14 days",
    getRange: () => {
      const today = getTodayString();
      return { start: addDaysToDateString(today, -13), end: today };
    },
  },
  {
    label: "Last 30 days",
    getRange: () => {
      const today = getTodayString();
      return { start: addDaysToDateString(today, -29), end: today };
    },
  },
  {
    label: "This week",
    getRange: () => {
      const now = new Date();
      return { start: formatDateInput(startOfWeek(now)), end: formatDateInput(now) };
    },
  },
  {
    label: "Last week",
    getRange: () => {
      const now = new Date();
      const thisWeekStart = startOfWeek(now);
      const lastWeekEnd = new Date(thisWeekStart);
      lastWeekEnd.setDate(thisWeekStart.getDate() - 1);
      const lastWeekStart = startOfWeek(lastWeekEnd);
      return { start: formatDateInput(lastWeekStart), end: formatDateInput(lastWeekEnd) };
    },
  },
  {
    label: "This month",
    getRange: () => {
      const now = new Date();
      return { start: formatDateInput(startOfMonth(now)), end: formatDateInput(now) };
    },
  },
  {
    label: "Last month",
    getRange: () => {
      const now = new Date();
      const lastMonth = addMonths(now, -1);
      return { start: formatDateInput(startOfMonth(lastMonth)), end: formatDateInput(endOfMonth(lastMonth)) };
    },
  },
];

const statusPalette = ["#1d5fd0", "#1f8f5f", "#c78322", "#d9485f", "#7c3aed", "#0f766e", "#ea580c", "#475569"];

export function formatTZS(value) {
  return `TSh ${Math.round(Number(value || 0)).toLocaleString()}`;
}

export function formatUSD(value) {
  return `$${Number(value || 0).toFixed(2)}`;
}

export function formatUsdFromTzs(value) {
  return formatUSD(Number(value || 0) / USD_TO_TZS);
}

export function formatMetaLeadSourceLabel(value) {
  const normalized = String(value || "").trim().toLowerCase();
  if (!normalized || normalized === "no_signal") return "No tracking signal";
  if (normalized === "lead_action") return "Lead action";
  if (normalized === "landing_page_view") return "Landing page view";
  if (normalized === "link_click") return "Link click";
  if (normalized === "mixed") return "Mixed lead sources";
  return value;
}

export function formatInteger(value) {
  return Math.round(Number(value || 0)).toLocaleString();
}

export function parseLooseNumber(value) {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  const normalized = String(value || "")
    .replace(/[^\d,.-]/g, "")
    .replace(/,(?=\d{3}\b)/g, "")
    .replace(",", ".");
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : 0;
}

export function normalizeHeaderName(value) {
  return String(value || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

export function getMetaApiBase() {
  const configuredBase = typeof import.meta !== "undefined" ? String(import.meta.env?.VITE_META_API_BASE || "").trim() : "";
  if (configuredBase) return configuredBase.replace(/\/$/, "");
  if (typeof window === "undefined") return `http://127.0.0.1:${META_API_PORT}/api/meta`;
  return `${window.location.protocol}//${window.location.hostname}:${META_API_PORT}/api/meta`;
}

export function getSharedApiBase() {
  const configuredBase = typeof import.meta !== "undefined" ? String(import.meta.env?.VITE_SHARED_API_BASE || "").trim() : "";
  if (configuredBase) return configuredBase.replace(/\/$/, "");
  if (typeof window === "undefined") return `http://127.0.0.1:${SHARED_API_PORT}/api/state`;
  return `${window.location.protocol}//${window.location.hostname}:${SHARED_API_PORT}/api/state`;
}

export function matchProductIdFromText(value, products) {
  const query = normalizeHeaderName(value);
  if (!query) return "";

  const directMatch = products.find(
    (product) =>
      normalizeHeaderName(product.id) === query ||
      normalizeHeaderName(product.name) === query
  );
  if (directMatch) return directMatch.id;

  const queryTokens = query.split(" ").filter(Boolean);
  const scored = products
    .map((product) => {
      const productTokens = normalizeHeaderName(product.name).split(" ").filter(Boolean);
      const intersection = queryTokens.filter((token) => productTokens.includes(token)).length;
      const startsWithBonus = normalizeHeaderName(product.name).startsWith(query) ? 2 : 0;
      return {
        id: product.id,
        score: intersection + startsWithBonus,
      };
    })
    .filter((entry) => entry.score > 0)
    .sort((a, b) => b.score - a.score);

  return scored[0]?.id || "";
}

export function buildMappedMetaRows(rows, products, campaignMappings = {}) {
  return rows.map((row) => {
    const manualMatch = campaignMappings[row.id] || "";
    const autoMatch = manualMatch || matchProductIdFromText(`${row.campaignName} ${row.adsetName || ""}`, products);
    return {
      ...row,
      mappedProductId: autoMatch || "",
      mappedProductName: products.find((product) => product.id === autoMatch)?.name || "",
      autoMatch: !manualMatch && Boolean(autoMatch),
    };
  });
}

export function normalizeOrderStatus(value) {
  const normalized = normalizeHeaderName(value);
  if (!normalized) return "new-order";
  return normalized.replace(/\s+/g, "-");
}

export function normalizeProductOffers(offers) {
  if (!Array.isArray(offers)) return [];
  return offers
    .map((offer) => ({
      minQty: Math.max(2, Number(offer?.minQty || 0)),
      totalPrice: Math.max(0, parseLooseNumber(offer?.totalPrice)),
    }))
    .filter((offer) => offer.minQty >= 2 && offer.totalPrice > 0)
    .sort((a, b) => a.minQty - b.minQty);
}

export function getProductPricing(product, quantity) {
  const qty = Math.max(1, Number(quantity || 1));
  const offers = normalizeProductOffers(product?.offers);
  const matchedOffer = [...offers].reverse().find((offer) => qty >= offer.minQty);
  if (matchedOffer) {
    return {
      unitPrice: matchedOffer.totalPrice / qty,
      totalPrice: matchedOffer.totalPrice,
      offerApplied: matchedOffer,
    };
  }
  const basePrice = Number(product?.sellingPrice || 0);
  return {
    unitPrice: basePrice,
    totalPrice: basePrice * qty,
    offerApplied: null,
  };
}

export function getCustomerOrderTotalTzs(customer, product) {
  if (Number(customer?.orderTotalTzs || 0) > 0) return Number(customer.orderTotalTzs);
  return getProductPricing(product, customer?.quantity).totalPrice;
}

export function formatOffersSummary(offers) {
  const normalized = normalizeProductOffers(offers);
  if (!normalized.length) return "No offers";
  return normalized.map((offer) => `${offer.minQty} pcs = ${formatTZS(offer.totalPrice)}`).join(" | ");
}

export function sanitizeProductRecord(product) {
  return {
    ...product,
    sellingPrice: Math.max(0, parseLooseNumber(product?.sellingPrice)),
    purchaseUnitPrice: Math.max(0, parseLooseNumber(product?.purchaseUnitPrice)),
    totalQty: Math.max(0, Number(product?.totalQty || 0)),
    shippingTotal: Math.max(0, parseLooseNumber(product?.shippingTotal)),
    otherCharges: Math.max(0, parseLooseNumber(product?.otherCharges)),
    delivery: Math.max(0, parseLooseNumber(product?.delivery)),
    estimatedArrivalDays: Math.max(0, Number(product?.estimatedArrivalDays || 0)),
    offers: normalizeProductOffers(product?.offers),
  };
}

export function getDefaultSituationData() {
  return {
    salaries: [],
    fixedCharges: [],
    adInputs: {},
    hourlyAdsSnapshots: [],
    cumulativeAdsTotalTzs: 0,
    cumulativeAdsByProduct: {},
    lastObservedAdsSpendTzs: 0,
    lastObservedAdsByProduct: {},
    lastAdsAccumulatedAt: null,
  };
}

export function sanitizeSituationData(value) {
  const defaults = getDefaultSituationData();
  return {
    salaries: Array.isArray(value?.salaries)
      ? value.salaries.map((entry) => ({
          id: String(entry?.id || `salary-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`),
          name: String(entry?.name || ""),
          amountTzs: Math.max(0, parseLooseNumber(entry?.amountTzs)),
        }))
      : defaults.salaries,
    fixedCharges: Array.isArray(value?.fixedCharges)
      ? value.fixedCharges.map((entry) => ({
          id: String(entry?.id || `fixed-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`),
          label: String(entry?.label || ""),
          amountTzs: Math.max(0, parseLooseNumber(entry?.amountTzs)),
        }))
      : defaults.fixedCharges,
    adInputs:
      value?.adInputs && typeof value.adInputs === "object"
        ? Object.fromEntries(
            Object.entries(value.adInputs).map(([productId, entry]) => [
              productId,
              {
                averageLeadCostTzs: Math.max(0, parseLooseNumber(entry?.averageLeadCostTzs)),
                incomingLeads: Math.max(0, Math.round(parseLooseNumber(entry?.incomingLeads))),
              },
            ])
          )
        : defaults.adInputs,
    hourlyAdsSnapshots: Array.isArray(value?.hourlyAdsSnapshots)
      ? value.hourlyAdsSnapshots
          .map((entry) => ({
            id: String(entry?.id || `ads-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`),
            bucket: String(entry?.bucket || ""),
            amountTzs: Math.max(0, parseLooseNumber(entry?.amountTzs)),
            capturedAt: entry?.capturedAt || null,
            source: String(entry?.source || "tracking"),
          }))
          .sort((a, b) => String(b.bucket || "").localeCompare(String(a.bucket || "")))
      : defaults.hourlyAdsSnapshots,
    cumulativeAdsTotalTzs: Math.max(0, parseLooseNumber(value?.cumulativeAdsTotalTzs)),
    cumulativeAdsByProduct:
      value?.cumulativeAdsByProduct && typeof value.cumulativeAdsByProduct === "object"
        ? Object.fromEntries(
            Object.entries(value.cumulativeAdsByProduct).map(([productId, amount]) => [
              productId,
              Math.max(0, parseLooseNumber(amount)),
            ])
          )
        : defaults.cumulativeAdsByProduct,
    lastObservedAdsSpendTzs: Math.max(0, parseLooseNumber(value?.lastObservedAdsSpendTzs)),
    lastObservedAdsByProduct:
      value?.lastObservedAdsByProduct && typeof value.lastObservedAdsByProduct === "object"
        ? Object.fromEntries(
            Object.entries(value.lastObservedAdsByProduct).map(([productId, amount]) => [
              productId,
              Math.max(0, parseLooseNumber(amount)),
            ])
          )
        : defaults.lastObservedAdsByProduct,
    lastAdsAccumulatedAt: value?.lastAdsAccumulatedAt || null,
  };
}

export function getEmptyExpeditionForm() {
  return {
    name: "",
    source: "china",
    sellingPrice: 0,
    purchaseUnitPrice: 0,
    totalQty: 0,
    shippingTotal: 0,
    otherCharges: 0,
    delivery: 0,
    estimatedArrivalDays: 3,
    supplierName: "",
    supplierContact: "",
    lifecycleStatus: "test",
    defectRate: 0,
    notes: "",
    offers: [],
  };
}

export function getEmptyCustomerForm(productId = "P001") {
  return {
    customerName: "",
    phone: "",
    city: "",
    address: "",
    productId,
    quantity: 1,
    orderDate: getTodayString(),
    paymentMethod: "COD",
    status: "new-order",
    notes: "",
    leadSource: "facebook",
    campaignName: "",
    adsetName: "",
    creativeName: "",
    priority: "normal",
    customerType: "new",
    callAttempts: 0,
    cancelReason: "",
    unreachedReason: "",
    carrierName: "",
    trackingNumber: "",
    expectedDeliveryDate: "",
    returnReason: "",
  };
}

export function getDefaultMetaAdsState() {
  const today = getTodayString();
  return {
    accessToken: "",
    accountId: "",
    dateStart: addDaysToDateString(today, -6),
    dateEnd: today,
    campaignMappings: {},
    autoSync: false,
    autoSyncIntervalMinutes: 5,
    lastSyncAt: null,
    lastSyncSummary: null,
    dailySpendSnapshots: [],
    cumulativeTrackedSpendTzs: 0,
    lifetimeSpendTzs: 0,
    lifetimeSpendCapturedAt: null,
    lastLifetimeSpendSyncDate: null,
  };
}

export function sanitizeServiceForm(value) {
  const totalLeads = Math.max(0, parseLooseNumber(value?.totalLeads));
  const adSpendUsd = Math.max(0, parseLooseNumber(value?.adSpendUsd));

  return {
    totalLeads,
    confirmationRate: Math.max(0, Math.min(100, parseLooseNumber(value?.confirmationRate))),
    deliveryRate: Math.max(0, Math.min(100, parseLooseNumber(value?.deliveryRate))),
    sellingPriceTzs: Math.max(0, parseLooseNumber(value?.sellingPriceTzs)),
    productCostTzs: Math.max(0, parseLooseNumber(value?.productCostTzs)),
    cplUsd: totalLeads > 0 ? adSpendUsd / totalLeads : 0,
    adSpendUsd,
  };
}

export function getDefaultServiceForm() {
  return sanitizeServiceForm({
    totalLeads: 150,
    confirmationRate: 50,
    deliveryRate: 62,
    sellingPriceTzs: 85000,
    productCostTzs: 35000,
    adSpendUsd: 320,
  });
}

export function getDefaultImportMeta() {
  return {
    lastOrdersImportAt: null,
    lastShippingImportAt: null,
  };
}

export function getDefaultCloudWorkspaceState() {
  return {
    products: [],
    tracking: [],
    customers: [],
    serviceForm: getDefaultServiceForm(),
    situationData: getDefaultSituationData(),
    metaAdsState: getDefaultMetaAdsState(),
    importMeta: getDefaultImportMeta(),
  };
}

export function sanitizeMetaAdsState(value) {
  const defaults = getDefaultMetaAdsState();
  const dailySpendSnapshots = Array.isArray(value?.dailySpendSnapshots)
    ? value.dailySpendSnapshots
        .map((entry, index) => {
          const bucket = String(entry?.bucket || entry?.date || "").trim();
          const totalSpendTzs = Math.max(0, parseLooseNumber(entry?.totalSpendTzs ?? entry?.spendTzs));
          const newSpendTzs = Math.max(0, parseLooseNumber(entry?.newSpendTzs ?? entry?.spendTzs));

          return {
            id: String(entry?.id || `meta-snapshot-${bucket || "entry"}-${index}`),
            bucket,
            totalSpendTzs,
            newSpendTzs,
            capturedAt: entry?.capturedAt || null,
            source: String(entry?.source || "meta_maximum"),
          };
        })
        .filter((entry) => entry.bucket)
        .sort((a, b) => {
          const bucketGap = String(b.bucket || "").localeCompare(String(a.bucket || ""));
          if (bucketGap !== 0) return bucketGap;
          return String(b.capturedAt || "").localeCompare(String(a.capturedAt || ""));
        })
    : defaults.dailySpendSnapshots;

  const snapshotCumulativeSpendTzs = dailySpendSnapshots.reduce((sum, entry) => sum + Number(entry.newSpendTzs || 0), 0);
  const fallbackCumulativeSpendTzs =
    snapshotCumulativeSpendTzs > 0
      ? snapshotCumulativeSpendTzs
      : Math.max(0, parseLooseNumber(value?.lifetimeSpendTzs));
  const explicitCumulativeTrackedSpendTzs = Math.max(0, parseLooseNumber(value?.cumulativeTrackedSpendTzs));

  return {
    accessToken: String(value?.accessToken || defaults.accessToken),
    accountId: String(value?.accountId || defaults.accountId),
    dateStart: String(value?.dateStart || defaults.dateStart),
    dateEnd: String(value?.dateEnd || defaults.dateEnd),
    campaignMappings:
      value?.campaignMappings && typeof value.campaignMappings === "object" ? value.campaignMappings : defaults.campaignMappings,
    autoSync: Boolean(value?.autoSync ?? defaults.autoSync),
    autoSyncIntervalMinutes: Math.max(1, Number(value?.autoSyncIntervalMinutes || defaults.autoSyncIntervalMinutes)),
    lastSyncAt: value?.lastSyncAt || null,
    lastSyncSummary: value?.lastSyncSummary || null,
    dailySpendSnapshots,
    cumulativeTrackedSpendTzs: explicitCumulativeTrackedSpendTzs > 0 ? explicitCumulativeTrackedSpendTzs : fallbackCumulativeSpendTzs,
    lifetimeSpendTzs: Math.max(0, parseLooseNumber(value?.lifetimeSpendTzs)),
    lifetimeSpendCapturedAt: value?.lifetimeSpendCapturedAt || null,
    lastLifetimeSpendSyncDate: value?.lastLifetimeSpendSyncDate || null,
  };
}

export const CONFIRMATION_STATUS_RULES = {
  new: { bucket: "new" },
  "new-order": { bucket: "new" },
  confirmed: { bucket: "confirmed" },
  cancelled: { bucket: "cancelled" },
  cancel: { bucket: "cancelled" },
  canceled: { bucket: "cancelled" },
  unreached: { bucket: "pending" },
  unreachable: { bucket: "pending" },
  "not-reachable": { bucket: "pending" },
  "not-joined": { bucket: "pending" },
  "not-joinable": { bucket: "pending" },
  "not-joignable": { bucket: "pending" },
  "non-joignable": { bucket: "pending" },
  "n-est-pas-joignable": { bucket: "pending" },
  "unreachable-text-sent": { bucket: "pending" },
  "no-reply": { bucket: "pending" },
  "no-answer": { bucket: "pending" },
  pending: { bucket: "pending" },
  scheduled: { bucket: "pending" },
  remind: { bucket: "pending" },
  "client-to-revert": { bucket: "pending" },
  reprogrammed: { bucket: "pending" },
  "back-to-stock": { bucket: "cancelled" },
  "out-of-stock": { bucket: "cancelled" },
  "out-of-region": { bucket: "cancelled" },
  wrong: { bucket: "cancelled" },
  "wrong-number": { bucket: "cancelled" },
  spam: { bucket: "cancelled" },
  double: { bucket: "cancelled" },
};

export const SHIPPING_STATUS_RULES = {
  confirmed: { bucket: "to_prepare" },
  "to-prepare": { bucket: "to_prepare" },
  to_prepare: { bucket: "to_prepare" },
  "in-preparation": { bucket: "to_prepare" },
  prepared: { bucket: "to_prepare" },
  "in-delivery": { bucket: "shipped" },
  "out-delivered": { bucket: "shipped" },
  "out-for-delivery": { bucket: "shipped" },
  shipped: { bucket: "shipped" },
  shipping: { bucket: "shipped" },
  "sending-to-agent": { bucket: "shipped" },
  received: { bucket: "delivered" },
  paid: { bucket: "delivered" },
  facture: { bucket: "delivered" },
  factured: { bucket: "delivered" },
  factur: { bucket: "delivered" },
  delivered: { bucket: "delivered" },
  "already-delivered": { bucket: "delivered" },
  return: { bucket: "returned" },
  returned: { bucket: "returned" },
  returning: { bucket: "returned" },
  rejected: { bucket: "returned" },
  refunded: { bucket: "returned" },
  damaged: { bucket: "returned" },
  report: { bucket: "returned" },
  reported: { bucket: "returned" },
  cancelled: { bucket: "returned" },
  canceled: { bucket: "returned" },
  "return-from-agency": { bucket: "returned" },
  "back-to-stock": { bucket: "returned" },
  "returned-to-stock": { bucket: "returned" },
  "return-to-stock": { bucket: "returned" },
  "return-stock": { bucket: "returned" },
  "return-request": { bucket: "returned" },
  "validate-return": { bucket: "returned" },
  "out-of-stock": { bucket: "returned" },
  "on-hold": { bucket: "to_prepare" },
};

export const DEFAULT_CONFIRMATION_STATUSES = [
  "new-order",
  "confirmed",
  "unreached",
  "cancelled",
  "out-of-stock",
  "wrong",
  "remind",
  "client-to-revert",
  "spam",
  "double",
];

export const DEFAULT_POST_CONFIRMATION_STATUSES = [
  "to-prepare",
  "shipped",
  "delivered",
  "returned",
  "refunded",
  "cancelled",
  "out-of-stock",
  "on-hold",
  "return-from-agency",
  "returning",
  "validate-return",
];

const STATUS_COLOR_OVERRIDES = {
  new: "#4f7cf3",
  "new-order": "#4f7cf3",
  confirmed: "#5ad818",
  cancelled: "#ef4444",
  canceled: "#ef4444",
  "back-to-stock": "#f59e0b",
  "out-of-stock": "#ea7a1f",
  "out-of-region": "#fdba74",
  wrong: "#f87171",
  "wrong-number": "#f87171",
  spam: "#111111",
  double: "#111111",
  pending: "#6ad0d5",
  scheduled: "#f5e897",
  remind: "#16a34a",
  "client-to-revert": "#dc43dc",
  reprogrammed: "#8d0a76",
  unreached: "#d020eb",
  unreachable: "#ef4444",
  "no-reply": "#f3c364",
  shipped: "#26b36a",
  "in-delivery": "#7ddff7",
  delivered: "#2e49b9",
  received: "#e8e3e8",
  paid: "#e6e8e8",
  facture: "#4caf50",
  refunded: "#d4cf11",
  return: "#da95d0",
  returned: "#1193ad",
  rejected: "#d4cf11",
  damaged: "#f59e0b",
  reported: "#f4e46e",
  "return-from-agency": "#0f766e",
  "return-stock": "#dfe9c5",
  "return-request": "#ef4444",
  "validate-return": "#41a1f1",
  "on-hold": "#d9e234",
};

export function formatStatusLabel(status) {
  const key = normalizeOrderStatus(status);
  const knownLabels = {
    new: "New Order",
    "new-order": "New Order",
    confirmed: "Confirmed",
    delivered: "Delivered",
    cancelled: "Cancelled",
    canceled: "Cancelled",
    unreached: "Unreached",
    unreachable: "Unreachable",
    "no-reply": "No Reply",
    pending: "Pending",
    remind: "Remind",
    scheduled: "Scheduled",
    shipped: "Shipped",
    "to-prepare": "To Prepare",
    to_prepare: "To Prepare",
    returned: "Returned",
    rejected: "Rejected",
    paid: "Paid",
    facture: "Facture",
    factured: "Factured",
    refunded: "Refunded",
    "out-of-stock": "Out Of Stock",
    "client-to-revert": "Client To Revert",
    "return-request": "Return Request",
    "validate-return": "Validate Return",
    "already-delivered": "Already Delivered",
  };
  if (knownLabels[key]) return knownLabels[key];
  return key
    .split("-")
    .map((segment) => segment.charAt(0).toUpperCase() + segment.slice(1))
    .join(" ");
}

export function getConfirmationStatusRule(status) {
  return CONFIRMATION_STATUS_RULES[normalizeOrderStatus(status)] || null;
}

export function getShippingStatusRule(status) {
  return SHIPPING_STATUS_RULES[normalizeOrderStatus(status)] || null;
}

export function getConfirmationBucket(status) {
  return getConfirmationStatusRule(status)?.bucket || "pending";
}

export function getShippingBucket(status) {
  return getShippingStatusRule(status)?.bucket || "to_prepare";
}

export function getStatusSemantic(status) {
  const confirmationRule = getConfirmationStatusRule(status);
  if (confirmationRule) return confirmationRule.bucket;
  const shippingRule = getShippingStatusRule(status);
  if (shippingRule) return shippingRule.bucket;
  return "pending";
}

export function getStatusColor(status) {
  const key = normalizeOrderStatus(status);
  if (STATUS_COLOR_OVERRIDES[key]) return STATUS_COLOR_OVERRIDES[key];

  const semantic = getStatusSemantic(status);
  if (semantic === "delivered") return "#16a34a";
  if (semantic === "confirmed") return "#f59e0b";
  if (semantic === "cancelled") return "#dc2626";
  if (semantic === "returned") return "#0f766e";
  if (semantic === "shipped") return "#0ea5e9";
  if (semantic === "to_prepare") return "#f59e0b";
  if (semantic === "new") return "#4f7cf3";

  const hash = key.split("").reduce((sum, char) => sum + char.charCodeAt(0), 0);
  return statusPalette[hash % statusPalette.length];
}

export function getStatusBadgeStyle(status) {
  const color = getStatusColor(status);
  return {
    display: "inline-flex",
    alignItems: "center",
    gap: 6,
    padding: "7px 11px",
    borderRadius: 999,
    fontWeight: 800,
    fontSize: 11,
    letterSpacing: 0.22,
    color,
    background: `${color}14`,
    border: `1px solid ${color}30`,
  };
}

export function excelDateToInput(value) {
  if (value == null || value === "") return getTodayString();
  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (!parsed) return getTodayString();
    return formatDateInput(new Date(parsed.y, parsed.m - 1, parsed.d));
  }

  const raw = String(value).trim();
  if (!raw) return getTodayString();
  const isoMatch = raw.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (isoMatch) {
    return `${isoMatch[1]}-${String(isoMatch[2]).padStart(2, "0")}-${String(isoMatch[3]).padStart(2, "0")}`;
  }

  const frMatch = raw.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
  if (frMatch) {
    return `${frMatch[3]}-${String(frMatch[2]).padStart(2, "0")}-${String(frMatch[1]).padStart(2, "0")}`;
  }

  const compactIsoMatch = raw.match(/^(\d{4})(\d{2})(\d{2})(?:\D|$)/);
  if (compactIsoMatch) {
    return `${compactIsoMatch[1]}-${compactIsoMatch[2]}-${compactIsoMatch[3]}`;
  }

  const fallback = new Date(raw);
  return Number.isNaN(fallback.getTime()) ? getTodayString() : formatDateInput(fallback);
}

export function ensureShippingStatusForConfirmed(confirmationStatus, shippingStatus) {
  const normalizedShipping = normalizeOrderStatus(shippingStatus);
  if (normalizedShipping) return normalizedShipping;
  return isConfirmationConfirmed(confirmationStatus) ? "to-prepare" : "";
}

export function buildNextId(items, prefix) {
  const maxId = items.reduce((max, item) => {
    const match = String(item?.id || "").match(new RegExp(`^${prefix}(\\d+)$`));
    return match ? Math.max(max, Number(match[1])) : max;
  }, 0);

  return `${prefix}${String(maxId + 1).padStart(3, "0")}`;
}

export function buildHistoryEntry({ action, source = "system", details = "" }) {
  return {
    id: `hist-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
    at: new Date().toISOString(),
    action: String(action || "updated"),
    source: String(source || "system"),
    details: String(details || ""),
  };
}

export function sanitizeHistoryEntries(history) {
  if (!Array.isArray(history)) return [];
  return history
    .map((entry) => ({
      id: String(entry?.id || `hist-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`),
      at: entry?.at || new Date().toISOString(),
      action: String(entry?.action || "updated"),
      source: String(entry?.source || "system"),
      details: String(entry?.details || ""),
    }))
    .sort((a, b) => String(b.at).localeCompare(String(a.at)));
}

export function appendCustomerHistory(customer, entry) {
  return [entry, ...sanitizeHistoryEntries(customer?.history)].slice(0, 24);
}

export function getWeekStartString(value) {
  const date = parseDateInput(value) || new Date(value || Date.now());
  if (Number.isNaN(date.getTime())) return getTodayString();
  return formatDateInput(startOfWeek(date));
}

export function getWeekEndString(weekStartString) {
  return addDaysToDateString(weekStartString, 6);
}

export function getWeekLabel(weekStartString) {
  return `${weekStartString} -> ${getWeekEndString(weekStartString)}`;
}

export function getDayBucket(value = Date.now()) {
  return new Date(value).toISOString().slice(0, 13);
}

export function hasMeaningfulWorkspaceData(snapshot = {}) {
  return Boolean(
    (Array.isArray(snapshot.products) && snapshot.products.length > 0) ||
    (Array.isArray(snapshot.customers) && snapshot.customers.length > 0) ||
    (Array.isArray(snapshot.tracking) && snapshot.tracking.length > 0)
  );
}

export function sanitizeCustomerRecord(customer) {
  const separated = inferSeparatedStatuses(customer);
  return {
    ...customer,
    id: String(customer?.id || `C${Math.random().toString(36).slice(2, 6).toUpperCase()}`),
    customerName: String(customer?.customerName || ""),
    phone: String(customer?.phone || ""),
    city: String(customer?.city || ""),
    address: String(customer?.address || ""),
    productId: String(customer?.productId || ""),
    quantity: Math.max(1, Number(customer?.quantity || 1)),
    orderDate: String(customer?.orderDate || getTodayString()),
    paymentMethod: String(customer?.paymentMethod || "COD"),
    notes: String(customer?.notes || ""),
    leadSource: String(customer?.leadSource || "facebook"),
    campaignName: String(customer?.campaignName || ""),
    adsetName: String(customer?.adsetName || ""),
    creativeName: String(customer?.creativeName || ""),
    priority: String(customer?.priority || "normal"),
    customerType: String(customer?.customerType || "new"),
    callAttempts: Math.max(0, Number(customer?.callAttempts || 0)),
    cancelReason: String(customer?.cancelReason || ""),
    unreachedReason: String(customer?.unreachedReason || ""),
    carrierName: String(customer?.carrierName || ""),
    trackingNumber: String(customer?.trackingNumber || ""),
    expectedDeliveryDate: String(customer?.expectedDeliveryDate || ""),
    actualDeliveryDate: String(customer?.actualDeliveryDate || ""),
    returnReason: String(customer?.returnReason || ""),
    orderTotalTzs: Math.max(0, parseLooseNumber(customer?.orderTotalTzs)),
    sourceOrderId: customer?.sourceOrderId || null,
    importSource: customer?.importSource || null,
    lastImportedAt: customer?.lastImportedAt || null,
    lastShippingImportedAt: customer?.lastShippingImportedAt || null,
    assignedTo: String(customer?.assignedTo || ""),
    history: sanitizeHistoryEntries(customer?.history),
    confirmationStatus: separated.confirmationStatus,
    shippingStatus: separated.shippingStatus,
    status: separated.effectiveStatus,
  };
}

export function normalizePhoneValue(value) {
  return String(value || "").replace(/\D+/g, "");
}

export function inferSeparatedStatuses(customer) {
  const rawStatus = normalizeOrderStatus(customer?.status || customer?.confirmationStatus || customer?.shippingStatus);
  const explicitConfirmation = normalizeOrderStatus(customer?.confirmationStatus);
  const explicitShipping = normalizeOrderStatus(customer?.shippingStatus);
  const hasShippingImport = Boolean(customer?.lastShippingImportedAt);
  const isKnownShipping = Boolean(getShippingStatusRule(rawStatus));
  const isKnownConfirmation = Boolean(getConfirmationStatusRule(rawStatus));

  let confirmationStatus = explicitConfirmation || "";
  let shippingStatus = explicitShipping || "";
  const explicitConfirmationIsShipping =
    Boolean(confirmationStatus) &&
    getShippingStatusRule(confirmationStatus) &&
    !getConfirmationStatusRule(confirmationStatus);

  if (explicitConfirmationIsShipping && !shippingStatus) {
    shippingStatus = confirmationStatus;
    confirmationStatus = "confirmed";
  }

  if (!confirmationStatus && !shippingStatus) {
    if (isKnownShipping && !isKnownConfirmation) {
      shippingStatus = rawStatus;
      confirmationStatus = "confirmed";
    } else if (isKnownShipping && hasShippingImport && rawStatus !== "confirmed") {
      shippingStatus = rawStatus;
      confirmationStatus = "confirmed";
    } else {
      confirmationStatus = rawStatus || "new";
    }
  }

  if (!shippingStatus && hasShippingImport && rawStatus && isKnownShipping) {
    shippingStatus = rawStatus;
  }

  if (!confirmationStatus && shippingStatus) {
    confirmationStatus = "confirmed";
  }

  if (!confirmationStatus) confirmationStatus = "new";
  if (confirmationStatus && !shippingStatus && isKnownShipping && confirmationStatus === "confirmed" && rawStatus !== "confirmed") {
    shippingStatus = rawStatus;
  }

  return {
    confirmationStatus,
    shippingStatus,
    effectiveStatus: shippingStatus || confirmationStatus,
  };
}

export function getCustomerConfirmationStatus(customer) {
  return inferSeparatedStatuses(customer).confirmationStatus;
}

export function getCustomerShippingStatus(customer) {
  return inferSeparatedStatuses(customer).shippingStatus;
}

export function getCustomerEffectiveStatus(customer) {
  return inferSeparatedStatuses(customer).effectiveStatus;
}

export function isConfirmationConfirmed(status) {
  return getConfirmationBucket(status) === "confirmed";
}

export function isConfirmationNew(status) {
  return getConfirmationBucket(status) === "new";
}

export function isConfirmationCancelled(status) {
  return getConfirmationBucket(status) === "cancelled";
}

export function isShippingDelivered(status) {
  return getShippingBucket(status) === "delivered";
}

export function isShippingInProgress(status) {
  const bucket = getShippingBucket(status);
  return bucket === "to_prepare" || bucket === "shipped";
}

export function isShippingReturned(status) {
  return getShippingBucket(status) === "returned";
}

export function getUnitProductCostUSD(product) {
  if (!product) return 0;
  const qty = Number(product.totalQty || 0);
  const purchaseUnitPrice = Number(product.purchaseUnitPrice || 0);
  const shippingTotalUsd = Number(product.shippingTotal || 0) / USD_TO_TZS;
  const otherChargesUsd = Number(product.otherCharges || 0) / USD_TO_TZS;
  const totalImportCost = purchaseUnitPrice * qty + shippingTotalUsd + otherChargesUsd;
  return qty > 0 ? totalImportCost / qty : 0;
}
