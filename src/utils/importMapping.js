import { DEFAULT_EXCHANGE_RATE, tshToUsd } from "./currency";
import { normalizeStatus } from "../config/statusMapping";

function normalizeHeaderName(value) {
  return String(value || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function parseLooseNumber(value) {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  const normalized = String(value || "")
    .replace(/[^\d,.-]/g, "")
    .replace(/,(?=\d{3}\b)/g, "")
    .replace(",", ".");
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : 0;
}

export function normalizePhoneNumber(value) {
  return String(value || "").replace(/\D+/g, "");
}

// Priority-aware alias lookup: tries aliases in order, returns first match
export function getRowValue(row, aliases = []) {
  const entries = Object.entries(row || {});
  for (const alias of aliases) {
    const found = entries.find(([key]) => normalizeHeaderName(key) === alias);
    if (found !== undefined) return found[1];
  }
  return "";
}

function splitMultiValue(value) {
  return String(value || "")
    .split(",")
    .map((entry) => entry.trim())
    .filter(Boolean);
}

/**
 * Detect which Excel format the file uses from its header row.
 *
 * Format 1 (COD platform): CODE, AMOUNT, CONF.STATUS, SHIPPING STATUS, RECIPIENT, PRODUCT NAME, PRODUCT QT
 * Format 2 (new platform): Id, Total price, Confirmation status, Delivery status, Full name, Quantity, Order date
 */
export function detectExcelFormat(rows) {
  if (!rows || !rows.length) return { format: "format1", label: "Format 1 (CODE/AMOUNT)" };
  const keys = Object.keys(rows[0]).map(normalizeHeaderName);

  // Strong Format 2 signals
  if (keys.includes("total price"))    return { format: "format2", label: "Format 2 (Id/Total price)" };
  if (keys.includes("delivery status")) return { format: "format2", label: "Format 2 (Id/Total price)" };
  if (keys.includes("livreur"))        return { format: "format2", label: "Format 2 (Id/Total price)" };
  if (keys.includes("teleconsultant")) return { format: "format2", label: "Format 2 (Id/Total price)" };

  // Strong Format 1 signals
  if (keys.includes("amount"))         return { format: "format1", label: "Format 1 (CODE/AMOUNT)" };
  if (keys.includes("conf.status") || keys.includes("conf status")) return { format: "format1", label: "Format 1 (CODE/AMOUNT)" };
  if (keys.includes("recipient"))      return { format: "format1", label: "Format 1 (CODE/AMOUNT)" };

  // Fallback: Format 2 uses "id" without "code"
  if (keys.includes("id") && !keys.includes("code")) return { format: "format2", label: "Format 2 (Id/Total price)" };

  return { format: "format1", label: "Format 1 (CODE/AMOUNT)" };
}

// Format 2 "Delivery status: cancelled" maps to "cancelled after shipping" in our system
// (different from a confirmation "cancelled" which is pre-shipment)
function mapFormat2ShippingStatus(raw) {
  const v = normalizeStatus(raw);
  if (v === "cancelled" || v === "canceled") return "cancelled after shipping";
  return raw;
}

// All normalized header names that are "known" and should NOT end up in extra_fields
const KNOWN_HEADERS = new Set([
  // Format 1
  "code", "recipient", "address", "city", "phone", "amount", "product name", "product",
  "produit", "item name", "product ref", "product id", "sku", "reference produit",
  "product qt", "quantite", "qte", "created at", "date",
  "conf status", "conf.status", "status confirmation",
  "conf updated at", "conf.updated at", "confirmation updated at",
  "shipping status", "shipment status", "updated at",
  // Format 2
  "id", "order id", "order date", "full name", "name", "customer", "customer name",
  "quantity", "qty", "total price", "price total", "total amount", "total", "montant",
  "confirmation status", "delivery status", "teleconsultant", "livreur", "note", "notes",
  "utm source", "utm source", "external order id",
  // Shared
  "orderid", "order no", "reference", "ref",
]);

export function parseImportedExcelRows(rows = [], { exchangeRate = DEFAULT_EXCHANGE_RATE, resolveProductId, resolveProductRef } = {}) {
  const { format, label: formatLabel } = detectExcelFormat(rows);
  const isFmt2 = format === "format2";

  // Format-aware primary aliases (detected format's column listed first)
  const orderIdAliases = isFmt2
    ? ["id", "code", "order id", "orderid", "order no", "reference", "ref"]
    : ["code", "order id", "orderid", "order no", "reference", "ref", "id"];

  const amountAliases = isFmt2
    ? ["total price", "price total", "total amount", "total", "amount", "montant"]
    : ["amount", "total price", "total", "total amount", "montant", "price total"];

  const clientNameAliases = isFmt2
    ? ["full name", "name", "recipient", "customer", "customer name"]
    : ["recipient", "full name", "customer", "customer name", "name"];

  const confirmationAliases = ["confirmation status", "conf status", "conf.status", "status confirmation"];
  const shippingAliases = isFmt2
    ? ["delivery status", "shipping status", "shipment status"]
    : ["shipping status", "delivery status", "shipment status"];

  const quantityAliases = ["quantity", "qty", "product qt", "quantite", "qte"];
  const createdAtAliases = ["order date", "created at", "date"];

  const report = {
    totalRowsImported: rows.length,
    newLeadsAdded: 0,
    existingLeadsUpdated: 0,
    duplicatesSkipped: 0,
    statusChangesDetected: 0,
    unknownConfirmationStatuses: 0,
    unknownShippingStatuses: 0,
    missingCodeRows: 0,
    missingPhoneRows: 0,
    missingAmountRows: 0,
    missingProductRows: 0,
  };

  const parsedRows = [];

  rows.forEach((row) => {
    const orderId          = String(getRowValue(row, orderIdAliases)).trim();
    const clientName       = String(getRowValue(row, clientNameAliases)).trim();
    const address          = String(getRowValue(row, ["address"])).trim();
    const city             = String(getRowValue(row, ["city"])).trim();
    const phone            = String(getRowValue(row, ["phone"])).trim();
    const normalizedPhone  = normalizePhoneNumber(phone);
    const amountTsh        = Math.max(0, parseLooseNumber(getRowValue(row, amountAliases)));
    const productNames     = splitMultiValue(getRowValue(row, ["product name", "product", "produit", "item name"]));
    const productRefs      = splitMultiValue(getRowValue(row, ["product ref", "product id", "sku", "reference produit"]));
    const quantityParts    = splitMultiValue(getRowValue(row, quantityAliases));
    const createdAt        = String(getRowValue(row, createdAtAliases)).trim();
    const confirmationStatusRaw  = String(getRowValue(row, confirmationAliases)).trim();
    const confirmationUpdatedAt  = String(getRowValue(row, ["conf updated at", "conf.updated at", "confirmation updated at"])).trim();
    const shippingStatusRaw = (() => {
      const raw = String(getRowValue(row, shippingAliases)).trim();
      return isFmt2 ? mapFormat2ShippingStatus(raw) : raw;
    })();
    const updatedAt = String(getRowValue(row, ["updated at"])).trim();

    // Format 2 extra fields
    const assignedTo       = isFmt2 ? String(getRowValue(row, ["teleconsultant"])).trim() : "";
    const deliveryAgent    = isFmt2 ? String(getRowValue(row, ["livreur"])).trim() : "";
    const notesRaw         = isFmt2 ? String(getRowValue(row, ["note", "notes"])).trim() : "";
    const utmSource        = isFmt2 ? String(getRowValue(row, ["utm source"])).trim() : "";
    const externalOrderId  = isFmt2 ? String(getRowValue(row, ["order id"])).trim() : "";

    const extraFields = Object.fromEntries(
      Object.entries(row || {}).filter(([key]) => !KNOWN_HEADERS.has(normalizeHeaderName(key)))
    );

    if (!orderId)          report.missingCodeRows    += 1;
    if (!normalizedPhone)  report.missingPhoneRows   += 1;
    if (amountTsh <= 0)    report.missingAmountRows  += 1;
    if (!productNames.length) report.missingProductRows += 1;

    const lineCount = Math.max(productNames.length || 1, quantityParts.length || 1, productRefs.length || 1);
    const quantityNumbers = Array.from({ length: lineCount }, (_, index) =>
      Math.max(1, Math.round(parseLooseNumber(quantityParts[index] || quantityParts[0] || 1)))
    );

    const hasMultipleProducts = lineCount > 1;
    const allocationWeights = quantityNumbers.map((qty, index) => {
      if (typeof resolveProductId !== "function") return qty;
      const resolvedId = resolveProductId(productRefs[index] || productNames[index] || "");
      const resolvedRef = typeof resolveProductRef === "function" ? resolveProductRef(resolvedId) : null;
      const sellingPriceTsh = Math.max(0, parseLooseNumber(resolvedRef?.sellingPrice));
      return sellingPriceTsh > 0 ? sellingPriceTsh * qty : qty;
    });
    const totalWeight = allocationWeights.reduce((sum, v) => sum + v, 0) || 1;

    for (let lineIndex = 0; lineIndex < lineCount; lineIndex += 1) {
      const productName       = productNames[lineIndex] || productNames[0] || "";
      const productRef        = productRefs[lineIndex]  || productRefs[0]  || "";
      const quantity          = quantityNumbers[lineIndex];
      const allocatedRevenue  = amountTsh > 0 ? (amountTsh * allocationWeights[lineIndex]) / totalWeight : 0;
      const resolvedProductId = typeof resolveProductId === "function" ? resolveProductId(productRef || productName) : "";
      const productIdentifier = normalizeHeaderName(productRef || productName) || "unknown-product";
      const fallbackKey       = `${normalizedPhone}::${productIdentifier}`;
      const importKey         = orderId ? `${orderId}::${lineIndex}` : `${fallbackKey}::${lineIndex}`;

      parsedRows.push({
        order_id:                    orderId || "",
        external_order_id:           externalOrderId || "",
        import_key:                  importKey,
        client_name:                 clientName,
        address,
        city,
        phone,
        normalized_phone:            normalizedPhone,
        amount_tsh:                  allocatedRevenue,
        amount_usd:                  tshToUsd(allocatedRevenue, exchangeRate),
        product_name:                productName,
        product_ref:                 productRef,
        quantity,
        created_at:                  createdAt,
        confirmation_status_raw:     confirmationStatusRaw,
        confirmation_status_normalized: normalizeStatus(confirmationStatusRaw),
        confirmation_updated_at:     confirmationUpdatedAt,
        shipping_status_raw:         shippingStatusRaw,
        shipping_status_normalized:  normalizeStatus(shippingStatusRaw),
        updated_at:                  updatedAt,
        // Format 2 specific
        assignedTo:                  assignedTo,
        delivery_agent:              deliveryAgent,
        notes_raw:                   notesRaw,
        utm_source:                  utmSource,
        // Meta
        raw_row_data:                row,
        extra_fields:                extraFields,
        productId:                   resolvedProductId,
        multi_product_revenue_allocated: hasMultipleProducts && amountTsh > 0,
        line_item_index:             lineIndex,
        detected_format:             format,
      });
    }
  });

  return { parsedRows, report, format, formatLabel };
}
