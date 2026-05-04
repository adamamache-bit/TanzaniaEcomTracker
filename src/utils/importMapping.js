import { DEFAULT_EXCHANGE_RATE, tshToUsd } from "./currency";
import { normalizeStatus } from "../config/statusMapping";

function normalizeHeaderName(value) {
  return String(value || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
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

export function getRowValue(row, aliases = []) {
  const entries = Object.entries(row || {});
  for (const [key, value] of entries) {
    if (aliases.includes(normalizeHeaderName(key))) return value;
  }
  return "";
}

function splitMultiValue(value) {
  return String(value || "")
    .split(",")
    .map((entry) => entry.trim())
    .filter(Boolean);
}

export function parseImportedExcelRows(rows = [], { exchangeRate = DEFAULT_EXCHANGE_RATE, resolveProductId, resolveProductRef } = {}) {
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
    const orderId = String(getRowValue(row, ["code"])).trim();
    const clientName = String(getRowValue(row, ["recipient", "customer", "customer name", "full name", "name"])).trim();
    const address = String(getRowValue(row, ["address"])).trim();
    const city = String(getRowValue(row, ["city"])).trim();
    const phone = String(getRowValue(row, ["phone"])).trim();
    const normalizedPhone = normalizePhoneNumber(phone);
    const amountTsh = Math.max(0, parseLooseNumber(getRowValue(row, ["amount", "total", "total amount", "montant", "price total"])));
    const productNames = splitMultiValue(getRowValue(row, ["product name", "product", "produit", "item name"]));
    const productRefs = splitMultiValue(getRowValue(row, ["product ref", "product id", "sku", "reference produit"]));
    const quantityParts = splitMultiValue(getRowValue(row, ["product qt", "quantity", "qty", "quantite"]));
    const createdAt = String(getRowValue(row, ["created at", "order date", "date"])).trim();
    const confirmationStatusRaw = String(getRowValue(row, ["conf status", "conf.status", "confirmation status", "status confirmation"])).trim();
    const confirmationUpdatedAt = String(getRowValue(row, ["conf updated at", "conf.updated at", "confirmation updated at"])).trim();
    const shippingStatusRaw = String(getRowValue(row, ["shipping status", "delivery status", "shipment status"])).trim();
    const updatedAt = String(getRowValue(row, ["updated at"])).trim();
    const extraFields = Object.fromEntries(
      Object.entries(row || {}).filter(([key]) => ![
        "code", "recipient", "customer", "customer name", "full name", "name", "address", "city", "phone", "amount", "total", "total amount",
        "montant", "price total", "product name", "product", "produit", "item name", "product ref", "product id", "sku", "reference produit",
        "product qt", "quantity", "qty", "quantite", "created at", "order date", "date", "conf status", "conf.status", "confirmation status",
        "status confirmation", "conf updated at", "conf.updated at", "confirmation updated at", "shipping status", "delivery status", "shipment status",
        "updated at"
      ].includes(normalizeHeaderName(key)))
    );

    if (!orderId) report.missingCodeRows += 1;
    if (!normalizedPhone) report.missingPhoneRows += 1;
    if (amountTsh <= 0) report.missingAmountRows += 1;
    if (!productNames.length) report.missingProductRows += 1;

    const lineCount = Math.max(productNames.length || 1, quantityParts.length || 1, productRefs.length || 1);
    const quantityNumbers = Array.from({ length: lineCount }, (_, index) =>
      Math.max(1, Math.round(parseLooseNumber(quantityParts[index] || quantityParts[0] || 1)))
    );

    const hasMultipleProducts = lineCount > 1;
    let allocationWeights = quantityNumbers.map((qty, index) => {
      if (typeof resolveProductId !== "function") return qty;
      const resolvedId = resolveProductId(productRefs[index] || productNames[index] || "");
      const resolvedRef = typeof resolveProductRef === "function" ? resolveProductRef(resolvedId) : null;
      const sellingPriceTsh = Math.max(0, parseLooseNumber(resolvedRef?.sellingPrice));
      return sellingPriceTsh > 0 ? sellingPriceTsh * qty : qty;
    });
    const totalWeight = allocationWeights.reduce((sum, value) => sum + value, 0) || 1;

    for (let lineIndex = 0; lineIndex < lineCount; lineIndex += 1) {
      const productName = productNames[lineIndex] || productNames[0] || "";
      const productRef = productRefs[lineIndex] || productRefs[0] || "";
      const quantity = quantityNumbers[lineIndex];
      const allocatedRevenueTsh = amountTsh > 0 ? (amountTsh * allocationWeights[lineIndex]) / totalWeight : 0;
      const resolvedProductId = typeof resolveProductId === "function" ? resolveProductId(productRef || productName) : "";
      const productIdentifier = normalizeHeaderName(productRef || productName) || "unknown-product";
      const fallbackKey = `${normalizedPhone}::${productIdentifier}`;
      const importKey = orderId ? `${orderId}::${lineIndex}` : `${fallbackKey}::${lineIndex}`;

      parsedRows.push({
        order_id: orderId || "",
        import_key: importKey,
        client_name: clientName,
        address,
        city,
        phone,
        normalized_phone: normalizedPhone,
        amount_tsh: allocatedRevenueTsh,
        amount_usd: tshToUsd(allocatedRevenueTsh, exchangeRate),
        product_name: productName,
        product_ref: productRef,
        quantity,
        created_at: createdAt,
        confirmation_status_raw: confirmationStatusRaw,
        confirmation_status_normalized: normalizeStatus(confirmationStatusRaw),
        confirmation_updated_at: confirmationUpdatedAt,
        shipping_status_raw: shippingStatusRaw,
        shipping_status_normalized: normalizeStatus(shippingStatusRaw),
        updated_at: updatedAt,
        raw_row_data: row,
        extra_fields: extraFields,
        productId: resolvedProductId,
        multi_product_revenue_allocated: hasMultipleProducts && amountTsh > 0,
        line_item_index: lineIndex,
      });
    }
  });

  return { parsedRows, report };
}
