const stripAccents = (value) =>
  String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");

const compactSpaces = (value) => value.replace(/\s+/g, " ").trim();

export function normalizeStatus(status) {
  const normalized = compactSpaces(
    stripAccents(String(status || "").toLowerCase())
      .replace(/\(\s*\d+\s*\)/g, "")
      .replace(/[_/-]+/g, " ")
      .replace(/[^\p{L}\p{N}\s]/gu, " ")
  );

  if (!normalized) return "";
  return normalized;
}

const CONFIRMED_STATUSES = new Set([
  "confirmed",
  "confirmee",
  "confirme",
  "commande confirmee",
  "confirmed order",
]);

const NO_REPLY_STATUSES = new Set([
  "unreached",
  "no reply",
  "no answer",
  "not answering",
  "scheduled",
  "pas de reponse",
  "ne repond pas",
]);

const CANCELLED_STATUSES = new Set([
  "cancelled",
  "canceled",
  "annule",
  "annulation",
  "client cancelled",
  "rejected",
  "refused confirmation",
]);

const INVALID_LEAD_STATUSES = new Set([
  "wrong",
  "out of region",
  "duplicate",
  "double",
]);

const STOCK_HOLD_STATUSES = new Set([
  "waiting stock",
  "back to stock",
]);

const DELIVERED_STATUSES = new Set([
  "facture",
  "facturee",
  "facture payee",
  "facture paye",
  "facturee payee",
  "delivered",
  "livre",
  "paid",
  "completed",
  "delivered paid",
]);

const RETURNED_STATUSES = new Set([
  "return stock",
  "return from agency",
  "return",
  "returned",
  "retour",
  "rejected",
  "refused",
  "client refused",
  "failed delivery",
  "undelivered",
  "not delivered",
  "cancelled after shipping",
  "damaged",
  "reported",
]);

const PENDING_SHIPPING_STATUSES = new Set([
  "in delivery",
  "sending to agent",
  "out of stock",
  "in preparation",
  "printed",
  "out for delivery",
  "shipping",
  "shipped",
  "in transit",
  "at agency",
  "with courier",
  "en livraison",
  "expedie",
  "chez livreur",
  "to prepare",
  "to prepare now",
  "prepared",
  "en preparation",
  "pret",
  "en cours",
  "being prepared",
]);

export function isBlankShippingStatus(status) {
  return !normalizeStatus(status);
}

export function isConfirmedStatus(status) {
  const normalized = normalizeStatus(status);
  return CONFIRMED_STATUSES.has(normalized);
}

export function isNoReplyStatus(status) {
  const normalized = normalizeStatus(status);
  return NO_REPLY_STATUSES.has(normalized);
}

export function isCancelledStatus(status) {
  const normalized = normalizeStatus(status);
  return CANCELLED_STATUSES.has(normalized);
}

export function isInvalidLeadStatus(status) {
  const normalized = normalizeStatus(status);
  return INVALID_LEAD_STATUSES.has(normalized);
}

export function isStockHoldStatus(status) {
  const normalized = normalizeStatus(status);
  return STOCK_HOLD_STATUSES.has(normalized);
}

export function isDeliveredStatus(status) {
  const normalized = normalizeStatus(status);
  return DELIVERED_STATUSES.has(normalized);
}

export function isReturnedStatus(status) {
  const normalized = normalizeStatus(status);
  return RETURNED_STATUSES.has(normalized);
}

export function isPendingShippingStatus(status) {
  const normalized = normalizeStatus(status);
  return PENDING_SHIPPING_STATUSES.has(normalized);
}

export function getConfirmationStatusBucket(status) {
  if (isConfirmedStatus(status)) return "confirmed";
  if (isCancelledStatus(status)) return "cancelled";
  if (isInvalidLeadStatus(status)) return "invalid";
  if (isStockHoldStatus(status)) return "stock_hold";
  if (isNoReplyStatus(status)) return "pending";
  return "new";
}

export function getShippingStatusBucket(status) {
  if (isBlankShippingStatus(status)) return "blank";
  if (isDeliveredStatus(status)) return "delivered";
  if (isReturnedStatus(status)) return "returned";
  if (isPendingShippingStatus(status)) return "pending";
  return "blank";
}
