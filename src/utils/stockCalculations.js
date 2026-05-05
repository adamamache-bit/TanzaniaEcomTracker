import { DEFAULT_EXCHANGE_RATE, tshToUsd, usdToTsh } from "./currency";
import { isConfirmedStatus, isDeliveredStatus, isReturnedStatus } from "../config/statusMapping";

function toNumber(value) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

function getNormalizedProductId(value) {
  return String(value || "").trim();
}

function getOrderQuantity(order) {
  return Math.max(1, Math.round(toNumber(order?.quantity || order?.product_qt || 1)));
}

export function calculateLandedCostPerUnit(stockPurchase = {}, exchangeRate = DEFAULT_EXCHANGE_RATE) {
  const quantityOrdered = Math.max(0, toNumber(stockPurchase.quantity_ordered ?? stockPurchase.totalQty ?? stockPurchase.quantityReceived ?? stockPurchase.quantity_received));
  if (quantityOrdered <= 0) return 0;
  const totalBuyCostUsd = quantityOrdered * Math.max(0, toNumber(stockPurchase.buy_price_per_unit_usd ?? stockPurchase.purchaseUnitPrice));
  const shippingCostUsd = Math.max(0, toNumber(stockPurchase.shipping_cost_usd ?? tshToUsd(stockPurchase.shippingTotal, exchangeRate)));
  const sourcingCostUsd = Math.max(0, toNumber(stockPurchase.sourcing_cost_usd));
  const otherChargesUsd = Math.max(
    0,
    toNumber(stockPurchase.other_charges_usd ?? tshToUsd(stockPurchase.other_charges_tsh ?? stockPurchase.otherCharges, exchangeRate))
  );

  return (totalBuyCostUsd + shippingCostUsd + sourcingCostUsd + otherChargesUsd) / quantityOrdered;
}

function getReceivedStockUnits(productId, stockPurchases = []) {
  const normalizedProductId = getNormalizedProductId(productId);
  const relevant = stockPurchases.filter((purchase) => getNormalizedProductId(purchase?.product_id ?? purchase?.id) === normalizedProductId);
  if (!relevant.length) return 0;

  return relevant.reduce((sum, purchase) => {
    const status = String(purchase?.status || purchase?.stockArrivalStatus || "").trim().toLowerCase();
    if (!["arrived", "received", "partially_received"].includes(status)) return sum;
    return sum + Math.max(0, toNumber(purchase.quantity_received ?? purchase.quantityReceived ?? purchase.totalQty));
  }, 0);
}

export function calculateReservedStock(productId, orders = []) {
  const normalizedProductId = getNormalizedProductId(productId);
  return orders.reduce((sum, order) => {
    if (getNormalizedProductId(order?.productId) !== normalizedProductId) return sum;
    const confirmationStatus = order?.confirmationStatusRaw ?? order?.confirmationStatus ?? order?.status;
    const shippingStatus = order?.shippingStatusRaw ?? order?.shippingStatus;
    if (!isConfirmedStatus(confirmationStatus)) return sum;
    if (isDeliveredStatus(shippingStatus) || isReturnedStatus(shippingStatus)) return sum;
    return sum + getOrderQuantity(order);
  }, 0);
}

export function calculateDeliveredStock(productId, orders = []) {
  const normalizedProductId = getNormalizedProductId(productId);
  return orders.reduce((sum, order) => {
    if (getNormalizedProductId(order?.productId) !== normalizedProductId) return sum;
    return isDeliveredStatus(order?.shippingStatusRaw ?? order?.shippingStatus)
      ? sum + getOrderQuantity(order)
      : sum;
  }, 0);
}

export function calculateReturnedStock(productId, orders = []) {
  const normalizedProductId = getNormalizedProductId(productId);
  return orders.reduce((sum, order) => {
    if (getNormalizedProductId(order?.productId) !== normalizedProductId) return sum;
    return isReturnedStatus(order?.shippingStatusRaw ?? order?.shippingStatus)
      ? sum + getOrderQuantity(order)
      : sum;
  }, 0);
}

export function calculateAvailableStock(productId, stockPurchases = [], orders = []) {
  const receivedUnits = getReceivedStockUnits(productId, stockPurchases);
  const deliveredUnits = calculateDeliveredStock(productId, orders);
  const reservedUnits = calculateReservedStock(productId, orders);
  const returnedUnits = calculateReturnedStock(productId, orders);
  return Math.max(0, receivedUnits - deliveredUnits - reservedUnits + returnedUnits);
}

export function calculateAvailableStockValue(productId, stockPurchases = [], orders = [], exchangeRate = DEFAULT_EXCHANGE_RATE) {
  const relevantPurchases = stockPurchases.filter((purchase) => getNormalizedProductId(purchase?.product_id ?? purchase?.id) === getNormalizedProductId(productId));
  const fallbackPurchase = relevantPurchases[0] || {};
  const costPerUnitUsd = calculateLandedCostPerUnit(fallbackPurchase, exchangeRate);
  const availableUnits = calculateAvailableStock(productId, stockPurchases, orders);
  const valueUsd = availableUnits * costPerUnitUsd;
  return {
    availableUnits,
    valueUsd,
    valueTsh: usdToTsh(valueUsd, exchangeRate),
  };
}

export function normalizeRegion(region) {
  return String(region || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

export function isDarEsSalaam(region) {
  const normalized = normalizeRegion(region);
  return ["dar es salaam", "dar", "daresalaam", "dsm"].includes(normalized);
}

export function getDeliveryFeeByRegion(order = {}, settings = {}) {
  if (isDarEsSalaam(order?.city || order?.region)) return Math.max(0, toNumber(settings.darFeeUsd ?? 8));
  return Math.max(0, toNumber(settings.otherFeeUsd ?? 9));
}

export function calculateServiceFeeForOrder(order = {}, settings = {}) {
  if (!isDeliveredStatus(order?.shippingStatusRaw ?? order?.shippingStatus)) {
    return { usd: 0, tsh: 0, auditFlags: [] };
  }
  const usd = getDeliveryFeeByRegion(order, settings);
  const exchangeRate = settings.exchangeRate ?? DEFAULT_EXCHANGE_RATE;
  return {
    usd,
    tsh: usdToTsh(usd, exchangeRate),
    auditFlags: order?.city ? [] : ["missing_region_fee_source"],
  };
}

export function calculateProductServiceFee(productId, orders = [], settings = {}) {
  const normalizedProductId = getNormalizedProductId(productId);
  return orders.reduce(
    (acc, order) => {
      if (getNormalizedProductId(order?.productId) !== normalizedProductId) return acc;
      const fee = calculateServiceFeeForOrder(order, settings);
      acc.usd += fee.usd;
      acc.tsh += fee.tsh;
      acc.auditFlags.push(...fee.auditFlags);
      return acc;
    },
    { usd: 0, tsh: 0, auditFlags: [] }
  );
}

export function calculateTotalServiceFee(orders = [], settings = {}) {
  return orders.reduce(
    (acc, order) => {
      const fee = calculateServiceFeeForOrder(order, settings);
      acc.usd += fee.usd;
      acc.tsh += fee.tsh;
      acc.auditFlags.push(...fee.auditFlags);
      return acc;
    },
    { usd: 0, tsh: 0, auditFlags: [] }
  );
}
