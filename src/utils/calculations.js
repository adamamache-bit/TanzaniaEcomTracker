import { DEFAULT_EXCHANGE_RATE, tshToUsd, usdToTsh } from "./currency";
import { calculateProductServiceFee, calculateTotalServiceFee, calculateLandedCostPerUnit } from "./stockCalculations";
import { isDeliveredStatus, isConfirmedStatus, isReturnedStatus, isPendingShippingStatus } from "../config/statusMapping";
import { deduplicateCampaigns } from "./adsMapping";

function toNumber(value) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

export function calculateRate(numerator, denominator) {
  return denominator > 0 ? (numerator / denominator) * 100 : 0;
}

export function getOrderRevenueAmounts(order = {}, product = {}, exchangeRate = DEFAULT_EXCHANGE_RATE) {
  const amountTsh = Math.max(0, toNumber(order.amount_tsh ?? order.orderTotalTzs ?? order.amountTsh));
  if (amountTsh > 0) {
    return {
      revenueTsh: amountTsh,
      revenueUsd: tshToUsd(amountTsh, exchangeRate),
      estimatedRevenueUsed: false,
      auditFlags: [],
    };
  }

  const quantity = Math.max(1, toNumber(order.quantity || 1));
  const sellingPriceUsd = Math.max(0, toNumber(product.sellingPriceUsd));
  const sellingPriceTsh = Math.max(0, toNumber(product.sellingPrice ?? usdToTsh(sellingPriceUsd, exchangeRate)));
  const revenueTsh = sellingPriceTsh * quantity;
  return {
    revenueTsh,
    revenueUsd: tshToUsd(revenueTsh, exchangeRate),
    estimatedRevenueUsed: revenueTsh > 0,
    auditFlags: revenueTsh > 0 ? ["estimated_revenue_used"] : ["missing_amount_delivered_order"],
  };
}

export function calculateProductFinancials(product = {}, orders = [], campaigns = [], settings = {}) {
  const exchangeRate = settings.exchangeRate ?? DEFAULT_EXCHANGE_RATE;
  const deliveredOrders = orders.filter((order) => isDeliveredStatus(order.shippingStatusRaw ?? order.shippingStatus));
  const deliveredUnits = deliveredOrders.reduce((sum, order) => sum + Math.max(1, toNumber(order.quantity || 1)), 0);
  const landedCostPerUnitUsd = calculateLandedCostPerUnit(
    {
      quantity_ordered: product.totalQty,
      buy_price_per_unit_usd: product.purchaseUnitPrice,
      shippingTotal: product.shippingTotal,
      otherCharges: product.otherCharges,
      sourcing_cost_usd: product.sourcingCostUsd,
    },
    exchangeRate
  );
  const revenueTsh = deliveredOrders.reduce((sum, order) => sum + getOrderRevenueAmounts(order, product, exchangeRate).revenueTsh, 0);
  const adsSpendTsh = deduplicateCampaigns(campaigns).reduce((sum, campaign) => sum + Math.max(0, toNumber(campaign.spendTsh ?? campaign.adSpend ?? 0)), 0);
  const productCostUsd = deliveredUnits * landedCostPerUnitUsd;
  const productCostTsh = usdToTsh(productCostUsd, exchangeRate);
  const serviceFee = calculateProductServiceFee(product.id, deliveredOrders, { exchangeRate, ...settings });
  const profitTsh = revenueTsh - productCostTsh - serviceFee.tsh - adsSpendTsh;
  const revenueUsd = tshToUsd(revenueTsh, exchangeRate);
  const profitUsd = tshToUsd(profitTsh, exchangeRate);
  return {
    deliveredUnits,
    revenueTsh,
    revenueUsd,
    adsSpendTsh,
    adsSpendUsd: tshToUsd(adsSpendTsh, exchangeRate),
    landedCostPerUnitUsd,
    productCostUsd,
    productCostTsh,
    serviceFeeUsd: serviceFee.usd,
    serviceFeeTsh: serviceFee.tsh,
    profitUsd,
    profitTsh,
    marginPct: revenueUsd > 0 ? (profitUsd / revenueUsd) * 100 : null,
  };
}

// Section 11 — WINNER / TESTING / LOSS / RESTOCK / WATCH
export function getProductDecisionStatus(metrics = {}, thresholds = {}) {
  const minOrders = Math.max(1, toNumber(thresholds.minDeliveredOrders ?? 5));
  const minAdSpendUsd = Math.max(0, toNumber(thresholds.minAdSpendForDecision ?? 0));
  const minMarginPct = toNumber(thresholds.winnerMarginPct ?? 10);
  const maxCpaUsd = toNumber(thresholds.breakEvenCpaUsd ?? 0);
  const lowStockQty = Math.max(0, toNumber(thresholds.minStockQuantity ?? 3));

  const deliveredUnits = toNumber(metrics.deliveredUnits ?? metrics.deliveredOrders ?? 0);
  const profitUsd = toNumber(metrics.profitUsd ?? 0);
  const marginPct = metrics.marginPct ?? null;
  const cpaUsd = toNumber(metrics.cpaUsd ?? 0);
  const stockQty = toNumber(metrics.availableStock ?? metrics.stockQty ?? 0);
  const adsSpendUsd = toNumber(metrics.adsSpendUsd ?? 0);

  const hasEnoughData = deliveredUnits >= minOrders || (minAdSpendUsd > 0 && adsSpendUsd >= minAdSpendUsd);
  if (!hasEnoughData) return "testing";

  if (profitUsd < 0) return "loss";

  if (profitUsd > 0 && stockQty <= lowStockQty) return "restock";

  const marginOk = marginPct === null || marginPct >= minMarginPct;
  const cpaOk = maxCpaUsd <= 0 || cpaUsd <= 0 || cpaUsd <= maxCpaUsd;
  if (profitUsd > 0 && marginOk && cpaOk) return "winner";

  return "watch";
}

export function calculateGlobalBusinessMetrics({ products = [], orders = [], productCampaigns = {}, unmappedAdsTsh = 0, salariesTsh = 0, toolsTsh = 0, extraGlobalChargesTsh = 0, otherGlobalExpensesTsh = 0, ownerInjectionTsh = 0, exchangeRate = DEFAULT_EXCHANGE_RATE } = {}) {
  const productMetrics = products.map((product) =>
    calculateProductFinancials(product, orders.filter((order) => order.productId === product.id), productCampaigns[product.id] || [], { exchangeRate })
  );
  const totalRevenueTsh = productMetrics.reduce((sum, metric) => sum + metric.revenueTsh, 0);
  const totalProductProfitTsh = productMetrics.reduce((sum, metric) => sum + metric.profitTsh, 0);
  const totalMappedAdsTsh = productMetrics.reduce((sum, metric) => sum + metric.adsSpendTsh, 0);
  const resolvedUnmappedAdsTsh = Math.max(0, toNumber(unmappedAdsTsh));
  const globalExpensesTsh = Math.max(0, toNumber(salariesTsh) + toNumber(toolsTsh) + toNumber(extraGlobalChargesTsh) + toNumber(otherGlobalExpensesTsh));
  // unmapped ads reduce global profit but are not part of any product's P&L
  const businessProfitWithoutInjectionTsh = totalProductProfitTsh - resolvedUnmappedAdsTsh - globalExpensesTsh;
  const cashBalanceWithInjectionTsh = totalRevenueTsh + Math.max(0, toNumber(ownerInjectionTsh)) - (totalMappedAdsTsh + resolvedUnmappedAdsTsh + globalExpensesTsh + productMetrics.reduce((sum, metric) => sum + metric.productCostTsh + metric.serviceFeeTsh, 0));
  return {
    totalRevenueTsh,
    totalRevenueUsd: tshToUsd(totalRevenueTsh, exchangeRate),
    businessProfitWithoutInjectionTsh,
    businessProfitWithoutInjectionUsd: tshToUsd(businessProfitWithoutInjectionTsh, exchangeRate),
    cashBalanceWithInjectionTsh,
    cashBalanceWithInjectionUsd: tshToUsd(cashBalanceWithInjectionTsh, exchangeRate),
    globalMargin: totalRevenueTsh > 0 ? (businessProfitWithoutInjectionTsh / totalRevenueTsh) * 100 : null,
    totalMappedAdsTsh,
    totalMappedAdsUsd: tshToUsd(totalMappedAdsTsh, exchangeRate),
    unmappedAdsTsh: resolvedUnmappedAdsTsh,
    unmappedAdsUsd: tshToUsd(resolvedUnmappedAdsTsh, exchangeRate),
  };
}

export function buildCalculationAudit({ orders = [], products = [], productCampaigns = {}, unmappedAdsTsh = 0, salariesTsh = 0, toolsTsh = 0, extraGlobalChargesTsh = 0, otherGlobalExpensesTsh = 0, ownerInjectionTsh = 0, exchangeRate = DEFAULT_EXCHANGE_RATE } = {}) {
  const totalLeads = orders.length;
  const confirmedOrders = orders.filter((order) => isConfirmedStatus(order.confirmationStatusRaw ?? order.confirmationStatus)).length;
  const deliveredOrders = orders.filter((order) => isDeliveredStatus(order.shippingStatusRaw ?? order.shippingStatus)).length;
  const returnedOrders = orders.filter((order) => isReturnedStatus(order.shippingStatusRaw ?? order.shippingStatus)).length;
  const pendingShipping = orders.filter((order) => isPendingShippingStatus(order.shippingStatusRaw ?? order.shippingStatus)).length;
  const deliveredRevenueTsh = orders.reduce((sum, order) => {
    if (!isDeliveredStatus(order.shippingStatusRaw ?? order.shippingStatus)) return sum;
    const product = products.find((entry) => entry.id === order.productId) || {};
    return sum + getOrderRevenueAmounts(order, product, exchangeRate).revenueTsh;
  }, 0);
  const serviceFee = calculateTotalServiceFee(orders, { exchangeRate });
  const globalMetrics = calculateGlobalBusinessMetrics({
    products,
    orders,
    productCampaigns,
    unmappedAdsTsh,
    salariesTsh,
    toolsTsh,
    extraGlobalChargesTsh,
    otherGlobalExpensesTsh,
    ownerInjectionTsh,
    exchangeRate,
  });
  const missingAmountOrders = orders.filter(
    (order) =>
      isDeliveredStatus(order.shippingStatusRaw ?? order.shippingStatus) &&
      Math.max(0, toNumber(order.amount_tsh ?? order.orderTotalTzs)) <= 0
  ).length;
  const zeroAmountOrders = missingAmountOrders;
  const estimatedRevenueOrders = missingAmountOrders;
  const missingRegionOrders = orders.filter(
    (order) => isDeliveredStatus(order.shippingStatusRaw ?? order.shippingStatus) && !String(order.city || "").trim()
  ).length;

  return {
    totalLeads,
    confirmedOrders,
    deliveredOrders,
    returnedOrders,
    pendingShipping,
    confirmationRate: calculateRate(confirmedOrders, totalLeads),
    deliveryRate: calculateRate(deliveredOrders, confirmedOrders),
    deliveredRevenueTsh,
    deliveredRevenueUsd: tshToUsd(deliveredRevenueTsh, exchangeRate),
    totalServiceFeeTsh: serviceFee.tsh,
    totalServiceFeeUsd: serviceFee.usd,
    salariesTsh,
    toolsTsh,
    extraGlobalChargesTsh,
    otherGlobalExpensesTsh,
    ownerInjectionTsh,
    totalMappedAdsTsh: globalMetrics.totalMappedAdsTsh,
    totalMappedAdsUsd: globalMetrics.totalMappedAdsUsd,
    unmappedAdsTsh: globalMetrics.unmappedAdsTsh,
    unmappedAdsUsd: globalMetrics.unmappedAdsUsd,
    businessProfitWithoutInjectionTsh: globalMetrics.businessProfitWithoutInjectionTsh,
    businessProfitWithoutInjectionUsd: globalMetrics.businessProfitWithoutInjectionUsd,
    cashBalanceWithInjectionTsh: globalMetrics.cashBalanceWithInjectionTsh,
    cashBalanceWithInjectionUsd: globalMetrics.cashBalanceWithInjectionUsd,
    globalMargin: globalMetrics.globalMargin,
    missingAmountOrders,
    zeroAmountOrders,
    estimatedRevenueOrders,
    missingRegionOrders,
  };
}
