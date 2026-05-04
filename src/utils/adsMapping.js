import { DEFAULT_EXCHANGE_RATE, usdToTsh } from "./currency";

function normalizeCode(value) {
  return String(value || "")
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, "");
}

function toNumber(value) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

// Deduplicate campaigns by campaignId + date to prevent double-counting
// when the same campaign appears in overlapping Meta sync windows.
export function deduplicateCampaigns(campaigns) {
  const seen = new Set();
  return campaigns.filter((campaign) => {
    const id = String(campaign?.id ?? campaign?.campaignId ?? campaign?.campaign_id ?? "");
    const date = String(campaign?.date ?? campaign?.dateStart ?? campaign?.bucket ?? "");
    if (!id) return true; // no ID means we can't deduplicate, keep it
    const key = `${id}:${date}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

export function generateProductMappingCode(product = {}) {
  const explicit = normalizeCode(product.code || product.mappingCode);
  if (explicit) return explicit;
  const fromRef = normalizeCode(product.product_ref || product.id);
  if (fromRef) return fromRef;
  const words = String(product.name || "")
    .split(/[^A-Za-z0-9]+/)
    .filter(Boolean);
  return normalizeCode(words.map((word) => word.slice(0, 2)).join("").slice(0, 8) || "PRD");
}

export function extractMappingCodeFromCampaignName(campaignName, products = []) {
  const campaign = String(campaignName || "");
  const structuredParts = campaign.split("|").map((part) => part.trim());
  if (structuredParts.length >= 2) {
    const directCode = normalizeCode(structuredParts[1]);
    if (directCode) return directCode;
  }

  const productCodes = products.map((product) => generateProductMappingCode(product)).filter(Boolean);
  const normalizedCampaign = normalizeCode(campaign);
  return productCodes.find((code) => normalizedCampaign.includes(code)) || "";
}

export function mapCampaignSpendToProduct(campaign = {}, products = [], exchangeRate = DEFAULT_EXCHANGE_RATE) {
  const mappingCode = extractMappingCodeFromCampaignName(campaign.campaignName || campaign.campaign_name, products);
  const matchedProduct = products.find((product) => generateProductMappingCode(product) === mappingCode) || null;
  const spendUsd = Math.max(0, toNumber(campaign.spendUsd ?? campaign.spend));
  const spendTsh = Math.max(0, toNumber(campaign.spendTzs ?? usdToTsh(spendUsd, exchangeRate)));

  return {
    ...campaign,
    mappingCode,
    productId: matchedProduct?.id || "",
    mapped: Boolean(matchedProduct),
    spendUsd,
    spendTsh,
  };
}

export function calculateProductAdsSpend(productId, campaigns = [], products = [], exchangeRate = DEFAULT_EXCHANGE_RATE) {
  return deduplicateCampaigns(campaigns).reduce((sum, campaign) => {
    const mapped = mapCampaignSpendToProduct(campaign, products, exchangeRate);
    return mapped.productId === productId ? sum + mapped.spendTsh : sum;
  }, 0);
}

export function calculateUnmappedAdsSpend(campaigns = [], products = [], exchangeRate = DEFAULT_EXCHANGE_RATE) {
  return deduplicateCampaigns(campaigns).reduce((sum, campaign) => {
    const mapped = mapCampaignSpendToProduct(campaign, products, exchangeRate);
    return !mapped.productId ? sum + mapped.spendTsh : sum;
  }, 0);
}

export function calculateTotalAdsSpend(campaigns = [], products = [], exchangeRate = DEFAULT_EXCHANGE_RATE) {
  return deduplicateCampaigns(campaigns).reduce((sum, campaign) => {
    const mapped = mapCampaignSpendToProduct(campaign, products, exchangeRate);
    return sum + mapped.spendTsh;
  }, 0);
}
