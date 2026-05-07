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
  const parts = campaign.split("|").map((p) => p.trim()).filter(Boolean);
  const productCodes = products.map((p) => generateProductMappingCode(p)).filter(Boolean);

  // When no products to match against, return raw segment[1] code
  if (!productCodes.length) {
    return parts.length >= 2 ? normalizeCode(parts[1]) : "";
  }

  // Pass 1: exact match of any pipe-segment against a product code
  for (const part of parts) {
    const code = normalizeCode(part);
    if (!code) continue;
    if (productCodes.includes(code)) return code;
  }

  // Pass 2: any pipe-segment contains a product code (e.g. "DSP4V2" contains "DSP4")
  for (const part of parts) {
    const code = normalizeCode(part);
    if (!code) continue;
    const found = productCodes.find((pc) => code.includes(pc));
    if (found) return found;
  }

  // Pass 3: full normalized campaign name contains a product code
  const normalizedCampaign = normalizeCode(campaign);
  return productCodes.find((pc) => normalizedCampaign.includes(pc)) || "";
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
