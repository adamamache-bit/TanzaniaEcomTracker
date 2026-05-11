import { supabase, supabaseEnabled, supabaseWorkspaceId } from "./supabaseClient";
import { isDeliveredStatus } from "../config/statusMapping";

export async function checkNormalizedTablesEmpty(workspaceId = supabaseWorkspaceId) {
  if (!supabaseEnabled || !supabase) return true;
  try {
    const { count } = await supabase
      .from("products")
      .select("id", { count: "exact", head: true })
      .eq("workspace_id", workspaceId);
    return (count || 0) === 0;
  } catch {
    return true;
  }
}

async function syncProducts(products, workspaceId) {
  if (!Array.isArray(products) || !products.length) return;
  const rows = products.map((p) => ({
    id: p.id,
    workspace_id: workspaceId,
    name: p.name || "",
    source: p.source || "",
    selling_price: Number(p.sellingPrice || 0),
    purchase_unit_price: Number(p.purchaseUnitPrice || 0),
    total_qty: Number(p.totalQty || 0),
    shipping_total: Number(p.shippingTotal || 0),
    other_charges: Number(p.otherCharges || 0),
    delivery: Number(p.delivery || 0),
    estimated_arrival_days: Number(p.estimatedArrivalDays || 0),
    stock_arrival_status: p.stockArrivalStatus || "",
    stock_ordered_at: p.stockOrderedAt || null,
    next_arrival_check_date: p.nextArrivalCheckDate || null,
    stock_arrived_at: p.stockArrivedAt || null,
    offers: Array.isArray(p.offers) ? p.offers : [],
    extra: { mappingCode: p.mappingCode || "" },
    updated_at: new Date().toISOString(),
  }));
  const { error } = await supabase
    .from("products")
    .upsert(rows, { onConflict: "id,workspace_id" });
  if (error) throw error;
}

async function syncOrders(customers, workspaceId) {
  if (!Array.isArray(customers) || !customers.length) return;
  const rows = customers.map((c) => ({
    id: c.id,
    workspace_id: workspaceId,
    customer_name: c.customerName || "",
    phone: c.phone || "",
    city: c.city || "",
    address: c.address || "",
    product_id: c.productId || "",
    quantity: Number(c.quantity || 1),
    order_date: c.orderDate || null,
    payment_method: c.paymentMethod || "COD",
    notes: c.notes || "",
    lead_source: c.leadSource || "facebook",
    campaign_name: c.campaignName || "",
    adset_name: c.adsetName || "",
    creative_name: c.creativeName || "",
    priority: c.priority || "normal",
    customer_type: c.customerType || "new",
    call_attempts: Number(c.callAttempts || 0),
    cancel_reason: c.cancelReason || "",
    unreached_reason: c.unreachedReason || "",
    carrier_name: c.carrierName || "",
    tracking_number: c.trackingNumber || "",
    expected_delivery_date: c.expectedDeliveryDate || null,
    actual_delivery_date: c.actualDeliveryDate || null,
    return_reason: c.returnReason || "",
    order_total_tzs: Number(c.orderTotalTzs || 0),
    amount_tsh: Number(c.amount_tsh || 0),
    amount_usd: Number(c.amount_usd || 0),
    source_order_id: c.sourceOrderId || null,
    import_source: c.importSource || null,
    last_imported_at: c.lastImportedAt || null,
    last_shipping_imported_at: c.lastShippingImportedAt || null,
    updated_at: c.updatedAt || null,
    confirmation_updated_at: c.confirmation_updated_at || null,
    assigned_to: c.assignedTo || "",
    confirmation_status: c.confirmationStatus || "",
    shipping_status: c.shippingStatus || "",
    status: c.status || "",
    history: Array.isArray(c.history) ? c.history : [],
    extra: { raw_row_data: c.raw_row_data || null, extra_fields: c.extra_fields || {} },
  }));
  for (let i = 0; i < rows.length; i += 500) {
    const { error } = await supabase
      .from("orders")
      .upsert(rows.slice(i, i + 500), { onConflict: "id,workspace_id" });
    if (error) throw error;
  }
}

async function syncTracking(tracking, workspaceId) {
  if (!Array.isArray(tracking) || !tracking.length) return;
  const rows = tracking.map((t) => ({
    id: t.id,
    workspace_id: workspaceId,
    product_id: t.productId || "",
    ad_spend: Number(t.adSpend || 0),
    orders: Number(t.orders || 0),
    confirmed: Number(t.confirmed || 0),
    delivered: Number(t.delivered || 0),
    name: t.name || "",
    date_start: t.dateStart || t.metaSince || null,
    date_end: t.dateEnd || t.metaUntil || null,
    meta_managed: Boolean(t.metaManaged),
    meta_imported_at: t.metaImportedAt || null,
    meta_currency: t.metaCurrency || "USD",
    extra: {},
    updated_at: new Date().toISOString(),
  }));
  const { error } = await supabase
    .from("tracking")
    .upsert(rows, { onConflict: "id,workspace_id" });
  if (error) throw error;
}

async function syncSettings(snapshot, workspaceId) {
  const entries = [
    { key: "serviceForm", value: snapshot.serviceForm },
    { key: "situationData", value: snapshot.situationData },
    {
      key: "metaAdsState",
      value: snapshot.metaAdsState ? { ...snapshot.metaAdsState, accessToken: "" } : null,
    },
    { key: "importMeta", value: snapshot.importMeta },
  ].filter((e) => e.value != null && typeof e.value === "object");
  if (!entries.length) return;
  const rows = entries.map((e) => ({
    key: e.key,
    workspace_id: workspaceId,
    value: e.value,
    updated_at: new Date().toISOString(),
  }));
  const { error } = await supabase
    .from("settings")
    .upsert(rows, { onConflict: "key,workspace_id" });
  if (error) throw error;
}

async function syncAuditTrail(customers, workspaceId) {
  if (!Array.isArray(customers)) return;
  const entries = [];
  for (const c of customers) {
    if (!Array.isArray(c.history)) continue;
    for (const h of c.history) {
      if (!h?.id) continue;
      entries.push({
        id: h.id,
        workspace_id: workspaceId,
        customer_id: c.id,
        customer_name: c.customerName || "",
        product_id: c.productId || "",
        action: h.action || "",
        source: h.source || "system",
        details: h.details || "",
        occurred_at: h.at || null,
      });
    }
  }
  if (!entries.length) return;
  for (let i = 0; i < entries.length; i += 200) {
    await supabase
      .from("audit_trail")
      .upsert(entries.slice(i, i + 200), { onConflict: "id" });
  }
}

export async function clearNormalizedProducts(workspaceId = supabaseWorkspaceId) {
  if (!supabaseEnabled || !supabase) return;
  const { error } = await supabase.from("products").delete().eq("workspace_id", workspaceId);
  if (error) throw error;
}

// Fire-and-forget: sync all normalized tables in parallel. Errors are silently swallowed.
export async function syncNormalizedTables(snapshot = {}, workspaceId = supabaseWorkspaceId) {
  if (!supabaseEnabled || !supabase) return;
  await Promise.allSettled([
    syncProducts(snapshot.products, workspaceId),
    syncOrders(snapshot.customers, workspaceId),
    syncTracking(snapshot.tracking, workspaceId),
    syncSettings(snapshot, workspaceId),
    syncAuditTrail(snapshot.customers, workspaceId),
  ]);
}

// ── ADS CAMPAIGNS: optional persistent spend table ───────────────────────────

function isAdsCampaignsTableMissing(error) {
  const message = String(error?.message || error?.details || error?.hint || "").toLowerCase();
  return error?.code === "42P01" || message.includes("ads_campaigns");
}

export async function saveAdsCampaignsToSupabase(campaigns = []) {
  if (!supabaseEnabled || !supabase || !campaigns.length) return { saved: false, notice: "" };
  const rows = campaigns
    .map((c) => ({
      campaign_id: String(c.campaignId || ""),
      campaign_name: String(c.campaignName || ""),
      date_start: String(c.dateStart || ""),
      date_end: String(c.dateEnd || ""),
      spend_usd: Number(c.spendUsd || 0),
      spend_tsh: Number(c.spendTsh || 0),
      product_id: String(c.productId || ""),
      product_name: String(c.productName || ""),
      is_mapped: Boolean(c.isMapped),
      impressions: Number(c.impressions || 0),
      clicks: Number(c.clicks || 0),
      leads: Number(c.leads || 0),
      updated_at: new Date().toISOString(),
    }))
    .filter((r) => r.campaign_id && r.date_start && r.date_end);
  if (!rows.length) return { saved: false, notice: "" };
  console.log("[AdsSync] Saving", rows.length, "campaign(s) to ads_campaigns:", rows.map((r) => ({ campaign_id: r.campaign_id, product_id: r.product_id, is_mapped: r.is_mapped, spend_usd: r.spend_usd, spend_tsh: r.spend_tsh })));
  try {
    const { error } = await supabase
      .from("ads_campaigns")
      .upsert(rows, { onConflict: "campaign_id,date_start", ignoreDuplicates: false });
    if (error) {
      console.error("[AdsSync] Save error:", error);
      if (isAdsCampaignsTableMissing(error)) return { saved: false, notice: "Run migration 003_ads_campaigns.sql in your Supabase SQL editor to enable campaign tracking." };
      throw error;
    }
    console.log("[AdsSync] Save OK —", rows.length, "row(s) upserted");
    return { saved: true, notice: "" };
  } catch (err) {
    console.error("[AdsSync] Save exception:", err);
    if (isAdsCampaignsTableMissing(err)) return { saved: false, notice: "Run migration 003_ads_campaigns.sql in your Supabase SQL editor to enable campaign tracking." };
    return { saved: false, notice: String(err?.message || "Failed to save ads campaigns.") };
  }
}

export async function loadAdsSpendByProductFromSupabase() {
  if (!supabaseEnabled || !supabase) return {};
  try {
    const { data, error } = await supabase
      .from("ads_campaigns")
      .select("product_id, spend_usd, spend_tsh, leads")
      .eq("is_mapped", true);
    if (error) {
      if (isAdsCampaignsTableMissing(error)) return {};
      throw error;
    }
    const map = {};
    for (const row of (data || [])) {
      const pid = String(row.product_id || "");
      if (!pid) continue;
      if (!map[pid]) map[pid] = { spendUsd: 0, spendTsh: 0, leads: 0, campaignCount: 0 };
      map[pid].spendUsd += Number(row.spend_usd || 0);
      map[pid].spendTsh += Number(row.spend_tsh || 0);
      map[pid].leads += Number(row.leads || 0);
      map[pid].campaignCount += 1;
    }
    console.log("[AdsSync] Spend by product from Supabase:", map);
    return map;
  } catch (err) {
    console.error("[AdsSync] loadAdsSpendByProductFromSupabase failed:", err);
    return {};
  }
}

export async function loadAdsCampaignsFromSupabase() {
  if (!supabaseEnabled || !supabase) return { available: false, campaigns: [] };
  try {
    const { data, error } = await supabase
      .from("ads_campaigns")
      .select("*")
      .order("created_at", { ascending: false });
    if (error) {
      if (isAdsCampaignsTableMissing(error)) return { available: false, campaigns: [] };
      throw error;
    }
    const campaigns = (data || []).map((row) => ({
      campaignId: row.campaign_id || "",
      campaignName: row.campaign_name || "",
      dateStart: String(row.date_start || ""),
      dateEnd: String(row.date_end || ""),
      spendUsd: Number(row.spend_usd || 0),
      spendTsh: Number(row.spend_tsh || 0),
      productId: row.product_id || "",
      productName: row.product_name || "",
      isMapped: Boolean(row.is_mapped),
      isUnmapped: !row.is_mapped,
      impressions: Number(row.impressions || 0),
      clicks: Number(row.clicks || 0),
      leads: Number(row.leads || 0),
      savedAt: row.created_at || null,
    }));
    return { available: true, campaigns };
  } catch {
    return { available: false, campaigns: [] };
  }
}

// ── READ: reconstruct workspace state from normalized tables ─────────────────

function rowToProduct(row) {
  return {
    id: row.id,
    name: row.name || "",
    source: row.source || "china",
    sellingPrice: Number(row.selling_price || 0),
    purchaseUnitPrice: Number(row.purchase_unit_price || 0),
    totalQty: Number(row.total_qty || 0),
    shippingTotal: Number(row.shipping_total || 0),
    otherCharges: Number(row.other_charges || 0),
    delivery: Number(row.delivery || 0),
    estimatedArrivalDays: Number(row.estimated_arrival_days || 0),
    stockArrivalStatus: row.stock_arrival_status || "arrived",
    stockOrderedAt: row.stock_ordered_at || null,
    nextArrivalCheckDate: row.next_arrival_check_date || null,
    stockArrivedAt: row.stock_arrived_at || null,
    offers: Array.isArray(row.offers) ? row.offers : [],
    mappingCode: row.extra?.mappingCode || "",
  };
}

function rowToCustomer(row) {
  return {
    id: row.id,
    customerName: row.customer_name || "",
    phone: row.phone || "",
    city: row.city || "",
    address: row.address || "",
    productId: row.product_id || "",
    quantity: Number(row.quantity || 1),
    orderDate: row.order_date || null,
    paymentMethod: row.payment_method || "COD",
    notes: row.notes || "",
    leadSource: row.lead_source || "facebook",
    campaignName: row.campaign_name || "",
    adsetName: row.adset_name || "",
    creativeName: row.creative_name || "",
    priority: row.priority || "normal",
    customerType: row.customer_type || "new",
    callAttempts: Number(row.call_attempts || 0),
    cancelReason: row.cancel_reason || "",
    unreachedReason: row.unreached_reason || "",
    carrierName: row.carrier_name || "",
    trackingNumber: row.tracking_number || "",
    expectedDeliveryDate: row.expected_delivery_date || null,
    actualDeliveryDate: row.actual_delivery_date || null,
    returnReason: row.return_reason || "",
    orderTotalTzs: Number(row.order_total_tzs || 0),
    amount_tsh: Number(row.amount_tsh || 0),
    amount_usd: Number(row.amount_usd || 0),
    sourceOrderId: row.source_order_id || null,
    importSource: row.import_source || null,
    lastImportedAt: row.last_imported_at || null,
    lastShippingImportedAt: row.last_shipping_imported_at || null,
    updatedAt: row.updated_at || null,
    confirmation_updated_at: row.confirmation_updated_at || null,
    assignedTo: row.assigned_to || "",
    confirmationStatus: row.confirmation_status || "",
    shippingStatus: row.shipping_status || "",
    status: row.status || "",
    history: Array.isArray(row.history) ? row.history : [],
    raw_row_data: row.extra?.raw_row_data || null,
    extra_fields: row.extra?.extra_fields || {},
  };
}

function rowToTracking(row) {
  return {
    id: row.id,
    productId: row.product_id || "",
    adSpend: Number(row.ad_spend || 0),
    orders: Number(row.orders || 0),
    confirmed: Number(row.confirmed || 0),
    delivered: Number(row.delivered || 0),
    name: row.name || "",
    dateStart: row.date_start || null,
    dateEnd: row.date_end || null,
    metaManaged: Boolean(row.meta_managed),
    metaImportedAt: row.meta_imported_at || null,
    metaCurrency: row.meta_currency || "USD",
  };
}

export async function loadWorkspaceFromNormalizedTables(workspaceId = supabaseWorkspaceId) {
  if (!supabaseEnabled || !supabase) return null;

  const [productsRes, ordersRes, trackingRes, settingsRes] = await Promise.all([
    supabase.from("products").select("*").eq("workspace_id", workspaceId),
    supabase.from("orders").select("*").eq("workspace_id", workspaceId),
    supabase.from("tracking").select("*").eq("workspace_id", workspaceId),
    supabase.from("settings").select("*").eq("workspace_id", workspaceId),
  ]);

  const products = (productsRes.data || []).map(rowToProduct);
  const customers = (ordersRes.data || []).map(rowToCustomer);
  const tracking = (trackingRes.data || []).map(rowToTracking);

  const settingsMap = {};
  for (const row of settingsRes.data || []) {
    settingsMap[row.key] = row.value;
  }

  return {
    products,
    customers,
    tracking,
    serviceForm: settingsMap.serviceForm || null,
    situationData: settingsMap.situationData || null,
    metaAdsState: settingsMap.metaAdsState ? { ...settingsMap.metaAdsState, accessToken: "" } : null,
    importMeta: settingsMap.importMeta || null,
  };
}

// Full migration with result report (used for auto-migration and manual button).
export async function migrateWorkspaceToNormalizedTables(
  snapshot = {},
  workspaceId = supabaseWorkspaceId
) {
  if (!supabaseEnabled || !supabase) return { success: false, counts: {}, errors: ["Supabase not configured"] };
  const results = await Promise.allSettled([
    syncProducts(snapshot.products, workspaceId),
    syncOrders(snapshot.customers, workspaceId),
    syncTracking(snapshot.tracking, workspaceId),
    syncSettings(snapshot, workspaceId),
    syncAuditTrail(snapshot.customers, workspaceId),
  ]);
  const errors = results
    .filter((r) => r.status === "rejected")
    .map((r) => r.reason?.message || "unknown error");
  return {
    success: errors.length === 0,
    counts: {
      products: snapshot.products?.length || 0,
      orders: snapshot.customers?.length || 0,
      tracking: snapshot.tracking?.length || 0,
    },
    errors,
  };
}

// ── PROFIT OVERVIEW: direct Supabase aggregate for the Overview tab ───────────

export async function loadProfitOverviewFromSupabase() {
  if (!supabaseEnabled || !supabase) return null;
  try {
    const [ordersResult, adsResult] = await Promise.all([
      supabase.from("orders").select("quantity, amount_tsh, order_total_tzs, city, shipping_status"),
      supabase.from("ads_campaigns").select("spend_usd"),
    ]);

    let revenueTsh = 0;
    let deliveredUnits = 0;
    let serviceChargesUsd = 0;

    for (const row of (ordersResult.data || [])) {
      if (!isDeliveredStatus(row.shipping_status)) continue;
      const qty = Math.max(1, Number(row.quantity || 1));
      revenueTsh += Number(row.amount_tsh || row.order_total_tzs || 0);
      deliveredUnits += qty;
      const city = String(row.city || "").toLowerCase();
      const isDar = city.includes("dar");
      serviceChargesUsd += qty * (isDar ? 8 : 9);
    }

    const adsSpendUsd = (adsResult.data || []).reduce((sum, row) => sum + Number(row.spend_usd || 0), 0);

    console.log("[ProfitSync] Overview from Supabase:", { revenueTsh, deliveredUnits, serviceChargesUsd, adsSpendUsd });
    return { revenueTsh, deliveredUnits, serviceChargesUsd, adsSpendUsd };
  } catch (err) {
    console.error("[ProfitSync] loadProfitOverviewFromSupabase failed:", err);
    return null;
  }
}

// ── REVENUE IMPORT: persist Excel-imported revenue totals ─────────────────────

export async function saveRevenueImportToSupabase(data) {
  if (!supabaseEnabled || !supabase) return false;
  try {
    const { error } = await supabase
      .from("settings")
      .upsert(
        { key: "revenue_import", workspace_id: supabaseWorkspaceId, value: data, updated_at: new Date().toISOString() },
        { onConflict: "key,workspace_id" }
      );
    if (error) throw error;
    return true;
  } catch (err) {
    console.error("[RevenueSync] saveRevenueImportToSupabase failed:", err);
    return false;
  }
}

export async function loadRevenueImportFromSupabase() {
  if (!supabaseEnabled || !supabase) return null;
  try {
    const { data, error } = await supabase
      .from("settings")
      .select("value")
      .eq("key", "revenue_import")
      .eq("workspace_id", supabaseWorkspaceId)
      .maybeSingle();
    if (error || !data) return null;
    return data.value || null;
  } catch {
    return null;
  }
}

// ── MANUAL ADS SPEND ──────────────────────────────────────────────────────────

export async function loadManualAdsSpendFromSupabase() {
  if (!supabaseEnabled || !supabase) return [];
  try {
    const { data, error } = await supabase.from("manual_ads_spend").select("*").order("week_start", { ascending: false });
    if (error) { if (error.code === "42P01") return []; throw error; }
    return (data || []).map((row) => ({
      id: row.id,
      weekStart: row.week_start || "",
      weekEnd: row.week_end || "",
      productId: row.product_id || "",
      productName: row.product_name || "",
      amountUsd: Number(row.amount_usd || 0),
      amountTsh: Number(row.amount_tsh || 0),
      notes: row.notes || "",
      createdAt: row.created_at || "",
    }));
  } catch { return []; }
}

export async function saveManualAdsSpendToSupabase(entry) {
  if (!supabaseEnabled || !supabase) return null;
  const row = {
    week_start: entry.weekStart,
    week_end: entry.weekEnd,
    product_id: entry.productId || "",
    product_name: entry.productName || "",
    amount_usd: Number(entry.amountUsd || 0),
    amount_tsh: Number(entry.amountTsh || 0),
    notes: entry.notes || "",
  };
  if (entry.id) row.id = entry.id;
  const { data, error } = await supabase.from("manual_ads_spend").upsert([row], { onConflict: "id" }).select().single();
  if (error) throw error;
  return data;
}

export async function deleteManualAdsSpendFromSupabase(id) {
  if (!supabaseEnabled || !supabase) return;
  await supabase.from("manual_ads_spend").delete().eq("id", id);
}

// ── EXTRA CHARGES ─────────────────────────────────────────────────────────────

export async function loadExtraChargesFromSupabase() {
  if (!supabaseEnabled || !supabase) return [];
  try {
    const { data, error } = await supabase.from("extra_charges").select("*").order("date", { ascending: false });
    if (error) { if (error.code === "42P01") return []; throw error; }
    return (data || []).map((row) => ({
      id: row.id,
      date: row.date || "",
      category: row.category || "other",
      description: row.description || "",
      amountUsd: Number(row.amount_usd || 0),
      amountTsh: Number(row.amount_tsh || 0),
      createdAt: row.created_at || "",
    }));
  } catch { return []; }
}

export async function saveExtraChargeToSupabase(entry) {
  if (!supabaseEnabled || !supabase) return null;
  const row = {
    date: entry.date,
    category: entry.category || "other",
    description: entry.description || "",
    amount_usd: Number(entry.amountUsd || 0),
    amount_tsh: Number(entry.amountTsh || 0),
  };
  if (entry.id) row.id = entry.id;
  const { data, error } = await supabase.from("extra_charges").upsert([row], { onConflict: "id" }).select().single();
  if (error) throw error;
  return data;
}

export async function deleteExtraChargeFromSupabase(id) {
  if (!supabaseEnabled || !supabase) return;
  await supabase.from("extra_charges").delete().eq("id", id);
}

// ── OWNER INJECTIONS ──────────────────────────────────────────────────────────

export async function loadOwnerInjectionsFromSupabase() {
  if (!supabaseEnabled || !supabase) return [];
  try {
    const { data, error } = await supabase.from("owner_injections").select("*").order("date", { ascending: false });
    if (error) { if (error.code === "42P01") return []; throw error; }
    return (data || []).map((row) => ({
      id: row.id,
      date: row.date || "",
      amountUsd: Number(row.amount_usd || 0),
      amountTsh: Number(row.amount_tsh || 0),
      notes: row.notes || "",
      createdAt: row.created_at || "",
    }));
  } catch { return []; }
}

export async function saveOwnerInjectionToSupabase(entry) {
  if (!supabaseEnabled || !supabase) return null;
  const row = {
    date: entry.date,
    amount_usd: Number(entry.amountUsd || 0),
    amount_tsh: Number(entry.amountTsh || 0),
    notes: entry.notes || "",
  };
  if (entry.id) row.id = entry.id;
  const { data, error } = await supabase.from("owner_injections").upsert([row], { onConflict: "id" }).select().single();
  if (error) throw error;
  return data;
}

export async function deleteOwnerInjectionFromSupabase(id) {
  if (!supabaseEnabled || !supabase) return;
  await supabase.from("owner_injections").delete().eq("id", id);
}
