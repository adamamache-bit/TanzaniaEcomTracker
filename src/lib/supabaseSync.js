import { supabase, supabaseEnabled, supabaseWorkspaceId } from "./supabaseClient";

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
    extra: {},
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
