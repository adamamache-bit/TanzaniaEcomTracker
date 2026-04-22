import express from "express";
import cors from "cors";
import dotenv from "dotenv";
import { promises as fs } from "fs";
import path from "path";
import { fileURLToPath } from "url";

dotenv.config();

const app = express();
const PORT = Number(process.env.API_PORT || 4174);
const META_GRAPH_VERSION = process.env.META_GRAPH_VERSION || "v23.0";
const META_GRAPH_BASE = `https://graph.facebook.com/${META_GRAPH_VERSION}`;
const META_LEAD_ACTION_PRIORITY = [
  "onsite_conversion.lead_grouped",
  "omni_lead",
  "offsite_conversion.fb_pixel_lead",
  "lead",
  "onsite_web_lead",
  "onsite_conversion.messaging_first_reply",
];
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const DATA_DIR = process.env.DATA_DIR || path.join(__dirname, "data");
const SHARED_STATE_FILE = path.join(DATA_DIR, "shared-workspace.json");
const SHARED_STATE_BACKUP_FILE = path.join(DATA_DIR, "shared-workspace.backup.json");

app.use(cors({ origin: true }));
app.use(express.json({ limit: "1mb" }));

function toNumber(value) {
  const parsed = Number(value || 0);
  return Number.isFinite(parsed) ? parsed : 0;
}

function ensureAccountId(accountId) {
  const normalized = String(accountId || "").trim();
  if (!normalized) return "";
  return normalized.startsWith("act_") ? normalized : `act_${normalized}`;
}

function findActionValue(list = [], actionType) {
  const matched = list.find((entry) => String(entry?.action_type || "") === actionType);
  return matched ? toNumber(matched.value) : 0;
}

function pickLeadValue(actions = []) {
  for (const actionType of META_LEAD_ACTION_PRIORITY) {
    const value = findActionValue(actions, actionType);
    if (value > 0) {
      return { leads: value, leadType: actionType };
    }
  }

  const fallback = actions
    .filter((entry) => String(entry?.action_type || "").toLowerCase().includes("lead"))
    .map((entry) => ({ actionType: String(entry?.action_type || ""), value: toNumber(entry?.value) }))
    .sort((a, b) => b.value - a.value)[0];

  return fallback ? { leads: fallback.value, leadType: fallback.actionType } : { leads: 0, leadType: "" };
}

function pickActionMetric(actions = [], priority = []) {
  for (const actionType of priority) {
    const value = findActionValue(actions, actionType);
    if (value > 0) {
      return { value, actionType };
    }
  }

  return { value: 0, actionType: "" };
}

function pickTrackedLeadMetric({ actualLeads = 0, landingPageViews = 0, inlineLinkClicks = 0 }) {
  if (actualLeads > 0) {
    return { trackedLeads: actualLeads, trackedLeadType: "lead_action" };
  }

  if (landingPageViews > 0) {
    return { trackedLeads: landingPageViews, trackedLeadType: "landing_page_view" };
  }

  if (inlineLinkClicks > 0) {
    return { trackedLeads: inlineLinkClicks, trackedLeadType: "link_click" };
  }

  return { trackedLeads: 0, trackedLeadType: "no_signal" };
}

function getDefaultSharedState() {
  return {
    products: [],
    tracking: [],
    customers: [],
    serviceForm: null,
    situationData: null,
    metaAdsState: null,
    importMeta: {
      lastOrdersImportAt: null,
      lastShippingImportAt: null,
    },
  };
}

async function ensureSharedStateFile() {
  await fs.mkdir(DATA_DIR, { recursive: true });
  try {
    await fs.access(SHARED_STATE_FILE);
  } catch {
    const initialPayload = {
      version: 0,
      updatedAt: null,
      state: getDefaultSharedState(),
    };
    await fs.writeFile(SHARED_STATE_FILE, JSON.stringify(initialPayload, null, 2), "utf8");
  }
}

async function readSharedState() {
  await ensureSharedStateFile();
  const raw = await fs.readFile(SHARED_STATE_FILE, "utf8");
  const parsed = JSON.parse(raw || "{}");
  return {
    version: Number(parsed.version || 0),
    updatedAt: parsed.updatedAt || null,
    state: parsed.state || getDefaultSharedState(),
  };
}

async function writeSharedState(nextState) {
  const current = await readSharedState();
  await fs.writeFile(SHARED_STATE_BACKUP_FILE, JSON.stringify(current, null, 2), "utf8");
  const payload = {
    version: current.version + 1,
    updatedAt: new Date().toISOString(),
    state: nextState || getDefaultSharedState(),
  };
  await fs.writeFile(SHARED_STATE_FILE, JSON.stringify(payload, null, 2), "utf8");
  return payload;
}

async function fetchMetaJson(url) {
  const response = await fetch(url);
  const payload = await response.json().catch(() => ({}));

  if (!response.ok || payload?.error) {
    const metaMessage = payload?.error?.message || `Meta API request failed (${response.status})`;
    const metaCode = payload?.error?.code;
    const metaSubcode = payload?.error?.error_subcode;
    const error = new Error(metaMessage);
    error.metaCode = metaCode;
    error.metaSubcode = metaSubcode;
    throw error;
  }

  return payload;
}

async function fetchMetaCollection(path, params) {
  const query = new URLSearchParams(params);
  let nextUrl = `${META_GRAPH_BASE}${path}?${query.toString()}`;
  const rows = [];

  while (nextUrl) {
    const payload = await fetchMetaJson(nextUrl);
    rows.push(...(payload?.data || []));
    nextUrl = payload?.paging?.next || "";
  }

  return rows;
}

app.get("/api/meta/health", (_req, res) => {
  res.json({
    ok: true,
    version: META_GRAPH_VERSION,
    baseUrl: META_GRAPH_BASE,
  });
});

app.get("/api/state/meta", async (_req, res) => {
  try {
    const payload = await readSharedState();
    return res.json({
      ok: true,
      version: payload.version,
      updatedAt: payload.updatedAt,
    });
  } catch (error) {
    return res.status(500).json({
      ok: false,
      error: error instanceof Error ? error.message : "Unable to read shared workspace meta.",
    });
  }
});

app.get("/api/state", async (_req, res) => {
  try {
    const payload = await readSharedState();
    return res.json({
      ok: true,
      version: payload.version,
      updatedAt: payload.updatedAt,
      state: payload.state,
    });
  } catch (error) {
    return res.status(500).json({
      ok: false,
      error: error instanceof Error ? error.message : "Unable to read shared workspace.",
    });
  }
});

app.post("/api/state", async (req, res) => {
  try {
    const state = req.body?.state;
    if (!state || typeof state !== "object") {
      return res.status(400).json({ ok: false, error: "State payload is required." });
    }

    const saved = await writeSharedState(state);
    return res.json({
      ok: true,
      version: saved.version,
      updatedAt: saved.updatedAt,
    });
  } catch (error) {
    return res.status(500).json({
      ok: false,
      error: error instanceof Error ? error.message : "Unable to save shared workspace.",
    });
  }
});

app.post("/api/meta/ad-accounts", async (req, res) => {
  try {
    const accessToken = String(req.body?.accessToken || "").trim();
    if (!accessToken) {
      return res.status(400).json({ ok: false, error: "Access token is required." });
    }

    const accounts = await fetchMetaCollection("/me/adaccounts", {
      access_token: accessToken,
      fields: "id,account_id,name,currency,timezone_name,account_status,business_country_code",
      limit: "200",
    });

    return res.json({
      ok: true,
      accounts: accounts.map((account) => ({
        id: account.id || account.account_id,
        accountId: account.account_id || account.id,
        name: account.name || account.id,
        currency: account.currency || "USD",
        timezoneName: account.timezone_name || "",
        accountStatus: account.account_status ?? null,
        businessCountryCode: account.business_country_code || "",
      })),
    });
  } catch (error) {
    const rawMessage = error instanceof Error ? error.message : "Unable to load ad accounts.";
    const isUnsupportedMeRequest =
      rawMessage.includes("Unsupported get request") ||
      rawMessage.includes("does not support this operation");

    return res.status(500).json({
      ok: false,
      error: isUnsupportedMeRequest
        ? "This token cannot list /me/adaccounts. Use a Meta user access token with ads_read permission, or enter your Ad Account ID manually in the app and go straight to Refresh insights."
        : rawMessage,
    });
  }
});

app.post("/api/meta/insights", async (req, res) => {
  try {
    const accessToken = String(req.body?.accessToken || "").trim();
    const accountId = ensureAccountId(req.body?.accountId);
    const since = String(req.body?.since || "").trim();
    const until = String(req.body?.until || "").trim();

    if (!accessToken || !accountId || !since || !until) {
      return res.status(400).json({ ok: false, error: "Access token, account and date range are required." });
    }

    const rows = await fetchMetaCollection(`/${accountId}/insights`, {
      access_token: accessToken,
      level: "campaign",
      fields: "campaign_id,campaign_name,objective,adset_id,adset_name,spend,impressions,reach,clicks,ctr,cpc,cpp,cpm,frequency,inline_link_clicks,unique_inline_link_clicks,actions,cost_per_action_type",
      time_range: JSON.stringify({ since, until }),
      limit: "500",
    });

    const normalizedRows = rows.map((row) => {
      const spend = toNumber(row.spend);
      const impressions = Math.round(toNumber(row.impressions));
      const reach = Math.round(toNumber(row.reach));
      const clicks = Math.round(toNumber(row.clicks));
      const linkClickMetric = pickActionMetric(row.actions || [], ["link_click", "inline_link_click", "outbound_click"]);
      const inlineLinkClicks = Math.round(toNumber(row.inline_link_clicks) || linkClickMetric.value);
      const uniqueInlineLinkClicks = Math.round(toNumber(row.unique_inline_link_clicks));
      const ctr = impressions > 0 ? (inlineLinkClicks / impressions) * 100 : 0;
      const cpc = inlineLinkClicks > 0 ? spend / inlineLinkClicks : toNumber(row.cpc);
      const cpp = toNumber(row.cpp);
      const cpm = toNumber(row.cpm);
      const frequency = toNumber(row.frequency);
      const leadMetric = pickLeadValue(row.actions || []);
      const landingPageViewMetric = pickActionMetric(row.actions || [], ["landing_page_view", "onsite_web_landing_page_view"]);
      const actualLeads = leadMetric.leads;
      const landingPageViews = Math.round(landingPageViewMetric.value);
      const trackedLeadMetric = pickTrackedLeadMetric({
        actualLeads,
        landingPageViews,
        inlineLinkClicks,
      });
      const trackedLeads = trackedLeadMetric.trackedLeads;
      const costPerLead = trackedLeads > 0 ? spend / trackedLeads : 0;

      return {
        id: String(row.campaign_id || row.adset_id || row.campaign_name || Math.random()),
        campaignId: String(row.campaign_id || ""),
        campaignName: row.campaign_name || "Unnamed campaign",
        objective: row.objective || "",
        adsetId: String(row.adset_id || ""),
        adsetName: row.adset_name || "",
        spend,
        impressions,
        reach,
        clicks,
        inlineLinkClicks,
        uniqueInlineLinkClicks,
        linkClickType: linkClickMetric.actionType,
        ctr,
        cpc,
        cpp,
        cpm,
        frequency,
        leads: actualLeads,
        actualLeads,
        leadType: leadMetric.leadType,
        trackedLeads,
        trackedLeadType: trackedLeadMetric.trackedLeadType,
        landingPageViews,
        costPerLead,
      };
    });

    const summary = normalizedRows.reduce(
      (acc, row) => {
        acc.spend += row.spend;
        acc.impressions += row.impressions;
        acc.reach += row.reach;
        acc.clicks += row.clicks;
        acc.inlineLinkClicks += row.inlineLinkClicks;
        acc.uniqueInlineLinkClicks += row.uniqueInlineLinkClicks;
        acc.landingPageViews += row.landingPageViews;
        acc.leads += row.actualLeads;
        acc.actualLeads += row.actualLeads;
        acc.trackedLeads += row.trackedLeads;
        acc.trackedLeadSources[row.trackedLeadType] = (acc.trackedLeadSources[row.trackedLeadType] || 0) + row.trackedLeads;
        return acc;
      },
      {
        spend: 0,
        impressions: 0,
        reach: 0,
        clicks: 0,
        inlineLinkClicks: 0,
        uniqueInlineLinkClicks: 0,
        landingPageViews: 0,
        leads: 0,
        actualLeads: 0,
        trackedLeads: 0,
        trackedLeadSources: {},
      }
    );

    const activeTrackedLeadSources = Object.entries(summary.trackedLeadSources).filter(([, value]) => Number(value || 0) > 0);
    const trackedLeadSource =
      activeTrackedLeadSources.length === 1 ? activeTrackedLeadSources[0][0] : activeTrackedLeadSources.length > 1 ? "mixed" : "no_signal";

    return res.json({
      ok: true,
      summary: {
        ...summary,
        ctr: summary.impressions > 0 ? (summary.inlineLinkClicks / summary.impressions) * 100 : 0,
        cpc: summary.inlineLinkClicks > 0 ? summary.spend / summary.inlineLinkClicks : 0,
        cpm: summary.impressions > 0 ? (summary.spend / summary.impressions) * 1000 : 0,
        cpp: summary.reach > 0 ? summary.spend / summary.reach : 0,
        frequency: summary.reach > 0 ? summary.impressions / summary.reach : 0,
        costPerLead: summary.trackedLeads > 0 ? summary.spend / summary.trackedLeads : 0,
        trackedLeadSource,
      },
      rows: normalizedRows,
    });
  } catch (error) {
    return res.status(500).json({
      ok: false,
      error: error instanceof Error ? error.message : "Unable to load Meta insights.",
    });
  }
});

app.post("/api/meta/spend-total", async (req, res) => {
  try {
    const accessToken = String(req.body?.accessToken || "").trim();
    const accountId = ensureAccountId(req.body?.accountId);

    if (!accessToken || !accountId) {
      return res.status(400).json({ ok: false, error: "Access token and account are required." });
    }

    const rows = await fetchMetaCollection(`/${accountId}/insights`, {
      access_token: accessToken,
      fields: "spend",
      date_preset: "maximum",
      limit: "5",
    });

    const totalSpend = rows.reduce((sum, row) => sum + toNumber(row?.spend), 0);

    return res.json({
      ok: true,
      spend: totalSpend,
      accountId,
      datePreset: "maximum",
      capturedAt: new Date().toISOString(),
    });
  } catch (error) {
    return res.status(500).json({
      ok: false,
      error: error instanceof Error ? error.message : "Unable to load Meta total spend.",
    });
  }
});

app.post("/api/meta/spend-daily", async (req, res) => {
  try {
    const accessToken = String(req.body?.accessToken || "").trim();
    const accountId = ensureAccountId(req.body?.accountId);
    const date = String(req.body?.date || "").trim();

    if (!accessToken || !accountId || !date) {
      return res.status(400).json({ ok: false, error: "Access token, account and date are required." });
    }

    const rows = await fetchMetaCollection(`/${accountId}/insights`, {
      access_token: accessToken,
      fields: "spend",
      time_range: JSON.stringify({ since: date, until: date }),
      limit: "5",
    });

    const totalSpend = rows.reduce((sum, row) => sum + toNumber(row?.spend), 0);

    return res.json({
      ok: true,
      spend: totalSpend,
      accountId,
      date,
      capturedAt: new Date().toISOString(),
    });
  } catch (error) {
    return res.status(500).json({
      ok: false,
      error: error instanceof Error ? error.message : "Unable to load Meta daily spend.",
    });
  }
});

app.listen(PORT, () => {
  console.log(`Meta API bridge listening on http://127.0.0.1:${PORT}`);
});
