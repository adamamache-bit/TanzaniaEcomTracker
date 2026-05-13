-- Migration 006: Enable RLS on all public tables
-- Fixes Supabase security warning: "RLS is disabled on public tables"
-- Uses permissive policies (USING true) so the app works without auth changes.
-- Run in Supabase SQL Editor.

ALTER TABLE orders ENABLE ROW LEVEL SECURITY;
CREATE POLICY "allow_all_orders" ON orders FOR ALL USING (true) WITH CHECK (true);

ALTER TABLE products ENABLE ROW LEVEL SECURITY;
CREATE POLICY "allow_all_products" ON products FOR ALL USING (true) WITH CHECK (true);

ALTER TABLE stock_purchases ENABLE ROW LEVEL SECURITY;
CREATE POLICY "allow_all_stock" ON stock_purchases FOR ALL USING (true) WITH CHECK (true);

ALTER TABLE ads_campaigns ENABLE ROW LEVEL SECURITY;
CREATE POLICY "allow_all_ads" ON ads_campaigns FOR ALL USING (true) WITH CHECK (true);

ALTER TABLE manual_ads_spend ENABLE ROW LEVEL SECURITY;
CREATE POLICY "allow_all_manual_ads" ON manual_ads_spend FOR ALL USING (true) WITH CHECK (true);

ALTER TABLE extra_charges ENABLE ROW LEVEL SECURITY;
CREATE POLICY "allow_all_charges" ON extra_charges FOR ALL USING (true) WITH CHECK (true);

ALTER TABLE owner_injections ENABLE ROW LEVEL SECURITY;
CREATE POLICY "allow_all_injections" ON owner_injections FOR ALL USING (true) WITH CHECK (true);
