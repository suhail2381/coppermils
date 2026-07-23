# ⚡ Wire & Cable Business Portal

A free, single-file Streamlit application for a wire/cable manufacturing business. It combines
a **public storefront** (customers browse products and place orders — no login needed) with a
**staff/admin portal** covering pricing, raw materials, production, scrap, orders, customers,
staff, and reports.

## What's included

| Area | What it does |
|---|---|
| 🛍️ Storefront | Product grid with filters, cart, checkout form that auto-creates the customer record and the order |
| 📦 Product Catalog & Pricing | Edit prices any time (with a price-change history log), add new products |
| 🧵 Raw Materials | Stock for Copper (Pure / Loose Gauge / China), Aluminium, PVC (FR/HR/Weatherproof), XLPE — log Purchase / Consumption / Scrap / Adjustment |
| 🏭 Production Entry | Log a production run: product, qty produced, qty rejected/broken, supervisor, production manager, labour count, and which raw materials (and how much) were consumed. Stock updates automatically. |
| 🗑️ Scrap / Condemned Material | Manual log for condemned/broken raw material (copper, PVC, XLPE, aluminium), plus auto-entries from rejected production |
| 🧾 Orders & Sales | View/update order status, inspect order line items, sales trend chart |
| 👥 Customers | Auto-built customer list from storefront checkouts |
| 🧑‍🏭 Staff | Labour / Supervisor / Production Manager records |
| 📈 Reports | Product-wise stock, raw-material-wise stock, sales-wise, and scrap-wise — each with an Excel export |

Product data is pre-seeded from your `wire_detail.xlsx` (aluminium & copper building-wire specs),
plus starter entries for Power Cables and Cat6 Networking Cable — all editable/expandable from
the Product Catalog page.

## Roles

Set staff logins under `[users]` in Streamlit secrets (see `secrets_example.toml`):

- **Admin** — sees and edits everything
- **Production Manager** — raw materials, production, scrap, staff, reports
- **Sales** — pricing (view), orders, customers, reports

The storefront needs no account at all.

## Local setup

```bash
pip install -r requirements.txt
mkdir -p .streamlit
cp secrets_example.toml .streamlit/secrets.toml   # then edit with real passwords
streamlit run app.py
```

## Deploying to Streamlit Community Cloud (free)

1. Push `app.py`, `requirements.txt`, and `.gitignore` to a GitHub repo (do **not** push
   `secrets.toml` or the `.db` file — `.gitignore` already excludes them).
2. On [share.streamlit.io](https://share.streamlit.io), create a new app pointing at the repo.
3. Under the app's **Settings → Secrets**, paste the contents of `secrets_example.toml` filled
   in with real staff passwords.
4. Deploy.

## Important: database persistence

This app uses a local **SQLite** file (`wire_cable_erp.db`) for zero setup cost. On Streamlit
Community Cloud, the filesystem resets whenever the app restarts or redeploys — so data can be
lost over time. Two ways to make it durable:

- **Simple**: from the Reports pages, regularly download the Excel exports as backups.
- **Robust** (recommended once you have real customers/orders): add a small Firebase Storage
  sync step, the same pattern used in your other portal — download the `.db` file from Firebase
  Storage on startup, upload it back after every write. Happy to wire this in if you'd like.

## Customizing

- Change `BUSINESS_NAME` and `CURRENCY` at the top of `app.py`.
- Add/edit products directly from **Product Catalog & Pricing** in the app — no code changes needed.
- Add labour/supervisors/production managers from the **Staff** page before using Production Entry
  (the dropdowns pull from active staff records).
