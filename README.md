# Sport Factory Panamá — KPI Dashboard

Interactive, offline-capable dashboard for the Sport Factory store KPIs
(VENTAS, UDS, FACT, UPT, VPT) across stores and years.

**Live page:** _(GitHub Pages link appears here once published)_

## Files
- `index.html` / `dashboard.html` — the dashboard. Self-contained, works offline; the
  data is embedded (encrypted) in the file. `index.html` is the copy GitHub Pages serves.
- `build_dashboard.py` — regenerates the dashboard from the Excel workbooks + database.
- `build_kpi_2024_from_db.py` — generates `SPORT FACTORY_KPI 2024.xlsx` from ICG.

## What the dashboard shows
- **Two tabs:** *KPIs* (VENTAS, UDS, FACT, UPT, VPT, **MARGEN %**) and
  *Presupuesto vs Real* (budget vs actual sales, per store + brand totals).
- **Period filter:** Año completo / Q1 (Ene–Jun) / Q2 (Jul–Dic).
- Year-over-year, monthly trend, all-KPIs-together, and per-store tables.

## Data sources
- `SPORT FACTORY_KPI <year>.xlsx` → VENTAS/UDS/FACT/UPT/VPT per store/month.
- ICG **[NEWSPF]** (local SQL Server) → **MARGEN %** per store/month.
- `Presupuesto de ventas SF 2026_JP_actualizado.xlsx` (sheet *Presupuesto de vts 2026*)
  → the monthly budget column, for the Presupuesto vs Real tab.

## Build requirements
The build needs the ICG database (for margin), so run it with the Python that has
`openpyxl`, `cryptography` **and** `pyodbc`, with SQL Server running locally:
```
"C:\Claude local\Database\.venv\Scripts\python.exe" build_dashboard.py
```
If the database is unreachable the dashboard still builds — only the margin KPI is
omitted (a warning is printed).

## Password protection
The data is **encrypted** inside the HTML (AES-256-GCM, key derived from the
passphrase via PBKDF2-SHA256, all done in the browser with the Web Crypto API).
The page opens to a passphrase prompt; nothing is readable — not even in the
page source on GitHub — without the passphrase. This is what makes it safe to
keep the repo public: visitors see only an encrypted blob.

- Share the **link** and the **passphrase separately** (e.g. link by email, passphrase by phone).
- To change/rotate the passphrase, just rebuild (see below) with a new one and push.
- Anyone who has the passphrase can decrypt; "revoking" means rebuilding with a new
  passphrase and re-sharing.

## Updating the data (or changing the passphrase)
1. Update the `SPORT FACTORY_KPI <year>.xlsx` files (same layout as before).
2. Run the build. It will ask for the passphrase (hidden input):
   ```
   python build_dashboard.py
   ```
   You can also pass it without a prompt:
   ```
   # Windows PowerShell
   $env:DASH_PASSPHRASE="your-passphrase"; python build_dashboard.py
   ```
   This rewrites `index.html` and `dashboard.html`.
3. Commit and push to update the live page:
   ```
   git add -A && git commit -m "Update KPIs" && git push
   ```

## Notes
- Year-over-year and YTD figures compare only the months present in **both** years
  (so a partial current year is compared like-for-like).
- UPT and VPT are recalculated over the accumulated totals, not averaged.
- The raw `.xlsx` files are git-ignored by default (see `.gitignore`).
