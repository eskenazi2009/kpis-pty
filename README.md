# Sport Factory Panamá — KPI Dashboard

Interactive, offline-capable dashboard for the Sport Factory store KPIs
(VENTAS, UDS, FACT, UPT, VPT) across stores and years.

**Live page:** _(GitHub Pages link appears here once published)_

## Files
- `index.html` / `dashboard.html` — the dashboard. Self-contained, works offline; the
  data is embedded directly in the file. `index.html` is the copy GitHub Pages serves.
- `build_dashboard.py` — regenerates the dashboard from the Excel workbooks.

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
