# Run several markets, each with its own live link

One codebase can power many market dashboards - each its own Render
service and URL - by pinning each service to a dataset with the
`MARKET_DATA` environment variable. Code fixes you push apply to every
market automatically (no duplicated code).

## One-time: pin the existing service

Once the repo holds MORE than one `Dataset_*.xlsx`, a service with no
`MARKET_DATA` picks the alphabetically-first file - which can change
which market it shows. So pin the current one first:

- Render → **fmi-market-report** service → **Environment** →
  Add `MARKET_DATA` = `Dataset_Global Stretch Hood Film Market.xlsx`
  (the exact filename of its current dataset) → **Save** (it redeploys).

## Add a new market link

1. **Put the new dataset in the repo and push** (from the project folder):
   ```bash
   cp "C:\path\to\Dataset_Global Green Tea Market.xlsx" .
   git add -A
   git commit -m "Add Green Tea dataset"
   git push
   ```
   The repo now has both datasets.

2. **Create a new Render service from the SAME repo:**
   - Render → **New + → Web Service** → pick `fmi_report_generator`.
   - **Name**: `fmi-green-tea` (becomes `https://fmi-green-tea.onrender.com`).
   - **Build command**: `pip install -r requirements.txt`
   - **Start command**:
     `gunicorn run:server --workers 1 --threads 4 --timeout 120`
   - **Environment variables**:
     - `MARKET_DATA` = `Dataset_Global Green Tea Market.xlsx`
     - `PYTHON_VERSION` = `3.11.9`
   - **Create Web Service** → it builds and gets its own link.

3. **(Optional) branded subdomain** for it, e.g.
   `greentea.futuremarketinsights.com` - Render service → Settings →
   Custom Domains, then add the CNAME in DNS (same as the main site).

`MARKET_DATA` may be just the filename (it is resolved inside the repo)
or a full path. Each new market = repeat steps 1-2 with its dataset.
