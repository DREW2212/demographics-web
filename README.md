# Demographics Merge Tool

A web app that merges three demographic radius reports (1-mile, 3-mile, 5-mile) into a single formatted Full Service workbook. Built in Python with Flask.

---

## What It Does

Users upload three Esri/demographic report Excel files — one for each radius (1 mi, 3 mi, 5 mi) — and the tool merges them into a single output workbook using the Full Service Demographic Template. The output file is downloaded directly to the user's computer.

### Sheet Mapping

Each radius file has these source sheets copied into the template:

| Source Sheet | Destination Tab in Template |
|---|---|
| Summary | `1 mile summary tab` / `3 mile summary` / `5 mile summary` |
| Demographic Quick Facts | `1 mile demo quick` / `3 mile demo quick` / `5 mile demo quick` |
| 2025 | `1 mile 2025` / `3 mile 2025` / `5 mile 2025` |
| 2030 | `1 mile 2030` / `3 mile 2030` / `5 mile 2030` |
| Income by Age of Household | `1 mile age of HHer` / `3 mile age of HHer` / `5 mile age of HHer` |

Cell formatting, merged cells, column widths, and row heights are all preserved. If an address is provided (or detected from the filename), it is written to cell A1 of the Summary tab.

---

## File Structure

```
demographics-web/
├── app.py                  # Flask web server — handles uploads and serves the UI
├── merger.py               # Core merge logic (DemographicsMerger class)
├── requirements.txt        # Python dependencies
├── Procfile                # Tells Railway/Heroku how to start the app
├── runtime.txt             # Python version for deployment
├── template/
│   └── Full_Service_Demographic_Template.xlsx   # Bundled template
└── templates/
    └── index.html          # Frontend UI (single page, vanilla HTML/CSS/JS)
```

---

## How to Run Locally

### 1. Install Python
Download Python 3.10+ from [python.org](https://python.org). Make sure to check "Add to PATH" during install.

### 2. Install dependencies
Open a terminal in the `demographics-web` folder and run:
```
pip install -r requirements.txt
```

### 3. Start the app
```
python app.py
```

Open your browser to `http://localhost:5000`. The password is `AGI`.

---

## How to Deploy (Railway)

The live site is hosted on [Railway](https://railway.app) and connected to this GitHub repo. Railway auto-deploys whenever new code is pushed to the `master` branch.

### To redeploy after making changes:
1. Edit the files locally
2. Run in the `demographics-web` folder:
```
git add .
git commit -m "describe your change"
git push
```
Railway picks up the push and redeploys automatically (takes ~2 minutes).

### To set up hosting from scratch (if the Railway project is ever lost):
1. Go to [railway.app](https://railway.app) and sign in with GitHub
2. New Project → Deploy from GitHub repo → select `DREW2212/demographics-web`
3. Railway detects Flask automatically and deploys it
4. Go to Settings → Domains → add your custom domain
5. Update DNS at your domain registrar to point to the Railway-provided CNAME

---

## How to Update the Template

If the Full Service Demographic Template changes:
1. Replace `template/Full_Service_Demographic_Template.xlsx` with the new version
2. Make sure the sheet names inside it still match the destination tab names in the Sheet Mapping table above
3. Push to GitHub — the site redeploys automatically

---

## How to Change the Password

Open `templates/index.html` and find this line near the top of the `<script>` block:
```js
const PASSWORD = 'AGI';
```
Change `AGI` to whatever you want, then push to GitHub.

---

## Dependencies

| Package | Purpose |
|---|---|
| Flask | Web framework — handles HTTP requests and file uploads |
| openpyxl | Reads and writes Excel files, copies sheet data and formatting |
| gunicorn | Production web server (used by Railway) |

---

## Updating Dependencies

If you need to add or upgrade a package:
```
pip install <package-name>
pip freeze > requirements.txt
git add requirements.txt
git commit -m "update dependencies"
git push
```
