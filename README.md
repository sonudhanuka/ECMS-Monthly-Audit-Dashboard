# ECMS Monthly Audit Dashboard

Auto-built HTML dashboard hosted on GitHub Pages.
Data is read from SharePoint and injected into the original HTML dashboard nightly.

## Repository structure

```
your-repo/
├── ECMS_Monthly_Audit_Dashboard.html   ← your original HTML (rename to this)
├── build_dashboard.py                  ← reads SharePoint, injects data, saves HTML
├── requirements.txt                    ← Python packages
├── README.md
├── docs/
│   └── index.html                      ← built output served by GitHub Pages
└── .github/
    └── workflows/
        └── build.yml                   ← runs build_dashboard.py on schedule
```

## One-time setup (15 minutes)

### 1. Add your SharePoint URL as a secret

Go to: **Settings → Secrets and variables → Actions → New repository secret**
- Name: `SHAREPOINT_FOLDER_URL`
- Value: your SharePoint folder share URL (Anyone with link → Copy)

### 2. Enable GitHub Pages

Go to: **Settings → Pages**
- Source: Deploy from a branch
- Branch: `main` · Folder: `/docs`
- Save

### 3. Run the first build

Go to: **Actions → Build Dashboard → Run workflow → Run workflow**

Your dashboard is now live at: `https://YOUR-USERNAME.github.io/YOUR-REPO-NAME`

## Monthly workflow (after setup)

1. Get your new monthly audit Excel file
2. Drop it into your SharePoint folder
3. GitHub Actions runs at 1am and rebuilds the dashboard automatically

Or trigger immediately: **Actions → Build Dashboard → Run workflow**

## Excel file format

Each monthly file must have these column headers:

| Column | Example values |
|--------|---------------|
| `Week` | CW14, CW15, CW16 |
| `Ticket Number` | INC56052293 |
| `Audit Parameter` | Closure Information |
| `Evaluation` | **Met** or **Not Met** |
| `Tower` | IMS |
| `Teams` | Workplace, Network, Cloud |
| `Agent` | Komal Rao |
| `Comments` | New closure template not followed |

## Changing the schedule

Edit `.github/workflows/build.yml`, line 8:
```yaml
- cron: '0 1 * * *'    # every night at 1am UTC (current)
- cron: '0 1 * * 1'    # every Monday at 1am
- cron: '0 1 1 * *'    # first day of each month at 1am
```
