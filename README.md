
# BSNL FTTH Warangal OA – Monthly Total Connections Dashboard

Streamlit + Pandas application for generating a downloadable Excel dashboard from a monthly
**Total Connections** workbook.

## Input

Upload any Excel filename, for example:

`TOTAL CONNECTIONS AUG2026.xlsx`

The source must contain these columns:

- `OLT IP`
- `FR Service Code`

All original source columns are retained in the `Data` sheet.

## Mapping logic

The application reads two editable Excel master files from the repository root:

1. `OLT_BBC_MAP.xlsx`
   - Maps `OLT IP` → `BBC`

2. `BBC_Master.xlsx`
   - Maps `BBC` → `DE / Manager`
   - Optional `Area / TIP`
   - Optional `Monthly Target`

The code accepts common alternate header names, so the mapping files do not have to use one exact spelling.

### Important

Whenever OLT/BBC/DE/Manager mapping changes:

1. Replace `OLT_BBC_MAP.xlsx` in GitHub.
2. Replace `BBC_Master.xlsx` in GitHub.
3. Redeploy/reboot Streamlit.

**No Python code change is required.**

Unmapped OLTs are not discarded. They are shown as `UNMAPPED` and listed in the dashboard.

## Generated Excel workbook

The downloadable workbook contains:

- `Dashboard` – executive KPI page
- `BBC_Report` – BBC-wise connections, OLTs and franchisees
- `DE_Manager_Report` – DE / Manager-wise connections
- `Franchisee_Report` – FR Service Code-wise connections
- `OLT_Report` – OLT IP-wise connections with BBC and DE / Manager
- `Data` – every source row retained with mapping columns added
- `Charts` – embedded Excel charts
- `OLT_BBC_Map` – copy of the mapping used
- `BBC_Master` – copy of the master used

## GitHub structure

```text
bsnl-ftth-monthly-dashboard/
├── app.py
├── report_generator.py
├── requirements.txt
├── README.md
├── .gitignore
├── OLT_BBC_MAP.xlsx
└── BBC_Master.xlsx
```

## Streamlit Cloud

1. Create a GitHub repository.
2. Upload all files from this folder.
3. Set the main file to `app.py`.
4. Deploy on Streamlit Community Cloud.
5. Upload the monthly Total Connections Excel file in the app.
6. Click **GENERATE MONTHLY EXCEL DASHBOARD**.
7. Download the generated workbook.

## Master file templates

The repository includes blank template files. Replace them with your actual masters.

### OLT_BBC_MAP.xlsx

Required columns:

| OLT IP | BBC Name |
|---|---|
| 10.x.x.x | BBC NAME |

### BBC_Master.xlsx

Recommended columns:

| BBC Name | DE / Manager | Area / TIP | Monthly Target |
|---|---|---|---:|
| BBC NAME | OFFICER NAME | TIP | 0 |

