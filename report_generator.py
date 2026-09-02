
from __future__ import annotations

from io import BytesIO
from pathlib import Path
import re
import pandas as pd
import xlsxwriter


def norm_col(x):
    return re.sub(r"[^A-Z0-9]+", " ", str(x).strip().upper()).strip()


def pick_col(df, aliases, required=True):
    lookup = {norm_col(c): c for c in df.columns}
    for alias in aliases:
        key = norm_col(alias)
        if key in lookup:
            return lookup[key]
    if required:
        raise ValueError(
            f"Required column not found. Expected one of {aliases}. "
            f"Found: {list(df.columns)}"
        )
    return None


def read_source(source_bytes):
    bio = BytesIO(source_bytes)
    xl = pd.ExcelFile(bio)
    sheet = "Sheet0" if "Sheet0" in xl.sheet_names else xl.sheet_names[0]

    # First try row 1 headers. If required fields are not found, inspect first 10 rows.
    preview = pd.read_excel(BytesIO(source_bytes), sheet_name=sheet, header=None, nrows=10)
    wanted = {"OLT IP", "FR Service Code"}

    header_row = None
    for i in range(min(10, len(preview))):
        vals = {norm_col(v) for v in preview.iloc[i].tolist()}
        if norm_col("OLT IP") in vals and norm_col("FR Service Code") in vals:
            header_row = i
            break

    if header_row is None:
        raise ValueError(
            "Could not locate the source header row. The workbook must contain "
            "'OLT IP' and 'FR Service Code'."
        )

    df = pd.read_excel(BytesIO(source_bytes), sheet_name=sheet, header=header_row)
    df = df.dropna(how="all").copy()
    df.columns = [str(c).strip() for c in df.columns]
    pick_col(df, ["OLT IP"])
    pick_col(df, ["FR Service Code"])
    return df


def read_master(path, required=True):
    if not path.exists():
        if required:
            raise FileNotFoundError(
                f"{path.name} is missing. Put it in the GitHub repository root."
            )
        return pd.DataFrame()
    return pd.read_excel(path, sheet_name=0)


def load_mappings(map_path, bbc_master_path):
    olt = read_master(map_path)
    bbc = read_master(bbc_master_path)

    olt_ip = pick_col(olt, ["OLT IP", "OLT_IP", "OLT", "IP"])
    olt_bbc = pick_col(olt, ["BBC Name", "BBC", "BBM Name", "Employee", "BBC_NAME"])

    olt_map = olt[[olt_ip, olt_bbc]].copy()
    olt_map.columns = ["OLT IP", "BBC"]
    olt_map["OLT IP"] = olt_map["OLT IP"].fillna("").astype(str).str.strip()
    olt_map["BBC"] = olt_map["BBC"].fillna("UNMAPPED").astype(str).str.strip()
    olt_map = olt_map.drop_duplicates("OLT IP", keep="last")

    bbc_name = pick_col(bbc, ["BBC Name", "BBC", "BBM Name", "Employee"])
    manager = pick_col(bbc, ["DE / Manager", "DE/Manager", "Manager", "MT", "DE", "AGM/ Manager(MT)"], required=False)
    area = pick_col(bbc, ["Area / TIP", "Area/TIP", "Area", "TIP"], required=False)
    target = pick_col(bbc, ["Monthly Target", "BBC Target", "Target"], required=False)

    info = bbc[[bbc_name]].copy()
    info.columns = ["BBC"]

    if manager:
        info["DE / Manager"] = bbc[manager]
    else:
        info["DE / Manager"] = "UNMAPPED"

    if area:
        info["Area / TIP"] = bbc[area]
    else:
        info["Area / TIP"] = ""

    if target:
        info["Monthly Target"] = pd.to_numeric(bbc[target], errors="coerce").fillna(0)
    else:
        info["Monthly Target"] = 0

    for c in ["BBC", "DE / Manager", "Area / TIP"]:
        info[c] = info[c].fillna("").astype(str).str.strip()

    info["BBC"] = info["BBC"].replace("", "UNMAPPED")
    info["DE / Manager"] = info["DE / Manager"].replace("", "UNMAPPED")
    info = info.drop_duplicates("BBC", keep="last")

    return olt_map, info


def add_mappings(df, olt_map, bbc_info):
    olt_col = pick_col(df, ["OLT IP"])
    fr_col = pick_col(df, ["FR Service Code"])

    out = df.copy()
    out[olt_col] = out[olt_col].fillna("").astype(str).str.strip()
    out[fr_col] = out[fr_col].fillna("").astype(str).str.strip()

    out["BBC"] = out[olt_col].map(dict(zip(olt_map["OLT IP"], olt_map["BBC"]))).fillna("UNMAPPED")

    info_map = bbc_info.set_index("BBC").to_dict("index")
    out["DE / Manager"] = out["BBC"].map(lambda x: info_map.get(x, {}).get("DE / Manager", "UNMAPPED"))
    out["Area / TIP"] = out["BBC"].map(lambda x: info_map.get(x, {}).get("Area / TIP", ""))
    out["Monthly Target"] = out["BBC"].map(lambda x: info_map.get(x, {}).get("Monthly Target", 0))

    out = out.rename(columns={olt_col: "OLT IP", fr_col: "FR Service Code"})
    return out


def aggregate(df):
    bbc = (
        df.groupby("BBC", dropna=False)
        .agg(Connections=("BBC", "size"), OLT_IPs=("OLT IP", "nunique"), Franchisees=("FR Service Code", "nunique"))
        .reset_index()
    )

    bbc = bbc.merge(
        df[["BBC", "DE / Manager", "Area / TIP", "Monthly Target"]].drop_duplicates("BBC"),
        on="BBC", how="left"
    )
    bbc["Achievement %"] = bbc.apply(
        lambda r: r["Connections"] / r["Monthly Target"] if r["Monthly Target"] else 0,
        axis=1
    )
    bbc = bbc[["DE / Manager", "Area / TIP", "BBC", "Connections", "OLT_IPs", "Franchisees", "Monthly Target", "Achievement %"]]
    bbc = bbc.sort_values(["DE / Manager", "Connections"], ascending=[True, False])

    manager = (
        df.groupby("DE / Manager", dropna=False)
        .agg(Connections=("BBC", "size"), BBCs=("BBC", "nunique"), OLT_IPs=("OLT IP", "nunique"), Franchisees=("FR Service Code", "nunique"))
        .reset_index()
        .sort_values("Connections", ascending=False)
    )

    franchisee = (
        df.groupby("FR Service Code", dropna=False)
        .agg(Connections=("FR Service Code", "size"), BBCs=("BBC", "nunique"), OLT_IPs=("OLT IP", "nunique"))
        .reset_index()
        .sort_values("Connections", ascending=False)
    )

    olt = (
        df.groupby(["OLT IP", "BBC", "DE / Manager", "Area / TIP"], dropna=False)
        .agg(Connections=("OLT IP", "size"), Franchisees=("FR Service Code", "nunique"))
        .reset_index()
        .sort_values("Connections", ascending=False)
    )

    return bbc, manager, franchisee, olt


def write_excel(df, bbc, manager, franchisee, olt, olt_map, bbc_info, source_name):
    out = BytesIO()

    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        wb = writer.book

        navy = "#102A43"
        blue = "#1F4EBA"
        light = "#F4F7FB"
        white = "#FFFFFF"
        green = "#E8F8E8"
        red = "#FDE8E8"
        border = "#D9E1EA"

        title = wb.add_format({
            "bold": True, "font_size": 22, "font_color": white,
            "bg_color": navy, "align": "center", "valign": "vcenter"
        })
        subtitle = wb.add_format({
            "italic": True, "font_color": "#667085", "bg_color": light,
            "align": "center"
        })
        header = wb.add_format({
            "bold": True, "font_color": white, "bg_color": navy,
            "border": 1, "border_color": border, "align": "center",
            "valign": "vcenter", "text_wrap": True
        })
        cell = wb.add_format({"border":1, "border_color":border, "align":"center"})
        pct = wb.add_format({"num_format":"0.0%", "border":1, "border_color":border, "align":"center"})
        good = wb.add_format({"bg_color":green, "font_color":"#087A25", "bold":True, "border":1, "border_color":border, "align":"center"})
        bad = wb.add_format({"bg_color":red, "font_color":"#C00000", "bold":True, "border":1, "border_color":border, "align":"center"})

        # Dashboard
        dash = wb.add_worksheet("Dashboard")
        dash.hide_gridlines(2)
        dash.set_column("A:A", 3)
        dash.set_column("B:Z", 14)
        dash.merge_range("B2:Z4", "BSNL FTTH WARANGAL OA – MONTHLY TOTAL CONNECTIONS DASHBOARD", title)
        dash.merge_range("B5:Z5", f"Source: {source_name}", subtitle)

        total = len(df)
        olt_count = df["OLT IP"].nunique()
        fr_count = df["FR Service Code"].nunique()
        bbc_count = df["BBC"].nunique()
        mgr_count = df["DE / Manager"].nunique()
        unmapped = int((df["BBC"] == "UNMAPPED").sum())

        cards = [
            ("TOTAL CONNECTIONS", total),
            ("OLT IPs", olt_count),
            ("FRANCHISEES", fr_count),
            ("BBCs", bbc_count),
            ("DE / MANAGERS", mgr_count),
            ("UNMAPPED CONNECTIONS", unmapped),
        ]
        for i, (lab, val) in enumerate(cards):
            col = 1 + i * 4
            dash.merge_range(6, col, 7, col+2, f"{val:,}", wb.add_format({
                "bold":True, "font_size":20, "font_color":blue, "bg_color":white,
                "border":1, "border_color":"#E0E0E0", "align":"center"
            }))
            dash.merge_range(8, col, 8, col+2, lab, wb.add_format({
                "font_size":9, "font_color":"#667085", "bg_color":white,
                "border":1, "border_color":"#E0E0E0", "align":"center"
            }))

        dash.write("B11", "Quick navigation", wb.add_format({"bold":True, "font_size":14, "font_color":navy}))
        links = [("BBC_Report","BBC-wise report"),("DE_Manager_Report","DE / Manager report"),
                 ("Franchisee_Report","FR Service Code report"),("OLT_Report","OLT IP report"),
                 ("Charts","Charts")]
        for i,(sheet,label) in enumerate(links, start=12):
            dash.write_url(i, 1, f"internal:'{sheet}'!A1", string=label)

        # Data and reports
        sheets = [
            ("Data", df),
            ("BBC_Report", bbc),
            ("DE_Manager_Report", manager),
            ("Franchisee_Report", franchisee),
            ("OLT_Report", olt),
            ("OLT_BBC_Map", olt_map),
            ("BBC_Master", bbc_info),
        ]
        for name, data in sheets:
            data.to_excel(writer, sheet_name=name, index=False)
            ws = writer.sheets[name]
            ws.hide_gridlines(2)
            ws.freeze_panes(1, 0)
            ws.set_tab_color(blue)
            for j, c in enumerate(data.columns):
                vals = data[c].head(200).fillna("").astype(str)
                width = min(max(12, max([len(str(c))] + [len(v) for v in vals]) + 2), 36)
                ws.set_column(j, j, width)
                ws.write(0, j, c, header)
            if len(data):
                ws.add_table(0, 0, len(data), len(data.columns)-1, {
                    "style":"Table Style Medium 2",
                    "columns":[{"header":c} for c in data.columns]
                })
            if "Achievement %" in data.columns:
                j = list(data.columns).index("Achievement %")
                ws.set_column(j,j,15,pct)

        # Conditional highlighting for unmapped data
        data_ws = writer.sheets["Data"]
        if "BBC" in df.columns and len(df):
            j = list(df.columns).index("BBC")
            data_ws.conditional_format(1, j, len(df), j, {
                "type":"text", "criteria":"containing", "value":"UNMAPPED", "format":bad
            })

        # Charts
        chs = wb.add_worksheet("Charts")
        chs.hide_gridlines(2)
        chs.set_column("A:A", 2)
        chs.set_column("B:Q", 13)
        chs.merge_range("B2:Q3", "MONTHLY CONNECTIONS – CHARTS", title)

        n = len(manager)
        if n:
            chart = wb.add_chart({"type":"bar"})
            chart.add_series({
                "name":"Connections",
                "categories":["DE_Manager_Report",1,0,n,0],
                "values":["DE_Manager_Report",1,1,n,1],
                "data_labels":{"value":True},
            })
            chart.set_title({"name":"Connections by DE / Manager"})
            chart.set_x_axis({"name":"Connections"})
            chart.set_y_axis({"name":"DE / Manager"})
            chart.set_legend({"none":True})
            chart.set_size({"width":1050,"height":620})
            chs.insert_chart("B5", chart)

        n2 = min(20, len(bbc))
        if n2:
            chart2 = wb.add_chart({"type":"column"})
            chart2.add_series({
                "name":"Connections",
                "categories":["BBC_Report",1,2,n2,2],
                "values":["BBC_Report",1,3,n2,3],
                "data_labels":{"value":True},
            })
            chart2.set_title({"name":"Top BBCs by Total Connections"})
            chart2.set_y_axis({"name":"Connections"})
            chart2.set_size({"width":1050,"height":600})
            chs.insert_chart("B38", chart2)

        n3 = min(20, len(olt))
        if n3:
            chart3 = wb.add_chart({"type":"bar"})
            chart3.add_series({
                "name":"Connections",
                "categories":["OLT_Report",1,0,n3,0],
                "values":["OLT_Report",1,4,n3,4],
                "data_labels":{"value":True},
            })
            chart3.set_title({"name":"Top OLT IPs by Connections"})
            chart3.set_x_axis({"name":"Connections"})
            chart3.set_y_axis({"name":"OLT IP"})
            chart3.set_legend({"none":True})
            chart3.set_size({"width":1050,"height":600})
            chs.insert_chart("B71", chart3)

    return out.getvalue()


def build_monthly_report(source_bytes, source_name, map_path, bbc_master_path):
    df = read_source(source_bytes)
    olt_map, bbc_info = load_mappings(map_path, bbc_master_path)
    df = add_mappings(df, olt_map, bbc_info)
    bbc, manager, franchisee, olt = aggregate(df)

    unmapped_olt = sorted(
        set(df.loc[df["BBC"] == "UNMAPPED", "OLT IP"].astype(str)) - {""}
    )

    xlsx = write_excel(
        df, bbc, manager, franchisee, olt,
        olt_map, bbc_info, source_name
    )

    stem = Path(source_name).stem
    output_name = f"{stem}_MONTHLY_DASHBOARD.xlsx"

    return {
        "xlsx_bytes": xlsx,
        "output_name": output_name,
        "bbc_report": bbc,
        "manager_report": manager,
        "franchisee_report": franchisee,
        "olt_report": olt,
        "unmapped_olt_list": unmapped_olt,
        "summary": {
            "total_connections": len(df),
            "olt_count": int(df["OLT IP"].nunique()),
            "franchisee_count": int(df["FR Service Code"].nunique()),
            "bbc_count": int(df["BBC"].nunique()),
            "manager_count": int(df["DE / Manager"].nunique()),
            "mapped_olts": int(df["OLT IP"].isin(olt_map["OLT IP"]).sum()),
            "unmapped_olts": len(unmapped_olt),
        }
    }
