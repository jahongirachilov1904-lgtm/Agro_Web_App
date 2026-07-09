import pandas as pd
from datetime import datetime

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side


AIR_COL = "avg_temp"
BASE_TEMP = 10

PHASES = [
    "Kurtakning bo'rtishi",
    "1-barg yozilishi",
    "Gullashi",
    "Pishib yetilish boshlanishi"
]


def make_date(y, m, d):
    try:
        if pd.isna(y) or pd.isna(m) or pd.isna(d):
            return None
        return datetime(int(y), int(m), int(d))
    except Exception:
        return None


def date_text(dt):
    if dt is None or pd.isna(dt):
        return ""
    return f"{dt.day}/{dt.month}"


def read_agro_data(path, sheet_name):
    xls = pd.ExcelFile(path)
    sheet_map = {str(s).strip().lower(): s for s in xls.sheet_names}
    key = str(sheet_name).strip().lower()

    if key not in sheet_map:
        raise ValueError(
            f"'{sheet_name}' nomli sheet bazada topilmadi! "
            f"Mavjud sheetlar: {', '.join(xls.sheet_names)}"
        )

    df = pd.read_excel(path, sheet_name=sheet_map[key])
    df.columns = [str(c).strip().lower() for c in df.columns]

    required = ["year", "month", "days", AIR_COL]
    for col in required:
        if col not in df.columns:
            raise ValueError(f"Excel bazada '{col}' ustuni topilmadi!")

    for col in required:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df = df.dropna(subset=["year", "month", "days", AIR_COL])

    df["year"] = df["year"].astype(int)
    df["month"] = df["month"].astype(int)
    df["days"] = df["days"].astype(int)

    df["date"] = pd.to_datetime(
        dict(year=df["year"], month=df["month"], day=df["days"]),
        errors="coerce"
    )

    df = df.dropna(subset=["date"])
    df = df.sort_values("date").reset_index(drop=True)

    return df


def read_phase_excel(path_or_file):
    raw = pd.read_excel(path_or_file, header=None)

    rows = []

    for i in range(1, len(raw)):
        r = raw.iloc[i]

        stansiya = r.iloc[0]

        p1 = make_date(r.iloc[1], r.iloc[2], r.iloc[3])
        p2 = make_date(r.iloc[4], r.iloc[5], r.iloc[6])
        p3 = make_date(r.iloc[7], r.iloc[8], r.iloc[9])
        p4 = make_date(r.iloc[10], r.iloc[11], r.iloc[12])

        dates = [p1, p2, p3, p4]
        years = [d.year for d in dates if d is not None]

        if not years:
            continue

        rows.append({
            "year": years[0],
            "stansiya": stansiya,
            "Kurtakning bo'rtishi": p1,
            "1-barg yozilishi": p2,
            "Gullashi": p3,
            "Pishib yetilish boshlanishi": p4
        })

    return pd.DataFrame(rows)


def find_start_by_10_degree(df, year):
    data = df[
        (df["year"] == year) &
        (df[AIR_COL] >= BASE_TEMP)
    ].copy()

    if data.empty:
        return None

    return data.sort_values("date").iloc[0]["date"]


def calc_period(df, start_date, end_date):
    if start_date is None or end_date is None or end_date < start_date:
        return {
            "kun": 0,
            "aktiv": 0,
            "effektiv": 0
        }

    data = df[
        (df["date"] >= start_date) &
        (df["date"] <= end_date)
    ].copy()

    data = data.dropna(subset=[AIR_COL])
    above10 = data[data[AIR_COL] >= BASE_TEMP]

    return {
        "kun": len(above10),
        "aktiv": round(above10[AIR_COL].sum(), 1),
        "effektiv": round((above10[AIR_COL] - BASE_TEMP).sum(), 1)
    }


def build_one_sheet(df, phase_df, station_name):
    rows = []

    for _, r in phase_df.iterrows():
        year = int(r["year"])

        start_date = find_start_by_10_degree(df, year)

        p1 = r["Kurtakning bo'rtishi"]
        p2 = r["1-barg yozilishi"]
        p3 = r["Gullashi"]
        p4 = r["Pishib yetilish boshlanishi"]

        stansiya = r.get("stansiya", station_name)
        if pd.isna(stansiya) or str(stansiya).strip() == "":
            stansiya = station_name

        # Fazalar orasidagi hisob
        h1 = calc_period(df, start_date, p1)
        h2 = calc_period(df, p1, p2)
        h3 = calc_period(df, p2, p3)
        h4 = calc_period(df, p3, p4)

        # Mavsum boshidan hisob
        c1 = calc_period(df, start_date, p1)
        c2 = calc_period(df, start_date, p2)
        c3 = calc_period(df, start_date, p3)
        c4 = calc_period(df, start_date, p4)

        row = {
            "YILLAR": year,
            "Stansiya": stansiya,
            "Hisoblash boshlangan sana": date_text(start_date),

            "Kurtakning bo'rtishi | Sana": date_text(p1),
            "Kurtakning bo'rtishi | Kun - Havo": h1["kun"],
            "Kurtakning bo'rtishi | Havo Aktiv ΣT": h1["aktiv"],
            "Kurtakning bo'rtishi | Havo Effektiv ΣT": h1["effektiv"],
            "Kurtakning bo'rtishi | Mavsum boshidan Aktiv ΣT": c1["aktiv"],
            "Kurtakning bo'rtishi | Mavsum boshidan Effektiv ΣT": c1["effektiv"],

            "1-barg yozilishi | Sana": date_text(p2),
            "1-barg yozilishi | Kun - Havo": h2["kun"],
            "1-barg yozilishi | Havo Aktiv ΣT": h2["aktiv"],
            "1-barg yozilishi | Havo Effektiv ΣT": h2["effektiv"],
            "1-barg yozilishi | Mavsum boshidan Aktiv ΣT": c2["aktiv"],
            "1-barg yozilishi | Mavsum boshidan Effektiv ΣT": c2["effektiv"],

            "Gullashi | Sana": date_text(p3),
            "Gullashi | Kun - Havo": h3["kun"],
            "Gullashi | Havo Aktiv ΣT": h3["aktiv"],
            "Gullashi | Havo Effektiv ΣT": h3["effektiv"],
            "Gullashi | Mavsum boshidan Aktiv ΣT": c3["aktiv"],
            "Gullashi | Mavsum boshidan Effektiv ΣT": c3["effektiv"],

            "Pishib yetilish boshlanishi | Sana": date_text(p4),
            "Pishib yetilish boshlanishi | Kun - Havo": h4["kun"],
            "Pishib yetilish boshlanishi | Havo Aktiv ΣT": h4["aktiv"],
            "Pishib yetilish boshlanishi | Havo Effektiv ΣT": h4["effektiv"],
            "Pishib yetilish boshlanishi | Mavsum boshidan Aktiv ΣT": c4["aktiv"],
            "Pishib yetilish boshlanishi | Mavsum boshidan Effektiv ΣT": c4["effektiv"],
        }

        rows.append(row)

    return pd.DataFrame(rows)


def format_excel(path):
    wb = load_workbook(path)

    green = PatternFill("solid", fgColor="C6E0B4")
    blue = PatternFill("solid", fgColor="D9EAF7")
    yellow = PatternFill("solid", fgColor="FFF2CC")
    thin = Side(style="thin", color="999999")

    for ws in wb.worksheets:
        ws.freeze_panes = "A2"

        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)

            value = str(cell.value)

            if "Sana" in value or "sana" in value:
                cell.fill = blue
            elif "Mavsum boshidan" in value:
                cell.fill = yellow
            else:
                cell.fill = green

        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)

        for col in ws.columns:
            max_len = 0
            col_letter = col[0].column_letter

            for cell in col:
                if cell.value is not None:
                    max_len = max(max_len, len(str(cell.value)))

            ws.column_dimensions[col_letter].width = min(max_len + 3, 35)

    wb.save(path)


def create_result_excel(data_excel_path, phase_file, station_name, output_path):
    df = read_agro_data(data_excel_path, station_name)
    phase_df = read_phase_excel(phase_file)

    result_df = build_one_sheet(df, phase_df, station_name)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        result_df.to_excel(
            writer,
            sheet_name="Natija",
            index=False
        )

    format_excel(output_path)

    return output_path