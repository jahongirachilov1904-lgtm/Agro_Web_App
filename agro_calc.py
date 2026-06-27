import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side


AIR_COL = "avg_temp"

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
    if dt is None:
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

    real_sheet_name = sheet_map[key]

    df = pd.read_excel(path, sheet_name=real_sheet_name)
    df.columns = [str(c).strip().lower() for c in df.columns]

    required = ["year", "month", "days", AIR_COL]

    for col in required:
        if col not in df.columns:
            raise ValueError(f"Excel bazada '{col}' ustuni topilmadi!")

    for col in required:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df = df.dropna(subset=["year", "month", "days"])
    df["year"] = df["year"].astype(int)
    df["month"] = df["month"].astype(int)
    df["days"] = df["days"].astype(int)

    df["date"] = pd.to_datetime(
        dict(
            year=df["year"],
            month=df["month"],
            day=df["days"]
        ),
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

        veg = make_date(r.iloc[1], r.iloc[2], r.iloc[3])
        p1 = make_date(r.iloc[4], r.iloc[5], r.iloc[6])
        p2 = make_date(r.iloc[7], r.iloc[8], r.iloc[9])
        p3 = make_date(r.iloc[10], r.iloc[11], r.iloc[12])
        p4 = make_date(r.iloc[13], r.iloc[14], r.iloc[15])

        dates = [veg, p1, p2, p3, p4]
        years = [d.year for d in dates if d is not None]

        if not years:
            continue

        rows.append({
            "year": years[0],
            "veg_excel": veg,
            "Kurtakning bo'rtishi": p1,
            "1-barg yozilishi": p2,
            "Gullashi": p3,
            "Pishib yetilish boshlanishi": p4
        })

    return pd.DataFrame(rows)


def find_veg_start_by_air(df, year):
    data = df[(df["year"] == year) & (df[AIR_COL] > 5)]

    if data.empty:
        return None

    return data.iloc[0]["date"]


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

    above10 = data[data[AIR_COL] > 10]

    return {
        "kun": len(above10),
        "aktiv": above10[AIR_COL].sum(),
        "effektiv": (above10[AIR_COL] - 10).sum()
    }


def build_mavsum_boshidan(df, phase_df):
    rows = []

    for _, r in phase_df.iterrows():
        year = int(r["year"])
        veg_start = find_veg_start_by_air(df, year)

        if veg_start is None:
            veg_start = r["veg_excel"]

        row = {
            "YILLAR": year,
            "Uzum mavsumi boshlanish sanasi": date_text(veg_start)
        }

        for phase in PHASES:
            phase_date = r[phase]

            air = calc_period(df, veg_start, phase_date)

            row[f"{phase} | Sana"] = date_text(phase_date)
            row[f"{phase} | Kun - Havo"] = air["kun"]
            row[f"{phase} | Havo Aktiv ΣT"] = round(air["aktiv"], 1)
            row[f"{phase} | Havo Effektiv ΣT"] = round(air["effektiv"], 1)

        rows.append(row)

    return pd.DataFrame(rows)


def build_fazalar_orasi(df, phase_df):
    rows = []

    phase_pairs = [
        ("Uzum mavsumi boshlanish sanasi", "Kurtakning bo'rtishi"),
        ("Kurtakning bo'rtishi", "1-barg yozilishi"),
        ("1-barg yozilishi", "Gullashi"),
        ("Gullashi", "Pishib yetilish boshlanishi")
    ]

    for _, r in phase_df.iterrows():
        year = int(r["year"])
        veg_start = find_veg_start_by_air(df, year)

        if veg_start is None:
            veg_start = r["veg_excel"]

        dates = {
            "Uzum mavsumi boshlanish sanasi": veg_start,
            "Kurtakning bo'rtishi": r["Kurtakning bo'rtishi"],
            "1-barg yozilishi": r["1-barg yozilishi"],
            "Gullashi": r["Gullashi"],
            "Pishib yetilish boshlanishi": r["Pishib yetilish boshlanishi"]
        }

        row = {
            "YILLAR": year
        }

        for start_name, end_name in phase_pairs:
            start_date = dates[start_name]
            end_date = dates[end_name]

            air = calc_period(df, start_date, end_date)

            block = f"{start_name} -> {end_name}"

            row[f"{block} | Boshlanish sana"] = date_text(start_date)
            row[f"{block} | Faza o'zgargan sana"] = date_text(end_date)
            row[f"{block} | Kun - Havo"] = air["kun"]
            row[f"{block} | Havo Aktiv ΣT"] = round(air["aktiv"], 1)
            row[f"{block} | Havo Effektiv ΣT"] = round(air["effektiv"], 1)

        rows.append(row)

    return pd.DataFrame(rows)


def format_excel(path):
    wb = load_workbook(path)

    green = PatternFill("solid", fgColor="C6E0B4")
    yellow = PatternFill("solid", fgColor="FFF2CC")
    blue = PatternFill("solid", fgColor="D9EAF7")
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
            cell.border = Border(
                left=thin,
                right=thin,
                top=thin,
                bottom=thin
            )

            if "Faza o'zgargan sana" in str(cell.value):
                cell.fill = yellow
            elif "Sana" in str(cell.value) or "sana" in str(cell.value):
                cell.fill = blue
            else:
                cell.fill = green

        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center"
                )
                cell.border = Border(
                    left=thin,
                    right=thin,
                    top=thin,
                    bottom=thin
                )

        for col in ws.columns:
            max_len = 0
            col_letter = col[0].column_letter

            for cell in col:
                if cell.value is not None:
                    max_len = max(max_len, len(str(cell.value)))

            ws.column_dimensions[col_letter].width = min(max_len + 3, 28)

    wb.save(path)


def create_result_excel(data_excel_path, phase_file, station_name, output_path):
    df = read_agro_data(data_excel_path, station_name)
    phase_df = read_phase_excel(phase_file)

    mavsum_df = build_mavsum_boshidan(df, phase_df)
    fazalar_df = build_fazalar_orasi(df, phase_df)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        mavsum_df.to_excel(
            writer,
            sheet_name="Mavsum_boshidan",
            index=False
        )

        fazalar_df.to_excel(
            writer,
            sheet_name="Fazalar_orasi",
            index=False
        )

    format_excel(output_path)

    return output_path