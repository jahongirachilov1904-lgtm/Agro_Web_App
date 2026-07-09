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
    """
    Yangi yuklanadigan Excel formati:

    0-ustun: Stansiya nomi

    1-3 ustun: Kurtakning bo'rtishi yil/oy/kun
    4-6 ustun: 1-barg yozilishi yil/oy/kun
    7-9 ustun: Gullashi yil/oy/kun
    10-12 ustun: Pishib yetilish boshlanishi yil/oy/kun

    Bu faylda mavsum boshlanish sanasi bo'lmaydi.
    """

    raw = pd.read_excel(path_or_file, header=None)

    rows = []

    for i in range(1, len(raw)):
        r = raw.iloc[i]

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
            "Kurtakning bo'rtishi": p1,
            "1-barg yozilishi": p2,
            "Gullashi": p3,
            "Pishib yetilish boshlanishi": p4
        })

    return pd.DataFrame(rows)


def find_start_by_10_degree(df, year):
    """
    Har bir yil uchun avg_temp >= 10°C bo'lgan birinchi kunni topadi.
    10.0°C bo'lsa ham hisobga oladi.
    """

    data = df[
        (df["year"] == year) &
        (df[AIR_COL] >= BASE_TEMP)
    ].copy()

    if data.empty:
        return None

    data = data.sort_values("date")

    return data.iloc[0]["date"]


def calc_period(df, start_date, end_date):
    """
    start_date dan end_date gacha hisoblaydi.

    Kun - Havo:
        avg_temp >= 10°C bo'lgan kunlar soni

    Havo Aktiv ΣT:
        avg_temp >= 10°C bo'lgan kunlarning harorat yig'indisi

    Havo Effektiv ΣT:
        Σ(avg_temp - 10), faqat avg_temp >= 10°C kunlar uchun
    """

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
        "aktiv": above10[AIR_COL].sum(),
        "effektiv": (above10[AIR_COL] - BASE_TEMP).sum()
    }


def build_mavsum_boshidan(df, phase_df):
    rows = []

    for _, r in phase_df.iterrows():
        year = int(r["year"])

        start_date = find_start_by_10_degree(df, year)

        row = {
            "YILLAR": year,
            "Hisoblash boshlangan sana": date_text(start_date)
        }

        for phase in PHASES:
            phase_date = r[phase]
            air = calc_period(df, start_date, phase_date)

            row[f"{phase} | Sana"] = date_text(phase_date)
            row[f"{phase} | Kun - Havo"] = air["kun"]
            row[f"{phase} | Havo Aktiv ΣT"] = round(air["aktiv"], 1)
            row[f"{phase} | Havo Effektiv ΣT"] = round(air["effektiv"], 1)

        rows.append(row)

    return pd.DataFrame(rows)


def build_fazalar_orasi(df, phase_df):
    rows = []

    phase_pairs = [
        ("Hisoblash boshlangan sana", "Kurtakning bo'rtishi"),
        ("Kurtakning bo'rtishi", "1-barg yozilishi"),
        ("1-barg yozilishi", "Gullashi"),
        ("Gullashi", "Pishib yetilish boshlanishi")
    ]

    for _, r in phase_df.iterrows():
        year = int(r["year"])

        start_date = find_start_by_10_degree(df, year)

        dates = {
            "Hisoblash boshlangan sana": start_date,
            "Kurtakning bo'rtishi": r["Kurtakning bo'rtishi"],
            "1-barg yozilishi": r["1-barg yozilishi"],
            "Gullashi": r["Gullashi"],
            "Pishib yetilish boshlanishi": r["Pishib yetilish boshlanishi"]
        }

        row = {
            "YILLAR": year
        }

        for start_name, end_name in phase_pairs:
            s_date = dates[start_name]
            e_date = dates[end_name]

            air = calc_period(df, s_date, e_date)

            block = f"{start_name} -> {end_name}"

            row[f"{block} | Boshlanish sana"] = date_text(s_date)
            row[f"{block} | Faza o'zgargan sana"] = date_text(e_date)
            row[f"{block} | Kun - Havo"] = air["kun"]
            row[f"{block} | Havo Aktiv ΣT"] = round(air["aktiv"], 1)
            row[f"{block} | Havo Effektiv ΣT"] = round(air["effektiv"], 1)

        rows.append(row)

    return pd.DataFrame(rows)