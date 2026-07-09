import os
import re
import tempfile

import streamlit as st

from agro_calc import create_result_excel


st.set_page_config(
    page_title="Agro Meva fazalar",
    page_icon="🌱",
    layout="centered"
)


st.title("🌱 Agrometeorologik mevalarning fazalarini hisoblash")

st.write(
    "Fenologik Excel faylni yuklang. "
    "Harorat chegarasini 1°C dan 15°C gacha tanlang. "
    "Dastur fayl nomidan stansiyani avtomatik aniqlaydi."
)


BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DATA_EXCEL_PATH = os.path.join(
    BASE_DIR,
    "Agro_T2M_kunlik_1991_2025.xlsx"
)


if not os.path.exists(DATA_EXCEL_PATH):
    st.error(
        "Agro_T2M_kunlik_1991_2025.xlsx baza fayli loyiha papkasida topilmadi!"
    )
    st.stop()


def clean_station_name(file_name):
    base_name = os.path.splitext(file_name)[0]

    # Masalan: Shahrisabz(1) -> Shahrisabz
    base_name = re.sub(r"\(\d+\)", "", base_name)

    # Masalan: Shahrisabz_2020 -> Shahrisabz
    station_name = base_name.split("_")[0]

    return station_name.strip()


base_temp = st.number_input(
    "Hisoblash boshlanadigan harorat chegarasini tanlang (°C)",
    min_value=1,
    max_value=15,
    value=10,
    step=1
)


uploaded_file = st.file_uploader(
    "Fenologik Excel faylni yuklang",
    type=["xlsx"]
)


if uploaded_file is not None:
    file_name = uploaded_file.name
    base_name = os.path.splitext(file_name)[0]
    station_name = clean_station_name(file_name)

    st.success(f"Yuklangan fayl: {file_name}")
    st.info(f"Aniqlangan stansiya: {station_name}")
    st.info(f"Tanlangan harorat chegarasi: {base_temp}°C")

    if st.button("Hisoblashni boshlash"):
        try:
            with st.spinner("Hisoblanmoqda..."):
                temp_dir = tempfile.mkdtemp()

                phase_path = os.path.join(temp_dir, file_name)

                with open(phase_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())

                output_name = f"{base_name}_natija_{base_temp}C.xlsx"
                output_path = os.path.join(temp_dir, output_name)

                result_path = create_result_excel(
                    data_excel_path=DATA_EXCEL_PATH,
                    phase_file=phase_path,
                    station_name=station_name,
                    output_path=output_path,
                    base_temp=base_temp
                )

            st.success("Hisoblash yakunlandi!")

            with open(result_path, "rb") as f:
                st.download_button(
                    label="📥 Natijani yuklab olish",
                    data=f,
                    file_name=output_name,
                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "spreadsheetml.sheet"
                    )
                )

        except Exception as e:
            st.error("Hisoblashda xatolik yuz berdi.")
            st.exception(e)