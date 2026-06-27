import os
import tempfile

import streamlit as st

from agro_calc import create_result_excel


st.set_page_config(
    page_title="Agroklimatik Hisoblash",
    page_icon="🌱",
    layout="centered"
)

st.title("🌱 Agroklimatik hisoblash tizimi")

st.write(
    "Fenologik Excel faylni yuklang. "
    "Dastur fayl nomidan stansiyani avtomatik aniqlaydi."
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_EXCEL_PATH = os.path.join(BASE_DIR, "Agro_mevalar.xlsx")

if not os.path.exists(DATA_EXCEL_PATH):
    st.error("Agro_mevalar.xlsx baza fayli loyiha papkasida topilmadi!")
    st.stop()

uploaded_file = st.file_uploader(
    "Fenologik Excel faylni yuklang",
    type=["xlsx"]
)

if uploaded_file is not None:
    file_name = uploaded_file.name
    base_name = os.path.splitext(file_name)[0]

    station_name = base_name.split("_")[0]

    st.success(f"Yuklangan fayl: {file_name}")
    st.info(f"Aniqlangan stansiya: {station_name}")

    if st.button("Hisoblashni boshlash"):
        try:
            with st.spinner("Hisoblanmoqda..."):
                temp_dir = tempfile.mkdtemp()

                phase_path = os.path.join(temp_dir, file_name)

                with open(phase_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())

                output_name = f"{base_name}_natija.xlsx"
                output_path = os.path.join(temp_dir, output_name)

                create_result_excel(
                    data_excel_path=DATA_EXCEL_PATH,
                    phase_file=phase_path,
                    station_name=station_name,
                    output_path=output_path
                )

            st.success("Hisoblash yakunlandi!")

            with open(output_path, "rb") as f:
                st.download_button(
                    label="📥 Natijani yuklab olish",
                    data=f,
                    file_name=output_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error("Hisoblashda xatolik yuz berdi.")
            st.exception(e)