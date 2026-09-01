# Agro Web App

Agrometeorologik meva fazalari orasidagi aktiv va effektiv harorat yig‘indisi,
jami yog‘in hamda yog‘inli kunlar sonini hisoblaydigan Streamlit ilovasi.

## Talablar

- Python 3.12
- Loyiha ildizidagi `Agro_T2M_kunlik_1991_2025.xlsx` ma’lumot fayli

## Lokal ishga tushirish

PowerShell’da loyiha papkasiga o‘ting:

```powershell
cd C:\Users\WECAN\PyCharmProjects\Agro_Web_App
```

Virtual muhit yarating va faollashtiring:

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\Activate.ps1
```

Kutubxonalarni o‘rnating:

```powershell
python -m pip install -r requirements.txt
```

Ilovani ishga tushiring:

```powershell
streamlit run app.py
```

Brauzerda odatda `http://localhost:8501` manzili ochiladi.

## Streamlit Community Cloud’ga joylash

1. Repository fayllarini GitHub’ga yuboring. Katta meteorologik Excel fayli ham
   repository ichida qolishi kerak, chunki dastur uni lokal fayl sifatida o‘qiydi.
2. [Streamlit Community Cloud](https://share.streamlit.io/) saytiga GitHub orqali kiring.
3. **Create app** tugmasini bosing va GitHub repository’ni tanlang.
4. Branch sifatida `main`ni tanlang.
5. **Main file path** maydoniga `app.py` yozing.
6. Deploy jarayonini boshlang va build logida dependency yoki fayl xatosi
   yo‘qligini tekshiring.

`requirements.txt` va `.streamlit/config.toml` repository ildizida joylashgan.
Community Cloud Linux muhitida ishlaydi; dastur ma’lumot faylini `app.py`
joylashgan papkaga nisbatan topadi va Windows absolute pathga bog‘liq emas.

## Secrets

Hozirgi ilova database yoki tashqi maxfiy API credential ishlatmaydi, shuning
uchun Streamlit secrets sozlamasi talab qilinmaydi. Kelajakda maxfiy qiymatlar
kerak bo‘lsa, ularni kod yoki GitHub repository’ga yozmang; Community Cloud’dagi
**App settings → Secrets** bo‘limidan kiriting. Lokal `.streamlit/secrets.toml`
fayli `.gitignore`ga kiritilgan.
