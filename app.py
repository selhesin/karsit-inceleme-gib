import io
import os
import re
import zipfile
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

st.set_page_config(
    page_title="191 KDV Karşıt İnceleme & GİB Dönüştürücü",
    page_icon="🧾",
    layout="wide",
)

GIB_COLUMNS = [
    "Faturanın Tarihi",
    "Faturanın Serisi",
    "Faturanın Numarası",
    "Faturanın Tutarı (TL)",
    "K.D.V(TL)",
    "Defter Kayıt Tarihi",
    "Yevmiye Numarası",
    "Ödeme Şekli ve Ödemeye İlişkin Bilgiler",
    "Açıklama",
    "Hatalı Satır Açıklama",
]


def tl_str(val: float) -> str:
    return f"{val:,.2f} ₺".replace(",", "X").replace(".", ",").replace("X", ".")


def temiz_dosya_adi(metin: str) -> str:
    temiz = re.sub(r'[\\/*?:"<>|]+', "_", str(metin)).strip()
    return temiz[:45] if temiz else "Firma"


def kdv_listesi_oku(file_obj) -> pd.DataFrame:
    xl = pd.ExcelFile(file_obj)
    hedef_sayfa = xl.sheet_names[0]
    for sh in xl.sheet_names:
        if "indirilecek" in sh.lower() or "kdv" in sh.lower():
            hedef_sayfa = sh
            break

    raw = pd.read_excel(xl, sheet_name=hedef_sayfa, header=None)

    baslik_idx = 2
    for idx, row in raw.iterrows():
        satir_metni = " ".join(str(x) for x in row.values if pd.notna(x)).lower()
        if "tarihi" in satir_metni and ("kdv" in satir_metni or "vergi kimlik" in satir_metni):
            baslik_idx = idx
            break

    df = raw.iloc[baslik_idx + 1 :].copy().reset_index(drop=True)
    if df.shape[1] >= 10 and df.iloc[:, 0].isna().all():
        df = df.iloc[:, 1:10].copy()
    else:
        df = df.iloc[:, :9].copy()

    df.columns = [
        "Tarih",
        "Seri",
        "SiraNo",
        "Unvan",
        "VKN",
        "Cins",
        "Miktar",
        "Matrah",
        "KDV",
    ]

    df = df[df["Tarih"].notna() & df["KDV"].notna()].copy()
    df["Tarih_dt"] = pd.to_datetime(df["Tarih"], dayfirst=True, errors="coerce")
    df = df[df["Tarih_dt"].notna()].copy()
    df["Ay"] = df["Tarih_dt"].dt.to_period("M").astype(str)

    df["VKN"] = df["VKN"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    df["Unvan"] = df["Unvan"].astype(str).str.strip()
    df["Seri"] = df["Seri"].fillna("").astype(str).str.strip()
    df["SiraNo"] = df["SiraNo"].fillna("").astype(str).str.strip()
    df["Cins"] = df["Cins"].fillna("").astype(str).str.strip()
    df["Matrah"] = pd.to_numeric(df["Matrah"], errors="coerce").fillna(0.0)
    df["KDV"] = pd.to_numeric(df["KDV"], errors="coerce").fillna(0.0)
    df["KDV_Dahil_Tutar"] = df["Matrah"] + df["KDV"]
    return df


def gib_excel_olustur(df_faturalar: pd.DataFrame, tutar_secimi: str, aciklama_yaz: bool) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Veriler"

    hdr_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    hdr_fill = PatternFill("solid", fgColor="1E293B")
    body_font = Font(name="Calibri", size=11, color="000000")
    thin = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1"),
    )

    for col_idx, col_name in enumerate(GIB_COLUMNS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = hdr_font
        cell.fill = hdr_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin
    ws.row_dimensions[1].height = 26

    for r_idx, (_, r) in enumerate(df_faturalar.iterrows(), start=2):
        fatura_tutari = r["KDV_Dahil_Tutar"] if tutar_secimi.startswith("KDV Dahil") else r["Matrah"]
        aciklama_metni = r["Cins"] if aciklama_yaz else ""

        satir = [
            r["Tarih_dt"].to_pydatetime(),
            r["Seri"],
            r["SiraNo"],
            round(float(fatura_tutari), 2),
            round(float(r["KDV"]), 2),
            "",
            "",
            "",
            aciklama_metni,
            "",
        ]

        for c_idx, val in enumerate(satir, start=1):
            c = ws.cell(row=r_idx, column=c_idx, value=val)
            c.font = body_font
            c.border = thin
            if c_idx == 1:
                c.number_format = r"dd\.mm\.yyyy"
                c.alignment = Alignment(horizontal="center")
            elif c_idx in (4, 5):
                c.number_format = "#,##0.00"
                c.alignment = Alignment(horizontal="right")
            else:
                c.number_format = "@"

    genislikler = [16, 16, 22, 22, 18, 18, 18, 34, 30, 22]
    for i, w in enumerate(genislikler, start=1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


st.title("🧾 191 İndirilecek KDV Karşıt İnceleme & GİB Excel Dönüştürücü")
st.caption("Ay bazında KDV %80 kümülatif pay ve İlk N alış parametresiyle firmaları ayrıştırır, faturalarını ayırır ve GİB 'Veriler' şablonuna çevirir.")

with st.sidebar:
    st.header("⚙️ Karşıt İnceleme Filtreleri")
    kdv_hedef_pct = st.slider("Kümülatif KDV Eşiği (%)", min_value=50, max_value=100, value=80, step=5)
    ilk_n_firma = st.number_input("Maksimum Firma Sayısı (İlk N)", min_value=1, max_value=50, value=10, step=1)
    kural_tipi = st.radio(
        "Seçim Kuralı",
        [
            "%80 KDV'ye Ulaşanlar (En Fazla İlk N)",
            "%80 KDV ve İlk N Birleşimi",
        ],
        index=0,
    )
    st.divider()
    st.header("📄 GİB Şablon Ayarları")
    tutar_secimi = st.radio(
        "Faturanın Tutarı (TL)",
        ["KDV Dahil (Matrah + KDV)", "KDV Hariç (Matrah)"],
        index=0,
    )
    aciklama_yaz = st.checkbox("Mal/Hizmet Cinsini 'Açıklama' sütununa aktar", value=True)

yuklenen_excel = st.file_uploader(
    "191 İndirilecek KDV Listesi Excel Dosyasını Yükleyin (.xlsx)",
    type=["xlsx", "xls"],
)

if not yuklenen_excel:
    st.info("Devam etmek için İndirilecek KDV Listesi Excel dosyanızı yükleyin.")
    st.stop()

df_tum = kdv_listesi_oku(yuklenen_excel)
ay_listesi = sorted(df_tum["Ay"].unique())

secili_ay = st.selectbox("📅 Dönem (Ay) Seçin", ["TÜM DÖNEM"] + ay_listesi, index=1 if ay_listesi else 0)
df_ay = df_tum if secili_ay == "TÜM DÖNEM" else df_tum[df_tum["Ay"] == secili_ay].copy()

toplam_kdv = df_ay["KDV"].sum()
firma_tablosu = (
    df_ay.groupby(["VKN", "Unvan"], as_index=False)
    .agg(
        Fatura_Adedi=("SiraNo", "count"),
        KDV_Haric=("Matrah", "sum"),
        KDV_Tutari=("KDV", "sum"),
        KDV_Dahil=("KDV_Dahil_Tutar", "sum"),
    )
    .sort_values("KDV_Tutari", ascending=False)
    .reset_index(drop=True)
)

firma_tablosu["KDV_Payi_%"] = (firma_tablosu["KDV_Tutari"] / toplam_kdv * 100) if toplam_kdv > 0 else 0.0
firma_tablosu["Kumulatif_%"] = firma_tablosu["KDV_Payi_%"].cumsum()

esik_idx = firma_tablosu[firma_tablosu["Kumulatif_%"] >= kdv_hedef_pct].index.min()
n_kdv_esik = int(esik_idx + 1) if pd.notna(esik_idx) else len(firma_tablosu)

if kural_tipi.startswith("%80 KDV'ye"):
    sec_adet = min(n_kdv_esik, int(ilk_n_firma))
else:
    sec_adet = min(len(firma_tablosu), max(n_kdv_esik, int(ilk_n_firma)))

firma_tablosu["Secili"] = firma_tablosu.index < sec_adet

c1, c2, c3, c4 = st.columns(4)
c1.metric("Dönem Fatura Sayısı", f"{len(df_ay):,}")
c2.metric("Dönem Toplam KDV", tl_str(toplam_kdv))
c3.metric("Ayrıştırılan Firma", f"{sec_adet} / {len(firma_tablosu)}")
c4.metric("Ayrıştırılan KDV Payı", f"%{firma_tablosu.loc[firma_tablosu['Secili'], 'KDV_Payi_%'].sum():.1f}")

st.subheader("1️⃣ Karşıt İnceleme Kapsamındaki Firmalar")
onay_tablosu = st.data_editor(
    firma_tablosu,
    use_container_width=True,
    hide_index=True,
    column_config={
        "Secili": st.column_config.CheckboxColumn("Seç", default=True),
        "KDV_Haric": st.column_config.NumberColumn("KDV Hariç (₺)", format="%.2f"),
        "KDV_Tutari": st.column_config.NumberColumn("KDV (₺)", format="%.2f"),
        "KDV_Dahil": st.column_config.NumberColumn("KDV Dahil (₺)", format="%.2f"),
        "KDV_Payi_%": st.column_config.NumberColumn("KDV Payı (%)", format="%.2f%%"),
        "Kumulatif_%": st.column_config.NumberColumn("Kümülatif (%)", format="%.2f%%"),
    },
    disabled=["VKN", "Unvan", "Fatura_Adedi", "KDV_Haric", "KDV_Tutari", "KDV_Dahil", "KDV_Payi_%", "Kumulatif_%"],
)

secili_firmalar = onay_tablosu[onay_tablosu["Secili"]].copy()
df_secili_faturalar = df_ay[df_ay["VKN"].isin(secili_firmalar["VKN"])].sort_values(["Unvan", "Tarih_dt"]).copy()

st.subheader("2️⃣ Ayrıştırılan Faturalar & GİB 'Veriler' Çıktısı")
if df_secili_faturalar.empty:
    st.warning("En az bir firma seçili olmalıdır.")
    st.stop()

st.dataframe(
    pd.DataFrame(
        {
            "Satıcı Ünvanı": df_secili_faturalar["Unvan"],
            "Faturanın Tarihi": df_secili_faturalar["Tarih_dt"].dt.strftime("%d.%m.%Y"),
            "Faturanın Serisi": df_secili_faturalar["Seri"],
            "Faturanın Numarası": df_secili_faturalar["SiraNo"],
            "Faturanın Tutarı (TL)": (
                df_secili_faturalar["KDV_Dahil_Tutar"]
                if tutar_secimi.startswith("KDV Dahil")
                else df_secili_faturalar["Matrah"]
            ).round(2),
            "K.D.V(TL)": df_secili_faturalar["KDV"].round(2),
            "Açıklama": df_secili_faturalar["Cins"] if aciklama_yaz else "",
        }
    ),
    use_container_width=True,
    hide_index=True,
)

zip_buf = io.BytesIO()
with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
    for sira, (_, f_row) in enumerate(secili_firmalar.iterrows(), start=1):
        f_df = df_secili_faturalar[df_secili_faturalar["VKN"] == f_row["VKN"]]
        xl_data = gib_excel_olustur(f_df, tutar_secimi, aciklama_yaz)
        ad = f"1.{sira}-{temiz_dosya_adi(f_row['Unvan'])}_{secili_ay}.xlsx"
        zf.writestr(ad, xl_data)
zip_buf.seek(0)

col_b1, col_b2 = st.columns(2)
with col_b1:
    st.download_button(
        label=f"📦 Firma Bazlı GİB Excel'leri İndir (.ZIP — {len(secili_firmalar)} Firma)",
        data=zip_buf.getvalue(),
        file_name=f"GIB_Karsit_Inceleme_{secili_ay}.zip",
        mime="application/zip",
        type="primary",
        use_container_width=True,
    )

with col_b2:
    st.download_button(
        label="📥 Seçili Tüm Faturaları Tek GİB Excel İndir (.xlsx)",
        data=gib_excel_olustur(df_secili_faturalar, tutar_secimi, aciklama_yaz),
        file_name=f"GIB_Veriler_Toplu_{secili_ay}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
