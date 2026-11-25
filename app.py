import io
import re
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

APP_SIFRE = "1903"

st.set_page_config(
    page_title="191 KDV Karşıt İnceleme & GİB Dönüştürücü",
    page_icon="🧾",
    layout="wide",
)

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0


def sifre_ekrani():
    st.markdown("## 🔒 Karşıt İnceleme Portalı — Giriş")
    st.caption("Uygulamaya erişmek için şifreyi girin.")

    col_form, _ = st.columns([1.2, 2.8])
    with col_form:
        with st.form("login_form", clear_on_submit=False):
            girilen_sifre = st.text_input("Şifre", type="password", placeholder="Şifreyi girin...")
            giris_btn = st.form_submit_button("Giriş Yap", type="primary", use_container_width=True)

        if giris_btn:
            if girilen_sifre == APP_SIFRE:
                st.session_state["authenticated"] = True
                st.rerun()
            else:
                st.error("Hatalı şifre. Lütfen tekrar deneyin.")


if not st.session_state["authenticated"]:
    sifre_ekrani()
    st.stop()

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


def guvenli_sayfa_adi(ad: str, kullanilanlar: set) -> str:
    temiz = re.sub(r'[\\/?*\[\]:]+', " ", str(ad)).strip()
    temiz = re.sub(r"\s+", " ", temiz)[:31].strip() or "Firma"
    aday = temiz
    sayac = 2
    while aday in kullanilanlar:
        ek = f" ({sayac})"
        aday = f"{temiz[: 31 - len(ek)]}{ek}"
        sayac += 1
    kullanilanlar.add(aday)
    return aday


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


def sayfa_doldur_gib(ws, df_faturalar: pd.DataFrame, tutar_secimi: str, aciklama_yaz: bool):
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


def firma_sayfali_excel_olustur(
    firma_kayitlari: list[tuple[str, pd.DataFrame]],
    tutar_secimi: str,
    aciklama_yaz: bool,
) -> bytes:
    wb = Workbook()
    kullanilan_isimler = set()
    ilk_sayfa = True

    for etiket, f_df in firma_kayitlari:
        if f_df.empty:
            continue
        sayfa_adi = guvenli_sayfa_adi(etiket, kullanilan_isimler)
        if ilk_sayfa:
            ws = wb.active
            ws.title = sayfa_adi
            ilk_sayfa = False
        else:
            ws = wb.create_sheet(title=sayfa_adi)
        sayfa_doldur_gib(ws, f_df, tutar_secimi, aciklama_yaz)

    if ilk_sayfa:
        ws = wb.active
        ws.title = "Bos"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


col_baslik, col_yenile, col_sil, col_cikis = st.columns([5.2, 1, 1, 1.2])
with col_baslik:
    st.title("🧾 191 İndirilecek KDV Karşıt İnceleme & GİB Dönüştürücü")
with col_yenile:
    st.write("")
    if st.button("🔄 Yenile", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
with col_sil:
    st.write("")
    if st.button("🗑️ Sil", use_container_width=True):
        st.session_state["uploader_key"] += 1
        st.cache_data.clear()
        st.rerun()
with col_cikis:
    st.write("")
    if st.button("🔐 Kilitle", use_container_width=True):
        st.session_state["authenticated"] = False
        st.rerun()

st.caption("Seçilen dönemdeki tüm firmaları KDV tutarları ve indirim/kapsama yüzdeleriyle listeler; tek Excel'de firma bazlı sayfalar üretir.")

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
    key=f"kdv_uploader_{st.session_state['uploader_key']}",
)

if not yuklenen_excel:
    st.info("Devam etmek için İndirilecek KDV Listesi Excel dosyanızı yükleyin.")
    st.stop()

df_tum = kdv_listesi_oku(yuklenen_excel)
ay_listesi = sorted(df_tum["Ay"].unique())

secili_ay = st.selectbox(
    "📅 Dönem (Ay) Seçin",
    ["TÜM AYLAR (AY AY SAYFA)", "TÜM DÖNEM (TEK TOPLAM)"] + ay_listesi,
    index=0,
)


def firma_ozet_hesapla(df_parca: pd.DataFrame) -> tuple[pd.DataFrame, float]:
    tot_kdv = float(df_parca["KDV"].sum())
    ozet = (
        df_parca.groupby(["VKN", "Unvan"], as_index=False)
        .agg(
            Fatura_Adedi=("SiraNo", "count"),
            KDV_Haric=("Matrah", "sum"),
            KDV_Tutari=("KDV", "sum"),
            KDV_Dahil=("KDV_Dahil_Tutar", "sum"),
        )
        .sort_values("KDV_Tutari", ascending=False)
        .reset_index(drop=True)
    )
    ozet["KDV_Payi_%"] = (ozet["KDV_Tutari"] / tot_kdv * 100) if tot_kdv > 0 else 0.0
    ozet["Kumulatif_%"] = ozet["KDV_Payi_%"].cumsum()

    esik_i = ozet[ozet["Kumulatif_%"] >= kdv_hedef_pct].index.min()
    n_esik = int(esik_i + 1) if pd.notna(esik_i) else len(ozet)

    if kural_tipi.startswith("%80 KDV'ye"):
        sec_n = min(n_esik, int(ilk_n_firma))
    else:
        sec_n = min(len(ozet), max(n_esik, int(ilk_n_firma)))

    ozet["Secili"] = ozet.index < sec_n
    return ozet, tot_kdv


def tum_firmalari_liste_goster(ozet_df: pd.DataFrame, toplam_kdv: float, secili_kdv: float):
    indirim_orani = (secili_kdv / toplam_kdv * 100) if toplam_kdv > 0 else 0.0
    st.markdown(
        f"**İndirim / KDV Kapsama Özeti:** Seçilen firmaların toplam KDV'si **{tl_str(secili_kdv)}** olup, "
        f"toplam **{tl_str(toplam_kdv)}** tutarındaki 191 İndirilecek KDV'nin **%{indirim_orani:.2f}**'sini kapsamaktadır."
    )
    with st.expander(f"📋 Tüm Firmalar Listesi ({len(ozet_df)} Firma — Toplam KDV ve Payları)", expanded=True):
        for idx, r in ozet_df.iterrows():
            rozet = "✅ Seçili" if r["Secili"] else "⬜ Hariç"
            st.markdown(
                f"{idx + 1}. **{r['Unvan']}** (`VKN: {r['VKN']}`) — "
                f"Toplam KDV: **{tl_str(r['KDV_Tutari'])}** "
                f"(Pay: **%{r['KDV_Payi_%']:.2f}** | Kümülatif: **%{r['Kumulatif_%']:.2f}** | {r['Fatura_Adedi']} fatura) — {rozet}"
            )


if secili_ay == "TÜM AYLAR (AY AY SAYFA)":
    tum_ozet_list = []
    sayfa_paketleri = []
    for ay_kodu in ay_listesi:
        df_m = df_tum[df_tum["Ay"] == ay_kodu].copy()
        oz_m, _ = firma_ozet_hesapla(df_m)
        oz_m.insert(0, "Ay", ay_kodu)
        tum_ozet_list.append(oz_m)

        for _, f_r in oz_m[oz_m["Secili"]].iterrows():
            f_fat = df_m[df_m["VKN"] == f_r["VKN"]].sort_values("Tarih_dt")
            sayfa_paketleri.append((f"{ay_kodu} {f_r['Unvan']}", f_fat))

    birlesik_ozet = pd.concat(tum_ozet_list, ignore_index=True) if tum_ozet_list else pd.DataFrame()
    toplam_kdv_tum = float(df_tum["KDV"].sum())
    secili_kdv_tum = float(birlesik_ozet.loc[birlesik_ozet["Secili"], "KDV_Tutari"].sum()) if not birlesik_ozet.empty else 0.0
    indirim_pct_tum = (secili_kdv_tum / toplam_kdv_tum * 100) if toplam_kdv_tum > 0 else 0.0

    genel_firmalar = (
        df_tum.groupby(["VKN", "Unvan"], as_index=False)
        .agg(Fatura_Adedi=("SiraNo", "count"), KDV_Tutari=("KDV", "sum"))
        .sort_values("KDV_Tutari", ascending=False)
        .reset_index(drop=True)
    )
    genel_firmalar["KDV_Payi_%"] = (genel_firmalar["KDV_Tutari"] / toplam_kdv_tum * 100) if toplam_kdv_tum > 0 else 0.0
    genel_firmalar["Kumulatif_%"] = genel_firmalar["KDV_Payi_%"].cumsum()
    secili_vkn_set = set(birlesik_ozet.loc[birlesik_ozet["Secili"], "VKN"]) if not birlesik_ozet.empty else set()
    genel_firmalar["Secili"] = genel_firmalar["VKN"].isin(secili_vkn_set)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Toplam 191 KDV", tl_str(toplam_kdv_tum))
    c2.metric("Seçili Firmalar KDV", tl_str(secili_kdv_tum))
    c3.metric("İndirim / KDV Oranı", f"%{indirim_pct_tum:.2f}")
    c4.metric("Oluşacak Firma Sayfası", f"{len(sayfa_paketleri)} Sayfa")

    st.subheader("1️⃣ Ön İzleme — Tüm Firmalar ve Toplam KDV Listesi")
    tum_firmalari_liste_goster(genel_firmalar, toplam_kdv_tum, secili_kdv_tum)

    excel_cikti = firma_sayfali_excel_olustur(sayfa_paketleri, tutar_secimi, aciklama_yaz)
    st.download_button(
        label=f"📥 Tüm Karşıt İnceleme Firmalarını Tek Excel'de İndir ({len(sayfa_paketleri)} Ayrı Sayfa)",
        data=excel_cikti,
        file_name="Karsit_Inceleme_Tum_Aylar_Firmalar.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

else:
    df_ay = df_tum if secili_ay == "TÜM DÖNEM (TEK TOPLAM)" else df_tum[df_tum["Ay"] == secili_ay].copy()
    firma_tablosu, toplam_kdv = firma_ozet_hesapla(df_ay)

    st.subheader("1️⃣ Karşıt İnceleme Kapsamındaki Firmalar (İstediğini Seç/Kaldır)")
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
    secili_kdv = float(secili_firmalar["KDV_Tutari"].sum()) if not secili_firmalar.empty else 0.0
    indirim_pct = (secili_kdv / toplam_kdv * 100) if toplam_kdv > 0 else 0.0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Dönem Toplam KDV", tl_str(toplam_kdv))
    c2.metric("Seçili Firmalar KDV", tl_str(secili_kdv))
    c3.metric("İndirim / KDV Oranı", f"%{indirim_pct:.2f}")
    c4.metric("Seçili Firma / Sayfa", f"{len(secili_firmalar)} / {len(onay_tablosu)}")

    st.subheader("2️⃣ Ön İzleme — Tüm Firmalar Listesi ve İndirim Oranı")
    tum_firmalari_liste_goster(onay_tablosu, toplam_kdv, secili_kdv)

    sayfa_paketleri = []
    for _, f_r in secili_firmalar.iterrows():
        f_fat = df_ay[df_ay["VKN"] == f_r["VKN"]].sort_values("Tarih_dt")
        sayfa_paketleri.append((f_r["Unvan"], f_fat))

    if not sayfa_paketleri:
        st.warning("En az bir firma seçili olmalıdır.")
        st.stop()

    excel_cikti = firma_sayfali_excel_olustur(sayfa_paketleri, tutar_secimi, aciklama_yaz)
    st.download_button(
        label=f"📥 Seçili {len(sayfa_paketleri)} Firmayı Tek Excel'de (Ayrı Sayfalarla) İndir",
        data=excel_cikti,
        file_name=f"Karsit_Inceleme_Firmalar_{secili_ay}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )
