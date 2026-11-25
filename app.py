import io
import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from docx import Document
from docx.shared import RGBColor

# --- SAYFA YAPILANDIRMASI ---
st.set_page_config(
    page_title="ProEnergy | GES & RES Fizibilite Platformu",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- 81 İL METEOROLOJİK VERİ SETİ (Yıllık Işınım kWh/m2-yıl, Ort. Rüzgar Hızı m/s) ---
ILLER_VERI = {
    "01 - Adana": (1685, 5.4), "02 - Adıyaman": (1710, 5.1), "03 - Afyonkarahisar": (1620, 5.8),
    "04 - Ağrı": (1560, 4.9), "05 - Amasya": (1480, 5.2), "06 - Ankara": (1585, 5.5),
    "07 - Antalya": (1765, 5.6), "08 - Artvin": (1340, 4.7), "09 - Aydın": (1695, 6.7),
    "10 - Balıkesir": (1540, 7.6), "11 - Bilecik": (1510, 6.1), "12 - Bingöl": (1630, 5.0),
    "13 - Bitlis": (1640, 5.2), "14 - Bolu": (1440, 5.3), "15 - Burdur": (1715, 5.6),
    "16 - Bursa": (1505, 6.4), "17 - Çanakkale": (1560, 8.1), "18 - Çankırı": (1520, 5.2),
    "19 - Çorum": (1500, 5.1), "20 - Denizli": (1680, 5.9), "21 - Diyarbakır": (1725, 5.1),
    "22 - Edirne": (1490, 6.8), "23 - Elazığ": (1665, 5.2), "24 - Erzincan": (1590, 5.0),
    "25 - Erzurum": (1580, 5.1), "26 - Eskişehir": (1565, 5.9), "27 - Gaziantep": (1735, 5.6),
    "28 - Giresun": (1330, 4.8), "29 - Gümüşhane": (1460, 4.9), "30 - Hakkari": (1675, 5.0),
    "31 - Hatay": (1715, 7.1), "32 - Isparta": (1710, 5.5), "33 - Mersin": (1745, 6.2),
    "34 - İstanbul": (1465, 7.2), "35 - İzmir": (1685, 7.8), "36 - Kars": (1545, 5.1),
    "37 - Kastamonu": (1425, 5.3), "38 - Kayseri": (1645, 6.0), "39 - Kırklareli": (1485, 7.1),
    "40 - Kırşehir": (1605, 5.7), "41 - Kocaeli": (1455, 6.5), "42 - Konya": (1690, 6.1),
    "43 - Kütahya": (1545, 5.7), "44 - Malatya": (1680, 5.4), "45 - Manisa": (1645, 7.3),
    "46 - Kahramanmaraş": (1705, 5.8), "47 - Mardin": (1755, 5.5), "48 - Muğla": (1730, 6.9),
    "49 - Muş": (1615, 4.9), "50 - Nevşehir": (1635, 5.8), "51 - Niğde": (1675, 6.0),
    "52 - Ordu": (1345, 4.9), "53 - Rize": (1310, 4.6), "54 - Sakarya": (1450, 5.8),
    "55 - Samsun": (1410, 5.5), "56 - Siirt": (1710, 4.9), "57 - Sinop": (1415, 6.3),
    "58 - Sivas": (1585, 5.9), "59 - Tekirdağ": (1495, 7.4), "60 - Tokat": (1490, 5.4),
    "61 - Trabzon": (1335, 4.7), "62 - Tunceli": (1625, 5.0), "63 - Şanlıurfa": (1760, 5.4),
    "64 - Uşak": (1625, 6.0), "65 - Van": (1685, 5.3), "66 - Yozgat": (1575, 5.6),
    "67 - Zonguldak": (1405, 5.6), "68 - Aksaray": (1660, 6.0), "69 - Bayburt": (1510, 5.0),
    "70 - Karaman": (1715, 6.2), "71 - Kırıkkale": (1580, 5.5), "72 - Batman": (1720, 5.0),
    "73 - Şırnak": (1730, 5.1), "74 - Bartın": (1410, 5.5), "75 - Ardahan": (1515, 5.0),
    "76 - Iğdır": (1575, 4.8), "77 - Yalova": (1475, 6.6), "78 - Karabük": (1435, 5.2),
    "79 - Kilis": (1740, 5.7), "80 - Osmaniye": (1700, 6.1), "81 - Düzce": (1435, 5.3)
}

AYLAR = ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"]
GUNES_DAGILIM = np.array([0.045, 0.055, 0.078, 0.092, 0.110, 0.122, 0.126, 0.116, 0.095, 0.072, 0.049, 0.040])
RUZGAR_PROFIL = np.array([1.14, 1.12, 1.06, 0.94, 0.86, 0.88, 0.96, 0.98, 0.92, 0.96, 1.05, 1.13])

# --- ÜST SANTRAL SEÇİMİ (GES / RES) ---
st.markdown("<h1 style='margin-bottom:0;'>⚡ ProEnergy Fizibilite Mühendislik Platformu</h1>", unsafe_allow_html=True)
st.markdown("<p style='color:#cbd5e1; margin-top:4px;'>Güneş (GES) ve Rüzgar (RES) Enerji Santralleri İçin Teknik Simülasyon ve Finansal Analiz Sistemi</p>", unsafe_allow_html=True)

santral_tipi = st.radio(
    "Santral Teknolojisini Seçin:",
    ["☀️ Güneş Enerjisi Santrali (GES)", "🌬️ Rüzgar Enerjisi Santrali (RES)"],
    horizontal=True
)
is_ges = "GES" in santral_tipi

# --- DİNAMİK ŞEFFAF ARKA PLAN VE MODERN GLASSMORPHISM CSS ---
bg_url = (
    "https://images.unsplash.com/photo-1509391366360-2e959784a276?auto=format&fit=crop&w=1920&q=80"
    if is_ges else
    "https://images.unsplash.com/photo-1466611653911-95081537e5b7?auto=format&fit=crop&w=1920&q=80"
)
accent_color = "#f59e0b" if is_ges else "#38bdf8"

st.markdown(f"""
<style>
    .stApp {{
        background: linear-gradient(rgba(11, 17, 32, 0.82), rgba(15, 23, 42, 0.88)), url("{bg_url}");
        background-size: cover;
        background-position: center;
        background-attachment: fixed;
        color: #f8fafc;
    }}
    [data-testid="stSidebar"] {{
        background-color: rgba(15, 23, 42, 0.88) !important;
        backdrop-filter: blur(12px);
        border-right: 1px solid rgba(255, 255, 255, 0.12);
    }}
    div[data-testid="stMetric"] {{
        background: rgba(30, 41, 59, 0.72);
        backdrop-filter: blur(10px);
        border: 1px solid rgba(255, 255, 255, 0.14);
        border-top: 3px solid {accent_color};
        padding: 16px 18px;
        border-radius: 12px;
        box-shadow: 0 8px 24px rgba(0,0,0,0.25);
    }}
    div[data-testid="stMetricLabel"] p {{
        color: #cbd5e1 !important;
        font-size: 0.9rem !important;
    }}
    div[data-testid="stMetricValue"] div {{
        color: #ffffff !important;
        font-weight: 700 !important;
    }}
    .stTabs [data-baseweb="tab-list"] {{
        background-color: rgba(30, 41, 59, 0.65);
        border-radius: 10px;
        padding: 4px;
    }}
    .stTabs [data-baseweb="tab"] {{
        color: #e2e8f0;
        border-radius: 8px;
    }}
    .stTabs [aria-selected="true"] {{
        background-color: {accent_color} !important;
        color: #0f172a !important;
        font-weight: 700;
    }}
</style>
""", unsafe_allow_html=True)

# --- SOL PANEL: KONUM, EKİPMAN VE FİNANSAL GİRDİLER ---
with st.sidebar:
    st.header("📍 1. Konum ve Meteoroloji")
    secilen_il = st.selectbox("Kurulum İli (81 İl)", list(ILLER_VERI.keys()), index=41)
    varsayilan_isinim, varsayilan_ruzgar = ILLER_VERI[secilen_il]

    if "aktif_il" not in st.session_state or st.session_state.aktif_il != secilen_il:
        st.session_state.aktif_il = secilen_il
        st.session_state.aylik_df = pd.DataFrame({
            "Ay": AYLAR,
            "Işınım (kWh/m²)": np.round(varsayilan_isinim * GUNES_DAGILIM, 1),
            "Rüzgar Hızı (m/s)": np.round(varsayilan_ruzgar * RUZGAR_PROFIL, 2)
        })

    st.divider()
    if is_ges:
        st.header("☀️ 2. GES Panel ve Sistem Girdileri")
        panel_secimi = st.selectbox(
            "Panel Modeli / Teknolojisi",
            [
                "Monokristal PERC 550 Wp (PR: %82)",
                "Bifacial N-Type 600 Wp (PR: %86)",
                "Polikristal 450 Wp (PR: %79)",
                "✏️ Özel Panel Değeri Gir..."
            ]
        )
        if "Özel" in panel_secimi:
            panel_adi = st.text_input("Özel Panel Marka / Model Adı", value="Özel Yüksek Verimli Panel")
            panel_gucu_w = st.number_input("Birim Panel Gücü (Wp)", min_value=100, max_value=1500, value=580, step=10)
            ges_pr = st.slider("Sistem Performans Oranı (PR %)", 50.0, 95.0, 84.0, 0.5) / 100.0
        else:
            panel_adi = panel_secimi.split(" (")[0]
            default_w = 550 if "550" in panel_secimi else (600 if "600" in panel_secimi else 450)
            default_pr = 0.82 if "82" in panel_secimi else (0.86 if "86" in panel_secimi else 0.79)
            panel_gucu_w = st.number_input("Birim Panel Gücü (Wp)", min_value=100, max_value=1500, value=default_w, step=10)
            ges_pr = st.slider("Sistem Performans Oranı (PR %)", 50.0, 95.0, default_pr * 100, 0.5) / 100.0

        kapasite_yontemi = st.radio("Kapasite Hesaplama Yöntemi", ["Panel Sayısı Gir", "Toplam Kurulu Güç (kWp) Gir"])
        if kapasite_yontemi == "Panel Sayısı Gir":
            panel_sayisi = st.number_input("Toplam Panel Sayısı (Adet)", min_value=1, max_value=1000000, value=2000, step=50)
            kurulu_guc_kw = (panel_gucu_w * panel_sayisi) / 1000.0
        else:
            kurulu_guc_kw = st.number_input("Toplam Kurulu Güç (kWp)", min_value=1.0, max_value=500000.0, value=1100.0, step=50.0)
            panel_sayisi = int(np.ceil((kurulu_guc_kw * 1000) / panel_gucu_w))

        st.success(f"Kurulu Güç: **{kurulu_guc_kw:,.1f} kWp** ({panel_sayisi:,} Panel)")
        ekipman_ozeti = f"{panel_adi} ({panel_gucu_w} Wp x {panel_sayisi:,} adet, PR: %{ges_pr*100:.1f})"

    else:
        st.header("🌬️ 2. RES Türbin ve Sistem Girdileri")
        turbin_secimi = st.selectbox(
            "Türbin Modeli",
            [
                "Standart Onshore 2.0 MW (Rotor: 114m)",
                "Yüksek Verim 3.6 MW (Rotor: 136m)",
                "Yeni Nesil 5.5 MW (Rotor: 158m)",
                "✏️ Özel Türbin Değeri Gir..."
            ]
        )
        if "Özel" in turbin_secimi:
            turbin_adi = st.text_input("Özel Türbin Marka / Model", value="Özel Rüzgar Türbini")
            turbin_gucu_mw = st.number_input("Birim Türbin Nominal Gücü (MW)", min_value=0.1, max_value=20.0, value=4.2, step=0.1)
            rotor_capi = st.number_input("Rotor Çapı (m)", min_value=20, max_value=260, value=145, step=2)
        else:
            turbin_adi = turbin_secimi.split(" (")[0]
            default_mw = 2.0 if "2.0" in turbin_secimi else (3.6 if "3.6" in turbin_secimi else 5.5)
            default_rotor = 114 if "114" in turbin_secimi else (136 if "136" in turbin_secimi else 158)
            turbin_gucu_mw = st.number_input("Birim Türbin Nominal Gücü (MW)", min_value=0.1, max_value=20.0, value=default_mw, step=0.1)
            rotor_capi = st.number_input("Rotor Çapı (m)", min_value=20, max_value=260, value=default_rotor, step=2)

        turbin_sayisi = st.number_input("Türbin Adedi", min_value=1, max_value=200, value=2, step=1)
        kurulu_guc_kw = turbin_gucu_mw * turbin_sayisi * 1000.0
        st.success(f"Kurulu Güç: **{kurulu_guc_kw:,.0f} kW ({kurulu_guc_kw/1000:.2f} MW)**")
        ekipman_ozeti = f"{turbin_adi} ({turbin_gucu_mw} MW x {turbin_sayisi} adet, Rotor: {rotor_capi}m)"

    st.divider()
    st.header("💰 3. Yatırım ve Finans Girdileri")
    varsayilan_capex = int(kurulu_guc_kw * (780 if is_ges else 1250))
    capex = st.number_input("Toplam Yatırım Maliyeti - CAPEX ($)", min_value=1000, max_value=1000000000, value=max(5000, varsayilan_capex), step=10000)
    opex_yillik = st.number_input("Yıllık İşletme & Bakım - OPEX ($/yıl)", min_value=0, max_value=100000000, value=int(capex * 0.018), step=1000)
    elektrik_fiyati = st.number_input("Elektrik Satış Birim Fiyatı ($/kWh)", min_value=0.01, max_value=1.00, value=0.085, step=0.005, format="%.3f")
    iskonto_orani = st.slider("İskonto / Faiz Oranı (%)", 1.0, 25.0, 8.0, 0.5) / 100.0
    sistem_omru = st.slider("Ekonomik Proje Ömrü (Yıl)", 10, 30, 25, 1)

# --- AYLIK VE YILLIK METEOROLOJİK VERİ DÜZENLEME ALANI ---
with st.expander(f"🛠️ {secilen_il} — Aylık ve Yıllık Güneş Işınımı & Rüzgar Hızı Değerlerini Düzenle", expanded=False):
    c_yillik1, c_yillik2, c_info = st.columns([1, 1, 2])
    mevcut_yillik_isinim = float(st.session_state.aylik_df["Işınım (kWh/m²)"].sum())
    mevcut_ort_ruzgar = float(st.session_state.aylik_df["Rüzgar Hızı (m/s)"].mean())

    with c_yillik1:
        yeni_yillik_isinim = st.number_input(
            "Yıllık Toplam Işınım Hedefi (kWh/m²-yıl)",
            min_value=500.0, max_value=3000.0, value=round(mevcut_yillik_isinim, 1), step=10.0
        )
        if abs(yeni_yillik_isinim - mevcut_yillik_isinim) > 0.5 and mevcut_yillik_isinim > 0:
            oran = yeni_yillik_isinim / mevcut_yillik_isinim
            st.session_state.aylik_df["Işınım (kWh/m²)"] = np.round(st.session_state.aylik_df["Işınım (kWh/m²)"] * oran, 1)
            st.rerun()

    with c_yillik2:
        yeni_ort_ruzgar = st.number_input(
            "Yıllık Ortalama Rüzgar Hızı Hedefi (m/s)",
            min_value=1.5, max_value=16.0, value=round(mevcut_ort_ruzgar, 2), step=0.1
        )
        if abs(yeni_ort_ruzgar - mevcut_ort_ruzgar) > 0.02 and mevcut_ort_ruzgar > 0:
            oran_r = yeni_ort_ruzgar / mevcut_ort_ruzgar
            st.session_state.aylik_df["Rüzgar Hızı (m/s)"] = np.round(st.session_state.aylik_df["Rüzgar Hızı (m/s)"] * oran_r, 2)
            st.rerun()

    with c_info:
        st.info("💡 Yukarıdan **yıllık toplam/ortalama** değeri değiştirdiğinizde 12 ay orantılı güncellenir; aşağıdaki tablodan **tek tek ayların** üzerine tıklayarak da istediğiniz ayın değerini manuel değiştirebilirsiniz.")

    duzenlenen_df = st.data_editor(
        st.session_state.aylik_df,
        use_container_width=True,
        hide_index=True,
        disabled=["Ay"],
        key="meteo_editor"
    )
    st.session_state.aylik_df = duzenlenen_df

# Güncel Aylık & Yıllık Değerler
aylik_isinim_serisi = st.session_state.aylik_df["Işınım (kWh/m²)"].to_numpy(dtype=float)
aylik_ruzgar_serisi = st.session_state.aylik_df["Rüzgar Hızı (m/s)"].to_numpy(dtype=float)
toplam_isinim_yillik = float(np.sum(aylik_isinim_serisi))
ortalama_ruzgar_yillik = float(np.mean(aylik_ruzgar_serisi))

# --- MÜHENDİSLİK VE FİNANS HESAPLAMA MOTORU ---
if is_ges:
    aylik_uretim_kwh = kurulu_guc_kw * aylik_isinim_serisi * ges_pr
    ilk_yil_uretim_kwh = float(np.sum(aylik_uretim_kwh))
    birim_uretim_bilgisi = f"Panel Başı: {ilk_yil_uretim_kwh / panel_sayisi:,.0f} kWh/yıl" if panel_sayisi > 0 else "-"
    ek_teknik_parametre = f"Performans Oranı (PR): %{ges_pr*100:.1f}"
else:
    tahmini_cf_genel = float(np.clip((ortalama_ruzgar_yillik - 2.5) * 0.072, 0.12, 0.62))
    with st.sidebar:
        res_cf = st.slider("Santral Kapasite Faktörü (CF %)", 10.0, 65.0, round(tahmini_cf_genel * 100, 1), 0.5) / 100.0
    ilk_yil_uretim_kwh = kurulu_guc_kw * 8760 * res_cf
    ruzgar_agirlik = (aylik_ruzgar_serisi ** 2.2) / np.sum(aylik_ruzgar_serisi ** 2.2)
    aylik_uretim_kwh = ilk_yil_uretim_kwh * ruzgar_agirlik
    birim_uretim_bilgisi = f"Türbin Başı: {(ilk_yil_uretim_kwh / turbin_sayisi)/1000:,.1f} MWh/yıl"
    ek_teknik_parametre = f"Kapasite Faktörü (CF): %{res_cf*100:.1f}"

kapasite_faktoru = (ilk_yil_uretim_kwh / (kurulu_guc_kw * 8760) * 100) if kurulu_guc_kw > 0 else 0.0

# 25 Yıllık Nakit Akışı Hesabı
nakit_kayitlari = []
kumulatif_nakit = -float(capex)
amortisman_yili = None
npv = -float(capex)
indirgenmis_enerji = 0.0
indirgenmis_maliyet = float(capex)

nakit_kayitlari.append({
    "Yıl": 0, "Üretim (MWh)": 0.0, "Brüt Gelir ($)": 0.0,
    "OPEX ($)": 0.0, "Net Nakit Akışı ($)": -capex, "Kümülatif Nakit ($)": kumulatif_nakit
})

for yil in range(1, sistem_omru + 1):
    degradasyon = (1 - 0.005) ** (yil - 1)
    yil_uretim = ilk_yil_uretim_kwh * degradasyon
    gelir = yil_uretim * elektrik_fiyati
    net_nakit = gelir - opex_yillik
    onceki_kum = kumulatif_nakit
    kumulatif_nakit += net_nakit

    if amortisman_yili is None and kumulatif_nakit >= 0:
        amortisman_yili = (yil - 1) + abs(onceki_kum) / net_nakit

    iskonto = (1 + iskonto_orani) ** yil
    npv += net_nakit / iskonto
    indirgenmis_enerji += yil_uretim / iskonto
    indirgenmis_maliyet += opex_yillik / iskonto

    nakit_kayitlari.append({
        "Yıl": yil,
        "Üretim (MWh)": round(yil_uretim / 1000, 1),
        "Brüt Gelir ($)": round(gelir, 0),
        "OPEX ($)": round(opex_yillik, 0),
        "Net Nakit Akışı ($)": round(net_nakit, 0),
        "Kümülatif Nakit ($)": round(kumulatif_nakit, 0)
    })

df_nakit = pd.DataFrame(nakit_kayitlari)
lcoe = (indirgenmis_maliyet / indirgenmis_enerji) if indirgenmis_enerji > 0 else 0.0
co2_tasarruf = (ilk_yil_uretim_kwh / 1000.0) * 0.48

# --- YÖNETİCİ ÖZETİ KARTLARI ---
k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Toplam Kurulu Güç", f"{kurulu_guc_kw:,.1f} kW", f"CF: %{kapasite_faktoru:.1f}")
k2.metric("İlk Yıl Üretimi", f"{ilk_yil_uretim_kwh/1000:,.1f} MWh", birim_uretim_bilgisi)
k3.metric("Amortisman Süresi", f"{amortisman_yili:.1f} Yıl" if amortisman_yili else f">{sistem_omru} Yıl", f"Yatırım: ${capex:,.0f}")
k4.metric("Net Bugünkü Değer (NPV)", f"${npv:,.0f}", f"LCOE: ${lcoe:.3f}/kWh")
k5.metric("Yıllık CO₂ Tasarrufu", f"{co2_tasarruf:,.1f} Ton", f"≈ {int(co2_tasarruf * 45):,} Ağaç")

st.markdown("<br>", unsafe_allow_html=True)

# --- SEKMELİ ANALİZ GÖRÜNÜMÜ ---
tab1, tab2, tab3 = st.tabs(["📊 Aylık Meteoroloji ve Üretim Grafikleri", "💰 Finansal Projeksiyon ve Nakit Akışı", "📥 Word & Excel Rapor İndir"])

with tab1:
    g1, g2 = st.columns(2)
    with g1:
        fig_meteo = go.Figure()
        if is_ges:
            fig_meteo.add_trace(go.Bar(
                x=AYLAR, y=aylik_isinim_serisi, name="Aylık Global Işınım (kWh/m²)",
                marker_color="#f59e0b"
            ))
            baslik_meteo = f"{secilen_il} — Aylık Güneş Işınımı (Toplam: {toplam_isinim_yillik:,.0f} kWh/m²-yıl)"
            y_ekseni = "Işınım (kWh/m²-ay)"
        else:
            fig_meteo.add_trace(go.Scatter(
                x=AYLAR, y=aylik_ruzgar_serisi, name="Aylık Ort. Rüzgar Hızı (m/s)",
                mode="lines+markers+text", text=np.round(aylik_ruzgar_serisi, 1), textposition="top center",
                line=dict(color="#38bdf8", width=4), marker=dict(size=9)
            ))
            baslik_meteo = f"{secilen_il} — Aylık Rüzgar Hızı Rejimi (Ort: {ortalama_ruzgar_yillik:.2f} m/s)"
            y_ekseni = "Rüzgar Hızı (m/s)"

        fig_meteo.update_layout(
            title=baslik_meteo, yaxis_title=y_ekseni,
            paper_bgcolor="rgba(15, 23, 42, 0.65)", plot_bgcolor="rgba(30, 41, 59, 0.45)",
            font=dict(color="#f8fafc"), height=400
        )
        st.plotly_chart(fig_meteo, use_container_width=True)

    with g2:
        fig_uretim = go.Figure()
        fig_uretim.add_trace(go.Bar(
            x=AYLAR, y=aylik_uretim_kwh / 1000.0,
            name="Aylık Elektrik Üretimi (MWh)",
            marker_color="#10b981" if is_ges else "#0ea5e9",
            text=np.round(aylik_uretim_kwh / 1000.0, 1), textposition="auto"
        ))
        fig_uretim.update_layout(
            title=f"Aylara Göre Tahmini Elektrik Üretimi (Toplam: {ilk_yil_uretim_kwh/1000:,.1f} MWh/yıl)",
            yaxis_title="Üretim (MWh)",
            paper_bgcolor="rgba(15, 23, 42, 0.65)", plot_bgcolor="rgba(30, 41, 59, 0.45)",
            font=dict(color="#f8fafc"), height=400
        )
        st.plotly_chart(fig_uretim, use_container_width=True)

with tab2:
    f1, f2 = st.columns([3, 2])
    with f1:
        fig_nakit = go.Figure()
        renkler = ["#ef4444" if val < 0 else "#10b981" for val in df_nakit["Kümülatif Nakit ($)"]]
        fig_nakit.add_trace(go.Bar(
            x=df_nakit["Yıl"], y=df_nakit["Kümülatif Nakit ($)"],
            marker_color=renkler, name="Kümülatif Net Nakit ($)"
        ))
        fig_nakit.add_hline(y=0, line_dash="dash", line_color="#f8fafc")
        fig_nakit.update_layout(
            title=f"{sistem_omru} Yıllık Kümülatif Nakit Akışı ve Geri Ödeme Noktası",
            xaxis_title="Proje Yılı", yaxis_title="Kümülatif Nakit ($)",
            paper_bgcolor="rgba(15, 23, 42, 0.65)", plot_bgcolor="rgba(30, 41, 59, 0.45)",
            font=dict(color="#f8fafc"), height=420
        )
        st.plotly_chart(fig_nakit, use_container_width=True)

    with f2:
        st.markdown("#### Yıllık Nakit Akışı Tablosu")
        st.dataframe(df_nakit, use_container_width=True, height=380, hide_index=True)

with tab3:
    st.markdown("### 📄 Profesyonel Fizibilite Raporu Çıktıları")
    st.write("Hazırlanan tüm teknik simülasyon, aylık meteoroloji tablosu ve 25 yıllık finansal projeksiyon verilerini tek tıkla **Excel (.xlsx)** veya **Word (.docx)** formatında indirebilirsiniz.")

    # --- EXCEL (.XLSX) OLUŞTURMA ---
    excel_buffer = io.BytesIO()
    ozet_df = pd.DataFrame({
        "Parametre": [
            "Santral Türü", "Kurulum İli", "Ekipman Konfigürasyonu", "Toplam Kurulu Güç (kW)",
            "Ek Teknik Parametre", "Yıllık Global Işınım (kWh/m²-yıl)", "Ortalama Rüzgar Hızı (m/s)",
            "İlk Yıl Enerji Üretimi (kWh)", "Kapasite Faktörü (%)", "Toplam Yatırım - CAPEX ($)",
            "Yıllık İşletme Gideri - OPEX ($)", "Elektrik Birim Satış Fiyatı ($/kWh)",
            "Amortisman Süresi (Yıl)", "Net Bugünkü Değer - NPV ($)", "LCOE ($/kWh)", "Yıllık CO2 Tasarrufu (Ton)"
        ],
        "Değer": [
            santral_tipi, secilen_il, ekipman_ozeti, round(kurulu_guc_kw, 2),
            ek_teknik_parametre, round(toplam_isinim_yillik, 1), round(ortalama_ruzgar_yillik, 2),
            round(ilk_yil_uretim_kwh, 0), round(kapasite_faktoru, 2), capex,
            opex_yillik, elektrik_fiyati,
            round(amortisman_yili, 2) if amortisman_yili else f">{sistem_omru}",
            round(npv, 2), round(lcoe, 4), round(co2_tasarruf, 1)
        ]
    })

    aylik_rapor_df = st.session_state.aylik_df.copy()
    aylik_rapor_df["Tahmini Üretim (MWh)"] = np.round(aylik_uretim_kwh / 1000.0, 2)

    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        ozet_df.to_excel(writer, sheet_name="Yönetici Özeti", index=False)
        aylik_rapor_df.to_excel(writer, sheet_name="Aylık Veriler", index=False)
        df_nakit.to_excel(writer, sheet_name="25 Yıllık Nakit Akışı", index=False)

    # --- WORD (.DOCX) OLUŞTURMA ---
    doc = Document()
    baslik = doc.add_heading("YENİLENEBİLİR ENERJİ SANTRALİ FİZİBİLİTE RAPORU", 0)
    baslik.runs[0].font.color.rgb = RGBColor(15, 23, 42)

    doc.add_paragraph(f"Proje Konumu: {secilen_il}  |  Santral Türü: {santral_tipi}")

    doc.add_heading("1. Teknik ve Finansal Yönetici Özeti", level=1)
    table = doc.add_table(rows=1, cols=2)
    table.style = "Light Grid Accent 1"
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "Parametre"
    hdr_cells[1].text = "Değer"

    for _, row in ozet_df.iterrows():
        row_cells = table.add_row().cells
        row_cells[0].text = str(row["Parametre"])
        row_cells[1].text = str(row["Değer"])

    doc.add_heading("2. Aylık Meteorolojik Potansiyel ve Üretim Dağılımı", level=1)
    t_ay = doc.add_table(rows=1, cols=4)
    t_ay.style = "Light Grid Accent 1"
    for idx, col_name in enumerate(aylik_rapor_df.columns):
        t_ay.rows[0].cells[idx].text = col_name

    for _, row in aylik_rapor_df.iterrows():
        rc = t_ay.add_row().cells
        rc[0].text = str(row["Ay"])
        rc[1].text = str(row["Işınım (kWh/m²)"])
        rc[2].text = str(row["Rüzgar Hızı (m/s)"])
        rc[3].text = str(row["Tahmini Üretim (MWh)"])

    word_buffer = io.BytesIO()
    doc.save(word_buffer)

    btn1, btn2 = st.columns(2)
    with btn1:
        st.download_button(
            label="📊 Excel Fizibilite Dosyasını İndir (.xlsx)",
            data=excel_buffer.getvalue(),
            file_name=f"Fizibilite_Raporu_{secilen_il[:2]}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with btn2:
        st.download_button(
            label="📝 Word Fizibilite Raporunu İndir (.docx)",
            data=word_buffer.getvalue(),
            file_name=f"Fizibilite_Raporu_{secilen_il[:2]}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )