import os
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Yalçın Market - Teknik Servis Takip",
    page_icon="🛠️",
    layout="wide",
)

st.title("🛠️ Yalçın Market Teknik Servis ve Arıza Takip Sistemi")
st.markdown("---")

file_path = "Yalcin_Market_Gelismis_Teknik_Servis_Takip_Sistemi.xlsx"


# Dosya yoksa otomatik oluşturan fonksiyon
def ensure_excel_exists():
    if not os.path.exists(file_path):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Arıza Takip Listesi"

        # Başlık satırlarını ve örnek tablo başlığını yazalım
        ws.cell(
            row=1,
            column=2,
            value="YALÇIN MARKET - DETAYLI TEKNİK ARIZA VE MÜDAHALE TAKİP FORMU",
        )
        ws.cell(row=2, column=2, value="SLA Süreçleri ve Durum Takibi")

        # Tablo başlıkları (16. satır)
        headers = [
            None,
            "Sıra",
            "Bildirim Tarih/Saat",
            "Şube Adı",
            "Sorun / Arıza Açıklaması",
            "Kategori",
            "Öncelik",
            "Durum",
            "Atanan Personel",
            "SLA Durumu",
            "Çözüm Süresi / Açıklama",
            "Yön. Onay",
        ]
        ws.row_dimensions[16].height = 25
        for col_idx, header in enumerate(headers):
            if header:
                ws.cell(row=16, column=col_idx, value=header)

        wb.save(file_path)


ensure_excel_exists()


@st.cache_data
def load_data():
    try:
        df = pd.read_excel(file_path, sheet_name="Arıza Takip Listesi", skiprows=15)
        if "Sıra" not in df.columns and len(df.columns) > 1:
            df.columns = df.iloc[0]
            df = df.iloc[1:].reset_index(drop=True)
        df = df.dropna(subset=["Sıra"])
        return df
    except Exception:
        # Eğer sayfa yapısı farklıysa boş bir şablon döndür
        return pd.DataFrame(
            columns=[
                "Sıra",
                "Bildirim Tarih/Saat",
                "Şube Adı",
                "Sorun / Arıza Açıklaması",
                "Kategori",
                "Öncelik",
                "Durum",
                "Atanan Personel",
                "SLA Durumu",
                "Çözüm Süresi / Açıklama",
                "Yön. Onay",
            ]
        )


df_ariza = load_data()

# Üst Özet Kartları (KPIs)
col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric("Toplam Arıza", len(df_ariza))
with col2:
    devam_eden = (
        len(df_ariza[df_ariza["Durum"] == "Devam Ediyor"])
        if not df_ariza.empty
        else 0
    )
    st.metric("Devam Eden (Açık)", devam_eden)
with col3:
    kritik = (
        len(df_ariza[df_ariza["Öncelik"] == "Kritik"])
        if not df_ariza.empty
        else 0
    )
    st.metric("Kritik Arızalar", kritik)
with col4:
    geciken = (
        len(df_ariza[df_ariza["SLA Durumu"] == "Gecikti"])
        if not df_ariza.empty
        else 0
    )
    st.metric("SLA Geciken", geciken, delta_color="inverse")

st.markdown("### 📋 Arıza Takip Listesi")

if not df_ariza.empty and "Şube Adı" in df_ariza.columns:
    sube_filtre = st.selectbox(
        "Şube Seçin", ["Tümü"] + list(df_ariza["Şube Adı"].dropna().unique())
    )
    if sube_filtre != "Tümü":
        df_goster = df_ariza[df_ariza["Şube Adı"] == sube_filtre]
    else:
        df_goster = df_ariza

    st.dataframe(
        df_goster[
            [
                "Sıra",
                "Bildirim Tarih/Saat",
                "Şube Adı",
                "Sorun / Arıza Açıklaması",
                "Kategori",
                "Öncelik",
                "Durum",
                "Atanan Personel",
                "SLA Durumu",
            ]
        ],
        use_container_width=True,
    )
else:
    st.info(
        "Henüz arıza kaydı bulunmuyor. Sol menüden ilk arıza kaydınızı oluşturabilirsiniz."
    )

# Yeni Arıza Kaydı Formu (Sidebar)
st.sidebar.header("➕ Yeni Arıza Bildirimi")
with st.sidebar.form("ariza_form"):
    yeni_sube = st.selectbox(
        "Şube Adı",
        [
            "Merkez Şube",
            "Şube 02 - Bahçelievler",
            "Şube 03 - Meydan",
            "Şube 05 - Çarşı",
        ],
    )
    yeni_kategori = st.selectbox(
        "Kategori",
        [
            "Soğutma / Soğuk Zincir",
            "Kasa & IT Donanım",
            "Elektrik & Aydınlatma",
            "HVAC (Klima / Havalandırma)",
            "Mekanik & Raf / Kapı",
        ],
    )
    yeni_oncelik = st.selectbox("Öncelik", ["Kritik", "Normal", "Düşük"])
    yeni_aciklama = st.text_area("Arıza Açıklaması")

    submit = st.form_submit_button("Arıza Kaydı Oluştur")

    if submit:
        if not yeni_aciklama.strip():
            st.sidebar.error("Lütfen bir arıza açıklaması girin!")
        else:
            try:
                wb = openpyxl.load_workbook(file_path)
                if "Arıza Takip Listesi" in wb.sheetnames:
                    ws = wb["Arıza Takip Listesi"]
                else:
                    ws = wb.active

                yeni_sira = len(df_ariza) + 1
                simdiki_zaman = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M")

                yeni_satir = [
                    None,
                    yeni_sira,
                    simdiki_zaman,
                    yeni_sube,
                    yeni_aciklama,
                    yeni_kategori,
                    yeni_oncelik,
                    "Devam Ediyor",
                    "Atanmadı",
                    "Zamanında",
                    "İşlem Bekliyor",
                    "Bekliyor",
                ]

                ws.append(yeni_satir)
                wb.save(file_path)

                st.sidebar.success(
                    "Arıza kaydı Excel dosyasına başarıyla eklendi!"
                )
                st.cache_data.clear()
                st.rerun()
            except Exception as e:
                st.sidebar.error(f"Kayıt eklenirken hata oluştu: {e}")
