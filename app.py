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


@st.cache_data
def load_data():
    # Başlıkların olduğu satırı doğru okumak için skiprows=15 yapıp ilk satırı kolon yapıyoruz
    df = pd.read_excel(file_path, sheet_name="Arıza Takip Listesi", skiprows=15)
    df.columns = df.iloc[
        0
    ]  # İlk satırı kolon isimleri olarak ata (Sıra, Şube Adı vs.)
    df = df.iloc[1:].reset_index(drop=True)  # Başlık satırını veri tabanından sil
    df = df.dropna(
        subset=["Sıra"]
    )  # Sıra numarası boş olan (boş) satırları temizle
    return df


df_ariza = load_data()

# Üst Özet Kartları (KPIs)
col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric("Toplam Arıza", len(df_ariza))
with col2:
    devam_eden = len(df_ariza[df_ariza["Durum"] == "Devam Ediyor"])
    st.metric("Devam Eden (Açık)", devam_eden)
with col3:
    kritik = len(df_ariza[df_ariza["Öncelik"] == "Kritik"])
    st.metric("Kritik Arızalar", kritik)
with col4:
    geciken = len(df_ariza[df_ariza["SLA Durumu"] == "Gecikti"])
    st.metric("SLA Geciken", geciken, delta_color="inverse")

st.markdown("### 📋 Arıza Takip Listesi")

# Filtreleme Seçenekleri
sube_filtre = st.selectbox(
    "Şube Seçin", ["Tümü"] + list(df_ariza["Şube Adı"].unique())
)
if sube_filtre != "Tümü":
    df_goster = df_ariza[df_ariza["Şube Adı"] == sube_filtre]
else:
    df_goster = df_ariza

# Tabloyu Gösterme
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

# Yeni Arıza Kaydı Formu (Sidebar)
st.sidebar.header("➕ Yeni Arıza Bildirimi")
with st.sidebar.form("ariza_form"):
    yeni_sube = st.selectbox(
        "Şube", ["Merkez Şube", "Şube 02 - Bahçelievler", "Şube 03 - Meydan"]
    )
    yeni_kategori = st.selectbox(
        "Kategori",
        [
            "Soğutma / Soğuk Zincir",
            "Kasa & IT Donanım",
            "Elektrik & Aydınlatma",
            "HVAC",
        ],
    )
    yeni_oncelik = st.selectbox("Öncelik", ["Kritik", "Normal", "Düşük"])
    yeni_aciklama = st.text_area("Arıza Açıklaması")

    submit = st.form_submit_button("Arıza Kaydı Oluştur")
    if submit:
        st.success("Arıza kaydı başarıyla oluşturuldu!")
