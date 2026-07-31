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


def ensure_excel_exists():
    if os.path.exists(file_path):
        try:
            df_test = pd.read_excel(
                file_path, sheet_name="Arıza Takip Listesi", skiprows=16
            )
            df_test = df_test.dropna(subset=["Sorun / Arıza Açıklaması"])
            if df_test.empty:
                os.remove(file_path)
        except:
            try:
                os.remove(file_path)
            except:
                pass

    if not os.path.exists(file_path):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Arıza Takip Listesi"
        ws.cell(
            row=1,
            column=2,
            value="YALÇIN MARKET - DETAYLI TEKNİK ARIZA VE MÜDAHALE TAKİP FORMU",
        )

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
        for col_idx, header in enumerate(headers):
            if header:
                ws.cell(row=17, column=col_idx, value=header)

        ornekler = [
            [
                None,
                1,
                "2026-07-29 08:30",
                "Merkez Şube",
                "Şarküteri reyonu soğutma derecesi yükseliyor",
                "Soğutma / Soğuk Zincir",
                "Kritik",
                "Tamamlandı",
                "Ali Usta",
                "Zamanında",
                "Kompresör gazı yenilendi, test edildi.",
                "Onaylandı",
            ],
            [
                None,
                2,
                "2026-07-29 09:15",
                "Şube 02 - Bahçelievler",
                "Kasa 2 barkod okuyucu temassızlık",
                "Kasa & IT Donanım",
                "Normal",
                "Devam Ediyor",
                "Caner Bey",
                "Devam Ediyor",
                "Kablo değişimi yapılacak.",
                "Bekliyor",
            ],
            [
                None,
                3,
                "2026-07-29 10:00",
                "Merkez Şube",
                "Depo koridor aydınlatma armatürü arızalı",
                "Elektrik & Aydınlatma",
                "Düşük",
                "Tamamlandı",
                "Ahmet Usta",
                "Zamanında",
                "LED ampul değiştirildi.",
                "Onaylandı",
            ],
        ]

        for idx, satir in enumerate(ornekler):
            for col_idx, val in enumerate(satir):
                if val is not None:
                    ws.cell(row=18 + idx, column=col_idx, value=val)

        wb.save(file_path)


ensure_excel_exists()


def load_data():
    try:
        df = pd.read_excel(file_path, sheet_name="Arıza Takip Listesi", skiprows=16)
        if "Sıra" not in df.columns and len(df.columns) > 1:
            df.columns = df.iloc[0]
            df = df.iloc[1:].reset_index(drop=True)
        df = df.dropna(subset=["Sorun / Arıza Açıklaması"])
        return df
    except Exception:
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

st.markdown("---")

# --- YÖNETİCİ DASHBOARD VE GRAFİKLER ---
st.markdown("### 📊 Yönetici Dashboard ve Analiz Grafikleri")

if not df_ariza.empty:
    g_col1, g_col2 = st.columns(2)

    with g_col1:
        st.markdown("#### Şubelere Göre Arıza Dağılımı")
        sube_data = df_ariza["Şube Adı"].value_counts()
        st.bar_chart(sube_data)

    with g_col2:
        st.markdown("#### Arıza Kategorileri Dağılımı")
        kat_data = df_ariza["Kategori"].value_counts()
        st.bar_chart(kat_data, color="#7c3aed")

st.markdown("---")
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
                "Çözüm Süresi / Açıklama",
            ]
        ],
        use_container_width=True,
    )

    # --- ARIZA DURUMU VE ÇÖZÜM GÜNCELLEME ---
    st.markdown("---")
    st.markdown("### ⚙️ Arıza Müdahale ve Durum Güncelleme")

    with st.form("guncelleme_formu"):
        col_g1, col_g2, col_g3 = st.columns(3)

        with col_g1:
            secilen_sira = st.selectbox(
                "İşlem Yapılacak Arıza (Sıra No)",
                df_ariza["Sıra"].astype(str).tolist(),
            )
        with col_g2:
            yeni_durum = st.selectbox(
                "Yeni Durum", ["Devam Ediyor", "Tamamlandı", "İptal Edildi"]
            )
        with col_g3:
            atanan_personel = st.text_input(
                "Atanan Teknik Personel", value="Ali Usta"
            )

        cozum_aciklamasi = st.text_area(
            "Çözüm Açıklaması / Yapılan İşlem",
            placeholder="Örn: Parça değiştirildi, test edildi.",
        )

        guncelle_submit = st.form_submit_button("Arızayı Güncelle")

        if guncelle_submit:
            try:
                wb = openpyxl.load_workbook(file_path)
                ws = wb["Arıza Takip Listesi"]

                bulundu = False
                for row in range(18, ws.max_row + 1):
                    sira_hucre = ws.cell(row=row, column=2).value
                    if sira_hucre is not None and str(sira_hucre) == str(
                        secilen_sira
                    ):
                        ws.cell(row=row, column=8, value=yeni_durum)
                        ws.cell(row=row, column=9, value=atanan_personel)
                        ws.cell(row=row, column=11, value=cozum_aciklamasi)
                        bulundu = True
                        break

                if bulundu:
                    wb.save(file_path)
                    st.success(
                        f"#{secilen_sira} numaralı arıza başarıyla güncellendi!"
                    )
                    st.rerun()
                else:
                    st.error("Arıza kaydı Excel dosyasında bulunamadı.")
            except Exception as e:
                st.error(f"Güncelleme sırasında hata oluştu: {e}")

else:
    st.info(
        "Henüz arıza kaydı bulunmuyor. Sol menüden ilk arıza kaydınızı oluşturabilirsiniz."
    )

# --- YENİ ARIZA KAYDI FORMU (SİDEBAR) ---
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
                ws = wb["Arıza Takip Listesi"]

                gercek_son_satir = 17
                for r in range(18, ws.max_row + 2):
                    val = ws.cell(row=r, column=5).value
                    if val is not None and str(val).strip() != "":
                        gercek_son_satir = r

                yeni_hedef_satir = gercek_son_satir + 1
                yeni_sira = len(df_ariza) + 1
                simdiki_zaman = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M")

                ws.cell(row=yeni_hedef_satir, column=2, value=yeni_sira)
                ws.cell(row=yeni_hedef_satir, column=3, value=simdiki_zaman)
                ws.cell(row=yeni_hedef_satir, column=4, value=yeni_sube)
                ws.cell(row=yeni_hedef_satir, column=5, value=yeni_aciklama)
                ws.cell(row=yeni_hedef_satir, column=6, value=yeni_kategori)
                ws.cell(row=yeni_hedef_satir, column=7, value=yeni_oncelik)
                ws.cell(row=yeni_hedef_satir, column=8, value="Devam Ediyor")
                ws.cell(row=yeni_hedef_satir, column=9, value="Atanmadı")
                ws.cell(row=yeni_hedef_satir, column=10, value="Zamanında")
                ws.cell(row=yeni_hedef_satir, column=11, value="İşlem Bekliyor")
                ws.cell(row=yeni_hedef_satir, column=12, value="Bekliyor")

                wb.save(file_path)
                st.sidebar.success(
                    "Arıza kaydı Excel dosyasına başarıyla eklendi!"
                )
                st.rerun()
            except Exception as e:
                st.sidebar.error(f"Kayıt eklenirken hata oluştu: {e}")
