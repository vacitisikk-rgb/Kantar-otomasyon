import streamlit as st
import pandas as pd
import json
import os

# Sayfa Genişlik Ayarı
st.set_page_config(page_title="Kantar Vardiya Özeti", layout="wide")
st.title("🚚 Kantar Çoklu Vardiya Özeti ve Sağlaması Otomasyonu")

KURAL_DOSYASI = "kantar_kurallari.json"

# Harf ve Boşluk Temizleme (Garantili Eşleşme İçin)
def temizle(metin):
    metin = str(metin).upper().strip()
    metin = metin.replace('I', 'İ').replace('Ş', 'S').replace('Ğ', 'G').replace('Ç', 'C').replace('Ü', 'U').replace('Ö', 'O').replace('İ', 'I')
    return metin

# Tam Eşleşme Şablonu
def kurallari_yukle():
    varsayilan = {
        "DENİŞ TÜVENAN KÖMÜR": {"kelime1": "DENİS", "kelime2": "TÜVENAN"},
        "TÜRK PİYALE KÖMÜR": {"kelime1": "PİYALE", "kelime2": ""},
        "KENT ÇİM TABAN KÜLÜ": {"kelime1": "KENT", "kelime2": "TABAN"},
        "KENT ÇİM UÇUCU KÜL": {"kelime1": "KENT", "kelime2": "UÇUCU"},
        "SOMA ÇİMENTO TABAN KÜLÜ": {"kelime1": "SOMA", "kelime2": "TABAN"},
        "SOMA ÇİMENTO KOOP. TABAN": {"kelime1": "KOOP", "kelime2": ""},
        "SOMA ÇİMENTO UÇUCU KÜL": {"kelime1": "SOMA", "kelime2": "UÇUCU"},
        "BATISÖKE KÜL": {"kelime1": "SÖKE", "kelime2": ""},
        "BATIÇİM UÇUCU KÜL (T)": {"kelime1": "BATI", "kelime2": "(T)"},
        "BATIÇİM UÇUCU KÜL": {"kelime1": "BATI", "kelime2": "UÇUCU"},
        "LİMAK TABAN KÜLÜ": {"kelime1": "LİMAK", "kelime2": "TABAN"},
        "LİMAK UÇUCU KÜL": {"kelime1": "LİMAK", "kelime2": "UÇUCU"},
        "ÇİMENTAŞ": {"kelime1": "ÇİMENTAŞ", "kelime2": ""},
        "ALTIN ÇİMENTO TABAN KÜLÜ": {"kelime1": "ALTIN", "kelime2": ""},
        "NARETRA UÇUCU KÜL": {"kelime1": "NARETRA", "kelime2": ""}
    }
    if os.path.exists(KURAL_DOSYASI):
        with open(KURAL_DOSYASI, "r", encoding="utf-8") as f:
            return json.load(f)
    return varsayilan

if 'kurallar' not in st.session_state:
    st.session_state.kurallar = kurallari_yukle()

# --- YAN MENÜ ---
st.sidebar.header("⚙️ Firma / Kural Yönetimi")
yeni_baslik = st.sidebar.text_input("Rapor Başlığı (Örn: ENGIN KÖMÜR)").upper().strip()
aranacak_k1 = st.sidebar.text_input("Aranacak 1. Kelime (Örn: ENGIN)").upper().strip()
aranacak_k2 = st.sidebar.text_input("Aranacak 2. Kelime (İsteğe Bağlı)").upper().strip()

if st.sidebar.button("➕ Yeni Firmayı / Kuralı Ekle"):
    if yeni_baslik and aranacak_k1:
        st.session_state.kurallar[yeni_baslik] = {"kelime1": aranacak_k1, "kelime2": aranacak_k2}
        st.sidebar.success(f"'{yeni_baslik}' başarıyla eklendi!")
        st.rerun()

# --- ANA EKRAN ---
yuklenen_dosyalar = st.file_uploader("📂 Kantar Excel Dosyalarını Buraya Sürükleyin veya Seçin", type=["xlsx", "xls"], accept_multiple_files=True)

if yuklenen_dosyalar:
    try:
        tüm_veriler = []
        for dosya in yuklenen_dosyalar:
            tek_df = pd.read_excel(dosya)
            tek_df.columns = [str(c).strip().upper() for c in tek_df.columns]
            tüm_veriler.append(tek_df)
            
        df = pd.concat(tüm_veriler, ignore_index=True)
        
        # Sütun İsimlerini Standartlaştırma
        sutun_haritasi = {
            'NET AĞIRLIK': 'NET_AGIRLIK', 'NET AGIRLIK': 'NET_AGIRLIK', 'NET': 'NET_AGIRLIK',
            'PLAKA': 'PLAKA', 'MALZEME': 'MALZEME', 'ÜRÜN': 'MALZEME',
            'NEREDEN GELDİ': 'NEREDEN_GELDI', 'NEREDEN GELDI': 'NEREDEN_GELDI',
            'NEREYE GİTTİ': 'NEREYE_GITTI', 'NEREYE GITTI': 'NEREYE_GITTI',
            'FİRMA': 'FIRMA', 'FIRMA': 'FIRMA'
        }
        
        yeni_sutunlar = {}
        for c in df.columns:
            for anahtar, hedef in sutun_haritasi.items():
                if anahtar in c:
                    yeni_sutunlar[c] = hedef
                    break
        df = df.rename(columns=yeni_sutunlar)
        
        gerekli = ['PLAKA', 'MALZEME', 'NEREDEN_GELDI', 'NEREYE_GITTI', 'NET_AGIRLIK', 'FIRMA']
        for g in gerekli:
            if g not in df.columns:
                df[g] = 0 if g == 'NET_AGIRLIK' else ""
        
        df['NET_AGIRLIK'] = pd.to_numeric(df['NET_AGIRLIK'], errors='coerce').fillna(0).astype(int)
        
        # Arama havuzunu hem normal hem temizlenmiş formatta düz metne çeviriyoruz
        df['ARAMA_DUZ'] = (df['FIRMA'].astype(str) + " " + df['MALZEME'].astype(str) + " " + df['NEREDEN_GELDI'].astype(str) + " " + df['NEREYE_GITTI'].astype(str)).str.upper()
        df['ARAMA_TEMIZ'] = df['ARAMA_DUZ'].apply(temizle)

        # --- HESAPLAMALAR ---
        rapor_listesi = []
        toplam_arac = 0
        toplam_tonaj = 0
        
        for isim, detay in st.session_state.kurallar.items():
            k1 = detay["kelime1"].upper()
            k2 = detay["kelime2"].upper() if "kelime2" in detay else ""
            
            # Hem Türkçe karakterli hem karaktersiz havuzda eş zamanlı arama yapıyoruz
            mask = (df['ARAMA_DUZ'].str.contains(k1, na=False) | df['ARAMA_TEMIZ'].str.contains(temizle(k1), na=False))
            if k2:
                mask = mask & (df['ARAMA_DUZ'].str.contains(k2, na=False) | df['ARAMA_TEMIZ'].str.contains(temizle(k2), na=False))
                
            alt_df = df[mask]
            a_sayisi = len(alt_df)
            t_sum = alt_df['NET_AGIRLIK'].sum()
            
            toplam_arac += a_sayisi
            toplam_tonaj += t_sum
            
            rapor_listesi.append({"AÇIKLAMA / MALZEME": isim, "ARAÇ SAYISI": a_sayisi, "TONAJ (kg)": f"{t_sum:,}".replace(",", ".")})
            
        rapor_df = pd.DataFrame(rapor_listesi)
        toplam_df = pd.DataFrame([{"AÇIKLAMA / MALZEME": "🏆 TOPLAM SEFER/TONAJ", "ARAÇ SAYISI": toplam_arac, "TONAJ (kg)": f"{toplam_tonaj:,}".replace(",", ".")}])
        
        # SAĞLAMA BÖLÜMÜ
        kolin_df = df[df['ARAMA_DUZ'].str.contains("KOLİN|KOLIN", na=False)]
        tuvenan_df = df[df['ARAMA_DUZ'].str.contains("TÜVENAN|TUVENAN", na=False)]
        hidrogen_df = df[df['ARAMA_DUZ'].str.contains("HİDRO|HIDRO|ENGİN|ENGIN", na=False)]
        kul_df = df[df['ARAMA_DUZ'].str.contains("KÜL|KUL", na=False)]
        cimento_maden_df = df[df['ARAMA_DUZ'].str.contains("SOMA|BATI|KENT|NARETRA|MADEN|ÇİMEN|CIMEN|LİMAK|LIMAK", na=False)]
        hurda_df = df[df['ARAMA_DUZ'].str.contains("HURDA", na=False)]
        
        saglama_listesi = [
            {"SAĞLAMASI": "KOLİN", "ARAÇ SAYISI": len(kolin_df), "TONAJ": f"{kolin_df['NET_AGIRLIK'].sum():,}".replace(",", ".")},
            {"SAĞLAMASI": "TÜVENAN KÖMÜR", "ARAÇ SAYISI": len(tuvenan_df), "TONAJ": f"{tuvenan_df['NET_AGIRLIK'].sum():,}".replace(",", ".")},
            {"SAĞLAMASI": "HİDRO-GEN A.Ş.", "ARAÇ SAYISI": len(hidrogen_df), "TONAJ": f"{hidrogen_df['NET_AGIRLIK'].sum():,}".replace(",", ".")},
            {"SAĞLAMASI": "TABAN-UÇUCU KÜL-UÇUCU KÜL T", "ARAÇ SAYISI": len(kul_df), "TONAJ": f"{kul_df['NET_AGIRLIK'].sum():,}".replace(",", ".")},
            {"SAĞLAMASI": "SOMA-BATI-KENTÇİM-NARETRA-TKS...", "ARAÇ SAYISI": len(cimento_maden_df), "TONAJ": f"{cimento_maden_df['NET_AGIRLIK'].sum():,}".replace(",", ".")},
            {"SAĞLAMASI": "HİDRO-GEN HURDA VS.", "ARAÇ SAYISI": len(hurda_df), "TONAJ": f"{hurda_df['NET_AGIRLIK'].sum():,}".replace(",", ".")}
        ]
        saglama_df = pd.DataFrame(saglama_listesi)

        # --- EKRANDA GÖSTERİM ---
        sol_kolon, sag_kolon = st.columns([6, 5])
        with sol_kolon:
            st.subheader(f"📋 Birleştirilmiş Vardiya Özet Taslağı ({len(yuklenen_dosyalar)} Dosya)")
            st.dataframe(rapor_df, use_container_width=True, hide_index=True)
            st.dataframe(toplam_df, use_container_width=True, hide_index=True)
            
            st.subheader("✅ Vardiya Sağlaması")
            st.dataframe(saglama_df, use_container_width=True, hide_index=True)
            
        with sag_kolon:
            st.subheader("🔍 Birleştirilmiş Tüm Kantar Listesi Satırları")
            st.dataframe(df[['PLAKA', 'FIRMA', 'MALZEME', 'NEREDEN_GELDI', 'NET_AGIRLIK']], use_container_width=True, hide_index=True)
            
    except Exception as e:
        st.error(f"Excel dosyaları işlenirken bir hata oluştu: {e}")
else:
    st.info(f"💡 Lütfen yukarıdaki alana bir veya birden fazla kantar Excel dosyasını yükleyin.")
