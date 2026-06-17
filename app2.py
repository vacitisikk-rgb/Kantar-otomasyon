import streamlit as st
import pandas as pd
import json
import os
import re

# Sayfa Genişlik Ayarı
st.set_page_config(page_title="Kantar Vardiya Özeti", layout="wide")
st.title("🚚 Kantar Çoklu Vardiya Özeti ve Sağlaması Otomasyonu")

KURAL_DOSYASI = "kantar_kurallari.json"

# Türkçe karakterleri temizleme fonksiyonu (Arama garantisi için)
def turkce_temizle(metin):
    metin = str(metin).lower().strip()
    metin = metin.replace('ı', 'i').replace('ş', 's').replace('ğ', 'g').replace('ç', 'c').replace('ü', 'u').replace('ö', 'o')
    # Sadece harf ve rakamları bırak
    metin = re.sub(r'[^a-z0-9\s]', '', metin)
    return metin

# Varsayılan Şablon Kuralları
def kurallari_yukle():
    varsayilan = {
        "DENİŞ TÜVENAN KÖMÜR": {"kelime": "denis", "ek_kelime": "tuvenan"},
        "TÜRK PİYALE KÖMÜR": {"kelime": "piyale", "ek_kelime": ""},
        "KENT ÇİM TABAN KÜLÜ": {"kelime": "kent", "ek_kelime": "taban"},
        "KENT ÇİM UÇUCU KÜL": {"kelime": "kent", "ek_kelime": "ucucu"},
        "SOMA ÇİMENTO TABAN KÜLÜ": {"kelime": "soma", "ek_kelime": "taban"},
        "SOMA ÇİMENTO KOOP. TABAN": {"kelime": "koop", "ek_kelime": ""},
        "SOMA ÇİMENTO UÇUCU KÜL": {"kelime": "soma", "ek_kelime": "ucucu"},
        "BATISÖKE KÜL": {"kelime": "soke", "ek_kelime": ""},
        "BATIÇİM UÇUCU KÜL (T)": {"kelime": "bati", "ek_kelime": "t"},
        "BATIÇİM UÇUCU KÜL": {"kelime": "bati", "ek_kelime": "ucucu"},
        "LİMAK TABAN KÜLÜ": {"kelime": "limak", "ek_kelime": "taban"},
        "LİMAK UÇUCU KÜL": {"kelime": "limak", "ek_kelime": "ucucu"},
        "ÇİMENTAŞ": {"kelime": "cimentas", "ek_kelime": ""},
        "ALTIN ÇİMENTO TABAN KÜLÜ": {"kelime": "altin", "ek_kelime": ""},
        "NARETRA UÇUCU KÜL": {"kelime": "naretra", "ek_kelime": ""}
    }
    if os.path.exists(KURAL_DOSYASI):
        with open(KURAL_DOSYASI, "r", encoding="utf-8") as f:
            return json.load(f)
    return varsayilan

def kurallari_kaydet(kurallar):
    with open(KURAL_DOSYASI, "w", encoding="utf-8") as f:
        json.dump(kurallar, f, ensure_ascii=False, indent=4)

if 'kurallar' not in st.session_state:
    st.session_state.kurallar = kurallari_yukle()

# --- YAN MENÜ ---
st.sidebar.header("⚙️ Firma / Kural Yönetimi")
yeni_baslik = st.sidebar.text_input("Rapor Başlığı (Örn: ENGIN KÖMÜR)").upper().strip()
aranacak_kelime = st.sidebar.text_input("Aranacak Anahtar Kelime (Örn: ENGIN)").strip()

if st.sidebar.button("➕ Yeni Firmayı / Kuralı Ekle"):
    if yeni_baslik and aranacak_kelime:
        st.session_state.kurallar[yeni_baslik] = {"kelime": turkce_temizle(aranacak_kelime), "ek_kelime": ""}
        kurallari_kaydet(st.session_state.kurallar)
        st.sidebar.success(f"'{yeni_baslik}' başarıyla eklendi!")
        st.rerun()
    else:
        st.sidebar.error("Lütfen tüm alanları doldurun.")

st.sidebar.subheader("🗑️ Kayıtlı Kuralları Sil")
silinecek = st.sidebar.selectbox("Silmek istediğiniz kural", list(st.session_state.kurallar.keys()))
if st.sidebar.button("❌ Seçili Kuralı Sil"):
    del st.session_state.kurallar[silinecek]
    kurallari_kaydet(st.session_state.kurallar)
    st.sidebar.warning(f"'{silinecek}' silindi.")
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
        
        sutun_haritasi = {
            'NET AĞIRLIK': 'NET_AGIRLIK', 'NET AGIRLIK': 'NET_AGIRLIK', 'NET': 'NET_AGIRLIK',
            'PLAKA': 'PLAKA', 'ARAÇ': 'PLAKA', 'ARAC': 'PLAKA',
            'MALZEME': 'MALZEME', 'ÜRÜN': 'MALZEME', 'URUN': 'MALZEME',
            'NEREDEN GELDİ': 'NEREDEN_GELDI', 'NEREDEN GELDI': 'NEREDEN_GELDI', 'GELDİĞİ YER': 'NEREDEN_GELDI',
            'NEREYE GİTTİ': 'NEREYE_GITTI', 'NEREYE GITTI': 'NEREYE_GITTI', 'GİTTİĞİ YER': 'NEREYE_GITTI',
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
        
        # Karakter temizleme işlemi uygulanıyor
        df['ARAMA_HAVUZU'] = (df['FIRMA'].astype(str) + " " + 
                              df['MALZEME'].astype(str) + " " + 
                              df['NEREDEN_GELDI'].astype(str) + " " + 
                              df['NEREYE_GITTI'].astype(str)).apply(turkce_temizle)

        # --- HESAPLAMALAR ---
        rapor_listesi = []
        toplam_arac = 0
        toplam_tonaj = 0
        
        for isim, detay in st.session_state.kurallar.items():
            kelime = turkce_temizle(detay["kelime"])
            ek_kelime = turkce_temizle(detay.get("ek_kelime", ""))
            
            alt_df = df[df['ARAMA_HAVUZU'].str.contains(kelime, na=False)]
            if ek_kelime:
                alt_df = alt_df[alt_df['ARAMA_HAVUZU'].str.contains(ek_kelime, na=False)]
            
            a_sayisi = len(alt_df)
            t_sum = alt_df['NET_AGIRLIK'].sum()
            toplam_arac += a_sayisi
            toplam_tonaj += t_sum
            
            rapor_listesi.append({"AÇIKLAMA / MALZEME": isim, "ARAÇ SAYISI": a_sayisi, "TONAJ (kg)": f"{t_sum:,}".replace(",", ".")})
            
        rapor_df = pd.DataFrame(rapor_listesi)
        toplam_df = pd.DataFrame([{"AÇIKLAMA / MALZEME": "🏆 TOPLAM SEFER/TONAJ", "ARAÇ SAYISI": toplam_arac, "TONAJ (kg)": f"{toplam_tonaj:,}".replace(",", ".")}])
        
        # SAĞLAMA BÖLÜMÜ
        kolin_df = df[df['ARAMA_HAVUZU'].str.contains("kolin", na=False)]
        tuvenan_df = df[df['ARAMA_HAVUZU'].str.contains("tuvenan", na=False)]
        hidrogen_df = df[df['ARAMA_HAVUZU'].str.contains("hidro|engin", na=False)]
        kul_df = df[df['ARAMA_HAVUZU'].str.contains("kul", na=False)]
        cimento_maden_df = df[df['ARAMA_HAVUZU'].str.contains("soma|bati|kent|naretra|maden|cimen|limak", na=False)]
        hurda_df = df[df['ARAMA_HAVUZU'].str.contains("hurda", na=False)]
        
        saglama_listesi = [
            {"SAĞLAMASI": "KOLİN", "ARAÇ SAYISI": len(kolin_df), "TONAJ": f"{kolin_df['NET_AGIRLIK'].sum():,}".replace(",", ".")},
            {"SAĞLAMASI": "TÜVENAN KÖMÜR", "ARAÇ SAYISI": len(tuvenan_df), "TONAJ": f"{tuvenan_df['NET_AGIRLIK'].sum():,}".replace(",", ".")},
            {"SAĞLAMASI": "HİDRO-GEN A.Ş.", "ARAÇ SAYISI": len(hidrogen_df), "TONAJ": f"{hidrogen_df['NET_AGIRLIK'].sum():,}".replace(",", ".")},
            {"SAĞLAMASI": "TABAN-UÇUCU KÜL-UÇUCU KÜL T", "ARAÇ SAYISI": len(kul_df), "TONAJ": f"{kul_df['NET_AGIRLIK'].sum():,}".replace(",", ".")},
            {"SAĞLAMASI": "SOMA-BATI-KENTÇİM-NARETRA-TKS...", "ARAÇ SAYISI": len(cimento_maden_df), "TONAJ": f"{cimento_maden_df['NET_AGIRLIK'].sum():,}".replace(",", ".")},
            {"SAĞLAMASI": "HİDRO-GEN HURDA VS.", "ARAÇ SAYISI": len(hurda_df), "TONAJ": f"{hurda_df['NET_AGIRLIK'].sum():,}".replace(",", ".")}
        ]
        saglama_df = pd.DataFrame(saglama_listesi)

        # --- EKRANDA GÖSTERİM (2 KOLON) ---
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
