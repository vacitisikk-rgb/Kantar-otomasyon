import streamlit as st
import pandas as pd
import json
import os

# Sayfa Genişlik Ayarı
st.set_page_config(page_title="Kantar Vardiya Özeti", layout="wide")
st.title("🚚 Kantar Çoklu Vardiya Özeti ve Sağlaması Otomasyonu")

KURAL_DOSYASI = "kantar_kurallari.json"

# Varsayılan Şablon Kuralları
def kurallari_yukle():
    varsayilan = {
        "DENİŞ TÜVENAN KÖMÜR": {"sutun": "Nereden Geldi", "kelime": "DENİŞ", "ek_sutun": "Malzeme", "ek_kelime": "TÜVENAN"},
        "TÜRK PİYALE KÖMÜR": {"sutun": "Nereden Geldi", "kelime": "PİYALE", "ek_sutun": "", "ek_kelime": ""},
        "KENT ÇİM TABAN KÜLÜ": {"sutun": "Firma", "kelime": "KENT", "ek_sutun": "Malzeme", "ek_kelime": "TABAN"},
        "KENT ÇİM UÇUCU KÜL": {"sutun": "Firma", "kelime": "KENT", "ek_sutun": "Malzeme", "ek_kelime": "UÇUCU"},
        "SOMA ÇİMENTO TABAN KÜLÜ": {"sutun": "Firma", "kelime": "SOMA", "ek_sutun": "Malzeme", "ek_kelime": "TABAN"},
        "SOMA ÇİMENTO KOOP. TABAN": {"sutun": "Firma", "kelime": "KOOP", "ek_sutun": "", "ek_kelime": ""},
        "SOMA ÇİMENTO UÇUCU KÜL": {"sutun": "Firma", "kelime": "SOMA", "ek_sutun": "Malzeme", "ek_kelime": "UÇUCU"},
        "BATISÖKE KÜL": {"sutun": "Firma", "kelime": "SÖKE", "ek_sutun": "", "ek_kelime": ""},
        "BATIÇİM UÇUCU KÜL (T)": {"sutun": "Firma", "kelime": "BATI", "ek_sutun": "Malzeme", "ek_kelime": "(T)"},
        "BATIÇİM UÇUCU KÜL": {"sutun": "Firma", "kelime": "BATI", "ek_sutun": "Malzeme", "ek_kelime": "UÇUCU"},
        "LİMAK TABAN KÜLÜ": {"sutun": "Firma", "kelime": "LİMAK", "ek_sutun": "Malzeme", "ek_kelime": "TABAN"},
        "LİMAK UÇUCU KÜL": {"sutun": "Firma", "kelime": "LİMAK", "ek_sutun": "Malzeme", "ek_kelime": "UÇUCU"},
        "ÇİMENTAŞ": {"sutun": "Firma", "kelime": "ÇİMENTAŞ", "ek_sutun": "", "ek_kelime": ""},
        "ALTIN ÇİMENTO TABAN KÜLÜ": {"sutun": "Firma", "kelime": "ALTIN", "ek_sutun": "", "ek_kelime": ""},
        "NARETRA UÇUCU KÜL": {"sutun": "Firma", "kelime": "NARETRA", "ek_sutun": "", "ek_kelime": ""}
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
sutun_secim = st.sidebar.selectbox("Aranacak Öncelikli Sütun", ["Firma", "Malzeme", "Nereden Geldi", "Nereye Gitti"])
aranacak_kelime = st.sidebar.text_input("Aranacak Anahtar Kelime (Örn: ENGIN)").upper().strip()

if st.sidebar.button("➕ Yeni Firmayı / Kuralı Ekle"):
    if yeni_baslik and aranacak_kelime:
        st.session_state.kurallar[yeni_baslik] = {"sutun": sutun_secim, "kelime": aranacak_kelime, "ek_sutun": "", "ek_kelime": ""}
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
            # Sütun isimlerindeki boşlukları temizle ve büyük harf yap
            tek_df.columns = [str(c).strip().upper() for c in tek_df.columns]
            tüm_veriler.append(tek_df)
            
        df = pd.concat(tüm_veriler, ignore_index=True)
        
        # Olası Türkçe/İngilizce sütun başlığı varyasyonlarını eşitle
        sutun_haritasi = {
            'NET AĞIRLIK': 'NET_AGIRLIK', 'NET AGIRLIK': 'NET_AGIRLIK', 'NET': 'NET_AGIRLIK',
            'PLAKA': 'PLAKA', 'ARAÇ': 'PLAKA', 'ARAC': 'PLAKA',
            'MALZEME': 'MALZEME', 'ÜRÜN': 'MALZEME', 'URUN': 'MALZEME',
            'NEREDEN GELDİ': 'NEREDEN_GELDI', 'NEREDEN GELDI': 'NEREDEN_GELDI', 'GELDİĞİ YER': 'NEREDEN_GELDI',
            'NEREYE GİTTİ': 'NEREYE_GITTI', 'NEREYE GITTI': 'NEREYE_GITTI', 'GİTTİĞİ YER': 'NEREYE_GITTI',
            'FİRMA': 'FIRMA', 'FIRMA': 'FIRMA'
        }
        
        # Mevcut sütunları haritaya göre yeniden adlandır
        yeni_sutunlar = {}
        for c in df.columns:
            for anahtar, hedef in sutun_haritasi.items():
                if anahtar in c:
                    yeni_sutunlar[c] = hedef
                    break
        df = df.rename(columns=yeni_sutunlar)
        
        # Eksik sütunları sıfırla/boş bırak
        gerekli = ['PLAKA', 'MALZEME', 'NEREDEN_GELDI', 'NEREYE_GITTI', 'NET_AGIRLIK', 'FIRMA']
        for g in gerekli:
            if g not in df.columns:
                df[g] = 0 if g == 'NET_AGIRLIK' else ""
        
        df['NET_AGIRLIK'] = pd.to_numeric(df['NET_AGIRLIK'], errors='coerce').fillna(0).astype(int)
        
        # Tüm metin alanlarını stringe çevir, büyüt ve temizle
        for col in ['PLAKA', 'MALZEME', 'NEREDEN_GELDI', 'NEREYE_GITTI', 'FIRMA']:
            df[col] = df[col].astype(str).str.upper().str.strip()
            
        # Akıllı kelime arama için tüm satırı birleştiren gizli bir arama sütunu oluşturuyoruz
        df['ARAMA_HAVUZU'] = df['FIRMA'] + " " + df['MALZEME'] + " " + df['NEREDEN_GELDI'] + " " + df['NEREYE_GITTI']

        # --- HESAPLAMALAR ---
        rapor_listesi = []
        toplam_arac = 0
        toplam_tonaj = 0
        
        for isim, detay in st.session_state.kurallar.items():
            kelime = detay["kelime"]
            ek_kelime = detay.get("ek_kelime", "")
            
            # Satırda aranan kelime geçiyor mu? (Sütun bağımsız akıllı arama)
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
        kolin_df = df[df['ARAMA_HAVUZU'].str.contains("KOLİN|KOLIN", na=False)]
        tuvenan_df = df[df['ARAMA_HAVUZU'].str.contains("TÜVENAN|TUVENAN", na=False)]
        hidrogen_df = df[df['ARAMA_HAVUZU'].str.contains("HİDRO|HIDRO|ENGİN|ENGIN", na=False)]
        kul_df = df[df['ARAMA_HAVUZU'].str.contains("KÜL|KUL", na=False)]
        cimento_maden_df = df[df['ARAMA_HAVUZU'].str.contains("SOMA|BATI|KENT|NARETRA|MADEN|ÇİMEN|CIMEN|LİMAK|LIMAK", na=False)]
        hurda_df = df[df['ARAMA_HAVUZU'].str.contains("HURDA", na=False)]
        
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
