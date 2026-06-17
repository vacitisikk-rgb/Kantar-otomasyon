import streamlit as st
import pandas as pd
import json
import os

# Sayfa Genişlik Ayarı
st.set_page_config(page_title="Kantar Vardiya Özeti", layout="wide")
st.title("🚚 Kantar Vardiya Özeti ve Sağlaması Otomasyonu")

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

# --- YAN MENÜ (Ayarlar & El İle Firma Ekleme) ---
st.sidebar.header("⚙️ Firma / Kural Yönetimi")
yeni_baslik = st.sidebar.text_input("Rapor Başlığı (Örn: ENGIN KÖMÜR)").upper().strip()
sutun_secim = st.sidebar.selectbox("Aranacak Sütun", ["Firma", "Malzeme", "Nereden Geldi", "Nereye Gitti"])
aranacak_kelime = st.sidebar.text_input("Aranacak Kelime Anahtarı (Örn: ENGIN)").upper().strip()

if st.sidebar.button("➕ Yeni Firmayı / Kuralı Ekle"):
    if yeni_baslik and aranacak_kelime:
        st.session_state.kurallar[yeni_baslik] = {"sutun": sutun_secim, "kelime": aranacak_kelime, "ek_sutun": "", "ek_kelime": ""}
        kurallari_kaydet(st.session_state.kurallar)
        st.sidebar.success(f"'{yeni_baslik}' başarıyla eklendi!")
        st.rerun()
    else:
        st.sidebar.error("Lütfen tüm alanları doldurun.")

# Mevcut Kuralları Silme Alanı
st.sidebar.subheader("🗑️ Kayıtlı Kuralları Sil")
silinecek = st.sidebar.selectbox("Silmek istediğiniz kural", list(st.session_state.kurallar.keys()))
if st.sidebar.button("❌ Seçili Kuralı Sil"):
    del st.session_state.kurallar[silinecek]
    kurallari_kaydet(st.session_state.kurallar)
    st.sidebar.warning(f"'{silinecek}' silindi.")
    st.rerun()

# --- ANA EKRAN (Excel Yükleme ve Raporlama) ---
yuklenen_dosya = st.file_uploader("📂 Kantar Excel Dosyasını Buraya Sürükleyin veya Seçin", type=["xlsx", "xls"])

if yuklenen_dosya:
    try:
        df = pd.read_excel(yuklenen_dosya)
        df.columns = [str(c).strip() for c in df.columns]
        
        # Sütunları standartlaştır
        gerekli = ['Plaka', 'Malzeme', 'Nereden Geldi', 'Nereye Gitti', 'Net Ağırlık', 'Firma']
        for g in gerekli:
            if g not in df.columns:
                df[g] = 0 if g == 'Net Ağırlık' else ""
        
        df['Net Ağırlık'] = pd.to_numeric(df['Net Ağırlık'], errors='coerce').fillna(0).astype(int)
        for col in ['Plaka', 'Malzeme', 'Nereden Geldi', 'Nereye Gitti', 'Firma']:
            df[col] = df[col].astype(str).str.upper().str.strip()

        # --- HESAPLAMALAR ---
        rapor_listesi = []
        toplam_arac = 0
        toplam_tonaj = 0
        
        for isim, detay in st.session_state.kurallar.items():
            alt_df = df[df[detay["sutun"]].str.contains(detay["kelime"], na=False)]
            if detay.get("ek_sutun") and detay.get("ek_kelime"):
                alt_df = alt_df[alt_df[detay["ek_sutun"]].str.contains(detay["ek_kelime"], na=False)]
            
            a_sayisi = len(alt_df)
            t_sum = alt_df['Net Ağırlık'].sum()
            toplam_arac += a_sayisi
            toplam_tonaj += t_sum
            
            rapor_listesi.append({"AÇIKLAMA / MALZEME": isim, "ARAÇ SAYISI": a_sayisi, "TONAJ (kg)": f"{t_sum:,}".replace(",", ".")})
            
        rapor_df = pd.DataFrame(rapor_listesi)
        toplam_df = pd.DataFrame([{"AÇIKLAMA / MALZEME": "🏆 TOPLAM SEFER/TONAJ", "ARAÇ SAYISI": toplam_arac, "TONAJ (kg)": f"{toplam_tonaj:,}".replace(",", ".")}])
        
        # SAĞLAMA BÖLÜMÜ
        kolin_df = df[df['Firma'].str.contains("KOLİN", na=False)]
        tuvenan_df = df[df['Malzeme'].str.contains("TÜVENAN", na=False)]
        hidrogen_df = df[df['Firma'].str.contains("HİDRO", na=False) | df['Firma'].str.contains("ENGIN", na=False)]
        kul_df = df[df['Malzeme'].str.contains("KÜL", na=False)]
        cimento_maden_df = df[df['Firma'].str.contains("SOMA|BATI|KENT|NARETRA|MADEN|ÇİMEN|LİMAK", na=False)]
        hurda_df = df[df['Malzeme'].str.contains("HURDA", na=False)]
        
        saglama_listesi = [
            {"SAĞLAMASI": "KOLİN", "ARAÇ SAYISI": len(kolin_df), "TONAJ": f"{kolin_df['Net Ağırlık'].sum():,}".replace(",", ".")},
            {"SAĞLAMASI": "TÜVENAN KÖMÜR", "ARAÇ SAYISI": len(tuvenan_df), "TONAJ": f"{tuvenan_df['Net Ağırlık'].sum():,}".replace(",", ".")},
            {"SAĞLAMASI": "HİDRO-GEN A.Ş.", "ARAÇ SAYISI": len(hidrogen_df), "TONAJ": f"{hidrogen_df['Net Ağırlık'].sum():,}".replace(",", ".")},
            {"SAĞLAMASI": "TABAN-UÇUCU KÜL-UÇUCU KÜL T", "ARAÇ SAYISI": len(kul_df), "TONAJ": f"{kul_df['Net Ağırlık'].sum():,}".replace(",", ".")},
            {"SAĞLAMASI": "SOMA-BATI-KENTÇİM-NARETRA-TKS...", "ARAÇ SAYISI": len(cimento_maden_df), "TONAJ": f"{cimento_maden_df['Net Ağırlık'].sum():,}".replace(",", ".")},
            {"SAĞLAMASI": "HİDRO-GEN HURDA VS.", "ARAÇ SAYISI": len(hurda_df), "TONAJ": f"{hurda_df['Net Ağırlık'].sum():,}".replace(",", ".")}
        ]
        saglama_df = pd.DataFrame(saglama_listesi)

        # --- EKRANDA GÖSTERİM (2 KOLON) ---
        sol_kolon, sag_kolon = st.columns([6, 5])
        
        with sol_kolon:
            st.subheader("📋 Vardiya Özet Taslağı")
            st.dataframe(rapor_df, use_container_width=True, hide_index=True)
            st.dataframe(toplam_df, use_container_width=True, hide_index=True)
            
            st.subheader("✅ Vardiya Sağlaması")
            st.dataframe(saglama_df, use_container_width=True, hide_index=True)
            
        with sag_kolon:
            st.subheader("🔍 Yüklenen Orijinal Kantar Listesi")
            st.dataframe(df[['Plaka', 'Firma', 'Malzeme', 'Nereden Geldi', 'Net Ağırlık']], use_container_width=True, hide_index=True)
            
    except Exception as e:
        st.error(f"Excel okunurken bir hata oluştu: {e}")
else:
    st.info("💡 Lütfen yukarıdaki alana vardiya kantar Excel dosyasını yükleyin.")
