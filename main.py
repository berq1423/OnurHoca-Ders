import pdfplumber
import re
from collections import Counter
from textblob import TextBlob
from deep_translator import GoogleTranslator
import nltk
from nltk.corpus import stopwords
from PIL import Image
import numpy as np
from sklearn.cluster import KMeans
import matplotlib.pyplot as plt
import pandas as pd
import time
import os

# Gerekli dil paketlerini indir
nltk.download('punkt')
nltk.download('stopwords')

class RaporAnalizcisi:
    def __init__(self, pdf_path):
        # Ayarlar
        self.pdf_path = pdf_path
        self.full_text = ""       
        self.cumle_verileri = []  # Duygu analizi verileri
        self.en_sik_kelimeler = [] # Kelime sıklığı verileri
        self.sayfa_renkleri = []  
        
        # --- GÜNCELLENEN KISIM: ÇIKTI KLASÖRÜ ---
        self.output_folder = "Çıktı"
        os.makedirs(self.output_folder, exist_ok=True)
        print(f"📁 Sonuçlar '{self.output_folder}' klasörüne kaydedilecek.")

    def metin_ve_duygu_analizi(self):
        print("\n--- 📖 CÜMLELER OKUNUYOR VE DUYGU ANALİZİ YAPILIYOR ---")
        translator = GoogleTranslator(source='auto', target='en')
        
        with pdfplumber.open(self.pdf_path) as pdf:
            total_pages = len(pdf.pages)
            tum_metin_listesi = []

            for i, page in enumerate(pdf.pages):
                text = page.extract_text()
                if not text: continue
                
                tum_metin_listesi.append(text)
                
                # Cümlelere böl
                cumleler = text.split('.')
                print(f"Sayfa {i+1}/{total_pages} işleniyor...")
                
                for cumle in cumleler:
                    temiz_cumle = cumle.strip()
                    # Başlıkları ve çok kısa anlamsız cümleleri atla
                    if len(temiz_cumle) < 20: continue 
                    
                    try:
                        # 1. Çevir -> 2. Analiz Et
                        ceviri = translator.translate(temiz_cumle)
                        puan = TextBlob(ceviri).sentiment.polarity
                        
                        durum = "Nötr"
                        if puan > 0.1: durum = "Pozitif"
                        elif puan < -0.1: durum = "Negatif"

                        self.cumle_verileri.append({
                            "Sayfa": i+1,
                            "Cümle": temiz_cumle,
                            "Duygu Puanı": puan,
                            "Durum": durum
                        })
                        # Google Translate'i yormamak için kısa bekleme
                        time.sleep(0.05) 
                    except:
                        continue
            
            self.full_text = " ".join(tum_metin_listesi)

    def kelime_sikligi_analizi(self):
        print("\n--- 🔢 EN SIK GEÇEN KELİMELER SAYILIYOR ---")
        
        # 1. Metni temizle (Küçük harf yap, noktalama sil)
        text = self.full_text.lower()
        text = re.sub(r'[^\w\s]', '', text) # Sadece harf ve boşluk kalsın
        words = text.split()
        
        # 2. Gereksiz kelimeleri (Stopwords) belirle
        etkisiz_kelimeler = set(stopwords.words('turkish'))
        # Listeye manuel eklemeler yapıyoruz (bunlar analizde çıkmasın)
        ekstra_etkisizler = {"bir", "ve", "ile", "bu", "de", "da", "için", "olarak", "olan", "daha", "veya", "gibi", "kadar", "sonra", "ancak", "yılında", "tarafından"}
        etkisiz_kelimeler.update(ekstra_etkisizler)
        
        # 3. Temizlenmiş kelime listesi oluştur
        anlamli_kelimeler = [w for w in words if w not in etkisiz_kelimeler and len(w) > 2]
        
        # 4. Sayım yap
        sayac = Counter(anlamli_kelimeler)
        
        # En çok geçen 100 kelimeyi al
        self.en_sik_kelimeler = sayac.most_common(100)
        print(f"✅ Toplam {len(anlamli_kelimeler)} anlamlı kelime tarandı. İlk 100 çıkarıldı.")

    def renk_analizi(self):
        print("\n--- 🎨 GÖRSEL TASARIM ANALİZİ ---")
        with pdfplumber.open(self.pdf_path) as pdf:
            for i, page in enumerate(pdf.pages):
                best_color = [255, 255, 255]
                max_area = 0
                
                for img in page.images:
                    try:
                        x0, top, x1, bottom = img['x0'], img['top'], img['x1'], img['bottom']
                        area = (x1 - x0) * (bottom - top)
                        # Sadece büyük resimleri (grafik/foto) al
                        if area > 5000 and area > max_area:
                            cropped = page.crop((x0, top, x1, bottom)).to_image(resolution=72)
                            img_np = np.array(cropped.original.resize((50,50)))
                            if img_np.shape[-1] == 4: img_np = img_np[:,:,:3]
                            pixels = img_np.reshape(-1, 3)
                            kmeans = KMeans(n_clusters=1, n_init=5).fit(pixels)
                            best_color = kmeans.cluster_centers_[0].astype(int)
                            max_area = area
                    except:
                        pass
                self.sayfa_renkleri.append(best_color)

    def dosyalara_kaydet(self):
        print("\n--- 💾 DOSYALAR KAYDEDİLİYOR ---")
        
        excel_path = f"{self.output_folder}/Analiz_Raporu.xlsx"
        
        # 1. EXCEL OLUŞTURMA
        df_duygu = pd.DataFrame(self.cumle_verileri)
        df_kelimeler = pd.DataFrame(self.en_sik_kelimeler, columns=["Kelime", "Tekrar Sayısı"])
        
        with pd.ExcelWriter(excel_path) as writer:
            df_duygu.to_excel(writer, sheet_name='Duygu Analizi', index=False)
            df_kelimeler.to_excel(writer, sheet_name='En Sık Geçen Kelimeler', index=False)
        
        print(f"✅ Excel dosyası hazır: {excel_path}")

        # 2. GRAFİKLER
        # Duygu Grafiği
        if not df_duygu.empty:
            plt.figure(figsize=(12, 6))
            plt.plot(df_duygu.index, df_duygu['Duygu Puanı'], alpha=0.3, color='gray')
            plt.plot(df_duygu.index, df_duygu['Duygu Puanı'].rolling(window=5).mean(), color='blue', linewidth=2, label='Trend')
            plt.axhline(0, color='red', linestyle='--')
            plt.title("Raporun Duygu Grafiği")
            plt.ylabel("Duygu (Pozitif/Negatif)")
            plt.xlabel("Cümle Sırası")
            plt.legend()
            plt.savefig(f"{self.output_folder}/Duygu_Grafigi.png")
            plt.close()

        # Renk Grafiği
        if self.sayfa_renkleri:
            colors = np.array(self.sayfa_renkleri)
            plt.figure(figsize=(12, 2))
            plt.imshow([colors], aspect='auto')
            plt.axis('off')
            plt.title("Sayfa Bazlı Renk Haritası")
            plt.savefig(f"{self.output_folder}/Renk_Haritasi.png")
            plt.close()
            
        print("🎉 İŞLEM TAMAMLANDI! 'Çıktı' klasörünü kontrol edebilirsin.")

# --- ÇALIŞTIRMA ---
# GÜNCELLENEN KISIM: Senin PDF dosyanın tam adı
dosya = "2024-tsrs-uyumlu-surdurulebilirlik-raporu.pdf" 

try:
    analiz = RaporAnalizcisi(dosya)
    analiz.metin_ve_duygu_analizi()
    analiz.kelime_sikligi_analizi() 
    analiz.renk_analizi()
    analiz.dosyalara_kaydet()
except Exception as e:
    print(f"Hata oluştu: {e}")