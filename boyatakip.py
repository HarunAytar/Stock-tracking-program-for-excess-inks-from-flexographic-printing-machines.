import pandas as pd
import numpy as np
import os
import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk  


ctk.set_appearance_mode("Dark")  
ctk.set_default_color_theme("blue")  


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DOSYA_YOLU = os.path.join(BASE_DIR, "Bilgi Tablosu.xlsx")

class BoyahaneUygulamasi(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        
        self.title("Boya Takip ve Renk Yönetimi Pro")
        self.geometry("1000x800")
        
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1) 

        
        self.font_title = ("Roboto Medium", 20)
        self.font_main = ("Roboto", 14)
        self.font_bold = ("Roboto", 14, "bold")
        
        self.df = None
        self.veriyi_yukle()
        self.arayuz_olustur()

    def veriyi_yukle(self):
        if os.path.exists(DOSYA_YOLU):
            try:
                self.df = pd.read_excel(DOSYA_YOLU)
                
                self.df.columns = [str(c).strip() for c in self.df.columns]
            except Exception as e:
                messagebox.showerror("Hata", f"Excel okunamadı. Dosya açık olabilir!\nHata: {e}")
        else:
            messagebox.showerror("Hata", f"Dosya bulunamadı: {DOSYA_YOLU}")

    def arayuz_olustur(self):
        
        lbl_baslik = ctk.CTkLabel(self, text="BOYAHANE YÖNETİM PANELİ", font=self.font_title, text_color="#3B8ED0")
        lbl_baslik.grid(row=0, column=0, pady=(20, 10), sticky="ew")

        
        self.frame_sorgu = ctk.CTkFrame(self)
        self.frame_sorgu.grid(row=1, column=0, padx=20, pady=10, sticky="nsw") 
        self.frame_sorgu.grid_columnconfigure((1, 3), weight=1)

        ctk.CTkLabel(self.frame_sorgu, text="İş Adı:", font=self.font_main).grid(row=0, column=0, padx=10, pady=15)
        self.ent_is = ctk.CTkEntry(self.frame_sorgu, placeholder_text="Ör: sleepy", font=self.font_main, width=200)
        self.ent_is.grid(row=0, column=1, padx=10, pady=15)

        ctk.CTkLabel(self.frame_sorgu, text="Boya İsmi:", font=self.font_main).grid(row=0, column=2, padx=10, pady=15)
        self.ent_boya = ctk.CTkEntry(self.frame_sorgu, placeholder_text="Ör: pan 185", font=self.font_main, width=200)
        self.ent_boya.grid(row=0, column=3, padx=10, pady=15)

        
        btn_ara = ctk.CTkButton(self.frame_sorgu, text="BOYAYI BUL & ANALİZ ET", 
                                command=self.sorgula, 
                                font=self.font_bold,
                                height=40,
                                fg_color="#1F6AA5", hover_color="#144870")
        btn_ara.grid(row=0, column=4, padx=20, pady=15)

        
        self.txt_sonuc = ctk.CTkTextbox(self, font=("Consolas", 14), activate_scrollbars=True)
        self.txt_sonuc.grid(row=2, column=0, padx=20, pady=10, sticky="nsew")
        self.txt_sonuc.insert("0.0", "...\n")

        
        self.frame_islem = ctk.CTkFrame(self)
        self.frame_islem.grid(row=3, column=0, padx=20, pady=20, sticky="ew")
        
        ctk.CTkLabel(self.frame_islem, text="STOK GÜNCELLEME", font=self.font_bold, text_color="gray").grid(row=0, column=0, columnspan=5, pady=(10,5))

        ctk.CTkLabel(self.frame_islem, text="Yeni Eklenen (kg):", font=self.font_main).grid(row=1, column=0, padx=10, pady=15)
        self.ent_eklenen = ctk.CTkEntry(self.frame_islem, font=self.font_main, width=120)
        self.ent_eklenen.insert(0, "0")
        self.ent_eklenen.grid(row=1, column=1, padx=10, pady=15)

        ctk.CTkLabel(self.frame_islem, text="Geri Dönen (kg):", font=self.font_main).grid(row=1, column=2, padx=10, pady=15)
        self.ent_donen = ctk.CTkEntry(self.frame_islem, font=self.font_main, width=120, placeholder_text="Kalan miktar")
        self.ent_donen.grid(row=1, column=3, padx=10, pady=15)

        btn_kaydet = ctk.CTkButton(self.frame_islem, text="VERİTABANINA KAYDET", 
                                   command=self.kaydet, 
                                   font=self.font_bold,
                                   height=40,
                                   fg_color="#2CC985", hover_color="#228B5E", text_color="white") # Yeşil buton
        btn_kaydet.grid(row=1, column=4, padx=20, pady=15)

    def sorgula(self):
        is_adi = self.ent_is.get().strip()
        boya_adi = self.ent_boya.get().strip()
        
        self.txt_sonuc.delete('1.0', tk.END)
        
        if self.df is None:
            self.txt_sonuc.insert(tk.END, "Veritabanı yüklenemedi!")
            return

        hedef = self.df[(self.df['Is_Adi'] == is_adi) & (self.df['Boya_Ismi'] == boya_adi)]
        
        if hedef.empty:
            self.txt_sonuc.insert(tk.END, f"HATA: '{is_adi}' işinde '{boya_adi}' bulunamadı!\nLütfen yazımı kontrol ediniz.")
            return

        
        h_L = hedef.iloc[0]['L']
        h_a = hedef.iloc[0]['a']
        h_b = hedef.iloc[0]['b']
        
        
        ayrac = "="*60 + "\n"
        alt_ayrac = "-"*60 + "\n"
        
        metin = f"{ayrac} ARANAN BOYA BİLGİLERİ \n{ayrac}"
        metin += f"• İş / Boya    : {is_adi} / {boya_adi}\n"
        metin += f"• Aniloks      : {hedef.iloc[0]['Aniloks']}\n"
        metin += f"• Film Rengi   : {hedef.iloc[0]['Film_Rengi']}\n"
        metin += f"• Lab Değerleri: L={h_L}, a={h_a}, b={h_b}\n"
        metin += f"• Mevcut Stok  : {hedef.iloc[0]['Miktar']} kg\n"
        metin += f"• KONUM        : RAF {hedef.iloc[0]['Raf']} | KAT {hedef.iloc[0]['Kat']} | SIRA {hedef.iloc[0]['Sira']}\n\n"

        
        temp_df = self.df.copy()
        for c in ['L', 'a', 'b']: 
            temp_df[c] = pd.to_numeric(temp_df[c], errors='coerce').fillna(0)
        
        temp_df['DeltaE'] = np.sqrt((temp_df['L'] - h_L)**2 + (temp_df['a'] - h_a)**2 + (temp_df['b'] - h_b)**2)
        
        
        oneriler = temp_df[temp_df.index != hedef.index[0]].sort_values(by='DeltaE').head(3)

        metin += f"{ayrac} ALTERNATİF / ÇIKMA BOYA ÖNERİLERİ (En Yakın Renkler) \n{ayrac}"
        
        for i, (_, row) in enumerate(oneriler.iterrows(), 1):
            metin += f"#{i} [Delta E: {row['DeltaE']:.2f}] -> {row['Boya_Ismi']} ({row['Is_Adi']})\n"
            metin += f"   Konum : {row['Raf']}-{row['Kat']}-{row['Sira']} | Stok: {row['Miktar']} kg\n"
            metin += f"   Detay : Aniloks {row['Aniloks']} | Film {row['Film_Rengi']}\n"
            metin += alt_ayrac

        self.txt_sonuc.insert(tk.END, metin)

    def kaydet(self):
        is_adi = self.ent_is.get().strip()
        boya_adi = self.ent_boya.get().strip()
        
        try:
            eklenen = float(self.ent_eklenen.get() or 0)
            if not self.ent_donen.get():
                messagebox.showwarning("Uyarı", "Lütfen geri dönen miktarı giriniz.")
                return
            donen = float(self.ent_donen.get())
        except ValueError:
            messagebox.showerror("Hata", "Lütfen ağırlık alanlarına sadece SAYI giriniz (kg).")
            return

        mask = (self.df['Is_Adi'] == is_adi) & (self.df['Boya_Ismi'] == boya_adi)
        if not self.df[mask].empty:
            eski_miktar = self.df.loc[mask, 'Miktar'].values[0]
            toplam_verilen = eski_miktar + eklenen
            tuketim = toplam_verilen - donen
            
            
            self.df.loc[mask, 'Miktar'] = donen
            
            try:
                self.df.to_excel(DOSYA_YOLU, index=False)
                
                
                ozet = f"İşlem Başarılı!\n\n" \
                       f"Eski Stok: {eski_miktar:.2f} kg\n" \
                       f"Eklenen: {eklenen:.2f} kg\n" \
                       f"Harcanan: {tuketim:.2f} kg\n" \
                       f"-------------------\n" \
                       f"YENİ STOK: {donen:.2f} kg"
                
                messagebox.showinfo("Kayıt Başarılı", ozet)
                
                
                self.sorgula()
                
            except Exception as e:
                messagebox.showerror("Hata", f"Excel dosyası kaydedilemedi.\nDosya açık olabilir, lütfen kapatıp tekrar deneyin.\nHata: {e}")
        else:
            messagebox.showerror("Hata", "Güncellenecek iş/boya bulunamadı. Önce sorgulama yapın.")

if __name__ == "__main__":
    app = BoyahaneUygulamasi()
    app.mainloop()