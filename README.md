# GMapsX

**Google Maps verilerini hızlı ve kolay bir şekilde dışa aktarın.**

![GMapsX](maps.png)

GMapsX, Google Maps'te arama yaparak işletmelerin ad, kategori, puan, yorum sayısı, adres, telefon, website ve Google Maps URL bilgilerini toplayıp Excel, PDF veya TXT olarak dışa aktaran masaüstü uygulamasıdır.

> ⚠️ **Sorumluluk Reddi:** Bu araç yalnızca eğitim ve araştırma amaçlıdır. Kullanımı tamamen kullanıcının sorumluluğundadır. Google Maps kullanım şartlarını inceleyiniz. Ticari amaçla veya Google'ın hizmet şartlarını ihlal edecek şekilde kullanılması önerilmez.

---

## Özellikler

- 🔍 **Google Maps arama** — işletme türü ve konuma göre arama
- ⚡ **Çoklu thread desteği** — 1-20 arası eş zamanlı thread ile hızlı veri toplama
- 🌍 **27 dil desteği** — Türkçe, İngilizce, Almanca, Fransızca, Arapça, Çince, Japonca ve daha fazlası
- 📊 **Excel export** — biçimlendirilmiş tablo + özet sayfası
- 📄 **PDF export** — taşınabilir belge formatında rapor
- 📝 **TXT export** — TSV formatında düz metin çıktısı
- 🕵️ **Headless mod** — arka planda gösterimsiz çalıştırma
- 🎯 **Maksimum sonuç kontrolü** — 5-200 arası sonuç sınırı
- 🖤 **Minimal koyu tema** — göz yormayan CustomTkinter arayüz

## Export Edilen Veriler

| Sütun | Açıklama |
|-------|----------|
| İsim | İşletme adı |
| Kategori | İşletme kategorisi (restoran, kafe, otel vb.) |
| Puan | Google puanı (1-5) |
| Yorum Sayısı | Toplam değerlendirme sayısı |
| Adres | Açık adres bilgisi |
| Telefon | İletişim numarası |
| Website | Web sitesi URL |
| Google Maps URL | Google Maps profil bağlantısı |

## Gereksinimler

- **Python** 3.10+
- **Chrome** tarayıcı (güncel sürüm)
- İnternet bağlantısı

## Kurulum

```bash
# 1. Repoyu klonlayın
git clone https://github.com/cangurel81/GMapsX.git
cd GMapsX

# 2. Bağımlılıkları yükleyin
pip install -r requirements.txt

# 3. Çalıştırın
python main.py
```

Exe olarak derlemek için `auto-py-to-exe` kullanabilir, `maps.ico` dosyasını "Additional Files" kısmına eklemeyi unutmayabilirsiniz.

## Kullanım

1. **İşletme türü** — örn. `restoran`, `avukat`, `dişçi`, `kuaför`
2. **Konum** — örn. `İstanbul`, `Ankara`, `Berlin`
3. **Maksimum sonuç** — toplanacak maksimum işletme sayısı
4. **Thread sayısı** — eş zamanlı çalışacak thread adedi
5. **Dil** — Google Maps arayüz dili
6. **▶ Aramayı Başlat** ile scraping başlar
7. İşlem tamamlandığında Excel / PDF / TXT olarak dışa aktarın

## Mimari

```
main.py          → CustomTkinter UI, export işlemleri
scraper.py       → Selenium tabanlı scraping motoru
maps.ico / .png  → Uygulama ikonu
requirements.txt → Python bağımlılıkları
```

## Teknolojiler

- **Arayüz:** CustomTkinter (CTk)
- **Scraping:** Selenium + ChromeDriver
- **Export:** openpyxl (Excel), fpdf2 (PDF)
- **Thread Yönetimi:** Python threading
- **Tema:** Özel koyu gri renk paleti

## Geliştirmeye Açık

Katkıda bulunmak, hata bildirmek veya öneri paylaşmak için **Issues** ve **Pull Requests** açık. Projeyi forklayıp geliştirebilirsiniz.

---

**[WebAdHere Software](https://webadhere.com/)** &nbsp;·&nbsp; [GitHub @cangurel81](https://github.com/cangurel81)
