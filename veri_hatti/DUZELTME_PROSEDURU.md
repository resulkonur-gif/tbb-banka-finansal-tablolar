# PDF ek verisi — elle düzeltme prosedürü

Bu prosedür `ANA_KLASOR/4_pdf_ek_veri/ek_veri.xlsx` dosyası içindir. Bu dosya, TBB'de olmayan ve yalnız
banka raporlarının metninden okunan kalemleri tutar: serbest karşılık ile TÜFE'ye endeksli kıymetlerin
varsayımı ve duyarlılığı. Kalemlerin tanımı `taksonomi.json` dosyasında ve Excel'in **Taksonomi** sayfasındadır.

## İlke: eski değer silinmez, doğrusu yanına yazılır

Her satırda iki değer birlikte saklanır:

| Sütun | Kim doldurur | Ne anlama gelir |
|---|---|---|
| **Okunan değer** | Program | PDF'ten okunan değer. Program bu sütunu yalnız ilk okumada yazar ve sonra değiştirmez. |
| **Güncel değer (elle)** *(sarı)* | Siz | Doğru olması gereken değer. Boş bırakılırsa okunan değer geçerli olur. |
| **Açıklama (elle)** *(sarı)* | Siz | Düzeltmenin nedeni ve kaynağı, ör. "s.98 elle kontrol; tablo 2 satır kaymış". |
| **Geçerli değer** | Formül | Güncel değer doluysa o, boşsa okunan değer. Site ve değerlendirme yalnız bu değeri kullanır. |
| **Durum** | Formül | Üç durumdan biri: *okundu*, *okunamadı — elle girin* ya da *elle düzeltildi*. |

Program sarı sütunlara hiçbir zaman yazmaz. Okunan değer de ilk okumadan sonra değişmez. Böylece
düzeltilen her satırda ilk okunan değer ile sizin girdiğiniz doğru değer yan yana durur.

## Adımlar

1. Hattı bir kez çalıştırın (`python veri_hatti.py ...`). Excel ilk çalıştırmada oluşur, sonraki
   çalıştırmalarda yalnız yeni dönemlerin satırları eklenir.
2. `ek_veri.xlsx` dosyasını açın. **Durum** sütununda şu satırları süzün:
   - *okunamadı — elle girin*: cümle bulunamamıştır. Raporda varsa değeri girin. Kalem raporda hiç
     yoksa, örneğin serbest karşılık bulunmuyorsa, değer olarak `0` yazın ve açıklamaya "raporda yok"
     ekleyin.
   - *okundu*: **Okunan cümle** sütunu ile **Kaynak** sütunundaki sayfayı karşılaştırarak kontrol edin.
3. Yanlış ya da eksik değer için yalnız **Güncel değer (elle)** sütununu doldurun. **Açıklama (elle)**
   sütununa nedenini yazın. Okunan değere dokunmayın.
4. Birimler: tutarlar **milyon TL**, oranlar **%** (29,5 → `29.5`). Azalış duyarlılığı pozitif yazılır.
5. Kaydedip kapatın. Dosya açıkken program yazamaz.
6. Düzeltmenin siteye yansıması için hattı yeniden çalıştırın. Ek veri adımı yalnız şu iki dosyayı
   yeniden yazar:

   ```
   python veri_hatti.py --kok ANA_KLASOR --depo DEPO --sunum-atla
   ```

   Ya da sadece depoyu güncelleyin:

   ```
   python pdf_ek_veri.py --kok ANA_KLASOR
   ```
7. Program, bir önceki çalıştırmadan beri değişen her elle değeri **Düzeltme günlüğü** sayfasına yazar.
   Her kayıtta zaman, kalem, eski ve yeni değer ile açıklama bulunur. Hangi değerin ne zaman neden
   değiştiğini oradan izleyebilirsiniz.

## Elle girişi geri almak

**Güncel değer (elle)** hücresini boşaltın. Geçerli değer yeniden okunan değer olur. Bu geri alma da
günlüğe yazılır.

## PDF'i yeniden okutmak

Okuma kuralı iyileştirildiğinde PDF'leri yeniden okutmak için:

```
python pdf_ek_veri.py --kok ANA_KLASOR --yeniden-oku
```

ya da `veri_hatti.py ... --ek-veri-yeniden` çalıştırın. Okunan değerler yenilenir, sarı sütunlar
korunur. Okunan değeri değişen satırlar günlüğe "yeniden okuma" olarak düşer.

## Yeni kalem eklemek (taksonomi)

1. `pdf_ek_veri.py` içindeki `TAKSONOMI["kalemler"]` listesine yeni kalemi ekleyin: kod, ad, birim,
   grup, kaynak ve not.
2. `oku()` fonksiyonuna o kalemin cümlesini bulan kuralı yazın.
3. Eski dönemler için `--yeniden-oku` ile çalıştırın. Yeni kalemin satırları eklenir, mevcut elle
   düzeltmeler korunur.

## Bilinen durumlar (Eylül 2026 raporları)

- **Halkbank** TÜFE kıymetlerinde tek oran değil enflasyon eğrisi kullanır. Duyarlılığı vergi öncesi
  kâr için değil **faiz geliri** için verir; Kaynak sütununda "faiz geliri" yazar. Varsayım oranını
  isterseniz sunumdan ya da analist notundan elle girin.
- **Garanti, Akbank ve Yapı Kredi** duyarlılığı tek tutar olarak verir; program artış ile azalışı
  simetrik kabul eder.
- Serbest karşılık cümlesi şimdilik yalnız **VakıfBank** ve **Akbank** raporlarında bulundu. Diğer
  bankalarda satır *okunamadı* olarak gelir. Raporda serbest karşılık yoksa `0` girin.
