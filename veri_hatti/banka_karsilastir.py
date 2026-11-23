#!/usr/bin/env python3
"""
KAP'tan büyük bankaların finansal raporlarını indirir, rapordaki ana tabloları
okur ve TBB Veri Sistemi verisiyle karşılaştıran bir Excel dosyası üretir.

Tablolar: Aktif (Varlıklar), Pasif (Yükümlülükler), Nazım Hesaplar,
          Gelir-Gider (Kar veya Zarar Tablosu), Kâr veya Zarar ve Diğer Kapsamlı Gelir,
          Nakit Akış, Özkaynak kalemleri (yasal özkaynak, sermaye yeterliliği oranları),
          Kur riski, Faiz oranı riski, Ortalama faiz oranları, Hisse senedi pozisyon riski,
          Likidite karşılama oranı (ve son üç ayın en düşük/en yüksek değeri),
          Kalan vadeye göre aktif-pasif, Net istikrarlı fonlama oranı, Kaldıraç oranı,
          TBB göstergeleri (YP aktif, YKLV, LKO, NİFO bileşenleri), Mevduat vade yapısı ve
          öteki bütün TBB dipnotları (genel eşleyici; kapsam ve bulunamayanlar ayrı sayfada)

Adımlar
  1. Hangi bankaları ve hangi dönemi istediğinizi sorar (7 büyük banka).
  2. Her banka için KAP'taki finansal tablo paketini indirir.
  3. Paketteki PDF'lerden Türkçe solo ve konsolide raporları ayırır.
  4. Raporlardan tabloları okur.
  5. TBB verisini GitHub'daki raporunuzdan alır, kalem kalem eşler ve
     farkları gösteren Excel'i üretir: her tablo ayrı sayfada (solo ve
     konsolide birlikte), iş kuralları ve kalem eşleşmeleri ayrı sayfada.

Kurulum (bir kez):  python -m pip install pymupdf openpyxl
Çalıştırma:         python banka_karsilastir.py
"""
from __future__ import annotations

import difflib
import io
import json
import re
import subprocess
import sys
import time
import unicodedata
import urllib.error
import urllib.request
import zipfile
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

# ============================================================== ayarlar

BANKALAR = [
    # no, ad, kısa ad, borsa kodu, TBB banka kodu, KAP üye kimliği (OID)
    (1, "T.C. Ziraat Bankası", "Ziraat", None, 3, None),
    (2, "Türkiye Halk Bankası", "Halkbank", "HALKB", 5, "1DE05DAA82C3073AE0530A4A622A2EBD"),
    (3, "Türkiye Vakıflar Bankası", "VakıfBank", "VAKBN", 6, "4028e4a1415f4d9b01415fe3340f36e4"),
    (4, "Türkiye İş Bankası", "İş Bankası", "ISCTR", 32, "4028e4a140f2ed7201411682b0cb05c6"),
    (5, "Türkiye Garanti Bankası", "Garanti BBVA", "GARAN", 30, "4028e4a140f2ed720140f37cb2a601b7"),
    (6, "Akbank", "Akbank", "AKBNK", 8, "4028e4a240e8d1830140e905edcd0006"),
    (7, "Yapı ve Kredi Bankası", "Yapı Kredi", "YKBNK", 37, "4028e4a240f2ef4c01412ae6d6630538"),
]

KAP_API = "https://www.kap.org.tr/tr/api"
TBB_URL = "https://raw.githubusercontent.com/resulkonur-gif/tbb-banka-finansal-tablolar/main/docs/index.html"
KLASOR = Path(__file__).resolve().parent
AY_ADI = {3: "Mart", 6: "Haziran", 9: "Eylül", 12: "Aralık"}

TABLOLAR = [   # PDF bölüm anahtarı → TBB tablo adı, Excel'de görünen ad
    ("VARLIKLAR", "1. Varlıklar", "Aktif"),
    ("YUKUMLULUKLER", "2. Yükümlülükler", "Pasif"),
    ("NAZIM", "3. Nazım Hesaplar", "Nazım"),
    ("KZ", "4. Gelir-Gider Tablosu", "Gelir-Gider"),
    ("DKG", "5. Kar veya Zarar ve DKG", "Kâr-Zarar ve DKG"),
    ("NAKIT", "7. Nakit Akış Tablosu", "Nakit Akış"),
    ("OZK", "9. Yasal Özkaynak", "Özkaynak"),
]

# Özkaynak tablosundaki oran satırları: TBB'de tablo kalemi değil, ayrı gösterge
OZK_ORANLAR = [   # sade() ile Türkçe karakterlerden arındırılmış ada göre
    (r"CEKIRDEK SERMAYE YETERLILIGI ORANI", "cet1"),
    (r"^(KONSOLIDE )?ANA SERMAYE YETERLILIGI ORANI", "t1"),
    (r"^(KONSOLIDE )?SERMAYE YETERLILIGI ORANI", "car"),
    (r"TOPLAM ILAVE CEKIRDEK SERMAYE GEREKSINIMI", "buf"),
    (r"SERMAYE KORUMA TAMPONU", "ccb"),
    (r"DONGUSEL SERMAYE TAMPONU", "ccyb"),
    (r"SISTEMIK ONEMLI BANKA TAMPONU", "sib"),
    (r"RISK AGIRLIKLI VARLIKLAR TUTARINA ORANI", "bufR"),
]
OZK_ARA_BASLIK = (r"(KATKI SERMAYE HESAPLAMASINDA DIKKATE ALINAN KARSILIKLARA ILISKIN SINIRLAR"
                  r"|UYGULANACAK INDIRIM ESASLARINDA ASIM TUTARININ ALTINDA KALAN TUTARLAR"
                  r"|SERMAYE YETERLILIGI ORANLARI|TAMPONLAR|ILAVE ANA SERMAYE|KATKI SERMAYE|OZKAYNAK)")
OZK_SON = "GECICI MADDE 4 HUKUMLERINE TABI KATKI SERMAYE KALEMLERININ UST SINIRI ASAN"

# TBB'de karşılığı olmadığı bilinen PDF satırları: aranmaz, "beklenen" olarak işaretlenir.
# (tablo anahtarı ya da None = hepsi, rapor türü ya da None = ikisi, sade() ada uygulanan desen, açıklama)
BEKLENEN = [
    ("NAKIT", None, r"^(BANKACILIK|YATIRIM|FINANSMAN) FAALIYETLERINE ILISKIN NAKIT AKI(SLARI|MLARI)",
     "Nakit akış bölüm başlıkları (A, B, C) TBB'de yok"),
    (None, "solo", r"^AZINLIK PAYLARI|^GRUBUN KARI", "Konsolidasyona özgü kalemler (azınlık payları, grubun kârı) "
                                                   "solo raporda boş; TBB solo listesinde yok"),
    ("OZK", "kons", r"UCUNCU KISILERIN ILAVE ANA SERMAYE",
     "Konsolide özkaynakta 'üçüncü kişilerin ilave ana sermayedeki payları' TBB tablosunda yok"),
    (None, None, r"^BILANCO$|TABLOSU$", "Tablo başlığı satırı"),
]


def beklenen_mi(anahtar: str, tur: str, ad: str):
    for t_, r_, desen, aciklama in BEKLENEN:
        if (t_ is None or t_ == anahtar) and (r_ is None or r_ == tur) and re.search(desen, sade(ad)):
            return aciklama
    return None

# ============================================================== dipnot tabloları (mali bünye)
# Her tablo: PDF'te bulma koşulu, sütunlar, satır eşleme yöntemi ve TBB dipnot adı.
#   esle: "kanon"  → kalem adından anahtar (nakit, krediler, …) ve sütun adından alt kalem
#         "lko"    → satır no (1–23) → TBB LKO kalemi; sütunlar TP+YP / YP
#         "nifo"   → satır no (1–34) → TBB NİFO kalemi; a–ç kalan vade, d uygulanmış tutar
#         "sira"   → satır no → TBB şablonundaki aynı sıradaki kalem (kaldıraç)
#         None     → TBB'de bu tablo yok, yalnız PDF değeri yazılır
DIPNOT_TABLOLARI = [
    dict(anahtar="FAIZ", sayfa="Faiz Oranı Riski", tbb="Faiz Oranı Riski", esle="kanon",
         sutun=["1 Aya Kadar", "1-3 Ay", "3-12 Ay", "1-5 Yıl", "5 Yıl ve Üzeri", "Faizsiz", "Toplam"],
         kanon=["1ay", "1-3", "3-12", "1-5", "5+", "faizsiz", "toplam"],
         bul=lambda s: ("FAIZ ORANI RISK" in s or "FAIZE DUYARLILI" in s) and "FAIZSIZ" in s, xmin=0.30, sayfa_say=3,
         son=lambda r: r["donem"] == "onceki" and kanon_kalem(r["ad"]) == "top_poz"),
    dict(anahtar="ORTFAIZ", sayfa="Ortalama Faiz Oranları", esle="kanon",
         tbb="Faiz Oranı Riski-Parasal Fin.Araç.Uyg.Ort.Faiz Oranları (%)",
         sutun=["Avro", "ABD Doları", "Yen", "TL"], kanon=["eur", "usd", "yen", "tl"],
         bul=lambda s: "ORTALAMA FAIZ ORAN" in s and re.search(r"\b(AVRO|EURO|EUR)\b.{0,40}\b(USD|DOLARI)\b", s),
         xmin=0.40, sayfa_say=2, oran=lambda no, ad: True,
         basliklar=[r"^(AVRO|EURO|EUR)$", r"^(USD|DOLARI)$", r"^(YEN|JPY)$", r"^(TL|TRY)$"]),
    dict(anahtar="HISSE", sayfa="Hisse Senedi Riski", tbb=None, esle=None,
         sutun=["Bilanço Değeri", "Gerçeğe Uygun Değer", "Piyasa Değeri"],
         bul=lambda s: "HISSE SENEDI POZISYON" in s and "BILANCO DEGERI" in s and "PIYASA DEGERI" in s,
         xmin=0.40, sayfa_say=2, son=lambda r: r["donem"] == "onceki" and "GORMEYEN" in sade(r["ad"])),
    dict(anahtar="LKO", sayfa="Likidite Karşılama Oranı", tbb="Likidite Karşılama Oranı", esle="lko",
         numarali=True, sutun=["Uygulanmamış TP+YP", "Uygulanmamış YP", "Uygulanmış TP+YP", "Uygulanmış YP"],
         bul=lambda s: "LIKIDITE KARSILAMA ORAN" in s and "UYGULANMAMIS" in s, xmin=0.45, sayfa_say=4,
         ilk=r"^YUKSEK KALITELI LIKIT VARLIK", desen="LKO_DESEN",
         son=lambda r: r["donem"] == "onceki" and r["no"] == 23, son_no=23, oran=lambda no, ad: no == 23),
    dict(anahtar="LKO_UC", sayfa="LKO En Düşük-En Yüksek", tbb=None, esle=None,
         sutun=["TP+YP oran (%)", "YP oran (%)"], oran=lambda no, ad: True),
    dict(anahtar="VADE", sayfa="Kalan Vade", tbb="Aktif ve pasiflerin kalan vadelerine göre gösterimi", esle="kanon",
         sutun=["Vadesiz", "1 Aya Kadar", "1-3 Ay", "3-12 Ay", "1-5 Yıl", "5 Yıl ve Üzeri", "Dağıtılamayan", "Toplam"],
         kanon=["vadesiz", "1ay", "1-3", "3-12", "1-5", "5+", "dagit", "toplam"],
         bul=lambda s: "VADESIZ" in s and "DAGITILAMAYAN" in s and "VARLIK" in s, xmin=0.30, sayfa_say=3,
         son=lambda r: r["donem"] == "onceki" and kanon_kalem(r["ad"]) == "gnakdi"),
    dict(anahtar="NIFO", sayfa="NİFO", tbb="Net İstikrarlı Fonlama Oranı", esle="nifo", numarali=True,
         sutun=["a) Vadesiz", "b) 6 aydan kısa", "c) 6 ay – 1 yıl", "ç) 1 yıl ve üzeri", "d) Uygulanmış toplam"],
         bul=lambda s: "MEVCUT ISTIKRARLI FON" in s and ("CARI DONEM" in s or "ONCEKI DONEM" in s),
         ilk=r"^MEVCUT ISTIKRARLI FON$", desen="NIFO_DESEN", bos_satir=r"^BIRBIRLERINE BAGLI",
         xmin=0.45, sayfa_say=5, son=lambda r: r["donem"] == "onceki" and r["no"] == 34, son_no=34,
         oran=lambda no, ad: no == 34),
    dict(anahtar="KALDIRAC", sayfa="Kaldıraç Oranı", tbb="Kaldıraç oranı kamuya açıklama şablonu", esle="sira",
         numarali=True, tek_blok=True, sutun=["Tutar"], desen="KALDIRAC_DESEN", ilk=r"^BILANCO ICI VARLIK",
         bul=lambda s: "KALDIRAC ORANINA ILISKIN" in s and "BILANCO ICI VARLIK" in s and "TOPLAM RISK TUTARI" in s,
         xmin=0.55, sayfa_say=2, son=lambda r: r["no"] == 15, son_no=15, oran=lambda no, ad: no == 15),
]

# Kalem adından anahtar: faiz oranı riski, ortalama faiz ve kalan vade tablolarında PDF ve TBB adları farklı
# yazılır ("İtfa Edilmiş Maliyeti ile Ölçülen FV" / "İtfa Edilmiş Maliyeti Üzerinden Değerlenen FV").
KALEM_ANAHTAR = [
    (r"GAYRINAKDI", "gnakdi"), (r"TUREV.*ALACAK", "turev_al"), (r"TUREV.*BORC", "turev_borc"),
    (r"NET BILANCO DISI", "net_bd"), (r"LIKIDITE ACIGI", "net_lik"),
    (r"BILANCODAKI UZUN", "bil_uzun"), (r"BILANCODAKI KISA", "bil_kisa"),
    (r"NAZIM.*UZUN", "naz_uzun"), (r"NAZIM.*KISA", "naz_kisa"), (r"TOPLAM POZISYON", "top_poz"),
    (r"TOPLAM VARLIK|VARLIKLAR TOPLAMI", "top_var"), (r"TOPLAM YUKUM|YUKUMLULUKLER TOPLAMI", "top_yuk"),
    (r"NAKIT DEGER", "nakit"), (r"BANKALAR MEVDUAT", "bmev"),
    (r"PARA PIYASALARINDAN", "pp_al"), (r"PARA PIYASALARINA", "pp_borc"),
    (r"DIGER KAPSAMLI", "gud_dkg"), (r"GERCEGE UYGUN", "gud_kz"), (r"ITFA", "itfa"),
    (r"KREDI", "kredi"), (r"DIGER VARLIK", "dvar"), (r"MEVDUAT", "dmev"), (r"MUHTELIF", "muhtelif"),
    (r"IHRAC", "ihrac"), (r"MALI KURULUS|SAGLANAN FON", "dmali"), (r"DIGER YUKUM", "dyuk"),
    (r"^FONLAR$", "fonlar"), (r"^BANKALAR", "bankalar"),
]


def kanon_kalem(ad: str):
    s = sade(ad)
    for desen, anahtar in KALEM_ANAHTAR:
        if re.search(desen, s):
            return anahtar
    return None


def kanon_sutun(ad: str):
    s = sade(ad)
    for desen, k in ((r"^VADESIZ", "vadesiz"), (r"^1 AYA?\b", "1ay"), (r"^1 3 AY", "1-3"), (r"^3 12 AY", "3-12"),
                     (r"^1 5 YIL", "1-5"), (r"^5 YIL", "5+"), (r"^FAIZSIZ", "faizsiz"), (r"^DAGITILAMAYAN", "dagit"),
                     (r"^(EURO|AVRO)$", "eur"), (r"^(USD|ABD DOLARI)$", "usd"), (r"^YEN$", "yen"), (r"^TL$", "tl"),
                     (r"^TP YP$", "tpyp"), (r"^YP$", "yp")):
        if re.search(desen, s):
            return k
    return None


# LKO şablonunun satırları (BDDK şablonu; numaralar bütün bankalarda aynı) → TBB kalem adının başı
LKO_SATIR = {1: "YUKSEK KALITELI LIKIT VARLIKLAR", 2: "GERCEK KISI MEVDUAT VE PERAKENDE MEVDUAT",
             3: "ISTIKRARLI MEVDUAT", 4: "DUSUK ISTIKRARLI", 5: "GERCEK KISI MEVDUAT VE PERAKENDE MEVDUAT DISIND",
             6: "OPERASYONEL MEVDUAT", 7: "OPERASYONEL OLMAYAN", 8: "DIGER TEMINATSIZ", 9: "TEMINATLI BORCLAR",
             10: "DIGER NAKIT CIKIS", 11: "TUREV YUKUMLULUKLER", 12: "YAPILANDIRILMIS", 13: "FINANSAL PIYASALARA",
             14: "HERHANGI BIR SARTA", 15: "DIGER CAYILAMAZ", 16: "TOPLAM NAKIT CIKISLARI", 17: "TEMINATLI ALACAKLAR",
             18: "TEMINATSIZ ALACAKLAR", 19: "DIGER NAKIT GIRIS", 20: "TOPLAM NAKIT GIRISLERI",
             21: "TOPLAM YUKSEK KALITELI", 22: "TOPLAM NET NAKIT", 23: "LIKIDITE KARSILAMA ORANI"}
# NİFO şablonunun satırları → TBB kalem adının başı (10 ve 25 TBB'de yok)
NIFO_SATIR = {1: "OZKAYNAK UNSURLARI", 2: "ANA SERMAYE VE KATKI", 3: "DIGER OZKAYNAK UNSURLARI",
              4: "GERCEK KISI VE PERAKENDE", 5: "ISTIKRARLI MEVDUAT", 6: "DUSUK ISTIKRARLI", 7: "DIGER KISILERE",
              8: "OPERASYONEL MEVDUAT", 9: "DIGER BORCLAR", 10: "BIRBIRLERINE BAGLI VARLIK", 11: "DIGER YUKUMLULUKLER",
              12: "TUREV YUKUMLULUKLER", 13: "YUKARIDA YER ALMAYAN DIGER OZKAYNAK", 14: "MEVCUT ISTIKRARLI FON",
              15: "YUKSEK KALITELI LIKIT", 16: "KREDI KUR", 17: "CANLI ALACAKLAR", 18: "TEMINATI BIRINCI",
              19: "KREDI K", 20: "KRE K", 21: "35 YA DA", 22: "IKAMET", 23: "35 YA DA", 24: "YUKSEK KALITELI LIK",
              25: "BIRBIRLERINE BAGLI YUKUM", 26: "DIGER VARLIKLAR", 27: "ALTIN", 28: "TUREV SOZLES",
              29: "TUREV VARLIKLAR", 30: "TUREV YUKUMLULUKLERIN", 31: "YUKARIDA YER ALMAYAN DIGER VARLIK",
              32: "BILANCO DISI BORCLAR", 33: "GEREKLI ISTIKRARLI FON", 34: "NET ISTIKRARLI FONLAMA ORANI"}

# Satır numarası yazmayan raporlar (ör. Ziraat) için: PDF kalem adı → şablon satır no (sırayla aranır)
LKO_DESEN = {1: r"^YUKSEK KALITELI LIKIT VARLIK", 2: r"^GERCEK KISI MEVDUAT VE PERAKENDE MEVDUAT$",
             3: r"^ISTIKRARLI MEVDUAT", 4: r"^DUSUK ISTIKRARLI", 5: r"DISINDA KALAN", 6: r"^OPERASYONEL MEVDUAT",
             7: r"^OPERASYONEL OLMAYAN", 8: r"^DIGER TEMINATSIZ", 9: r"^TEMINATLI BORC", 10: r"^DIGER NAKIT CIKIS",
             11: r"^TUREV YUKUMLULUK", 12: r"^YAPILANDIRILMIS", 13: r"FINANSAL PIYASALARA", 14: r"HERHANGI BIR SARTA",
             15: r"CAYILAMAZ", 16: r"TOPLAM NAKIT CIKIS", 17: r"^TEMINATLI ALACAK", 18: r"^TEMINATSIZ ALACAK",
             19: r"^DIGER NAKIT GIRIS", 20: r"TOPLAM NAKIT GIRIS", 21: r"STOKU", 22: r"TOPLAM NET NAKIT",
             23: r"KARSILAMA ORANI"}
NIFO_DESEN = {1: r"^OZKAYNAK UNSURLARI", 2: r"^ANA SERMAYE", 3: r"^DIGER OZKAYNAK", 4: r"^GERCEK KISI VE PERAKENDE",
              5: r"^ISTIKRARLI MEVDUAT", 6: r"^DUSUK ISTIKRARLI", 7: r"^DIGER KISILERE", 8: r"^OPERASYONEL MEVDUAT",
              9: r"^DIGER BORCLAR", 10: r"^BIRBIRLERINE BAGLI VARLIK", 11: r"^DIGER YUKUMLULUKLER$",
              12: r"^TUREV YUKUMLULUKLER$", 13: r"^YUKARIDA YER ALMAYAN DIGER OZKAYNAK", 14: r"^MEVCUT ISTIKRARLI FON",
              15: r"^YUKSEK KALITELI LIKIT VARLIK", 16: r"DEPO EDILEN", 17: r"^CANLI ALACAK", 18: r"^TEMINATI BIRINCI",
              19: r"TEMINATSIZ VEYA", 20: r"DISINDAKI", 21: r"^35 YA DA", 22: r"^IKAMET", 23: r"^35 YA DA",
              24: r"NITELIGINI HAIZ", 25: r"^BIRBIRLERINE BAGLI YUKUM", 26: r"^DIGER VARLIKLAR$", 27: r"^ALTIN",
              28: r"BASLANGIC TEMINATI", 29: r"^TUREV VARLIKLAR", 30: r"DEGISIM TEMINATI",
              31: r"^YUKARIDA YER ALMAYAN DIGER VARLIK", 32: r"^BILANCO DISI BORC", 33: r"^GEREKLI ISTIKRARLI FON",
              34: r"FONLAMA ORANI"}
KALDIRAC_DESEN = {1: r"^BILANCO ICI VARLIKLAR\b", 2: r"ANA SERMAYEDEN INDIRILEN", 3: r"^BILANCO ICI VARLIKLARA",
                  4: r"YENILEME MALIYETI", 5: r"PO[TN]+[AO]NSIYEL", 6: r"TUREVLERINE ILISKIN TOPLAM",
                  7: r"BILANCO ICI HARIC|FINANSMAN ISLEMLERININ MENKUL", 8: r"ARACILIK",
                  9: r"ISLEMLERINE ILISKIN TOPLAM", 10: r"BRUT NOMINAL", 11: r"KREDIYE DONUSTURME",
                  12: r"BILANCO DISI ISLEMLERE ILISKIN TOPLAM", 13: r"^ANA SERMAYE$", 14: r"^TOPLAM RISK TUTARI",
                  15: r"KALDIRAC ORANI"}


def sade2(s: str) -> str:
    """sade() ama parantez içini atmadan ("(Ana sermayeden indirilen varlıklar)")."""
    return sade((s or "").replace("(", " ").replace(")", " "))


# ============================================================== kütüphaneler


def kutuphaneleri_hazirla():
    eksik = []
    try:
        import pymupdf  # noqa: F401
    except ImportError:
        try:
            import fitz  # noqa: F401
        except ImportError:
            eksik.append("pymupdf")
    try:
        import openpyxl  # noqa: F401
    except ImportError:
        eksik.append("openpyxl")
    if not eksik:
        return
    print(f"Gerekli kütüphaneler eksik: {', '.join(eksik)}")
    if input("Şimdi kurulsun mu? (E/h) [E]: ").strip().lower() in ("", "e", "evet", "y"):
        subprocess.check_call([sys.executable, "-m", "pip", "install", *eksik])
    else:
        sys.exit(f"Kurmak için: {sys.executable} -m pip install {' '.join(eksik)}")


try:   # Windows konsolunda Türkçe karakterler bozulmasın
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass
kutuphaneleri_hazirla()
try:
    import pymupdf
except ImportError:
    import fitz as pymupdf
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ============================================================== HTTP

BASLIKLAR = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/130.0 Safari/537.36"),
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "tr-TR,tr;q=0.9",
    "Referer": "https://www.kap.org.tr/tr/bildirim-sorgu",
    "Origin": "https://www.kap.org.tr",
}
ACICI = urllib.request.build_opener(urllib.request.ProxyHandler())   # sistem proxy ayarı
ARALIK = 2.5
_son = [0.0]


class IstekHatasi(Exception):
    pass


def istek(method: str, url: str, govde: dict | None = None, deneme: int = 4) -> bytes:
    veri = json.dumps(govde).encode() if govde is not None else None
    h = dict(BASLIKLAR)
    if veri is not None:
        h["Content-Type"] = "application/json"
    son_hata = None
    for i in range(deneme):
        bekle = ARALIK - (time.monotonic() - _son[0])
        if bekle > 0 and "kap.org.tr" in url:
            time.sleep(bekle)
        _son[0] = time.monotonic()
        try:
            req = urllib.request.Request(url, data=veri, headers=h, method=method)
            with ACICI.open(req, timeout=120) as r:
                return r.read()
        except urllib.error.HTTPError as e:
            son_hata = e
            if e.code == 429:
                s = 60 * (i + 1)
                print(f"      KAP istek sınırı (429); {s} sn bekleniyor…", flush=True)
                time.sleep(s)
                continue
            if e.code not in (500, 502, 503, 504):
                break
        except (urllib.error.URLError, TimeoutError, ConnectionError) as e:
            son_hata = e
        time.sleep(5 * (i + 1))
    raise IstekHatasi(f"{method} {url} → {son_hata}")


def js(b: bytes):
    try:
        return json.loads(b.decode("utf-8"))
    except Exception:
        return None


# ============================================================== metin yardımcıları

def sade(s: str) -> str:
    s = (s or "").upper().replace("İ", "I").replace("ı", "I")
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = re.sub(r"\([^)]*\)", " ", s)
    return " ".join(re.sub(r"[^A-Z0-9 ]", " ", s).split())


def benzer(a: str, b: str) -> float:
    a, b = sade(a), sade(b)
    return difflib.SequenceMatcher(None, a, b).ratio() if a and b else 0.0


DIPNOT_REF = re.compile(r"(\s*\(?\b[IVX]{1,4}-[IVX]{1,4}(?:-\d+(?:\.\d+)*)?\)?(?:\s*[,/]\s*)?)+\s*$")


def ad_temizle(ad: str) -> str:
    """Kalem adının sonundaki 'V-I-1' gibi dipnot referanslarını atar."""
    return DIPNOT_REF.sub("", ad or "").strip()


# ============================================================== 1) seçimler

def son_ceyrek() -> int:
    n = datetime.now()
    ay = (n.month - 1) // 3 * 3
    return (n.year - 1) * 100 + 12 if ay == 0 else n.year * 100 + ay


def sor(soru: str, varsayilan: str) -> str:
    try:
        c = input(f"{soru} [{varsayilan}]: ").strip()
    except EOFError:
        c = ""
    return c or varsayilan


def secimleri_al():
    print("=" * 64)
    print(" KAP finansal raporları ↔ TBB karşılaştırması")
    print("=" * 64)
    for no, ad, kisa, borsa, *_ in BANKALAR:
        print(f"  {no}) {ad}" + (f"  ({borsa})" if borsa else ""))
    while True:
        c = sor("\nHangi bankalar? Numaraları virgülle yazın (ör. 1,3,6) ya da 'hepsi'", "hepsi").lower()
        if c in ("hepsi", "tümü", "tumu", "h"):
            secilen = [b for b in BANKALAR]
            break
        try:
            nolar = sorted({int(x) for x in re.split(r"[,\s]+", c) if x})
            if nolar and all(1 <= n <= 7 for n in nolar):
                secilen = [b for b in BANKALAR if b[0] in nolar]
                break
        except ValueError:
            pass
        print("   1 ile 7 arasında numaralar girin, ör. 1,3,6")
    son = son_ceyrek()
    while True:
        d = sor("Dönem (YYYYAA, ay 03/06/09/12)", str(son))
        if re.fullmatch(r"20\d\d(03|06|09|12)", d) and int(d) <= son:
            break
        print(f"   Geçersiz dönem. Ör. {son}")
    yil, ay = int(d[:4]), int(d[4:])
    print(f"\nSeçilen: {', '.join(b[2] for b in secilen)} | {AY_ADI[ay]} {yil}")
    if sor("Başlatılsın mı? (E/h)", "E").lower() not in ("e", "evet", "y", "yes"):
        sys.exit("İptal edildi.")
    return secilen, yil, ay


# ============================================================== 2) KAP'tan indirme

OID_DOSYA = KLASOR / "kap_oid.json"


def oid_bul(banka) -> str:
    no, ad, kisa, borsa, tbb, oid = banka
    if oid:
        return oid
    kayit = js(OID_DOSYA.read_bytes()) if OID_DOSYA.exists() else {}
    if kayit and kisa in kayit:
        return kayit[kisa]
    hedef = sade(ad.replace("T.C.", "")).replace("TURKIYE CUMHURIYETI ", "")
    adaylar = []
    try:   # KAP akıllı arama
        sonuc = js(istek("POST", f"{KAP_API}/search/smart", {"keyword": "Ziraat Bankası"})) or []
        for kat in sonuc if isinstance(sonuc, list) else []:
            for it in (kat.get("results") or []) if isinstance(kat, dict) else []:
                if it.get("memberOrFundOid"):
                    adaylar.append((it.get("searchValue", ""), it["memberOrFundOid"]))
    except IstekHatasi:
        pass
    if not adaylar:   # üye listeleri
        for tur in ["IGS", "DDK", "YK", "PYS", "BDK", "DCS", "DK", "HT", "KVH"]:
            try:
                for u in js(istek("GET", f"{KAP_API}/company/items/{tur}/A")) or []:
                    o = u.get("memberOid") or u.get("mkkMemberOid")
                    if o:
                        adaylar.append((u.get("memberTitle") or u.get("title") or "", o))
            except IstekHatasi:
                continue
            if any(hedef in sade(a) for a, _ in adaylar):
                break
    uygun = [(a, o) for a, o in adaylar if hedef in sade(a) and "KATILIM" not in sade(a) and "DINAMIK" not in sade(a)]
    if not uygun:
        raise IstekHatasi(f"{ad}: KAP üye kimliği bulunamadı")
    oid = uygun[0][1]
    kayit = kayit or {}
    kayit[kisa] = oid
    OID_DOSYA.write_text(json.dumps(kayit, ensure_ascii=False, indent=1), encoding="utf-8")
    return oid


def zip_adi(bilgi: zipfile.ZipInfo) -> str:
    ad = bilgi.filename
    if not bilgi.flag_bits & 0x800:
        try:
            ad = ad.encode("cp437").decode("cp857")
        except (UnicodeEncodeError, UnicodeDecodeError):
            pass
    return Path(ad).name


def dosya_adi(ad: str) -> str:
    return re.sub(r'[\\/:*?"<>|]+', "_", ad).strip() or "dosya"


def kap_paketi_indir(banka, yil: int, ay: int, hedef: Path) -> list[Path]:
    oid = oid_bul(banka)
    donem = ay // 3
    meta = js(istek("GET", f"{KAP_API}/financialTable/listCompanyExcelMembers/{oid}/{yil}/{donem}"))
    if not meta:
        raise IstekHatasi(f"{AY_ADI[ay]} {yil} için KAP'ta finansal tablo yok (henüz yayımlanmamış olabilir)")
    icerik = istek("GET", f"{KAP_API}/home-financial/download-file/{oid}/{yil}/{donem}")
    hedef.mkdir(parents=True, exist_ok=True)
    pdfler = []
    if icerik[:4] == b"PK\x03\x04":
        (hedef / "kap_paketi.zip").write_bytes(icerik)
        with zipfile.ZipFile(io.BytesIO(icerik)) as zf:
            for bilgi in zf.infolist():
                if bilgi.is_dir():
                    continue
                yol = hedef / dosya_adi(zip_adi(bilgi))
                yol.write_bytes(zf.read(bilgi))
                if yol.suffix.lower() == ".pdf":
                    pdfler.append(yol)
    if not pdfler:   # pakette PDF yoksa: dönemin "Finansal Rapor" bildirimlerinin ekleri
        pdfler = bildirim_pdfleri(oid, yil, ay, hedef)
    return pdfler


def bildirim_pdfleri(oid: str, yil: int, ay: int, hedef: Path) -> list[Path]:
    bas = f"{yil}-{ay:02d}-01"
    bit_y, bit_a = (yil + 1, ay + 4 - 12) if ay + 4 > 12 else (yil, ay + 4)
    govde = {"fromDate": bas, "toDate": f"{bit_y}-{bit_a:02d}-28", "memberType": "", "mkkMemberOidList": [oid],
             "inactiveMkkMemberOidList": [], "disclosureClass": "FR", "subjectList": [], "isLate": "",
             "mainSector": "", "sector": "", "subSector": "", "marketOid": "", "index": "", "bdkReview": "",
             "bdkMemberOidList": [], "year": "", "term": "", "ruleType": "", "period": "", "fromSrc": False,
             "srcCategory": "", "disclosureIndexList": []}
    out = []
    for b in (js(istek("POST", f"{KAP_API}/disclosure/members/byCriteria", govde)) or [])[:6]:
        idx = b.get("disclosureIndex")
        for kayit in js(istek("GET", f"{KAP_API}/notification/attachment-detail/{idx}")) or []:
            for ek in kayit.get("attachments") or []:
                ad = ek.get("fileName") or ""
                if (ek.get("fileExtension") or Path(ad).suffix).lower().lstrip(".") != "pdf" or not ek.get("objId"):
                    continue
                yol = hedef / dosya_adi(f"{idx}_{ad if ad.lower().endswith('.pdf') else ad + '.pdf'}")
                yol.write_bytes(istek("GET", f"{KAP_API}/file/download/{ek['objId']}"))
                out.append(yol)
    return out


def pdf_turu(yol: Path) -> str | None:
    """'solo', 'kons' ya da None (İngilizce/ilgisiz)."""
    try:
        d = pymupdf.open(yol)
    except Exception:
        return None
    bas = "\n".join(d[i].get_text() for i in range(min(4, len(d)))).upper()
    if "BİLANÇO" not in bas and "FİNANSAL TABLO" not in bas and "FİNANSAL RAPOR" not in bas:
        return None
    if len(d) < 20:          # faaliyet raporu, sorumluluk beyanı vb.
        return None
    return "solo" if "KONSOLİDE OLMAYAN" in bas else "kons"


# ============================================================== 3) PDF okuma

KOD = re.compile(r"^(?:[A-H]\.|[IVXL]+\.|\d+(?:\.\d{1,2})+\.?|\d+\.)$")
KOD_METIN = re.compile(r"^((?:[A-H]|[IVXL]+)\.|\d+(?:\.\d{1,2})+\.?)\s+(\S.*)$")
SAYI_TR = re.compile(r"^\(?-?\d{1,3}(?:\.\d{3})*(?:,\d+)?\)?$|^-$")
SAYI_EN = re.compile(r"^\(?-?\d{1,3}(?:,\d{3})*(?:\.\d+)?\)?$|^-$")


def kur_baslik(sm: str) -> bool:
    """Kur riski tablosunun sütun başlığı: Avro/EURO/EUR, ABD Doları/USD ... Toplam."""
    return bool(re.search(r"\b(AVRO|EURO|EUR)\b", sm) and re.search(r"\b(ABD DOLARI|USD)\b", sm)
                and re.search(r"\bTOPLAM\b", sm))


def eksik_ad(ad):
    """Adın alt satırda sürdüğünü gösteren işaretler: açık parantez, virgül, bağlaç vb. ile bitiş."""
    ad = ad.strip()
    son = sade(ad.split()[-1]) if ad.split() else ""
    return (ad.count("(") > ad.count(")") or ad.endswith((",", "T.C."))
            or son in ("ZARARA", "GELIRE", "FINANSAL", "YANSITILAN", "KAPSAMLI", "KONTROL", "AMACLI", "VE", "VEYA", "ILE", "OLARAK", "UZERINDEN"))


class Okuyucu:
    """Bir raporun ana tablolarını okur. Sayı biçimi (1.234,5 / 1,234.5) ve
    birim (bin / milyon TL) rapordan bulunur."""

    def __init__(self, yol: Path):
        self.yol = yol
        self.bicim = "tr"
        self.carpan = 1
        self.tablolar: dict[str, list] = {}

    # --- sayı
    def sayi_mi(self, t: str) -> bool:
        return bool((SAYI_TR if self.bicim == "tr" else SAYI_EN).match(t))

    def sayi(self, t: str) -> float:
        t = t.strip()
        if t == "-":
            return 0.0
        neg = (t.startswith("(") and t.endswith(")")) or t.startswith("-")
        t = t.strip("()-")
        t = t.replace(".", "").replace(",", ".") if self.bicim == "tr" else t.replace(",", "")
        return -float(t) if neg else float(t)

    @staticmethod
    def kod_mu(t: str) -> bool:
        if not KOD.match(t):
            return False
        if re.fullmatch(r"\d+\.", t):
            return True
        return all(len(p) <= 2 for p in t.rstrip(".").split(".")[1:])

    @staticmethod
    def bolum(satir: str):
        ham = satir.strip()
        u = ham.upper().replace("KÂR", "KAR")
        if re.fullmatch(r"(KONSOLİDE )?VARLIKLAR", ham):
            return "VARLIKLAR"
        if re.fullmatch(r"(KONSOLİDE )?YÜKÜMLÜLÜKLER", ham):
            return "YUKUMLULUKLER"
        sm = sade(ham.replace("Â", "A").replace("â", "a"))
        if len(sm) > 140 or re.search(r"TABLOSUN[AD]|TABLOSUYLA|ILE ILISKILENDIR|YANSITIL|MUHASEBELES", sm):
            return None                                        # başlık değil, metin içinde geçiyor
        if re.search(r"NAZIM HESAPLAR( TABLOSU)?$|NAZIM HESAPLAR TABLOSU", sm):
            return "NAZIM"
        if re.search(r"KAPSAMLI GELIR TABLOSU", sm):
            return "DKG"
        if re.search(r"ZARAR TABLOSU( |$)|GELIR TABLOSU$", sm):
            return "KZ"
        if re.search(r"(DEGISIM|DEGISIKLIKLER) TABLOSU", sm):
            return "ATLA"
        if re.search(r"NAKIT AKIS TABLOSU", sm):
            return "NAKIT"
        return None

    def oku(self):
        d = pymupdf.open(self.yol)
        sayfalar = [p.get_text() for p in d]

        def bilanco_mu(t):
            st = sade(t)
            return (any(re.fullmatch(r"(KONSOLİDE )?VARLIKLAR", l.strip()) for l in t.split("\n"))
                    and "FINANSAL VARLIKLAR" in st and ("BILANCO" in st or "FINANSAL DURUM" in st
                                                        or "FINANSAL TABLOLAR" in st))
        bas = next((i for i, t in enumerate(sayfalar[:60]) if bilanco_mu(t)), None)
        if bas is None:
            raise ValueError("raporda bilanço bulunamadı")
        # birim: önce bilanço sayfası, sonra yakın sayfalar, sonra ilk sayfalar
        self.carpan = 1
        for i in [bas, bas - 1, bas + 1, bas - 2, bas + 2] + list(range(0, min(15, len(sayfalar)))):
            if not 0 <= i < len(sayfalar):
                continue
            st = sade(sayfalar[i].replace("(", " ").replace(")", " "))
            if re.search(r"\bMILYON (TURK LIRASI|TL)\b", st):
                self.carpan = 1000
                break
            if re.search(r"\bBIN (TURK LIRASI|TL)\b", st):
                break
        tr = len(re.findall(r"(?<![\d.,])\d{1,3}(?:\.\d{3})+(?![\d.,])", sayfalar[bas]))
        en = len(re.findall(r"(?<![\d.,])\d{1,3}(?:,\d{3})+(?![\d.,])", sayfalar[bas]))
        self.bicim = "en" if en > tr else "tr"
        sonuc: dict[str, list] = defaultdict(list)
        bolum = None
        for i in range(bas, min(len(sayfalar), bas + 20)):
            t = sayfalar[i]
            if i > bas and ("ÜÇÜNCÜ BÖLÜM" in t or "MUHASEBE POLİTİKALARI" in t.upper()):
                break
            parca: list[str] = []
            satirlar_ = t.split("\n")
            # bazı raporlarda tablo başlığı metinde sayfanın sonunda gelir (Garanti, İş): sayfanın üst
            # kısmındaki başlıklar öne alınır
            H = d[i].rect.height
            onde = []
            for j, satir in enumerate(satirlar_):
                if self.bolum(satir) not in (None, "VARLIKLAR", "YUKUMLULUKLER"):
                    hit = d[i].search_for(satir.strip()[:60])
                    if hit and hit[0].y0 < 0.3 * H and j > 3:
                        onde.append(j)
            if onde:
                satirlar_ = [satirlar_[j] for j in onde] + [x for j, x in enumerate(satirlar_) if j not in onde]
            for satir in satirlar_:
                b = self.bolum(satir)
                if b in ("VARLIKLAR", "YUKUMLULUKLER") and bolum not in (None, "VARLIKLAR"):
                    b = None                                   # bilanço geçildikten sonra aranmaz
                if b:
                    if bolum and parca:
                        sonuc[bolum].extend(self._satirlar(parca))
                    parca, bolum = [], b
                    continue
                parca.append(satir)
            if bolum and parca:
                sonuc[bolum].extend(self._satirlar(parca))
        # aynı tablo ikinci kez başlıyorsa (ör. Yapı Kredi'nin üç aylık gelir tablosu) ikinci kopya atılır
        for k_ in ("KZ", "DKG", "NAKIT"):
            v = sonuc.get(k_, [])
            tekrar = [j for j, r in enumerate(v) if r["kod"] == "I"]
            if len(tekrar) > 1 and tekrar[1] - tekrar[0] >= 15:
                sonuc[k_] = v[:tekrar[1]]
        self.tablolar = {k: v for k, v in sonuc.items() if k in dict((t[0], 1) for t in TABLOLAR)}
        try:
            self.tablolar["OZK"] = self._ozkaynak(d, bas)
        except Exception as e:                                  # tablo okunamazsa diğerleri sürer
            print(f"      ! özkaynak tablosu okunamadı: {e}")
            self.tablolar["OZK"] = []
        try:
            self.tablolar["KUR"] = self._kur_riski(d, bas)
        except Exception as e:
            print(f"      ! kur riski tablosu okunamadı: {e}")
            self.tablolar["KUR"] = []
        for spec in DIPNOT_TABLOLARI:
            try:
                self.tablolar[spec["anahtar"]] = (self._lko_uc(d, bas) if spec["anahtar"] == "LKO_UC"
                                                  else self._matris(d, bas, spec))
            except Exception as e:
                print(f"      ! {spec['sayfa']} tablosu okunamadı: {e}")
                self.tablolar[spec["anahtar"]] = []
        return self

    # --- mali bünye dipnot tabloları
    def _satirlar_xy(self, sayfa) -> list[list]:
        """Kelimeleri satırlara toplar: [(x0, x1, kelime, kalın_mı), ...] (soldan sağa)."""
        kalin = []
        for blok in sayfa.get_text("dict")["blocks"]:
            for sat in blok.get("lines", []):
                for sp in sat.get("spans", []):
                    if sp.get("flags", 0) & 16 or "bold" in sp.get("font", "").lower():
                        kalin.append(sp["bbox"])

        def kalin_mi(x0, y0, x1, y1):
            cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
            return any(b[0] - 1 <= cx <= b[2] + 1 and b[1] - 1 <= cy <= b[3] + 1 for b in kalin)
        kel = sorted(((y0 + y1) / 2, x0, x1, w, kalin_mi(x0, y0, x1, y1))
                     for x0, y0, x1, y1, w, *_ in sayfa.get_text("words"))
        gruplar: list[list] = []
        for yc, x0, x1, w, kl in kel:
            if gruplar and abs(yc - gruplar[-1][0]) <= 3.0:
                g = gruplar[-1]
                g[1].append((x0, x1, w, kl))
                g[0] = (g[0] * (len(g[1]) - 1) + yc) / len(g[1])
            else:
                gruplar.append([yc, [(x0, x1, w, kl)]])
        return [sorted(g[1]) for g in gruplar]

    @staticmethod
    def sayfa_basligi_mi(metin: str) -> bool:
        sm = sade(metin.replace("(", " ").replace(")", " "))
        return (re.fullmatch(r"\d{1,3}", metin.strip()) is not None or "(DEVAMI)" in metin.upper()
                or any(k in sm for k in ("DIPNOTLAR", "TUTARLAR AKSI", "ANONIM ORTAKLIGI", "BANKASI A S",
                                          "BANKASI T A", "TARIHI ITIBARIYLA", "MALI BUNYEYE", "DEVAMI")))

    def _matris(self, d, bas: int, spec: dict, numarali=None) -> list[dict]:
        """Satır × sütun dipnot tablosu. Cari / Önceki Dönem başlıklarıyla bloklara ayrılır; rakamlar satırın
        sağındaki sayılardır ve x konumuna göre sütunlara dağıtılır (boş hücreler atlanmaz)."""
        def bulundu(i):
            st = sade(d[i].get_text())
            if spec["bul"](st):
                return True
            # başlık önceki sayfanın sonunda, tablo bu sayfada başlıyor (İş, Garanti, YKB faiz riski)
            onceki_son = sade(d[i - 1].get_text()[-700:]) if i > bas else ""
            return bool(onceki_son) and spec["bul"](onceki_son + " " + st)
        bas_sayfa = next((i for i in range(bas, len(d)) if bulundu(i)), None)
        if bas_sayfa is None:
            return []
        bas_roma = next((m.group(1) for m in (re.match(r"^(I|II|III|IV|V|VI|VII|VIII|IX|X)\.(\s|$)", x.strip())
                                              for x in d[bas_sayfa].get_text().split("\n")) if m), None)
        ncol = len(spec["sutun"]) * (2 if spec.get("tek_blok") else 1)
        numarali = spec.get("numarali", False) if numarali is None else numarali
        satirlar, tampon, acik, bekleyen_toks, basliklar_ = [], [], None, [], []
        donem, bolum, dipnot_modu, bitti, aktif = None, None, False, False, not spec.get("ilk")
        for i in range(bas_sayfa, min(len(d), bas_sayfa + spec.get("sayfa_say", 3))):
            sayfa = d[i]
            W = sayfa.rect.width
            for kel in self._satirlar_xy(sayfa):
                metin = " ".join(w for _, _, w, _ in kel).strip()
                sm = sade(metin)
                sag = []
                for x0, x1, w, kl in reversed(kel):
                    ww = w.strip("%") if w not in ("%",) else w
                    if spec.get("basliklar") and re.fullmatch(r"\(\d\)", ww):
                        break                                   # dipnot atfı, ör. "Değerler (3)"
                    if x0 > W * spec["xmin"] and (self.sayi_mi(ww) or ww in ("-", "–")) and ww != "%":
                        sag.insert(0, (x1, ww, (x0 + x1) / 2))
                    else:
                        break
                sol = kel[:len(kel) - len(sag)]
                no = None
                if numarali and sol and re.fullmatch(r"\d{1,2}\.?", sol[0][2]) and sol[0][0] < W * 0.24:
                    no = int(sol[0][2].rstrip("."))
                    sol = sol[1:]
                etiket = " ".join(w for _, _, w, _ in sol).strip()
                kalin = bool(sol) and sum(kl for *_, kl in sol) >= 0.6 * len(sol)
                if spec.get("basliklar"):
                    bulunan = [None] * len(spec["basliklar"])
                    for q, (x0, x1, w, kl) in enumerate(kel):
                        for j, desen in enumerate(spec["basliklar"]):
                            if bulunan[j] is None and re.fullmatch(desen, sade(w)):
                                x1b = kel[q + 1][1] if q + 1 < len(kel) and sade(kel[q + 1][2]) == "DOLARI" else x1
                                bulunan[j] = (x0 + x1b) / 2
                    if sum(x is not None for x in bulunan) >= 2 and not sag:
                        basliklar_.append((len(satirlar), bulunan))
                        if not re.search(r"\b(CARI|ONCEKI) DONEM", sm):
                            continue
                dm = re.search(r"\b(CARI|ONCEKI) DONEM", sm)
                if dm and (dm.start() == 0 or (kalin and len(sm) < 90)) and all(
                        w in ("-", "–") or re.fullmatch(r"\(\d\)", w) for _, w, *_ in sag):
                    donem = "cari" if dm.group(1) == "CARI" else "onceki"
                    bolum, tampon, acik, dipnot_modu, bekleyen_toks = None, [], None, False, []
                    aktif = not spec.get("ilk")
                    continue
                if (not sag and kel[0][0] < W * 0.1 and re.match(r"^\d+(\.\d+)+\.?\s+\S", metin)
                        and any(r["donem"] == "onceki" for r in satirlar)):
                    break_ = True
                    bitti = True                                  # "4.5 Hisse senedi …" yeni bölüm
                    break
                satir_no_yalniz = (numarali and no is not None and not etiket and not sag
                                   and kel[0][0] < W * 0.24)       # "4" tek başına: satır numarası
                if donem is None or (self.sayfa_basligi_mi(metin) and not satir_no_yalniz):
                    continue
                if spec.get("basliklar") and not basliklar_:
                    continue                                    # para birimi başlığı görülmeden satır alma
                if not aktif:
                    aktif = bool(re.search(spec["ilk"], sm))
                    if not aktif:
                        continue
                if spec.get("bos_satir") and not sag and no is None and re.search(spec["bos_satir"], sm) and not dipnot_modu:
                    satirlar.append({"donem": donem, "bolum": bolum, "no": None, "ad": etiket, "toks": []})
                    tampon = []
                    continue
                if not sag and no is None and re.match(r"^\((\*+|\d{1,2})\)", metin):
                    if any(r["donem"] == "onceki" for r in satirlar) or spec.get("tek_blok"):
                        bitti = True
                        break
                    dipnot_modu, tampon, acik = True, [], None     # cari bloğun altındaki açıklamalar
                    continue
                if dipnot_modu:
                    continue
                roma = re.match(r"^(I|II|III|IV|V|VI|VII|VIII|IX|X)\.\s", metin)
                if not sag and no is None and kalin and roma and roma.group(1) != bas_roma:
                    bitti = True                                  # sonraki bölüm başladı
                    break
                se = sade(etiket)
                if no is None and sag and all(w in ("-", "–") for _, w, *_ in sag) and (
                        etiket.endswith(":") or se in ("VARLIKLAR", "YUKUMLULUKLER", "AKTIFLER", "PASIFLER")):
                    bolum, tampon = se, []                        # "Yükümlülükler:  -  -  -" bölüm başlığı
                    continue
                if no is not None:
                    son_no = max((r["no"] for r in satirlar if r["donem"] == donem and r["no"]), default=0)
                    if no <= son_no or (spec.get("son_no") and no > spec["son_no"]):
                        continue                                  # tablo dışı metin (ör. "30 Eylül 2026 …")
                    acik = {"donem": donem, "bolum": bolum, "no": no,
                            "ad": " ".join(tampon + [etiket]).strip(), "toks": sorted(bekleyen_toks + sag)}
                    tampon, bekleyen_toks = [], []
                    satirlar.append(acik)
                elif sag:
                    if numarali:
                        if acik is None or acik["toks"]:
                            # satır numarasından önce yazılmış rakam / ad parçası: sonraki numaralı satıra eklenir
                            bekleyen_toks += sag
                            if etiket:
                                tampon.append(etiket)
                            continue
                        acik["ad"] = " ".join([acik["ad"]] + tampon + [etiket]).strip()
                        acik["toks"] = sag
                        tampon = []
                    else:
                        if (spec.get("basliklar") and tampon and etiket[:1].isupper()
                                and not eksik_ad(" ".join(tampon))):
                            satirlar.append({"donem": donem, "bolum": bolum, "no": None,
                                             "ad": " ".join(tampon), "toks": []})   # boş satır ("Muhtelif Borçlar")
                            tampon = []
                        satirlar.append({"donem": donem, "bolum": bolum, "no": None,
                                         "ad": " ".join(tampon + [etiket]).strip(), "toks": sag})
                        tampon = []
                else:
                    if sm in ("VARLIKLAR", "YUKUMLULUKLER", "AKTIFLER", "PASIFLER") or (metin.endswith(":") and len(sm) < 40):
                        bolum, tampon = sm, []
                    elif kalin:
                        bolum, tampon = sm, []
                    elif kel[0][0] < W * 0.35:
                        if spec.get("basliklar") or spec.get("esle") == "kanon":
                            # rakamları dikeyde ortalanmış çok satırlı adlar (İş, Garanti)
                            onceki_s = satirlar[-1] if satirlar and satirlar[-1]["donem"] == donem else None
                            if not tampon and onceki_s and onceki_s["toks"] and (
                                    eksik_ad(onceki_s["ad"]) or etiket.count(")") > etiket.count("(")):
                                onceki_s["ad"] += " " + etiket
                                continue
                            if spec.get("basliklar") and tampon and not eksik_ad(" ".join(tampon)) and etiket[:1].isupper():
                                satirlar.append({"donem": donem, "bolum": bolum, "no": None,
                                                 "ad": " ".join(tampon), "toks": []})
                                tampon = []
                        tampon.append(etiket)
                    continue
                son = satirlar[-1]
                son["ad"] = ad_temizle(re.sub(r"\s*\((\*+|\d{1,2})\)", " ", son["ad"])).strip(" :")
                if spec.get("son") and son["toks"] and spec["son"](son):
                    bitti = True
                    break
                if spec.get("son_no") and son.get("no") == spec["son_no"]:
                    dipnot_modu = True                            # blok bitti; sonraki dönem başlığına kadar atla
            if bitti:
                break
        # rakamları sütunlara dağıt: dolu satırlardaki sağ kenarlar sütun çapalarıdır
        capa = {}
        for dn in ("cari", "onceki"):
            dolu = [r["toks"] for r in satirlar if r["donem"] == dn and len(r["toks"]) == ncol]
            if dolu:
                capa[dn] = [sorted(t[j][0] for t in dolu)[len(dolu) // 2] for j in range(ncol)]
        for ri, r in enumerate(satirlar):
            c = capa.get(r["donem"]) or next(iter(capa.values()), None)
            sayilar = [None] * ncol
            toks = r.pop("toks")
            bc = next((b_ for i_, b_ in reversed(basliklar_) if i_ <= ri), None)
            if bc is not None:                         # sütun başlıklarının konumuna göre (ör. Yen sütunu olmayan rapor)
                for x1, w, xc in toks:
                    js = [j for j in range(ncol) if bc[j] is not None and sayilar[j] is None]
                    if js:
                        j = min(js, key=lambda jj: abs(bc[jj] - xc))
                        sayilar[j] = w
            elif c is None or len(toks) >= ncol:
                for j, (_, w, *_r) in enumerate(toks[-ncol:]):
                    sayilar[ncol - min(len(toks), ncol) + j] = w
            else:
                onceki_j = -1
                for k, (x1, w, *_r) in enumerate(toks):
                    kalan = len(toks) - k - 1
                    adaylar = range(onceki_j + 1, ncol - kalan)
                    j = min(adaylar, key=lambda jj: abs(c[jj] - x1))
                    sayilar[j], onceki_j = w, j
            r["sayilar"] = [None if w is None else (0.0 if w in ("-", "–") else self.sayi(w)) for w in sayilar]
        sonuc = [r for r in satirlar if r["ad"] or r["no"] is not None]
        if numarali and spec.get("desen") and not any(r["no"] for r in sonuc):
            # satır numarası yazılmamış (ör. Ziraat): numarasız oku, numarayı kalem adından bul
            desen = globals()[spec["desen"]]
            sonuc = self._matris(d, bas, spec, numarali=False)
            for dn in ("cari", "onceki"):
                son_no = 0
                for r in [x for x in sonuc if x["donem"] == dn]:
                    ad = sade2(r["ad"])
                    r["no"] = next((n for n in sorted(desen) if n > son_no and re.search(desen[n], ad)), None)
                    son_no = r["no"] or son_no
            sonuc = [r for r in sonuc if r["no"] is not None]
        return sonuc

    def _lko_uc(self, d, bas: int) -> list[dict]:
        """Son üç aydaki haftalık LKO'nun en düşük / en yüksek değerleri ve tarihleri."""
        satirlar, donem, ters = [], None, False
        bas_sayfa = next((i for i in range(bas, len(d)) if all(k in sade(d[i].get_text()) for k in
                         ("LIKIDITE KARSILAMA ORAN", "EN DUSUK", "EN YUKSEK", "TARIH"))), None)
        if bas_sayfa is None:
            return []
        tarih = r"((?:\d{1,2}\s+)?[A-Za-zÇĞİÖŞÜçğıöşü]+\s+\d{4})\s+(\(?-?[\d.,]+\)?)"
        for i in range(bas_sayfa, min(len(d), bas_sayfa + 3)):
            for kel in self._satirlar_xy(d[i]):
                metin = " ".join(w for _, _, w, _ in kel).strip()
                sm = sade(metin)
                if re.match(r"^(CARI|ONCEKI) DONEM", sm):
                    donem = "cari" if sm.startswith("CARI") else "onceki"
                    continue
                if "TP YP" in sm and re.search(r"\bYP\b", sm.replace("TP YP", "")) and len(sm) < 30:
                    ters = sm.index("TP YP") > sm.replace("TP YP", "    ").index("YP")   # önce YP yazılmış
                    continue
                m = re.match(r"^EN (DUSUK|YUKSEK)\b", sm)
                if donem and m:
                    bul = re.findall(tarih, metin)
                    if ters:
                        bul = bul[1::-1] + bul[2:]
                    if len(bul) >= 2 and not any(r["donem"] == donem and sade(r["ad"]).endswith(m.group(1)) for r in satirlar):
                        satirlar.append({"donem": donem, "bolum": None, "no": None,
                                         "ad": "En düşük" if m.group(1) == "DUSUK" else "En yüksek",
                                         "tarih": [bul[0][0], bul[1][0]],
                                         "sayilar": [self.sayi(bul[0][1]), self.sayi(bul[1][1])]})
            if sum(1 for r in satirlar if r["donem"] == "onceki") >= 2:
                break
        return satirlar

    def _koordinat_satirlari(self, sayfa) -> list[list]:
        """Sayfadaki kelimeleri yüksekliğe göre satırlara toplar: [(x0, kelime, kalın_mı), ...]"""
        kalin = []
        for blok in sayfa.get_text("dict")["blocks"]:
            for sat in blok.get("lines", []):
                for sp in sat.get("spans", []):
                    if sp.get("flags", 0) & 16 or "bold" in sp.get("font", "").lower():
                        kalin.append(sp["bbox"])

        def kalin_mi(x0, y0, x1, y1):
            cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
            return any(b[0] - 1 <= cx <= b[2] + 1 and b[1] - 1 <= cy <= b[3] + 1 for b in kalin)
        kelimeler = sorted(((y0 + y1) / 2, x0, w, kalin_mi(x0, y0, x1, y1))
                           for x0, y0, x1, y1, w, *_ in sayfa.get_text("words"))
        gruplar: list[list] = []
        for yc, x0, w, kl in kelimeler:
            if gruplar and abs(yc - gruplar[-1][0]) <= 3.0:
                g = gruplar[-1]
                g[1].append((x0, w, kl))
                g[0] = (g[0] * (len(g[1]) - 1) + yc) / len(g[1])
            else:
                gruplar.append([yc, [(x0, w, kl)]])
        return [sorted(g[1]) for g in gruplar]

    def _kur_riski(self, d, bas: int) -> list[dict]:
        """'Kur riskine ilişkin bilgiler' tablosu: her satır Avro, ABD Doları, Diğer YP, Toplam."""
        bas_sayfa = None
        for i in range(bas, len(d)):
            st = sade(d[i].get_text())
            onceki_st = sade(d[i - 1].get_text()) if i > 0 else ""
            t = d[i].get_text()
            if ("Avro" in t and "ABD Doları" in t and "KUR RISKINE ILISKIN" in st and "Toplam" in t) or (
                    ("KUR RISK" in st or "KUR RISK" in onceki_st)
                    and any(kur_baslik(sade(" ".join(w for _, w, _ in kel)))
                            for kel in self._koordinat_satirlari(d[i]))):
                bas_sayfa = i
                break
        if bas_sayfa is None:
            return []
        satirlar, bekleyen = [], []
        donem, bolum, basladi = "cari", None, False
        for i in range(bas_sayfa, min(len(d), bas_sayfa + 3)):
            sayfa = d[i]
            genislik = sayfa.rect.width
            for kel in self._koordinat_satirlari(sayfa):
                metin = " ".join(w for _, w, _ in kel).strip()
                sm = sade(metin)
                sag = []
                for x, w, _ in reversed(kel):
                    if x > genislik * 0.45 and (self.sayi_mi(w) or w in ("-", "–")):
                        sag.insert(0, w)
                    else:
                        break
                etiket = " ".join(w for _, w, _ in kel[:len(kel) - len(sag)]).strip()
                if not basladi:
                    basladi = kur_baslik(sm)
                    if basladi and "ONCEKI DONEM" in sm:
                        donem = "onceki"
                    continue
                if "ONCEKI DONEM" in sm and not sag:
                    donem, bolum, bekleyen = "onceki", None, []
                    continue
                if "CARI DONEM" in sm and not sag:
                    donem, bekleyen = "cari", []
                    continue
                if not sag:
                    if kur_baslik(sm) or ("AVRO" in sm and "ABD DOLARI" in sm):
                        continue
                    son_s = satirlar[-1] if satirlar and satirlar[-1]["donem"] == donem else None
                    if sm == "VARLIKLAR" and not bekleyen and son_s and eksik_ad(son_s["ad"]):
                        son_s["ad"] += " " + etiket                 # "… Finansal / Varlıklar" devam satırı
                        continue
                    if sm == "VARLIKLAR":
                        bolum, bekleyen = "varlik", []
                        continue
                    if sm == "YUKUMLULUKLER":
                        bolum, bekleyen = "yukum", []
                        continue
                    if any(k in sm for k in ("DIPNOTLAR", "TUTARLAR AKSI", "ANONIM ORTAKLIGI", "BANKASI A S",
                                             "BANKASI T A", "TARIHI ITIBARIYLA", "MALI BUNYEYE")) \
                            or "(DEVAMI)" in metin.upper() or re.fullmatch(r"\d{1,3}", metin):
                        continue
                    if satirlar and (re.match(r"^\(\d\)", metin) or re.match(r"^(\d+|[IVX]+)\.(\s|$)", metin)):
                        if any(r["donem"] == "onceki" for r in satirlar):
                            return satirlar                      # tablo bitti (dipnot açıklamaları)
                        continue
                    son_s = satirlar[-1] if satirlar and satirlar[-1]["donem"] == donem else None
                    if not bekleyen and son_s and (eksik_ad(son_s["ad"]) or etiket.count(")") > etiket.count("(")):
                        son_s["ad"] = ad_temizle(son_s["ad"] + " " + etiket)   # dikeyde ortalanmış çok satırlı ad
                        continue
                    bekleyen.append(etiket)
                    continue
                if len(sag) < 4:
                    bekleyen = []
                    continue
                ad = " ".join(bekleyen + [etiket])
                bekleyen = []
                ad = ad_temizle(re.sub(r"\s*\(\d\)", "", ad)).strip(" :")
                sa = sade(ad)
                b = bolum
                if donem == "onceki" or sa.startswith(("NET BILANCO", "NET NAZIM", "TUREV", "GAYRINAKDI")):
                    b = "varlik" if sa.startswith("TOPLAM VARLIK") else ("yukum" if sa.startswith("TOPLAM YUKUM") else "ozet")
                satirlar.append({"kod": "", "ad": ad, "bolum": b, "donem": donem,
                                 "sayilar": [0.0 if w in ("-", "–") else self.sayi(w) for w in sag[-4:]]})
                if donem == "onceki" and sa.startswith("GAYRINAKDI"):
                    return satirlar
        return satirlar

    def _ozkaynak(self, d, bas: int) -> list[dict]:
        def yeni_kalem(onceki_ad: str, etiket: str) -> bool:
            """")" ile biten tam bir ad + büyük harfle başlayan yeni satır: ayrı kalemler
            ("… (Geçici Madde 4 kapsamında olanlar)" / "Karşılıklar (…)" – İş)."""
            if not etiket[:1].isupper() or etiket.startswith("("):
                return False
            # tam ad: ")" ile ya da kalem sonu kelimesiyle biter ("… kaynaklanan tutar" / "Ortaklık paylarının …" – İş)
            return onceki_ad.rstrip().endswith(")") or (
                len(onceki_ad.split()) >= 8 and re.search(r"(TUTAR|TUTARI|KISMI|TOPLAMI|YATIRIMLAR|PAYLAR)$", sade(onceki_ad))
                is not None and re.match(r"^(ORTAKLIK|BANKA|KANUNUN|GECICI|IPOTEK|STANDART|TOPLAM|KURUMCA|NET|SEREFIYE)",
                                         sade(etiket)) is not None)
        """'Özkaynak kalemlerine ilişkin açıklamalar' tablosu. Kalem numarası yoktur; satırlar
        kelime koordinatlarından kurulur, rakamlar sayfanın sağındaki sütunlardan alınır."""
        def tablo_sayfasi(t):
            st = sade(t)
            return ("OZKAYNAK KALEMLERINE ILISKIN" in st
                    and any(sade(x) == "CEKIRDEK SERMAYE" for x in t.split("\n")))
        bas_sayfa = next((i for i in range(bas, min(len(d), bas + 80)) if tablo_sayfasi(d[i].get_text())), None)
        if bas_sayfa is None:
            return []
        satirlar, bekleyen, basladi = [], [], False
        bekleyen_kalin = [False]
        # rakamları dikeyde ortalanmış düzen (İş): adı olmayan, yalnız iki rakamlı satırlar sık görülür
        yalniz_rakam = sum(1 for x in d[bas_sayfa].get_text().split("\n")[:0]) if False else 0
        for kel in self._koordinat_satirlari(d[bas_sayfa]):
            if kel and all(self.sayi_mi(w) or w in ("-", "–") for _, w, _ in kel):
                yalniz_rakam += 1
        self._ozk_dikey = yalniz_rakam >= 3

        def bekleyeni_bosalt():
            for x in bekleyen:
                satirlar.append({"kod": "", "ad": x, "sayilar": []})
            bekleyen.clear()
        for i in range(bas_sayfa, min(len(d), bas_sayfa + 5)):
            sayfa = d[i]
            genislik = sayfa.rect.width
            # kalın (başlık) yazı bölgeleri
            kalin = []
            for blok in sayfa.get_text("dict")["blocks"]:
                for sat in blok.get("lines", []):
                    for sp in sat.get("spans", []):
                        if sp.get("flags", 0) & 16 or "bold" in sp.get("font", "").lower():
                            kalin.append(sp["bbox"])

            def kalin_mi(x0, y0, x1, y1):
                cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
                return any(b[0] - 1 <= cx <= b[2] + 1 and b[1] - 1 <= cy <= b[3] + 1 for b in kalin)
            # aynı yükseklikteki kelimeler bir satır (yakınlık toleranslı kümeleme)
            kelimeler = sorted(((y0 + y1) / 2, x0,
                                w.strip("%") if "%" in w and self.sayi_mi(w.strip("%")) else w,   # "11,98%"
                                kalin_mi(x0, y0, x1, y1))
                               for x0, y0, x1, y1, w, *_ in sayfa.get_text("words"))
            gruplar: list[list] = []
            for yc, x0, w, kl in kelimeler:
                if gruplar and abs(yc - gruplar[-1][0]) <= 3.0:
                    g = gruplar[-1]
                    g[1].append((x0, w, kl))
                    g[0] = (g[0] * (len(g[1]) - 1) + yc) / len(g[1])
                else:
                    gruplar.append([yc, [(x0, w, kl)]])
            bitti = False
            for _, kel in gruplar:
                kel = sorted(kel)
                yazi = [(x, w, kl) for x, w, kl in kel if not (self.sayi_mi(w) or w in ("-", "–"))]
                satir_kalin = bool(yazi) and sum(kl for *_, kl in yazi) >= 0.6 * len(yazi)
                ks = [(x, w) for x, w, kl in kel]
                if re.fullmatch(OZK_ARA_BASLIK, sade(" ".join(w for _, w, _ in yazi))):
                    satir_kalin = True                         # kalın yazılmamış ara başlık (Akbank)
                metin = " ".join(w for _, w in ks).strip()
                if not basladi:
                    basladi = metin.upper().startswith("ÇEKİRDEK SERMAYE") and len(metin) < 30
                    if not basladi:
                        continue
                if (re.match(r"^(II|III|IV|V)\.\s+[A-ZÇĞİÖŞÜ]", metin) or metin.startswith("(*)")
                        or sade(metin).startswith(("BORCLANMA ARACININ", "IHRACCI"))):
                    bitti = True
                    break
                # rakam sütunları: sayfanın sağında ve sağında başka kelime olmayan sayılar
                sag = []
                for x, w in reversed(ks):
                    if x > genislik * 0.62 and (self.sayi_mi(w) or w in ("-", "–")):
                        sag.insert(0, (x, w))
                    else:
                        break
                etiket = " ".join(w for x, w in ks if (x, w) not in sag).strip()
                if re.fullmatch(r"\d{1,3}", metin) or "Tutarlar aksi belirtilmedikçe" in metin \
                        or metin.upper().startswith(("30 ", "31 ", "MALİ BÜNYEYE", "I. ÖZKAYNAK")) \
                        or metin in ("Cari", "Önceki", "Dönem", "Cari Önceki", "Dönem Dönem") \
                        or any(k in sade(metin) for k in ("DIPNOTLAR", "DEVAMI", "BANKASI A S", "BANKASI T A",
                                                            "ANONIM ORTAKLIGI", "TARIHI ITIBARIYLA"))\
                        or sade(metin).startswith(("CARI DONEM", "ONCEKI DONEM", "TUTARLAR AKSI")) \
                        or "(DEVAMI)" in metin.upper() or "MALİ BÜNYEYE" in metin.upper() \
                        or sade(metin) in ("FINANSAL KURULUSLARI", "VE FINANSAL KURULUSLARI") \
                        or (len(metin) < 90 and re.search(r"FINANSAL (RAPORU|TABLOLAR)$|^PARA BIRIMI|^\(?PARA BIRIMI",
                                                          sade(metin))) \
                        or re.match(r"^\(?\s*Para birimi", metin, re.I):
                    continue                                    # sayfa başlığı / numarası
                if not sag:
                    if (not bekleyen and satirlar and not satirlar[-1]["ad"].strip() and satirlar[-1]["sayilar"]
                            and not satirlar[-1].get("baslik")):
                        satirlar[-1]["ad"] = etiket           # rakamlar adın üstündeki satırda (Garanti)
                        continue
                    if satir_kalin:                            # kalın, rakamsız satır: başlık
                        bekleyeni_bosalt()
                        onceki = satirlar[-1] if satirlar else None
                        if onceki and onceki.get("baslik") and not onceki["ad"].isupper() and not etiket.isupper():
                            onceki["ad"] += " " + etiket               # iki satıra taşan başlık
                        else:
                            satirlar.append({"kod": "", "ad": etiket, "sayilar": [], "baslik": True})
                    elif (not bekleyen and etiket[:1].islower() and satirlar and satirlar[-1]["sayilar"]
                          and not satirlar[-1].get("baslik")):
                        satirlar[-1]["ad"] += " " + etiket    # rakamları dikeyde ortalanmış adın son satırı
                    elif bekleyen and not bekleyen_kalin[0] and not yeni_kalem(bekleyen[-1], etiket):
                        bekleyen[-1] += " " + etiket          # adın devamı
                    else:
                        bekleyeni_bosalt()
                        bekleyen.append(etiket)
                    bekleyen_kalin[0] = False
                    continue
                if bekleyen and (not satir_kalin or not etiket or etiket[:1].islower()) and not (
                        etiket[:1].isupper() and not satir_kalin and self._ozk_dikey) and not yeni_kalem(bekleyen[-1], etiket):
                    tam = (bekleyen.pop() + " " + etiket).strip()
                    bekleyeni_bosalt()
                else:
                    bekleyeni_bosalt()
                    tam = etiket
                sayilar = [0.0 if w in ("-", "–") else self.sayi(w) for _, w in sag]
                satirlar.append({"kod": "", "ad": ad_temizle(re.sub(r"\s*\(\*+\)\s*", " ", tam)).strip(),
                                 "sayilar": sayilar[-2:]})
                if OZK_SON in sade(tam):                       # tablonun son kalemi
                    bitti = True
                    break
            if bitti:
                break
        bekleyeni_bosalt()
        # ikinci "Toplam Özkaynak" satırının üstünde "ÖZKAYNAK" ara başlığı yoksa ekle (YKB): TBB'de bu satır
        # "ÖZKAYNAK" başlığının altındaki "Yasal Özkaynak" kalemi
        for r in satirlar:
            if r.get("baslik") and re.match(r"^OZKAYNAK CARI DONEM", sade(r["ad"])):
                r["ad"] = "ÖZKAYNAK"                          # "ÖZKAYNAK  Cari Dönem  Önceki Dönem" (YKB)
        gorulen = 0
        for j, r in enumerate(list(satirlar)):
            sa = sade(r["ad"])
            if re.match(r"^(OZKAYNAK )?TOPLAM OZKAYNAK", sa) and r["sayilar"]:
                gorulen += 1
                onceki_ad = sade(satirlar[satirlar.index(r) - 1]["ad"]) if satirlar.index(r) else ""
                if gorulen == 2 and not sa.startswith("OZKAYNAK") and onceki_ad != "OZKAYNAK":
                    satirlar.insert(satirlar.index(r), {"kod": "", "ad": "ÖZKAYNAK", "sayilar": [], "baslik": True})
        for r in satirlar:     # oran satırlarını işaretle
            for desen, anahtar in OZK_ORANLAR:
                if re.search(desen, sade(r["ad"])) and r["sayilar"]:
                    r["oran"] = anahtar
                    break
        return satirlar

    def _satirlar(self, satirlar: list[str]) -> list[dict]:
        toks = []
        for x in (y.strip() for y in satirlar):
            if not x:
                continue
            # aynı satırda yazılmış rakamları ayır ("606.529   647.095   1.253.624", "Bankalar  (3)  42.007")
            w = x.split()
            j = len(w)
            while j > 0 and (self.sayi_mi(w[j - 1]) or w[j - 1] in ("-", "–")):
                j -= 1
            if len(w) > 1 and j < len(w) and not (j == 0 and len(w) == 1):
                if j:
                    toks.append(" ".join(w[:j]))
                toks.extend(w[j:])
            else:
                toks.append(x)
        out, cur, basladi = [], None, False
        for x in toks:
            km = KOD_METIN.match(x)
            if cur and cur["ad"] and not cur["sayilar"] and re.fullmatch(r"[1-9]\d?(\.\d{1,2})+\.?", x):
                continue                                       # dipnot sütunundaki referans (5.1.1, 1.1)
            if self.kod_mu(x) or (km and not self.sayi_mi(x)):
                if cur:
                    out.append(cur)
                if km and not self.kod_mu(x):
                    cur = {"kod": km.group(1).rstrip("."), "ad": km.group(2).strip(), "sayilar": []}
                else:
                    cur = {"kod": x.rstrip("."), "ad": "", "sayilar": []}
                basladi = True
                continue
            if not basladi or cur is None:
                continue
            if not cur["sayilar"] and re.fullmatch(r"\(\d{1,2}\)(?:\s*,\s*\(\d{1,2}\))*", x):
                if not re.fullmatch(r"\(\d{1,2}\)", x):
                    continue                                   # dipnot referansı (5),(6)
                cur["atif_adayi"] = True                      # "(1)" atıf da olabilir, "(88)" tutar da: sonra karar
            if self.sayi_mi(x):
                cur["sayilar"].append(self.sayi(x))
            elif re.fullmatch(r"\(?[IVX]{1,4}-[IVX]{1,4}(?:-\d+(?:\.\d+)*)?\)?[,.]?", x) and not cur["sayilar"]:
                continue                                       # dipnot referansı V-I-1
            elif not cur["sayilar"]:
                cur["ad"] = (cur["ad"] + " " + x).strip()
            else:                                              # numarasız satır (ör. VARLIKLAR TOPLAMI)
                out.append(cur)
                cur = {"kod": "", "ad": x, "sayilar": []}
        if cur:
            out.append(cur)
        # "(1)" adayları: satırda olağan sütun sayısından bir fazla rakam varsa ilki dipnot atfıdır
        sayim = Counter(len(r["sayilar"]) for r in out if r["sayilar"])
        mod = sayim.most_common(1)[0][0] if sayim else 0
        for r in out:
            if r.pop("atif_adayi", False) and len(r["sayilar"]) == mod + 1:
                r["sayilar"] = r["sayilar"][1:]
        for r in out:
            r["ad"] = ad_temizle(r["ad"])
        return [r for r in out if r["sayilar"] or (r["kod"] and r["ad"])]


# ============================================================== 4) TBB verisi

def tbb_verisi() -> dict:
    onbellek = KLASOR / "tbb_veri_onbellek.json"
    try:
        print("   TBB verisi GitHub'dan indiriliyor (~7 MB)…", flush=True)
        html = istek("GET", TBB_URL).decode("utf-8")
        i = html.index("const DATA = ") + len("const DATA = ")
        D, _ = json.JSONDecoder().raw_decode(html, i)
        onbellek.write_text(json.dumps(D), encoding="utf-8")
        return D
    except Exception as e:
        if onbellek.exists():
            print(f"   ! TBB verisi indirilemedi ({e}); önceki kopya kullanılıyor.")
            return json.loads(onbellek.read_text(encoding="utf-8"))
        raise


def tbb_dipnotu(D: dict, tur: str, ad: str) -> dict | None:
    """TBB raporunun dipnot dosyasını (docs/dipnot/<id>.js) indirir; önbelleğe yazar."""
    dn = next((x for x in D[tur].get("DN", []) if x["ad"] == ad), None)
    if not dn:
        return None
    onbellek = KLASOR / f"tbb_dipnot_{dn['id']}.json"
    try:
        t = istek("GET", TBB_URL.replace("index.html", f"dipnot/{dn['id']}.js")).decode("utf-8")
        m = re.match(r'\s*__DN\("([^"]+)",(.*)\);\s*$', t, re.S)
        veri = json.loads(m.group(2))
        onbellek.write_text(json.dumps(veri), encoding="utf-8")
        return veri
    except Exception as e:
        if onbellek.exists():
            return json.loads(onbellek.read_text(encoding="utf-8"))
        print(f"      ! TBB dipnotu alınamadı ({ad}): {e}")
        return None


def dn_deger(d: dict, donem: str, kod, bs: int):
    """Dipnot verisi seyrek tutulur: v[dönem][kod] = [banka sırası, değer, …]"""
    v = d["v"].get(donem, {}).get(str(kod))
    if v is None:
        return None
    for i in range(0, len(v), 2):
        if v[i] == bs:
            return v[i + 1]
    return 0


def ad_hizala(a: list[str], b: list[str], esik: float) -> list:
    """İki ad listesini sırayı koruyarak en iyi hizalar; [(i, j|None, puan)]"""
    n, m = len(a), len(b)
    benz = [[0.0] * m for _ in range(n)]
    for i in range(n):
        for j in range(m):
            if a[i] and b[j]:
                r = difflib.SequenceMatcher(None, a[i], b[j]).ratio()
                benz[i][j] = r if r >= esik else 0.0
    P = [[0.0] * (m + 1) for _ in range(n + 1)]
    for i in range(n - 1, -1, -1):
        for j in range(m - 1, -1, -1):
            P[i][j] = max(P[i + 1][j], P[i][j + 1], (benz[i][j] + P[i + 1][j + 1]) if benz[i][j] else 0)
    out, i, j = [], 0, 0
    while i < n:
        if j < m and benz[i][j] and P[i][j] == benz[i][j] + P[i + 1][j + 1]:
            out.append((i, j, round(benz[i][j], 2)))
            i, j = i + 1, j + 1
        elif j < m and P[i][j] == P[i][j + 1]:
            j += 1
        else:
            out.append((i, None, 0.0))
            i += 1
    return out


KUR_ESANLAM = {"DIGER MEVDUAT": "DOVIZ TEVDIAT HESABI"}


def kur_net_satirlari(d: dict) -> dict:
    """Sitede net bilanço / net nazım pozisyonunun döviz cinsi kırılımları artık ayrı satır değil ('Net Genel Pozisyon'
    tek satır). PDF'teki döviz cinsine göre net satırlarla karşılaştırmak için kırılımlar kaynak satırlardan
    (toplam varlıklar − |toplam yükümlülükler|; türev alacakları − |türev borçları|) geçici olarak eklenir."""
    items = d["items"]
    if any(sade(a).startswith("NET BILANCO POZISYONU VARLIK") for _, a, *_ in items):
        return d                                                  # eski yapı: kırılımlar zaten var
    yig, cocuk, ad_of = [], defaultdict(list), {}
    for kod, ad, sv, *_ in items:
        ad_of[kod] = ad
        while yig and yig[-1][1] >= sv:
            yig.pop()
        if yig:
            cocuk[yig[-1][0]].append(kod)
        yig.append((kod, sv))
    bul = lambda ust, f: next((c for c in cocuk.get(ust, []) if f(sade(ad_of[c]))), None)
    kok = items[0][0]
    nb, nn = bul(kok, lambda n: n == "NET BILANCO POZISYONU"), bul(kok, lambda n: n == "NET NAZIM HESAP POZISYONU")
    if not (nb and nn):
        return d
    va, yu = bul(nb, lambda n: n == "VARLIKLAR"), bul(nb, lambda n: n == "YUKUMLULUKLER")
    tv = bul(va, lambda n: n.startswith("TOPLAM VARLIKLAR")) if va else None
    ty = bul(yu, lambda n: n.startswith("TOPLAM YUKUMLULUKLER")) if yu else None
    al, bo = bul(nn, lambda n: "TUREV" in n and "ALACAK" in n), bul(nn, lambda n: "TUREV" in n and "BORC" in n)
    if not (tv and ty and al and bo):
        return d
    yeni_items, V = list(items), {p: dict(x) for p, x in d["v"].items()}
    kod = 9580000
    for ust, adi, (a_, b_) in ((nb, "Net Bilanço Pozisyonu (varlıklar − yükümlülükler)", (tv, ty)),
                                (nn, "Net Nazım Hesap Pozisyonu (alacaklar − borçlar)", (al, bo))):
        kod += 1
        ana = kod
        yeni_items.append([ana, adi, 2, 0])
        for pb in ("EURO", "USD", "DIGER YP"):
            kod += 1
            yeni_items.append([kod, {"DIGER YP": "Diğer YP"}.get(pb, pb), 3, 0])
            ka = next((c for c in cocuk.get(a_, []) if sade(ad_of[c]) == pb), None)
            kb = next((c for c in cocuk.get(b_, []) if sade(ad_of[c]) == pb), None)
            for p, x in V.items():
                va_ = x.get(str(ka)) or []
                vb_ = x.get(str(kb)) or []
                A_ = {va_[i]: va_[i + 1] for i in range(0, len(va_), 2)}
                B_ = {vb_[i]: vb_[i + 1] for i in range(0, len(vb_), 2)}
                dd = {bk: (A_.get(bk) or 0) - abs(B_.get(bk) or 0) for bk in set(A_) | set(B_)}
                if dd:
                    x[str(kod)] = [z for bk, v in sorted(dd.items()) for z in (bk, v)]
    return {**d, "items": yeni_items, "v": V}


def kur_tbb_satirlari(d: dict) -> list[dict]:
    """TBB kur riski dipnotundaki satırlar: alt kalemleri EURO/USD/Diğer YP olan düğümler."""
    d = kur_net_satirlari(d)
    items = d["items"]
    yig, ebeveyn, cocuk, ad_of = [], {}, defaultdict(list), {}
    for kod, ad, sv, *_ in items:
        ad_of[kod] = ad
        while yig and yig[-1][1] >= sv:
            yig.pop()
        ebeveyn[kod] = yig[-1][0] if yig else None
        if yig:
            cocuk[yig[-1][0]].append(kod)
        yig.append((kod, sv))
    out = []
    for kod, ad, sv, *_ in items:
        ck = {sade(ad_of[c]): c for c in cocuk.get(kod, [])}
        if {"EURO", "USD", "DIGER YP"} <= set(ck) and not sade(ad).startswith("NET GENEL POZISYON"):
            x, bolum = ebeveyn.get(kod), "ozet"
            while x is not None:
                if sade(ad_of[x]) == "VARLIKLAR":
                    bolum = "varlik"
                if sade(ad_of[x]) == "YUKUMLULUKLER":
                    bolum = "yukum"
                x = ebeveyn.get(x)
            out.append({"kod": kod, "ad": ad, "bolum": bolum,
                        "sutun": [ck["EURO"], ck["USD"], ck["DIGER YP"], kod]})
    return out


class DipnotEsleyici:
    """Bir mali bünye tablosunun PDF satırlarını TBB dipnot kalemlerine eşler.
    ref(r) → (sütun başına TBB kodu / ("top", [kodlar]) / None, TBB kalem adı, yöntem)"""

    def __init__(self, spec: dict, d: dict):
        self.spec, self.d = spec, d
        self.dugum, self.cocuk, sira = {}, defaultdict(list), []
        yig = []
        for kod, ad, sv, *_ in d["items"]:
            while yig and yig[-1][1] >= sv:
                yig.pop()
            ebeveyn = yig[-1][0] if yig else None
            self.dugum[kod] = {"kod": kod, "ad": ad, "sv": sv, "ebeveyn": ebeveyn}
            if ebeveyn is not None:
                self.cocuk[ebeveyn].append(kod)
            sira.append(kod)
            yig.append((kod, sv))
        self.sira = sira
        e = spec["esle"]
        if e == "kanon":
            self.anahtar = {}
            for kod in sira:
                sut = {kanon_sutun(self.dugum[c]["ad"]): c for c in self.cocuk.get(kod, [])}
                sut.pop(None, None)
                if len(sut) >= 2:
                    k = kanon_kalem(self.dugum[kod]["ad"])
                    if k and k not in self.anahtar:
                        self.anahtar[k] = (kod, sut)
        elif e == "lko":
            bolum = {}
            for kod in sira:
                a = sade(self.dugum[kod]["ad"])
                if "UYGULANMAMIS" in a:
                    bolum.setdefault("un", kod)
                elif "UYGULANMIS TOPLAM" in a:
                    bolum.setdefault("ap", kod)
                elif "UST SINIR" in a:
                    bolum.setdefault("ust", kod)
            self.bolum = bolum
        elif e == "nifo":
            adaylar = [k for k in sira if self.dugum[k]["sv"] >= 2 and not re.match(
                r"^(KALAN VADE|DIKKATE ALMA|VADESIZ|6 AY|1 YIL)", sade(self.dugum[k]["ad"]))]
            self.no_dugum, j = {}, 0
            for n in sorted(NIFO_SATIR):
                bas = NIFO_SATIR[n]
                kalan = adaylar[j:]
                bul = next((k for k in kalan if sade(self.dugum[k]["ad"]) == bas), None) or \
                    next((k for k in kalan if sade(self.dugum[k]["ad"]).startswith(bas)), None)
                if bul is not None:
                    self.no_dugum[n] = bul
                    j = adaylar.index(bul) + 1
                else:   # toplam satırları (mevcut / gerekli istikrarlı fon) TBB'de kalemlerin üstündedir
                    bul = next((k for k in adaylar if sade(self.dugum[k]["ad"]).startswith(bas)), None)
                    if bul is not None and bul not in self.no_dugum.values():
                        self.no_dugum[n] = bul
        elif e == "sira":
            ust = next((k for k in sira if self.dugum[k]["sv"] == 1), None)
            self.sirali = self.cocuk.get(ust, [])

    def _bul(self, ebeveyn, bas):
        cc = self.cocuk.get(ebeveyn, [])
        return next((k for k in cc if sade(self.dugum[k]["ad"]) == bas), None) or \
            next((k for k in cc if sade(self.dugum[k]["ad"]).startswith(bas)), None)

    def _tp_yp(self, kod):
        cc = {kanon_sutun(self.dugum[c]["ad"]): c for c in self.cocuk.get(kod, [])}
        return [cc.get("tpyp"), cc.get("yp")]

    def ad(self, kod):
        return self.dugum[kod]["ad"] if kod in self.dugum else ""

    def ref(self, r: dict):
        e, n = self.spec["esle"], len(self.spec["sutun"])
        if e == "kanon":
            k = kanon_kalem(r["ad"])
            if k not in self.anahtar:
                return None, "", "eşleşmedi"
            kod, sut = self.anahtar[k]
            refs = [("top", list(sut.values())) if kk == "toplam" else sut.get(kk) for kk in self.spec["kanon"]]
            return refs, self.ad(kod), f"kalem anahtarı: {k}"
        if e == "lko":
            no = r.get("no")
            if no not in LKO_SATIR:
                return None, "", "eşleşmedi"
            bas = LKO_SATIR[no]
            if no >= 21:
                kod = self._bul(self.bolum.get("ust"), bas)
                return ([None, None] + self._tp_yp(kod) if kod else None), self.ad(kod), f"LKO satır {no}"
            un, ap = self._bul(self.bolum.get("un"), bas), self._bul(self.bolum.get("ap"), bas)
            if un is None and ap is None:
                return None, "", "eşleşmedi"
            refs = (self._tp_yp(un) if un else [None, None]) + (self._tp_yp(ap) if ap else [None, None])
            return refs, self.ad(ap or un), f"LKO satır {no}"
        if e == "nifo":
            kod = self.no_dugum.get(r.get("no"))
            if kod is None:
                return None, "", "TBB'de bu satır yok" if r.get("no") in (10, 25) else "eşleşmedi"
            cc = self.cocuk.get(kod, [])
            kv = next((c for c in cc if sade(self.dugum[c]["ad"]).startswith("KALAN VADE")), None)
            uy = next((c for c in cc if sade(self.dugum[c]["ad"]).startswith("DIKKATE ALMA ORANI UYGULANMIS")), None)
            if kv is None and uy is None:
                return [None] * (n - 1) + [kod], self.ad(kod), f"NİFO satır {r['no']} (yalnız d sütunu)"
            refs = [None] * n
            for c in self.cocuk.get(kv, []):
                a = sade(self.dugum[c]["ad"])
                j = 0 if a.startswith("VADESIZ") else 1 if a.startswith("6 AYDAN") else 2 if a.startswith("6 AY ILE") \
                    else 3 if a.startswith("1 YIL") else None
                if j is not None:
                    refs[j] = c
            dolu = [j for j in range(n - 1) if r["sayilar"][j] is not None]
            if kv is not None and len(dolu) == 1:       # birleşik hücre: tutar a–ç sütunlarının toplamı
                refs = [None] * n
                refs[dolu[0]] = kv
            refs[n - 1] = uy
            return refs, self.ad(kod), f"NİFO satır {r['no']}"
        if e == "sira":
            no = r.get("no")
            if not no or no > len(self.sirali):
                return None, "", "eşleşmedi"
            kod = self.sirali[no - 1]
            return [kod], self.ad(kod), f"şablon sırası {no}"
        return None, "", "TBB'de yok"

    def deger(self, donem: str, ref, bs: int):
        if ref is None:
            return None
        if isinstance(ref, tuple):
            v = [dn_deger(self.d, donem, k, bs) for k in ref[1]]
            v = [x for x in v if x is not None]
            return sum(v) if v else None
        return dn_deger(self.d, donem, ref, bs)


# ============================================================== öteki TBB dipnotları (genel eşleyici)
# Özel okuyucusu olan dipnotlar dışındaki bütün TBB dipnotları ve "8. Mevduat Vade Yapısı" tablosu,
# TBB'deki önceki dönem değerlerinin PDF'te aranmasıyla eşlenir (dipnot düzenleri bankadan bankaya değişir).
OZEL_DIPNOTLAR = {"Kur riskine ilişkin bilgiler"} | {sp["tbb"] for sp in DIPNOT_TABLOLARI if sp["tbb"]}
MEVDUAT_VADE = "8. Mevduat Vade Yapısı"
MEVDUAT_VADE_DN = "2.a.1.i) Mevduatın vade yapısına"   # aynı kodlu pasif dipnotu: özel okuyucuyla doldurulur
GD_TIRE = {"-", "–", "—"}
GD_REF = re.compile(r"^\(\d{1,2}\)$")


class GenelDipnot:
    """PDF'in bütün satırları (etiket + sayılar + dönem bağlamı) üzerinde TBB dipnot kalemlerini arar.
    Düzen A: TBB'de alt kalemleri sütun olan satırın değer dizisi PDF satırında bitişik aranır.
    Düzen B: tek TBB kalemi = PDF satırının bir sütunu ya da sütunlarının toplamı (ad benzerliği şartıyla).
    Cari dönem değeri: yan yana tabloda aynı satırın cari yarısı, alt alta tabloda cari bloktaki aynı sıradaki satır."""

    def __init__(self, ok: "Okuyucu"):
        self.ok = ok
        self.P = self._satirlar()
        self.indeks = defaultdict(set)              # yuvarlanmış |değer| → satır sırası
        for li, L in enumerate(self.P):
            if L["donem"] == "cari":
                continue
            for v in L["sayilar"]:
                self.indeks[round(abs(v))].add(li)
            n = len(L["sayilar"])
            if n >= 2:
                self.indeks[round(abs(sum(L["sayilar"])))].add(li)
            if n >= 4 and n % 2 == 0:                    # yan yana tablo: cari ve önceki yarıların toplamı
                self.indeks[round(abs(sum(L["sayilar"][:n // 2])))].add(li)
                self.indeks[round(abs(sum(L["sayilar"][n // 2:])))].add(li)

    def _sayi_mi(self, w):
        w2 = w.strip("%")
        return bool(w2) and (self.ok.sayi_mi(w2) or re.fullmatch(r"\(?-?\d+([.,]\d+)?\)?", w2) is not None)

    def _satirlar(self) -> list[dict]:
        d = pymupdf.open(self.ok.yol)
        out, donem, c_bekliyor = [], "?", False
        for pno in range(len(d)):
            gruplar = defaultdict(list)
            for x0, y0, x1, y1, w, *_ in d[pno].get_text("words"):
                gruplar[round((y0 + y1) / 2 / 2.5)].append((x0, w))
            onceki_etiket, son_bos = "", None
            for y in sorted(gruplar):
                ks = sorted(gruplar[y])
                metin = " ".join(w for _, w in ks)
                etiket_k, sayilar = [], []
                for i, (x, w) in enumerate(ks):
                    if (self._sayi_mi(w) or w in GD_TIRE) and not (i == 0 and re.fullmatch(r"\d{1,2}\.?", w)) \
                            and not (etiket_k and not sayilar and re.fullmatch(r"\d{1,2}", w) and x < 150) \
                            and not re.fullmatch(r"(19|20)\d\d", w.strip("()")):
                        sayilar.append((x, w))
                    else:
                        etiket_k.append(w)
                etiket = " ".join(etiket_k)
                sm = sade(metin)
                c, o = "CARI DONEM" in sm, "ONCEKI DONEM" in sm
                if sayilar and re.search(r"(CARI|ONCEKI) DONEM(IN)? (SONU|BASI|BAKIYE|KARI|ZARARI|NET)", sm):
                    c = o = False                     # satır adı ("Önceki Dönem Sonu Bakiyesi"), dönem başlığı değil
                if c and o:
                    donem = "yan"
                elif c:
                    donem = "cari"
                    c_bekliyor = not sayilar
                elif o:
                    # "Cari Dönem" ile "Önceki Dönem" arada rakamlı satır olmadan alt alta: yan yana tablo (Garanti)
                    donem = "yan" if c_bekliyor and not sayilar else "onceki"
                if sayilar:
                    c_bekliyor = False
                if not sayilar:
                    # etiketi rakamların biraz altına düşmüş satır (Akbank "Kambiyo İşlemlerinden Zarar"): hemen
                    # ardından gelen kısa, rakamsız metin o satırın etiketidir
                    if son_bos is not None and etiket and len(etiket) < 80:
                        son_bos["etiket"], son_bos = etiket, None
                        continue
                    onceki_etiket = etiket if len(etiket) < 120 else ""
                    son_bos = None
                    continue
                son_bos = None
                if not etiket and onceki_etiket:
                    etiket = onceki_etiket
                elif etiket and re.search(r"\b(ve|ile|veya)$", onceki_etiket.strip()):
                    # iki satıra bölünmüş etiket ("Tahsil İmkanı Sınırlı Krediler ve" / "Diğer Alacaklar İçin …")
                    etiket = onceki_etiket.strip() + " " + etiket
                onceki_etiket = ""
                if all(GD_REF.match(w) for _, w in sayilar) or (not etiket and all(w in GD_TIRE for _, w in sayilar)):
                    continue                                  # dipnot işareti ya da satırdan kopmuş tek "-"
                try:
                    deg = [0.0 if w in GD_TIRE else self.ok.sayi(w.strip("%")) for _, w in sayilar]
                except Exception:
                    continue
                out.append({"sayfa": pno + 1, "etiket": etiket, "donem": donem, "sayilar": deg,
                            "ham": [w for _, w in sayilar]})
                son_bos = out[-1] if not etiket else None
        # yüzde payı sütunları (tutar | % | tutar | % …, VakıfBank sektör dağılımı): ardışık ≥3 satırda tek sıralı
        # sütunlar hep iki ondalıklı ya da "-" ve ≤100 ise atılır
        def yuzde_mi(w, v):
            return w in GD_TIRE or (re.fullmatch(r"-?\d{1,3}[.,]\d{2}", w) is not None and abs(v) <= 100)
        i = 0
        while i < len(out):
            n = len(out[i]["sayilar"])
            j = i
            while j + 1 < len(out) and len(out[j + 1]["sayilar"]) == n and out[j + 1]["sayfa"] == out[i]["sayfa"]:
                j += 1
            blok = out[i:j + 1]
            if n >= 4 and n % 2 == 0 and len(blok) >= 3 and all(
                    all(yuzde_mi(x["ham"][q], x["sayilar"][q]) for q in range(1, n, 2)) for x in blok) and any(
                    re.fullmatch(r"-?\d{1,3}[.,]\d{2}", x["ham"][q]) for x in blok for q in range(1, n, 2)):
                for x in blok:
                    x["sayilar"], x["ham"] = x["sayilar"][0::2], x["ham"][0::2]
            i = j + 1
        # bloktaki olağan sütun sayısını aşan baştaki "(6)" gibi işaretleri at
        i = 0
        while i < len(out):
            j = i
            while j + 1 < len(out) and out[j + 1]["donem"] == out[i]["donem"]:
                j += 1
            mod = Counter(len(x["sayilar"]) for x in out[i:j + 1]).most_common(1)[0][0]
            for x in out[i:j + 1]:
                if len(x["sayilar"]) == mod + 1 and GD_REF.match(x["ham"][0]):
                    x["sayilar"], x["ham"] = x["sayilar"][1:], x["ham"][1:]
            i = j + 1
        return out

    # --- TBB ağacı
    @staticmethod
    def tbb_satirlari(d: dict) -> list[dict]:
        items = d["items"]
        ebeveyn, cocuk, yigin = {}, defaultdict(list), []
        for kod, ad, sv, *_ in items:
            while yigin and yigin[-1][1] >= sv:
                yigin.pop()
            ebeveyn[kod] = yigin[-1][0] if yigin else None
            if ebeveyn[kod] is not None:
                cocuk[ebeveyn[kod]].append(kod)
            yigin.append((kod, sv))
        ad_of = {k: a for k, a, *_ in items}

        def yolu(k):
            out, x = [], ebeveyn.get(k)
            while x is not None and ebeveyn.get(x) is not None:
                out.append(ad_of[x])
                x = ebeveyn[x]
            return list(reversed(out))
        satirlar = []
        for kod, ad, sv, *_ in items:
            ck = cocuk.get(kod, [])
            if ck and all(not cocuk.get(c) for c in ck):
                satirlar.append({"kod": kod, "ad": ad, "yol": yolu(kod), "sutunlar": [(c, ad_of[c]) for c in ck]})
        for kod, ad, sv, *_ in items:
            if not cocuk.get(kod) and ebeveyn.get(kod) is not None:
                if any(cocuk.get(k) for k in cocuk[ebeveyn[kod]]):
                    satirlar.append({"kod": kod, "ad": ad, "yol": yolu(kod), "sutunlar": [(kod, ad)], "tekil": True})
        return satirlar

    @staticmethod
    def _hizala(V, N, tol):
        en_iyi, m, k = None, len(V), len(N)
        if m == 0 or k < m:
            return None
        for o in range(k - m + 1):
            es = yanlis = sifir = 0
            for i, v in enumerate(V):
                if v is None:
                    continue
                n = N[o + i]
                if abs(v) < 1e-9:
                    if abs(n) < 1e-9:
                        sifir += 1
                    elif abs(n) > tol:
                        yanlis += 1
                elif abs(abs(n) - abs(v)) <= (1.5 * tol if tol >= 1 and abs(v) >= 1000 else tol):
                    # TBB tutarı TL'den bine yuvarlanırken bankanın kendi yuvarlamasıyla 1–2 birim ayrışabiliyor
                    es += 1
                else:
                    yanlis += 1
            if yanlis == 0 and es > 0:
                puan = es + 0.1 * sifir
                if en_iyi is None or puan > en_iyi[0]:
                    en_iyi = (puan, o)
        return en_iyi

    def _aday_satirlar(self, degerler):
        out = set()
        for v in degerler:
            if v:
                r = round(abs(v))
                for k in ((r - 2, r - 1, r, r + 1, r + 2) if r >= 1000 else (r - 1, r, r + 1)):
                    out |= self.indeks.get(k, set())
        return out

    def _blok(self, i):
        P, t = self.P, self.P[i]["donem"]
        s = e = i
        while s - 1 >= 0 and P[s - 1]["donem"] == t:
            s -= 1
        while e + 1 < len(P) and P[e + 1]["donem"] == t:
            e += 1
        return s, e

    def cari_satir(self, li):
        P = self.P
        so, eo = self._blok(li)
        k = li - so
        j = so - 1
        while j >= 0 and P[j]["donem"] == "?":
            j -= 1
        if j < 0 or P[j]["donem"] != "cari" or abs(P[j]["sayfa"] - P[li]["sayfa"]) > 4:
            return None
        sc, ec = self._blok(j)
        n = len(P[li]["sayilar"])
        aday = []
        ad = sade(P[li]["etiket"])
        if ad:
            ayni = [c for c in range(sc, ec + 1) if sade(P[c]["etiket"]) == ad]
            if ayni:   # aynı adlı cari satır varsa yalnız o; sütun sayısı tutmuyorsa karşılık yok sayılır
                ayni = [c for c in ayni if len(P[c]["sayilar"]) == n]
                return min(ayni, key=lambda c: abs((c - sc) - k)) if ayni else None
        if (eo - so) != (ec - sc) and P[li]["etiket"]:
            for c in range(sc, ec + 1):
                if len(P[c]["sayilar"]) == n:
                    b = benzer(P[c]["etiket"], P[li]["etiket"])
                    if b >= 0.85:
                        aday.append((b - 0.002 * abs((c - sc) / max(1, ec - sc) - k / max(1, eo - so)) * 100, c))
            return max(aday)[1] if aday else None
        hedef = sc + k
        for c in range(max(sc, hedef - 3), min(ec, hedef + 3) + 1):
            if len(P[c]["sayilar"]) != n:
                continue
            b = benzer(P[c]["etiket"], P[li]["etiket"]) if (P[c]["etiket"] or P[li]["etiket"]) else 1.0
            aday.append((b - 0.05 * abs(c - hedef), c, b))
        if not aday:
            return None
        _, c, b = max(aday)
        # iki satırın da adı varsa adlar neredeyse aynı olmalı (ör. "Yurt İçinde" ≠ "Yurt Dışında")
        if P[c]["etiket"] and P[li]["etiket"]:
            return c if b >= 0.85 else None
        return c if c == hedef else None

    # --- düzen A
    def _esle_a(self, d, tsat, bs, birim, donem):
        P, adaylar = self.P, []
        for r in tsat:
            ham = [dn_deger(d, donem, c, bs) for c, _ in r["sutunlar"]]
            if all(h is None for h in ham):
                continue
            oran = any(isinstance(h, float) and not float(h).is_integer() for h in ham if h is not None) or \
                "ORAN" in sade(" ".join(r["yol"] + [r["ad"]]))
            adaylar.append((r, ham, oran))
        sayfa_puan, on_es = defaultdict(float), []
        sayfa_satir = Counter(L["sayfa"] for L in P)
        for r, ham, oran in adaylar:
            for carp, tol in (((1, 0.011), (birim, 1.0)) if oran else ((birim, 1.0),)):
                V = [None if h is None else (round(h / carp) if carp != 1 else h) for h in ham]
                nz = sum(1 for v in V if v)
                if nz == 0:
                    continue
                for li in self._aday_satirlar(V):
                    L = P[li]
                    h = self._hizala(V, L["sayilar"], tol)
                    # etiketi ayrı satıra düşmüş (boş etiketli) satır: tek değer birebir ve büyükse kabul
                    bos_tam = (nz == 1 and not L["etiket"].strip() and
                               any(v and abs(v) >= 1000 and abs(abs(L["sayilar"][h[1] + i]) - abs(v)) <= 1
                                   for i, v in enumerate(V))) if h else False
                    if h and (h[0] >= 2 or (nz == 1 and benzer(r["ad"], L["etiket"]) >= 0.55) or bos_tam):
                        on_es.append((r, ham, carp, li, h))
                        sayfa_puan[L["sayfa"]] += h[0] / max(1, sayfa_satir[L["sayfa"]]) ** 0.5
        secim = {}
        for r, ham, carp, li, h in on_es:
            skor = h[0] + 2 * benzer(r["ad"], P[li]["etiket"]) + 0.2 * sayfa_puan[P[li]["sayfa"]]
            if r["kod"] not in secim or skor > secim[r["kod"]][0]:
                secim[r["kod"]] = (skor, r, ham, carp, li, h)
        kayit = []
        for kod, (skor, r, ham, carp, li, (puan, o)) in secim.items():
            L, m = P[li], len(r["sutunlar"])
            n = len(L["sayilar"])
            cari, kaynak = None, ""
            if L["donem"] == "yan" and n % 2 == 0 and o >= n // 2:
                cari, kaynak = L["sayilar"][o - n // 2: o - n // 2 + m], "aynı satır (yan yana)"
            else:
                cj = self.cari_satir(li)
                if cj is not None:
                    cari, kaynak = P[cj]["sayilar"][o:o + m], f"cari tablo s.{P[cj]['sayfa']}"
            for i, (c, cad) in enumerate(r["sutunlar"]):
                kayit.append({"kod": c, "satir": r["ad"], "yol": " › ".join(r["yol"]),
                              "sutun": "" if r.get("tekil") else cad,
                              "c": None if cari is None or i >= len(cari) else cari[i],
                              "o": L["sayilar"][o + i], "birim": carp, "sayfa": L["sayfa"],
                              "etiket": L["etiket"], "kaynak": kaynak, "duzen": "A"})
        return kayit

    # --- düzen B
    def _turet(self, L, cari):
        s, n = L["sayilar"], len(L["sayilar"])
        if L["donem"] == "yan" and n % 2 == 0 and n >= 2:
            h = n // 2
            parca, bas = (s[:h], 0) if cari else (s[h:], h)
        else:
            parca, bas = s, 0
        out = [("toplam", sum(parca), 0.5 * len(parca) + 0.5, ("sum", bas, len(parca)))] if len(parca) >= 2 else []
        out += [(f"sütun {i + 1}", v, 1.0, ("col", bas + i, 1)) for i, v in enumerate(parca)]
        if len(parca) >= 4 and len(parca) % 2 == 0:              # iki grup × (TP, YP): I. grup / II. grup (2.a.3.iii)
            h = len(parca) // 2
            out += [("yarı 1", sum(parca[:h]), 0.5 * h + 0.5, ("sum", bas, h)),
                    ("yarı 2", sum(parca[h:]), 0.5 * h + 0.5, ("sum", bas + h, h))]
        return out

    @staticmethod
    def _uygula(L, dn):
        tur, bas, k = dn
        s = L["sayilar"]
        if bas < 0 or bas + k > len(s):
            return None
        return sum(s[bas:bas + k]) if tur == "sum" else s[bas]

    def _esle_b(self, d, bs, birim, donem, zaten):
        P, kayit = self.P, []
        ad_of = {k: a for k, a, *_ in d["items"]}
        for kod, ad, sv, *_ in d["items"]:
            if kod in zaten or sv == 0:
                continue
            v = dn_deger(d, donem, kod, bs)
            if not v:
                continue
            hedef = v / birim
            adaylar = []
            for li in self._aday_satirlar([hedef]):
                L = P[li]
                b = benzer(ad, L["etiket"])
                for nm, val, tol, dn in self._turet(L, cari=False):
                    if abs(abs(val) - abs(hedef)) <= (1.5 * tol if tol >= 1 and abs(hedef) >= 1000 else tol):
                        adaylar.append((b, li, nm, dn))
            if not adaylar:
                continue
            adaylar.sort(key=lambda a: (a[0], -a[1]), reverse=True)
            b, li, nm, dn = adaylar[0]
            tekil = len({a[1] for a in adaylar}) == 1
            if not (b >= 0.6 or (b >= 0.5 and tekil and abs(hedef) >= 1000)):
                continue
            L = P[li]
            cari, kaynak = None, ""
            if L["donem"] == "yan" and len(L["sayilar"]) % 2 == 0:
                h = len(L["sayilar"]) // 2
                cari, kaynak = self._uygula(L, (dn[0], dn[1] - h, dn[2])), "aynı satır (yan yana)"
            else:
                cj = self.cari_satir(li)
                if cj is not None:
                    cari, kaynak = self._uygula(P[cj], dn), f"cari tablo s.{P[cj]['sayfa']}"
            kayit.append({"kod": kod, "satir": ad, "yol": "", "sutun": nm, "c": cari, "o": self._uygula(L, dn),
                          "birim": birim, "sayfa": L["sayfa"], "etiket": L["etiket"], "kaynak": kaynak,
                          "duzen": "B"})
        return kayit

    def esle(self, d: dict, bs: int, birim: int, donemler) -> dict:
        """Önceki dönem adayları içinde en çok kalem bulunanı seçer."""
        tsat = self.tbb_satirlari(d)
        en_iyi = None
        for donem in donemler:
            if donem not in d["v"]:
                continue
            a = self._esle_a(d, tsat, bs, birim, donem)
            b = self._esle_b(d, bs, birim, donem, {k["kod"] for k in a})
            sayfalar = sorted(k["sayfa"] for k in a + b)
            if len(sayfalar) >= 3:
                orta = sayfalar[len(sayfalar) // 2]
                b = [k for k in b if abs(k["sayfa"] - orta) <= 3]
            kayit = a + b
            self._cari_tamamla(kayit)
            dolu = [k for k, *_ in d["items"] if dn_deger(d, donem, k, bs)]
            sonuc = {"donem": donem, "kayit": kayit, "dolu": dolu}
            if en_iyi is None or len(kayit) > len(en_iyi["kayit"]):
                en_iyi = sonuc
        if en_iyi and not en_iyi["kayit"]:
            for donem in donemler:
                if donem not in d["v"]:
                    continue
                aday = {"donem": donem, "kayit": [], "dolu": [k for k, *_ in d["items"] if dn_deger(d, donem, k, bs)]}
                self.ad_esle_yakin(d, bs, birim, aday)
                if len(aday["kayit"]) > len(en_iyi["kayit"]):
                    en_iyi = aday
        if en_iyi:
            self.ad_tamamla(d, en_iyi, bs)
        return en_iyi or {"donem": None, "kayit": [], "dolu": []}

    def _cari_tamamla(self, kayit):
        """Cari değeri bulunamayan kayıt (aynı rakamlar önceki dönem bloğundaki başka bir tabloda da geçiyor – Garanti
        donuk alacak hareketi): aynı adlı, önceki değeri aynı olan yan yana (cari | önceki) satırdan cari alınır."""
        for k in kayit:
            if k["c"] is not None or k["o"] is None:
                continue
            ad = k["satir"]
            for li in self._aday_satirlar([k["o"]]):
                L = self.P[li]
                n = len(L["sayilar"])
                if L["donem"] != "yan" or n % 2 or max(benzer(L["etiket"], ad), benzer(L["etiket"], k.get("sutun") or "")) < 0.8:
                    continue
                h = n // 2
                for i in range(h, n):
                    if abs(L["sayilar"][i] - k["o"]) <= 0.5:
                        k.update(c=L["sayilar"][i - h], sayfa=L["sayfa"], etiket=L["etiket"], kaynak="aynı satır (yan yana)")
                        break
                if k["c"] is not None:
                    break

    def ad_esle_yakin(self, d: dict, bs: int, birim: int, sonuc: dict):
        """Hiçbir kalem değerle tutmadıysa (banka önceki dönemi baştan yeniden düzenlemiş – Garanti konsolide):
        adı TBB kalemine çok benzeyen ve önceki dönem değeri TBB'ye %3 yakın en az iki satır aynı sayfadaysa
        bu satırlar çapa olur; cari değer önceki değerin sütun karşılığından alınır."""
        donem, P = sonuc["donem"], self.P
        items = d["items"]
        cocuklu = {items[i][0] for i in range(len(items) - 1) if items[i + 1][2] > items[i][2]}
        cifti = {}                                             # önceki fonksiyonu → cari fonksiyonu

        def fonklar(n):
            out = []
            h = n // 2
            for i in range(h, n):
                out.append((f"s{i}", lambda x, i=i: x[i], lambda x, i=i, h=h: x[i - h]))
            if n >= 4 and n % 2 == 0:
                out.append(("y2", lambda x, h=h: sum(x[h:]), lambda x, h=h: sum(x[:h])))
            return out
        adaylar = []
        for kod, ad, sv, *_ in items:
            tv = dn_deger(d, donem, kod, bs)
            a = sade(ad)
            if not tv or sv < 1 or len(a) < 6 or a in ("DIGER", "TOPLAM", "KAR", "ZARAR"):
                continue
            hedef = abs(tv / birim)
            for li, L in enumerate(P):
                n = len(L["sayilar"])
                if n < 2 or n % 2 or L["donem"] == "cari":
                    continue
                b_ = benzer(L["etiket"], ad)
                if b_ < 0.88:
                    continue
                for fad, fo, fc in fonklar(n):
                    v = abs(fo(L["sayilar"]))
                    if hedef > 0 and abs(v - hedef) <= 0.03 * hedef:
                        adaylar.append((kod, ad, li, fad, fo, fc, b_))
                        break
        sayfa_say = Counter(P[x[2]]["sayfa"] for x in adaylar)
        if not sayfa_say:
            return
        sayfa, n_ = sayfa_say.most_common(1)[0]
        if n_ < 2:
            return
        goruldu = set()
        for kod, ad, li, fad, fo, fc, b_ in sorted(adaylar, key=lambda x: -x[6]):
            L = P[li]
            if L["sayfa"] != sayfa or kod in goruldu or li in goruldu:
                continue
            goruldu.update({kod, li})
            sonuc["kayit"].append({"kod": kod, "satir": ad, "yol": "", "sutun": "", "c": fc(L["sayilar"]),
                                   "o": fo(L["sayilar"]), "birim": birim, "sayfa": L["sayfa"], "etiket": L["etiket"],
                                   "kaynak": "ad + yakın değer (önceki dönem yeniden düzenlenmiş)", "duzen": "ad"})

    def ad_tamamla(self, d: dict, sonuc: dict, bs: int = None):
        """Değerle eşleşmeyen TBB kalemleri: aynı tablodaki satırlarda adla aranır (banka önceki dönemi yeniden
        düzenlemişse önceki dönem değeri TBB'yle tutmaz – Garanti). Cari/önceki değer, eşleşen komşu satırla aynı
        sütun kuralıyla (tek sütun, yarıların toplamı …) alınır."""
        kayit = sonuc["kayit"]
        if not kayit:
            return
        P = self.P
        fonk = []
        for n in range(1, 9):
            for i in range(n):
                fonk.append((n, f"s{i}", lambda x, i=i: x[i]))
            if n % 2 == 0 and n >= 2:
                h = n // 2
                fonk += [(n, "y1", lambda x, h=h: sum(x[:h])), (n, "y2", lambda x, h=h: sum(x[h:]))]
            fonk.append((n, "t", lambda x: sum(x)))
        kural, kullanilan, sayfalar, satir_of = Counter(), set(), set(), {}
        for k in kayit:
            if k.get("sutun") not in ("", None, "toplam") or k["c"] is None or k["o"] is None:
                continue
            for li, L in enumerate(P):
                if L["sayfa"] == k["sayfa"] and L["etiket"] == k["etiket"]:
                    x = L["sayilar"]
                    fc = [ad for n, ad, f in fonk if n == len(x) and abs(f(x) - k["c"]) <= 0.5]
                    fo = [ad for n, ad, f in fonk if n == len(x) and abs(f(x) - k["o"]) <= 0.5]
                    for a_ in fc:
                        for b_ in fo:
                            kural[(len(x), a_, b_)] += 1
                    if fc and fo:
                        kullanilan.add(li)
                        sayfalar.add(L["sayfa"])
                        satir_of[id(k)] = x
                    break
        if not kural:
            return
        # en çok kayıtla tutan kural; eşitlikte yarı toplamları (TP+YP) tek sütuna yeğlenir
        (n, fc, fo), _ = max(kural.items(), key=lambda q: (q[1], q[0][1] in ("y1", "y2") and q[0][2] in ("y1", "y2")))
        f_of = {ad: f for m, ad, f in fonk if m == n}
        for k in kayit:                     # önceki değeri kurala uyan kaydın cari değeri de kuralla (TP 0 iken YP seçilmiş)
            x = satir_of.get(id(k))
            if x is not None and len(x) == n and abs(f_of[fo](x) - k["o"]) <= 0.5 and not k.get("sutun"):
                k["c"] = f_of[fc](x)
        var = {k["kod"] for k in kayit}
        items = d["items"]
        cocuklu = {items[i][0] for i in range(len(items) - 1) if items[i + 1][2] > items[i][2]}
        for kod, ad, sv, *_ in items:
            a = sade(ad)
            if kod in var or kod in cocuklu or sv < 1 or len(a) < 6 or a in ("DIGER", "TOPLAM"):
                continue
            aday = []
            for li, L in enumerate(P):
                if li in kullanilan or len(L["sayilar"]) != n or not (min(sayfalar) - 1 <= L["sayfa"] <= max(sayfalar) + 1):
                    continue
                b_ = benzer(L["etiket"], ad)
                if b_ >= 0.9:
                    aday.append((b_, li))
            if not aday:
                continue
            aday.sort(reverse=True)
            if len(aday) > 1 and aday[1][0] >= aday[0][0] - 0.02:
                continue                                         # belirsiz
            li = aday[0][1]
            L = P[li]
            tv = dn_deger(d, sonuc["donem"], kod, bs) if bs is not None else None
            if tv:                                            # önceki dönem TBB'den %25'ten çok sapıyorsa başka kalem
                o_ = f_of[fo](L["sayilar"]) * kayit[0]["birim"]
                if abs(abs(o_) - abs(tv)) > 0.25 * abs(tv):
                    continue
            kullanilan.add(li)
            kayit.append({"kod": kod, "satir": ad, "yol": "", "sutun": "", "c": f_of[fc](L["sayilar"]),
                          "o": f_of[fo](L["sayilar"]), "birim": kayit[0]["birim"], "sayfa": L["sayfa"], "etiket": L["etiket"],
                          "kaynak": "ad eşlemesi (önceki dönem TBB'yle tutmuyor)", "duzen": "ad"})


def mevduat_vade_dipnotu(D: dict, tur: str) -> dict | None:
    """'8. Mevduat Vade Yapısı' ana tablosunu dipnot biçimine çevirir (seyrek v)."""
    items = [it for it in D[tur]["items"] if it[3] == MEVDUAT_VADE]
    if not items:
        return None
    A, v = D[tur]["A"], {}
    for donem, tablo in A["v"].items():
        dd = {}
        for kod, *_ in items:
            dizi = tablo.get(str(kod))
            if dizi is not None:
                dd[str(kod)] = [x for i, val in enumerate(dizi) if val is not None for x in (i, val)]
        v[donem] = dd
    return {"items": [[k, a, sv] for k, a, sv, *_ in items], "v": v}


# ---------------------------------------------------------------------------
# Mevduatın vade yapısı (TBB "8. Mevduat Vade Yapısı") – özel okuyucu
MV_SUTUN_DESEN = [   # (anahtar, başlık kelimesi deseni)
    ("vadesiz", r"^VADESIZ$"), ("ihbar", r"^IHBARLI$"), ("1aya", r"^KADAR$"),
    ("1_3", r"^1-3$"), ("3_6", r"^3-6$"), ("6_12", r"^AY-1$"), ("1yil", r"^(USTU|UZERI|UZUN)$"),
    ("birikimli", r"^BIRIKIMLI$"), ("toplam", r"^TOPLAM$"),
]
MV_TBB_SUTUN = {"VADESIZ": "vadesiz", "7 GUN IHBARLI": "ihbar", "1AYA KADAR": "1aya", "1 AYA KADAR": "1aya",
                "1-3 AY": "1_3", "3-6 AY": "3_6", "6 AY-1 YIL": "6_12", "1 YIL VE USTU": "1yil",
                "BIRIKIMLI MEVDUAT": "birikimli"}
MV_SATIR = [   # (anahtar, PDF satır adı deseni, TBB kalem adı deseni)
    ("tas", r"TASARRUF", r"^TASARRUF"),
    ("dth", r"DOVIZ TEVDIAT", r"^DOVIZ TEVDIAT"),
    ("dth_yi", r"^YURT ?ICI(NDE)? YER", r"^YURTICINDE YERLESIK"),
    ("dth_yd", r"^YURT ?DISI(NDA)? YER", r"^YURTDISINDA YERLESIK"),
    ("resmi", r"^RESM", r"^RESMI KUR"),
    ("ticari", r"^TIC", r"^TIC\.? KUR"),
    ("diger_kur", r"^DIG(ER)?\.? KUR", r"^DIG\.? KUR"),
    ("kiymetli", r"^KIYMETLI", r"^KIYMETLI"),
    ("bank", r"^BANKALAR(ARASI)? ?MEVDUAT|^BANKALARARASI$|^BANKALAR MEVDUATI", r"^BANKALAR MEVDUATI"),
    ("b_tcmb", r"^(TCMB|T ?C ?MERKEZ|MERKEZ BANKASI)", r"^TC MERKEZ"),
    ("b_yi", r"^YURT ?ICI BANKALAR", r"^YURTICI BANKALAR"),
    ("b_yd", r"^YURT ?DISI BANKALAR", r"^YURTDISI BANKALAR"),
    ("b_kat", r"^KATILIM", r"^KATILIM"),
    ("b_diger", r"^DIGER$", r"^DIGER$"),
    ("toplam", r"^TOPLAM", r"^TOPLAM MEVDUAT"),
]


def mv_kod_haritasi(items) -> dict:
    """TBB kalemlerinden {(satır anahtarı, sütun anahtarı): kod}; satırın 'toplam' sütunu kalemin kendisi."""
    harita, satir = {}, None
    for kod, ad, sv, *_ in items:
        sa = sade(ad).replace(".", "")
        sa2 = re.sub(r"\s+", " ", sa).strip()
        sut = {re.sub(r"\s+", " ", sade(k_).replace(".", "")).strip(): v_ for k_, v_ in MV_TBB_SUTUN.items()}.get(sa2)
        if sut and satir:
            harita[(satir, sut)] = kod
            continue
        satir = None
        for anahtar, _, desen in MV_SATIR:
            if re.search(desen, sa2):
                satir = anahtar
                harita[(anahtar, "toplam")] = kod
                break
    return harita


def mevduat_vade_oku(ok: "Okuyucu") -> dict:
    """Mevduatın vade yapısı tablosu: {'cari': {(satır, sütun): değer}, 'onceki': {...}, 'sayfa': n}."""
    d = pymupdf.open(ok.yol)
    sonuc = {"cari": {}, "onceki": {}, "sayfa": None, "etiket": {}}
    sayfalar = []
    for i in range(len(d)):
        st = sade(d[i].get_text())
        if "BIRIKIMLI" in st and "TASARRUF" in st and "VADESIZ" in st and "KIYMETLI" in st:
            sayfalar.append(i)
    adaylar = [[i, i + 1] if i + 1 < len(d) else [i] for i in sayfalar]
    for sayfa_grubu in adaylar:
      donem, merkez, bekleyen, okundu, bitti = None, None, None, False, False
      sonuc["cari"], sonuc["onceki"], sonuc["etiket"] = {}, {}, {}
      for i in sayfa_grubu:
        if bitti or (i != sayfa_grubu[0] and not okundu):
            break
        W = d[i].rect.width
        satirlar = ok._satirlar_xy(d[i])
        for j, kel in enumerate(satirlar):
            kelimeler = [w for _, _, w, _ in kel]
            sm = sade(" ".join(kelimeler))
            esik = min(W * 0.24, min(merkez.values()) - W * 0.06) if merkez else W * 0.24
            rakam = [(x0, x1, w) for x0, x1, w, _ in kel
                     if x0 > esik and (ok.sayi_mi(w) or w in ("-", "–"))]
            if "VADESIZ" in sm and not any(len(w) > 2 for _, _, w in rakam):
                # sütun merkezleri: bu satır ve yakınındaki rakamsız başlık satırları
                yakin = [satirlar[x] for x in range(max(0, j - 2), min(len(satirlar), j + 3))]
                merkez = {}
                for s2 in yakin:
                    if any(ok.sayi_mi(w) and len(w) > 3 for _, _, w, _ in s2):
                        continue
                    for x0, x1, w, _ in s2:
                        ws = sade(w.replace("-", "TIRE")).replace("TIRE", "-").replace(" ", "")
                        for anahtar, desen in MV_SUTUN_DESEN:
                            if re.match(desen, ws) and (anahtar not in merkez or anahtar == "1aya" and False):
                                merkez.setdefault(anahtar, (x0 + x1) / 2)
                if "ihbar" not in merkez:
                    for s2 in yakin:
                        for a, b in zip(s2, s2[1:]):
                            if a[2] == "7" and sade(b[2]) == "GUN":
                                merkez["ihbar"] = (a[0] + b[1]) / 2
                if "6_12" not in merkez and "3_6" in merkez and "1yil" in merkez:
                    merkez["6_12"] = (merkez["3_6"] + merkez["1yil"]) / 2   # "6 Ay - 1 Yıl" ayrık yazılmış
                if re.search(r"ONCEKI DONEM|ARALIK", sm):
                    donem = "onceki"
                elif re.search(r"CARI DONEM|EYLUL|HAZIRAN|MART", sm) or donem is None:
                    donem = "cari" if donem is None else donem
                continue
            if re.match(r"(\d+(\.\d+)*\.?\s*)?ONCEKI DONEM", sm):
                donem = "onceki"
                continue
            if re.match(r"(\d+(\.\d+)*\.?\s*)?CARI DONEM", sm) and not rakam:
                donem = "cari"
                continue
            if merkez is None or donem is None:
                continue
            etiket = " ".join(w for x0, _, w, _ in kel if x0 <= esik or not (ok.sayi_mi(w) or w in ("-", "–")))
            es = sade(re.sub(r"\(\*+\)|\(\d\)", "", etiket)).replace(".", "").strip()
            es = re.sub(r"\s+", " ", es)
            if not rakam:
                bekleyen = es if es and len(es) <= 40 else None    # çok satırlı ad: "Döviz Tevdiat / Hesabı"
                continue
            anahtar = None
            for ad in [x for x in (es, ((bekleyen or "") + " " + es).strip(), bekleyen) if x]:
                anahtar = next((a for a, desen, _ in MV_SATIR if re.search(desen, ad)), None)
                if anahtar:
                    break
            bekleyen = None
            if anahtar is None:
                continue
            for x0, x1, w in rakam:
                xc = (x0 + x1) / 2
                sut = min(merkez, key=lambda k: abs(merkez[k] - xc))
                v = ok.sayi(w)
                if (anahtar, sut) not in sonuc[donem]:
                    sonuc[donem][(anahtar, sut)] = v
                if not sonuc.get("_bitti_" + donem):
                    sonuc.setdefault("_x1", []).append((donem, anahtar, sut, x1, v))
            sonuc["etiket"].setdefault((donem, anahtar), etiket)
            okundu = True
            if anahtar == "toplam" and donem == "onceki":
                bitti = True
                break
      if okundu:
          sonuc["sayfa"] = sayfa_grubu[0] + 1
          if sonuc["onceki"] and sonuc["cari"]:
              break
    # rakamlar sağa yaslı: bir sütunda tek kalmış rakamın sağ kenarı başka bir sütunun rakamlarıyla hizalıysa o sütuna
    # aittir (İş: dar "7 Gün İhbarlı" sütununa düşen vadesiz "99")
    x1ler = defaultdict(list)
    kayitlar = sonuc.pop("_x1", [])
    for donem, anahtar, sut, x1, v in kayitlar:
        x1ler[(donem, sut)].append(x1)
    med = {k_: sorted(v_)[len(v_) // 2] for k_, v_ in x1ler.items() if len(v_) >= 3}
    if med:
        yeni = {"cari": {}, "onceki": {}}
        for donem, anahtar, sut, x1, v in kayitlar:
            hedef = sut
            m = med.get((donem, sut))
            if m is None or abs(x1 - m) > 4:
                aday = [s_ for (d_, s_), m_ in med.items() if d_ == donem and abs(m_ - x1) <= 3]
                if len(aday) == 1:
                    hedef = aday[0]
            if (anahtar, hedef) not in yeni[donem]:
                yeni[donem][(anahtar, hedef)] = v
            elif (anahtar, sut) not in yeni[donem]:
                yeni[donem][(anahtar, sut)] = v
        for donem in ("cari", "onceki"):
            anahtarlar = {a_ for a_, _ in yeni[donem]}
            sonuc[donem] = {k_: v_ for k_, v_ in sonuc[donem].items() if k_[0] not in anahtarlar}
            sonuc[donem].update(yeni[donem])
    return sonuc


MV_SUTUN_AD = {"vadesiz": "Vadesiz", "ihbar": "7 Gün İhbarlı", "1aya": "1 Aya Kadar", "1_3": "1-3 Ay", "3_6": "3-6 Ay",
               "6_12": "6 Ay-1 Yıl", "1yil": "1 Yıl ve Üstü", "birikimli": "Birikimli Mevduat", "toplam": "Toplam"}


def mevduat_vade_kayit(ok: "Okuyucu", d: dict, bs: int, donemler) -> dict:
    """Özel okuyucunun sonucunu genel eşleyicinin kayıt biçimine çevirir (önceki dönem = geçen yıl sonu)."""
    r = getattr(ok, "_mv_sonuc", None)
    if r is None:
        r = ok._mv_sonuc = mevduat_vade_oku(ok)
    donem = next((x for x in donemler if x in d["v"]), None)
    ad_of = {k: a for k, a, *_ in d["items"]}
    harita = mv_kod_haritasi(d["items"])
    kayit = []
    for (satir, sut), kod in harita.items():
        o, c = r["onceki"].get((satir, sut)), r["cari"].get((satir, sut))
        if o is None and c is None:
            continue
        kayit.append({"kod": kod, "satir": ad_of.get(harita.get((satir, "toplam")), satir),
                      "yol": "", "sutun": MV_SUTUN_AD.get(sut, sut), "c": c, "o": o, "birim": ok.carpan,
                      "sayfa": r["sayfa"] or 0, "etiket": r["etiket"].get(("cari", satir), ""),
                      "kaynak": "mevduat vade yapısı tablosu", "duzen": "M"})
    dolu = [k for k, *_ in d["items"] if donem and dn_deger(d, donem, k, bs)]
    return {"donem": donem, "kayit": kayit, "dolu": dolu}


# ---------------------------------------------------------------------------
# Risk grubu dipnotları (RİSK-1 a, c.1, c.2) – özel okuyucu
RG_DIPNOTLAR = ("RİSK-1. a)", "RİSK-1. c.1)", "RİSK-1. c.2)")


def rg_satir(sa: str):
    if re.search(r"DONEM BASI", sa):
        return "bas"
    if re.search(r"DONEM SONU", sa):
        return "son"
    if re.search(r"ALINAN FAIZ|MEVDUAT FAIZ GIDERI|FAIZ GIDERI|^KOMISYON GELIRLERI", sa):
        return "faiz"
    if re.search(r"TOPLAM KAR|KAR ?/ ?\(?ZARAR|KAR ZARAR", sa):
        return "kar"
    return None


def rg_kod_haritasi(items, tablo: str) -> dict:
    """TBB kalemleri → {(grup, alt, satır): kod}. alt: a) için NAKDI/GNAKDI, c.2) için GUD/RK, c.1) için ''."""
    harita, grup, alt = {}, -1, ""
    grup_sv = None
    for kod, ad, sv, *_ in items:
        sa = sade(ad).replace(".", " ")
        sa = re.sub(r"\s+", " ", sa).strip()
        if re.match(r"^(ISTIRAK|BANKANIN DOGRUDAN|RISK GR)", sa) and "DAHIL OLDUGU RISK GRUBU" not in sa:
            grup += 1
            grup_sv, alt = sv, ""
            continue
        if grup < 0:
            continue
        if sa in ("NAKDI",):
            alt = "NAKDI"; continue
        if sa in ("G NAKDI", "GNAKDI", "GAYRINAKDI"):
            alt = "GNAKDI"; continue
        if sa.startswith("GERCEGE UYGUN"):
            alt = "GUD"; continue
        if sa.startswith("RISKTEN KORUNMA"):
            alt = "RK"; continue
        s = rg_satir(sa)
        if s:
            harita[(grup, alt, s)] = kod
    return harita


def risk_grubu_oku(ok: "Okuyucu") -> dict:
    """{'a': {(dönem, grup, NAKDI|GNAKDI, satır): değer}, 'c1': {(dönem, grup, '', satır)}, 'c2': {(dönem, grup, GUD|RK, satır)}}"""
    d = pymupdf.open(ok.yol)
    sonuc = {"a": {}, "c1": {}, "c2": {}, "sayfa": None}
    basla = None
    for i in range(len(d)):
        st = sade(d[i].get_text())
        if re.search(r"RISK GRUBU? ?(ILE|NA|NUN)?", st) and "DONEM BASI" in st and re.search(r"NAKDI", st):
            basla = i
            break
    if basla is None:
        return sonuc
    sonuc["sayfa"] = basla + 1
    tablo, donem, alt, yan_yana, bitti = None, "cari", "", False, False
    merkez = []                                    # sütun merkezleri (6'lı satırlardan)
    son_etiket = ""                                # rakamsız son satır ("Alınan Faiz ve" / rakamlar / "Komisyon …")
    for i in range(basla, min(len(d), basla + 4)):
        if bitti:
            break
        W = d[i].rect.width
        for kel in ok._satirlar_xy(d[i]):
            kelimeler = [w for _, _, w, _ in kel]
            sm = sade(" ".join(kelimeler))
            # satır sonundaki sayı dizisi (ilk rakam sütunu bankaya göre %30–%45 arasında)
            n_r = 0
            for _, _, w, _ in reversed(kel):
                if ok.sayi_mi(w) or w in ("-", "–"):
                    n_r += 1
                else:
                    break
            while n_r and (kel[len(kel) - n_r][0] < W * 0.25 or re.fullmatch(r"\(\d\)", kel[len(kel) - n_r][2])):
                n_r -= 1                                  # sol kenardaki madde no / "(3)" dipnot atfı rakam değil
            rakam = [((x0 + x1) / 2, w) for x0, x1, w, _ in kel[len(kel) - n_r:]] if n_r else []
            etiket = sade(" ".join(w for _, _, w, _ in kel[:len(kel) - n_r]))
            # sütun başlıkları: Nakdi/G.Nakdi ya da Cari/Önceki altı kez → sütun merkezleri
            bas_k = []
            for q_, (x0, x1, w, _) in enumerate(kel):
                if re.fullmatch(r"(G\.?)?NAKDI|CARI|ONCEKI", sade(w).replace(" ", "")):
                    if q_ + 1 < len(kel) and sade(kel[q_ + 1][2]) == "DONEM":
                        x1 = kel[q_ + 1][1]                # "Cari Dönem" bütün olarak ortalanır
                    bas_k.append((x0 + x1) / 2)
            if not rakam and len(bas_k) == 6:
                merkez = bas_k
            if re.search(r"UST YONETIM|YONETIM KURULU|BANKA UST", sm) and tablo:
                bitti = True
                break
            # tablo ve dönem başlıkları
            if not rakam:
                son_etiket = etiket
                if re.search(r"MEVDUATA? (ILISKIN|AIT)|AIT MEVDUAT", sm) or re.fullmatch(r"MEVDUAT( \S+)*", etiket or "x") and "CARI" in sm:
                    if tablo != "c1":
                        merkez = []
                    tablo, alt = "c1", ""
                if re.search(r"VADELI ISLEM|OPSIYON SOZLESME", sm):
                    tablo, alt, merkez = "c2", "", []
                if re.search(r"KREDILER(E)? (VE DIGER ALACAK|ILISKIN)|AIT KREDI", sm) or etiket in ("KREDILER", "KREDILER VE DIGER ALACAKLAR"):
                    if tablo != "a":
                        tablo = "a"
                if tablo is None and re.search(r"\bNAKDI\b", sm):
                    tablo = "a"
                cari_v, onc_v = "CARI DONEM" in sm or re.search(r"\bCARI\b", sm), "ONCEKI DONEM" in sm or re.search(r"\bONCEKI\b", sm)
                if cari_v and onc_v:
                    yan_yana = True
                elif onc_v and not cari_v:
                    donem, yan_yana = "onceki", False
                elif cari_v and not onc_v:
                    donem, yan_yana = "cari", False
                if etiket.startswith("GERCEGE UYGUN"):
                    alt = "GUD"
                elif etiket.startswith("RISKTEN KORUNMA"):
                    alt = "RK"
                continue
            if tablo is None:
                continue
            if etiket.startswith("GERCEGE UYGUN"):
                alt = "GUD"
            elif etiket.startswith("RISKTEN KORUNMA"):
                alt = "RK"
            s = rg_satir(etiket) or (rg_satir((son_etiket + " " + etiket).strip()) if son_etiket else None)
            son_etiket = ""
            if s is None:
                continue
            if len(rakam) == 6:
                merkez = [x for x, _ in rakam]
            if len(rakam) == 6 or not merkez:
                yerler = list(range(len(rakam)))
            else:                                     # boş hücreli satır: en yakın sütuna
                yerler = [min(range(6), key=lambda q: abs(merkez[q] - x)) for x, _ in rakam]
            for (x, w), q in zip(rakam, yerler):
                v = 0.0 if w in ("-", "–") else ok.sayi(w)
                if tablo == "a":
                    if yan_yana:
                        continue
                    sonuc["a"][(donem, q // 2, "NAKDI" if q % 2 == 0 else "GNAKDI", s)] = v
                else:
                    dn = ("cari" if q % 2 == 0 else "onceki") if yan_yana else donem
                    g = q // 2 if yan_yana else q
                    sonuc[tablo][(dn, g, alt if tablo == "c2" else "", s)] = v
            if tablo == "c2" and alt == "RK" and s == "kar":
                bitti = True
                break
    return sonuc


def risk_grubu_kayit(ok: "Okuyucu", ad: str, d: dict, bs: int, donemler) -> dict:
    """Özel okuyucunun sonucunu genel eşleyicinin kayıt biçimine çevirir. Bakiyeler önceki yıl sonuyla,
    faiz ve kâr/zarar önceki yılın aynı çeyreğiyle karşılaştırılır (raporun önceki dönem sütunu öyle)."""
    r = getattr(ok, "_rg_sonuc", None)
    if r is None:
        r = ok._rg_sonuc = risk_grubu_oku(ok)
    tablo = "a" if ad.startswith("RİSK-1. a)") else "c1" if ad.startswith("RİSK-1. c.1)") else "c2"
    harita = rg_kod_haritasi(d["items"], tablo)
    yil_sonu = donemler[0]
    ceyrek = donemler[1] if len(donemler) > 1 else donemler[0]
    ad_of = {k: a for k, a, *_ in d["items"]}
    grup_ad = ["İştirak/bağlı ortaklık", "Doğrudan ve dolaylı ortaklar", "Diğer gerçek ve tüzel kişiler"]
    kayit = []
    for (g, alt, s), kod in harita.items():
        c = r[tablo].get(("cari", g, alt, s))
        o = r[tablo].get(("onceki", g, alt, s))
        if c is None and o is None:
            continue
        kayit.append({"kod": kod, "satir": ad_of.get(kod, s), "yol": " › ".join(x for x in (grup_ad[g] if g < 3 else "", alt) if x),
                      "sutun": "", "c": c, "o": o, "birim": ok.carpan, "sayfa": r["sayfa"] or 0, "etiket": "",
                      "kaynak": "risk grubu tablosu", "duzen": "R",
                      "donem": ceyrek if s in ("faiz", "kar") else yil_sonu})
    dolu = [k for k, *_ in d["items"] if dn_deger(d, yil_sonu, k, bs)]
    return {"donem": yil_sonu, "kayit": kayit, "dolu": dolu}


# ---------------------------------------------------------------------------
# Sigorta kapsamındaki / kapsamında bulunmayan tasarruf mevduatı – özel okuyucu
SG_DIPNOTLAR = ("2.a.2.i) Sigorta kaps", "2.a.4.i) Sigorta kaps")
SG_SATIR = [   # (anahtar, boşluksuz sade ad deseni)
    ("kiyi", r"KIYIB(N?K|ANKACILIG)"), ("b_kiyi", r"MUNHASIRANKIYI"),
    ("b_yd", r"YURTDISI(NDAKI)?SUBELERDEBULUNANMEVDUAT"), ("yd", r"YURTDISI(NDAKI)?SUBELERDEBULUNANYABANCI"),
    ("b_hakim", r"HAKIMORTAK"), ("b_yon", r"YONETIMVEYAMUDURLER"), ("b_tck", r"5237|TCK"),
    ("dth", r"NITELIGINIHAIZDTH"), ("dig", r"NITELIGINIHAIZDIG"), ("tas", r"^TASARRUFMEVDUATI$"),
]


def sg_anahtar(ad: str):
    a = sade(ad).replace(" ", "").replace(".", "")
    return next((k for k, d in SG_SATIR if re.search(d, a)), None)


def sg_kod_haritasi(items) -> dict:
    harita, grup = {}, ""
    for kod, ad, sv, *_ in items:
        a = sade(ad).replace(" ", "").replace(".", "")
        if a.startswith("SIGORTAKAPSAMINDABULUNAN"):
            grup = "kap"; continue
        if a.startswith("SIGORTALIMITINIASAN"):
            grup = "asan"; continue
        k = sg_anahtar(ad)
        if k:
            harita[(grup if not k.startswith("b_") else "", k)] = kod
    return harita


def sigorta_oku(ok: "Okuyucu") -> dict:
    """{'cari'|'onceki': {(kap|asan|'', satır): değer}}"""
    d = pymupdf.open(ok.yol)
    sonuc = {"cari": {}, "onceki": {}, "sayfa": None}
    bas = next((i for i in range(len(d)) if "LIMITINI ASAN" in sade(d[i].get_text())
                and "TASARRUF" in sade(d[i].get_text())), None)
    if bas is None:
        return sonuc
    sonuc["sayfa"] = bas + 1
    bolum, tampon, merkez, bitti, donem_once = None, [], [], False, False
    for i in range(bas, min(len(d), bas + 2)):
        if bitti:
            break
        W = d[i].rect.width
        for kel in ok._satirlar_xy(d[i]):
            sm = sade(" ".join(w for _, _, w, _ in kel))
            n_r = 0
            for _, _, w, _ in reversed(kel):
                if ok.sayi_mi(w) or w in ("-", "–"):
                    n_r += 1
                else:
                    break
            while n_r and (kel[len(kel) - n_r][0] < W * 0.35 or re.fullmatch(r"\(\d\)", kel[len(kel) - n_r][2])):
                n_r -= 1
            rakam = [((x0 + x1) / 2, w) for x0, x1, w, _ in kel[len(kel) - n_r:]] if n_r else []
            etiket = " ".join(w for _, _, w, _ in kel[:len(kel) - n_r])
            es = sade(etiket)
            if not rakam:
                if "LIMITINI ASAN" in sm and bolum in (None, "A"):
                    bolum, tampon = "A", []
                    continue
                if "KAPSAMINDA BULUNMAYAN" in sm:
                    bolum, tampon, merkez = "B", [], []
                    continue
                if bolum and re.match(r"^(TICARI|TUZEL KISI) MEVDUAT", es):
                    bolum = bolum + "_tic"; tampon = []
                    continue
                if bolum == "B_tic" and re.match(r"^([a-z]|\d+)(\.\d+)*\.\s", etiket.strip()):
                    bitti = True; break
                if bolum == "B" and re.match(r"^([a-zç]|\d+)(\.\d+)*\.\s", etiket.strip()) and sonuc["cari"]:
                    bitti = True; break
                hb = []
                for q, (x0, x1, w, _) in enumerate(kel):
                    if sade(w) in ("CARI", "ONCEKI"):
                        x1b = kel[q + 1][1] if q + 1 < len(kel) and sade(kel[q + 1][2]) == "DONEM" else x1
                        hb.append((x0 + x1b) / 2)
                if len(hb) == 2 and bolum == "A" and not sonuc["cari"]:
                    donem_once = True                     # Cari Dönem (kapsam, aşan) | Önceki Dönem (kapsam, aşan)
                    continue
                if len(hb) in (2, 4):
                    merkez = hb
                    continue
                if es:
                    tampon.append(etiket)
                continue
            if bolum not in ("A", "B"):
                tampon = []
                continue
            if re.match(r"^(TICARI|TUZEL KISI) MEVDUAT", es):
                bolum = bolum + "_tic"; tampon = []
                continue
            anahtar = None
            for n in range(0, len(tampon) + 1):
                aday = " ".join(tampon[len(tampon) - n:] + [etiket])
                anahtar = sg_anahtar(aday)
                if anahtar:
                    break
            if anahtar is None and bolum == "A":              # YKB: "Türk Lirası mevduat" / "Döviz tevdiat" / "Diğer mevduat"
                anahtar = {"TURK LIRASI MEVDUAT": "tas", "DOVIZ TEVDIAT HESAPLARI": "dth",
                           "DIGER MEVDUAT HESAPLARI": "dig"}.get(es)
            tampon = []
            if anahtar is None or (bolum == "A") == anahtar.startswith("b_"):
                continue
            ns = 4 if bolum == "A" else 2
            if len(rakam) == ns or len(merkez) != ns:
                yer = list(range(len(rakam)))[:ns]
            else:
                yer = [min(range(ns), key=lambda q: abs(merkez[q] - x)) for x, _ in rakam]
            for (x, w), q in zip(rakam, yer):
                v = 0.0 if w in ("-", "–") else ok.sayi(w)
                if bolum == "B":
                    grup, dn = "", ("cari" if q == 0 else "onceki")
                elif donem_once:
                    grup, dn = ("kap" if q % 2 == 0 else "asan"), ("cari" if q < 2 else "onceki")
                else:
                    grup, dn = ("kap" if q < 2 else "asan"), ("cari" if q % 2 == 0 else "onceki")
                sonuc[dn][(grup, anahtar)] = v
    return sonuc


def sigorta_kayit(ok: "Okuyucu", ad: str, d: dict, bs: int, donemler) -> dict:
    r = getattr(ok, "_sg_sonuc", None)
    if r is None:
        r = ok._sg_sonuc = sigorta_oku(ok)
    harita = sg_kod_haritasi(d["items"])
    yil_sonu = donemler[0]
    ad_of = {k: a for k, a, *_ in d["items"]}
    kayit = []
    for key, kod in harita.items():
        c, o = r["cari"].get(key), r["onceki"].get(key)
        if c is None and o is None:
            continue
        kayit.append({"kod": kod, "satir": ad_of.get(kod, ""), "yol": {"kap": "Sigorta kapsamında", "asan": "Limiti aşan"}.get(key[0], ""),
                      "sutun": "", "c": c, "o": o, "birim": ok.carpan, "sayfa": r["sayfa"] or 0, "etiket": "",
                      "kaynak": "sigorta kapsamı tablosu", "duzen": "S"})
    dolu = [k for k, *_ in d["items"] if dn_deger(d, yil_sonu, k, bs)]
    return {"donem": yil_sonu, "kayit": kayit, "dolu": dolu}


# ---------------------------------------------------------------------------
# Mevduata ödenen faizin vade yapısı (2.b.4.i) – özel okuyucu
MF_DIPNOT = "2.b.4.i) Mevduata ödenen faiz"
MF_SATIR = [("kiymetli", r"KIYMETLI"), ("ihbar", r"7GUNIHBARLI"), ("dth", r"^(DTH|DOVIZTEVDIAT)"),
            ("bank", r"ANKALAR(ARASI)?MEVDUAT"), ("tas", r"TASARRUF"), ("resmi", r"^RESMI"),
            ("ticari", r"^TICARI"), ("diger", r"^DIGER")]
MF_SUTUN_TBB = {"VADESIZ": "vadesiz", "VADESIZMEVDUAT": "vadesiz", "1AYAKADAR": "1a", "3AYAKADAR": "3a",
                "6AYAKADAR": "6a", "1YILAKADAR": "1y", "1YILDANUZUN": "uzun", "BIRIKIMLIMEVDUAT": "birikimli"}


def mf_kalem(ad: str):
    a = re.sub(r"^\d+B", "", sade(ad).replace(" ", "").replace(".", ""))
    return next((k for k, d in MF_SATIR if re.search(d, a)), None)


def mf_kod_haritasi(items) -> dict:
    """{(TP|YP, satır, sütun): kod}; sütun 'toplam' = satır kalemi."""
    harita, bolum, satir = {}, None, None
    for kod, ad, sv, *_ in items:
        a = sade(ad).replace(" ", "").replace(".", "")
        if a.startswith("TURKPARASI"):
            bolum, satir = "TP", None; continue
        if a.startswith("YABANCIPARA"):
            bolum, satir = "YP", None; continue
        if bolum is None:
            continue
        s = MF_SUTUN_TBB.get(a)
        if s and satir:
            harita[(bolum, satir, s)] = kod
            continue
        if a == "VADELIMEVDUAT":
            continue
        k = mf_kalem(ad)
        if k:
            satir = k
            harita[(bolum, satir, "toplam")] = kod
    return harita


def mevduat_faiz_oku(ok: "Okuyucu") -> dict:
    """{'cari'|'onceki': {(TP|YP, satır, sütun): değer}}"""
    d = pymupdf.open(ok.yol)
    sonuc = {"cari": {}, "onceki": {}, "sayfa": None}
    sayfalar = [i for i in range(len(d)) if "MEVDUATA ODENEN FAIZ" in sade(d[i].get_text())
                and "KADAR" in sade(d[i].get_text())]
    if not sayfalar:
        return sonuc
    bas = sayfalar[-1]                      # içindekiler sayfasını atla
    sonuc["sayfa"] = bas + 1
    donem, bolum, merkez, tampon, basladi, bitti = "cari", None, {}, [], False, False
    kad_top, yedek = [], {}
    for i in range(bas, min(len(d), bas + 2)):
        if bitti:
            break
        W = d[i].rect.width
        for kel in ok._satirlar_xy(d[i]):
            sm = sade(" ".join(w for _, _, w, _ in kel))
            if "MEVDUATA ODENEN FAIZ" in sm:
                basladi = True
                continue
            if not basladi:
                continue
            n_r = 0
            for _, _, w, _ in reversed(kel):
                if ok.sayi_mi(w) or w in ("-", "–"):
                    n_r += 1
                else:
                    break
            while n_r and kel[len(kel) - n_r][0] < W * 0.20:
                n_r -= 1
            rakam = [((x0 + x1) / 2, w) for x0, x1, w, _ in kel[len(kel) - n_r:]] if n_r else []
            etiket = " ".join(w for _, _, w, _ in kel[:len(kel) - n_r])
            es = re.sub(r"^\d+B", "", sade(etiket).replace(" ", "").replace(".", ""))
            if es.startswith("TURKPARASI"):
                bolum = "TP"; tampon = []; continue
            if es.startswith("YABANCIPARA"):           # "Yabancı Para: - - -" (VakıfBank)
                bolum = "YP"; tampon = []; continue
            if not rakam and re.match(r"^ONCEKI DONEM", sm) and kel[0][0] < W * 0.3 and sonuc["cari"]:
                yedek = dict(merkez)
                donem, merkez, bolum, kad_top = "onceki", {}, None, []
            if not rakam:
                # başlık kelimeleri → sütun merkezleri
                kad_top += [(x0 + x1) / 2 for x0, x1, w, _ in kel if sade(w) == "KADAR"]
                if len(kad_top) >= 4:                      # başlık birkaç satıra bölünmüş olabilir (VakıfBank)
                    for anahtar, x in zip(("1a", "3a", "6a", "1y"), sorted(kad_top)[:4]):
                        merkez[anahtar] = x
                for x0, x1, w, _ in kel:                   # Garanti konsolide: "1-3 Ay", "3-6 Ay", "6-12 Ay"
                    k_ = {"1-3": "3a", "3-6": "6a", "6-12": "1y"}.get(w.strip())
                    if k_:
                        merkez[k_] = (x0 + x1) / 2
                if len(kad_top) < 4 and kad_top and any(k_ in merkez for k_ in ("3a", "6a", "1y")):
                    merkez["1a"] = min(kad_top)
                uz = [(x0 + x1) / 2 for x0, x1, w, _ in kel if sade(w) == "UZUN"]
                if uz:
                    merkez["uzun"] = uz[0]
                for x0, x1, w, _ in kel:
                    sw = sade(w)
                    if sw == "VADESIZ":
                        merkez["vadesiz"] = (x0 + x1) / 2
                    elif sw.startswith("BIRIKIMLI"):
                        merkez["birikimli"] = (x0 + x1) / 2
                    elif sw == "TOPLAM" and x0 > W * 0.5:
                        merkez["toplam"] = (x0 + x1) / 2
                    elif sw == "ONCEKI" and x0 > W * 0.5:
                        merkez["onc_toplam"] = (x0 + x1) / 2      # YKB: cari tabloda önceki dönem toplamı
                if donem == "onceki" and yedek:
                    for k_, v_ in yedek.items():       # önceki tabloda tekrarlanmayan başlık (Akbank "Vadesiz")
                        merkez.setdefault(k_, v_)
                if re.match(r"^CARI DONEM", sm) and kel[0][0] < W * 0.3 and sonuc["cari"]:
                    bitti = True; break
                if es.startswith("TURKPARASI"):
                    bolum = "TP"; tampon = []; continue
                if es.startswith("YABANCIPARA"):
                    bolum = "YP"; tampon = []; continue
                if bolum and es and not re.match(r"^(HESAPADI|VADELI|CARIDONEM|ONCEKIDONEM)", es):
                    tampon.append(etiket)
                if bolum == "YP" and donem == "onceki" and re.match(r"^(\d+(\.\d+)*\.|[a-zç]\.)\s", etiket.strip()):
                    bitti = True; break
                continue
            if bolum is None or not merkez:
                tampon = []
                continue
            if re.match(r"^(GENEL)?TOPLAM", es):
                tampon = []
                if es.startswith("GENEL") and (donem == "onceki" or "onc_toplam" in merkez):
                    bitti = True; break
                continue
            satir = None
            for n in range(0, len(tampon) + 1):
                satir = mf_kalem(" ".join(tampon[len(tampon) - n:] + [etiket]))
                if satir:
                    break
            tampon = []
            if satir is None:
                continue
            dolan = set()
            for x, w in rakam:
                sut = min(merkez, key=lambda k_: abs(merkez[k_] - x))
                v = 0.0 if w in ("-", "–") else ok.sayi(w)
                if sut == "onc_toplam":
                    sonuc["onceki"][(bolum, satir, "toplam")] = v
                else:
                    sonuc[donem][(bolum, satir, sut)] = v
                    dolan.add(sut)
            for sut in merkez:                          # boş hücre "-" yazılmamış (İş): 0
                if sut != "onc_toplam" and sut not in dolan:
                    sonuc[donem].setdefault((bolum, satir, sut), 0.0)
    return sonuc


def mevduat_faiz_kayit(ok: "Okuyucu", ad: str, d: dict, bs: int, donemler) -> dict:
    r = getattr(ok, "_mf_sonuc", None)
    if r is None:
        r = ok._mf_sonuc = mevduat_faiz_oku(ok)
    harita = mf_kod_haritasi(d["items"])
    ceyrek = donemler[1] if len(donemler) > 1 else donemler[0]     # gelir tablosu dipnotu: önceki yılın aynı çeyreği
    ad_of = {k: a for k, a, *_ in d["items"]}
    kayit = []
    for key, kod in harita.items():
        c, o = r["cari"].get(key), r["onceki"].get(key)
        if c is None and o is None:
            continue
        kayit.append({"kod": kod, "satir": ad_of.get(kod, ""), "yol": f"{key[0]} › {key[1]}", "sutun": key[2],
                      "c": c, "o": o, "birim": ok.carpan, "sayfa": r["sayfa"] or 0, "etiket": "",
                      "kaynak": "mevduata ödenen faiz tablosu", "duzen": "F", "donem": ceyrek})
    dolu = [k for k, *_ in d["items"] if dn_deger(d, ceyrek, k, bs)]
    return {"donem": ceyrek, "kayit": kayit, "dolu": dolu}


# ---------------------------------------------------------------------------
# Donuk alacak dipnotları (2. f.10 i–iv) – özel okuyucu
DA_DIPNOTLAR = ("2. f.10 i)", "2. f.10 ii)", "2. f.10 iii)", "2. f.10 iv)")


def _da_sade(ad: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", sade(ad.replace("(", " ").replace(")", " ")))


def da_satir(tablo: str, a: str, kat: str = ""):
    """Boşluksuz sade ad → satır anahtarı."""
    if tablo == "i":
        if "KARSILIKLARDANONCEKIBRUT" in a: return "brut"
        if "YENIDENYAPILANDIRILAN" in a: return "yy"
        return None
    if tablo in ("ii", "iii"):
        for k, d in (("acilis", r"^ONCEKIDONEMSONUBAKIYE"), ("intikal", r"INTIKAL"),
                     ("giris", r"HESAPLARINDANGIRIS"), ("cikis", r"HESAPLARINACIKIS"), ("tahsilat", r"TAHSILAT"),
                     ("kayit", r"KAYITTANDUS"), ("satilan", r"^SATILAN"), ("kurumsal", r"^KURUMSAL"),
                     ("bireysel", r"^BIREYSEL"), ("kk", r"^KREDIKART"), ("diger", r"^DIGER$"),
                     ("son", r"^(CARIDONEM)?DONEMSONUBAKIYE"), ("karsilik", r"^KARSILIK"), ("net", r"^BILANCODAKINET")):
            if re.search(d, a):
                return k
        return None
    # iv
    if a.startswith("KARSILIK"):
        return (kat, "karsilik") if kat else None
    for kk, d in (("gt", "GERCEKVETUZEL"), ("bank", "^BANKALAR"), ("diger", "^DIGERKREDI")):
        if re.search(d, a):
            return (kk, "brut" if "BRUT" in a else "net")
    return None


def da_grup(a: str):
    m = re.search(r"(?<![A-Z])(III|IV|V)GRUP", a)
    return {"III": 0, "IV": 1, "V": 2}[m.group(1)] if m else None


def da_kod_haritasi(items, tablo: str) -> dict:
    harita, g, kat = {}, None, ""
    for kod, ad, sv, *_ in items:
        a = _da_sade(ad)
        gg = da_grup(a)
        if gg is not None:
            g, kat = gg, ""
            harita[(g, "grup")] = kod
            continue
        if g is None:
            continue
        s = da_satir(tablo, a, kat)
        if tablo == "iv" and s:
            kat = s[0]
        if s:
            harita[(g,) + (s if isinstance(s, tuple) else (s,))] = kod
    return harita


def donuk_alacak_oku(ok: "Okuyucu") -> dict:
    """{tablo: {'cari'|'onceki': {(grup, ...satır): değer}}}"""
    d = pymupdf.open(ok.yol)
    sonuc = {t: {"cari": {}, "onceki": {}} for t in ("i", "ii", "iii", "iv")}
    sonuc["sayfa"] = None
    sayfalar = [i for i in range(len(d)) if "DONUK ALACAK" in sade(d[i].get_text())
                and re.search(r"\bIV\.? ?GRUP", sade(d[i].get_text()).replace("IV GRUP", "IV.GRUP") + " ")
                or "DONUK ALACAK HAREKET" in sade(d[i].get_text())]
    sayfalar = [i for i in range(len(d)) if re.search(r"KARSILIKLARDAN ONCEKI BRUT|DONUK ALACAK HAREKET|DONUK ALACAK.{0,40}YENIDEN YAPILANDIR", sade(d[i].get_text()))
                and re.search(r"TAHSIL IMKANI|TAHSILI SUPHELI|III\.? GRUP", sade(d[i].get_text()))
                and re.search(r"YENIDEN YAPILANDIRILAN KREDILER|INTIKAL", sade(d[i].get_text()))]
    if not sayfalar:
        return sonuc
    # içindekiler sayfasını atla: rakamlı tablo içeren ilk sayfa
    bas = sayfalar[0]
    sonuc["sayfa"] = bas + 1
    tablo, donem, merkez, kat, tampon, bitti = None, "cari", [], "", "", False
    for i in range(bas, min(len(d), bas + 4)):
        if bitti:
            break
        W = d[i].rect.width
        for kel in ok._satirlar_xy(d[i]):
            sm = sade(" ".join(w for _, _, w, _ in kel))
            n_r = 0
            for _, _, w, _ in reversed(kel):
                if ok.sayi_mi(w) or w in ("-", "–"):
                    n_r += 1
                else:
                    break
            while n_r and (kel[len(kel) - n_r][0] < W * 0.40 or re.fullmatch(r"\(\d\)", kel[len(kel) - n_r][2])):
                n_r -= 1
            rakam = [((x0 + x1) / 2, w) for x0, x1, w, _ in kel[len(kel) - n_r:]] if n_r else []
            etiket = " ".join(w for _, _, w, _ in kel[:len(kel) - n_r])
            a = _da_sade(etiket)
            # tablo başlıkları (numaralı başlık satırları)
            if not rakam and "DONUK ALACAK" in sm or (not rakam and re.search(r"KULLANICI GRUP|YABANCI PARA OLARAK", sm)):
                yeni = ("i" if "YENIDEN YAPILANDIR" in sm else "ii" if "HAREKET" in sm else
                        "iii" if "YABANCI PARA" in sm else "iv" if "KULLANICI GRUP" in sm else None)
                if "TFRS" in sm or "FAIZ TAHAKKUK" in sm:
                    if tablo:
                        bitti = True; break
                if yeni:
                    tablo, donem, merkez, kat, tampon = yeni, "cari", [], "", ""
                continue
            if tablo is None:
                continue
            if re.match(r"^(\d+\.)+\d*\.?\s", etiket.strip()) and not rakam and "DONUK" not in sm:
                if tablo == "iv":
                    bitti = True; break
                continue
            if not rakam:
                gm = [((x0 + x1) / 2) for q, (x0, x1, w, _) in enumerate(kel)
                      if sade(w) == "GRUP" and q > 0 and sade(kel[q - 1][2]).rstrip(".") in ("III", "IV", "V")]
                if len(gm) == 3:
                    merkez = gm
                    continue
                if re.match(r"^(CARI|ONCEKI) DONEM", sm) and "SONU" not in sm:
                    donem = "onceki" if sm.startswith("ONCEKI") else "cari"
                    continue
                tampon = (tampon + " " + etiket).strip() if len(tampon) < 60 else etiket
                continue
            if re.match(r"^(CARI|ONCEKI)DONEMNET", a):          # iv: "Cari Dönem (Net)" toplam satırı
                donem = "onceki" if a.startswith("ONCEKI") else "cari"
                tampon = ""
                continue
            s = da_satir(tablo, a, kat) or (da_satir(tablo, _da_sade(tampon + " " + etiket), kat) if tampon else None)
            tampon = ""
            if s is None:
                continue
            if tablo == "iv":
                kat = s[0]
            if len(rakam) == 3 or len(merkez) != 3:
                yer = list(range(len(rakam)))[:3]
            else:
                yer = [min(range(3), key=lambda q: abs(merkez[q] - x)) for x, _ in rakam]
            for (x, w), q in zip(rakam, yer):
                v = 0.0 if w in ("-", "–") else abs(ok.sayi(w))
                anahtar = (q,) + (s if isinstance(s, tuple) else (s,))
                if tablo == "ii" and s == "acilis":            # açılış bakiyesi = önceki yıl sonu bakiyesi
                    sonuc["ii"]["onceki"].setdefault((q, "son"), v)
                    continue
                sonuc[tablo][donem][anahtar] = v
    return sonuc


def donuk_alacak_kayit(ok: "Okuyucu", ad: str, d: dict, bs: int, donemler) -> dict:
    r = getattr(ok, "_da_sonuc", None)
    if r is None:
        r = ok._da_sonuc = donuk_alacak_oku(ok)
    tablo = next(t for t, p in zip(("i", "ii", "iii", "iv"), DA_DIPNOTLAR) if ad.startswith(p))
    harita = da_kod_haritasi(d["items"], tablo)
    yil_sonu = donemler[0]
    ad_of = {k: a for k, a, *_ in d["items"]}
    kayit = []
    for key, kod in harita.items():
        kaynak_key = (key[0], "net") if key[1:] == ("grup",) and tablo in ("ii", "iii") else key
        c, o = r[tablo]["cari"].get(kaynak_key), r[tablo]["onceki"].get(kaynak_key)
        if c is None and o is None:
            continue
        kayit.append({"kod": kod, "satir": ad_of.get(kod, ""), "yol": ["III. Grup", "IV. Grup", "V. Grup"][key[0]],
                      "sutun": " ".join(str(x) for x in key[1:]), "c": c, "o": o, "birim": ok.carpan,
                      "sayfa": r["sayfa"] or 0, "etiket": "", "kaynak": "donuk alacak tablosu", "duzen": "D"})
    dolu = [k for k, *_ in d["items"] if dn_deger(d, yil_sonu, k, bs)]
    return {"donem": yil_sonu, "kayit": kayit, "dolu": dolu}


# ---------------------------------------------------------------------------
# Şablon tablolar – ada ve sütun başlığına göre okuyucu (önceki dönemi olmayan / boş hücresi "-" yazılmayan
# raporlar için: İş, YKB). Kalemler adından, sütunlar başlık kelimelerinin konumundan bulunur.
def _st_sayilar(ok, kel, W, sol=0.30):
    n = 0
    for _, _, w, _ in reversed(kel):
        if (ok.sayi_mi(w) and not re.fullmatch(r"\(\d\)|\(\*+\)", w)) or w in GD_TIRE:
            n += 1
        else:
            break
    while n and kel[len(kel) - n][0] < W * sol:
        n -= 1
    rakam = [((x0 + x1) / 2, 0.0 if w in GD_TIRE else ok.sayi(w)) for x0, x1, w, _ in kel[len(kel) - n:]] if n else []
    etiket = " ".join(w for _, _, w, _ in kel[:len(kel) - n])
    return rakam, etiket


def _st_yeni_baslik(metin: str) -> bool:
    """Numaralı yeni bölüm başlığı mı ("5.2.13 …", "h.2. …", "c) …"); sayfa üstündeki "30 Eylül 2026 …" değil."""
    t = metin.strip()
    if re.search(r"\b20\d\d\b", t) or re.match(r"^[IVX]+\.\s", t) or (len(t) > 25 and t == t.upper()):
        return False                                   # tarihli ya da büyük harfli sayfa üst başlığı
    return re.match(r"^(\(?[A-Za-zÇĞİÖŞÜçğıöşü]{1,2}[.)](\d+\.?)*|\d+(\.\d+)*\.?)\s+\S", t) is not None


def _st_sade(s):
    return re.sub(r"[^A-Z0-9]", "", sade(s))


def _st_satirlar(ok, d, bas, n_sayfa, baslik_re):
    """Başlık satırından itibaren satırlar: {sayfa, W, kel, sm, etiket, rakam}"""
    out, basladi = [], False
    for i in range(bas, min(len(d), bas + n_sayfa)):
        W = d[i].rect.width
        for kel in ok._satirlar_xy(d[i]):
            sm = sade(" ".join(w for _, _, w, _ in kel))
            rakam, etiket = _st_sayilar(ok, kel, W)
            if not basladi:
                if re.search(baslik_re, sm) and not rakam:
                    basladi = True
                continue
            if Okuyucu.sayfa_basligi_mi(" ".join(w for _, _, w, _ in kel)) and not rakam:
                continue
            out.append({"sayfa": i + 1, "W": W, "kel": kel, "sm": sm, "etiket": etiket, "rakam": rakam})
    return out


def _st_merkez(satirlar, kelime_re, adet, ilk=40):
    """İlk satırlardaki başlık kelimelerinin x merkezleri (soldan sağa)."""
    for L in satirlar[:ilk]:
        if L["rakam"]:
            continue
        m = [(x0 + x1) / 2 for x0, x1, w, _ in L["kel"] if re.fullmatch(kelime_re, sade(w))]
        if len(m) == adet:
            return m
    return None


def _st_yerlestir(rakam, merkez):
    if not merkez:
        return None
    out = [None] * len(merkez)
    for cx, v in rakam:
        j = min(range(len(merkez)), key=lambda q: abs(merkez[q] - cx))
        out[j] = v if out[j] is None else out[j]
    return out


# kalem desenleri: (anahtar(lar), desen) – desen kalemin sadeleştirilmiş, boşluksuz adında aranır; birden çok anahtarlı
# desen her eşleşmede sıradaki boş anahtara yazılır. "son" anahtarı okunduğunda tablo biter.
BASIT_TABLOLAR = [
    {"tbb": "2.b.3) İhraç Edilen Menkul", "baslik": r"IHRAC EDILEN MENKUL KIYMETLERE VERILEN FAIZ", "mod": "tpyp",
     "gelir": True, "satirlar": [(("#1", "son"), r"IHRACEDILENMENKUL|^FAIZLER$|VERILENFAIZLER")]},
    {"tbb": "2.a.3) Menkul Değerlerden", "baslik": r"MENKUL DEGERLERDEN ALINAN FAIZ", "mod": "tpyp", "gelir": True,
     "satirlar": [(("#2",), r"KAR(VEYA)?ZARARA|ZARARAYANSITILAN|ALIMSATIM"),
                  (("#3",), r"DIGERKAPSAMLI|SATILMAYAHAZIR"),
                  (("#4",), r"ITFAEDILMIS|VADEYEKADAR"),
                  (("#1", "son"), r"^TOPLAM")]},
    {"tbb": "2.a.1) Kredilerden Alınan", "baslik": r"KREDILERDEN ALINAN FAIZ", "mod": "tpyp", "gelir": True,
     "satirlar": [(("#2",), r"^KISAVADELI"), (("#3",), r"^ORTAVEUZUN|^ORTAUZUN"),
                  (("#4",), r"TAKIPTEKI|DONUKALACAK"), (("#5",), r"KAYNAKKUL"),
                  (("#1", "son"), r"^TOPLAM")]},
    {"tbb": "2.ı.8) Menkul değerler değer artış", "baslik": r"MENKUL DEGERLER DEGER ARTIS FONU", "mod": "tpyp",
     "satirlar": [(("#2",), r"^ISTIRAK"), (("#3",), r"DIGERKAPSAMLI|SATILMAYAHAZIR|DEGERLEMEFARKI"),
                  (("#4",), r"^KURFARKI"), (("#1", "son"), r"^TOPLAM")]},
    {"tbb": "2.ı.1) Ödenmiş sermayenin", "baslik": r"ODENMIS SERMAYE(NIN)?( GOSTERIMI)?\W*$", "sayfa": r"ODENMIS SERMAYE", "mod": "co",
     "satirlar": [(("#3",), r"IMTIYAZLI"), (("#2",), r"HISSESENEDIKARSILIGI")]},
    {"tbb": "2.h) Bankanın sermaye benzeri", "baslik": r"SERMAYE BENZERI (BORCLANMA ARACLARINA|KREDILERE|BORCLANMA ARACLARI ILE) ILISKIN", "mod": "tpyp",
     "satirlar": [(("#2",), r"ILAVEANASERMAYE"), (("#5",), r"KATKISERMAYE"),
                  (("#3", "#6"), r"SERMAYEBENZERIKREDILER"),
                  (("#4", "#7"), r"SERMAYEBENZERIBORCLANMAARACLARI$"),
                  (("#1", "son"), r"^TOPLAM")]},
]
BASIT_TABLOLAR.append(
    {"tbb": "2. a) Nakit değerler ve TCMB", "baslik": r"NAKIT DEGERLER VE T ?C MERKEZ BANKASI|NAKIT DEGERLER VE TCMB|"
     r"NAKIT DEGERLER VE MERKEZ BANKASI", "mod": "tpyp",
     "satirlar": [(("#2",), r"^KASA|^EFEKTIF"), (("#3",), r"MERKEZBANKASI|^TCMB"), (("#4",), r"^DIGER"),
                  (("#1", "son"), r"^TOPLAM")]})
BASIT_TABLOLAR.append(
    {"tbb": "Kaldıraç Oranı -TMS-Özet", "baslik": r"KALDIRAC ORAN", "sayfa": r"TMS UYARINCA DUZENLENEN", "mod": "co",
     "yalniz_eslesen": True,
     "satirlar": [(("#4",), r"TUREVFINANSAL"), (("#5",), r"MENKULKIYMETVEYAEMTIA"), (("#6",), r"BILANCODISI"),
                  (("#7",), r"DIGERFARKLAR"), (("#8", "son"), r"TOPLAMRISKTUTARI"),
                  (("#2",), r"TOPLAMVARLIK"), (("#3",), r"VARLIKTUTARI|ARASINDAKIFARK")]})
BASIT_DIPNOTLAR = tuple(sp["tbb"] for sp in BASIT_TABLOLAR)


def basit_tablo_oku(ok, sp) -> dict:
    """Tablonun bütün rakamlı satırları: [{anahtar (ad deseninden), c, o, etiket}]; en çok kalemi tanınan sayfa seçilir."""
    d = pymupdf.open(ok.yol)
    en_iyi = None
    for bas in range(len(d)):
        if not re.search(sp.get("sayfa", sp["baslik"]), sade(d[bas].get_text())):
            continue
        S = _st_satirlar(ok, d, bas, 2, sp["baslik"])
        if not any(L["rakam"] for L in S[:30]):
            continue
        if sp["mod"] == "tpyp":
            merkez = _st_merkez(S, r"TP|YP", 4)
            beklenen = 4
        else:
            merkez = None
            for L in S[:30]:
                if L["rakam"]:
                    break
                c = [(x0 + x1) / 2 for x0, x1, w, _ in L["kel"] if sade(w) == "CARI"]
                o = [(x0 + x1) / 2 for x0, x1, w, _ in L["kel"] if sade(w) == "ONCEKI"]
                if len(c) == 1 and len(o) == 1:
                    merkez = [c[0], o[0]]
                    break
            beklenen = 2
        satirlar, tampon, atanmis, bos = [], [], set(), 0
        for j, L in enumerate(S):
            if not L["rakam"]:
                if satirlar and _st_yeni_baslik(" ".join(w for _, _, w, _ in L["kel"])):
                    break                                        # yeni numaralı başlık
                tampon = (tampon + [L["etiket"]])[-2:]
                bos += 1
                if satirlar and bos > 12:
                    break
                continue
            bos = 0
            deg = _st_yerlestir(L["rakam"], merkez) if merkez and len(L["rakam"]) != beklenen else None
            if deg is None:
                deg = [v for _, v in L["rakam"]]
                if len(deg) != beklenen:
                    tampon = []
                    continue
            sonraki = S[j + 1]["etiket"] if j + 1 < len(S) and not S[j + 1]["rakam"] else ""
            adaylar = [L["etiket"], " ".join(tampon + [L["etiket"]]), L["etiket"] + " " + sonraki]
            if _st_sade(L["etiket"]).startswith("TOPLAM"):
                adaylar = [L["etiket"]]                       # "Toplam" satırı üstündeki rakamsız adı almasın
            tampon = []
            anahtar, son = None, False
            for metin in adaylar:
                a = _st_sade(metin)
                if not a:
                    continue
                for anahtarlar, desen in sp["satirlar"]:
                    if re.search(desen, a):
                        bosta = [k for k in anahtarlar if k != "son" and k not in atanmis]
                        if bosta:
                            anahtar, son = bosta[0], "son" in anahtarlar
                            break
                if anahtar:
                    break
            if beklenen == 4:
                c = None if deg[0] is None and deg[1] is None else (deg[0] or 0) + (deg[1] or 0)
                o = None if deg[2] is None and deg[3] is None else (deg[2] or 0) + (deg[3] or 0)
            else:
                c, o = deg[0], deg[1]
            if anahtar:
                atanmis.add(anahtar)
            satirlar.append({"anahtar": anahtar, "c": c, "o": o, "etiket": L["etiket"]})
            if son:
                break
        son_k = {k for a_, _ in sp["satirlar"] if "son" in a_ for k in a_ if k != "son"}
        n = sum(1 for x in satirlar if x["anahtar"] and (x["anahtar"] not in son_k or len(sp["satirlar"]) == 1))
        if n and (en_iyi is None or n > en_iyi[0]):
            en_iyi = (n, {"satirlar": satirlar, "sayfa": bas + 1})
    return en_iyi[1] if en_iyi else {"satirlar": [], "sayfa": None}


def basit_tablo_kayit(ok: "Okuyucu", ad: str, d: dict, bs: int, donemler) -> dict:
    """Her TBB kalemi için: önceki dönem değeri TBB'dekiyle aynı olan satır (TBB bankaya göre farklı satıra
    yazabiliyor); yoksa ya da birden çoksa ad deseninin bulduğu satır."""
    sp = next(s for s in BASIT_TABLOLAR if ad.startswith(s["tbb"]))
    onbellek = ok.__dict__.setdefault("_basit", {})
    r = onbellek.get(sp["tbb"])
    if r is None:
        r = onbellek[sp["tbb"]] = basit_tablo_oku(ok, sp)
    karsi = donemler[1] if sp.get("gelir") and len(donemler) > 1 else donemler[0]
    ad_of = {str(k): a for k, a, *_ in d["items"]}
    anahtarlar = []
    for a_, _ in sp["satirlar"]:
        anahtarlar += [x for x in a_ if x != "son"]
    kayit, kullanilan = [], set()
    for anahtar in anahtarlar:
        q = int(anahtar[1:])                     # "#n": TBB kalem listesindeki sıra (solo ve konsolide aynı yapıda)
        if q >= len(d["items"]):
            continue
        kod = str(d["items"][q][0])
        tv = dn_deger(d, karsi, kod, bs)
        ad_satir = next((x for x in r["satirlar"] if x["anahtar"] == anahtar), None)
        sec = ad_satir
        if tv:
            es = [x for x in r["satirlar"] if x["o"] is not None and abs(x["o"] - tv / ok.carpan) <= 1.5
                  and id(x) not in kullanilan]
            if ad_satir is not None and ad_satir in es:
                sec = ad_satir
            elif len(es) == 1:
                sec = es[0]
        if sec is None or id(sec) in kullanilan:
            continue
        if sp.get("yalniz_eslesen") and not (tv and sec["o"] is not None and abs(sec["o"] - tv / ok.carpan) <= 1.5):
            continue                                      # satır düzeni güvenilmez: yalnız önceki dönemi TBB'yle tutanlar
        kullanilan.add(id(sec))
        kayit.append({"kod": int(kod), "satir": ad_of[kod], "yol": "", "sutun": "", "c": sec["c"], "o": sec["o"],
                      "birim": ok.carpan, "sayfa": r["sayfa"] or 0, "etiket": sec["etiket"],
                      "kaynak": "şablon tablo (ad + başlık)", "duzen": "S", "donem": karsi})
    dolu = [k for k, *_ in d["items"] if dn_deger(d, karsi, k, bs)]
    return {"donem": karsi, "kayit": kayit, "dolu": dolu}


# ---------------------------------------------------------------------------
# Tüketici (2.f.4) ve taksitli ticari (2.f.5) krediler – satır grubu × vade sütunu
KR_DIPNOTLAR = ("2. f.4) Tüketici kre", "2. f.5) Taksitli ticari")
KR_BASLIK = {"2. f.4) Tüketici kre": r"TUKETICI KREDILERI.{0,12}BIREYSEL KREDI KARTLARI",
             "2. f.5) Taksitli ticari": r"TAKSITLI TICARI KREDILER VE KURUMSAL KREDI KARTLARI"}
KR_TIP = [("PKK", r"^PERSONELKREDIKART"), ("BKK", r"^BIREYSELKREDIKART"), ("KKK", r"^KURUMSALKREDIKART"),
          ("PK", r"^PERSONELKREDI"), ("TTK", r"^TAKSITLITICARI"), ("TK", r"^TUKETICIKREDI"), ("KMH", r"^KREDILIMEVDUAT")]
KR_COCUK = [("KONUT", r"^KONUT"), ("TASIT", r"^TASIT"), ("IHTIYAC", r"^IHTIYAC"), ("ISYERI", r"^ISYERI"),
            ("TAKSITSIZ", r"^TAKSITSIZ"), ("TAKSITLI", r"^TAKSITLI$"), ("DIGER", r"^DIGER")]


def kr_grup(a: str):
    tip = next((t for t, d in KR_TIP if re.search(d, a)), None)
    if not tip:
        return None
    para = "DE" if "DOVIZE" in a else "YP" if re.search(r"YP(GERCEKKISI|TUZELKISI)?$", a) else \
        "TP" if re.search(r"TP(GERCEKKISI|TUZELKISI)?$", a) else None
    return tip, para


def kr_cocuk(a: str):
    return next((c for c, d in KR_COCUK if re.search(d, a)), None)


def kr_kod_haritasi(items) -> dict:
    """{(vade, tip, para, çocuk|None): kod}"""
    harita, vade, grup, kmh_sira = {}, None, None, {}
    for kod, ad, sv, *_ in items:
        a = _st_sade(ad)
        if sv == 1:
            vade = "kisa" if a.startswith("KISA") else "orta" if a.startswith("ORTA") else None
            grup = None
            continue
        if vade is None or sv < 2:
            continue
        if sv == 2:
            g = kr_grup(a)
            if g and g[0] == "KMH" and g[1] is None:          # f.5: iki "Kredili Mevduat Hesabı (Tüzel Kişi)" = TP, YP
                n = kmh_sira.get(vade, 0)
                kmh_sira[vade] = n + 1
                g = ("KMH", "TP" if n == 0 else "YP")
            grup = g
            if g:
                harita[(vade, g[0], g[1], None)] = kod
        elif sv == 3 and grup:
            c = kr_cocuk(a)
            if c:
                harita[(vade, grup[0], grup[1], c)] = kod
    return harita


def kredi_tablo_oku(ok, baslik) -> dict:
    d = pymupdf.open(ok.yol)
    en_iyi = None
    for bas in range(len(d)):
        if not re.search(baslik, sade(d[bas].get_text())):
            continue
        S = _st_satirlar(ok, d, bas, 3, baslik)
        sonuc = {"cari": {}, "onceki": {}, "sayfa": bas + 1}
        donem, merkez, grup, okunan, bos = "cari", {}, None, 0, 0
        for L in S:
            sm = L["sm"]
            if not L["rakam"]:
                if re.match(r"^ONCEKI DONEM", sm):
                    donem, grup = "onceki", None
                elif re.match(r"^CARI DONEM", sm):
                    donem, grup = "cari", None
                kel = [(x0, x1, sade(w)) for x0, x1, w, _ in L["kel"] if x0 >= L["W"] * 0.30]
                for q, (x0, x1, s_) in enumerate(kel):
                    x1_ = kel[q + 1][1] if q + 1 < len(kel) and kel[q + 1][2] == "VADELI" else x1
                    if s_ == "KISA":
                        merkez["kisa"] = (x0 + x1_) / 2
                    elif s_ == "ORTA":
                        merkez["orta"] = (x0 + (kel[q + 1][1] if q + 1 < len(kel) and kel[q + 1][2] == "VE" else x1)) / 2
                        if q + 3 < len(kel) and kel[q + 2][2] == "UZUN":
                            merkez["orta"] = (x0 + kel[q + 3][1]) / 2
                    elif s_ == "UZUN" and (q == 0 or kel[q - 1][2] != "VE"):
                        cx = (x0 + x1_) / 2
                        merkez["orta"] = (merkez["orta"] + cx) / 2 if "orta" in merkez and abs(merkez["orta"] - cx) < 40 else cx
                    elif s_ in ("FAIZ", "TAHAKKUK", "REESKONTLARI", "REESKONT") and "tah" not in merkez:
                        tah = [k_ for k_ in kel if k_[2] in ("FAIZ", "VE", "GELIR", "TAHAKKUK", "REESKONTLARI", "REESKONT")]
                        merkez["tah"] = (min(k_[0] for k_ in tah) + max(k_[1] for k_ in tah)) / 2
                    elif s_ == "TOPLAM":
                        merkez["top"] = (x0 + x1) / 2
                a = _st_sade(L["etiket"])
                if okunan and ((_st_yeni_baslik(" ".join(w for _, _, w, _ in L["kel"])) and not re.search(baslik, sm))
                               or any(re.search(b_, sm) for b_ in KR_BASLIK.values() if b_ != baslik)):
                    break
                g = kr_grup(a) if a else None
                if g:
                    grup = g
                bos += 1
                if okunan and bos > 15:
                    break
                continue
            bos = 0
            a = _st_sade(L["etiket"])
            if re.match(r"^TOPLAM", a):
                if donem == "onceki" or not any(re.match(r"^ONCEKI DONEM", x["sm"]) for x in S):
                    pass
                grup = None
                continue
            if len(merkez) < 2:
                continue
            ks = sorted(merkez, key=merkez.get)
            deg = {}
            for cx, v in L["rakam"]:
                k = min(ks, key=lambda q: abs(merkez[q] - cx))
                deg.setdefault(k, v)
            g = kr_grup(a)
            if g:
                grup, anahtar = g, (g[0], g[1], None)
            elif grup and kr_cocuk(a):
                anahtar = (grup[0], grup[1], kr_cocuk(a))
            else:
                continue
            if anahtar in sonuc[donem]:
                if anahtar[0] == "KMH":                         # "Gerçek Kişi" + "Personel" satırları (Garanti)
                    for k_, v_ in deg.items():
                        sonuc[donem][anahtar][k_] = (sonuc[donem][anahtar].get(k_) or 0) + v_
                continue
            sonuc[donem][anahtar] = deg
            okunan += 1
        if okunan and (en_iyi is None or okunan > en_iyi[0]):
            en_iyi = (okunan, sonuc)
    return en_iyi[1] if en_iyi else {"cari": {}, "onceki": {}, "sayfa": None}


def kredi_tablo_kayit(ok: "Okuyucu", ad: str, d: dict, bs: int, donemler) -> dict:
    on = next(x for x in KR_DIPNOTLAR if ad.startswith(x))
    onbellek = ok.__dict__.setdefault("_kredi", {})
    r = onbellek.get(on)
    if r is None:
        r = onbellek[on] = kredi_tablo_oku(ok, KR_BASLIK[on])
    harita = kr_kod_haritasi(d["items"])
    yil_sonu = donemler[0]
    ad_of = {k: a for k, a, *_ in d["items"]}
    kayit = []
    for (vade, tip, para, cocuk), kod in harita.items():
        vals = []
        for donem in ("cari", "onceki"):
            deg = r[donem].get((tip, para, cocuk))
            if deg is None:
                vals.append(None)
                continue
            # TBB'de faiz/gelir tahakkukları kısa vadeli tutara eklenir (toplam = kısa + orta-uzun)
            v = (deg.get("kisa", 0) or 0) + (deg.get("tah", 0) or 0) if vade == "kisa" else (deg.get("orta", 0) or 0)
            vals.append(v)
        c, o = vals
        if c is None and o is None:
            continue
        kayit.append({"kod": kod, "satir": ad_of.get(kod, ""), "yol": "Kısa vadeli" if vade == "kisa" else "Orta ve uzun vadeli",
                      "sutun": "", "c": c, "o": o, "birim": ok.carpan, "sayfa": r["sayfa"] or 0, "etiket": "",
                      "kaynak": "kredi tablosu (satır grubu × vade)", "duzen": "S"})
    dolu = [k for k, *_ in d["items"] if dn_deger(d, yil_sonu, k, bs)]
    return {"donem": yil_sonu, "kayit": kayit, "dolu": dolu}


# ---------------------------------------------------------------------------
# Standart nitelikli / yakın izlemedeki krediler (2.f.2.i): büyük tablo (satır × 4 sütun) ve karşılık tablosu
YI_DIPNOTLAR = ("2.f.2.i) Standart nitelikli ve yakın izle", "2.f.2.i) Standart Nitelikli - Yakın")
YI_BASLIK = r"STANDART NITELIKLI VE YAKIN IZLEME|BIRINCI VE IKINCI ASAMA KREDILER"
YI_SATIR = [("ISLETME", r"^ISLETME"), ("IHRACAT", r"^IHRACAT"), ("ITHALAT", r"^ITHALAT"), ("MALI", r"^MALIKESIM"),
            ("TUKETICI", r"^TUKETICI"), ("KART", r"^KREDIKART"), ("IDISI", r"^IHTISASDISI"), ("IHTISAS", r"^IHTISASKREDI"),
            ("DALACAK", r"^DIGERALACAK"), ("DIGER", r"^DIGER$")]
YI_KUCUK = [("12AY", r"12AYLIK"), ("ONEMLI", r"KREDIRISKINDEONEMLI"), ("GENEL", r"GENELKARSILIK")]
YI_SUTUN = {"std": ("STANDART", "NITELIKLI"), "yer": ("KAPSAMINDA", "ALMAYANLAR", "ALMAYAN"),
            "soz": ("SOZLESME", "KOSULLARINDA", "DEGISIKLIK"), "fin": ("FINANSMAN",)}


def yakin_izleme_oku(ok) -> dict:
    d = pymupdf.open(ok.yol)
    en_iyi = None
    for bas in range(len(d)):
        if not re.search(YI_BASLIK, sade(d[bas].get_text())):
            continue
        S = _st_satirlar(ok, d, bas, 3, YI_BASLIK)
        sonuc = {"buyuk": {"cari": {}, "onceki": {}}, "kucuk": {"cari": {}, "onceki": {}}, "sayfa": bas + 1}
        donem, bkel, kkel, tampon, okunan, bos, yan_k = "cari", [], [], [], 0, 0, False
        for L in S:
            sm = L["sm"]
            if not L["rakam"]:
                if okunan and _st_yeni_baslik(" ".join(w for _, _, w, _ in L["kel"])):
                    break
                bos += 1
                if okunan and bos > 15:
                    break
                ws = [((x0 + x1) / 2, sade(w)) for x0, x1, w, _ in L["kel"] if x0 >= L["W"] * 0.28]
                if re.match(r"^ONCEKI DONEM", sm):
                    donem, bkel = "onceki", []
                elif re.match(r"^CARI DONEM", sm):
                    donem, bkel = "cari", []
                if any(w in sum(YI_SUTUN.values(), ()) for _, w in ws):
                    bkel += ws
                tum = [sade(w) for _, _, w, _ in L["kel"]]
                if "CARI" in tum or "ONCEKI" in tum:          # yeni dönem başlığı: karşılık tablosu başlığı yeniden
                    kkel, yan_k = [], ("CARI" in tum and "ONCEKI" in tum)
                if any(w in ("STANDART", "YAKIN") for _, w in ws):
                    for q, (x, w) in enumerate(ws):           # "Standart Nitelikli" / "Yakın İzlemedeki" ifadesinin ortası
                        if w in ("STANDART", "YAKIN"):
                            x2 = ws[q + 1][0] if q + 1 < len(ws) and ws[q + 1][1] in ("NITELIKLI", "IZLEMEDEKI") else x
                            kkel.append(((x + x2) / 2, w))
                if L["etiket"].strip():
                    tampon = (tampon + [L["etiket"]])[-2:]
                continue
            bos = 0
            # başlık kelimeleri rakamlı satırda da olabilir (İş: "İhtisas Dışı Krediler … Değişiklik Finansman 2.341.890 …")
            n_et = len(L["kel"]) - len(L["rakam"])
            bas_k = [((x0 + x1) / 2, sade(w)) for x0, x1, w, _ in L["kel"][:n_et]
                     if x0 >= L["W"] * 0.28 and sade(w) in sum(YI_SUTUN.values(), ())]
            if bas_k:
                bkel += bas_k
                L = dict(L, etiket=" ".join(w for x0, _, w, _ in L["kel"][:n_et] if x0 < L["W"] * 0.28))
            a = _st_sade(L["etiket"]) or _st_sade(" ".join(tampon))
            if not _st_sade(L["etiket"]) and tampon:
                a = _st_sade(tampon[-1])
            tampon = []
            satir = next((k_ for k_, d_ in YI_SATIR if re.search(d_, a)), None)
            kucuk = next((k_ for k_, d_ in YI_KUCUK if re.search(d_, a)), None)
            if satir:
                merkez = {}
                for s_, kelimeler in YI_SUTUN.items():
                    xs = [x for x, w in bkel if w in kelimeler]
                    if xs:
                        merkez[s_] = sum(xs) / len(xs)
                if len(merkez) < 4:
                    continue
                deg = {}
                for cx, v in L["rakam"]:
                    s_ = min(merkez, key=lambda q: abs(merkez[q] - cx))
                    deg.setdefault(s_, v)
                if satir not in sonuc["buyuk"][donem]:
                    sonuc["buyuk"][donem][satir] = deg
                    okunan += 1
                kkel = []
            elif kucuk and False:
                # karşılık tablosu: Standart/Yakın sütunları; cari ve önceki yan yana olabilir
                std = sorted(x for x, w in kkel if w == "STANDART")[-2:]
                yak = sorted(x for x, w in kkel if w == "YAKIN")[-2:]
                yan = yan_k and len(std) == 2 and len(yak) == 2
                if not std or not yak:
                    continue
                merkez = ({("cari", "std"): std[0], ("cari", "yak"): yak[0], ("onceki", "std"): std[1], ("onceki", "yak"): yak[1]}
                          if yan else {(donem, "std"): std[-1], (donem, "yak"): yak[-1]})
                for cx, v in L["rakam"]:
                    dk = min(merkez, key=lambda q: abs(merkez[q] - cx))
                    sonuc["kucuk"][dk[0]].setdefault((dk[1], kucuk), v)
                okunan += 1
        # karşılık tablosu: ayrı geçiş (arada başka tablolar ve başlıklar olabilir)
        kkel, yan_k, donem = [], False, "cari"
        for L in S:
            tum = [sade(w) for _, _, w, _ in L["kel"]]
            if not L["rakam"]:
                if re.match(r"^ONCEKI DONEM", L["sm"]):
                    donem = "onceki"
                elif re.match(r"^CARI DONEM", L["sm"]):
                    donem = "cari"
            n_et = len(L["kel"]) - len(L["rakam"])
            if not L["rakam"] and ("CARI" in tum or "ONCEKI" in tum):
                kkel, yan_k = [], ("CARI" in tum and "ONCEKI" in tum)
            ws = [((x0 + x1) / 2, sade(w)) for x0, x1, w, _ in L["kel"][:n_et] if x0 >= L["W"] * 0.28]
            if not L["rakam"] and any(w in ("STANDART", "YAKIN") for _, w in ws) and not any(
                    w in ("KAPSAMINDA", "ALMAYANLAR", "ALMAYAN", "SOZLESME", "FINANSMAN", "YAPILANDIRMA") for _, w in ws):
                for q, (x, w) in enumerate(ws):
                    if w in ("STANDART", "YAKIN"):
                        x2 = ws[q + 1][0] if q + 1 < len(ws) and ws[q + 1][1] in ("NITELIKLI", "IZLEMEDEKI") else x
                        kkel.append(((x + x2) / 2, w))
                continue
            if not L["rakam"]:
                continue
            kucuk = next((k_ for k_, d_ in YI_KUCUK if re.search(d_, _st_sade(L["etiket"]))), None)
            if not kucuk:
                continue
            std = sorted(x for x, w in kkel if w == "STANDART")[-2:]
            yak = sorted(x for x, w in kkel if w == "YAKIN")[-2:]
            if not std or not yak:
                continue
            yan = yan_k and len(std) == 2 and len(yak) == 2
            merkez = ({("cari", "std"): std[0], ("cari", "yak"): yak[0], ("onceki", "std"): std[1], ("onceki", "yak"): yak[1]}
                      if yan else {(donem, "std"): std[-1], (donem, "yak"): yak[-1]})
            for cx, v in L["rakam"]:
                dk = min(merkez, key=lambda q: abs(merkez[q] - cx))
                sonuc["kucuk"][dk[0]].setdefault((dk[1], kucuk), v)
            okunan += 1
        if okunan and (en_iyi is None or okunan > en_iyi[0]):
            en_iyi = (okunan, sonuc)
    return en_iyi[1] if en_iyi else {"buyuk": {"cari": {}, "onceki": {}}, "kucuk": {"cari": {}, "onceki": {}}, "sayfa": None}


def yi_kod_haritasi(items, kucuk: bool) -> dict:
    harita, yigin = {}, []
    for kod, ad, sv, *_ in items:
        a = _st_sade(ad)
        while yigin and yigin[-1][0] >= sv:
            yigin.pop()
        yigin.append((sv, a))
        if sv == 0:
            continue
        yol = " ".join(x for _, x in yigin[1:])
        if kucuk:
            s_ = "std" if "STANDART" in yol else "yak" if "YAKIN" in yol else None
            r_ = next((k_ for k_, d_ in YI_KUCUK if re.search(d_, a)), None)
            if s_ and r_:
                harita[(s_, r_)] = kod
            continue
        s_ = ("fin" if "YENIDENFINANSMAN" in yol else "soz" if "SOZLESME" in yol else
              "yer" if "YERALMAYAN" in yol else "std" if "STANDART" in yol else None)
        a2 = re.sub(r"^STANDARTNITELIKLI", "", a)
        r_ = next((k_ for k_, d_ in YI_SATIR if re.search(d_, a2)), None)
        if s_ and r_ and r_ != "IDISI" and len(yigin) >= 3:     # "İhtisas dışı" üst kalemi ara toplamla dolar
            harita[(s_, r_)] = kod
    return harita


def yi_kucuk_metinden(ok):
    """Karşılık tablosu sütunları 'Standart | Yakın İzleme' × 'Cari | Önceki' düzeninde (Halkbank): satır başına
    dört rakam — std cari, std önceki, yakın cari, yakın önceki. Konum okuyucusu bu düzeni tanımadığında kullanılır."""
    d = pymupdf.open(ok.yol)
    rk = r"(\(?[\d.,]+\)?|-)"
    for pg in d:
        t = re.sub(r"\s+", " ", pg.get_text())
        # düzen 1 (Halkbank): Standart | Yakın, altında Cari | Önceki → std c, std o, yak c, yak o
        # düzen 2 (Akbank konsolide): Cari | Önceki, altında Standart | Yakın | Standart | Yakın → std c, yak c, std o, yak o
        m0 = re.search(r"Standart Nitelikli Krediler\s+Yakın İzlemedeki Krediler\s+Cari Dönem\s+Önceki Dönem\s+"
                       r"Cari Dönem\s+Önceki Dönem", t)
        duzen = 1
        if not m0:
            m0 = re.search(r"Standart Nitelikli Krediler\s+Yakın İzlemedeki Krediler\s+Standart Nitelikli Krediler\s+"
                           r"Yakın İzlemedeki Krediler", t)
            duzen = 2
            if m0 and not re.search(r"Cari Dönem.{0,80}Önceki Dönem", t[max(0, m0.start() - 250):m0.start()]):
                m0 = None
        if not m0 or "Kredi Riskinde Önemli Artış" not in t[m0.end():m0.end() + 400]:
            continue
        parca = t[m0.end():m0.end() + 400]
        T = {"cari": {}, "onceki": {}}
        for r_, desen in YI_KUCUK:
            for m in re.finditer(r"((?:12 )?[A-Za-zÇĞİÖŞÜçğıöşü' ]{6,60}?)\s+" + " ".join([rk] * 4), parca):
                if not re.search(desen, _st_sade(m.group(1))):
                    continue
                v = [0.0 if x == "-" else ok.sayi(x) for x in m.groups()[1:]]
                if duzen == 2:
                    v = [v[0], v[2], v[1], v[3]]
                T["cari"][("std", r_)], T["onceki"][("std", r_)] = v[0], v[1]
                T["cari"][("yak", r_)], T["onceki"][("yak", r_)] = v[2], v[3]
                break
        if T["cari"]:
            return T
    return None


def yakin_izleme_kayit(ok: "Okuyucu", ad: str, d: dict, bs: int, donemler) -> dict:
    r = getattr(ok, "_yi_sonuc", None)
    if r is None:
        r = ok._yi_sonuc = yakin_izleme_oku(ok)
    kucuk = ad.startswith(YI_DIPNOTLAR[1])
    harita = yi_kod_haritasi(d["items"], kucuk)
    T = r["kucuk"] if kucuk else r["buyuk"]
    if kucuk and any((s_, r_) not in T.get("cari", {}) for s_ in ("std", "yak") for r_ in ("12AY", "ONEMLI")):
        M = yi_kucuk_metinden(ok)
        if M:                                       # yalnız konum okuyucusunun bulamadıkları tamamlanır
            T = {"cari": {**M["cari"], **T.get("cari", {})}, "onceki": {**M["onceki"], **T.get("onceki", {})}}
    ad_of = {k: a for k, a, *_ in d["items"]}
    kayit = []
    for (s_, r_), kod in harita.items():
        if kucuk:
            c, o = T["cari"].get((s_, r_)), T["onceki"].get((s_, r_))
        else:
            c = T["cari"][r_].get(s_, 0.0) if r_ in T["cari"] else None
            o = T["onceki"][r_].get(s_, 0.0) if r_ in T["onceki"] else None
        if c is None and o is None:
            continue
        kayit.append({"kod": kod, "satir": ad_of.get(kod, ""), "yol": s_, "sutun": "", "c": c, "o": o,
                      "birim": ok.carpan, "sayfa": r["sayfa"] or 0, "etiket": "",
                      "kaynak": "standart nitelikli / yakın izleme tablosu", "duzen": "S"})
    dolu = [k for k, *_ in d["items"] if dn_deger(d, donemler[0], k, bs)]
    return {"donem": donemler[0], "kayit": kayit, "dolu": dolu}


# ---------------------------------------------------------------------------
# I ve II'nci grupta sınıflandırılan gayrinakdi krediler (2.a.3.iii): satır × (I. Grup TP, YP | II. Grup TP, YP).
# Genel eşleyici önceki dönem tablosu olmayan raporlarda (İş) bulamaz; o zaman bu okuyucu kullanılır.
GG_DIPNOT = "2.a.3.iii) I ve II"
GG_SATIR = [(0, r"^(TOPLAM)?GAYRINAKDIKREDILER$|^TOPLAMGAYRINAKDI"), (1, r"^TEMINATMEKTUP"), (2, r"^AVAL"),
            (3, r"^AKREDITIF"), (4, r"^CIRO"), (5, r"^MENKULKIYMET"), (6, r"^FAKTORING"), (7, r"^DIGERGARANTI")]


def grup_gayrinakdi_oku(ok) -> dict:
    d = pymupdf.open(ok.yol)
    baslik = r"GRUPTA SINIFLANDIRILAN GAYRI ?NAKDI"
    for bas in range(len(d)):
        if not re.search(baslik, sade(d[bas].get_text())):
            continue
        S = _st_satirlar(ok, d, bas, 2, baslik)
        sonuc, donem, merkez, okunan = {"cari": {}, "onceki": {}, "sayfa": bas + 1}, "cari", None, 0
        for L in S:
            if not L["rakam"]:
                if okunan and _st_yeni_baslik(" ".join(w for _, _, w, _ in L["kel"])):
                    break
                if "ONCEKI DONEM" in L["sm"]:
                    donem = "onceki"
                elif "CARI DONEM" in L["sm"]:
                    donem = "cari"
                m = [(x0 + x1) / 2 for x0, x1, w, _ in L["kel"] if sade(w) in ("TP", "YP")]
                if len(m) == 4:
                    merkez = m
                continue
            if not merkez:
                continue
            a = _st_sade(L["etiket"])
            q = next((i for i, d_ in GG_SATIR if re.search(d_, a)), None)
            if q is None:
                continue
            deg = _st_yerlestir(L["rakam"], merkez)
            sonuc[donem].setdefault((1, q), (deg[0] or 0) + (deg[1] or 0))
            sonuc[donem].setdefault((2, q), (deg[2] or 0) + (deg[3] or 0))
            okunan += 1
        if okunan >= 3:
            return sonuc
    return {"cari": {}, "onceki": {}, "sayfa": None}


def grup_gayrinakdi_kayit(ok, d: dict, bs: int, donemler) -> dict:
    r = grup_gayrinakdi_oku(ok)
    items = d["items"]
    harita, grup = {}, None
    for kod, ad, sv, *_ in items:
        a = _st_sade(ad)
        if sv == 1:
            grup = 2 if sade(ad).startswith("II ") else 1      # "I inci Grup…" / "II nci Grup…"
            harita[(grup, 0)] = kod
        elif sv == 2 and grup:
            q = next((i for i, d_ in GG_SATIR if i and re.search(d_, a)), None)
            if q is not None:
                harita[(grup, q)] = kod
    ad_of = {k: a for k, a, *_ in items}
    kayit = []
    for key, kod in harita.items():
        c, o = r["cari"].get(key), r["onceki"].get(key)
        if c is None and o is None:
            continue
        kayit.append({"kod": kod, "satir": ad_of[kod], "yol": f"{key[0]}. grup", "sutun": "", "c": c, "o": o,
                      "birim": ok.carpan, "sayfa": r["sayfa"] or 0, "etiket": "", "kaynak": "grup tablosu (TP + YP)",
                      "duzen": "S"})
    return {"donem": donemler[0], "kayit": kayit, "dolu": [k for k, *_ in items if dn_deger(d, donemler[0], k, bs)]}


def oran_notu(p, t) -> str:
    """Oranlar (%): TBB iki ondalık ya da tam sayıya yuvarlanmış tutar."""
    if p is None or t is None:
        return ""
    a, b = abs(p), abs(t)
    if abs(a - b) < 0.006:
        return "aynı"
    if float(b).is_integer() and abs(a - b) <= 0.5 + 1e-9:
        return "yuvarlama (±1)"
    if abs(a - b) <= 0.0151:
        return "küçük fark (≤3)"
    return "FARK"


def tbb_kod_ayir(ad: str):
    m = re.match(r"^\s*((?:[A-H]|[IVXL]+)\.|\d+(?:\.\d+)*\.?)\s*(.*)$", ad)
    return (m.group(1).rstrip("."), m.group(2)) if m else ("", ad)


def tbb_esle(satirlar: list[dict], kalemler: list[tuple]) -> list[tuple]:
    """PDF satırlarını TBB kalemlerine eşler: önce sırayla (ad + numara),
    bulunamazsa tüm listede aynı numara + benzer ad. TFRS 9 uygulamayan
    bankalara ait eski kalemler aday dışıdır."""
    tbb = [(c, *tbb_kod_ayir(n)) for c, n in kalemler if "uygulamayan" not in n.lower()]
    j, out, kullanildi = 0, [], set()
    for s in satirlar:
        en_iyi = None
        for k in range(j, min(j + 12, len(tbb))):
            c, no, ad = tbb[k]
            puan = (difflib.SequenceMatcher(None, sade(ad), sade(s["ad"])).ratio() if s["ad"] else 0.5) \
                + (0.3 if no == s["kod"] else 0) - 0.01 * (k - j)
            if en_iyi is None or puan > en_iyi[2]:
                en_iyi = (k, c, puan)
        if en_iyi and en_iyi[2] >= 0.75:
            j = en_iyi[0] + 1
            kullanildi.add(en_iyi[0])
            out.append((s, en_iyi[1], "sıralı: ad + numara", round(en_iyi[2], 2)))
            continue
        aday = None
        for k, (c, no, ad) in enumerate(tbb):
            if k in kullanildi or not s["kod"] or no != s["kod"]:
                continue
            b = difflib.SequenceMatcher(None, sade(ad), sade(s["ad"])).ratio()
            if b >= 0.6 and (aday is None or b > aday[2]):
                aday = (k, c, b)
        if aday:
            kullanildi.add(aday[0])
            out.append((s, aday[1], "sırasız: aynı numara + benzer ad", round(aday[2], 2)))
        else:
            out.append((s, None, "eşleşmedi", round(en_iyi[2], 2) if en_iyi else 0.0))
    return out


def hizali_esle(satirlar: list[dict], kalemler: list[tuple]) -> list[tuple]:
    """Kalem numarası olmayan tablolar için: iki listeyi baştan sona en iyi hizalayan
    eşleme (dinamik programlama). Her iki taraftaki fazla satırlar atlanabilir;
    yalnız ad benzerliği ≥ 0,55 olan çiftler eşleşir."""
    tbb = [(c, n) for c, n in kalemler if "uygulamayan" not in n.lower()]
    a, b = [sade(x["ad"]) for x in satirlar], [sade(n) for c, n in tbb]
    n, m = len(a), len(b)
    benz = [[0.0] * m for _ in range(n)]
    for i in range(n):
        for j in range(m):
            if a[i] and b[j]:
                r = difflib.SequenceMatcher(None, a[i], b[j]).ratio()
                benz[i][j] = r if r >= 0.55 else 0.0
    P = [[0.0] * (m + 1) for _ in range(n + 1)]
    for i in range(n - 1, -1, -1):
        for j in range(m - 1, -1, -1):
            P[i][j] = max(P[i + 1][j], P[i][j + 1], (benz[i][j] + P[i + 1][j + 1]) if benz[i][j] else 0)
    out, i, j = [], 0, 0
    while i < n:
        if j < m and benz[i][j] and P[i][j] == benz[i][j] + P[i + 1][j + 1]:
            out.append((satirlar[i], tbb[j][0], "hizalı: ad benzerliği", round(benz[i][j], 2)))
            i, j = i + 1, j + 1
        elif j < m and P[i][j] == P[i][j + 1]:
            j += 1
        else:
            out.append((satirlar[i], None, "eşleşmedi", 0.0))
            i += 1
    return out


def tbb_deger(A: dict, donem: str, kod, bs: int):
    liste = A["v"].get(donem, {}).get(str(kod)) if kod else None
    return liste[bs] if (liste and bs < len(liste)) else None


def sutun_sec(anahtar: str, eslesme, A, donem, bs, carpan):
    """Önceki dönem sütunu = TBB ile en çok tutan sütun; cari sütun ona göre.
    (Gelir tablosunda sütun sırası bankadan bankaya değişir.)"""
    varsayilan = {"KZ": (0, 1), "DKG": (0, 1), "NAKIT": (0, 1), "OZK": (0, 1)}.get(anahtar, (2, 5))
    n = Counter(len(s["sayilar"]) for s, k, *_ in eslesme if k and s["sayilar"]).most_common(1)
    n = n[0][0] if n else varsayilan[1] + 1
    puan = Counter()
    for s, kod, *_ in eslesme:
        t = tbb_deger(A, donem, kod, bs)
        if not t:
            continue
        tm = round(t / carpan)
        for j, x in enumerate(s["sayilar"][:n]):
            if abs(abs(x) - abs(tm)) <= 1:
                puan[j] += 1
    if not puan:
        return varsayilan, n
    oi = max(puan, key=lambda j: (puan[j], j == varsayilan[1]))
    ci = oi - n // 2 if (n >= 6 and oi >= n // 2) else (0 if oi != 0 else 1)
    return (ci, oi), n


def fark_notu(p, t) -> str:
    if p is None or t is None:
        return ""
    d = abs(abs(t) - abs(p))
    if d == 0:
        return "aynı"
    if d <= 1:
        return "yuvarlama (±1)"
    if d <= 3 and d <= 1e-4 * max(abs(t), abs(p)):
        return "küçük fark (≤3)"
    return "FARK"


# ============================================================== 5) Excel

BASLIK = PatternFill("solid", fgColor="E8EDF5")
TURUNCU = PatternFill("solid", fgColor="FDE2C8")
YESIL = PatternFill("solid", fgColor="E3F1E4")
TUTAR = '#,##0;(#,##0);"-"'


def bicimle(ws, genislik, dondur="A2"):
    for c in range(1, ws.max_column + 1):
        ws.cell(1, c).fill = BASLIK
        ws.cell(1, c).font = Font(bold=True)
        ws.cell(1, c).alignment = Alignment(wrap_text=True, vertical="top")
    for i, w in enumerate(genislik, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = dondur
    if ws.max_row > 1:
        ws.auto_filter.ref = ws.dimensions


def donem_et(p: str) -> str:
    y, m = p.split("-")
    return f"{AY_ADI[int(m)]} {y}"


KURALLAR = [
    ("Kaynak", "KAP finansal tablo paketindeki Türkçe solo ve konsolide rapor PDF'leri. PDF'in türü dosya adından değil "
               "içeriğinden bulunur (\"KONSOLİDE OLMAYAN\" → solo); İngilizce raporlar, faaliyet raporları atlanır."),
    ("TBB verisi", "tbb-banka-finansal-tablolar raporundaki veri (GitHub). Solo raporlar TBB solo, konsolide raporlar TBB "
                   "konsolide verisiyle karşılaştırılır."),
    ("Birim", "Karşılaştırma raporun kendi biriminde yapılır (rapordaki 'Milyon/Bin Türk Lirası' ifadesinden). TBB bin TL "
              "verisi =ROUND(ham/1000;0) ile milyon TL'ye çevrilir; ham değer sayfanın sağındadır."),
    ("Sayı biçimi", "1.234.567 (Türkçe) ve 1,234,567 (İngilizce) biçimleri rapordan otomatik tanınır. Parantezli sayı "
                    "negatiftir, \"-\" sıfırdır."),
    ("Önceki dönem", "Aktif, pasif, nazım: önceki yıl sonu (Aralık). Gelir-gider, DKG ve nakit akış: önceki yılın aynı "
                     "dönemi (birikimli)."),
    ("Cari dönem", "TBB seçilen dönemi yayımladıysa cari dönem de karşılaştırılır; yayımlamadıysa yalnız PDF değeri vardır."),
    ("Sütun seçimi", "Önceki dönem sütunu, TBB ile en çok tutan PDF sütunudur (gelir tablosunda sütun sırası bankaya göre "
                     "değişir). Bilanço/nazımda cari sütun = önceki sütun − 3; gelir-gider, DKG ve nakit akışta cari birikimli sütun ilk "
                     "sütundur. Seçilen sütunlar aşağıdaki eşleşme tablosunda yazar."),
    ("Kalem eşleşmesi", "Önce sıra korunarak kalem adı ve numarasına bakılır (puan ≥ 0,75). Bulunamazsa TBB listesinin tamamında "
                        "aynı numara ve benzer ad aranır (ad benzerliği ≥ 0,6); TBB'nin konsolide listesi kısmen sırasızdır."),
    ("Eski kalemler", "TBB listesindeki \"TFRS 9 uygulamayan bankalar\" kalemleri aday dışıdır."),
    ("Dipnot referansları", "Kalem adlarındaki \"V-I-1\", \"(1)\" gibi dipnot referansları atılır."),
    ("Fark etiketi", "aynı: fark yok | yuvarlama (±1): rapordaki üst toplamlar ayrı yuvarlanır | küçük fark (≤3): en fazla 3 "
                     "birim ve oransal olarak çok küçük | FARK: daha büyük fark (turuncu)."),
    ("İşaret", "Karşılaştırma mutlak değerle yapılır; TBB bazı gider kalemlerini ters işaretle tutar (Not: TBB ters işaretli)."),
    ("Hisse başına kâr", "\"Tam TL\" / \"Hisse başına\" satırları birime bölünmez; 0,006 toleransla karşılaştırılır."),
    ("Bilanço denkliği", "Aktif toplamı ile pasif toplamı (PDF) her rapor ve dönem için karşılaştırılır; ±1 denk sayılır."),
    ("Özkaynak tablosu", "Dipnotlardaki 'Özkaynak kalemlerine ilişkin açıklamalar' tablosu; önceki dönem önceki yıl sonudur. "
                         "Kalın yazılmış rakamsız satırlar başlıktır; kalın olmayan rakamsız satırlar, rakamın geldiği "
                         "satıra kadar birleştirilip kalem adı olur. Rakamlar sayfanın sağındaki, sağında başka kelime "
                         "olmayan sayılardır. Kalem numarası olmadığından eşleşme iki listeyi baştan sona hizalayarak "
                         "ad benzerliğiyle (≥ 0,55) yapılır; iki taraftaki fazla satırlar atlanabilir."),
    ("Beklenen eşleşmeyenler", "TBB'de karşılığı olmadığı bilinen satırlar aranmaz, 'beklenen (kural)' olarak işaretlenir: "
                               "nakit akış A/B/C başlıkları; solo raporlarda azınlık payları ve grubun kârı; konsolide "
                               "özkaynakta üçüncü kişilerin payları. Sayfa başlığı parçaları ('(Devamı)', 'Mali bünyeye…') okunmaz."),
    ("Kur riski", "Dipnotlardaki 'Kur riskine ilişkin bilgiler' tablosu. Her kalem dört satır olur: Avro, ABD Doları, Diğer YP, "
                  "Toplam; TBB'deki EURO / USD / Diğer YP alt kalemleri ve kalemin kendisi (toplam). Varlıklar, Yükümlülükler ve "
                  "Yükümlülükler bölümleri sırayla hizalanarak ad benzerliğiyle (≥ 0,45), özet bölümü (net pozisyonlar, türevler, "
                  "gayrinakdi krediler) sıraya bakılmadan en benzer adla (≥ 0,6) eşlenir; 'Diğer Mevduat' = 'Döviz Tevdiat "
                  "Hesabı'. Önceki dönem (önceki yıl sonu) için raporda yalnız özet satırlar vardır. TBB verisi raporunuzun "
                  "dipnot dosyasından (GitHub) alınır; TBB yükümlülükleri negatif tutar."),
    ("Özkaynak oranları", "Sermaye yeterliliği oranları ve tamponlar TBB'de tablo kalemi değil, ayrı göstergedir (cet1, t1, car, "
                          "buf, ccb, ccyb, sib, bufR); adlarından eşlenir, 0,006 toleransla karşılaştırılır (TBB iki ondalık tutar)."),
    ("Mali bünye tabloları", "Faiz oranı riski, ortalama faiz oranları, hisse senedi pozisyon riski, likidite karşılama oranı "
                             "(LKO), LKO en düşük/en yüksek, kalan vade, NİFO ve kaldıraç tabloları dipnotlardan okunur. Her "
                             "kalem × sütun ayrı satırdır. TBB verisi raporunuzun ilgili dipnot dosyasından (GitHub) alınır; "
                             "önceki dönem önceki yıl sonudur (Aralık)."),
    ("Mali bünye – okuma", "Tablo, başlığındaki 'Cari Dönem' / 'Önceki Dönem' satırlarıyla bloklara ayrılır. Rakamlar x "
                           "konumlarına göre sütunlara dağıtılır (boş hücreler kaymaz). Rakamsız ad satırları sonraki "
                           "rakamlı satırla birleşir; 'Varlıklar:' gibi satırlar bölüm başlığıdır. Tablo altındaki (1), (*) "
                           "açıklamaları okunmaz."),
    ("Mali bünye – eşleşme", "Faiz oranı riski, ortalama faiz ve kalan vade: kalem adından anahtar (nakit, bankalar, krediler, "
                             "itfa edilmiş maliyet, …) ve sütun adından TBB alt kalemi (1 aya kadar, faizsiz, Avro…). Toplam "
                             "sütunu TBB'de alt kalemlerin toplamıdır. LKO (1–23), NİFO (1–34) ve kaldıraç (1–15) BDDK "
                             "şablonudur: satır numarasıyla eşlenir; numara yazmayan raporlarda (ör. Ziraat) numara kalem "
                             "adından bulunur."),
    ("LKO / NİFO sütunları", "LKO: dikkate alınma oranı uygulanmamış TP+YP, YP; uygulanmış TP+YP, YP (TBB'deki 'TP+YP' alt "
                             "kalemi PDF'teki TP+YP sütunudur). NİFO: a–ç kalan vadeye göre uygulanmamış tutar, d uygulanmış "
                             "toplam. NİFO'da alt kalemleri olan satırların (1, 4, 7, 11, 17, 26) ve toplamların (14, 33) a–ç "
                             "sütunları TBB'de tutulmaz: 'yalnız PDF'. 10 ve 25. satırlar TBB'de yoktur. Birleşik hücreli "
                             "satırlarda (28–30) tek tutar, TBB'deki uygulanmamış toplamla karşılaştırılır."),
    ("Mali bünye oranları", "LKO (23), NİFO (34), kaldıraç (15) ve ortalama faiz oranları birime bölünmez. TBB LKO ve NİFO'yu "
                            "tam sayıya yuvarlanmış tutar: PDF değeri yuvarlandığında tutuyorsa 'yuvarlama (±1)'."),
    ("Yalnız PDF", "Hisse senedi pozisyon riski ve LKO en düşük/en yüksek tabloları TBB'de yoktur; PDF değerleri (ve LKO "
                   "tarihleri) yazılır, karşılaştırılmaz."),
    ("Göstergeler", "TBB'nin ayrı tuttuğu göstergeler: YP aktif toplamı (bilanço YP sütunu), YKLV stoku ve net nakit "
                    "çıkışları (LKO 21–22, uygulanmış TP+YP), LKO TP+YP ve YP (23), mevcut ve gerekli istikrarlı fon "
                    "(NİFO 14 ve 33, d), NİFO (34). TBB burada oranları iki ondalıkla tutar."),
    ("Öteki dipnotlar – kapsam", "Özel okuyucusu olmayan bütün TBB dipnotları (aktif, pasif, nazım, gelir tablosu, risk "
                                 "grubu ve öteki mali bünye dipnotları) ile '8. Mevduat Vade Yapısı' tablosu genel "
                                 "eşleyiciyle aranır. Önceki dönem olarak önceki yıl sonu ve önceki yılın aynı çeyreği "
                                 "denenir, daha çok kalem bulunan seçilir."),
    ("Öteki dipnotlar – yöntem", "Düzen A: TBB'de alt kalemleri sütun olan satırın önceki dönem değerleri PDF'te aynı "
                                 "satırda yan yana aranır. Düzen B: tek TBB kalemi, adı benzeyen (≥%60) PDF satırının bir "
                                 "sütunu ya da sütun toplamıdır (TBB çoğu dipnotta TP+YP toplamını tutar). Cari değer: yan "
                                 "yana tablolarda aynı satırın cari yarısı, alt alta tablolarda cari bloktaki aynı sıradaki "
                                 "satır. Bulunan kalemlerin önceki dönemi tanım gereği tutar; asıl çıktı cari değerdir."),
    ("Öteki dipnotlar – yıl sonu", "Bu banka için TBB'de önceki yılın aynı çeyreğinde verisi olmayan dipnotlar (ör. bölge "
                                   "ve sektör risk profilleri) yalnız yıl sonu raporlarında yayımlanır: aranmaz, kapsam "
                                   "sayfasında 'yalnız yıl sonu' yazılır."),
    ("PDF'te bulunamayanlar", "Ara dönemde yayımlanan dipnotlarda TBB'de dolu olup PDF'te bulunamayan kalemler ayrı "
                              "sayfada listelenir (çoğu ara toplam ya da raporda farklı ayrımla verilen kalem)."),
]


class Rapor:
    KARS_BASLIK = ["Banka", "Tür", "Kalem no", "Kalem adı (PDF)", "{cari} – PDF", "Önceki dönem", "Önceki dönem – PDF",
                   "Önceki dönem – TBB", "Fark (|TBB|−|PDF|)", "Not", "{cari} – TBB", "{cari} fark", "{cari} not",
                   "Birim", "TBB kodu", "Kalem adı (TBB)", "Önceki dönem – TBB ham (bin TL)", "{cari} – TBB ham (bin TL)",
                   "PDF sütun 1", "PDF sütun 2", "PDF sütun 3", "PDF sütun 4", "PDF sütun 5", "PDF sütun 6"]

    GD_BASLIK = ["Banka", "Tür", "Dipnot grubu", "Dipnot", "TBB kalemi", "Sütun", "{cari} – PDF", "Önceki dönem",
                 "Önceki dönem – PDF", "Önceki dönem – TBB", "Fark (|TBB|−|PDF|)", "Not", "{cari} – TBB", "{cari} fark",
                 "{cari} not", "Birim", "TBB kodu", "PDF sayfa", "PDF satırı", "Cari değer kaynağı", "Düzen",
                 "Önceki dönem – TBB ham", "{cari} – TBB ham"]
    GD_KAPSAM_BASLIK = ["Banka", "Tür", "Dipnot grubu", "Dipnot", "Önceki dönem", "TBB'de dolu kalem", "PDF'te bulunan",
                        "Kapsama %", "Cari değeri okunan", "aynı", "yuvarlama (±1)", "küçük fark (≤3)", "FARK",
                        "{cari}: aynı", "{cari}: ±1", "{cari}: küçük", "{cari}: FARK", "Durum"]
    GOSTERGE_BASLIK = ["Banka", "Tür", "Gösterge", "TBB anahtarı", "PDF kaynağı", "{cari} – PDF", "Önceki dönem",
                       "Önceki dönem – PDF", "Önceki dönem – TBB", "Fark (|TBB|−|PDF|)", "Not", "{cari} – TBB",
                       "{cari} fark", "{cari} not", "Birim", "Önceki dönem – TBB ham", "{cari} – TBB ham"]
    DIPNOT_BASLIK = ["Banka", "Tür", "Kalem no", "Kalem adı (PDF)", "Sütun", "{cari} – PDF", "Önceki dönem",
                     "Önceki dönem – PDF", "Önceki dönem – TBB", "Fark (|TBB|−|PDF|)", "Not", "{cari} – TBB",
                     "{cari} fark", "{cari} not", "Birim", "TBB kodu", "Kalem adı (TBB)", "Önceki dönem – TBB ham",
                     "{cari} – TBB ham", "Bölüm (PDF)"]

    def __init__(self, yil: int, ay: int, D: dict):
        self.yil, self.ay, self.D = yil, ay, D
        self.cari = f"{yil}-{ay:02d}"
        self.cari_et = f"{AY_ADI[ay]} {yil}"
        self.wb = Workbook()
        self.ozet = self.wb.active
        self.ozet.title = "Özet"
        self.ozet.append(["Banka", "Tür", "Tablo", "PDF satırı", "TBB ile eşleşen", "aynı", "yuvarlama (±1)",
                          "küçük fark (≤3)", "FARK", "TBB karşılığı yok", "beklenen (kural)",
                          f"{self.cari_et} TBB'de var mı", f"{self.cari_et}: aynı", f"{self.cari_et}: ±1",
                          f"{self.cari_et}: küçük fark", f"{self.cari_et}: FARK"])
        self.kayit = []        # konsol raporu için her karşılaştırma
        self.bilgi = {}        # (banka, tür) → sütun seçimi, denklik, birim
        self.sayfa = {}
        for anahtar, _, gorunen in TABLOLAR:            # her veri grubu ayrı sayfa
            ws = self.wb.create_sheet(gorunen)
            ws.append([h.format(cari=self.cari_et) for h in self.KARS_BASLIK])
            self.sayfa[anahtar] = ws
        self.kur = self.wb.create_sheet("Kur Riski")
        self.kur.append([("Para birimi" if h == "Kalem no" else h).format(cari=self.cari_et) for h in self.KARS_BASLIK[:-2]])
        self.dsayfa = {}
        for spec in DIPNOT_TABLOLARI:
            ws = self.wb.create_sheet(spec["sayfa"])
            ws.append([h.format(cari=self.cari_et) for h in self.DIPNOT_BASLIK])
            self.dsayfa[spec["anahtar"]] = ws
        self.gosterge = self.wb.create_sheet("Göstergeler")
        self.gosterge.append([h.format(cari=self.cari_et) for h in self.GOSTERGE_BASLIK])
        self.gd = self.wb.create_sheet("Diğer Dipnotlar")
        self.gd.append([h.format(cari=self.cari_et) for h in self.GD_BASLIK])
        self.gd_kapsam = self.wb.create_sheet("Dipnot Kapsamı")
        self.gd_kapsam.append([h.format(cari=self.cari_et) for h in self.GD_KAPSAM_BASLIK])
        self.gd_eksik = self.wb.create_sheet("PDF'te Bulunamayanlar")
        self.gd_eksik.append(["Banka", "Tür", "Dipnot grubu", "Dipnot", "TBB kodu", "TBB kalemi",
                              "Önceki dönem", "Önceki dönem – TBB", "Önceki dönem – TBB ham", "Birim"])
        self.denk = self.wb.create_sheet("Bilanço denkliği")
        self.denk.append(["Banka", "Tür", "Dönem", "Aktif toplamı (PDF)", "Pasif toplamı (PDF)", "Fark", "Durum"])
        self.eslesmeler = []   # İş kuralları sayfasının eşleşme tablosu

    def ekle(self, banka, tur: str, ok: Okuyucu, pdf_adi: str):
        no, ad, kisa, borsa, tbb_kodu, _ = banka
        S = self.D[tur]
        A = S["A"]
        if tbb_kodu not in A["codes"]:
            print(f"      ! {kisa} TBB {tur} verisinde yok, atlandı")
            return None
        bs = A["codes"].index(tbb_kodu)
        tbb_ad = {str(c): n for c, n, l, t in S["items"]}
        tur_ad = "Solo" if tur == "solo" else "Konsolide"
        birim = "milyon TL" if ok.carpan == 1000 else "bin TL"
        cari_tbbde = self.cari in A["v"]
        bil_sutun = (2, 5)
        bilgi = self.bilgi.setdefault((kisa, tur), {"birim": birim, "bicim": ok.bicim, "sutun": {}, "denk": []})

        for anahtar, tbb_tablo, gorunen in TABLOLAR:
            ws = self.sayfa[anahtar]
            satirlar = ok.tablolar.get(anahtar, [])
            onc = f"{self.yil - 1}-12" if anahtar in ("VARLIKLAR", "YUKUMLULUKLER", "NAZIM", "OZK") \
                else f"{self.yil - 1}-{self.ay:02d}"
            kalemler = [(c, n) for c, n, l, t in S["items"] if t == tbb_tablo]
            esle_f = hizali_esle if anahtar == "OZK" else tbb_esle
            # TBB'de karşılığı olmadığı bilinen satırlar eşlemeye girmez (yoksa sırayı kaydırıp komşu kalemi alır)
            tbb_adlar = [sade(n) for _, n in kalemler]

            def tbb_de_yok(ad_):
                sa_ = sade(ad_)
                return not any(difflib.SequenceMatcher(None, sa_, n_).ratio() >= 0.95 for n_ in tbb_adlar)
            aday = [x for x in satirlar if not x.get("oran")
                    and not (anahtar == "OZK" and beklenen_mi(anahtar, tur, x["ad"]) and tbb_de_yok(x["ad"]))]
            eslesme = esle_f(aday, kalemler)
            if anahtar == "OZK":
                es_id = {id(x[0]) for x in eslesme}
                eslesme += [(x, None, "beklenen", 0.0) for x in satirlar
                            if not x.get("oran") and id(x) not in es_id and beklenen_mi(anahtar, tur, x["ad"])
                            and tbb_de_yok(x["ad"])]
            if anahtar == "OZK":   # oran satırlarını sırasına göre geri koy
                es_map = {id(x[0]): x for x in eslesme}
                eslesme = [es_map[id(x)] if id(x) in es_map else
                           (x, f"oran:{x['oran']}", "oran: ad (TBB göstergesi)", 1.0) for x in satirlar]
            (ci, oi), n = sutun_sec(anahtar, eslesme, A, onc, bs, ok.carpan)
            if anahtar == "VARLIKLAR":
                bil_sutun = (ci, oi)
            bilgi["sutun"][gorunen] = (ci, oi, donem_et(onc))
            say, say_c = Counter(), Counter()
            for s, kod, yontem, puan in eslesme:
                c_ = s["sayilar"][ci] if len(s["sayilar"]) > ci else None
                o_ = s["sayilar"][oi] if len(s["sayilar"]) > oi else None
                if isinstance(kod, str) and kod.startswith("oran:"):
                    gosterge = A.get(kod[5:], {})
                    t = (gosterge.get(onc) or [None] * 999)[bs] if bs < len(gosterge.get(onc) or []) else None
                    tc = ((gosterge.get(self.cari) or [None] * 999)[bs]
                          if cari_tbbde and bs < len(gosterge.get(self.cari) or []) else None)
                else:
                    t = tbb_deger(A, onc, kod, bs)
                    tc = tbb_deger(A, self.cari, kod, bs) if cari_tbbde else None
                if kod and t is None and o_ is not None and not str(kod).startswith("oran:"):
                    t = 0
                hisse = "tam tl" in s["ad"].lower() or "hisse başına" in s["ad"].lower() or bool(s.get("oran"))
                carp = 1 if hisse else ok.carpan

                def karsilastir_(p_, t_):
                    if t_ is None:
                        return ""
                    if hisse:
                        return "" if p_ is None else ("aynı" if abs(abs(p_) - abs(t_)) < 0.006 else "FARK")
                    return fark_notu(p_, round(t_ / carp))
                bek = beklenen_mi(anahtar, tur, s["ad"]) if kod is None else None
                if bek:
                    notu, yontem = "beklenen (kural)", f"beklenen: {bek}"
                    say["beklenen"] += 1
                elif kod is None:
                    notu = "TBB karşılığı yok" if s["sayilar"] else "başlık"
                    say[notu] += 1
                else:
                    notu = karsilastir_(o_, t)
                    say[notu] += 1
                    say["eşleşen"] += 1
                    if t and o_ and (t > 0) != (o_ > 0):
                        notu += "; TBB ters işaretli"
                cnotu = karsilastir_(c_, tc) if kod and tc is not None else ""
                if cnotu:
                    say_c[cnotu] += 1
                self.kayit.append({"banka": kisa, "tur": tur, "tablo": gorunen, "kalem": f"{s['kod']} {s['ad']}".strip(),
                                   "notu": notu, "cnotu": cnotu, "o": o_, "c": c_,
                                   "t": None if t is None else (t if hisse else round(t / carp)),
                                   "tc": None if tc is None else (tc if hisse else round(tc / carp)),
                                   "kod": kod, "onc": donem_et(onc),
                                   "bs": bs, "hedef": ("G", kod[5:]) if str(kod).startswith("oran:") else "A",
                                   "carp": carp, "t_ham": t})
                i = ws.max_row + 1
                bol = f"/{carp}" if carp != 1 else ""
                ws.append([
                    kisa, tur_ad, s["kod"], s["ad"], c_, donem_et(onc), o_,
                    ((f"=ROUND(Q{i}{bol},0)" if bol else f"=Q{i}") if t is not None else None),
                    (f'=IF(OR(G{i}="",H{i}=""),"",ABS(H{i})-ABS(G{i}))' if (t is not None and o_ is not None) else None),
                    notu,
                    ((f"=ROUND(R{i}{bol},0)" if bol else f"=R{i}") if tc is not None else None),
                    (f'=IF(OR(E{i}="",K{i}=""),"",ABS(K{i})-ABS(E{i}))' if (tc is not None and c_ is not None) else None),
                    cnotu, ("oran (%)" if s.get("oran") else "TL") if hisse else birim, kod or "",
                    (tbb_ad.get(str(kod), "") if not str(kod).startswith("oran:") else f"TBB göstergesi: {kod[5:]}")
                    if kod else "", t, tc,
                    *s["sayilar"][:6]])
                bicim = "0.00#" if s.get("oran") else ("0.0000" if hisse else TUTAR)
                for col in "EGHIKLQRSTUVWX":
                    ws[f"{col}{i}"].number_format = bicim
                if notu.startswith("FARK") or cnotu == "FARK":
                    for col in "ABCDEFGHIJ":
                        ws[f"{col}{i}"].fill = TURUNCU
                self.eslesmeler.append([kisa, tur_ad, gorunen, s["kod"], s["ad"], kod or "",
                                        (tbb_ad.get(str(kod), "") if not str(kod).startswith("oran:")
                                         else f"TBB göstergesi: {kod[5:]}") if kod else "", yontem, puan,
                                        f"PDF sütun {ci + 1}", f"PDF sütun {oi + 1}"])
            self.ozet.append([kisa, tur_ad, gorunen, len(satirlar), say["eşleşen"], say["aynı"], say["yuvarlama (±1)"],
                              say["küçük fark (≤3)"], say["FARK"], say["TBB karşılığı yok"], say["beklenen"],
                              "evet" if cari_tbbde else "hayır", say_c["aynı"], say_c["yuvarlama (±1)"],
                              say_c["küçük fark (≤3)"], say_c["FARK"]])
            if say["FARK"]:
                self.ozet.cell(self.ozet.max_row, 9).fill = TURUNCU

        # bilanço denkliği
        ci, oi = bil_sutun
        for et, sira in ((self.cari_et, ci), ("Önceki dönem", oi)):
            akt = self._toplam(ok.tablolar.get("VARLIKLAR", []), "VARLIK", sira)
            pas = self._toplam(ok.tablolar.get("YUKUMLULUKLER", []), "YÜKÜMLÜLÜK", sira)
            fark = None if (akt is None or pas is None) else akt - pas
            bilgi["denk"].append((et, akt, pas, fark))
            self.denk.append([kisa, tur_ad, et, akt, pas, fark,
                              "—" if fark is None else ("denk" if abs(fark) <= 1 else "DENK DEĞİL")])
            for col in "DEF":
                self.denk[f"{col}{self.denk.max_row}"].number_format = TUTAR
            self.denk.cell(self.denk.max_row, 7).fill = YESIL if fark is not None and abs(fark) <= 1 else TURUNCU
        return birim

    def ekle_kur(self, banka, tur: str, ok: "Okuyucu"):
        no, ad, kisa, borsa, tbb_kodu, _ = banka
        A = self.D[tur]["A"]
        if tbb_kodu not in A["codes"]:
            return
        bs = A["codes"].index(tbb_kodu)
        tur_ad = "Solo" if tur == "solo" else "Konsolide"
        satirlar = ok.tablolar.get("KUR", [])
        d = tbb_dipnotu(self.D, tur, "Kur riskine ilişkin bilgiler")
        say = Counter()
        if not d or not satirlar:
            self.ozet.append([kisa, tur_ad, "Kur Riski", len(satirlar), 0, 0, 0, 0, 0, 0, 0,
                              "TBB dipnotu yok" if not d else "tablo okunamadı"])
            return
        onc = f"{self.yil - 1}-12"
        cari_tbbde = self.cari in d["v"]
        d = kur_net_satirlari(d)
        tbb = kur_tbb_satirlari(d)
        birim = "milyon TL" if ok.carpan == 1000 else "bin TL"
        PB = ["Avro", "ABD Doları", "Diğer YP", "Toplam"]

        def esle(donem):
            esl = {}
            for bolum in ("varlik", "yukum", "ozet"):
                P = [r for r in satirlar if r["donem"] == donem and r["bolum"] == bolum]
                T = [t for t in tbb if t["bolum"] == bolum]
                a = [KUR_ESANLAM.get(sade(r["ad"]), sade(r["ad"])) for r in P]
                b = [sade(t["ad"]) for t in T]
                if bolum == "ozet":     # özet satırlarının sırası bankadan bankaya ve TBB'de farklı
                    kullanildi = set()
                    for i, x in enumerate(a):
                        aday = max(((difflib.SequenceMatcher(None, x, y).ratio(), j) for j, y in enumerate(b)
                                    if j not in kullanildi), default=(0, None))
                        if aday[0] >= 0.6:
                            kullanildi.add(aday[1])
                            esl[id(P[i])] = (T[aday[1]], round(aday[0], 2))
                        else:
                            esl[id(P[i])] = (None, 0.0)
                    continue
                for i, j, puan in ad_hizala(a, b, 0.45):
                    esl[id(P[i])] = (T[j], puan) if j is not None else (None, 0.0)
            return esl
        es_c, es_o = esle("cari"), esle("onceki")
        say_c = Counter()
        onceki_kod = {t["kod"]: r for r in satirlar if r["donem"] == "onceki"
                      for t, _ in [es_o.get(id(r), (None, 0))] if t}
        for r in [x for x in satirlar if x["donem"] == "cari"]:
            T, puan = es_c.get(id(r), (None, 0.0))
            o_row = onceki_kod.get(T["kod"]) if T else None
            self.eslesmeler.append([kisa, tur_ad, "Kur Riski", "", r["ad"], T["kod"] if T else "",
                                    T["ad"] if T else "", "hizalı: ad benzerliği" if T else "eşleşmedi", puan,
                                    "Avro, ABD Doları, Diğer YP, Toplam", "önceki dönem satırı" if o_row else "raporda yok"])
            for k, pb in enumerate(PB):
                kod = T["sutun"][k] if T else None
                c_ = r["sayilar"][k]
                o_ = o_row["sayilar"][k] if o_row else None
                t = dn_deger(d, onc, kod, bs) if (kod and o_row) else None
                tc = dn_deger(d, self.cari, kod, bs) if (kod and cari_tbbde) else None
                if not T:
                    notu = "TBB karşılığı yok"
                elif o_row is None:
                    notu = "önceki dönem raporda yok"
                else:
                    notu = fark_notu(o_, None if t is None else round(t / ok.carpan))
                    if t and o_ and (t > 0) != (o_ > 0):
                        notu += "; TBB ters işaretli"
                say[notu.split(";")[0]] += 1
                if T:
                    say["eşleşen"] += 1
                cnotu = fark_notu(c_, round(tc / ok.carpan)) if tc is not None else ""
                if cnotu:
                    say_c[cnotu] += 1
                self.kayit.append({"banka": kisa, "tur": tur, "tablo": "Kur Riski", "kalem": f"{r['ad']} – {pb}",
                                   "notu": notu, "cnotu": cnotu, "o": o_, "c": c_,
                                   "t": None if t is None else round(t / ok.carpan),
                                   "tc": None if tc is None else round(tc / ok.carpan), "kod": kod, "onc": donem_et(onc),
                                   "bs": bs, "hedef": ("DN", "Kur riskine ilişkin bilgiler"), "carp": ok.carpan,
                                   "t_ham": dn_deger(d, onc, kod, bs) if kod else None})
                ws = self.kur
                i = ws.max_row + 1
                bol = f"/{ok.carpan}" if ok.carpan != 1 else ""
                ws.append([kisa, tur_ad, pb, r["ad"], c_, donem_et(onc), o_,
                           ((f"=ROUND(Q{i}{bol},0)" if bol else f"=Q{i}") if t is not None else None),
                           (f'=IF(OR(G{i}="",H{i}=""),"",ABS(H{i})-ABS(G{i}))' if (t is not None and o_ is not None) else None),
                           notu,
                           ((f"=ROUND(R{i}{bol},0)" if bol else f"=R{i}") if tc is not None else None),
                           (f'=IF(OR(E{i}="",K{i}=""),"",ABS(K{i})-ABS(E{i}))' if tc is not None else None),
                           cnotu, birim, kod or "", (T["ad"] + ("" if pb == "Toplam" else f" – {pb}")) if T else "",
                           t, tc, *r["sayilar"][:4]])
                for col in "EGHIKLQRSTUV":
                    ws[f"{col}{i}"].number_format = TUTAR
                if notu.startswith("FARK") or cnotu == "FARK":
                    for col in "ABCDEFGHIJ":
                        ws[f"{col}{i}"].fill = TURUNCU
        self.ozet.append([kisa, tur_ad, "Kur Riski (değer)", 4 * sum(1 for r in satirlar if r["donem"] == "cari"),
                          say["eşleşen"], say["aynı"], say["yuvarlama (±1)"],
                          say["küçük fark (≤3)"], say["FARK"], say["TBB karşılığı yok"], 0,
                          "evet" if cari_tbbde else "hayır", say_c["aynı"], say_c["yuvarlama (±1)"],
                          say_c["küçük fark (≤3)"], say_c["FARK"]])
        if say["FARK"]:
            self.ozet.cell(self.ozet.max_row, 9).fill = TURUNCU

    def ekle_dipnot(self, banka, tur: str, ok: "Okuyucu"):
        """Mali bünye tabloları: faiz oranı riski, ortalama faiz, hisse senedi, LKO, kalan vade, NİFO, kaldıraç."""
        no_, ad_, kisa, borsa, tbb_kodu, _ = banka
        A = self.D[tur]["A"]
        if tbb_kodu not in A["codes"]:
            return
        bs = A["codes"].index(tbb_kodu)
        tur_ad = "Solo" if tur == "solo" else "Konsolide"
        birim = "milyon TL" if ok.carpan == 1000 else "bin TL"
        onc = f"{self.yil - 1}-12"
        bilgi = self.bilgi.setdefault((kisa, tur), {"birim": birim, "bicim": ok.bicim, "sutun": {}, "denk": []})
        bilgi.setdefault("okunan", {})
        for spec in DIPNOT_TABLOLARI:
            ws = self.dsayfa[spec["anahtar"]]
            satirlar = ok.tablolar.get(spec["anahtar"], [])
            d = tbb_dipnotu(self.D, tur, spec["tbb"]) if spec["tbb"] else None
            es = DipnotEsleyici(spec, d) if d else None
            cari_tbbde = bool(d) and self.cari in d["v"]
            tek = spec.get("tek_blok", False)
            sutunlar = spec["sutun"]
            oran_f = spec.get("oran") or (lambda no, ad: False)

            def anahtar(r):
                return r["no"] if r.get("no") is not None else (kanon_kalem(r["ad"]) or sade(r["ad"]))
            # cari ↔ önceki satırları eşle (aynı anahtar birden çok kez geçebilir)
            onceki = defaultdict(list)
            for r in satirlar:
                if r["donem"] == "onceki":
                    onceki[anahtar(r)].append(r)
            ciftler, kullanilan = [], set()
            for r in satirlar:
                if r["donem"] == "cari" or tek:
                    adaylar = [x for x in onceki.get(anahtar(r), []) if id(x) not in kullanilan]
                    o = adaylar[0] if adaylar and not tek else None
                    if o is not None:
                        kullanilan.add(id(o))
                    ciftler.append((r, o))
            ciftler += [(None, r) for r in satirlar if r["donem"] == "onceki" and id(r) not in kullanilan and not tek]

            say, say_c = Counter(), Counter()
            hucre = 0
            for rc, ro in ciftler:
                r0 = rc or ro
                refs, tbb_ad, yontem = es.ref(r0) if es else (None, "", "TBB'de bu tablo yok" if not spec["tbb"]
                                                             else "TBB dipnotu alınamadı")
                no = r0.get("no")
                oran = oran_f(no, r0["ad"])
                carp = 1 if oran else ok.carpan
                self.eslesmeler.append([kisa, tur_ad, spec["sayfa"], no or "", r0["ad"],
                                        ", ".join(str(x if not isinstance(x, tuple) else "toplam") for x in refs if x)
                                        if refs else "", tbb_ad, yontem, "", ", ".join(sutunlar),
                                        "önceki dönem satırı" if (ro or tek) else "raporda yok"])
                for j, sut in enumerate(sutunlar):
                    if tek:
                        c_ = rc["sayilar"][0] if rc else None
                        o_ = rc["sayilar"][1] if rc and len(rc["sayilar"]) > 1 else None
                    else:
                        c_ = rc["sayilar"][j] if rc and j < len(rc["sayilar"]) else None
                        o_ = ro["sayilar"][j] if ro and j < len(ro["sayilar"]) else None
                    if c_ is None and o_ is None:
                        continue
                    hucre += 1
                    ref = refs[j] if refs and j < len(refs) else None
                    t = es.deger(onc, ref, bs) if (es and ref is not None) else None
                    tc = es.deger(self.cari, ref, bs) if (es and ref is not None and cari_tbbde) else None
                    if ref is not None and t is None and d is not None and onc in d["v"]:
                        t = 0
                    if not spec["tbb"]:
                        notu = "yalnız PDF (TBB'de yok)"
                    elif d is None:
                        notu = "TBB dipnotu alınamadı"
                    elif ref is None and refs:
                        notu = "yalnız PDF (TBB'de bu sütun yok)"
                    elif ref is None:
                        notu = "TBB karşılığı yok"
                    elif o_ is None:
                        notu = "önceki dönem raporda yok"
                    else:
                        notu = oran_notu(o_, t) if oran else fark_notu(o_, None if t is None else round(t / carp))
                    say["yalnız PDF" if notu.startswith("yalnız PDF") else notu.split(";")[0]] += 1
                    if ref is not None:
                        say["eşleşen"] += 1
                    cnotu = ""
                    if tc is not None and c_ is not None:
                        cnotu = oran_notu(c_, tc) if oran else fark_notu(c_, round(tc / carp))
                        say_c[cnotu] += 1
                    t_g = None if t is None else (t if oran else round(t / carp))
                    tc_g = None if tc is None else (tc if oran else round(tc / carp))
                    self.kayit.append({"banka": kisa, "tur": tur, "tablo": spec["sayfa"],
                                       "kalem": f"{(str(no) + ' ') if no else ''}{r0['ad']} – {sut}",
                                       "notu": notu, "cnotu": cnotu, "o": o_, "c": c_, "t": t_g, "tc": tc_g,
                                       "kod": ref if not isinstance(ref, tuple) else "toplam", "onc": donem_et(onc),
                                       "bs": bs, "hedef": ("DN", spec["tbb"]) if spec["tbb"] else None, "carp": carp,
                                       "t_ham": t})
                    i = ws.max_row + 1
                    bol = f"/{carp}" if carp != 1 else ""
                    tarih_c = rc["tarih"][j] if rc and rc.get("tarih") else None
                    tarih_o = ro["tarih"][j] if ro and ro.get("tarih") else None
                    if tarih_c or tarih_o:
                        notu = f"{notu} | tarih: {tarih_c or '–'} / önceki: {tarih_o or '–'}"
                    ws.append([kisa, tur_ad, no or "", r0["ad"], sut, c_, donem_et(onc), o_,
                               ((f"=ROUND(R{i}{bol},0)" if bol else f"=R{i}") if t is not None else None),
                               (f'=IF(OR(H{i}="",I{i}=""),"",ABS(I{i})-ABS(H{i}))' if (t is not None and o_ is not None) else None),
                               notu,
                               ((f"=ROUND(S{i}{bol},0)" if bol else f"=S{i}") if tc is not None else None),
                               (f'=IF(OR(F{i}="",L{i}=""),"",ABS(L{i})-ABS(F{i}))' if (tc is not None and c_ is not None) else None),
                               cnotu, "oran (%)" if oran else birim,
                               ("toplam: " + ", ".join(map(str, ref[1]))) if isinstance(ref, tuple) else (ref or ""),
                               (tbb_ad + ("" if isinstance(ref, tuple) or ref is None else
                                          f" → {es.ad(ref)}" if es and es.ad(ref) != tbb_ad else "")) if ref is not None else "",
                               t, tc, r0.get("bolum") or ""])
                    bicim = "0.00" if oran else TUTAR
                    for col in "FHIJLMRS":
                        ws[f"{col}{i}"].number_format = bicim
                    if notu.startswith("FARK") or cnotu == "FARK":
                        for col in "ABCDEFGHIJK":
                            ws[f"{col}{i}"].fill = TURUNCU
            bilgi["okunan"][spec["sayfa"]] = hucre
            self.ozet.append([kisa, tur_ad, spec["sayfa"], hucre, say["eşleşen"], say["aynı"], say["yuvarlama (±1)"],
                              say["küçük fark (≤3)"], say["FARK"], say["TBB karşılığı yok"], 0,
                              ("evet" if cari_tbbde else "hayır") if spec["tbb"] else "TBB'de tablo yok",
                              say_c["aynı"], say_c["yuvarlama (±1)"], say_c["küçük fark (≤3)"], say_c["FARK"]])
            if say["FARK"]:
                self.ozet.cell(self.ozet.max_row, 9).fill = TURUNCU

    def ekle_gosterge(self, banka, tur: str, ok: "Okuyucu"):
        """TBB göstergeleri: YP aktif, YKLV, net nakit çıkışı, LKO, mevcut/gerekli istikrarlı fon, NİFO."""
        no_, ad_, kisa, borsa, tbb_kodu, _ = banka
        A = self.D[tur]["A"]
        if tbb_kodu not in A["codes"]:
            return
        bs = A["codes"].index(tbb_kodu)
        tur_ad = "Solo" if tur == "solo" else "Konsolide"
        birim = "milyon TL" if ok.carpan == 1000 else "bin TL"
        onc = f"{self.yil - 1}-12"
        bilgi = self.bilgi[(kisa, tur)]
        ci, oi, _ = bilgi["sutun"].get("Aktif", (2, 5, ""))
        varl = ok.tablolar.get("VARLIKLAR", [])
        lko = {(r["donem"], r["no"]): r["sayilar"] for r in ok.tablolar.get("LKO", [])}
        nifo = {(r["donem"], r["no"]): r["sayilar"] for r in ok.tablolar.get("NIFO", [])}

        def al(t, dn, no, j):
            x = t.get((dn, no))
            return x[j] if x and j < len(x) else None
        satirlar = [
            ("YP aktif toplamı", "yp", "Bilanço, aktif toplamı YP sütunu",
             self._toplam(varl, "VARLIK", ci - 1), self._toplam(varl, "VARLIK", oi - 1), False),
            ("Yüksek kaliteli likit varlık stoku (TP+YP)", "hqla", "LKO satır 21, uygulanmış TP+YP",
             al(lko, "cari", 21, 2), al(lko, "onceki", 21, 2), False),
            ("Toplam net nakit çıkışları (TP+YP)", "out", "LKO satır 22, uygulanmış TP+YP",
             al(lko, "cari", 22, 2), al(lko, "onceki", 22, 2), False),
            ("Likidite karşılama oranı TP+YP (%)", "lcr", "LKO satır 23", al(lko, "cari", 23, 2), al(lko, "onceki", 23, 2), True),
            ("Likidite karşılama oranı YP (%)", "lcrYp", "LKO satır 23", al(lko, "cari", 23, 3), al(lko, "onceki", 23, 3), True),
            ("Mevcut istikrarlı fon", "asf", "NİFO satır 14, d", al(nifo, "cari", 14, 4), al(nifo, "onceki", 14, 4), False),
            ("Gerekli istikrarlı fon", "rsf", "NİFO satır 33, d", al(nifo, "cari", 33, 4), al(nifo, "onceki", 33, 4), False),
            ("Net istikrarlı fonlama oranı (%)", "nsfr", "NİFO satır 34, d", al(nifo, "cari", 34, 4), al(nifo, "onceki", 34, 4), True),
        ]
        say, say_c = Counter(), Counter()
        for ad, anahtar, kaynak, c_, o_, oran in satirlar:
            seri = A.get(anahtar) or {}

            def tbb(dn):
                x = seri.get(dn) if isinstance(seri, dict) else None
                return x[bs] if x and bs < len(x) else None
            t, tc = tbb(onc), tbb(self.cari)
            carp = 1 if oran else ok.carpan
            notu = "PDF'te okunamadı" if o_ is None else ("TBB'de yok" if t is None else
                                                          (oran_notu(o_, t) if oran else fark_notu(o_, round(t / carp))))
            cnotu = "" if (tc is None or c_ is None) else (oran_notu(c_, tc) if oran else fark_notu(c_, round(tc / carp)))
            say[notu] += 1
            if cnotu:
                say_c[cnotu] += 1
            self.kayit.append({"banka": kisa, "tur": tur, "tablo": "Göstergeler", "kalem": ad, "notu": notu,
                               "cnotu": cnotu, "o": o_, "c": c_, "t": None if t is None else (t if oran else round(t / carp)),
                               "tc": None if tc is None else (tc if oran else round(tc / carp)),
                               "kod": f"gösterge:{anahtar}", "onc": donem_et(onc),
                               "bs": bs, "hedef": ("G", anahtar), "carp": carp, "t_ham": t})
            i = self.gosterge.max_row + 1
            bol = "" if oran else f"/{carp}"
            self.gosterge.append([kisa, tur_ad, ad, anahtar, kaynak, c_, donem_et(onc), o_,
                                  (f"=ROUND(P{i}{bol},{2 if oran else 0})" if t is not None else None),
                                  (f'=IF(OR(H{i}="",I{i}=""),"",ABS(I{i})-ABS(H{i}))' if t is not None and o_ is not None else None),
                                  notu, (f"=ROUND(Q{i}{bol},{2 if oran else 0})" if tc is not None else None),
                                  (f'=IF(OR(F{i}="",L{i}=""),"",ABS(L{i})-ABS(F{i}))' if tc is not None and c_ is not None else None),
                                  cnotu, "oran (%)" if oran else birim, t, tc])
            for col in "FHIJLMPQ":
                self.gosterge[f"{col}{i}"].number_format = "0.00" if oran else TUTAR
            if notu == "FARK" or cnotu == "FARK":
                for col in "ABCDEFGHIJK":
                    self.gosterge[f"{col}{i}"].fill = TURUNCU
        bilgi.setdefault("okunan", {})["Göstergeler"] = len(satirlar)
        self.ozet.append([kisa, tur_ad, "Göstergeler", len(satirlar), sum(1 for x in satirlar),
                          say["aynı"], say["yuvarlama (±1)"], say["küçük fark (≤3)"], say["FARK"],
                          say["TBB'de yok"] + say["PDF'te okunamadı"], 0,
                          "evet" if self.cari in (A.get("hqla") or {}) else "hayır",
                          say_c["aynı"], say_c["yuvarlama (±1)"], say_c["küçük fark (≤3)"], say_c["FARK"]])

    def ekle_genel(self, banka, tur: str, ok: "Okuyucu"):
        """Özel okuyucusu olmayan bütün TBB dipnotları ve mevduat vade yapısı (genel eşleyici)."""
        no_, ad_, kisa, borsa, tbb_kodu, _ = banka
        A = self.D[tur]["A"]
        if tbb_kodu not in A["codes"]:
            return
        bs = A["codes"].index(tbb_kodu)
        tur_ad = "Solo" if tur == "solo" else "Konsolide"
        birim_ad = "milyon TL" if ok.carpan == 1000 else "bin TL"
        donemler = [f"{self.yil - 1}-12"] + ([f"{self.yil - 1}-{self.ay:02d}"] if self.ay != 12 else [])
        gd = GenelDipnot(ok)
        listeler = [("Ana tablo", MEVDUAT_VADE, mevduat_vade_dipnotu(self.D, tur))]
        for x in self.D[tur].get("DN", []):
            if x["ad"] not in OZEL_DIPNOTLAR:
                listeler.append((x.get("g", ""), x["ad"], None))
        toplam = Counter()
        konsol = []
        for grup, ad, d in listeler:
            if d is None:
                d = tbb_dipnotu(self.D, tur, ad)
            if d is None:
                continue
            ceyrek = f"{self.yil - 1}-{self.ay:02d}"
            ara_donemler = [ceyrek] + [f"{self.yil}-{m:02d}" for m in (3, 6, 9) if m < self.ay]
            ara_donemde = self.ay == 12 or any(dn_deger(d, p_, k, bs) for p_ in ara_donemler for k, *_ in d["items"])
            if not ara_donemde:
                dolu_y = [k for k, *_ in d["items"] if dn_deger(d, donemler[0], k, bs)]
                if dolu_y:
                    self.gd_kapsam.append([kisa, tur_ad, grup, ad, donem_et(donemler[0]), len(dolu_y), "", "", "", "", "", "",
                                           "", "", "", "", "", "yalnız yıl sonu raporlarında (ara dönemde TBB'de de yok)"])
                    toplam.update({"dipnot": 1, "yilsonu": 1})
                    konsol.append((ad, len(dolu_y), 0, 0, 0, "yılsonu"))
                continue
            mevduat_tablosu = ad == MEVDUAT_VADE or ad.startswith(MEVDUAT_VADE_DN)
            if mevduat_tablosu:
                sonuc = mevduat_vade_kayit(ok, d, bs, donemler)
            elif ad.startswith(RG_DIPNOTLAR):
                sonuc = risk_grubu_kayit(ok, ad, d, bs, donemler)
            elif ad.startswith(SG_DIPNOTLAR):
                sonuc = sigorta_kayit(ok, ad, d, bs, donemler)
            elif ad.startswith(MF_DIPNOT):
                sonuc = mevduat_faiz_kayit(ok, ad, d, bs, donemler)
            elif ad.startswith(DA_DIPNOTLAR):
                sonuc = donuk_alacak_kayit(ok, ad, d, bs, donemler)
            elif ad.startswith(BASIT_DIPNOTLAR):
                sonuc = basit_tablo_kayit(ok, ad, d, bs, donemler)
                if any(ad.startswith(s_["tbb"]) and s_.get("yalniz_eslesen") for s_ in BASIT_TABLOLAR):
                    g_ = gd.esle(d, bs, ok.carpan, donemler)            # genel eşleyicinin bulduklarıyla tamamla
                    var = {x["kod"] for x in sonuc["kayit"]}
                    sonuc["kayit"] += [x for x in g_["kayit"] if x["kod"] not in var and not x.get("donem")
                                       and g_["donem"] == sonuc["donem"]]
            elif ad.startswith(KR_DIPNOTLAR):
                sonuc = kredi_tablo_kayit(ok, ad, d, bs, donemler)
            elif ad.startswith(YI_DIPNOTLAR):
                sonuc = yakin_izleme_kayit(ok, ad, d, bs, donemler)
            else:
                sonuc = gd.esle(d, bs, ok.carpan, donemler)
                if ad.startswith(GG_DIPNOT) and sum(1 for x in sonuc["kayit"] if x["c"] is not None) < 4:
                    s2 = grup_gayrinakdi_kayit(ok, d, bs, donemler)
                    if len(s2["kayit"]) > len(sonuc["kayit"]):
                        sonuc = s2
            onc = sonuc["donem"]
            ad_of = {k: a for k, a, *_ in d["items"]}
            cari_tbbde = self.cari in d["v"]
            say, say_c = Counter(), Counter()
            bulunan = set()
            for k in sonuc["kayit"]:
                carp = k["birim"]
                oran = carp == 1
                t = dn_deger(d, k.get("donem") or onc, k["kod"], bs)
                tc = dn_deger(d, self.cari, k["kod"], bs) if cari_tbbde else None
                o_, c_ = k["o"], k["c"]
                notu = (oran_notu(o_, t) if oran else fark_notu(o_, round(t / carp))) if t is not None else "TBB'de yok"
                cnotu = "" if (tc is None or c_ is None) else (oran_notu(c_, tc) if oran else fark_notu(c_, round(tc / carp)))
                say[notu] += 1
                if c_ is not None:
                    say["cari"] += 1
                if cnotu:
                    say_c[cnotu] += 1
                if t:
                    bulunan.add(k["kod"])
                kalem = " › ".join(x for x in (k["yol"], k["satir"]) if x)
                self.kayit.append({"banka": kisa, "tur": tur, "tablo": "Diğer Dipnotlar",
                                   "kalem": f"{ad[:40]} | {kalem}" + (f" – {k['sutun']}" if k["sutun"] else ""),
                                   "notu": notu, "cnotu": cnotu, "o": o_, "c": c_,
                                   "t": None if t is None else (t if oran else round(t / carp)),
                                   "tc": None if tc is None else (tc if oran else round(tc / carp)),
                                   "kod": k["kod"], "onc": donem_et(onc),
                                   "bs": bs, "hedef": "A" if ad == MEVDUAT_VADE else ("DN", ad), "carp": carp,
                                   "t_ham": t, "dipnot": ad, "tbb_kalem": kalem + (f" – {k['sutun']}" if k["sutun"] else ""),
                                   "pdf_sayfa": k.get("sayfa"), "pdf_etiket": k.get("etiket") or "",
                                   "yontem": k.get("kaynak") or k.get("duzen") or ""})
                i = self.gd.max_row + 1
                bol = "" if oran else f"/{carp}"
                self.gd.append([kisa, tur_ad, grup, ad, kalem, k["sutun"], c_, donem_et(onc), o_,
                                (f"=ROUND(V{i}{bol},{2 if oran else 0})" if t is not None else None),
                                (f'=IF(OR(I{i}="",J{i}=""),"",ABS(J{i})-ABS(I{i}))' if t is not None else None),
                                notu, (f"=ROUND(W{i}{bol},{2 if oran else 0})" if tc is not None else None),
                                (f'=IF(OR(G{i}="",M{i}=""),"",ABS(M{i})-ABS(G{i}))' if tc is not None and c_ is not None else None),
                                cnotu, "oran (%)" if oran else birim_ad, k["kod"], k["sayfa"], k["etiket"][:120],
                                k["kaynak"], k["duzen"], t, tc])
                for col in "GIJKMNVW":
                    self.gd[f"{col}{i}"].number_format = "0.00" if oran else TUTAR
                if notu == "FARK" or cnotu == "FARK":
                    for col in "ABCDEFGHIJKL":
                        self.gd[f"{col}{i}"].fill = TURUNCU
            dolu = sonuc["dolu"]
            eksik = [k for k in dolu if k not in bulunan]
            for k in eksik:
                t = dn_deger(d, onc, k, bs)
                self.gd_eksik.append([kisa, tur_ad, grup, ad, k, ad_of.get(k, ""), donem_et(onc) if onc else "",
                                      round(t / ok.carpan) if isinstance(t, int) else t, t, birim_ad])
            kap = round(100 * len(bulunan) / len(dolu), 1) if dolu else None
            durum = ("TBB'de bu banka için veri yok" if not dolu else
                     "PDF'te bulunamadı (ara dönem raporunda yer almıyor olabilir)" if not bulunan else
                     "tam" if not eksik else "kısmi")
            self.gd_kapsam.append([kisa, tur_ad, grup, ad, donem_et(onc) if onc else "", len(dolu), len(bulunan), kap,
                                   say["cari"], say["aynı"], say["yuvarlama (±1)"], say["küçük fark (≤3)"], say["FARK"],
                                   say_c["aynı"], say_c["yuvarlama (±1)"], say_c["küçük fark (≤3)"], say_c["FARK"], durum])
            if say["FARK"]:
                self.gd_kapsam.cell(self.gd_kapsam.max_row, 13).fill = TURUNCU
            toplam.update({"dipnot": 1, "dolu": len(dolu), "bulunan": len(bulunan), "kayit": len(sonuc["kayit"]),
                           "cari": say["cari"], "aynı": say["aynı"], "±1": say["yuvarlama (±1)"],
                           "küçük": say["küçük fark (≤3)"], "FARK": say["FARK"], "c_aynı": say_c["aynı"],
                           "c_±1": say_c["yuvarlama (±1)"], "c_küçük": say_c["küçük fark (≤3)"],
                           "c_FARK": say_c["FARK"], "veri_var": 1 if dolu else 0,
                           "hic": 1 if dolu and not bulunan else 0, "cari_tbbde": 1 if cari_tbbde else 0})
            konsol.append((ad, len(dolu), len(bulunan), say["cari"], say["FARK"] + say_c["FARK"], ""))
        bilgi = self.bilgi[(kisa, tur)]
        bilgi["genel"] = (toplam, konsol)
        bilgi.setdefault("okunan", {})["Diğer Dipnotlar"] = toplam["kayit"]
        self.ozet.append([kisa, tur_ad, "Diğer Dipnotlar", toplam["kayit"], toplam["kayit"], toplam["aynı"], toplam["±1"],
                          toplam["küçük"], toplam["FARK"], toplam["dolu"] - toplam["bulunan"], 0,
                          "evet" if toplam["cari_tbbde"] else "hayır",
                          toplam["c_aynı"], toplam["c_±1"], toplam["c_küçük"], toplam["c_FARK"]])

    @staticmethod
    def _toplam(satirlar, kelime, sira):
        for s in reversed(satirlar):
            if "TOPLAM" in s["ad"].upper() and kelime in s["ad"].upper() and len(s["sayilar"]) > sira:
                return s["sayilar"][sira]
        return None

    def _is_kurallari(self):
        ws = self.wb.create_sheet("İş kuralları ve eşleşmeler")
        ws.append(["İş kuralları"])
        ws["A1"].font = Font(bold=True, size=12)
        ws.append(["No", "Kural", "Açıklama"])
        for c in range(1, 4):
            ws.cell(2, c).fill = BASLIK
            ws.cell(2, c).font = Font(bold=True)
        for no, (ad, aciklama) in enumerate(KURALLAR, 1):
            ws.append([no, ad, aciklama])
            ws.cell(ws.max_row, 3).alignment = Alignment(wrap_text=True, vertical="top")
            ws.cell(ws.max_row, 2).font = Font(bold=True)
        ws.append([])
        ws.append(["Kalem eşleşmeleri (PDF kalemi → TBB kalemi)"])
        ws.cell(ws.max_row, 1).font = Font(bold=True, size=12)
        bas = ws.max_row + 1
        ws.append(["Banka", "Tür", "Tablo", "PDF kalem no", "Kalem adı (PDF)", "TBB kodu", "Kalem adı (TBB)",
                   "Eşleşme yöntemi", "Eşleşme puanı", "Cari dönem sütunu", "Önceki dönem sütunu"])
        for c in range(1, 12):
            ws.cell(bas, c).fill = BASLIK
            ws.cell(bas, c).font = Font(bold=True)
        for r in self.eslesmeler:
            ws.append(r)
            if r[7] == "eşleşmedi" and r[3]:
                ws.cell(ws.max_row, 8).fill = TURUNCU
        for i, w in enumerate([14, 22, 90, 12, 50, 10, 50, 30, 12, 16, 18], 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = None

    ANA_TABLOLAR_HARITA = {"Aktif", "Pasif", "Nazım", "Gelir-Gider", "Kâr-Zarar ve DKG", "Nakit Akış", "Özkaynak"}

    def _esleme_haritasi(self):
        """Bütün PDF kalemi → TBB kalemi eşlemeleri tek sayfada: ana tablolar, kur riski, mali bünye, dipnotlar, göstergeler."""
        ws = self.wb.create_sheet("Eşleme haritası")
        bas = ["Banka", "Tür", "Kapsam", "Tablo / dipnot", "PDF sayfa", "PDF kalemi", "PDF sütunu", "TBB kodu",
               "TBB kalemi", "Eşleme yöntemi", "Eylül 2026 – PDF", "Önceki dönem", "Önceki dönem – PDF",
               "Önceki dönem – TBB", "Not"]
        ws.append(bas)
        for c in range(1, len(bas) + 1):
            ws.cell(1, c).fill = BASLIK
            ws.cell(1, c).font = Font(bold=True)
        tur_kod = {"Solo": "solo", "Konsolide": "kons"}
        indeks = defaultdict(list)
        for k in self.kayit:
            indeks[(k["banka"], k["tur"], k["tablo"], str(k["kod"]))].append(k)
        for r in self.eslesmeler:
            kisa, tur_ad, tablo, no, pdf_ad, kod, tbb_ad, yontem, puan, csut, osut = (list(r) + [""] * 11)[:11]
            if not kod:
                continue
            kk = [x for kd in str(kod).split(", ") for x in indeks.get((kisa, tur_kod.get(tur_ad, tur_ad), tablo, kd), [])]
            kapsam = ("Kur riski" if tablo == "Kur Riski" else "Ana tablo" if tablo in self.ANA_TABLOLAR_HARITA
                      else "Mali bünye")
            yon = f"{yontem} ({puan})" if puan not in ("", None) else yontem
            if len(kk) > 1:                       # çok sütunlu satır (NİFO, kalan vade, kur …): sütun başına bir satır
                for k in kk:
                    ws.append([kisa, tur_ad, kapsam, tablo, "", f"{no} {pdf_ad}".strip(), k["kalem"], kod, tbb_ad, yon,
                               k["c"], k["onc"], k["o"], k["t"], k["notu"]])
                continue
            k = kk[0] if kk else None
            ws.append([kisa, tur_ad, kapsam, tablo, "", f"{no} {pdf_ad}".strip(), f"cari: {csut} / önceki: {osut}"
                       if csut or osut else "", kod, tbb_ad, yon,
                       k["c"] if k else None, k["onc"] if k else "", k["o"] if k else None, k["t"] if k else None,
                       k["notu"] if k else ""])
        for k in self.kayit:
            if k["tablo"] not in ("Diğer Dipnotlar", "Göstergeler"):
                continue
            tur_ad = "Solo" if k["tur"] == "solo" else "Konsolide"
            if k["tablo"] == "Göstergeler":
                ws.append([k["banka"], tur_ad, "Gösterge", "TBB göstergeleri", "", k["kalem"], "", k["kod"], k["kalem"],
                           "", k["c"], k["onc"], k["o"], k["t"], k["notu"]])
                continue
            ws.append([k["banka"], tur_ad, "Dipnot", k.get("dipnot", ""), k.get("pdf_sayfa"), k.get("pdf_etiket", ""), "",
                       k["kod"], k.get("tbb_kalem", k["kalem"]), k.get("yontem", ""), k["c"], k["onc"], k["o"], k["t"],
                       k["notu"]])
        for i, w in enumerate([13, 10, 11, 44, 8, 50, 22, 10, 60, 30, 15, 12, 15, 15, 16], 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        for row in ws.iter_rows(min_row=2, min_col=11, max_col=14):
            for c in row:
                if isinstance(c.value, (int, float)):
                    c.number_format = "#,##0.##"
        ws.auto_filter.ref = f"A1:O{ws.max_row}"
        ws.freeze_panes = "A2"

    def kaydet(self, yol: Path):
        bicimle(self.ozet, [16, 11, 18, 11, 13, 9, 12, 12, 8, 14, 12, 14, 12, 12, 12, 12])
        for ws in self.sayfa.values():
            bicimle(ws, [14, 11, 9, 50, 16, 13, 16, 16, 12, 22, 16, 12, 14, 10, 9, 50, 20, 20, 14, 14, 14, 14, 14, 14], "E2")
        bicimle(self.kur, [14, 11, 11, 46, 16, 13, 16, 16, 12, 22, 16, 12, 14, 10, 9, 50, 20, 20, 14, 14, 14, 14], "E2")
        bicimle(self.gd, [14, 11, 22, 40, 50, 18, 15, 13, 15, 15, 12, 18, 15, 12, 14, 10, 10, 8, 40, 18, 7, 18, 18], "G2")
        bicimle(self.gd_kapsam, [14, 11, 24, 60, 13, 12, 12, 10, 12, 8, 10, 10, 8, 10, 8, 8, 8, 40], "E2")
        bicimle(self.gd_eksik, [14, 11, 24, 60, 10, 60, 13, 16, 18, 10], "E2")
        bicimle(self.gosterge, [14, 11, 40, 10, 34, 15, 13, 15, 15, 12, 18, 15, 12, 14, 10, 18, 18], "F2")
        for ws in self.dsayfa.values():
            bicimle(ws, [14, 11, 8, 46, 20, 15, 13, 15, 15, 12, 24, 15, 12, 14, 10, 12, 46, 18, 18, 16], "F2")
        bicimle(self.denk, [16, 11, 14, 20, 20, 12, 14])
        self._is_kurallari()
        banka_kurallari_sayfasi(self.wb)
        self._esleme_haritasi()
        self.wb.save(yol)


# ---------------------------------------------------------------------------
# Banka banka iş kuralları (Excel "Banka iş kuralları" sayfası). Tür: "Okuma" = PDF düzenine göre okuma kuralı,
# "TBB" = TBB verisinin bu bankada farklı tuttuğu kalem (bizim okumamız değil), "Eşleme" = PDF → TBB kalem eşlemesi.
BANKA_KURALLARI = [
    # (banka, kapsam, tablo/dipnot, tür, kural, kodda)
    ("Tüm bankalar", "solo+kons", "Gelir tablosu", "Eşleme",
     "İşaret banka banka belirlenir: bankanın önceki dönem değeri TBB'dekiyle karşılaştırılır. Önceki yıl değeri yoksa, "
     "yıl başından birikimli tutar Haziran'ı aşıyorsa bankanın Haziran TBB işareti esas alınır.", "eylul_ekle.isaret"),
    ("Tüm bankalar", "solo+kons", "Tüketici / taksitli ticari krediler (2.f.4, 2.f.5)", "TBB",
     "TBB 'Kısa vadeli' kalemi = kısa vadeli + faiz ve gelir tahakkuku; 'Orta ve uzun vadeli' = orta-uzun sütunu.",
     "kredi_tablo_kayit"),
    ("Tüm bankalar", "solo+kons", "Şablon dipnotlar (2.a.1, 2.a.3, 2.b.3, 2.h, 2.ı.8)", "TBB",
     "TBB tutarı = TP + YP. Gelir dipnotları önceki yılın aynı çeyreğiyle karşılaştırılır.", "basit_tablo_kayit"),
    ("Tüm bankalar", "solo+kons", "Genel dipnotlar", "Eşleme",
     "Kalem önce önceki dönem değeriyle bulunur; banka önceki dönemi yeniden düzenlemişse aynı tablodaki satırlarda adla "
     "(ya da adı çok benzer ve değeri %3 yakın satırlar çapa alınarak) aranır. Cari değer komşu satırın sütun kuralıyla alınır.",
     "GenelDipnot.ad_tamamla / ad_esle_yakin"),
    ("Tüm bankalar", "solo+kons", "Genel dipnotlar", "Okuma",
     "'Cari Dönem' ve 'Önceki Dönem' başlıkları arada rakamlı satır olmadan alt alta yazılmışsa tablo yan yanadır.",
     "GenelDipnot._satirlar"),
    ("Tüm bankalar", "solo+kons", "Özkaynak tablosu", "Okuma",
     "Rakamsız bir satır ')' ya da 'tutar / kısmı / toplamı' ile bitiyor ve alttaki satır yeni bir kalem adıyla başlıyorsa "
     "ayrı kalemlerdir (birleştirilmez).", "Okuyucu._ozkaynak.yeni_kalem"),
    ("Tüm bankalar", "solo+kons", "Mevduat vade yapısı", "Okuma",
     "Rakamlar sağa yaslı: bir sütunda tek kalan rakam, sağ kenarı hizalı olduğu sütuna verilir.", "mevduat_vade_oku"),

    ("Ziraat", "solo+kons", "Gelir tablosu – vergi karşılığı", "Okuma",
     "Cari / ertelenmiş vergi giderleri PDF'te parantezli (negatif) yazılır; TBB'de de negatif. Çoğunluk kuralıyla ters "
     "çevriliyordu, banka kuralıyla düzeldi.", "eylul_ekle.isaret"),
    ("Ziraat", "solo+kons", "KAP", "Okuma", "KAP OID borsa kodundan bulunamaz; aramayla bulunup kap_oid.json'a yazılır. "
     "Ziraat Katılım elenir.", "kap_oid"),
    ("Ziraat", "solo+kons", "Değer artış fonu (2.ı.8)", "TBB",
     "TBB 'GUD farkı DKG yansıtılan FV' satırını 'İştirakler…' kalemine yazmış; eşleme önceki dönem değeriyle yapılır.",
     "basit_tablo_kayit"),

    ("Halkbank", "solo+kons", "Gelir tablosu – vergi", "Okuma",
     "Cari vergi karşılığı PDF'te pozitif yazılır, TBB negatif tutar; işaret banka kuralıyla çevrilir.", "eylul_ekle.isaret"),
    ("Halkbank", "solo+kons", "Kredi tabloları (2.f.4)", "Okuma",
     "Önceki dönem tablosu ertesi sayfada, sayfa başında '30 Eylül 2026 tarihinde sona eren…' satırı var; bu satır yeni "
     "bölüm sayılmaz.", "_st_yeni_baslik"),
    ("Halkbank", "solo+kons", "NİFO / kaldıraç", "Okuma",
     "Satır numarası tek başına bir satırda; sayfa numarası sanılmaz. Kaldıraç dönem başlığındaki '(1)' atfı atılır.", "_matris"),
    ("Halkbank", "solo+kons", "Karşılık tablosu (2.f.2.i alt)", "Okuma",
     "12 aylık beklenen zarar / önemli artış tablosu okunamıyor (eksik).", "yakin_izleme_oku"),

    ("VakıfBank", "solo+kons", "Gelir tablosu – vergi karşılığı", "Okuma",
     "Ziraat gibi: vergi giderleri PDF'te negatif; banka kuralıyla işaret korunur.", "eylul_ekle.isaret"),
    ("VakıfBank", "solo+kons", "Sigorta kapsamı (2.a.2.i / 2.a.4.i)", "Okuma",
     "Sütun sırası dönem önce: (kapsam cari, aşan cari, kapsam önceki, aşan önceki).", "sigorta_oku"),
    ("VakıfBank", "solo+kons", "Mevduata ödenen faiz (2.b.4.i)", "Okuma",
     "'Kadar' başlıkları birkaç satıra bölünmüş; 'Yabancı Para: - - -' satırı bölüm başlığıdır.", "mevduat_faiz_oku"),
    ("VakıfBank", "solo+kons", "Donuk alacak iv)", "Okuma",
     "Dönem başlığı 'Cari Dönem - 30 Eylül 2026 Krediler Krediler Krediler'; net satırı 'Net Değer'.", "donuk_alacak_oku"),
    ("VakıfBank", "solo+kons", "2.b.3, 2.a.3, 2.ı.8", "Bilgi",
     "Ayrı dipnot tablosu yok (tutarlar gelir tablosunda); Eylül hücreleri boş kalır.", ""),
    ("VakıfBank", "solo", "Sigorta kapsamı", "TBB",
     "Aralık 2025'te limiti aşan DTH / diğer hesaplar toplamı doğru ama satırlara yanlış dağıtılmış.", ""),

    ("İş Bankası", "solo+kons", "Bütün tablolar", "Okuma",
     "Boş hücrelere '-' yazılmaz: sütunlar başlık kelimelerinin konumundan bulunur, okunan satırın rakamsız sütunları 0.", ""),
    ("İş Bankası", "solo+kons", "Kredi tabloları (e.2, e.3, e.4)", "Bilgi",
     "Yalnız cari dönem verilir (önceki dönem tablosu yok) → önceki dönemle TBB'ye karşı sınanamaz.", "kredi_tablo_oku"),
    ("İş Bankası", "solo+kons", "Özkaynak", "Okuma",
     "'… (Geçici Madde 4 kapsamında olanlar)' ile 'Karşılıklar (…8'inci madde…)' satırları ayrı kalemdir; "
     "'Uygulanacak indirim esaslarında aşım tutarının altında kalan tutarlar' satırları da ayrı.", "yeni_kalem"),
    ("İş Bankası", "solo+kons", "Özkaynak – genel karşılıklar", "Bilgi",
     "Temmuz 2026'dan itibaren İDDY: standart yaklaşım genel karşılıkları 56.274 → 891; Karşılıklar = 891 + 13.508.", ""),
    ("İş Bankası", "solo", "Özkaynak", "TBB",
     "Haziran 2026'da 'Ana sermaye ve katkı sermaye toplamı' kalemine 38 (toplam özkaynaktan indirimler) yazılmış.", ""),
    ("İş Bankası", "kons", "Özkaynak", "TBB",
     "Haziran 2026'da '%10'dan fazla paylı kuruluşlar' tutarı '%10 veya daha az' kalemine yazılmış (PDF: 5.337 '%10'dan fazla').", ""),
    ("İş Bankası", "solo+kons", "Mevduat vade", "Okuma",
     "Yurt içi bankalar vadesiz (99 / 260) dar '7 gün ihbarlı' sütununa düşüyordu; sağ kenar hizasıyla düzeltilir.",
     "mevduat_vade_oku"),
    ("İş Bankası", "solo+kons", "Donuk alacak i)", "Okuma",
     "'(Karşılıklardan Önceki Brüt Tutarlar)' parantez içinde; tablo 'donuk alacaklardan … yeniden yapılandır' başlığıyla bulunur.",
     "donuk_alacak_oku"),
    ("İş Bankası", "solo+kons", "Sermaye benzeri borçlanma (2.h)", "Bilgi", "Dipnot tablosu yok; TBB'de de Haziran'da 0.", ""),
    ("İş Bankası", "solo+kons", "Risk grubu", "TBB", "'Diğer gerçek ve tüzel kişiler' dönem sonu TBB'de farklı.", ""),

    ("Garanti BBVA", "solo+kons", "Özkaynak", "Okuma",
     "Bazı satırlarda rakamlar adın üstündeki satırda (Katkı sermaye toplamı, ÇSYO, tampon oranı); '(Para birimi: …)' ve "
     "'… Finansal Raporu' sayfa başlıkları atlanır.", "Okuyucu._ozkaynak"),
    ("Garanti BBVA", "kons", "Genel dipnotlar", "Okuma",
     "'Cari Dönem' / 'Önceki Dönem' başlıkları alt alta (yan yana TP/YP tablosu). Önceki dönem yeniden düzenlenmiş: "
     "Eylül 2025 değerleri TBB'yle tutmaz → ad eşlemesi (2.b.1, 2.ç, 2.f).", "GenelDipnot"),
    ("Garanti BBVA", "solo+kons", "Kredili mevduat hesabı (2.f.4)", "Eşleme",
     "'Gerçek Kişi' ve 'Personel' KMH satırları toplanır (solo TBB böyle; konsolide TBB personeli YP kalemine yazmış).",
     "kredi_tablo_oku"),
    ("Garanti BBVA", "solo+kons", "Yakın izleme (2.f.2.i)", "Okuma",
     "Tablo başlığı 'Birinci ve ikinci aşama krediler…'.", "YI_BASLIK"),
    ("Garanti BBVA", "solo+kons", "Donuk alacak", "Okuma",
     "Muhasebe politikası sayfaları da 'donuk alacak' geçiyor; tablo sayfası grup adları + 'intikal' ile seçilir. ii) III. grup "
     "'Satılan'/'Diğer' TBB'de farklı satırda.", "donuk_alacak_oku"),
    ("Garanti BBVA", "kons", "Mevduata ödenen faiz", "Okuma", "Vade sütunları '1-3 Ay / 3-6 Ay / 6-12 Ay'. YP DTH / bankalar "
     "önceki dönem TBB'de farklı.", "mevduat_faiz_oku"),
    ("Garanti BBVA", "kons", "Mevduat vade", "Okuma",
     "Resmi kuruluşlar satırındaki küçük rakamlar (94 | 1 | 2) sütun ortasına göre bir sağa kayıyor; sağ kenar hizasıyla "
     "yerleştirilir.", "mevduat_vade_oku"),
    ("VakıfBank", "solo+kons", "Gayrinakdi kredilerin sektör dağılımı (2.a.3.ii)", "Okuma",
     "Her tutarın yanında yüzde payı sütunu var (tutar | % | tutar | %); iki ondalıklı ≤100 sütunlar atılır.",
     "GenelDipnot._satirlar"),
    ("VakıfBank", "solo+kons", "I ve II. grup gayrinakdi krediler (2.a.3.iii)", "Eşleme",
     "Satır × (I. grup TP, YP | II. grup TP, YP); TBB grup tutarı = TP + YP.", "GenelDipnot._turet"),
    ("İş Bankası", "solo+kons", "I ve II. grup gayrinakdi krediler (2.a.3.iii)", "Okuma",
     "Yalnız cari dönem, boş hücreler '-' değil → grup tablosu okuyucusu (TP/YP başlık konumu).", "grup_gayrinakdi_kayit"),
    ("Garanti BBVA", "solo+kons", "Sermaye benzeri borçlanma (2.h)", "Bilgi", "Dipnot tablosu yok; TBB'de de Haziran'da 0.", ""),
    ("Garanti BBVA", "solo", "Özkaynak", "TBB", "Aralık 2025'te bazı kalemler bir satır kaymış.", ""),

    ("Akbank", "solo+kons", "Özkaynak", "Okuma",
     "Oranlar '11,98%' biçiminde; kalın yazılmamış ara başlıklar (OZK_ARA_BASLIK) başlık sayılır.", "Okuyucu._ozkaynak"),
    ("Akbank", "solo+kons", "Mevduat vade", "Okuma", "7 gün ihbarlı sütunu yok; '6 Ay - 1 Yıl' iki satıra bölünmüş.",
     "mevduat_vade_oku"),
    ("Akbank", "solo+kons", "Mevduata ödenen faiz", "Okuma",
     "Birikimli sütunu yok; önceki dönem tablosu yeni sayfada başlık tekrarlamaz.", "mevduat_faiz_oku"),
    ("Akbank", "solo", "Mevduat vade", "TBB", "Aralık 2025'te TBB sütunları bir kaydırmış ('1 aya kadar' → '7 gün ihbarlı').", ""),

    ("Yapı Kredi", "solo+kons", "Özkaynak", "Okuma",
     "İkinci 'Toplam Özkaynak' satırının üstündeki 'ÖZKAYNAK Cari Dönem Önceki Dönem' başlığı arındırılır.", "Okuyucu._ozkaynak"),
    ("Yapı Kredi", "solo+kons", "Mevduata ödenen faiz", "Okuma",
     "Ayrı önceki dönem tablosu yok; cari tabloda yalnız 'Önceki Dönem' toplam sütunu.", "mevduat_faiz_oku"),
    ("Yapı Kredi", "solo+kons", "Sigorta kapsamı", "Okuma",
     "Kalem adları 'Türk Lirası mevduat' / 'Döviz tevdiat hesapları' / 'Diğer mevduat hesapları'.", "sigorta_oku"),
    ("Yapı Kredi", "solo+kons", "Kredi tabloları", "Bilgi", "Yalnız cari dönem verilir.", "kredi_tablo_oku"),
    ("Yapı Kredi", "solo", "Sigorta kapsamı", "TBB", "Tasarruf satırlarına tüzel kişi mevduatı eklenmiş.", ""),
    ("Yapı Kredi", "solo+kons", "Gelir tablosu", "Okuma", "Çeyreklik gelir tablosunun ikinci kopyası atlanır.", ""),
]


def banka_kurallari_sayfasi(wb):
    ws = wb.create_sheet("Banka iş kuralları")
    ws.append(["Banka banka iş kuralları"])
    ws["A1"].font = Font(bold=True, size=12)
    ws.append(["Tür: Okuma = PDF düzenine göre okuma kuralı · Eşleme = PDF → TBB kalem eşlemesi · TBB = TBB verisinin farklı "
               "tuttuğu kalem (bizim okumamız değil) · Bilgi = yayımlanmayan / eksik tablo"])
    ws.append(["Banka", "Kapsam", "Tablo / dipnot", "Tür", "Kural", "Kodda"])
    for c in range(1, 7):
        ws.cell(3, c).fill = BASLIK
        ws.cell(3, c).font = Font(bold=True)
    for r in BANKA_KURALLARI:
        ws.append(list(r))
        ws.cell(ws.max_row, 5).alignment = Alignment(wrap_text=True, vertical="top")
        if r[3] == "TBB":
            ws.cell(ws.max_row, 4).fill = TURUNCU
    for i, w in enumerate([14, 11, 34, 9, 100, 32], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.auto_filter.ref = f"A3:F{ws.max_row}"
    ws.freeze_panes = "A4"


# ============================================================== konsol raporu

def sayi_yaz(x) -> str:
    if x is None:
        return "–"
    if isinstance(x, float) and not float(x).is_integer():
        return f"{x:,.3f}".replace(",", "_").replace(".", ",").replace("_", ".").rstrip("0").rstrip(",")
    return f"{x:,.0f}".replace(",", ".")


def rapor_konsol(rapor: "Rapor", kisa: str, tur: str, ok: "Okuyucu"):
    """Bir raporun tablo tablo ve kalem kalem sonuçlarını konsola yazar."""
    tur_ad = "Solo" if tur == "solo" else "Konsolide"
    bilgi = rapor.bilgi.get((kisa, tur), {})
    kay = [k for k in rapor.kayit if k["banka"] == kisa and k["tur"] == tur]
    cari_var = any(k["cnotu"] for k in kay)
    cizgi = "─" * 100
    print(f"\n   {cizgi}")
    print(f"   {kisa} – {tur_ad}   |   {ok.yol.name}")
    print(f"   birim: {bilgi.get('birim', '?')}, sayı biçimi: {'1.234,5 (Türkçe)' if ok.bicim == 'tr' else '1,234.5 (İngilizce)'}")
    print(f"   {cizgi}")
    bas = (f"   {'Tablo':<26}{'Okunan':>7}{'Eşleşen':>8}{'Aynı':>6}{'±1':>5}{'Küçük':>6}{'FARK':>6}{'Karşılıksız':>12}"
           f"{'Beklenen':>9}{'Yalnız PDF':>11}{'Öncekisi yok':>13}")
    if cari_var:
        bas += f"   │ {rapor.cari_et}: {'Aynı':>5}{'±1':>5}{'Küçük':>6}{'FARK':>6}"
    print(bas)
    sira = ([g for _, _, g in TABLOLAR] + ["Kur Riski"] + [sp["sayfa"] for sp in DIPNOT_TABLOLARI]
            + ["Göstergeler", "Diğer Dipnotlar"])
    okunan = {g: len(ok.tablolar.get(a, [])) for a, _, g in TABLOLAR}
    okunan["Kur Riski"] = 4 * sum(1 for r in ok.tablolar.get("KUR", []) if r["donem"] == "cari")
    okunan.update(bilgi.get("okunan", {}))
    for g in sira:
        kk = [k for k in kay if k["tablo"] == g]
        n = Counter(k["notu"].split(";")[0] for k in kk)
        nc = Counter(k["cnotu"] for k in kk if k["cnotu"])
        es = sum(1 for k in kk if k["kod"])
        sat = (f"   {g:<26}{okunan.get(g, 0):>7}{es:>8}{n['aynı']:>6}{n['yuvarlama (±1)']:>5}{n['küçük fark (≤3)']:>6}"
               f"{n['FARK']:>6}{n['TBB karşılığı yok']:>12}{n['beklenen (kural)']:>9}"
               f"{sum(v for k_, v in n.items() if k_.startswith('yalnız PDF')):>11}"
               f"{n['önceki dönem raporda yok']:>13}")
        if cari_var:
            sat += f"   │ {'':>{len(rapor.cari_et)}}  {nc['aynı']:>5}{nc['yuvarlama (±1)']:>5}{nc['küçük fark (≤3)']:>6}{nc['FARK']:>6}"
        print(sat)
    # sütun seçimi
    if bilgi.get("sutun"):
        print("\n   Kullanılan PDF sütunları (cari / önceki dönem):")
        for g, (ci, oi, et) in bilgi["sutun"].items():
            print(f"      {g:<20} cari: sütun {ci + 1}   önceki ({et}): sütun {oi + 1}")
    # bilanço denkliği
    for et, akt, pas, fark in bilgi.get("denk", []):
        durum = "denk" if fark is not None and abs(fark) <= 1 else "DENK DEĞİL"
        print(f"   Bilanço denkliği – {et:<13}: aktif {sayi_yaz(akt):>13}  pasif {sayi_yaz(pas):>13}  → {durum}")
    # öteki TBB dipnotları
    if bilgi.get("genel"):
        tp, liste = bilgi["genel"]
        print(f"\n   Öteki TBB dipnotları (genel eşleyici): {tp['dipnot']} dipnot; {tp['yilsonu']} tanesi yalnız yıl sonu "
              f"raporlarında. Ara dönemde de yayımlanan {tp['veri_var']} dipnotta TBB'de dolu {tp['dolu']} kalemin "
              f"{tp['bulunan']} tanesi PDF'te bulundu (%{100 * tp['bulunan'] / max(1, tp['dolu']):.0f}), "
              f"cari değeri okunan {tp['cari']}")
        print(f"   {'Dipnot':<62}{'TBB dolu':>9}{'Bulunan':>9}{'Cari':>6}{'FARK':>6}")
        for ad, dolu, bul, cari, fark, durum in liste:
            if dolu:
                ek = "   ← yalnız yıl sonu raporunda" if durum == "yılsonu" else ("   ← PDF'te yok" if not bul else "")
                print(f"   {ad[:61]:<62}{dolu:>9}{bul:>9}{cari:>6}{fark:>6}" + ek)
    # likidite ve kaldıraç oranları (PDF)
    uc = ok.tablolar.get("LKO_UC", [])
    if uc:
        print("\n   LKO son üç ay en düşük / en yüksek (TP+YP | YP):")
        for r in uc:
            print(f"      {('Cari' if r['donem'] == 'cari' else 'Önceki'):<7}{r['ad']:<10} "
                  f"{sayi_yaz(r['sayilar'][0]):>8} ({r['tarih'][0]})  |  {sayi_yaz(r['sayilar'][1]):>8} ({r['tarih'][1]})")
    oranlar = [k for k in kay if k["tablo"] in ("Likidite Karşılama Oranı", "NİFO", "Kaldıraç Oranı")
               and (re.search(r"ORANI", sade(k["kalem"])) and ("YP" in k["kalem"] or "Uygulanmış toplam" in k["kalem"]
                                                              or k["tablo"] == "Kaldıraç Oranı"))
               and k["kalem"].split()[0] in ("23", "34", "15")]
    if oranlar:
        print("\n   Likidite / NİFO / kaldıraç oranları (cari PDF / TBB | önceki PDF / TBB):")
        for k in oranlar:
            print(f"      {k['kalem'][:58]:<58} {sayi_yaz(k['c']):>8} / {sayi_yaz(k['tc']):>6} | "
                  f"{sayi_yaz(k['o']):>8} / {sayi_yaz(k['t']):>6}  {k['notu']}")
    # oranlar
    oran = [k for k in kay if str(k["kod"]).startswith("oran:")]
    if oran:
        print("\n   Özkaynak oranları (cari | önceki PDF / TBB):")
        for k in oran:
            print(f"      {k['kalem'][:58]:<58} {sayi_yaz(k['c']):>8} | {sayi_yaz(k['o']):>8} / {sayi_yaz(k['t']):>8}  {k['notu']}")
    # kalem listeleri

    def liste(baslik, secim, cari=False):
        kk = [k for k in kay if secim(k)]
        if not kk:
            return
        print(f"\n   {baslik} ({len(kk)}):")
        for k in kk:
            ad = f"[{k['tablo']}] {k['kalem']}"[:72]
            p_, t_ = (k["c"], k["tc"]) if cari else (k["o"], k["t"])
            fark = None if p_ is None or t_ is None else abs(t_) - abs(p_)
            print(f"      {ad:<72} PDF {sayi_yaz(p_):>12}  TBB {sayi_yaz(t_):>12}  fark {sayi_yaz(fark):>6}")
    liste("FARK – önceki dönem", lambda k: k["notu"].startswith("FARK"))
    liste("Küçük fark (≤3) – önceki dönem", lambda k: k["notu"].startswith("küçük"))
    if cari_var:
        liste(f"FARK – {rapor.cari_et}", lambda k: k["cnotu"] == "FARK", cari=True)
        liste(f"Küçük fark (≤3) – {rapor.cari_et}", lambda k: k["cnotu"].startswith("küçük"), cari=True)
    kk = [k for k in kay if k["notu"].startswith("TBB karşılığı yok")]
    if kk:
        print(f"\n   TBB karşılığı bulunamayan kalemler ({len(kk)}):")
        for k in kk:
            ad = f"[{k['tablo']}] {k['kalem']}"[:72]
            print(f"      {ad:<72} PDF önceki {sayi_yaz(k['o'])}")
    kk = [k for k in kay if k["notu"].startswith("beklenen")]
    if kk:
        print(f"\n   Beklenen (kurala göre aranmayan) kalemler ({len(kk)}):")
        for k in kk:
            print(f"      [{k['tablo']}] {k['kalem'][:80]}")


def genel_ozet(rapor: "Rapor"):
    print("\n" + "═" * 104)
    print(" GENEL ÖZET")
    print("═" * 104)
    cari_var = any(k["cnotu"] for k in rapor.kayit)
    bas = f" {'Banka':<14}{'Tür':<11}{'Eşleşen':>8}{'Aynı':>7}{'±1':>6}{'Küçük':>7}{'FARK':>6}{'Karşılıksız':>12}{'Beklenen':>9}"
    if cari_var:
        bas += f"  │ {rapor.cari_et}: {'Aynı':>6}{'±1':>5}{'Küçük':>6}{'FARK':>6}"
    print(bas)
    gruplar = []
    for k in rapor.kayit:
        if (k["banka"], k["tur"]) not in gruplar:
            gruplar.append((k["banka"], k["tur"]))
    top, topc = Counter(), Counter()
    for b, t in gruplar:
        kk = [k for k in rapor.kayit if k["banka"] == b and k["tur"] == t]
        n = Counter(k["notu"].split(";")[0] for k in kk)
        nc = Counter(k["cnotu"] for k in kk if k["cnotu"])
        es = sum(1 for k in kk if k["kod"])
        top.update(n); topc.update(nc); top["eşleşen"] += es
        sat = (f" {b:<14}{('Solo' if t == 'solo' else 'Konsolide'):<11}{es:>8}{n['aynı']:>7}{n['yuvarlama (±1)']:>6}"
               f"{n['küçük fark (≤3)']:>7}{n['FARK']:>6}{n['TBB karşılığı yok']:>12}{n['beklenen (kural)']:>9}")
        if cari_var:
            sat += f"  │ {'':>{len(rapor.cari_et)}}  {nc['aynı']:>6}{nc['yuvarlama (±1)']:>5}{nc['küçük fark (≤3)']:>6}{nc['FARK']:>6}"
        print(sat)
    print("─" * 104)
    sat = (f" {'TOPLAM':<25}{top['eşleşen']:>8}{top['aynı']:>7}{top['yuvarlama (±1)']:>6}{top['küçük fark (≤3)']:>7}"
           f"{top['FARK']:>6}{top['TBB karşılığı yok']:>12}{top['beklenen (kural)']:>9}")
    if cari_var:
        sat += (f"  │ {'':>{len(rapor.cari_et)}}  {topc['aynı']:>6}{topc['yuvarlama (±1)']:>5}"
                f"{topc['küçük fark (≤3)']:>6}{topc['FARK']:>6}")
    print(sat)
    if not cari_var:
        print(f" ({rapor.cari_et} TBB'de henüz yok; karşılaştırma önceki dönem sütunu üzerinden yapıldı.)")


# ============================================================== ana akış

def main():
    secilen, yil, ay = secimleri_al()
    kok = KLASOR / f"banka_karsilastirma_{yil}{ay:02d}"
    kok.mkdir(exist_ok=True)

    print(f"\n[1/4] KAP'tan raporlar indiriliyor ({len(secilen)} banka)")
    raporlar = []    # (banka, tür, pdf yolu)
    for b in secilen:
        print(f"   • {b[2]}", flush=True)
        try:
            pdfler = kap_paketi_indir(b, yil, ay, kok / dosya_adi(b[2]))
        except Exception as e:
            print(f"      ! atlandı: {e}")
            continue
        bulunan = {}
        for p in pdfler:
            t = pdf_turu(p)
            if t and t not in bulunan:
                bulunan[t] = p
        for t in ("solo", "kons"):
            if t in bulunan:
                print(f"      {('solo' if t == 'solo' else 'konsolide'):10s} → {bulunan[t].name}")
                raporlar.append((b, t, bulunan[t]))
        if not bulunan:
            print("      ! Türkçe rapor PDF'i bulunamadı")

    if not raporlar:
        sys.exit("\nOkunacak rapor yok.")

    print("\n[2/4] TBB verisi alınıyor")
    D = tbb_verisi()
    print(f"   TBB'deki son dönem: {donem_et(D['solo']['periods'][-1])}")

    print("\n[3/4] Raporlar okunuyor ve TBB ile eşleniyor")
    rapor = Rapor(yil, ay, D)
    for b, t, p in raporlar:
        try:
            ok = Okuyucu(p).oku()
            rapor.ekle(b, t, ok, p.name)
            rapor.ekle_kur(b, t, ok)
            rapor.ekle_dipnot(b, t, ok)
            rapor.ekle_gosterge(b, t, ok)
            print("      öteki TBB dipnotları aranıyor…", flush=True)
            rapor.ekle_genel(b, t, ok)
            rapor_konsol(rapor, b[2], t, ok)
        except Exception as e:
            print(f"   ! {b[2]} {t}: okunamadı ({e})")

    print("\n[4/4] Excel yazılıyor")
    cikti = kok / f"karsilastirma_{yil}{ay:02d}.xlsx"
    rapor.kaydet(cikti)

    genel_ozet(rapor)
    print(f"\nTamamlandı → {cikti}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit("\nDurduruldu.")
