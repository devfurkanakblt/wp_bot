"""
DEÜ CS&AI - WhatsApp grup ekleme botu (Selenium)

Akış:
  1. Aşama: Excel'deki her üyeyi WhatsApp Web üzerinden rehbere kaydeder ve gruba eklemeye çalışır.
            Gizlilik ayarı yüzünden eklenemeyenler ya da rehberde bulunamayanlar "DAVET_BEKLIYOR" olur;
            bir adım hata verirse kişi "HATA" olarak işaretlenir ve sonraki çalıştırmada yeniden denenir.
  2. Aşama: DAVET_BEKLIYOR durumundaki herkese tanıtım + davet linki içeren kişisel mesaj gönderilir.

Her adımın sonucu durum_raporu.csv dosyasına yazılır. Script yarıda kesilirse tekrar çalıştırdığında
tamamlanmış üyeleri atlayıp kaldığı yerden devam eder.

Kullanım:
  python wp_botu.py --onizle              # tarayıcı açmadan listeyi ve mesajı gösterir
  python wp_botu.py --test 5XXXXXXXXX     # yalnızca verilen numara ile deneme yapar
  python wp_botu.py --limit 5             # ilk 5 bekleyen üyeyle çalışır
  python wp_botu.py                       # herkes
  python wp_botu.py --kaydetme            # rehbere kaydetme adımını atlar
  python wp_botu.py --sadece-davet        # gruba ekleme yapmadan, bitmemiş herkese davet mesajı yollar
  python wp_botu.py --sifirla             # raporu yedekleyip sıfırlar, 1. sıradan yeniden başlar
  python wp_botu.py --izle                # botun penceresinde elle yaptığın hamleleri kaydeder
"""

import argparse
import base64
import csv
import json
import random
import re
import subprocess
import sys
import time
import urllib.parse
from datetime import datetime
from pathlib import Path

import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import (
    NoAlertPresentException,
    StaleElementReferenceException,
    WebDriverException,
)
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

# ------------------------------------------------------------------ AYARLAR
KLASOR = Path(__file__).resolve().parent
EXCEL = KLASOR / "Üye_Listesi_İsim_Telefon_E-posta.xlsx"
ULKE_KODU = "90"

# Grup adı ve davet linki git'e girmeyen ayarlar.json dosyasından okunur
# (örnek için ayarlar.ornek.json'a bak). Böylece davet linki herkese açık repoda görünmez.
AYAR_DOSYASI = KLASOR / "ayarlar.json"
if not AYAR_DOSYASI.exists():
    sys.exit("ayarlar.json bulunamadı. ayarlar.ornek.json dosyasını ayarlar.json olarak kopyalayıp doldur.")
_ayarlar = json.loads(AYAR_DOSYASI.read_text(encoding="utf-8"))
GRUP_ADI = _ayarlar["grup_adi"]
DAVET_LINKI = _ayarlar["davet_linki"]

PROFIL = KLASOR / "whatsapp_profil"          # QR kodu bir kere okutmak için kalıcı Chrome profili
RAPOR = KLASOR / "durum_raporu.csv"
HATA_KLASORU = KLASOR / "hata_ekranlari"

BEKLE = (8, 15)            # iki üye arasındaki rastgele bekleme (sn) - ban riskini azaltır
MOLA_HER = 20              # her 20 üyede bir
MOLA = (60, 120)           # bu kadar saniye mola ver

MESAJ = (
    "Merhaba {isim}! 👋\n\n"
    "Ben Furkan, Dokuz Eylül Üniversitesi Bilgisayar Bilimleri ve Yapay Zeka Topluluğu'nun "
    "(DEÜ CS&AI) başkan yardımcısıyım. Topluluğumuza katıldığın için çok teşekkür ederiz, "
    "aramıza hoş geldin! 🎉\n\n"
    "Etkinliklerimizden, duyurularımızdan ve projelerimizden ilk sen haberdar ol diye seni "
    "WhatsApp grubumuza eklemek istedim, ancak gizlilik ayarların nedeniyle doğrudan ekleyemedim. "
    "Aşağıdaki bağlantıya tıklayarak gruba hemen katılabilirsin:\n\n"
    "👉 {link}\n\n"
    "Aklına takılan bir şey olursa bana buradan yazabilirsin. Görüşmek üzere! 🚀\n\n"
    "— Furkan | DEÜ CS&AI"
)

# Durumlar
EKLENDI = "EKLENDI"
ZATEN_GRUPTA = "ZATEN_GRUPTA"
DAVET_BEKLIYOR = "DAVET_BEKLIYOR"
DAVET_GONDERILDI = "DAVET_GONDERILDI"
WA_YOK = "WHATSAPP_YOK"
GECERSIZ = "GECERSIZ_NUMARA"
HATA = "HATA"
BITMIS = {EKLENDI, ZATEN_GRUPTA, DAVET_GONDERILDI, WA_YOK}

# WhatsApp Web arayüzünde aranan uyarı metinleri
K_WA_YOK = ["WhatsApp'ta değil", "WhatsApp kullanmıyor", "not on WhatsApp", "isn't on WhatsApp"]
K_GECERSIZ_URL = ["geçersiz", "invalid"]
K_ZATEN = ["Zaten", "Already"]


# ------------------------------------------------------------------ YARDIMCILAR
def lit(s):
    """Python metnini güvenli bir XPath string literal'ine çevirir."""
    if "'" not in s:
        return f"'{s}'"
    if '"' not in s:
        return f'"{s}"'
    return "concat(" + ", \"'\", ".join(f"'{p}'" for p in s.split("'")) + ")"


def X_ICERIR(parcalar, kapsam="//"):
    """Metninde verilen parçalardan birini içeren en içteki elemanı bulan XPath."""
    kosul = " or ".join(
        f"(contains(., {lit(p)}) and not(./*[contains(., {lit(p)})]))" for p in parcalar
    )
    return f"{kapsam}*[{kosul}]"


DIALOG = "//div[@role='dialog']//"

# ------------------------------------------------------------------ SEÇİCİLER
# Kullanıcının botun penceresinde elle yaptığı hamlelerin kaydından (python wp_botu.py --izle) alındı.
# Tam XPath yerine WhatsApp'ın data-testid kimlikleri kullanılır: tam XPath'ler pencereye göre değişiyor
# (kullanıcının tarayıcısında /html/body/div[1]/..., botun penceresinde /html/body/div[2]/...).
S = {
    # Rehbere kaydetme
    "yeni_sohbet": "//button[@aria-label='Yeni sohbet' or @aria-label='New chat']",
    "yeni_kisi": "//*[@data-testid='new-chat-drawer-new-contact-cell']",
    "ad": "//*[@data-testid='contact-first-name-input']",
    "soyad": "//*[@data-testid='contact-last-name-input']",
    "telefon": "//input[@data-testid='phone-number-input']",
    "kisiyi_kaydet": "//*[@data-testid='save-contact-btn']",
    # Grubu açma
    "sohbet_arama": "//div[@id='side']//*[@data-testid='chat-list-search-container']//input",
    "sohbet_basligi": "//*[@data-testid='conversation-info-header']",
    # Gruba ekleme
    "ekle": ("//*[@data-testid='group-info-drawer-body']//button["
             ".//*[@data-icon='ic-person-add'] or .//*[name()='title' and text()='ic-person-add']"
             " or normalize-space(.)='Ekle']"),
    "pencere_arama": "//div[@role='dialog']//*[@data-testid='chat-list-search-container']//input",
    "ilk_sonuc": "(//div[@role='dialog']//*[@data-testid='list-item-1'])[1]",
    "secili_kisi": "//div[@role='dialog']//*[@data-testid='chat-controls'][contains(@aria-label, 'üyesini çıkar')]",
    "uye_ekle": "//div[@role='dialog']//button[normalize-space(.)='Üye ekle']",
    "onay_ekle": "//*[@data-testid='confirm-popup']//button[normalize-space(.)='Ekle']",
    # Gizlilik engeli -> WhatsApp'ın kendi davet akışı
    "gruba_davet_et": "//*[@data-testid='confirm-popup']//button[normalize-space(.)='Gruba davet et']",
    "davet_mesaji": "//*[@data-testid='invite-message-caption-input']",
    "davet_gonder": "//*[@data-testid='send-invitation-button']",
}


def normalize_tel(ham):
    """Telefonu 5XXXXXXXXX (10 hane) biçimine getirir, geçersizse None döner."""
    rakam = re.sub(r"\D", "", str(ham).split(".")[0])
    if rakam.startswith("0090"):
        rakam = rakam[4:]
    elif rakam.startswith("90") and len(rakam) == 12:
        rakam = rakam[2:]
    elif rakam.startswith("0") and len(rakam) == 11:
        rakam = rakam[1:]
    return rakam if re.fullmatch(r"5\d{9}", rakam) else None


def bekle(araligi):
    time.sleep(random.uniform(*araligi))


def log(sira, ad, mesaj):
    print(f"[{datetime.now():%H:%M:%S}] #{sira} {ad}: {mesaj}", flush=True)


# ------------------------------------------------------------------ RAPOR
ALANLAR = ["Sıra", "İsim", "Soyisim", "Telefon", "Durum", "Açıklama", "Rehber", "Zaman"]


def rapor_oku():
    if not RAPOR.exists():
        return {}
    with RAPOR.open(encoding="utf-8-sig", newline="") as f:
        return {r["Telefon"]: r for r in csv.DictReader(f)}


def rapor_yaz(rapor):
    with RAPOR.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=ALANLAR)
        w.writeheader()
        w.writerows(sorted(rapor.values(), key=lambda r: int(r["Sıra"])))


def guncelle(rapor, uye, durum, aciklama="", rehber=None):
    """rehber verilmezse kişinin önceki "Rehber" bilgisi (rehbere kaydedildi mi) korunur."""
    if rehber is None:
        rehber = rapor.get(uye["tel"], {}).get("Rehber", "")
    rapor[uye["tel"]] = {
        "Sıra": uye["sira"], "İsim": uye["isim"], "Soyisim": uye["soyisim"],
        "Telefon": uye["tel"], "Durum": durum, "Açıklama": aciklama, "Rehber": rehber,
        "Zaman": f"{datetime.now():%Y-%m-%d %H:%M:%S}",
    }
    rapor_yaz(rapor)


def uyeleri_oku():
    df = pd.read_excel(EXCEL, dtype={"Telefon": str})
    uyeler, gecersizler = [], []
    for _, r in df.iterrows():
        uye = {
            "sira": int(r["Sıra"]),
            "isim": str(r["İsim"]).strip(),
            "soyisim": str(r["Soyisim"]).strip(),
            "ham_tel": str(r["Telefon"]).strip(),
        }
        tel = normalize_tel(uye["ham_tel"])
        if tel:
            uye["tel"] = tel
            uyeler.append(uye)
        else:
            gecersizler.append(uye)
    return uyeler, gecersizler


# ------------------------------------------------------------------ WHATSAPP
class AdimHatasi(Exception):
    pass


class WhatsApp:
    def __init__(self):
        opts = webdriver.ChromeOptions()
        opts.add_argument(f"--user-data-dir={PROFIL}")
        opts.add_argument("--start-maximized")
        opts.add_argument("--lang=tr")
        opts.add_argument("--log-level=3")   # Chrome'un updater/GCM log satırlarını gizler
        opts.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
        self.d = webdriver.Chrome(options=opts)
        HATA_KLASORU.mkdir(exist_ok=True)

    # --- temel işlemler
    def bul(self, xpathler, sure=10, kok=None):
        """Verilen XPath'lerden herhangi birine uyan ilk görünür elemanı döndürür (yoksa None).
        kok verilirse arama o elemanın içinde yapılır (XPath'ler './/' ile başlamalı)."""
        if isinstance(xpathler, str):
            xpathler = [xpathler]
        bitis = time.time() + sure
        while True:
            for xp in xpathler:
                for el in (kok or self.d).find_elements(By.XPATH, xp):
                    try:
                        if el.is_displayed():
                            return el
                    except StaleElementReferenceException:
                        continue
            if time.time() >= bitis:
                return None
            time.sleep(0.3)

    def tikla(self, xpathler, sure=10, kok=None):
        el = self.bul(xpathler, sure, kok)
        if not el:
            return False
        try:
            el.click()
        except WebDriverException:
            self.d.execute_script("arguments[0].click();", el)
        time.sleep(0.8)
        return True

    def yaz(self, el, metin):
        el.click()
        el.send_keys(Keys.CONTROL, "a")
        el.send_keys(Keys.BACKSPACE)
        el.send_keys(metin)

    def esc(self, kez=1):
        for _ in range(kez):
            ActionChains(self.d).send_keys(Keys.ESCAPE).perform()
            time.sleep(0.5)

    def ekran_goruntusu(self, ad):
        try:
            self.d.save_screenshot(str(HATA_KLASORU / f"{datetime.now():%H%M%S}_{ad}.png"))
        except WebDriverException:
            pass

    def uyari_kapat(self):
        """Sayfadan ayrılırken çıkan 'Siteden ayrılılsın mı?' uyarısını kabul eder."""
        try:
            self.d.switch_to.alert.accept()
        except NoAlertPresentException:
            pass

    # --- oturum
    def giris(self):
        self.d.get("https://web.whatsapp.com")
        print("WhatsApp Web açılıyor... İlk çalıştırmada telefonundan QR kodu okut (5 dk süren var).")
        if not self.bul("//div[@id='pane-side']", sure=300):
            raise SystemExit("WhatsApp Web'e giriş yapılamadı.")
        print("Giriş başarılı, sohbetlerin yüklenmesi bekleniyor...")
        time.sleep(8)

    # --- klavye / pano
    def tuslar(self, *tuslar, ara=0.4):
        """Tuşlara sırayla basar (o an odakta olan elemana), aralarda kısa bekler."""
        zincir = ActionChains(self.d)
        for t in tuslar:
            zincir.send_keys(t).pause(ara)
        zincir.perform()

    def tumunu_sil(self):
        """Odaktaki alanda Ctrl+A ile her şeyi seçip siler."""
        ActionChains(self.d).key_down(Keys.CONTROL).send_keys("a").key_up(Keys.CONTROL)             .send_keys(Keys.BACKSPACE).perform()
        time.sleep(0.3)

    def tikla_el(self, el):
        try:
            el.click()
        except WebDriverException:
            self.d.execute_script("arguments[0].click();", el)
        time.sleep(0.8)

    def adim(self, anahtar, aciklama, sure=10):
        """S[anahtar] elemanına tıklar; bulunamazsa hangi adımda takıldığını söyleyen hata verir."""
        el = self.bul(S[anahtar], sure)
        if not el:
            raise AdimHatasi(f"{aciklama} bulunamadı ({anahtar}: {S[anahtar]})")
        self.tikla_el(el)
        return el

    def yapistir(self, metin):
        """Metni odaktaki alana Ctrl+V ile yapıştırır (emojiler klavyeyle yazılamadığı için).
        Panoya PowerShell ile yazılır; Türkçe karakter ve emoji bozulmasın diye metin base64 ile aktarılır."""
        b64 = base64.b64encode(metin.encode("utf-8")).decode("ascii")
        subprocess.run(
            ["powershell", "-NoProfile", "-Command",
             f"Set-Clipboard -Value ([Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('{b64}')))"],
            check=True, capture_output=True,
        )
        ActionChains(self.d).key_down(Keys.CONTROL).send_keys("v").key_up(Keys.CONTROL).perform()
        time.sleep(1)

    def dok(self, sira, ad):
        """Hata ayıklama için yalnızca açık pencerelerin, grup bilgisinin ve kişi formunun HTML'ini kaydeder;
        sohbet listesi ve mesajlar dahil edilmez."""
        parcalar = []
        for baslik, xp in (
            ("pencereler", "//div[@role='dialog']"),
            ("grup_bilgisi", "//*[@data-testid='group-info-drawer-body']"),
            ("kisi_formu", f"{S['ad']}/ancestor::*[.{S['kisiyi_kaydet']}][1]"),
        ):
            for el in self.d.find_elements(By.XPATH, xp):
                try:
                    parcalar.append(f"<!-- {baslik} -->\n{el.get_attribute('outerHTML')}")
                except WebDriverException:
                    pass
        if parcalar:
            yol = HATA_KLASORU / f"{datetime.now():%H%M%S}_{sira}_{ad}.html"
            yol.write_text("\n\n".join(parcalar), encoding="utf-8")

    # --- rehbere kaydetme (kayıttaki hamleler)
    def kisi_kaydet(self, uye):
        """Yeni sohbet > Yeni kişi > Ad, Tab, Soyadı > Telefon > yeşil tik. WA_YOK ya da None döner."""
        self.adim("yeni_sohbet", "'Yeni sohbet' butonu")
        self.adim("yeni_kisi", "'Yeni kişi' seçeneği")

        ad_el = self.adim("ad", "Ad alanı")
        self.tumunu_sil()
        ActionChains(self.d).send_keys(uye["isim"]).perform()
        self.tuslar(Keys.TAB)
        ActionChains(self.d).send_keys(uye["soyisim"]).perform()
        time.sleep(0.5)
        soyad_el = self.bul(S["soyad"], 3)
        if uye["isim"] not in ad_el.text or not soyad_el or uye["soyisim"] not in soyad_el.text:
            self.esc(2)
            raise AdimHatasi("Ad / Soyadı alanları doldurulamadı")

        tel_el = self.adim("telefon", "Telefon alanı")
        self.tumunu_sil()
        tel_el.send_keys(uye["tel"])                 # 10 hane, ülke kodu (+90) formda seçili
        time.sleep(3)                                # WhatsApp numarayı kontrol ediyor
        if self.bul(X_ICERIR(K_WA_YOK), 1):
            self.esc(2)
            return WA_YOK

        self.adim("kisiyi_kaydet", "Yeşil tik (Kişiyi kaydet)")
        time.sleep(3)
        if self.bul(S["kisiyi_kaydet"], 1):
            self.esc(2)
            raise AdimHatasi("Kayıt formu kapanmadı (kişi zaten kayıtlı olabilir)")
        return None

    # --- grup (kayıttaki hamleler)
    def grubu_ac(self):
        """Arama kutusuna grup adını yazar, 3 kez Tab + 1 kez Enter ile grubu açar."""
        kutu = self.adim("sohbet_arama", "Sohbet arama kutusu", 15)
        self.tumunu_sil()
        kutu.send_keys(GRUP_ADI)
        time.sleep(2)
        self.tuslar(Keys.TAB, Keys.TAB, Keys.TAB, Keys.ENTER)
        time.sleep(2)
        if not self.bul(f"{S['sohbet_basligi']}[contains(., {lit(GRUP_ADI)})]", 10):
            raise AdimHatasi("Grup sohbeti açılmadı (Tab x3 + Enter sonrası başlıkta grup adı yok)")

    def grup_bilgisini_ac(self):
        if self.bul(S["ekle"], 1):
            return                                   # grup bilgisi zaten açık
        self.grubu_ac()
        self.adim("sohbet_basligi", "Grup başlığı")
        if not self.bul(S["ekle"], 10):
            raise AdimHatasi(f"Grup bilgisinde 'Ekle' butonu bulunamadı (ekle: {S['ekle']})")

    def gruba_ekle(self, uye):
        """EKLENDI, ZATEN_GRUPTA, DAVET_GONDERILDI ya da DAVET_BEKLIYOR (kişi bulunamadı) döner."""
        self.grup_bilgisini_ac()
        self.adim("ekle", "Grup bilgisindeki 'Ekle' butonu")

        kutu = self.bul(S["pencere_arama"], 10)
        if not kutu:
            raise AdimHatasi("Kişi ekleme penceresindeki arama kutusu bulunamadı")
        kutu.send_keys(f"{uye['isim']} {uye['soyisim']}")
        time.sleep(2.5)

        ilk = self.bul(S["ilk_sonuc"], 3)
        if not ilk:
            self.esc()
            return DAVET_BEKLIYOR                    # rehberde bulunamadı -> mesajla davet edilecek
        if self.bul(DIALOG + X_ICERIR(K_ZATEN, ""), 0.5):
            self.esc()
            return ZATEN_GRUPTA
        if uye["isim"].split()[0].lower() not in ilk.text.lower():
            self.esc()
            raise AdimHatasi(f"İlk sonuç bu kişi değil: '{ilk.text.strip()[:40]}'")

        # İlk kişiyi seç: Tab, Tab, Enter
        self.tuslar(Keys.TAB, Keys.TAB, Keys.ENTER)
        if not self.bul(S["secili_kisi"], 3):
            self.esc()
            raise AdimHatasi("Kişi seçilemedi (Tab x2 + Enter sonrası seçili kişi yok)")

        self.adim("uye_ekle", "'Üye ekle' butonu", 5)
        self.adim("onay_ekle", "Onay penceresindeki 'Ekle' butonu", 5)

        # Gizlilik engeli varsa "X eklenemedi ... davet edebilirsiniz" penceresi açılır
        if not self.bul(S["gruba_davet_et"], 8):
            return EKLENDI
        self.adim("gruba_davet_et", "'Gruba davet et' butonu", 3)
        self.davet_mesaji_gonder(uye)
        return DAVET_GONDERILDI

    def davet_mesaji_gonder(self, uye):
        """WhatsApp'ın davet penceresinde mesajı seçip (Ctrl+A) bizim mesajla değiştirir (Ctrl+V) ve gönderir."""
        self.adim("davet_mesaji", "Davet mesajı kutusu")
        ActionChains(self.d).key_down(Keys.CONTROL).send_keys("a").key_up(Keys.CONTROL).perform()
        self.yapistir(mesaj_metni(uye))
        kutu = self.bul(S["davet_mesaji"], 2)
        if not kutu or uye["isim"] not in kutu.text or "Furkan" not in kutu.text:
            raise AdimHatasi("Davet mesajı kutuya yapıştırılamadı")
        self.adim("davet_gonder", "Davet gönder butonu", 5)
        bitis = time.time() + 15
        while time.time() < bitis and self.bul(S["davet_gonder"], 0):
            time.sleep(0.5)
        if self.bul(S["davet_gonder"], 0):
            raise AdimHatasi("Davet gönderildiği doğrulanamadı (pencere kapanmadı)")

    # --- mesaj
    def mesaj_gonder(self, tel, metin):
        """DAVET_GONDERILDI veya WA_YOK döner."""
        url = (f"https://web.whatsapp.com/send?phone={ULKE_KODU}{tel}"
               f"&text={urllib.parse.quote(metin)}")
        self.d.get(url)
        self.uyari_kapat()

        bitis = time.time() + 60
        kutu = None
        while time.time() < bitis:
            kutu = self.bul("//footer//div[@contenteditable='true']", 1)
            if kutu and kutu.text.strip():
                break
            if self.bul(DIALOG + X_ICERIR(K_GECERSIZ_URL, ""), 0):
                self.esc()
                return WA_YOK
        else:
            raise AdimHatasi("Sohbet açılmadı ya da mesaj kutusu dolmadı")

        time.sleep(1.5)
        if not self.tikla([
            "//footer//*[@data-icon='send']",
            "//footer//*[@data-icon='wds-ic-send-filled']",
            "//footer//button[@aria-label='Gönder' or @aria-label='Send']",
        ], 5):
            kutu.send_keys(Keys.ENTER)

        # Mesaj kutusu boşalsın ve "saat" (gönderiliyor) ikonu kaybolsun
        bitis = time.time() + 30
        while time.time() < bitis:
            kutu = self.bul("//footer//div[@contenteditable='true']", 1)
            bekleyen = self.d.find_elements(By.XPATH, "//div[@id='main']//*[@data-icon='msg-time']")
            if kutu and not kutu.text.strip() and not bekleyen:
                time.sleep(1)
                return DAVET_GONDERILDI
            time.sleep(0.5)
        raise AdimHatasi("Mesaj gönderildiği doğrulanamadı")

    def kapat(self):
        try:
            self.d.quit()
        except WebDriverException:
            pass


# ------------------------------------------------------------------ ANA AKIŞ
def mesaj_metni(uye):
    return MESAJ.format(isim=uye["isim"], link=DAVET_LINKI)


def mola(sayac):
    if sayac and sayac % MOLA_HER == 0:
        s = random.uniform(*MOLA)
        print(f"--- {sayac} kişi işlendi, {s:.0f} sn mola ---", flush=True)
        time.sleep(s)
    else:
        bekle(BEKLE)


def asama1_ekle(wa, uyeler, rapor, kaydet):
    print(f"\n=== 1. AŞAMA: Rehbere kaydetme + gruba ekleme ({len(uyeler)} kişi) ===")
    for i, uye in enumerate(uyeler, 1):
        ad = f"{uye['isim']} {uye['soyisim']}"
        try:
            if kaydet and rapor.get(uye["tel"], {}).get("Rehber") == "EVET":
                log(uye["sira"], ad, "zaten rehbere kaydedilmiş, kayıt adımı atlanıyor")
            elif kaydet:
                try:
                    if wa.kisi_kaydet(uye) == WA_YOK:
                        log(uye["sira"], ad, "WhatsApp kullanmıyor, atlanıyor")
                        guncelle(rapor, uye, WA_YOK)
                        continue
                    log(uye["sira"], ad, "rehbere kaydedildi")
                    guncelle(rapor, uye, HATA, "Rehbere kaydedildi, gruba eklenmedi", rehber="EVET")
                except AdimHatasi as e:
                    log(uye["sira"], ad, f"kayıt uyarısı: {e} (gruba eklemeye devam)")
                    wa.ekran_goruntusu(f"{uye['sira']}_kayit")
                    wa.dok(uye["sira"], "kayit")

            sonuc = wa.gruba_ekle(uye)
            if sonuc == DAVET_BEKLIYOR:
                # Ekleme penceresinde kişi bulunamadı -> davet linkini hemen mesajla gönder
                log(uye["sira"], ad, "ekleme penceresinde bulunamadı, davet mesajı şimdi gönderiliyor")
                davet_gonder(wa, uye, rapor)
            else:
                aciklama = {
                    EKLENDI: "Gruba eklendi",
                    ZATEN_GRUPTA: "Zaten grupta",
                    DAVET_GONDERILDI: "Eklenemedi (gizlilik), 'Gruba davet et' ile davet gönderildi",
                }[sonuc]
                log(uye["sira"], ad, aciklama)
                guncelle(rapor, uye, sonuc, aciklama)
        except (AdimHatasi, WebDriverException) as e:
            hata = str(e).splitlines()[0][:150]
            log(uye["sira"], ad, f"HATA: {hata} -> sonraki çalıştırmada tekrar denenecek")
            wa.ekran_goruntusu(f"{uye['sira']}_grup")
            wa.dok(uye["sira"], "grup")
            guncelle(rapor, uye, HATA, f"Ekleme hatası: {hata}")
            wa.esc(3)
        mola(i)


def davet_gonder(wa, uye, rapor):
    """Kişiye tanıtım + davet linki mesajını gönderir ve sonucu rapora yazar."""
    ad = f"{uye['isim']} {uye['soyisim']}"
    try:
        sonuc = wa.mesaj_gonder(uye["tel"], mesaj_metni(uye))
        aciklama = "Davet mesajı gönderildi" if sonuc == DAVET_GONDERILDI else "Numara WhatsApp'ta yok"
        log(uye["sira"], ad, aciklama)
        guncelle(rapor, uye, sonuc, aciklama)
    except (AdimHatasi, WebDriverException) as e:
        hata = str(e).splitlines()[0][:150]
        log(uye["sira"], ad, f"davet mesajı HATA: {hata} (sonraki çalıştırmada tekrar denenecek)")
        wa.ekran_goruntusu(f"{uye['sira']}_mesaj")
        guncelle(rapor, uye, DAVET_BEKLIYOR, f"Mesaj hatası: {hata}")


def asama2_davet(wa, uyeler, rapor):
    """Yalnızca önceki çalıştırmalarda daveti gönderilememiş kişiler için (yeniden deneme)."""
    print(f"\n=== Gönderilemeyen davetler tekrar deneniyor ({len(uyeler)} kişi) ===")
    for i, uye in enumerate(uyeler, 1):
        davet_gonder(wa, uye, rapor)
        mola(i)


# ------------------------------------------------------------------ İZLEME MODU
IZLEME_DOSYASI = KLASOR / "izleme_kaydi.txt"

# Sayfaya eklenen dinleyici: her tıklamayı / önemli tuşu / yazıyı tam XPath ve tanımlayıcılarıyla kaydeder
IZLEME_JS = r"""
if (!window.__izleme) {
  window.__izleme = [];
  const tamXPath = el => {
    const parcalar = [];
    for (; el && el.nodeType === 1; el = el.parentElement) {
      const ad = el.namespaceURI === 'http://www.w3.org/2000/svg'
        ? `*[name()='${el.localName}']` : el.localName;
      const kardes = el.parentElement
        ? [...el.parentElement.children].filter(c => c.localName === el.localName) : [el];
      parcalar.unshift(kardes.length > 1 ? `${ad}[${kardes.indexOf(el) + 1}]` : ad);
    }
    return '/' + parcalar.join('/');
  };
  const tanimla = el => {
    if (!el || el.nodeType !== 1) return null;
    const ikon = el.closest('[data-icon]') || el.querySelector('[data-icon]');
    const svgBaslik = el.querySelector('svg title') || el.closest('svg')?.querySelector('title');
    const buton = el.closest('button, [role="button"], [role="listitem"], [role="row"], [role="option"]');
    return {
      xpath: tamXPath(el), etiket: el.localName,
      role: el.getAttribute('role'), testid: el.closest('[data-testid]')?.getAttribute('data-testid'),
      aria: el.closest('[aria-label]')?.getAttribute('aria-label'),
      title: el.closest('[title]')?.getAttribute('title'),
      ikon: ikon?.getAttribute('data-icon') || svgBaslik?.textContent,
      yazi: (el.innerText || el.textContent || '').trim().replace(/\s+/g, ' ').slice(0, 80),
      tiklanabilir: buton ? {
        xpath: tamXPath(buton), etiket: buton.localName, role: buton.getAttribute('role'),
        testid: buton.getAttribute('data-testid'), aria: buton.getAttribute('aria-label'),
        yazi: (buton.innerText || '').trim().replace(/\s+/g, ' ').slice(0, 80),
      } : null,
      bolge: el.closest('#side') ? 'sol panel (#side)' : el.closest('#main') ? 'sohbet (#main)'
           : el.closest('section') ? 'bilgi paneli (section)' : el.closest('[role="dialog"]') ? 'dialog'
           : 'diğer / açılır pencere',
    };
  };
  const ekle = (tur, el, ek) => window.__izleme.push(
    Object.assign({ zaman: new Date().toLocaleTimeString('tr-TR'), tur }, tanimla(el), ek || {}));
  document.addEventListener('click', e => ekle('TIKLAMA', e.target), true);
  document.addEventListener('keydown', e => {
    if (['Enter', 'Tab', 'Escape', 'ArrowDown', 'ArrowUp'].includes(e.key) || e.ctrlKey)
      ekle('TUŞ', e.target, { tus: (e.ctrlKey ? 'Ctrl+' : '') + (e.shiftKey ? 'Shift+' : '') + e.key });
  }, true);
  document.addEventListener('input', e => {
    const t = e.target;
    const deger = t.value !== undefined ? t.value : (t.innerText || '');
    ekle('YAZI', t, { deger: deger.trim().slice(0, 80) });
  }, true);
}
"""


def izle():
    """Botun Chrome penceresini açar; kullanıcının yaptığı her hamleyi izleme_kaydi.txt dosyasına yazar."""
    wa = WhatsApp()
    satirlar = []
    son_yazi = None
    try:
        wa.giris()
        print("\n=== İZLEME MODU ===")
        print("Bu Chrome penceresinde rehbere kişi ekleme ve gruba üye ekleme işlemlerini elle yap.")
        print("Her tıklaman ve tuşun aşağıda görünecek. Bitince bu terminalde Ctrl+C'ye bas.\n")
        while True:
            try:
                wa.d.execute_script(IZLEME_JS)    # sayfa yenilenirse dinleyiciyi tekrar kur
                olaylar = wa.d.execute_script("const o = window.__izleme || []; window.__izleme = []; return o;")
            except WebDriverException:
                print("Chrome penceresi kapandı.")
                break
            for o in olaylar:
                # Arka arkaya aynı alana yazılan harfleri tek satırda birleştir
                if o["tur"] == "YAZI" and son_yazi and son_yazi["xpath"] == o["xpath"]:
                    son_yazi["deger"] = o["deger"]
                    satirlar[-1] = izleme_satiri(son_yazi)
                    continue
                son_yazi = o if o["tur"] == "YAZI" else None
                satirlar.append(izleme_satiri(o))
                print(satirlar[-1], flush=True)
            IZLEME_DOSYASI.write_text("\n".join(satirlar), encoding="utf-8")
            time.sleep(0.5)
    except KeyboardInterrupt:
        pass
    finally:
        IZLEME_DOSYASI.write_text("\n".join(satirlar), encoding="utf-8")
        print(f"\n{len(satirlar)} hamle kaydedildi: {IZLEME_DOSYASI}")
        wa.kapat()


def izleme_satiri(o):
    s = f"[{o['zaman']}] {o['tur']}"
    if o.get("tus"):
        s += f" {o['tus']}"
    if o.get("deger") is not None:
        s += f" = '{o['deger']}'"
    s += f"\n    bölge: {o.get('bolge')} | eleman: <{o.get('etiket')}>"
    for k in ("testid", "aria", "title", "ikon", "role"):
        if o.get(k):
            s += f" {k}={o[k]!r}"
    if o.get("yazi") and o["tur"] != "YAZI":
        s += f" yazı={o['yazi']!r}"
    s += f"\n    xpath: {o.get('xpath')}"
    t = o.get("tiklanabilir")
    if t:
        s += (f"\n    tıklanabilir üst: <{t['etiket']}> role={t['role']!r} testid={t['testid']!r} "
              f"aria={t['aria']!r} yazı={t['yazi']!r}\n    tıklanabilir xpath: {t['xpath']}")
    return s


def ozet(rapor, gecersizler):
    print("\n=== ÖZET ===")
    sayim = {}
    for r in rapor.values():
        sayim[r["Durum"]] = sayim.get(r["Durum"], 0) + 1
    for durum, adet in sorted(sayim.items()):
        print(f"  {durum:<18} {adet}")
    if gecersizler:
        print(f"  {GECERSIZ:<18} {len(gecersizler)}  (Excel'de düzeltilmesi gerekiyor):")
        for u in gecersizler:
            print(f"     #{u['sira']} {u['isim']} {u['soyisim']} -> '{u['ham_tel']}'")
    print(f"\nAyrıntılı rapor: {RAPOR}")


def main():
    p = argparse.ArgumentParser(description="DEÜ CS&AI WhatsApp grup ekleme botu")
    p.add_argument("--onizle", action="store_true", help="tarayıcı açmadan liste ve mesajı göster")
    p.add_argument("--test", metavar="TEL", help="sadece bu numarayla deneme yap (5XXXXXXXXX)")
    p.add_argument("--limit", type=int, help="en fazla bu kadar bekleyen üyeyi işle")
    p.add_argument("--kaydetme", action="store_true", help="rehbere kaydetme adımını atla")
    p.add_argument("--sadece-davet", action="store_true", help="gruba eklemeden herkese davet mesajı gönder")
    p.add_argument("--izle", action="store_true",
                   help="botun Chrome penceresini aç, elle yaptığın hamleleri izleme_kaydi.txt'ye kaydet")
    p.add_argument("--sifirla", action="store_true",
                   help="ilerleme raporunu yedekleyip sıfırla, herkese 1. sıradan yeniden başla")
    args = p.parse_args()

    if args.izle:
        izle()
        return

    global RAPOR
    uyeler, gecersizler = uyeleri_oku()
    if args.test:
        RAPOR = KLASOR / "test_raporu.csv"   # deneme sonuçları asıl raporu kirletmesin
        tel = normalize_tel(args.test)
        if not tel:
            sys.exit("Test numarası geçersiz. Örnek: 5321234567")
        uyeler = [{"sira": 0, "isim": "Test", "soyisim": "Kişi", "ham_tel": args.test, "tel": tel}]
        gecersizler = []

    if args.sifirla and RAPOR.exists():
        yedek = RAPOR.with_name(f"{RAPOR.stem}_yedek_{datetime.now():%Y%m%d_%H%M%S}.csv")
        RAPOR.rename(yedek)
        print(f"Rapor sıfırlandı, eskisi yedeklendi: {yedek.name}")

    rapor = rapor_oku()
    bekleyen =[u for u in uyeler if rapor.get(u["tel"], {}).get("Durum") not in BITMIS]
    if args.limit:
        bekleyen = bekleyen[: args.limit]

    print(f"Toplam geçerli üye: {len(uyeler)} | işlenecek: {len(bekleyen)} | geçersiz numara: {len(gecersizler)}")
    if args.onizle:
        print("\n--- Örnek mesaj ---\n" + mesaj_metni(bekleyen[0] if bekleyen else uyeler[0]))
        ozet(rapor, gecersizler)
        return
    if not bekleyen:
        ozet(rapor, gecersizler)
        return

    wa = WhatsApp()
    try:
        wa.giris()
        if args.sadece_davet:
            for u in bekleyen:
                if rapor.get(u["tel"], {}).get("Durum") != DAVET_BEKLIYOR:
                    guncelle(rapor, u, DAVET_BEKLIYOR, "Sadece davet modu")
        else:
            ilk_asama = [u for u in bekleyen if rapor.get(u["tel"], {}).get("Durum") != DAVET_BEKLIYOR]
            asama1_ekle(wa, ilk_asama, rapor, kaydet=not args.kaydetme)

        davet_listesi = [u for u in bekleyen if rapor.get(u["tel"], {}).get("Durum") == DAVET_BEKLIYOR]
        if davet_listesi:
            asama2_davet(wa, davet_listesi, rapor)
    except KeyboardInterrupt:
        print("\nKullanıcı tarafından durduruldu. Tekrar çalıştırınca kaldığı yerden devam eder.")
    finally:
        ozet(rapor, gecersizler)
        wa.kapat()


if __name__ == "__main__":
    main()
