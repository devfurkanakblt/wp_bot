"""
DEÜ CS&AI - WhatsApp grup ekleme botu (Selenium)

Akış:
  1. Aşama: Excel'deki her üyeyi WhatsApp Web üzerinden rehbere kaydeder ve gruba eklemeye çalışır.
            Gizlilik ayarı yüzünden (ya da başka bir sebeple) eklenemeyenler "DAVET_BEKLIYOR" olarak işaretlenir.
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
"""

import argparse
import csv
import json
import random
import re
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

# WhatsApp Web arayüz metinleri (Türkçe + İngilizce arayüz için)
T_YENI_SOHBET = ["Yeni sohbet", "New chat"]
T_YENI_KISI = ["Yeni kişi", "New contact"]
T_KAYDET = ["Kaydet", "Save"]
# "Yeni kişi" formundaki yeşil tik (kaydet) butonunun tam XPath'i
TIK_XPATH = "/html/body/div[1]/div/div/div/div/div[3]/div/div[2]/div[1]/div/span/div/span/div/div/div[2]/span/div/span"
T_UYE_EKLE = ["Kişi ekle", "Add member", "Add members", "Add participant"]
T_EKLE = ["Ekle", "Kişi ekle", "Add", "Add member", "Add participant"]
T_IPTAL = ["İptal", "Vazgeç", "Kapat", "Tamam", "Cancel", "Close", "OK"]
K_WA_YOK = ["WhatsApp'ta değil", "WhatsApp kullanmıyor", "not on WhatsApp", "isn't on WhatsApp"]
K_GECERSIZ_URL = ["geçersiz", "invalid"]
K_ZATEN = ["Zaten", "Already"]
K_DAVET = ["davet", "Davet", "invite", "Invite"]


# ------------------------------------------------------------------ YARDIMCILAR
def lit(s):
    """Python metnini güvenli bir XPath string literal'ine çevirir."""
    if "'" not in s:
        return f"'{s}'"
    if '"' not in s:
        return f'"{s}"'
    return "concat(" + ", \"'\", ".join(f"'{p}'" for p in s.split("'")) + ")"


def X_TAM(metinler, kapsam="//"):
    """Metni / aria-label'ı / title'ı tam olarak eşleşen en içteki elemanı bulan XPath."""
    kosul = " or ".join(
        f"(normalize-space(.)={lit(m)} and not(.//*[normalize-space(.)={lit(m)}]))"
        f" or @aria-label={lit(m)} or @title={lit(m)}"
        for m in metinler
    )
    return f"{kapsam}*[{kosul}]"


def X_ICERIR(parcalar, kapsam="//"):
    """Metninde verilen parçalardan birini içeren en içteki elemanı bulan XPath."""
    kosul = " or ".join(
        f"(contains(., {lit(p)}) and not(./*[contains(., {lit(p)})]))" for p in parcalar
    )
    return f"{kapsam}*[{kosul}]"


DIALOG = "//div[@role='dialog']//"


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
ALANLAR = ["Sıra", "İsim", "Soyisim", "Telefon", "Durum", "Açıklama", "Zaman"]


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


def guncelle(rapor, uye, durum, aciklama=""):
    rapor[uye["tel"]] = {
        "Sıra": uye["sira"], "İsim": uye["isim"], "Soyisim": uye["soyisim"],
        "Telefon": uye["tel"], "Durum": durum, "Açıklama": aciklama,
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

    def ara(self, metin):
        kutu = self.bul([
            "//div[@id='side']//div[@contenteditable='true']",
            "//div[@id='side']//input[@type='text']",
            "//div[@id='side']//*[@role='textbox']",
        ], 15)
        if not kutu:
            raise AdimHatasi("Sohbet arama kutusu bulunamadı")
        self.yaz(kutu, metin)
        time.sleep(2)

    # --- rehbere kaydetme
    def kisi_kaydet(self, uye):
        """Yeni sohbet > Yeni kişi formuyla kişiyi rehbere kaydeder. WA_YOK ya da None döner."""
        if not self.tikla([
            "//*[@data-icon='new-chat-outline']",
            "//*[@data-icon='new-chat']",
            X_TAM(T_YENI_SOHBET),
        ]):
            raise AdimHatasi("'Yeni sohbet' butonu bulunamadı")
        if not self.tikla(X_TAM(T_YENI_KISI)):
            self.esc()
            raise AdimHatasi("'Yeni kişi' seçeneği bulunamadı")

        alanlar = self.form_alanlari(sure=10)
        if not alanlar:
            self.esc(2)
            raise AdimHatasi("Yeni kişi formu açılmadı")
        time.sleep(1)                                   # formun tüm alanları çizilsin
        alanlar = self.form_alanlari(sure=3) or alanlar
        etiketler = [self.etiket(el).lower() for el in alanlar]
        turler = [self.alan_turu(el, e) for el, e in zip(alanlar, etiketler)]

        def tur(ad):
            return next((el for el, t in zip(alanlar, turler) if t == ad), None)

        # Telefon: etiketinden ya da tipinden; bulunamazsa kullanıcı adı olmayan son alan
        tel_el = tur("telefon") or next(
            (el for el, t in reversed(list(zip(alanlar, turler))) if t != "kullanici"), None)
        # Ad/Soyad: etiketinden; bulunamazsa formun ilk iki alanı (kullanıcı adı ve telefon hariç)
        isim_adaylari = [el for el, t in zip(alanlar, turler)
                         if t not in ("kullanici", "telefon") and el != tel_el]
        ad_el = tur("ad") or (isim_adaylari[0] if isim_adaylari else None)
        soyad_el = tur("soyad") or next((el for el in isim_adaylari if el != ad_el), None)

        panel = self.form_paneli(tel_el or alanlar[0])
        if not tel_el:
            self.form_dok(panel, uye)
            self.esc(2)
            raise AdimHatasi(f"Telefon alanı bulunamadı (bulunan alanlar: {etiketler})")
        if not ad_el:
            self.form_dok(panel, uye)
            self.esc(2)
            raise AdimHatasi(f"Ad alanı bulunamadı (bulunan alanlar: {etiketler})")

        if soyad_el is None:
            self.alana_yaz(ad_el, f"{uye['isim']} {uye['soyisim']}", "Ad", panel, uye)
        else:
            self.alana_yaz(ad_el, uye["isim"], "Ad", panel, uye)
            self.alana_yaz(soyad_el, uye["soyisim"], "Soyad", panel, uye)
        self.alana_yaz(tel_el, uye["tel"], "Telefon", panel, uye)   # ülke kodu +90 seçili olmalı
        time.sleep(3)                                               # WhatsApp numarayı kontrol ediyor

        if self.bul(X_ICERIR(K_WA_YOK), 1):
            self.esc(2)
            return WA_YOK

        # Kaydetme butonu formun altındaki yeşil tik (✓). Önce bilinen tam XPath denenir;
        # WhatsApp arayüzü değişip o yol bozulursa diğer yöntemler yedek olarak devreye girer.
        tik = self.bul(TIK_XPATH, 5) or self.bul([
            ".//*[contains(@data-icon, 'checkmark')]",
            ".//*[@role='button' or self::button][@aria-label='Onayla' or @aria-label='Kaydet'"
            " or @aria-label='Bitti' or @aria-label='Confirm' or @aria-label='Save' or @aria-label='Done']",
            X_TAM(T_KAYDET, ".//"),
            ".//button[@type='submit']",
        ], 2, kok=panel) or self.ustteki_tik()
        if not tik:
            self.form_dok(panel, uye)
            self.esc(2)
            raise AdimHatasi("Kaydetme (yeşil tik) butonu bulunamadı")
        # İkon <span> ise tıklanabilir üst elemanını (button / role=button) kullan
        buton = self.d.execute_script(
            "return arguments[0].closest('button, [role=\"button\"]') || arguments[0];", tik)
        if buton.get_attribute("aria-disabled") == "true" or buton.get_attribute("disabled"):
            self.form_dok(panel, uye)
            self.esc(2)
            raise AdimHatasi("Yeşil tik pasif (form eksik doldurulmuş olabilir)")
        try:
            buton.click()
        except WebDriverException:
            self.d.execute_script("arguments[0].click();", buton)
        time.sleep(3)

        # Form hâlâ açıksa (ör. kişi zaten kayıtlı) kapat ve devam et
        try:
            hala_acik = tel_el.is_displayed()
        except (StaleElementReferenceException, WebDriverException):
            hala_acik = False
        if hala_acik:
            self.esc(2)
            raise AdimHatasi("Kayıt formu kapanmadı (kişi zaten kayıtlı olabilir)")
        return None

    def form_alanlari(self, sure=10):
        """Ekranda en üstte görünen (başka panelin altında kalmayan) yazı alanlarını sırasıyla döndürür.
        WhatsApp bazı alanları <input>, bazılarını contenteditable <div> olarak çizer; ikisi de alınır.
        Sohbet penceresindeki (#main) mesaj kutusu hariç tutulur."""
        js = """
        const sonuc = [];
        document.querySelectorAll('input, [contenteditable="true"]').forEach(el => {
            if (el.closest('#main')) return;
            const tip = (el.getAttribute('type') || '').toLowerCase();
            if (['checkbox', 'radio', 'hidden', 'file', 'submit', 'button'].includes(tip)) return;
            const r = el.getBoundingClientRect();
            if (r.width < 2 || r.height < 2) return;
            const isabet = document.elementFromPoint(r.left + r.width / 2, r.top + r.height / 2);
            let kutu = el;
            for (let i = 0; i < 3 && kutu.parentElement; i++) kutu = kutu.parentElement;
            if (isabet && kutu.contains(isabet)) sonuc.push(el);
        });
        return sonuc;
        """
        bitis = time.time() + sure
        while True:
            alanlar = self.d.execute_script(js)
            if len(alanlar) >= 2 or time.time() >= bitis:
                return alanlar
            time.sleep(0.5)

    def etiket(self, el):
        """Alanın görünen etiketini döndürür: aria-label, placeholder, aria-labelledby, <label>
        ya da (hiçbiri yoksa) yalnızca bu alanı içeren en yakın kapsayıcının yazısı ("Soyad" gibi)."""
        return (self.d.execute_script("""
            const el = arguments[0];
            let t = el.getAttribute('aria-label') || el.getAttribute('placeholder')
                    || el.getAttribute('data-placeholder') || '';
            if (!t && el.getAttribute('aria-labelledby'))
                t = el.getAttribute('aria-labelledby').split(' ')
                      .map(id => (document.getElementById(id) || {}).textContent || '').join(' ');
            if (!t && el.id) {
                const l = document.querySelector('label[for="' + el.id + '"]');
                if (l) t = l.textContent;
            }
            const secici = 'input, [contenteditable="true"]';
            let kap = el.parentElement;
            for (let i = 0; !t && kap && i < 5; i++, kap = kap.parentElement) {
                if (kap.querySelectorAll(secici).length > 1) break;
                t = (kap.innerText || '').trim();
            }
            return t;
        """, el) or "").strip()

    def alan_turu(self, el, etiket):
        """Alanı etiketinden ve tipinden 'ad', 'soyad', 'kullanici', 'telefon' ya da '' olarak sınıflar."""
        if any(k in etiket for k in ("kullanıcı", "kullanici", "username", "@")):
            return "kullanici"
        if any(k in etiket for k in ("telefon", "phone")) or el.get_attribute("type") == "tel" \
                or el.get_attribute("inputmode") in ("tel", "numeric"):
            return "telefon"
        if etiket.startswith(("soyad", "last name", "surname")):
            return "soyad"
        if etiket.startswith(("ad", "first name", "name", "isim")):
            return "ad"
        return ""

    def ustteki_tik(self):
        """Ekranda en üstte görünen (panel arkasında kalmayan) tik ikonunu bulur; sohbet alanı hariç."""
        return self.d.execute_script("""
            for (const el of document.querySelectorAll('[data-icon*="checkmark"]')) {
                if (el.closest('#main') || el.closest('#pane-side')) continue;
                const r = el.getBoundingClientRect();
                if (r.width < 2 || r.height < 2) continue;
                const isabet = document.elementFromPoint(r.left + r.width / 2, r.top + r.height / 2);
                const buton = el.closest('button, [role="button"]') || el;
                if (isabet && buton.contains(isabet)) return el;
            }
            return null;
        """)

    def alan_degeri(self, el):
        if el.tag_name.lower() == "input":
            return (el.get_property("value") or "").strip()
        return el.text.strip()

    def alana_yaz(self, el, metin, ad, panel, uye):
        """Alana yazar ve gerçekten yazıldığını doğrular; olmazsa klavye ile tekrar dener."""
        self.yaz(el, metin)
        time.sleep(0.3)
        if self.alan_degeri(el).replace(" ", "") == metin.replace(" ", ""):
            return
        el.click()
        ActionChains(self.d).key_down(Keys.CONTROL).send_keys("a").key_up(Keys.CONTROL) \
            .send_keys(Keys.BACKSPACE).send_keys(metin).perform()
        time.sleep(0.3)
        if self.alan_degeri(el).replace(" ", "") != metin.replace(" ", ""):
            self.form_dok(panel, uye)
            self.esc(2)
            raise AdimHatasi(f"'{ad}' alanı doldurulamadı")

    def form_paneli(self, alan):
        """Alanı içeren 'Yeni kişi' panelini döndürür (bulunamazsa sayfanın tamamı)."""
        basliklar = " or ".join(f"normalize-space(.)={lit(t)}" for t in T_YENI_KISI)
        try:
            return alan.find_element(By.XPATH, f"./ancestor::*[.//*[{basliklar}]][1]")
        except WebDriverException:
            return self.d.find_element(By.TAG_NAME, "body")

    def form_dok(self, panel, uye):
        """Hata ayıklama için yalnızca formun HTML'ini kaydeder (sohbetler dahil edilmez)."""
        try:
            if panel.tag_name.lower() == "body":
                return
            yol = HATA_KLASORU / f"{datetime.now():%H%M%S}_{uye['sira']}_form.html"
            yol.write_text(panel.get_attribute("outerHTML"), encoding="utf-8")
        except WebDriverException:
            pass

    # --- grup
    def grubu_ac(self):
        baslik = f"//div[@id='main']//header//*[@title={lit(GRUP_ADI)} or normalize-space(.)={lit(GRUP_ADI)}]"
        if self.bul(baslik, 1):
            return
        self.ara(GRUP_ADI)
        if not self.tikla(f"//div[@id='pane-side']//span[@title={lit(GRUP_ADI)}]", 10):
            raise AdimHatasi(f"'{GRUP_ADI}' grubu bulunamadı")
        if not self.bul(baslik, 10):
            raise AdimHatasi("Grup sohbeti açılmadı")
        self.esc()  # arama kutusunu temizle

    def grup_bilgisini_ac(self):
        if self.bul(X_TAM(T_UYE_EKLE), 1):
            return
        self.grubu_ac()
        self.tikla("//div[@id='main']//header//*[@role='button'] | //div[@id='main']//header", 5)
        if not self.bul(X_TAM(T_UYE_EKLE), 10):
            raise AdimHatasi("Grup bilgisinde 'Kişi ekle' bulunamadı (grup yöneticisi misin?)")

    def gruba_ekle(self, uye):
        """EKLENDI, ZATEN_GRUPTA veya DAVET_BEKLIYOR döner."""
        self.grup_bilgisini_ac()
        self.tikla(X_TAM(T_UYE_EKLE))

        kutu = self.bul([DIALOG + "div[@contenteditable='true']", DIALOG + "input"], 10)
        if not kutu:
            raise AdimHatasi("Kişi ekleme penceresinde arama kutusu yok")

        tam_ad = f"{uye['isim']} {uye['soyisim']}"
        satir = None
        for aranan in (uye["tel"], tam_ad):
            self.yaz(kutu, aranan)
            time.sleep(2.5)
            if self.bul(DIALOG + X_ICERIR(K_ZATEN, ""), 1):
                self.esc()
                return ZATEN_GRUPTA
            satir = self.bul([
                DIALOG + f"span[@title={lit(tam_ad)}]",
                DIALOG + "div[@role='listitem'][.//span[@title]]",
                DIALOG + "div[@role='button'][.//span[@title]]",
            ], 3)
            if satir:
                break
        if not satir:
            self.esc()
            return DAVET_BEKLIYOR  # rehberde bulunamadı -> link ile davet edilecek

        satir.click()
        time.sleep(1)
        if not self.tikla([
            DIALOG + "*[@data-icon='checkmark-medium']",
            DIALOG + "*[@data-icon='checkmark']",
            DIALOG + "*[@aria-label='Onayla' or @aria-label='Confirm']",
        ], 5):
            self.esc()
            raise AdimHatasi("Onay (✓) butonu bulunamadı")

        # "X kişisi gruba eklensin mi?" onayı
        if not self.tikla(X_TAM(T_EKLE, DIALOG), 5):
            self.esc()
            raise AdimHatasi("Eklemeyi onaylama butonu bulunamadı")

        # Gizlilik engeli varsa WhatsApp "davet gönder" penceresi açar
        if self.bul(DIALOG + X_ICERIR(K_DAVET, ""), 8):
            if not self.tikla(X_TAM(T_IPTAL, DIALOG), 3):
                self.esc()
            return DAVET_BEKLIYOR
        return EKLENDI

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
            if kaydet:
                try:
                    if wa.kisi_kaydet(uye) == WA_YOK:
                        log(uye["sira"], ad, "WhatsApp kullanmıyor, atlanıyor")
                        guncelle(rapor, uye, WA_YOK)
                        continue
                    log(uye["sira"], ad, "rehbere kaydedildi")
                except AdimHatasi as e:
                    log(uye["sira"], ad, f"kayıt uyarısı: {e} (gruba eklemeye devam)")
                    wa.ekran_goruntusu(f"{uye['sira']}_kayit")

            sonuc = wa.gruba_ekle(uye)
            aciklama = {
                EKLENDI: "Gruba eklendi",
                ZATEN_GRUPTA: "Zaten grupta",
                DAVET_BEKLIYOR: "Doğrudan eklenemedi, davet linki gönderilecek",
            }[sonuc]
            log(uye["sira"], ad, aciklama)
            guncelle(rapor, uye, sonuc, aciklama)
        except (AdimHatasi, WebDriverException) as e:
            hata = str(e).splitlines()[0][:150]
            log(uye["sira"], ad, f"HATA: {hata} -> davet linki gönderilecek")
            wa.ekran_goruntusu(f"{uye['sira']}_grup")
            guncelle(rapor, uye, DAVET_BEKLIYOR, f"Ekleme hatası: {hata}")
            wa.esc(3)
        mola(i)


def asama2_davet(wa, uyeler, rapor):
    print(f"\n=== 2. AŞAMA: Davet mesajı gönderme ({len(uyeler)} kişi) ===")
    for i, uye in enumerate(uyeler, 1):
        ad = f"{uye['isim']} {uye['soyisim']}"
        try:
            sonuc = wa.mesaj_gonder(uye["tel"], mesaj_metni(uye))
            aciklama = "Davet mesajı gönderildi" if sonuc == DAVET_GONDERILDI else "Numara WhatsApp'ta yok"
            log(uye["sira"], ad, aciklama)
            guncelle(rapor, uye, sonuc, aciklama)
        except (AdimHatasi, WebDriverException) as e:
            hata = str(e).splitlines()[0][:150]
            log(uye["sira"], ad, f"HATA: {hata}")
            wa.ekran_goruntusu(f"{uye['sira']}_mesaj")
            guncelle(rapor, uye, DAVET_BEKLIYOR, f"Mesaj hatası: {hata}")
        mola(i)


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
    args = p.parse_args()

    global RAPOR
    uyeler, gecersizler = uyeleri_oku()
    if args.test:
        RAPOR = KLASOR / "test_raporu.csv"   # deneme sonuçları asıl raporu kirletmesin
        tel = normalize_tel(args.test)
        if not tel:
            sys.exit("Test numarası geçersiz. Örnek: 5321234567")
        uyeler = [{"sira": 0, "isim": "Test", "soyisim": "Kişi", "ham_tel": args.test, "tel": tel}]
        gecersizler = []

    rapor = rapor_oku()
    bekleyen = [u for u in uyeler if rapor.get(u["tel"], {}).get("Durum") not in BITMIS]
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
