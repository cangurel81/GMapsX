import time
import random
import re
import urllib.parse
import threading

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import (
    TimeoutException, StaleElementReferenceException
)
from webdriver_manager.chrome import ChromeDriverManager


class GoogleMapsScraper:

    SEL = {
        "feed": [
            "div[role='feed']",
            "#QA0Szd > div > div > div.w6VYqd > div:nth-child(2)",
        ],
        "place_links": [
            "div[role='feed'] a[href*='/maps/place/']",
            "div.Nv2PK a[href*='/maps/place/']",
            "a[href*='/maps/place/']",
        ],
        "name": [
            "h1.DUwDvf",
            "h1[class*='fontHeadline']",
            "h1",
        ],
        "category": [
            "button.DkEaL",
            "span.mgr77e button",
            "div.LBgpqf button.DkEaL",
            "button[aria-label*='kategori']",
        ],
        "rating": [
            "div.F7nice > span:first-child",
            "span.ceNzKf",
            "div.F7nice span[aria-hidden='true']",
            "span[aria-label*='yıldız']",
        ],
        "review_count": [
            "div.F7nice span[aria-label*='yorum']",
            "div.F7nice span[aria-label*='review']",
            "div.F7nice span[aria-label*='Review']",
            "div.F7nice > span:nth-child(2) span",
            "span[aria-label*='değerlendirme']",
        ],
        "address": [
            "button[data-item-id='address'] div.Io6YTe",
            "button[data-item-id='address']",
            "button[aria-label*='Adres']",
            "button[aria-label*='Address']",
            "div.TrHaRc button",
        ],
        "phone": [
            "button[data-item-id*='phone']",
            "button[aria-label*='Telefon']",
            "button[aria-label*='Phone']",
            "a[href^='tel:']",
            "button[aria-label*='Ara']",
        ],
        "website": [
            "a[data-item-id='authority']",
            "a[aria-label*='Web sitesi']",
            "a[aria-label*='website']",
            "a[aria-label*='Website']",
        ],
    }

    def __init__(self, headless: bool = False,
                 on_result=None, on_log=None):
        self.headless = headless
        self.on_result = on_result
        self.on_log = on_log
        self.driver = None
        self.running = False
        self.results: list[dict] = []

    def _log(self, msg: str):
        if self.on_log:
            self.on_log(msg)

    def _get_text(self, selectors: list[str], parent=None, driver=None) -> str:
        d = driver or self.driver
        ctx = parent or d
        for sel in selectors:
            try:
                el = ctx.find_element(By.CSS_SELECTOR, sel)
                txt = el.text.strip()
                if txt:
                    return txt
            except Exception:
                continue
        return ""

    def _get_attr(self, selectors: list[str], attr: str, parent=None, driver=None) -> str:
        d = driver or self.driver
        ctx = parent or d
        for sel in selectors:
            try:
                el = ctx.find_element(By.CSS_SELECTOR, sel)
                val = el.get_attribute(attr)
                if val:
                    return val.strip()
            except Exception:
                continue
        return ""

    def _build_driver(self):
        opts = webdriver.ChromeOptions()
        if self.headless:
            opts.add_argument("--headless=new")
        opts.add_argument("--no-sandbox")
        opts.add_argument("--disable-dev-shm-usage")
        opts.add_argument("--disable-blink-features=AutomationControlled")
        opts.add_argument("--window-size=1440,900")
        opts.add_argument("--disable-gpu")
        opts.add_argument("--lang=tr-TR,tr;q=0.9,en;q=0.8")
        opts.add_argument(
            "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36"
        )
        opts.add_experimental_option("excludeSwitches", ["enable-automation"])
        opts.add_experimental_option("useAutomationExtension", False)
        opts.add_experimental_option("prefs", {
            "profile.default_content_setting_values.notifications": 2
        })

        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=opts)

        driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
            "source": "Object.defineProperty(navigator,'webdriver',{get:()=>undefined})"
        })
        return driver

    def setup_driver(self):
        self.driver = self._build_driver()

    def close(self):
        self.running = False
        if self.driver:
            try:
                self.driver.quit()
            except Exception:
                pass
            self.driver = None

    def _handle_consent(self, driver=None):
        d = driver or self.driver
        patterns = [
            "button#L2AGLb",
            "button[aria-label*='Kabul']",
            "button[aria-label*='Accept all']",
            "button[aria-label*='Tümünü kabul']",
            "form:nth-child(2) button",
        ]
        for sel in patterns:
            try:
                btn = WebDriverWait(d, 4).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, sel))
                )
                btn.click()
                time.sleep(1.2)
                return
            except Exception:
                continue

    def _collect_urls(self, max_count: int) -> list[str]:
        urls: list[str] = []
        seen: set[str] = set()

        self._log("Sonuç paneli bekleniyor…")

        feed = None
        for sel in self.SEL["feed"]:
            try:
                feed = WebDriverWait(self.driver, 12).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, sel))
                )
                break
            except Exception:
                continue

        if not feed:
            self._log("Sonuç paneli bulunamadı.")
            return urls

        no_new_streak = 0
        last_count = 0

        while self.running and len(urls) < max_count:
            for sel in self.SEL["place_links"]:
                try:
                    links = self.driver.find_elements(By.CSS_SELECTOR, sel)
                    for a in links:
                        href = a.get_attribute("href") or ""
                        if "/maps/place/" in href and href not in seen:
                            seen.add(href)
                            clean = href.split("?")[0]
                            urls.append(clean)
                    break
                except StaleElementReferenceException:
                    continue

            self._log(f"Toplanan URL: {len(urls)} / kaydırılıyor…")

            if len(urls) >= max_count:
                break

            if len(urls) == last_count:
                no_new_streak += 1
            else:
                no_new_streak = 0
            last_count = len(urls)

            if no_new_streak >= 4:
                self._log("Daha fazla sonuç yok.")
                break

            try:
                self.driver.execute_script("arguments[0].scrollBy(0, 2000)", feed)
            except Exception:
                try:
                    items = self.driver.find_elements(
                        By.CSS_SELECTOR, "div[role='feed'] > div"
                    )
                    if items:
                        self.driver.execute_script(
                            "arguments[0].scrollIntoView()", items[-1]
                        )
                except Exception:
                    break

            time.sleep(random.uniform(1.8, 2.8))

        return urls[:max_count]

    def _extract_place(self, url: str, driver=None) -> dict | None:
        d = driver or self.driver
        try:
            d.get(url)
            time.sleep(random.uniform(2.0, 3.2))

            try:
                WebDriverWait(d, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "h1"))
                )
            except TimeoutException:
                return None

            time.sleep(1.5)

            data: dict[str, str] = {}

            data["İsim"] = self._get_text(self.SEL["name"], driver=d)
            if not data["İsim"]:
                return None

            data["Kategori"] = self._get_text(self.SEL["category"], driver=d)
            if not data["Kategori"]:
                data["Kategori"] = self._get_attr_from_label(
                    r"^[^,]*(cafe|restoran|otel|doktor|avukat|dişçi|market|barber)", driver=d
                )

            raw_rating = self._get_text(self.SEL["rating"], driver=d)
            data["Puan"] = re.sub(r"[^\d.,]", "", raw_rating)[:4]

            raw_reviews = self._get_text(self.SEL["review_count"], driver=d)
            data["Yorum Sayısı"] = re.sub(r"[^\d]", "", raw_reviews)

            data["Adres"] = self._extract_field("address", self.SEL["address"], driver=d)
            data["Telefon"] = self._extract_field("phone", self.SEL["phone"], driver=d)
            data["Website"] = self._extract_field("website", self.SEL["website"], driver=d)

            current_url = d.current_url
            data["Google Maps URL"] = current_url

            return data

        except Exception as e:
            self._log(f"  ⚠ Hata: {str(e)[:80]}")
            return None

    def _get_attr_from_label(self, pattern: str, driver=None) -> str:
        d = driver or self.driver
        try:
            script = """
            let labels = Array.from(document.querySelectorAll('[aria-label]'))
                .map(el => el.getAttribute('aria-label'))
                .filter(label => label && label.length > 0);
            return labels;
            """
            labels = d.execute_script(script)
            for label in labels:
                if re.search(pattern, label, re.IGNORECASE):
                    return label[:100]
            return ""
        except Exception:
            return ""

    def _extract_field(self, field_type: str, selectors: list[str],
                       driver=None) -> str:
        d = driver or self.driver
        value = ""

        if field_type != "phone":
            value = self._get_text(selectors, driver=d)
            if value:
                return value

        for sel in selectors:
            try:
                el = d.find_element(By.CSS_SELECTOR, sel)
                label = el.get_attribute("aria-label") or ""
                text = el.text.strip() or ""

                if field_type == "phone" and (label or text):
                    raw = label or text
                    raw = re.sub(r"(telefon|phone|aç|call|ara|dial).*?:", "",
                                raw, flags=re.IGNORECASE).strip()
                    phone = re.sub(r"[^\d+\-() ]", "", raw).strip()
                    if phone and len(phone) > 5 and re.search(r"\d", phone):
                        return phone

                if field_type == "address" and label:
                    addr = label.replace("Adresi kopyala", "") \
                                .replace("Copy address", "") \
                                .replace("Address", "").strip()
                    if len(addr) > 5 and not addr.startswith("http"):
                        return addr

                if field_type == "website" and label:
                    if "http" not in label and len(label) > 3:
                        return label[:100]
            except Exception:
                continue

        if field_type == "phone":
            try:
                tel_el = d.find_element(By.CSS_SELECTOR, "a[href^='tel:']")
                value = tel_el.get_attribute("href").replace("tel:", "").strip()
                if value and re.search(r"\d", value):
                    return value
            except Exception:
                pass

        if field_type == "website":
            for sel in selectors:
                try:
                    el = d.find_element(By.CSS_SELECTOR, sel)
                    href = el.get_attribute("href") or ""
                    if href.startswith("http") and "google.com" not in href:
                        return href
                except Exception:
                    continue

        return value

    def search(self, query: str, location: str,
               max_results: int = 20, language: str = "tr",
               threads: int = 1) -> list[dict]:
        self.running = True
        self.results = []

        search_query = f"{query} {location}"
        encoded = urllib.parse.quote(search_query)
        maps_url = f"https://www.google.com/maps/search/{encoded}/?hl={language}"

        self._log("Google Maps'e bağlanılıyor…")
        self.driver.get(maps_url)
        time.sleep(3)
        self._handle_consent()
        time.sleep(1.5)

        place_urls = self._collect_urls(max_results)
        total = len(place_urls)
        self._log(f"{total} yer bulundu. Detaylar çekiliyor…")

        if threads <= 1 or total <= 1:
            for idx, purl in enumerate(place_urls, 1):
                if not self.running:
                    break
                self._log(f"[{idx}/{total}] İşleniyor…")
                record = self._extract_place(purl)
                if record:
                    self.results.append(record)
                    if self.on_result:
                        self.on_result(record, len(self.results))
                time.sleep(random.uniform(0.8, 1.8))
        else:
            self._search_multi(place_urls, threads)

        self._log(f"Tamamlandı ✓  {len(self.results)} kayıt")
        return self.results

    def _search_multi(self, urls: list[str], max_workers: int):
        total = len(urls)
        lock = threading.Lock()
        idx = [0]

        def worker_thread():
            local_driver = self._build_driver()
            try:
                local_driver.get("https://www.google.com/maps")
                time.sleep(2)
                self._handle_consent(driver=local_driver)
                while self.running:
                    with lock:
                        if idx[0] >= total:
                            break
                        url = urls[idx[0]]
                        idx[0] += 1
                    try:
                        record = self._extract_place(url, driver=local_driver)
                        if record:
                            with lock:
                                self.results.append(record)
                                done = len(self.results)
                                if self.on_result:
                                    self.on_result(record, done)
                                self._log(f"[{done}/{total}] İşleniyor… ({max_workers} thread)")
                    except Exception as e:
                        self._log(f"  ⚠ Hata: {str(e)[:60]}")
            finally:
                try:
                    local_driver.quit()
                except Exception:
                    pass

        threads = []
        for _ in range(min(max_workers, total)):
            t = threading.Thread(target=worker_thread, daemon=True)
            t.start()
            threads.append(t)
        for t in threads:
            t.join()
