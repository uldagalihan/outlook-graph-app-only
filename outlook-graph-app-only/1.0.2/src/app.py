import json
import re
import requests
from walkoff_app_sdk.app_base import AppBase

GRAPH = "https://graph.microsoft.com/v1.0"

class OutlookGraphAppOnly(AppBase):
    __version__ = "1.0.2"
    app_name = "Outlook Graph AppOnly"

    def __init__(self, redis=None, logger=None, **kwargs):
        super().__init__(redis=redis, logger=logger, **kwargs)

    # === Auth ===
    def _token(self, tenant_id, client_id, client_secret):
        url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
        data = {
            "grant_type": "client_credentials",
            "client_id": client_id,
            "client_secret": client_secret,
            "scope": "https://graph.microsoft.com/.default",
        }
        r = requests.post(url, data=data, timeout=30)
        r.raise_for_status()
        return r.json()["access_token"]

    # === HTTP GET helper (Prefer text body) ===
    def _get(self, url, tok, params=None, prefer_text_body=True):
        headers = {"Authorization": f"Bearer {tok}"}
        if prefer_text_body:
            headers["Prefer"] = 'outlook.body-content-type="text"'

        # Log: encode edilmiş tam URL + params
        if self.logger:
            try:
                prepped = requests.Request("GET", url, params=params).prepare()
                self.logger.info(f"[Graph GET] url={prepped.url} params={json.dumps(params, ensure_ascii=False)}")
            except Exception:
                pass

        r = requests.get(url, headers=headers, params=params, timeout=30)
        r.raise_for_status()
        return r.json()

    # === Mesajları subject = '...' ile getir (body dahil) ===
    def _fetch_by_exact_subject(self, tenant_id, client_id, client_secret, mailbox, subject, top=None):
        tok = self._token(tenant_id, client_id, client_secret)
        url = f"{GRAPH}/users/{mailbox}/messages"

        # OData tek tırnak kaçışı: '' (iki tek tırnak)
        safe_subject = subject.replace("'", "''")

        # $orderby=receivedDateTime desc -> Graph kuralı gereği $filter içinde önce olmalı
        filter_expr = f"receivedDateTime ge 1900-01-01T00:00:00Z and subject eq '{safe_subject}'"

        params = {
            "$select": "id,sender,subject,receivedDateTime,body,uniqueBody,bodyPreview",
            "$filter": filter_expr,
            "$orderby": "receivedDateTime desc",
        }
        if top is not None:
            try:
                t = max(1, min(1000, int(top)))
            except Exception:
                t = 10
            params["$top"] = t

        data = self._get(url, tok, params=params, prefer_text_body=True)
        return data.get("value", [])

    # === Body metnini al (uniqueBody varsa onu tercih et) ===
    @staticmethod
    def _get_body_text(item):
        body = ""
        if isinstance(item, dict):
            ub = item.get("uniqueBody") or {}
            b  = item.get("body") or {}
            body = (ub.get("content") or b.get("content") or item.get("bodyPreview") or "")
        # Satırları koruyup boşlukları sadeleştir
        body = body.replace("\r\n", "\n").replace("\r", "\n")
        body = re.sub(r"[ \t]+", " ", body)
        body = re.sub(r"\n{2,}", "\n", body)
        return body.strip()

    # === İsim normalize ===
    @staticmethod
    def _clean_name(name):
        if not name:
            return ""
        name = re.sub(r"\s+", " ", name)
        return name.strip(" \t\r\n-–—.")

    # === NEW HIRE: Akıllı ayrıştırıcı (pozisyonu yeme!)
    @staticmethod
    def _extract_name_new_hire(text):
        # 1) CEP TELEFONU ... <sicil>  sonrası metni yakala (yoksa SİCİL NO ile dene)
        m = re.search(r"CEP\s*TELEFONU\b.*?\b(\d{3,})\b(?P<rest>.*)", text, flags=re.IGNORECASE | re.DOTALL)
        if not m:
            m = re.search(r"S[İI]C[İI]L\s*NO\b.*?\b(\d{3,})\b(?P<rest>.*)", text, flags=re.IGNORECASE | re.DOTALL)
        rest = m.group("rest") if m else text

        # 2) Tokenize (Türkçe harfler ve isim bağlaçları destekli)
        tokens = re.findall(r"[A-Za-zÇĞİÖŞÜçğıöşü'’\-]+", rest)

        # İsimde kullanılabilecek bağlaç/küçük kelimeler
        connectors = {"de","da","van","von","bin","ibn","al","el","oğlu","oglu","del","di","di","di’"}
        # Pozisyon ve diğer kolon/ayraç stopword’leri
        stops = {
            # pozisyon & unvan
            "stajyer","staj","uzman","kıdemli","kidemli","mühendis","muhendis","müdür","mudur","yönetici","yonetici",
            "uzmanı","uzmani","analist","koordinatör","koordinator","danışman","danisman","asistan","sorumlu",
            "lider","lead","direktör","direktor","director","manager","yöneticisi","yoneticisi","developer","geliştirici",
            "devops","architect","mimar","engineer","officer","specialist","product","process","security","siber","güvenlik",
            # diğer kolonlar
            "pozisyon","masraf","yeri","bağlı","bagli","olduğu","oldugu","ilk","yönetici","yonetici","işyeri","isyeri",
            "lokasyon","yaka","tipi","işe","ise","başlama","baslama","tarihi","mail","grubu","cep","telefonu"
        }

        name_tokens = []
        started = False

        def is_name_token(tok: str) -> bool:
            # Büyük harfle başlayan Türkçe isim parçaları veya TAM büyük (kısaltma)
            return bool(re.match(r"^[A-ZÇĞİÖŞÜ][A-Za-zÇĞİÖŞÜçğıöşü'’\-]+$", tok)) or bool(re.match(r"^[A-ZÇĞİÖŞÜ]{2,}$", tok))

        for tok in tokens:
            low = tok.lower()

            # Durma koşulları: pozisyon/stops veya diğer kolonlar gelirse isim bitti
            if low in stops:
                break

            if is_name_token(tok) or low in connectors:
                name_tokens.append(tok)
                started = True
                # 2–6 token arası makul bir ad-soyad (çok uzarsa kesiyoruz)
                if len(name_tokens) >= 6:
                    break
                continue

            # isim başladıktan sonra isme uymayan bir token gelirse dur
            if started:
                break
            # başlamadıysa (ilk kelimeler isme uygun değilse), devam et
            continue

        name = " ".join(name_tokens)
        name = OutlookGraphAppOnly._clean_name(name)

        # Fallback: "ADI SOYADI : <Ad Soyad>"
        if not name:
            m2 = re.search(
                r"ADI\s*SOYADI\s*[:\-]?\s*(?P<name>(?:[A-ZÇĞİÖŞÜ][A-Za-zÇĞİÖŞÜçğıöşü'’\-]+(?:\s+[A-ZÇĞİÖŞÜ][A-Za-zÇĞİÖŞÜçğıöşü'’\-]+){1,5}))",
                flags=re.IGNORECASE
            )
            if m2:
                name = OutlookGraphAppOnly._clean_name(m2.group("name"))

        return name

    # === TERMINATION Regex ===
    @staticmethod
    def _extract_name_termination(text):
        pat = re.compile(
            r"sicili\s+ile\s+çalışan\s+(?P<name>.+?)\s+için\b",
            flags=re.IGNORECASE | re.DOTALL
        )
        m = pat.search(text)
        return OutlookGraphAppOnly._clean_name(m.group("name")) if m else ""

    # === ACTION 1: New Hire -> İSİM LİSTESİ ===
    def list_new_hire_messages(self, tenant_id, client_id, client_secret, mailbox, top=None):
        subject = "[Kurum Dışı] Şirkete Yeni Katılım - New Comer"
        items = self._fetch_by_exact_subject(tenant_id, client_id, client_secret, mailbox, subject, top)
        names = []
        for it in items:
            body = self._get_body_text(it)
            name = self._extract_name_new_hire(body)
            if name:
                names.append(name)
                if self.logger:
                    self.logger.info(f"[NEW_HIRE] matched name: {name}")
            else:
                if self.logger:
                    prev = (it.get("bodyPreview") or "")[:100]
                    self.logger.info(f"[NEW_HIRE] no match. preview={prev}")
        return {"success": True, "names": names}

    # === ACTION 2: Termination -> İSİM LİSTESİ ===
    def list_termination_messages(self, tenant_id, client_id, client_secret, mailbox, top=None):
        subject = "[Kurum Dışı] Çalışan İlişik Kesme Bildirimi"
        items = self._fetch_by_exact_subject(tenant_id, client_id, client_secret, mailbox, subject, top)
        names = []
        for it in items:
            body = self._get_body_text(it)
            name = self._extract_name_termination(body)
            if name:
                names.append(name)
                if self.logger:
                    self.logger.info(f"[TERMINATION] matched name: {name}")
            else:
                if self.logger:
                    self.logger.info("[TERMINATION] no match in body")
        return {"success": True, "names": names}

if __name__ == "__main__":
    OutlookGraphAppOnly.run()
