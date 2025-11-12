#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import json
import datetime as _dt
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
        safe_subject = subject.replace("'", "''")
        # $orderby alanı $filter içinde de (ve önce) geçmeli kuralına uyar
        filter_expr = f"receivedDateTime ge 1900-01-01T00:00:00Z and subject eq '{safe_subject}'"
        params = {
            "$select": "sender,subject,receivedDateTime",
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


    # === İsim normalize ===
    @staticmethod
    def _clean_name(name):
        if not name:
            return ""
        name = re.sub(r"\s+", " ", name)
        return name.strip(" \t\r\n-–—.")

    # === Yardımcılar: token sınıfları ===
    @staticmethod
    def _is_all_caps_word(tok: str) -> bool:
        # Türkçe uyumlu ALL-CAPS: sadece harflerden oluşsun ve en az 2 harf olsun
        letters = re.sub(r"[^A-Za-zÇĞİÖŞÜçğıöşü]", "", tok)
        return len(letters) >= 2 and letters == letters.upper()

    @staticmethod
    def _is_title_like(tok: str) -> bool:
        # TitleCase/MixedCase isim: İlk harf büyük, devamı küçük/karmışık (oğlu- gibi tire destekli)
        return bool(re.match(r"^[A-ZÇĞİÖŞÜ][A-Za-zÇĞİÖŞÜçğıöşü'’\-]+$", tok))

    # === NEW HIRE: Ad Soyad = (sicil no'dan sonra) ilk ALL-CAPS gelene kadarki 2–6 token ===
    @staticmethod
    def _extract_name_new_hire(text):
        # 1) CEP TELEFONU ... <sicil> sonrası metni al; yoksa SİCİL NO ile dene
        m = re.search(r"CEP\s*TELEFONU\b.*?\b(\d{3,})\b(?P<rest>.*)", text, flags=re.IGNORECASE | re.DOTALL)
        if not m:
            m = re.search(r"S[İI]C[İI]L\s*NO\b.*?\b(\d{3,})\b(?P<rest>.*)", text, flags=re.IGNORECASE | re.DOTALL)
        rest = m.group("rest") if m else text

        # 2) Tokenize (TR harfler dahil)
        tokens = re.findall(r"[A-Za-zÇĞİÖŞÜçğıöşü'’\-]+", rest)

        connectors = {"de","da","van","von","bin","ibn","al","el","oğlu","oglu","del","di"}
        name_tokens = []

        for tok in tokens:
            low = tok.lower()

            # POZİSYON kolonu ALL-CAPS olduğundan: ilk ALL-CAPS kelimede kes
            if OutlookGraphAppOnly._is_all_caps_word(tok):
                break

            # İsim kuralları: TitleCase veya bağlaç
            if OutlookGraphAppOnly._is_title_like(tok) or low in connectors:
                name_tokens.append(tok)
                if len(name_tokens) >= 6:  # çok uzamasın
                    break
                continue

            # İsim başladıysa ve bu token isim formatına uymuyorsa kes
            if name_tokens:
                break
            continue

        # Sonda bağlaç kaldıysa at
        while name_tokens and name_tokens[-1].lower() in connectors:
            name_tokens.pop()

        name = " ".join(name_tokens)
        name = OutlookGraphAppOnly._clean_name(name)

        # 3) Fallback: "ADI SOYADI : <Ad Soyad>"
        if not name:
            m2 = re.search(
                r"ADI\s*SOYADI\s*[:\-]?\s*(?P<name>(?:[A-ZÇĞİÖŞÜ][A-Za-zÇĞİÖŞÜçğıöşü'’\-]+(?:\s+[A-ZÇĞİÖŞÜ][A-Za-zÇĞİÖŞÜçğıöşü'’\-]+){1,5}))",
                text, flags=re.IGNORECASE
            )
            if m2:
                name = OutlookGraphAppOnly._clean_name(m2.group("name"))

        # 4) Minimum 2 token şartı (Ad + Soyad)
        if len(name.split()) < 2:
            return ""
        return name

    # === Yardımcı: tarih yakala (TR formatlarına toleranslı) ===
    @staticmethod
    def _extract_first_date(text):
        """
        Metinden ilk 'olası tarih'i döndürür.
        Öncelik: 'tarihi itibari ile' civarındaki DD.MM.YYYY / DD/MM/YYYY / YYYY-MM-DD.
        Bulunamazsa ilk tarih benzeri ifadeyi arar.
        Return: datetime.date veya None
        """
        if not text:
            return None
        t = text.replace("\r\n", "\n").replace("\r", "\n")
        t = re.sub(r"\s+", " ", t)

        # 1) 'tarihi itibari ile' civarında 0-30 karakter içinde bir tarih ara
        ctx_pat = re.compile(r"tarihi\s+itibari\s+ile.{0,30}", re.IGNORECASE)
        date_pats = [
            r"(?P<d>\b\d{1,2}[./]\d{1,2}[./]\d{2,4}\b)",  # 17.10.2025, 17/10/25
            r"(?P<d>\b\d{4}-\d{1,2}-\d{1,2}\b)",          # 2025-10-17
        ]
        # önce bağlam içinde tara
        mctx = ctx_pat.search(t)
        search_ranges = []
        if mctx:
            s, e = mctx.start(), mctx.end()
            search_ranges.append(t[max(0, s-40): min(len(t), e+40)])
        # sonra tüm metin fallback
        search_ranges.append(t)

        def _parse_candidate(s):
            # DD.MM.YYYY veya DD/MM/YYYY
            m = re.search(r"\b(\d{1,2})[./](\d{1,2})[./](\d{2,4})\b", s)
            if m:
                d, M, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
                if y < 100:  # 25 -> 2025 varsay
                    y += 2000
                try:
                    return _dt.date(y, M, d)
                except ValueError:
                    pass
            # YYYY-MM-DD
            m = re.search(r"\b(\d{4})-(\d{1,2})-(\d{1,2})\b", s)
            if m:
                y, M, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
                try:
                    return _dt.date(y, M, d)
                except ValueError:
                    pass
            return None

        for seg in search_ranges:
            for p in date_pats:
                m = re.search(p, seg)
                if m:
                    got = _parse_candidate(m.group("d"))
                    if got:
                        return got
        return None

    # === TERMINATION Regex (iki formatı da doğru parse eder) ===
    @staticmethod
    def _extract_name_termination(text):
        """
        1) '... sicili ile çalışan <Ad Soyad> isimli çalışan için ...'
        2) '<Ad Soyad> için ...'
        """
        txt = (text or "").replace("\r\n", "\n").replace("\r", "\n")
        txt = re.sub(r"[ \t]+", " ", txt)
        txt = re.sub(r"\n+", " ", txt).strip()
        name_token = r"(?:[A-ZÇĞİÖŞÜ][A-Za-zÇĞİÖŞÜçğıöşü'’\-]+|[A-ZÇĞİÖŞÜ]{2,})"
        name_pattern = rf"(?P<name>{name_token}(?:\s+{name_token}){{1,5}})"
        pat_with_label = re.compile(
            rf"sicil\w*\s+ile\s+çalışan\s+{name_pattern}\s+isimli\s+çalışan\s+için\b",
            flags=re.IGNORECASE
        )
        pat_plain = re.compile(rf"\b{name_pattern}\s+için\b", flags=re.IGNORECASE)
        pat_without_label = re.compile(
            rf"sicil\w*\s+ile\s+çalışan\s+{name_pattern}\s+için\b", flags=re.IGNORECASE
        )
        for pat in (pat_with_label, pat_without_label, pat_plain):
            m = pat.search(txt)
            if m:
                raw = m.group("name").strip()
                raw = re.sub(r"\s+isimli\s+çalışan\b.*$", "", raw, flags=re.IGNORECASE).strip()
                connectors = {"de","da","van","von","bin","ibn","al","el","oğlu","oglu","del","di","di’"}
                toks = raw.split()
                while toks and toks[-1].lower() in connectors:
                    toks.pop()
                name = " ".join(toks)
                if len(name.split()) >= 2:
                    return OutlookGraphAppOnly._clean_name(name)
        return ""

    # === ISO yardımcıları ===
    @staticmethod
    def _to_iso_date(d: _dt.date) -> str:
        return d.isoformat() if d else ""

    @staticmethod
    def _to_iso_dt(dt: _dt.datetime) -> str:
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=_dt.timezone.utc)
        return dt.astimezone(_dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    @staticmethod
    def _midnight_utc_after_days(d: _dt.date, days: int) -> str:
        if not d:
            return ""
        target = _dt.datetime(d.year, d.month, d.day, tzinfo=_dt.timezone.utc) + _dt.timedelta(days=days)
        target = target.replace(hour=0, minute=0, second=0, microsecond=0)
        return OutlookGraphAppOnly._to_iso_dt(target)

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
                    self.logger.info(f"[NEW_HIRE] no match. preview={prev}")
        return {"success": True, "names": names}

    # === ACTION 2: Termination -> Detaylı JSON + geri uyumluluk ===
    def list_termination_messages(self, tenant_id, client_id, client_secret, mailbox, top=None):
        subject = "[Kurum Dışı] Çalışan İlişik Kesme Bildirimi"
        items = self._fetch_by_exact_subject(tenant_id, client_id, client_secret, mailbox, subject, top)
        names = []
        out_items = []
        for it in items:
            body = self._get_body_text(it)
            name = self._extract_name_termination(body)
            term_date = self._extract_first_date(body)  # datetime.date veya None

            if name:
                names.append(name)

            received_iso = it.get("receivedDateTime", "") or ""
            activate_at = self._midnight_utc_after_days(term_date, 3) if term_date else ""

            out_items.append({
                "name": name or "",
                "mail_received_at": received_iso,
                "termination_date": self._to_iso_date(term_date) if term_date else "",
                "activate_at": activate_at,
            })

            if self.logger:
                self.logger.info(f"[TERMINATION] name='{name}' term_date='{term_date}' activate_at='{activate_at}'")

        return {"success": True, "items": out_items, "names": names}


if __name__ == "__main__":
    OutlookGraphAppOnly.run()

