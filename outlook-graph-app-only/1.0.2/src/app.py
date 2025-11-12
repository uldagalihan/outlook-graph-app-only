#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import json
import datetime as _dt
import re
import requests
from walkoff_app_sdk.app_base import AppBase

GRAPH = "https://graph.microsoft.com/v1.0"

class OutlookGraphAppOnly(AppBase):
    __version__ = "1.2.0"
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

    # === Pagination aware fetch (subject exact veya contains) ===
    def _fetch_messages(self, tenant_id, client_id, client_secret, mailbox,
                        subject, mode="exact", max_items=None):
        """
        mode: 'exact' (eski davranış) | 'contains' (daha esnek)
        """
        tok = self._token(tenant_id, client_id, client_secret)
        url = f"{GRAPH}/users/{mailbox}/messages"

        if mode == "contains":
            filt = f"contains(subject,'{subject.replace(\"'\",\"''\")}')"
        else:
            safe_subject = subject.replace("'", "''")
            filt = f"receivedDateTime ge 1900-01-01T00:00:00Z and subject eq '{safe_subject}'"

        params = {
            "$select": "id,sender,subject,receivedDateTime,body,uniqueBody,bodyPreview",
            "$filter": filt,
            "$orderby": "receivedDateTime desc",
            "$top": min(1000, max_items) if max_items else 50  # sayfa başı
        }

        items = []
        page_url = url
        page_params = params
        while True:
            data = self._get(page_url, tok, params=page_params, prefer_text_body=True)
            vals = data.get("value", [])
            items.extend(vals)
            if max_items and len(items) >= max_items:
                items = items[:max_items]
                break
            next_link = data.get("@odata.nextLink")
            if not next_link:
                break
            page_url, page_params = next_link, None

        if self.logger:
            self.logger.info(f"[FETCH] mode={mode} subject='{subject}' -> {len(items)} items")
        return items

    # === Body metnini al (uniqueBody varsa onu tercih et) ===
    @staticmethod
    def _get_body_text(item):
        body = ""
        if isinstance(item, dict):
            ub = item.get("uniqueBody") or {}
            b  = item.get("body") or {}
            body = (ub.get("content") or b.get("content") or item.get("bodyPreview") or "")
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

    # === Yardımcılar ===
    @staticmethod
    def _is_all_caps_word(tok: str) -> bool:
        letters = re.sub(r"[^A-Za-zÇĞİÖŞÜçğıöşü]", "", tok)
        return len(letters) >= 2 and letters == letters.upper()

    @staticmethod
    def _is_title_like(tok: str) -> bool:
        return bool(re.match(r"^[A-ZÇĞİÖŞÜ][A-Za-zÇĞİÖŞÜçğıöşü'’\-]+$", tok))

    # === NEW HIRE extraction ===
    @staticmethod
    def _extract_name_new_hire(text):
        m = re.search(r"CEP\s*TELEFONU\b.*?\b(\d{3,})\b(?P<rest>.*)", text, flags=re.IGNORECASE | re.DOTALL)
        if not m:
            m = re.search(r"S[İI]C[İI]L\s*NO\b.*?\b(\d{3,})\b(?P<rest>.*)", text, flags=re.IGNORECASE | re.DOTALL)
        rest = m.group("rest") if m else text
        tokens = re.findall(r"[A-Za-zÇĞİÖŞÜçğıöşü'’\-]+", rest)
        connectors = {"de","da","van","von","bin","ibn","al","el","oğlu","oglu","del","di"}
        name_tokens = []
        for tok in tokens:
            low = tok.lower()
            if OutlookGraphAppOnly._is_all_caps_word(tok):
                break
            if OutlookGraphAppOnly._is_title_like(tok) or low in connectors:
                name_tokens.append(tok)
                if len(name_tokens) >= 6:
                    break
                continue
            if name_tokens:
                break
        while name_tokens and name_tokens[-1].lower() in connectors:
            name_tokens.pop()
        name = OutlookGraphAppOnly._clean_name(" ".join(name_tokens))
        if not name:
            m2 = re.search(
                r"ADI\s*SOYADI\s*[:\-]?\s*(?P<name>(?:[A-ZÇĞİÖŞÜ][A-Za-zÇĞİÖŞÜçğıöşü'’\-]+(?:\s+[A-ZÇĞİÖŞÜ][A-Za-zÇĞİÖŞÜçğıöşü'’\-]+){1,5}))",
                text, flags=re.IGNORECASE
            )
            if m2:
                name = OutlookGraphAppOnly._clean_name(m2.group("name"))
        if len(name.split()) < 2:
            return ""
        return name

    # === TR tarih yakala ===
    @staticmethod
    def _extract_first_date(text):
        if not text:
            return None
        t = text.replace("\r\n", "\n").replace("\r", "\n")
        t = re.sub(r"\s+", " ", t)
        # bağlam öncelikli
        ctx_pat = re.compile(r"tarihi\s+itibari\s+ile.{0,30}", re.IGNORECASE)
        segs = []
        mctx = ctx_pat.search(t)
        if mctx:
            s, e = mctx.start(), mctx.end()
            segs.append(t[max(0, s-40): min(len(t), e+40)])
        segs.append(t)

        def _parse_candidate(s):
            m = re.search(r"\b(\d{1,2})[./](\d{1,2})[./](\d{2,4})\b", s)
            if m:
                d, M, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
                if y < 100:
                    y += 2000
                try:
                    return _dt.date(y, M, d)
                except ValueError:
                    pass
            m = re.search(r"\b(\d{4})-(\d{1,2})-(\d{1,2})\b", s)
            if m:
                y, M, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
                try:
                    return _dt.date(y, M, d)
                except ValueError:
                    pass
            return None

        for seg in segs:
            dd = _parse_candidate(seg)
            if dd:
                return dd
        return None

    # === TERMINATION extraction ===
    @staticmethod
    def _extract_name_termination(text):
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

    # === ACTION 1: New Hire -> İSİM LİSTESİ (aynı API) ===
    def list_new_hire_messages(self, tenant_id, client_id, client_secret, mailbox, top=None, mode="exact"):
        subject = "[Kurum Dışı] Şirkete Yeni Katılım - New Comer"
        items = self._fetch_messages(tenant_id, client_id, client_secret, mailbox, subject, mode=mode, max_items=top)
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
                    prev = (it.get("bodyPreview") or "")[:120]
                    self.logger.info(f"[NEW_HIRE] no match. preview={prev}")
        return {"success": True, "names": names}

    # === ACTION 2: Termination ===
    def list_termination_messages(self, tenant_id, client_id, client_secret, mailbox,
                                  top=None,
                                  mode="exact",
                                  subject_text="[Kurum Dışı] Çalışan İlişik Kesme Bildirimi",
                                  output="detailed",
                                  qradar_entry_type="ALN"):
        """
        mode: 'exact' (eski) | 'contains' (esnek)
        output:
          - 'detailed'         -> items + names (eski geri uyumluluk)
          - 'compact'          -> items (name, mail_received_at, termination_date, activate_date, activate_at)
          - 'names'            -> {names:[...]}
          - 'qradar_values'    -> {values:[...]}  (QRadar ref-set loader için düz liste)
          - 'qradar_values_typed' -> {values:[{"type": "<ALN>", "value": "Ad Soyad"}, ...]}
        """
        items = self._fetch_messages(tenant_id, client_id, client_secret, mailbox,
                                     subject_text, mode=mode, max_items=top)

        names = []
        out_items = []

        for it in items:
            body = self._get_body_text(it)
            name = self._extract_name_termination(body)
            term_date = self._extract_first_date(body)

            if name:
                names.append(name)

            received_iso = it.get("receivedDateTime", "") or ""
            activate_at = self._midnight_utc_after_days(term_date, 3) if term_date else ""
            activate_date = self._to_iso_date(term_date + _dt.timedelta(days=3)) if term_date else ""

            base = {
                "name": name or "",
                "mail_received_at": received_iso,
                "termination_date": self._to_iso_date(term_date) if term_date else "",
                "activate_date": activate_date,  # TR’de gün kıyas için
                "activate_at": activate_at,      # UTC midnight (isteğe bağlı)
            }

            if output in ("detailed", "compact"):
                pass  # base yeterli
            else:
                # diğer modlarda items toplamak zorunda değiliz ama tutmak sorun değil
                pass

            # detailed ek alanlar:
            if output == "detailed":
                base.update({
                    "subject": it.get("subject") or "",
                    "message_id": it.get("id") or "",
                    "preview": (it.get("bodyPreview") or "")[:300]
                })

            out_items.append(base)

        if self.logger:
            self.logger.info(f"[TERMINATION] parsed={len(out_items)} names={len([n for n in names if n])}")

        # --- ÇIKTI MODLARI ---
        if output == "names":
            return {"success": True, "names": names}

        if output == "qradar_values":
            # QRadar ref-set loader çoğu zaman {values:[...]} bekliyor
            return {"success": True, "values": names}

        if output == "qradar_values_typed":
            # Eğer loader entry type istiyorsa:
            typed = [{"type": qradar_entry_type, "value": n} for n in names if n]
            return {"success": True, "values": typed}

        if output == "compact":
            return {"success": True, "items": out_items}

        # default (geri uyumluluk): detailed + names
        return {"success": True, "items": out_items, "names": names}


if __name__ == "__main__":
    OutlookGraphAppOnly.run()
