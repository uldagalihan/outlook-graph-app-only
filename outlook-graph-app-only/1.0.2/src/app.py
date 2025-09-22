import requests
from walkoff_app_sdk.app_base import AppBase

GRAPH = "https://graph.microsoft.com/v1.0"

class OutlookGraphAppOnly(AppBase):
    __version__ = "1.1.0"
    app_name = "Outlook Graph AppOnly"

    def __init__(self, redis=None, logger=None, **kwargs):
        super().__init__(redis=redis, logger=logger, **kwargs)

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

    def _get(self, url, tok, params=None, use_search=False):
        headers = {"Authorization": f"Bearer {tok}"}
        if use_search:
            headers["ConsistencyLevel"] = "eventual"
        r = requests.get(url, headers=headers, params=params, timeout=30)
        r.raise_for_status()
        return r.json()

    def _patch(self, url, tok, payload):
        headers = {
            "Authorization": f"Bearer {tok}",
            "Content-Type": "application/json"
        }
        r = requests.patch(url, headers=headers, json=payload, timeout=30)
        r.raise_for_status()
        return r.status_code in (200, 202)

    def list_unread_by_subject_inbox(self, tenant_id, client_id, client_secret, mailbox, subject_query, top=25, mark_read=False):
        tok = self._token(tenant_id, client_id, client_secret)
        url = f"{GRAPH}/users/{mailbox}/mailFolders/Inbox/messages"
        params = {
            "": subject_query,
            "": "isRead eq false",
            "": "receivedDateTime desc",
            "": int(top)
        }
        data = self._get(url, tok, params=params, use_search=True)
        items = data.get("value", [])

        if mark_read and items:
            for m in items:
                mid = m.get("id")
                if not mid:
                    continue
                patch_url = f"{GRAPH}/users/{mailbox}/messages/{mid}"
                try:
                    self._patch(patch_url, tok, {"isRead": True})
                except requests.HTTPError as e:
                    if self.logger:
                        self.logger.error(f"Mark read failed for {mid}: {e}")

        return {
            "success": True,
            "count": len(items),
            "data": items
        }

    def list_unread_new_hire(self, tenant_id, client_id, client_secret, mailbox, top=25, mark_read=False):
        subject_query = '"[Kurum Dışı] Şirkete Yeni Katılım - New Comer"'
        return self.list_unread_by_subject_inbox(
            tenant_id, client_id, client_secret, mailbox,
            subject_query=subject_query, top=top, mark_read=mark_read
        )

    def list_unread_termination(self, tenant_id, client_id, client_secret, mailbox, top=25, mark_read=False):
        subject_query = '"[Kurum Dışı] Çalışan İlişik Kesme Bildirimi"'
        return self.list_unread_by_subject_inbox(
            tenant_id, client_id, client_secret, mailbox,
            subject_query=subject_query, top=top, mark_read=mark_read
        )

if __name__ == "__main__":
    OutlookGraphAppOnly.run()
