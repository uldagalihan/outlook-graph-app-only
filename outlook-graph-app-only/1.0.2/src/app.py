import requests
from walkoff_app_sdk.app_base import AppBase

GRAPH = "https://graph.microsoft.com/v1.0"

class OutlookGraphAppOnly(AppBase):
    __version__ = "1.0.2"
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

    def _get(self, url, tok, params=None):
        headers = {"Authorization": f"Bearer {tok}"}
        r = requests.get(url, headers=headers, params=params, timeout=30)
        r.raise_for_status()
        return r.json()

    def _list_by_exact_subject(self, tenant_id, client_id, client_secret, mailbox, subject, top=None):
        tok = self._token(tenant_id, client_id, client_secret)
        url = f"{GRAPH}/users/{mailbox}/messages"
        safe_subject = subject.replace("'", "''")
        filter_expr = f"receivedDateTime ge 1900-01-01T00:00:00Z and subject eq '{safe_subject}'"
        params = {
            "": "id,sender,subject,receivedDateTime",
            "": filter_expr,
            "": "receivedDateTime desc"
        }
        if top:
            params[""] = int(top)
        data = self._get(url, tok, params=params)
        return {"success": True, "data": data.get("value", []), "next_link": data.get("@odata.nextLink")}

    def list_new_hire_messages(self, tenant_id, client_id, client_secret, mailbox, top=None):
        subject = "[Kurum Dışı] Şirkete Yeni Katılım - New Comer"
        return self._list_by_exact_subject(tenant_id, client_id, client_secret, mailbox, subject, top)

    def list_termination_messages(self, tenant_id, client_id, client_secret, mailbox, top=None):
        subject = "[Kurum Dışı] Çalışan İlişik Kesme Bildirimi"
        return self._list_by_exact_subject(tenant_id, client_id, client_secret, mailbox, subject, top)

if __name__ == "__main__":
    OutlookGraphAppOnly.run()
