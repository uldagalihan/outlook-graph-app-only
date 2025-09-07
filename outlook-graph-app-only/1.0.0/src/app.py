import requests
from walkoff_app_sdk.app_base import AppBase

GRAPH = "https://graph.microsoft.com/v1.0"

class OutlookGraphAppOnly(AppBase):
    1\.0\.1 = "1\.0\.1"
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
        try:
            r = requests.post(url, data=data, timeout=30)
            r.raise_for_status()
            return r.json()["access_token"]
        except requests.RequestException as e:
            return {"error": True, "reason": f"Token error: {e}"}

    def list_inbox(self, tenant_id, client_id, client_secret, mailbox, top=None):
        tok = self._token(tenant_id, client_id, client_secret)
        if isinstance(tok, dict) and tok.get("error"):
            return {"success": False, "reason": tok["reason"]}

        params = {}
        if top:
            params["$top"] = int(top)

        url = f"{GRAPH}/users/{mailbox}/mailFolders/Inbox/messages"
        try:
            r = requests.get(url, headers={"Authorization": f"Bearer {tok}"}, params=params, timeout=30)
            r.raise_for_status()
            data = r.json()
            return {
                "success": True,
                "data": data.get("value", []),
                "next_link": data.get("@odata.nextLink")
            }
        except requests.RequestException as e:
            return {"success": False, "reason": f"list_inbox error: {e}"}

    def list_next_page(self, tenant_id, client_id, client_secret, next_link):
        tok = self._token(tenant_id, client_id, client_secret)
        if isinstance(tok, dict) and tok.get("error"):
            return {"success": False, "reason": tok["reason"]}

        try:
            r = requests.get(next_link, headers={"Authorization": f"Bearer {tok}"}, timeout=30)
            r.raise_for_status()
            data = r.json()
            return {
                "success": True,
                "data": data.get("value", []),
                "next_link": data.get("@odata.nextLink")
            }
        except requests.RequestException as e:
            return {"success": False, "reason": f"list_next_page error: {e}"}

    def list_inbox_delta(self, tenant_id, client_id, client_secret, mailbox, delta_link=None, top=None):
        tok = self._token(tenant_id, client_id, client_secret)
        if isinstance(tok, dict) and tok.get("error"):
            return {"success": False, "reason": tok["reason"]}

        if delta_link:
            url = delta_link
            params = None
        else:
            url = f"{GRAPH}/users/{mailbox}/mailFolders/Inbox/messages/delta"
            params = {}
            if top:
                params["$top"] = int(top)

        try:
            r = requests.get(url, headers={"Authorization": f"Bearer {tok}"}, params=params, timeout=30)
            r.raise_for_status()
            data = r.json()
            return {
                "success": True,
                "data": data.get("value", []),
                "delta_link": data.get("@odata.deltaLink"),
                "next_link": data.get("@odata.nextLink")
            }
        except requests.RequestException as e:
            return {"success": False, "reason": f"list_inbox_delta error: {e}"}

if __name__ == "__main__":
    from walkoff_app_sdk.runner import run_app
    run_app(OutlookGraphAppOnly)
