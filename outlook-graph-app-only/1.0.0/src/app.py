import requests
from typing import Dict, Any, Optional
from walkoff_app_sdk.app_base import AppBase

GRAPH = "https://graph.microsoft.com/v1.0"


class OutlookGraphAppOnly(AppBase):
    __version__ = "1.0.0"
    app_name = "Outlook Graph (AppOnly)"

    def __init__(self, redis, logger, console_logger=None):
        super().__init__(redis, logger, console_logger)

    def _auth_headers(self, token: str) -> Dict[str, str]:
        return {
            "Authorization": f"Bearer {token}",
            "Accept": "application/json",
            'Prefer': 'outlook.body-content-type="text"',
        }

    def _get_token(self, tenant_id: str, client_id: str, client_secret: str) -> str:
        url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
        data = {
            "client_id": client_id,
            "client_secret": client_secret,
            "grant_type": "client_credentials",
            "scope": "https://graph.microsoft.com/.default",
        }
        resp = requests.post(url, data=data, timeout=30)
        resp.raise_for_status()
        return resp.json()["access_token"]

    def list_inbox(
        self,
        tenant_id: str,
        client_id: str,
        client_secret: str,
        mailbox: str,
        top: Optional[int] = 50,
    ) -> Dict[str, Any]:
        token = self._get_token(tenant_id, client_id, client_secret)
        url = f"{GRAPH}/users/{mailbox}/mailFolders/Inbox/messages"
        params: Dict[str, Any] = {}
        if top is not None:
            top = max(1, min(int(top), 1000))
            params[""] = top
        resp = requests.get(url, headers=self._auth_headers(token), params=params, timeout=60)
        resp.raise_for_status()
        data = resp.json()
        return {"success": True, "data": data.get("value", []), "next_link": data.get("@odata.nextLink")}

    def list_next_page(
        self,
        tenant_id: str,
        client_id: str,
        client_secret: str,
        next_link: str,
    ) -> Dict[str, Any]:
        if not next_link:
            return {"success": False, "error": "next_link is empty"}
        token = self._get_token(tenant_id, client_id, client_secret)
        resp = requests.get(next_link, headers=self._auth_headers(token), timeout=60)
        resp.raise_for_status()
        data = resp.json()
        return {"success": True, "data": data.get("value", []), "next_link": data.get("@odata.nextLink")}

    def list_inbox_delta(
        self,
        tenant_id: str,
        client_id: str,
        client_secret: str,
        mailbox: str,
        delta_link: Optional[str] = None,
        top: Optional[int] = 100,
    ) -> Dict[str, Any]:
        token = self._get_token(tenant_id, client_id, client_secret)
        if delta_link:
            url = delta_link
            params = None
        else:
            url = f"{GRAPH}/users/{mailbox}/mailFolders/Inbox/messages/delta"
            params = {}
            if top is not None:
                top = max(1, min(int(top), 1000))
                params[""] = top
        resp = requests.get(url, headers=self._auth_headers(token), params=params, timeout=60)
        resp.raise_for_status()
        data = resp.json()
        return {
            "success": True,
            "data": data.get("value", []),
            "delta_link": data.get("@odata.deltaLink"),
            "next_link": data.get("@odata.nextLink"),
        }


if __name__ == "__main__":
    OutlookGraphAppOnly.run()
