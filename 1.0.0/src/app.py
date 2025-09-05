import time
import requests
from typing import Optional, Dict, Any
from walkoff_app_sdk.app_base import AppBase

_TOKEN_CACHE = {"token": None, "exp": 0}

class OutlookGraphAppOnly(AppBase):
    __version__ = "1.0.0"
    app_name = "Outlook Graph (AppOnly)"

    def _get_token(self, tenant_id: str, client_id: str, client_secret: str) -> str:
        now = int(time.time())
        if _TOKEN_CACHE["token"] and _TOKEN_CACHE["exp"] - 60 > now:
            return _TOKEN_CACHE["token"]
        url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
        data = {
            "grant_type": "client_credentials",
            "client_id": client_id,
            "client_secret": client_secret,
            "scope": "https://graph.microsoft.com/.default",
        }
        headers = {"Content-Type": "application/x-www-form-urlencoded", "Accept": "application/json"}
        resp = requests.post(url, data=data, headers=headers, timeout=30)
        resp.raise_for_status()
        j = resp.json()
        token = j.get("access_token")
        expires_in = int(j.get("expires_in", 3000))
        if not token:
            raise RuntimeError(f"Token JSON missing access_token: {j}")
        _TOKEN_CACHE["token"] = token
        _TOKEN_CACHE["exp"] = now + expires_in
        return token

    def _auth_headers(self, token: str) -> Dict[str, str]:
        return {
            "Authorization": f"Bearer {token}",
            "Accept": "application/json",
            'Prefer': 'outlook.body-content-type="text"',
        }

    def list_inbox(
        self,
        tenant_id: str,
        client_id: str,
        client_secret: str,
        mailbox: str,
        top: int = 50,
        orderby: str = "receivedDateTime desc",
        select: str = "subject,from,receivedDateTime,bodyPreview,hasAttachments,webLink,id",
    ) -> Dict[str, Any]:
        token = self._get_token(tenant_id, client_id, client_secret)
        url = f"https://graph.microsoft.com/v1.0/users/{mailbox}/mailFolders/Inbox/messages"
        params = {"": top, "": orderby}
        if select:
            params[""] = select
        resp = requests.get(url, headers=self._auth_headers(token), params=params, timeout=60)
        resp.raise_for_status()
        data = resp.json()
        return {"success": True, "data": data, "next_link": data.get("@odata.nextLink")}

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
        return {"success": True, "data": data, "next_link": data.get("@odata.nextLink")}

    def list_inbox_delta(
        self,
        tenant_id: str,
        client_id: str,
        client_secret: str,
        mailbox: str,
        delta_link: Optional[str] = None,
        top: int = 100,
    ) -> Dict[str, Any]:
        token = self._get_token(tenant_id, client_id, client_secret)
        if delta_link:
            url = delta_link
            params = None
        else:
            url = f"https://graph.microsoft.com/v1.0/users/{mailbox}/mailFolders/Inbox/messages/delta"
            params = {"": top}
        resp = requests.get(url, headers=self._auth_headers(token), params=params, timeout=60)
        resp.raise_for_status()
        data = resp.json()
        return {
            "success": True,
            "data": data,
            "delta_link": data.get("@odata.deltaLink"),
            "next_link": data.get("@odata.nextLink"),
        }

if __name__ == "__main__":
    OutlookGraphAppOnly.run()
