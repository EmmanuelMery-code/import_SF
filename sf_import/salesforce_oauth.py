"""
Authentification Salesforce par session ID (sans Connected App).
L'utilisateur se connecte à Salesforce dans le navigateur, puis transmet
son session ID à l'application.
"""
import re
import requests
from simple_salesforce import Salesforce


def _normalize_instance_url(url: str) -> str:
    """
    Convertit l'URL Lightning (lightning.force.com) en URL API (my.salesforce.com).
    Ex: https://xxx.trailblaze.lightning.force.com -> https://xxx.trailblaze.my.salesforce.com
    """
    if not url or ".lightning.force.com" not in url.lower():
        return url
    url = url.strip().rstrip("/")
    match = re.match(r"(https?://)([^/]+)(.*)", url)
    if match:
        prefix, host, rest = match.groups()
        host = host.replace(".lightning.force.com", ".my.salesforce.com")
        return f"{prefix}{host}{rest}" if rest else f"{prefix}{host}"
    return url


def create_salesforce_client(instance_url: str, session_id: str) -> Salesforce:
    """Crée un client simple_salesforce à partir de l'URL d'instance et du session ID."""
    return Salesforce(instance_url=instance_url, session_id=session_id)


def get_org_info(sf_client, instance_url: str = "", session_id: str = "") -> dict:
    """
    Récupère le nom, l'Id et l'URL de l'org Salesforce.
    Utilise d'abord l'endpoint userinfo (plus fiable avec Session ID),
    puis tente la requête Organization pour le nom.
    """
    url = instance_url or ""
    if not url and hasattr(sf_client, "sf_instance"):
        url = f"https://{sf_client.sf_instance}"
    base = url.rstrip("/") if url else ""
    if not session_id:
        session_id = getattr(sf_client, "session_id", None) or getattr(getattr(sf_client, "auth", None), "access_token", None) or ""
    org_id = ""
    org_name = ""

    # 1. Endpoint userinfo - Bearer ou OAuth
    for auth_header in (f"Bearer {session_id}", f"OAuth {session_id}"):
            try:
                userinfo_url = f"{base}/services/oauth2/userinfo"
                resp = requests.get(
                    userinfo_url,
                    headers={"Authorization": auth_header},
                    timeout=10,
                )
                if resp.status_code == 200:
                    data = resp.json()
                    org_id = data.get("organization_id", "") or data.get("organizationId", "") or ""
                    if org_id:
                        break
            except Exception:
                continue

    # 2. Requête SOQL Organization - essayer URL d'abord, puis URL normalisée si Lightning
    clients_to_try = [sf_client]
    if base and ".lightning.force.com" in base.lower():
        try:
            alt_url = _normalize_instance_url(base)
            if alt_url != base:
                clients_to_try.append(Salesforce(instance_url=alt_url, session_id=session_id))
        except Exception:
            pass
    for client in clients_to_try:
        try:
            result = client.query("SELECT Id, Name FROM Organization LIMIT 1")
            records = result.get("records") if isinstance(result, dict) else []
            if records:
                rec = records[0]
                if hasattr(rec, "get"):
                    name = rec.get("Name", "") or ""
                    oid = rec.get("Id", "") or ""
                else:
                    name = getattr(rec, "Name", "") or ""
                    oid = getattr(rec, "Id", "") or ""
                org_name = str(name) if name else ""
                if not org_id:
                    org_id = str(oid) if oid else ""
                break
        except Exception:
            continue

    # 3. Si on a org_id mais pas de nom, essayer GET /sobjects/Organization/{id}
    if org_id and not org_name:
        try:
            org_url = f"{base}/services/data/v50.0/sobjects/Organization/{org_id}?fields=Name"
            resp = requests.get(
                org_url,
                headers={
                    "Authorization": f"Bearer {session_id}",
                    "Content-Type": "application/json",
                },
                timeout=10,
            )
            if resp.status_code == 200:
                data = resp.json()
                org_name = data.get("Name", "") or ""
        except Exception:
            pass

    return {
        "name": org_name,
        "id": org_id,
        "url": url,
    }
