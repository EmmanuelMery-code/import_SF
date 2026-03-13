"""Context processors pour les templates."""

# Mapping url_name -> libellé pour le menu
PAGE_LABELS = {
    "index": "Accueil",
    "upload": "Accueil",
    "explorer": "Explorateur",
    "preview": "Prévisualisation",
    "create_import_file": "Créer un fichier pour import",
    "sf_login": "Connexion Salesforce",
    "sf_logout": "Déconnexion",
}


def salesforce_context(request):
    """Ajoute les infos Salesforce (org, URL) et connected au contexte."""
    url_name = getattr(getattr(request, "resolver_match", None), "url_name", None) or "index"
    current_label = PAGE_LABELS.get(url_name, "Accueil")

    creds = request.session.get("salesforce_credentials")
    if not creds:
        return {
            "connected": False,
            "sf_org_name": "",
            "sf_org_id": "",
            "sf_instance_url": "",
            "current_page_label": current_label,
            "current_url_name": url_name,
        }
    org_name = creds.get("org_name") or ""
    org_id = creds.get("org_id") or ""
    # Si org_name ou org_id manquant (ex: ancienne session), tenter de les récupérer
    if (not org_name or not org_id) and creds.get("instance_url") and creds.get("session_id"):
        try:
            from .salesforce_oauth import create_salesforce_client, get_org_info
            sf = create_salesforce_client(creds["instance_url"], creds["session_id"])
            info = get_org_info(sf, creds["instance_url"], session_id=creds.get("session_id", ""))
            if info.get("name") or info.get("id"):
                creds = dict(creds)
                creds["org_name"] = info.get("name", "")
                creds["org_id"] = info.get("id", "")
                request.session["salesforce_credentials"] = creds
                request.session.modified = True
                org_name = creds["org_name"]
                org_id = creds["org_id"]
        except Exception:
            pass
    return {
        "connected": True,
        "sf_org_name": org_name,
        "sf_org_id": org_id,
        "sf_instance_url": creds.get("instance_url") or "",
        "current_page_label": current_label,
        "current_url_name": url_name,
    }
