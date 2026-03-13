"""
Module d'import Salesforce avec réplication des IDs entre onglets.
"""

from typing import Callable, Optional

import pandas as pd

from excel_import_config import (
    ImportConfig,
    ColumnMapping,
    SheetConfig,
    apply_id_mapping,
)


def prepare_records_for_salesforce(
    df: pd.DataFrame,
    field_mapping: dict[str, str],
    exclude_columns: Optional[list[str]] = None
) -> list[dict]:
    """
    Prépare les enregistrements pour l'import Salesforce.
    field_mapping : {nom_colonne_excel: nom_champ_salesforce}
    """
    exclude = set(exclude_columns or [])
    records = []
    for _, row in df.iterrows():
        rec = {}
        for excel_col, sf_field in field_mapping.items():
            if excel_col in df.columns and sf_field and excel_col not in exclude:
                val = row.get(excel_col)
                if pd.notna(val) and str(val).strip():
                    rec[sf_field] = str(val).strip() if not isinstance(val, (int, float)) else val
        records.append(rec)
    return records


def import_to_salesforce(
    sf_client,
    object_name: str,
    records: list[dict],
    batch_size: int = 200,
    on_progress: Optional[Callable[[int, int, str], None]] = None,
    mode: str = "insert",
    upsert_external_id_field: Optional[str] = None,
) -> list[dict]:
    """
    Importe les enregistrements dans Salesforce via l'API REST.
    Modes: insert (create), update (nécessite Id), upsert (nécessite champ externe ID).
    Retourne la liste des résultats avec success, id, errors.
    """
    results = []
    total = len(records)
    sf_object = getattr(sf_client, object_name)

    for i, record in enumerate(records):
        if on_progress and (i + 1) % 50 == 0:
            on_progress(i + 1, total, f"Import ligne {i + 1}/{total}...")
        try:
            if mode == "update":
                rec_id = record.get("Id")
                if not rec_id:
                    results.append({"success": False, "errors": ["Id manquant pour update"], "id": None})
                    continue
                update_data = {k: v for k, v in record.items() if k != "Id"}
                if not update_data:
                    results.append({"success": True, "id": rec_id, "errors": []})
                else:
                    sf_object.update(str(rec_id), update_data)
                    results.append({"success": True, "id": rec_id, "errors": []})
            elif mode == "upsert" and upsert_external_id_field:
                ext_val = record.get(upsert_external_id_field)
                if ext_val is None or str(ext_val).strip() == "":
                    results.append({
                        "success": False,
                        "errors": [f"Valeur manquante pour le champ {upsert_external_id_field}"],
                        "id": None,
                    })
                    continue
                ext_val = str(ext_val).strip()
                upsert_result = sf_object.upsert(f"{upsert_external_id_field}/{ext_val}", record)
                if isinstance(upsert_result, dict) and upsert_result.get("created") is False:
                    results.append({
                        "success": True,
                        "id": upsert_result.get("id"),
                        "errors": [],
                    })
                elif isinstance(upsert_result, bool) and upsert_result:
                    results.append({"success": True, "id": None, "errors": []})
                else:
                    results.append({"success": True, "id": str(upsert_result) if upsert_result else None, "errors": []})
            else:
                create_result = sf_object.create(record)
                if isinstance(create_result, str):
                    results.append({"success": True, "id": create_result, "errors": []})
                elif isinstance(create_result, dict):
                    results.append({
                        "success": create_result.get("success", True),
                        "id": create_result.get("id"),
                        "errors": create_result.get("errors", []),
                    })
                else:
                    results.append({"success": True, "id": str(create_result), "errors": []})
        except Exception as e:
            results.append({"success": False, "errors": [str(e)], "id": None})
    if on_progress:
        on_progress(total, total, "Terminé")
    return results


def replicate_ids_to_next_sheets(
    data: dict[str, pd.DataFrame],
    config: ImportConfig,
    results_by_sheet: dict[str, list]
) -> dict[str, pd.DataFrame]:
    """
    Réplique les IDs Salesforce des onglets importés vers les onglets cibles
    en utilisant les correspondances de colonnes.
    """
    updated_data = {name: df.copy() for name, df in data.items()}
    sheets_ordered = config.get_sheets_ordered()
    
    for i, sheet_config in enumerate(sheets_ordered):
        sheet_name = sheet_config.name
        if sheet_name not in data or sheet_name not in results_by_sheet:
            continue
        
        df = updated_data[sheet_name]
        results = results_by_sheet[sheet_config.name]
        
        # Ajouter les IDs au DataFrame source
        id_col = sheet_config.id_column or "Id"
        if id_col not in df.columns:
            df[id_col] = None
        
        # Mapper les résultats (index ligne -> Id)
        for j, res in enumerate(results):
            if res.get("success") and res.get("id") and j < len(df):
                df.iloc[j, df.columns.get_loc(id_col)] = res["id"]
        
        updated_data[sheet_name] = df
        
        # Appliquer les correspondances vers les onglets cibles
        for mapping in config.mappings:
            if mapping.source_sheet != sheet_name:
                continue
            
            target_df = updated_data.get(mapping.target_sheet)
            if target_df is None:
                continue
            
            # target_column = clé de correspondance dans la cible
            # target_salesforce_field = colonne où mettre l'ID (ex: AccountId)
            target_id_col = mapping.target_salesforce_field or f"{mapping.source_sheet}Id"
            target_df = apply_id_mapping(
                source_df=df,
                source_key_col=mapping.source_column,
                source_id_col=id_col,
                target_df=target_df,
                target_key_col=mapping.target_column,
                target_id_col=target_id_col
            )
            updated_data[mapping.target_sheet] = target_df
    
    return updated_data
