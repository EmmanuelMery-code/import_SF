from django.apps import AppConfig


class SfImportConfig(AppConfig):
    default_auto_field = "django.db.models.BigAutoField"
    name = "sf_import"
    verbose_name = "Import Salesforce"
