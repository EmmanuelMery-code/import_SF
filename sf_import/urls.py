from django.urls import path
from . import views

urlpatterns = [
    path("", views.index, name="index"),
    path("explorer/", views.explorer, name="explorer"),
    path("explorer/upload/", views.explorer_upload, name="explorer_upload"),
    path("upload/", views.upload, name="upload"),
    path("preview/", views.preview, name="preview"),
    path("sf-login/", views.sf_login, name="sf_login"),
    path("sf-login/submit/", views.sf_login_submit, name="sf_login_submit"),
    path("sf-callback/", views.sf_callback, name="sf_callback"),
    path("create-import-file/", views.create_import_file, name="create_import_file"),
    path("create-import-file/generate/", views.create_import_file_generate, name="create_import_file_generate"),
    path("sf-logout/", views.sf_logout, name="sf_logout"),
    path("run-import/", views.run_import, name="run_import"),
    path("save-config/", views.save_config, name="save_config"),
    path("add-sheet/", views.add_sheet, name="add_sheet"),
    path("export-config/", views.export_config, name="export_config"),
    path("reset-config/", views.reset_config, name="reset_config"),
    path("rename-sheet/", views.rename_sheet, name="rename_sheet"),
    path("delete-sheet/", views.delete_sheet, name="delete_sheet"),
    path("api/sobjects/", views.api_sobjects, name="api_sobjects"),
]
