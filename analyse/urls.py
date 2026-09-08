from django.urls import path

from . import views


urlpatterns = [
    path(
        "importer/",
        views.importer_excel,
        name="importer_excel"
    ),
    path(
        "dashboard/",
        views.dashboard,
        name="dashboard"
    ),
]