from django.urls import path

from . import views


urlpatterns = [
    path(
        "",
            views.dashboard,
        ),

    path(
        "dashboard/",
        views.dashboard,
        name="dashboard"
    ),
    
    path(
    "statistiques-matieres/",
    views.statistiques_matieres,
    name="statistiques_matieres"
    ),
    path(
        "export/excel/",
        views.exporter_statistiques_excel,
        name="export_statistiques_excel"
    ),
    path("export/eleves-sous-10/",
        views.exporter_eleves_sous_10_excel, 
        name="export_eleves_sous_10_excel"),
]