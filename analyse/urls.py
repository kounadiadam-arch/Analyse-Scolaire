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
        name="export_eleves_sous_10_excel"
    ),
    path(
        "statistiques/",
        views.statistiques,
        name="statistiques"
    ),
    path(
        "meilleurs-eleves/",
        views.meilleurs_eleves,
        name="meilleurs_eleves"
    ),
    path(
        "resultats/",
        views.resultats_eleves,
        name="resultats"
    ),
    path(
        "eleves-a-suivre/",
        views.eleves_a_suivre,
        name="eleves_a_suivre"
    ),

]