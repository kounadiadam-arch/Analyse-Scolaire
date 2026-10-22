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
    path("fiche-eleve/", 
        views.fiche_eleve,
        name="fiche_eleve"
    ),
    path("fiche-eleve/pdf/", 
        views.exporter_fiche_eleve_pdf,
        name="export_fiche_eleve_pdf"
    ),
    path("bilan-annuel/",
        views.bilan_annuel, 
        name="bilan_annuel"
    ),
    path("bilan-annuel/excel/",
        views.exporter_bilan_annuel_excel, 
        name="export_bilan_annuel_excel"
    ),
    path("bilan-annuel/pdf/", 
        views.exporter_bilan_annuel_pdf,
        name="export_bilan_annuel_pdf"
    ),
    path("bilan-trimestre/",
        views.bilan_trimestre,
        name="bilan_trimestre"
        ),
    path("bilan-trimestre/excel/",
        views.exporter_bilan_trimestre_excel,
        name="export_bilan_trimestre_excel"
        ),
    path("bilan-trimestre/pdf/",
        views.exporter_bilan_trimestre_pdf,
        name="export_bilan_trimestre_pdf"
        )

]