from django.contrib import messages
from django.shortcuts import render

from .forms import ImportExcelForm
from .services.importer import importer_fichier_excel

from django.db.models import Avg, Max, Min, Count
from .models import (
    AnneeScolaire,
    Classe,
    Eleve,
    Resultat,
)


def importer_excel(request):

    if request.method == "POST":

        form = ImportExcelForm(
            request.POST,
            request.FILES
        )

        if form.is_valid():

            fichier = form.cleaned_data["fichier"]

            try:
                resultat = importer_fichier_excel(
                    fichier
                )

                messages.success(
                    request,
                    f"Importation réussie : "
                    f"{resultat['eleves']} élèves et "
                    f"{resultat['resultats']} résultats importés."
                )

            except Exception as e:

                messages.error(
                    request,
                    f"Erreur lors de l'importation : {e}"
                )

    else:
        form = ImportExcelForm()

    return render(
        request,
        "analyse/importer.html",
        {
            "form": form
        }
    )




def dashboard(request):

    annee_id = request.GET.get("annee")
    classe_id = request.GET.get("classe")
    trimestre = request.GET.get("trimestre")

    annees = AnneeScolaire.objects.all().order_by("-nom")
    classes = Classe.objects.all().order_by("nom")

    resultats = Resultat.objects.select_related(
        "eleve",
        "annee_scolaire",
        "classe"
    )

    # Filtres
    if annee_id:
        resultats = resultats.filter(
            annee_scolaire_id=annee_id
        )

    if classe_id:
        resultats = resultats.filter(
            classe_id=classe_id
        )

    if trimestre:
        resultats = resultats.filter(
            trimestre=trimestre
        )

    statistiques = resultats.aggregate(
        nombre_eleves=Count("id"),
        moyenne_classe=Avg("moyenne_trimestrielle"),
        meilleure_moyenne=Max("moyenne_trimestrielle"),
        plus_faible_moyenne=Min("moyenne_trimestrielle"),
    )

    # Classement
    classement = resultats.order_by(
        "-moyenne_trimestrielle"
    )

    context = {
        "annees": annees,
        "classes": classes,
        "resultats": resultats,
        "classement": classement,
        "statistiques": statistiques,

        "annee_selectionnee": annee_id,
        "classe_selectionnee": classe_id,
        "trimestre_selectionne": trimestre,
    }

    return render(
        request,
        "analyse/dashboard.html",
        context
    )