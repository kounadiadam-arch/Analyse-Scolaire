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



def dashboard(request):

    # -----------------------------
    # IMPORTATION EXCEL
    # -----------------------------

    if request.method == "POST":

        form = ImportExcelForm(
            request.POST,
            request.FILES
        )

        if form.is_valid():

            fichier = form.cleaned_data["fichier"]

            try:

                resultat_import = importer_fichier_excel(
                    fichier
                )

                messages.success(
                    request,
                    f"Importation réussie ! "
                    f"{resultat_import['eleves']} élèves et "
                    f"{resultat_import['resultats']} résultats "
                    f"ont été importés."
                )

            except Exception as e:

                messages.error(
                    request,
                    f"Échec de l'importation : {str(e)}"
                )

        else:

            messages.error(
                request,
                "Échec de l'importation. "
                "Veuillez sélectionner un fichier Excel valide."
            )

    else:

        form = ImportExcelForm()


    # -----------------------------
    # FILTRES DU DASHBOARD
    # -----------------------------

    annee_id = request.GET.get("annee")
    classe_id = request.GET.get("classe")
    trimestre = request.GET.get("trimestre")
    niveau = request.GET.get("niveau")  # Nouveau paramètre pour le niveau

    annees = AnneeScolaire.objects.all().order_by("-nom")
    classes = Classe.objects.all().order_by("nom")
    
    # Récupération des niveaux uniques existants dans la base de données
    niveaux_uniques = Classe.objects.values_list('niveau', flat=True).distinct().exclude(niveau='').order_by('niveau')

    resultats = Resultat.objects.select_related(
        "eleve",
        "annee_scolaire",
        "classe"
    )

    # Filtre par niveau
    if niveau:
        resultats = resultats.filter(
            classe__niveau=niveau
        )

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

    # Si un niveau est sélectionné, on filtre aussi la liste des classes
    if niveau:
        classes = classes.filter(niveau=niveau)

    # -----------------------------
    # STATISTIQUES
    # -----------------------------

    statistiques = resultats.aggregate(

        nombre_eleves=Count("id"),

        moyenne_classe=Avg(
            "moyenne_trimestrielle"
        ),

        meilleure_moyenne=Max(
            "moyenne_trimestrielle"
        ),

        plus_faible_moyenne=Min(
            "moyenne_trimestrielle"
        ),

    )

    classement = resultats.order_by(
        "-moyenne_trimestrielle"
    )

    context = {

        "form": form,

        "annees": annees,

        "classes": classes,

        "niveaux": niveaux_uniques,  # Ajout des niveaux dans le contexte

        "resultats": resultats,

        "classement": classement,

        "statistiques": statistiques,

        "annee_selectionnee": annee_id,

        "classe_selectionnee": classe_id,

        "trimestre_selectionne": trimestre,

        "niveau_selectionne": niveau,  # Ajout du niveau sélectionné

    }


    return render(
        request,
        "analyse/dashboard.html",
        context
    )


def statistiques_matieres(request):

    annee_id = request.GET.get("annee")
    classe_id = request.GET.get("classe")
    trimestre = request.GET.get("trimestre")
    niveau = request.GET.get("niveau")  # Nouveau paramètre

    annees = AnneeScolaire.objects.all().order_by("-nom")
    classes = Classe.objects.all().order_by("nom")
    
    # Récupération des niveaux uniques existants
    niveaux_uniques = Classe.objects.values_list('niveau', flat=True).distinct().exclude(niveau='').order_by('niveau')

    resultats = Resultat.objects.select_related(
        "eleve",
        "annee_scolaire",
        "classe"
    )

    # Filtres
    if niveau:
        resultats = resultats.filter(
            classe__niveau=niveau
        )

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

    # Si un niveau est sélectionné, on filtre aussi la liste des classes
    if niveau:
        classes = classes.filter(niveau=niveau)

    matieres = {
        "Composition française": "composition_francaise",
        "Orthographe / Grammaire": "orthographe_grammaire",
        "Expression orale": "expression_orale",
        "Histoire-Géographie": "hist_geo",
        "Espagnol": "espagnol",
        "Allemand": "allemand",
        "Anglais": "anglais",
        "Français": "francais",
        "Conduite": "conduite",
        "EDHC": "edhc",
        "EPS": "eps",
        "SVT": "svt",
        "Physiques-Chimie": "physiques_chimie",
        "Mathématiques": "mathematiques",
        "Philosophie": "philosophie",
    }

    statistiques_matieres = []

    for nom, champ in matieres.items():

        valeurs = resultats.exclude(
            **{f"{champ}__isnull": True}
        )

        moyenne = valeurs.aggregate(
            moyenne=Avg(champ)
        )["moyenne"]

        nombre = valeurs.count()

        if moyenne is not None:
            moyenne = round(float(moyenne), 2)

            taux_reussite = (
                valeurs.filter(
                    **{f"{champ}__gte": 10}
                ).count()
                / nombre
                * 100
            )

            taux_reussite = round(
                taux_reussite,
                2
            )

        else:
            taux_reussite = 0

        statistiques_matieres.append({
            "nom": nom,
            "champ": champ,
            "moyenne": moyenne,
            "nombre": nombre,
            "taux_reussite": taux_reussite,
        })

    # Matières avec des données
    matieres_valides = [
        m for m in statistiques_matieres
        if m["moyenne"] is not None
    ]

    if matieres_valides:

        meilleure_matiere = max(
            matieres_valides,
            key=lambda x: x["moyenne"]
        )

        plus_faible_matiere = min(
            matieres_valides,
            key=lambda x: x["moyenne"]
        )

    else:

        meilleure_matiere = None
        plus_faible_matiere = None

    context = {
        "annees": annees,
        "classes": classes,
        "niveaux": niveaux_uniques,  # Ajout des niveaux

        "statistiques_matieres":
            statistiques_matieres,

        "meilleure_matiere":
            meilleure_matiere,

        "plus_faible_matiere":
            plus_faible_matiere,

        "annee_selectionnee": annee_id,
        "classe_selectionnee": classe_id,
        "trimestre_selectionne": trimestre,
        "niveau_selectionne": niveau,  # Ajout du niveau sélectionné
    }

    return render(
        request,
        "analyse/statistiques_matieres.html",
        context
    )