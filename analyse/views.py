import re
import unicodedata
from decimal import Decimal

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


# =====================================================================
# STATISTIQUES PAR NIVEAU (tableau type "STAT_1")
# =====================================================================
#
# Ces fonctions construisent le tableau récapitulatif par niveau
# (effectifs classés / non classés par sexe, répartition par tranche
# de moyenne, moyenne par niveau) affiché sur le dashboard, en
# respectant les filtres déjà appliqués (année, classe, trimestre,
# niveau).
#
# Hypothèses retenues (à ajuster si besoin) :
#   - "Effectif classé"      = élèves ayant une moyenne_trimestrielle
#                               renseignée et différente de 0.
#   - "Effectif non classé"  = élèves sans moyenne_trimestrielle ou
#                               avec une moyenne_trimestrielle égale à 0.
#   - Les pourcentages des tranches de moyenne sont calculés sur la
#     base de l'effectif classé (T) du niveau concerné.
#   - "F" = genre 'F' (Féminin), "G" = genre 'M' (Masculin/Garçon).
#   - Le "1er cycle" regroupe les niveaux 6e/5e/4e/3e, le
#     "2nd cycle" regroupe tout le reste (2nde, 1ère, Tle...), en se
#     basant sur les valeurs réelles du champ Classe.niveau, quelle
#     que soit leur orthographe exacte.


def _normaliser(texte):
    """Retire les accents et met en minuscules pour faciliter les
    comparaisons de niveaux ("2ndeA", "2NDE A", "2nde-A", ...)."""

    texte = texte or ""
    texte = unicodedata.normalize("NFKD", texte)
    texte = texte.encode("ascii", "ignore").decode("ascii")
    return texte.lower().strip()


def _info_tri_niveau(niveau):
    """Détermine à quel cycle appartient un niveau et calcule une clé
    de tri, à partir de la valeur brute stockée dans Classe.niveau."""

    n = _normaliser(niveau)

    # --- Collège : 6e, 5e, 4e, 3e (ou "6ème", "6 eme", ...) ---
    college = re.match(r"^([3-6])\s*e", n)

    if college:
        chiffre = int(college.group(1))
        # Ordre décroissant : 6e, 5e, 4e, 3e
        return "1er", (0, -chiffre, "")

    # --- Lycée : Seconde / Première / Terminale ---
    prefixes_lycee = [
        ("seconde", 0),
        ("2nde", 0),
        ("2de", 0),
        ("2e", 0),
        ("premiere", 1),
        ("1ere", 1),
        ("1re", 1),
        ("1e", 1),
        ("terminale", 2),
        ("tle", 2),
        ("term", 2),
    ]

    for prefixe, rang in prefixes_lycee:
        if n.startswith(prefixe):
            section = n[len(prefixe):].strip().upper()
            return "2nd", (1, rang, section)

    # --- Niveau non reconnu : placé en fin de tableau ---
    return "2nd", (2, 0, n)


CHAMPS_NUMERIQUES_NIVEAU = (
    "nombre_classes",
    "classe_f", "classe_g", "classe_t",
    "non_classe_f", "non_classe_g", "non_classe_t",
    "sup10_f", "sup10_g", "sup10_t",
    "moy_f", "moy_g", "moy_t",
    "inf_f", "inf_g", "inf_t",
)


def _nouvelle_ligne_niveau(niveau):

    ligne = {"niveau": niveau, "is_total": False, "classes_ids": set()}

    for champ in CHAMPS_NUMERIQUES_NIVEAU:
        ligne[champ] = 0

    ligne["somme_moyennes"] = Decimal("0")
    ligne["nb_moyennes"] = 0

    return ligne


def _finaliser_ligne(ligne):
    """Calcule nombre_classes, pourcentages et moyenne finale d'une ligne."""

    if "classes_ids" in ligne:
        ligne["nombre_classes"] = len(ligne.pop("classes_ids"))

    total_classe = ligne["classe_t"] or 0

    ligne["sup10_pct"] = (
        round(ligne["sup10_t"] / total_classe * 100, 2)
        if total_classe else 0
    )
    ligne["moy_pct"] = (
        round(ligne["moy_t"] / total_classe * 100, 2)
        if total_classe else 0
    )
    ligne["inf_pct"] = (
        round(ligne["inf_t"] / total_classe * 100, 2)
        if total_classe else 0
    )

    ligne["moyenne_niveau"] = (
        round(float(ligne["somme_moyennes"] / ligne["nb_moyennes"]), 2)
        if ligne["nb_moyennes"] else None
    )

    return ligne


def _construire_lignes_par_niveau(resultats):
    """Parcourt les résultats filtrés et construit une ligne de
    statistiques par niveau de classe."""

    lignes_par_niveau = {}

    for resultat in resultats:

        niveau = resultat.classe.niveau or "Non renseigné"

        ligne = lignes_par_niveau.setdefault(
            niveau, _nouvelle_ligne_niveau(niveau)
        )

        ligne["classes_ids"].add(resultat.classe_id)

        genre = resultat.eleve.genre
        moyenne = resultat.moyenne_trimestrielle

        est_fille = genre == "F"
        est_garcon = genre == "M"

        if moyenne is None or moyenne == 0:

            ligne["non_classe_t"] += 1

            if est_fille:
                ligne["non_classe_f"] += 1
            elif est_garcon:
                ligne["non_classe_g"] += 1

        else:

            ligne["classe_t"] += 1

            if est_fille:
                ligne["classe_f"] += 1
            elif est_garcon:
                ligne["classe_g"] += 1

            ligne["somme_moyennes"] += moyenne
            ligne["nb_moyennes"] += 1

            if moyenne >= 10:
                ligne["sup10_t"] += 1
                if est_fille:
                    ligne["sup10_f"] += 1
                elif est_garcon:
                    ligne["sup10_g"] += 1

            elif moyenne >= Decimal("8.5"):
                ligne["moy_t"] += 1
                if est_fille:
                    ligne["moy_f"] += 1
                elif est_garcon:
                    ligne["moy_g"] += 1

            else:
                ligne["inf_t"] += 1
                if est_fille:
                    ligne["inf_f"] += 1
                elif est_garcon:
                    ligne["inf_g"] += 1

    return [_finaliser_ligne(l) for l in lignes_par_niveau.values()]


def _fusionner_lignes(lignes, libelle):
    """Additionne plusieurs lignes de niveau pour obtenir une ligne
    de total (1er cycle / 2nd cycle / général)."""

    fusion = _nouvelle_ligne_niveau(libelle)
    fusion.pop("classes_ids")
    fusion["nombre_classes"] = 0
    fusion["is_total"] = True

    for ligne in lignes:

        for champ in CHAMPS_NUMERIQUES_NIVEAU:
            fusion[champ] += ligne[champ]

        fusion["somme_moyennes"] += ligne["somme_moyennes"]
        fusion["nb_moyennes"] += ligne["nb_moyennes"]

    return _finaliser_ligne(fusion)


def calculer_tableau_statistiques_niveaux(resultats):
    """Construit la liste finale de lignes (dans l'ordre d'affichage)
    à utiliser directement dans le template : niveaux de 1er cycle,
    total 1er cycle, niveaux de 2nd cycle, total 2nd cycle, total
    général."""

    lignes = _construire_lignes_par_niveau(resultats)

    for ligne in lignes:
        cycle, cle_tri = _info_tri_niveau(ligne["niveau"])
        ligne["_cycle"] = cycle
        ligne["_cle_tri"] = cle_tri

    lignes_1er_cycle = sorted(
        (l for l in lignes if l["_cycle"] == "1er"),
        key=lambda l: l["_cle_tri"],
    )

    lignes_2nd_cycle = sorted(
        (l for l in lignes if l["_cycle"] == "2nd"),
        key=lambda l: l["_cle_tri"],
    )

    tableau = list(lignes_1er_cycle)

    if lignes_1er_cycle:
        tableau.append(
            _fusionner_lignes(lignes_1er_cycle, "Total 1er cycle")
        )

    tableau += lignes_2nd_cycle

    if lignes_2nd_cycle:
        tableau.append(
            _fusionner_lignes(lignes_2nd_cycle, "Total 2nd cycle")
        )

    if lignes_1er_cycle or lignes_2nd_cycle:
        tableau.append(
            _fusionner_lignes(lignes, "Total général")
        )

    return tableau


# =====================================================================
# TABLEAU EFFECTIFS / AFFECTÉS / REDOUBLANTS PAR NIVEAU (tableau "STAT_2")
# =====================================================================
#
# Hypothèses retenues (à ajuster si besoin) :
#   - "EFFECTIF"    = nombre d'élèves distincts (par sexe) ayant un
#                      résultat correspondant aux filtres appliqués,
#                      pour le niveau concerné.
#   - "AFFECTÉS"    = sous-ensemble de ces élèves dont le champ
#                      Eleve.statut contient "affecté" (comparaison
#                      insensible à la casse et aux accents).
#   - "REDOUBLANTS" = sous-ensemble de ces élèves dont
#                      Eleve.redoublant vaut True.
#   - "G" = genre 'M' (Garçon), "F" = genre 'F' (Fille).
#   - En plus du total 1er cycle / 2nd cycle, le 2nd cycle est
#     également sous-totalisé par palier (2nde, 1ère, Tle), comme
#     dans le modèle fourni.


CHAMPS_NUMERIQUES_EFFECTIFS = (
    "effectif_g", "effectif_f", "effectif_t",
    "affectes_g", "affectes_f", "affectes_t",
    "redoublants_g", "redoublants_f", "redoublants_t",
)


def _est_affecte(statut):
    return "affecte" in _normaliser(statut or "")


def _calculer_ligne_effectifs(niveau, eleves):

    ligne = {"niveau": niveau, "is_total": False}

    for champ in CHAMPS_NUMERIQUES_EFFECTIFS:
        ligne[champ] = 0

    for eleve in eleves:

        genre = eleve.genre
        est_garcon = genre == "M"
        est_fille = genre == "F"

        ligne["effectif_t"] += 1
        if est_garcon:
            ligne["effectif_g"] += 1
        elif est_fille:
            ligne["effectif_f"] += 1

        if _est_affecte(eleve.statut):
            ligne["affectes_t"] += 1
            if est_garcon:
                ligne["affectes_g"] += 1
            elif est_fille:
                ligne["affectes_f"] += 1

        if eleve.redoublant:
            ligne["redoublants_t"] += 1
            if est_garcon:
                ligne["redoublants_g"] += 1
            elif est_fille:
                ligne["redoublants_f"] += 1

    return ligne


def _construire_lignes_effectifs(resultats):
    """Construit, pour chaque niveau, l'ensemble des élèves distincts
    (par élève, pas par résultat) correspondant aux filtres, puis
    calcule les compteurs effectif / affectés / redoublants."""

    niveau_eleves = {}

    for resultat in resultats:

        niveau = resultat.classe.niveau or "Non renseigné"

        eleves_niveau = niveau_eleves.setdefault(niveau, {})
        eleves_niveau[resultat.eleve_id] = resultat.eleve

    return [
        _calculer_ligne_effectifs(niveau, eleves.values())
        for niveau, eleves in niveau_eleves.items()
    ]


def _fusionner_lignes_effectifs(lignes, libelle):

    fusion = {"niveau": libelle, "is_total": True}

    for champ in CHAMPS_NUMERIQUES_EFFECTIFS:
        fusion[champ] = sum(l[champ] for l in lignes)

    return fusion


NOMS_PALIERS_LYCEE = {
    0: "Total 2nde",
    1: "Total 1ere",
    2: "Total Tle",
}


def calculer_tableau_effectifs_niveaux(resultats):
    """Construit la liste finale de lignes (dans l'ordre d'affichage)
    pour le tableau effectifs / affectés / redoublants : niveaux de
    1er cycle, total 1er cycle, puis pour le 2nd cycle un
    sous-total par palier (2nde / 1ère / Tle), un total 2nd cycle,
    et enfin le total général."""

    lignes = _construire_lignes_effectifs(resultats)

    for ligne in lignes:
        cycle, cle_tri = _info_tri_niveau(ligne["niveau"])
        ligne["_cycle"] = cycle
        ligne["_cle_tri"] = cle_tri

    lignes_1er_cycle = sorted(
        (l for l in lignes if l["_cycle"] == "1er"),
        key=lambda l: l["_cle_tri"],
    )

    lignes_2nd_cycle = sorted(
        (l for l in lignes if l["_cycle"] == "2nd"),
        key=lambda l: l["_cle_tri"],
    )

    tableau = list(lignes_1er_cycle)

    if lignes_1er_cycle:
        tableau.append(
            _fusionner_lignes_effectifs(lignes_1er_cycle, "Total 1er Cycle")
        )

    paliers = {}
    niveaux_non_reconnus = []

    for ligne in lignes_2nd_cycle:
        if ligne["_cle_tri"][0] == 1:
            palier = ligne["_cle_tri"][1]
            paliers.setdefault(palier, []).append(ligne)
        else:
            niveaux_non_reconnus.append(ligne)

    for palier in sorted(paliers.keys()):

        lignes_palier = paliers[palier]
        tableau += lignes_palier

        tableau.append(
            _fusionner_lignes_effectifs(
                lignes_palier,
                NOMS_PALIERS_LYCEE.get(palier, "Total palier")
            )
        )

    tableau += niveaux_non_reconnus

    if lignes_2nd_cycle:
        tableau.append(
            _fusionner_lignes_effectifs(lignes_2nd_cycle, "Total 2e Cycle")
        )

    if lignes_1er_cycle or lignes_2nd_cycle:
        tableau.append(
            _fusionner_lignes_effectifs(lignes, "Total Général")
        )

    return tableau



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

    # Une moyenne trimestrielle égale à 0 correspond à un élève non
    # classé : elle ne doit donc ni être retenue comme "plus faible
    # moyenne", ni entrer dans le calcul de la moyenne de la classe.
    resultats_avec_moyenne = resultats.exclude(
        moyenne_trimestrielle=0
    )

    statistiques = resultats.aggregate(

        nombre_eleves=Count("id"),

    )

    statistiques.update(

        resultats_avec_moyenne.aggregate(

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

    )

    classement = resultats.order_by(
        "-moyenne_trimestrielle"
    )

    # -----------------------------
    # TABLEAU STATISTIQUES PAR NIVEAU
    # -----------------------------

    tableau_statistiques_niveaux = calculer_tableau_statistiques_niveaux(
        resultats
    )

    tableau_effectifs_niveaux = calculer_tableau_effectifs_niveaux(
        resultats
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

        "tableau_statistiques_niveaux": tableau_statistiques_niveaux,

        "tableau_effectifs_niveaux": tableau_effectifs_niveaux,

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