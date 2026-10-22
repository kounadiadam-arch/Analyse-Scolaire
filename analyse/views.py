import io
import re
import unicodedata
from datetime import datetime
from decimal import ROUND_HALF_UP, Decimal

from django.contrib import messages
from django.core.paginator import Paginator
from django.http import Http404, HttpResponse
from django.shortcuts import get_object_or_404, redirect
from django.shortcuts import render as _render_django
from django.urls import reverse

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

from .forms import ImportExcelForm
from .services.importer import importer_fichier_excel

from django.db.models import Avg, Max, Min, Count, Q
from .models import (
    AnneeScolaire,
    Classe,
    Eleve,
    Resultat,
)


def render(request, template_name, contexte=None, *args, **kwargs):
    """Comme django.shortcuts.render, mais ajoute à chaque page la
    variable « bilan_disponible » : le menu « Bilan de fin d'année » de
    base.html n'est cliquable que si les moyennes des trois trimestres
    sont enregistrées pour au moins une année."""

    contexte = dict(contexte or {})
    contexte.setdefault("bilan_disponible", bool(_ids_annees_completes()))

    return _render_django(request, template_name, contexte, *args, **kwargs)




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



# =====================================================================
# STATISTIQUES PAR MATIÈRE
# =====================================================================

SEUIL_REUSSITE = 10

MATIERES = {
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


def calculer_statistiques_matieres(resultats):
    """Moyenne, effectif et taux de réussite (note >= 10) de chaque
    matière. Une seule requête SQL pour toutes les matières."""

    agregats = {}

    for champ in MATIERES.values():
        agregats[f"moy_{champ}"] = Avg(champ)
        agregats[f"nb_{champ}"] = Count(champ)
        agregats[f"ok_{champ}"] = Count(
            champ,
            filter=Q(**{f"{champ}__gte": SEUIL_REUSSITE})
        )

    totaux = resultats.aggregate(**agregats)

    statistiques = []

    for nom, champ in MATIERES.items():

        moyenne = totaux[f"moy_{champ}"]
        nombre = totaux[f"nb_{champ}"]
        reussis = totaux[f"ok_{champ}"]

        if moyenne is not None and nombre:
            moyenne = round(float(moyenne), 2)
            taux_reussite = round(reussis / nombre * 100, 2)
        else:
            moyenne = None
            taux_reussite = 0

        statistiques.append({
            "nom": nom,
            "champ": champ,
            "moyenne": moyenne,
            "nombre": nombre,
            "taux_reussite": taux_reussite,
        })

    return statistiques


# =====================================================================
# DONNÉES DES GRAPHIQUES DU DASHBOARD
# =====================================================================
#
# Trois graphiques sont affichés : répartition des élèves par tranche
# de moyenne, moyenne par niveau, et filles / garçons par tranche.
# Ils réutilisent le tableau "Statistiques par niveau" déjà calculé
# (aucune requête SQL supplémentaire). Les données sont transmises au
# template dans un dictionnaire sérialisé en JSON (filtre
# `json_script`), que Chart.js lit côté navigateur.


def construire_donnees_graphiques(tableau_niveaux):

    lignes_niveaux = [l for l in tableau_niveaux if not l["is_total"]]

    total = next(
        (l for l in tableau_niveaux if l["niveau"] == "Total général"),
        None
    )

    if not lignes_niveaux or total is None:
        return {"a_des_donnees": False}

    labels_tranches = [
        "≥ 10",
        "8,5 à < 10",
        "< 8,5",
        "Non classés",
    ]

    # ---- Répartition des élèves par tranche de moyenne ----

    tranches = {
        "labels": labels_tranches,
        "valeurs": [
            total["sup10_t"],
            total["moy_t"],
            total["inf_t"],
            total["non_classe_t"],
        ],
        "taux_reussite": (
            total["sup10_pct"] if total["classe_t"] else None
        ),
    }

    # ---- Moyenne par niveau ----

    niveaux = {
        "labels": [l["niveau"] for l in lignes_niveaux],
        "moyennes": [l["moyenne_niveau"] for l in lignes_niveaux],
    }

    # ---- Filles / garçons par tranche de moyenne ----

    genre = {
        "labels": labels_tranches,
        "filles": [
            total["sup10_f"], total["moy_f"],
            total["inf_f"], total["non_classe_f"],
        ],
        "garcons": [
            total["sup10_g"], total["moy_g"],
            total["inf_g"], total["non_classe_g"],
        ],
    }

    return {
        "a_des_donnees": True,
        "seuil": SEUIL_REUSSITE,
        "tranches": tranches,
        "niveaux": niveaux,
        "genre": genre,
    }


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
    niveau = request.GET.get("niveau")
    recherche = request.GET.get("recherche", "").strip()

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

    # Recherche par nom, prénom ou matricule
    if recherche:
        resultats = resultats.filter(
            Q(eleve__nom__icontains=recherche)
            | Q(eleve__prenoms__icontains=recherche)
            | Q(eleve__matricule__icontains=recherche)
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

    # -----------------------------
    # CLASSEMENT + PERFORMANCE
    # -----------------------------

    classement_queryset = resultats.order_by(
        "-moyenne_trimestrielle",
        "eleve__nom",
        "eleve__prenoms",
    )

    # On prépare le classement complet avant la pagination afin que le
    # rang reste correct sur toutes les pages.
    classement_liste = list(classement_queryset)

    for rang, resultat in enumerate(classement_liste, start=1):
        resultat.rang = rang

        moyenne = resultat.moyenne_trimestrielle

        if moyenne is None:
            resultat.performance = "À accompagner"
            resultat.performance_css = "a-accompagner"
        elif moyenne >= 16:
            resultat.performance = "Excellent"
            resultat.performance_css = "excellent"
        elif moyenne >= 10:
            resultat.performance = "Satisfaisant"
            resultat.performance_css = "satisfaisant"
        else:
            resultat.performance = "À accompagner"
            resultat.performance_css = "a-accompagner"

        _attacher_details_matieres(resultat)

    total_resultats = len(classement_liste)

    # -----------------------------
    # PAGINATION DU CLASSEMENT (10 élèves par page)
    # -----------------------------

    paginator_classement = Paginator(
        classement_liste,
        10
    )

    numero_page_classement = request.GET.get("page_classement")

    classement = paginator_classement.get_page(
        numero_page_classement
    )

    # Évolution par rapport au trimestre précédent (page affichée seulement)
    _attacher_evolution(classement)

    # Chaîne de requête (filtres actuels) sans le paramètre de page,
    # pour reconstruire les liens de pagination dans le template.
    querydict_classement = request.GET.copy()
    querydict_classement.pop("page_classement", None)
    querystring_classement = querydict_classement.urlencode()

    # -----------------------------
    # TABLEAU STATISTIQUES PAR NIVEAU
    # -----------------------------

    tableau_statistiques_niveaux = calculer_tableau_statistiques_niveaux(
        resultats
    )

    # -----------------------------
    # LIBELLÉS DES FILTRES ACTIFS
    # -----------------------------

    filtres_actifs = []

    annee_obj = None
    if annee_id:
        annee_obj = AnneeScolaire.objects.filter(id=annee_id).first()
        if annee_obj:
            filtres_actifs.append({
                "libelle": "Année scolaire",
                "valeur": annee_obj.nom,
            })

    if niveau:
        filtres_actifs.append({
            "libelle": "Niveau",
            "valeur": niveau,
        })

    classe_obj = None
    if classe_id:
        classe_obj = Classe.objects.filter(id=classe_id).first()
        if classe_obj:
            filtres_actifs.append({
                "libelle": "Classe",
                "valeur": classe_obj.nom,
            })

    trimestre_labels = {
        "T1": "Trimestre 1",
        "T2": "Trimestre 2",
        "T3": "Trimestre 3",
    }

    if trimestre:
        filtres_actifs.append({
            "libelle": "Trimestre",
            "valeur": trimestre_labels.get(trimestre, trimestre),
        })

    if recherche:
        filtres_actifs.append({
            "libelle": "Recherche",
            "valeur": recherche,
        })

    # Période clairement affichée dans le dashboard
    periode_annee = annee_obj.nom if annee_obj else "Toutes les années"
    periode_trimestre = trimestre_labels.get(
        trimestre,
        "Tous les trimestres"
    )

    if annee_obj and trimestre:
        periode_analysee = f"{periode_annee} — {periode_trimestre}"
    elif annee_obj:
        periode_analysee = f"{periode_annee} — tous les trimestres"
    elif trimestre:
        periode_analysee = f"Toutes les années — {periode_trimestre}"
    else:
        periode_analysee = "Toutes les années — tous les trimestres"

    # -----------------------------
    # ZONE D'ALERTE ET D'ACCOMPAGNEMENT
    # -----------------------------

    # Élèves dont la moyenne est inférieure à 10, du plus faible au plus élevé.
    eleves_sous_10 = list(
        resultats
        .filter(moyenne_trimestrielle__lt=10)
        .order_by("moyenne_trimestrielle", "eleve__nom", "eleve__prenoms")
    )
    
    # Matières ayant les moyennes les plus faibles.
    statistiques_matieres_dashboard = calculer_statistiques_matieres(resultats)
    matieres_faibles = sorted(
        [m for m in statistiques_matieres_dashboard if m["moyenne"] is not None],
        key=lambda m: m["moyenne"]
    )[:5]

    # Classes ayant les plus faibles taux de réussite.
    classes_stats = {}
    for resultat in resultats:
        classe_nom = resultat.classe.nom
        stats_classe = classes_stats.setdefault(
            resultat.classe_id,
            {
                "classe": classe_nom,
                "total": 0,
                "reussis": 0,
            }
        )

        moyenne = resultat.moyenne_trimestrielle
        if moyenne is not None and moyenne > 0:
            stats_classe["total"] += 1
            if moyenne >= SEUIL_REUSSITE:
                stats_classe["reussis"] += 1

    classes_faibles = []
    for stats_classe in classes_stats.values():
        if stats_classe["total"]:
            stats_classe["taux_reussite"] = round(
                stats_classe["reussis"] / stats_classe["total"] * 100,
                2
            )
            classes_faibles.append(stats_classe)

    classes_faibles.sort(
        key=lambda c: (c["taux_reussite"], c["classe"])
    )
    classes_faibles = classes_faibles[:10]

    # Liste prioritaire : les 10 élèves ayant les moyennes les plus faibles.
    eleves_a_accompagner = eleves_sous_10[:10]

    # -----------------------------
    # DONNÉES DES GRAPHIQUES
    # -----------------------------

    donnees_graphiques = construire_donnees_graphiques(
        tableau_statistiques_niveaux
    )

    context = {

        "form": form,

        "annees": annees,

        "classes": classes,

        "niveaux": niveaux_uniques,  # Ajout des niveaux dans le contexte

        "resultats": resultats,

        "classement": classement,

        "total_resultats": total_resultats,

        "querystring_classement": querystring_classement,

        "statistiques": statistiques,

        "annee_selectionnee": annee_id,

        "classe_selectionnee": classe_id,

        "trimestre_selectionne": trimestre,

        "niveau_selectionne": niveau,
        "recherche": recherche,
        "filtres_actifs": filtres_actifs,
        "periode_analysee": periode_analysee,

        "tableau_statistiques_niveaux": tableau_statistiques_niveaux,

        "donnees_graphiques": donnees_graphiques,

        "eleves_sous_10": eleves_sous_10,
        "matieres_faibles": matieres_faibles,
        "classes_faibles": classes_faibles,
        "eleves_a_accompagner": eleves_a_accompagner,

    }


    return render(
        request,
        "analyse/dashboard.html",
        context
    )


def statistiques(request):
    """Page dédiée aux statistiques par niveau : effectifs / affectés /
    redoublants, ainsi que le tableau de statistiques par niveau,
    en respectant les mêmes filtres que le dashboard (année, niveau,
    classe, trimestre)."""

    # -----------------------------
    # FILTRES (identiques à la vue dashboard)
    # -----------------------------

    annee_id = request.GET.get("annee")
    classe_id = request.GET.get("classe")
    trimestre = request.GET.get("trimestre")
    niveau = request.GET.get("niveau")

    annees = AnneeScolaire.objects.all().order_by("-nom")
    classes = Classe.objects.all().order_by("nom")

    niveaux_uniques = Classe.objects.values_list(
        'niveau', flat=True
    ).distinct().exclude(niveau='').order_by('niveau')

    resultats = Resultat.objects.select_related(
        "eleve",
        "annee_scolaire",
        "classe"
    )

    if niveau:
        resultats = resultats.filter(classe__niveau=niveau)

    if annee_id:
        resultats = resultats.filter(annee_scolaire_id=annee_id)

    if classe_id:
        resultats = resultats.filter(classe_id=classe_id)

    if trimestre:
        resultats = resultats.filter(trimestre=trimestre)

    # Si un niveau est sélectionné, on filtre aussi la liste des classes
    if niveau:
        classes = classes.filter(niveau=niveau)

    # -----------------------------
    # TABLEAUX
    # -----------------------------

    tableau_statistiques_niveaux = calculer_tableau_statistiques_niveaux(
        resultats
    )

    tableau_effectifs_niveaux = calculer_tableau_effectifs_niveaux(
        resultats
    )

    # -----------------------------
    # LIBELLÉS DES FILTRES ACTIFS + PÉRIODE
    # -----------------------------

    filtres_actifs = []

    annee_obj = None
    if annee_id:
        annee_obj = AnneeScolaire.objects.filter(id=annee_id).first()
        if annee_obj:
            filtres_actifs.append({
                "libelle": "Année scolaire",
                "valeur": annee_obj.nom,
            })

    if niveau:
        filtres_actifs.append({
            "libelle": "Niveau",
            "valeur": niveau,
        })

    classe_obj = None
    if classe_id:
        classe_obj = Classe.objects.filter(id=classe_id).first()
        if classe_obj:
            filtres_actifs.append({
                "libelle": "Classe",
                "valeur": classe_obj.nom,
            })

    trimestre_labels = {
        "T1": "Trimestre 1",
        "T2": "Trimestre 2",
        "T3": "Trimestre 3",
    }

    if trimestre:
        filtres_actifs.append({
            "libelle": "Trimestre",
            "valeur": trimestre_labels.get(trimestre, trimestre),
        })

    periode_annee = annee_obj.nom if annee_obj else "Toutes les années"
    periode_trimestre = trimestre_labels.get(
        trimestre,
        "Tous les trimestres"
    )

    if annee_obj and trimestre:
        periode_analysee = f"{periode_annee} — {periode_trimestre}"
    elif annee_obj:
        periode_analysee = f"{periode_annee} — tous les trimestres"
    elif trimestre:
        periode_analysee = f"Toutes les années — {periode_trimestre}"
    else:
        periode_analysee = "Toutes les années — tous les trimestres"

    context = {

        "annees": annees,
        "classes": classes,
        "niveaux": niveaux_uniques,

        "annee_selectionnee": annee_id,
        "classe_selectionnee": classe_id,
        "trimestre_selectionne": trimestre,
        "niveau_selectionne": niveau,

        "filtres_actifs": filtres_actifs,
        "periode_analysee": periode_analysee,

        "tableau_statistiques_niveaux": tableau_statistiques_niveaux,
        "tableau_effectifs_niveaux": tableau_effectifs_niveaux,

    }

    return render(
        request,
        "analyse/statistiques.html",
        context
    )


def meilleurs_eleves(request):
    """Page dédiée affichant les 5 meilleurs élèves (moyenne la plus
    élevée), en respectant les mêmes filtres que le dashboard (année,
    niveau, classe, trimestre)."""

    # -----------------------------
    # FILTRES (identiques à la vue dashboard)
    # -----------------------------

    annee_id = request.GET.get("annee")
    classe_id = request.GET.get("classe")
    trimestre = request.GET.get("trimestre")
    niveau = request.GET.get("niveau")

    annees = AnneeScolaire.objects.all().order_by("-nom")
    classes = Classe.objects.all().order_by("nom")

    niveaux_uniques = Classe.objects.values_list(
        'niveau', flat=True
    ).distinct().exclude(niveau='').order_by('niveau')

    resultats = Resultat.objects.select_related(
        "eleve",
        "annee_scolaire",
        "classe"
    )

    if niveau:
        resultats = resultats.filter(classe__niveau=niveau)

    if annee_id:
        resultats = resultats.filter(annee_scolaire_id=annee_id)

    if classe_id:
        resultats = resultats.filter(classe_id=classe_id)

    if trimestre:
        resultats = resultats.filter(trimestre=trimestre)

    # Si un niveau est sélectionné, on filtre aussi la liste des classes
    if niveau:
        classes = classes.filter(niveau=niveau)

    # -----------------------------
    # TOP 5 DES MEILLEURS ÉLÈVES
    # -----------------------------

    # Une moyenne trimestrielle égale à 0 correspond à un élève non
    # classé : on l'exclut pour ne pas fausser le classement des
    # meilleurs élèves.
    meilleurs_eleves_liste = list(
        resultats
        .exclude(moyenne_trimestrielle=0)
        .order_by("-moyenne_trimestrielle", "eleve__nom", "eleve__prenoms")[:5]
    )

    for rang, resultat in enumerate(meilleurs_eleves_liste, start=1):
        resultat.rang = rang

        moyenne = resultat.moyenne_trimestrielle

        if moyenne is not None and moyenne >= 16:
            resultat.performance = "Excellent"
            resultat.performance_css = "excellent"
        elif moyenne is not None and moyenne >= 10:
            resultat.performance = "Satisfaisant"
            resultat.performance_css = "satisfaisant"
        else:
            resultat.performance = "À accompagner"
            resultat.performance_css = "a-accompagner"

        _attacher_details_matieres(resultat)

    # -----------------------------
    # LIBELLÉS DES FILTRES ACTIFS + PÉRIODE
    # -----------------------------

    filtres_actifs = []

    annee_obj = None
    if annee_id:
        annee_obj = AnneeScolaire.objects.filter(id=annee_id).first()
        if annee_obj:
            filtres_actifs.append({
                "libelle": "Année scolaire",
                "valeur": annee_obj.nom,
            })

    if niveau:
        filtres_actifs.append({
            "libelle": "Niveau",
            "valeur": niveau,
        })

    classe_obj = None
    if classe_id:
        classe_obj = Classe.objects.filter(id=classe_id).first()
        if classe_obj:
            filtres_actifs.append({
                "libelle": "Classe",
                "valeur": classe_obj.nom,
            })

    trimestre_labels = {
        "T1": "Trimestre 1",
        "T2": "Trimestre 2",
        "T3": "Trimestre 3",
    }

    if trimestre:
        filtres_actifs.append({
            "libelle": "Trimestre",
            "valeur": trimestre_labels.get(trimestre, trimestre),
        })

    periode_annee = annee_obj.nom if annee_obj else "Toutes les années"
    periode_trimestre = trimestre_labels.get(
        trimestre,
        "Tous les trimestres"
    )

    if annee_obj and trimestre:
        periode_analysee = f"{periode_annee} — {periode_trimestre}"
    elif annee_obj:
        periode_analysee = f"{periode_annee} — tous les trimestres"
    elif trimestre:
        periode_analysee = f"Toutes les années — {periode_trimestre}"
    else:
        periode_analysee = "Toutes les années — tous les trimestres"

    context = {

        "annees": annees,
        "classes": classes,
        "niveaux": niveaux_uniques,

        "annee_selectionnee": annee_id,
        "classe_selectionnee": classe_id,
        "trimestre_selectionne": trimestre,
        "niveau_selectionne": niveau,

        "filtres_actifs": filtres_actifs,
        "periode_analysee": periode_analysee,

        "meilleurs_eleves": meilleurs_eleves_liste,

    }

    return render(
        request,
        "analyse/meilleurs_eleves.html",
        context
    )


# =====================================================================
# FILTRES COMMUNS AUX PAGES "RÉSULTATS" ET "ÉLÈVES À SUIVRE"
# =====================================================================
#
# Ces deux pages reprennent exactement les mêmes filtres que le
# dashboard (année, niveau, classe, trimestre, recherche). Le code
# ci-dessous évite de dupliquer ce bloc dans chaque vue.

TRIMESTRE_LABELS = {
    "T1": "Trimestre 1",
    "T2": "Trimestre 2",
    "T3": "Trimestre 3",
}


def _preparer_filtres_page(request):
    """Applique les filtres GET sur les résultats et renvoie
    (resultats, contexte) où `contexte` contient tout ce dont le
    formulaire de filtres a besoin (listes, valeurs sélectionnées,
    filtres actifs, période analysée, chaîne de requête)."""

    annee_id = request.GET.get("annee")
    classe_id = request.GET.get("classe")
    trimestre = request.GET.get("trimestre")
    niveau = request.GET.get("niveau")
    recherche = request.GET.get("recherche", "").strip()

    annees = AnneeScolaire.objects.all().order_by("-nom")
    classes = Classe.objects.all().order_by("nom")

    niveaux_uniques = Classe.objects.values_list(
        "niveau", flat=True
    ).distinct().exclude(niveau="").order_by("niveau")

    resultats = Resultat.objects.select_related(
        "eleve",
        "annee_scolaire",
        "classe"
    )

    if niveau:
        resultats = resultats.filter(classe__niveau=niveau)

    if annee_id:
        resultats = resultats.filter(annee_scolaire_id=annee_id)

    if classe_id:
        resultats = resultats.filter(classe_id=classe_id)

    if trimestre:
        resultats = resultats.filter(trimestre=trimestre)

    if recherche:
        resultats = resultats.filter(
            Q(eleve__nom__icontains=recherche)
            | Q(eleve__prenoms__icontains=recherche)
            | Q(eleve__matricule__icontains=recherche)
        )

    if niveau:
        classes = classes.filter(niveau=niveau)

    # ---- Libellés des filtres actifs ----

    filtres_actifs = []

    annee_obj = None
    if annee_id:
        annee_obj = AnneeScolaire.objects.filter(id=annee_id).first()
        if annee_obj:
            filtres_actifs.append({
                "libelle": "Année scolaire",
                "valeur": annee_obj.nom,
            })

    if niveau:
        filtres_actifs.append({"libelle": "Niveau", "valeur": niveau})

    if classe_id:
        classe_obj = Classe.objects.filter(id=classe_id).first()
        if classe_obj:
            filtres_actifs.append({
                "libelle": "Classe",
                "valeur": classe_obj.nom,
            })

    if trimestre:
        filtres_actifs.append({
            "libelle": "Trimestre",
            "valeur": TRIMESTRE_LABELS.get(trimestre, trimestre),
        })

    if recherche:
        filtres_actifs.append({"libelle": "Recherche", "valeur": recherche})

    # ---- Période analysée ----

    periode_annee = annee_obj.nom if annee_obj else "Toutes les années"
    periode_trimestre = TRIMESTRE_LABELS.get(
        trimestre, "Tous les trimestres"
    )

    if annee_obj and trimestre:
        periode_analysee = f"{periode_annee} — {periode_trimestre}"
    elif annee_obj:
        periode_analysee = f"{periode_annee} — tous les trimestres"
    elif trimestre:
        periode_analysee = f"Toutes les années — {periode_trimestre}"
    else:
        periode_analysee = "Toutes les années — tous les trimestres"

    # Chaîne de requête sans le numéro de page (liens de pagination
    # et liens d'export).
    querydict = request.GET.copy()
    querydict.pop("page", None)

    contexte = {
        "annees": annees,
        "classes": classes,
        "niveaux": niveaux_uniques,

        "annee_selectionnee": annee_id,
        "classe_selectionnee": classe_id,
        "trimestre_selectionne": trimestre,
        "niveau_selectionne": niveau,
        "recherche": recherche,

        "filtres_actifs": filtres_actifs,
        "periode_analysee": periode_analysee,
        "querystring": querydict.urlencode(),
    }

    return resultats, contexte


def _attacher_details_matieres(resultat):
    """Ajoute `details_matieres` à un résultat : la liste des matières
    (parmi celles définies dans MATIERES) où l'élève a une note
    renseignée, avec son nom et sa note. Utilisé pour l'affichage
    dépliable « Voir les matières » du classement."""

    resultat.details_matieres = [
        {"nom": nom, "note": getattr(resultat, champ)}
        for nom, champ in MATIERES.items()
        if getattr(resultat, champ) is not None
    ]


def _definir_performance(resultat):
    """Ajoute `performance` et `performance_css` à un résultat
    (mêmes seuils que le classement du dashboard)."""

    moyenne = resultat.moyenne_trimestrielle

    if moyenne is None:
        resultat.performance = "À accompagner"
        resultat.performance_css = "a-accompagner"
    elif moyenne >= 16:
        resultat.performance = "Excellent"
        resultat.performance_css = "excellent"
    elif moyenne >= 10:
        resultat.performance = "Satisfaisant"
        resultat.performance_css = "satisfaisant"
    else:
        resultat.performance = "À accompagner"
        resultat.performance_css = "a-accompagner"


# =====================================================================
# ÉVOLUTION DE LA MOYENNE PAR RAPPORT AU TRIMESTRE PRÉCÉDENT
# =====================================================================
#
# Pour chaque élève affiché, on compare sa moyenne au trimestre
# précédent :
#   - T2 -> comparé à T1 (même année scolaire)
#   - T3 -> comparé à T2 (même année scolaire)
#   - T1 -> aucune comparaison (rien n'est affiché)
#
# La comparaison ne dépend pas des filtres du dashboard : le trimestre
# précédent est toujours recherché directement dans la base. Une seule
# requête SQL est exécutée pour toute la page.

TRIMESTRE_PRECEDENT = {"T2": "T1", "T3": "T2"}


def _attacher_evolution(resultats_page):
    """Ajoute à chaque résultat :
      - evolution            : écart de moyenne (Decimal) ou None
      - evolution_sens       : "hausse", "baisse", "stable" ou None
      - evolution_ref        : trimestre de comparaison (ex. "T1")
      - evolution_precedente : moyenne du trimestre de comparaison
    """

    resultats_page = list(resultats_page)

    if not resultats_page:
        return

    eleves_ids = {r.eleve_id for r in resultats_page}

    moyennes = {}

    lignes = Resultat.objects.filter(
        eleve_id__in=eleves_ids
    ).values_list(
        "eleve_id", "annee_scolaire_id", "trimestre", "moyenne_trimestrielle"
    )

    for eleve_id, annee_id, trim, moy in lignes:
        moyennes[(eleve_id, annee_id, trim)] = moy

    for resultat in resultats_page:

        resultat.evolution = None
        resultat.evolution_sens = None
        resultat.evolution_ref = ""
        resultat.evolution_precedente = None

        actuelle = resultat.moyenne_trimestrielle

        if actuelle is None or actuelle == 0:
            continue

        trim = resultat.trimestre

        if trim in TRIMESTRE_PRECEDENT:
            cle = (
                resultat.eleve_id,
                resultat.annee_scolaire_id,
                TRIMESTRE_PRECEDENT[trim],
            )
            reference = TRIMESTRE_PRECEDENT[trim]

        else:
            continue

        precedente = moyennes.get(cle)

        if precedente is None or precedente == 0:
            continue

        ecart = (
            Decimal(str(actuelle)) - Decimal(str(precedente))
        ).quantize(Decimal("0.01"))

        resultat.evolution = ecart
        resultat.evolution_ref = reference
        resultat.evolution_precedente = precedente

        if ecart > 0:
            resultat.evolution_sens = "hausse"
        elif ecart < 0:
            resultat.evolution_sens = "baisse"
        else:
            resultat.evolution_sens = "stable"


# =====================================================================
# PAGE "RÉSULTATS" : CLASSEMENT COMPLET DES ÉLÈVES
# =====================================================================

def resultats_eleves(request):
    """Page dédiée au classement des élèves (même contenu que le
    bloc « Classement des élèves » du dashboard, qui y est conservé),
    avec les mêmes filtres et une pagination de 15 élèves par page."""

    resultats, contexte = _preparer_filtres_page(request)

    # ---- Indicateurs de synthèse ----

    # Une moyenne de 0 correspond à un élève non classé : elle est
    # exclue des indicateurs (comme sur le dashboard).
    statistiques = resultats.aggregate(nombre_eleves=Count("id"))

    statistiques.update(
        resultats.exclude(moyenne_trimestrielle=0).aggregate(
            moyenne_classe=Avg("moyenne_trimestrielle"),
            meilleure_moyenne=Max("moyenne_trimestrielle"),
            plus_faible_moyenne=Min("moyenne_trimestrielle"),
        )
    )

    # ---- Classement (rang calculé avant la pagination) ----

    classement_liste = list(
        resultats.order_by(
            "-moyenne_trimestrielle",
            "eleve__nom",
            "eleve__prenoms",
        )
    )

    for rang, resultat in enumerate(classement_liste, start=1):
        resultat.rang = rang
        _definir_performance(resultat)
        _attacher_details_matieres(resultat)

    paginator = Paginator(classement_liste, 15)
    classement = paginator.get_page(request.GET.get("page"))

    # Évolution par rapport au trimestre précédent (page affichée seulement)
    _attacher_evolution(classement)

    contexte.update({
        "classement": classement,
        "total_resultats": len(classement_liste),
        "statistiques": statistiques,
    })

    return render(request, "analyse/resultats.html", contexte)


# =====================================================================
# PAGE "ÉLÈVES À SUIVRE"
# =====================================================================

def eleves_a_suivre(request):
    """Page dédiée au suivi pédagogique : liste complète des élèves
    dont la moyenne est inférieure à 10, matières les plus faibles et
    classes ayant le plus faible taux de réussite, selon les mêmes
    filtres que le dashboard."""

    resultats, contexte = _preparer_filtres_page(request)

    # ---- Élèves sous 10 (même définition que sur le dashboard et
    #      que l'export Excel) ----

    eleves_sous_10 = list(
        resultats
        .filter(moyenne_trimestrielle__lt=10)
        .order_by("moyenne_trimestrielle", "eleve__nom", "eleve__prenoms")
    )

    for rang, resultat in enumerate(eleves_sous_10, start=1):
        resultat.rang = rang
        _attacher_details_matieres(resultat)

    paginator = Paginator(eleves_sous_10, 15)
    liste_eleves = paginator.get_page(request.GET.get("page"))

    # ---- Matières les plus faibles ----

    statistiques_mat = calculer_statistiques_matieres(resultats)
    matieres_faibles = sorted(
        [m for m in statistiques_mat if m["moyenne"] is not None],
        key=lambda m: m["moyenne"]
    )[:5]

    # ---- Classes ayant le plus faible taux de réussite ----

    lignes_classes = (
        resultats
        .filter(moyenne_trimestrielle__gt=0)
        .order_by()
        .values("classe_id", "classe__nom")
        .annotate(
            total=Count("id"),
            reussis=Count(
                "id",
                filter=Q(moyenne_trimestrielle__gte=SEUIL_REUSSITE)
            ),
        )
    )

    classes_faibles = [
        {
            "classe": ligne["classe__nom"],
            "total": ligne["total"],
            "reussis": ligne["reussis"],
            "taux_reussite": round(
                ligne["reussis"] / ligne["total"] * 100, 2
            ),
        }
        for ligne in lignes_classes
    ]

    classes_faibles.sort(key=lambda c: (c["taux_reussite"], c["classe"]))
    classes_faibles = classes_faibles[:10]

    # ---- Indicateurs de synthèse ----

    total_resultats = resultats.count()
    nombre_sous_10 = len(eleves_sous_10)

    part_sous_10 = (
        round(nombre_sous_10 / total_resultats * 100, 2)
        if total_resultats else 0
    )

    contexte.update({
        "liste_eleves": liste_eleves,
        "total_resultats": total_resultats,
        "nombre_sous_10": nombre_sous_10,
        "part_sous_10": part_sous_10,
        "matieres_faibles": matieres_faibles,
        "classes_faibles": classes_faibles,
    })

    return render(request, "analyse/eleves_a_suivre.html", contexte)


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

    statistiques_matieres = calculer_statistiques_matieres(resultats)

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


# =====================================================================
# EXPORT EXCEL DE TOUTES LES STATISTIQUES DU DASHBOARD
# =====================================================================

POLICE_ENTETE = Font(bold=True, name="Arial")
ALIGNEMENT_CENTRE = Alignment(horizontal="center", vertical="center")


def _ecrire_entetes(feuille, entetes):
    """Écrit une ligne d'en-têtes stylée (gras, sans couleur) en haut
    d'une feuille."""

    for colonne, texte in enumerate(entetes, start=1):

        cellule = feuille.cell(row=1, column=colonne, value=texte)
        cellule.font = POLICE_ENTETE
        cellule.alignment = ALIGNEMENT_CENTRE

    feuille.freeze_panes = "A2"


def _ajuster_largeurs_colonnes(feuille):
    """Ajuste automatiquement la largeur des colonnes selon leur contenu."""

    for colonne in feuille.columns:

        longueur_max = max(
            (len(str(cellule.value)) for cellule in colonne if cellule.value is not None),
            default=10
        )

        lettre = get_column_letter(colonne[0].column)
        feuille.column_dimensions[lettre].width = min(longueur_max + 4, 40)


# ---- Styles reproduisant la structure des tableaux "par niveau" du
#      dashboard, sans couleurs (gras + bordures uniquement) ----

POLICE_ENTETE_TABLEAU = Font(bold=True, name="Arial")

BORDURE_FINE = Border(
    left=Side(style="thin", color="333333"),
    right=Side(style="thin", color="333333"),
    top=Side(style="thin", color="333333"),
    bottom=Side(style="thin", color="333333"),
)


def _cellule_entete_sombre(feuille, ligne, colonne, texte):
    """Écrit une cellule d'en-tête au style des tableaux "par niveau"
    du dashboard (gras, centré, bordé, sans couleur)."""

    cellule = feuille.cell(row=ligne, column=colonne, value=texte)
    cellule.font = POLICE_ENTETE_TABLEAU
    cellule.alignment = ALIGNEMENT_CENTRE
    cellule.border = BORDURE_FINE

    return cellule


def _remplir_ligne_donnees(feuille, ligne_excel, valeurs, est_total, est_total_general,
                            colonnes_pourcentage=()):
    """Écrit une ligne de données, avec les lignes de total en gras
    (comme les classes CSS total-row / total-general du dashboard),
    mais sans couleur de fond."""

    for indice_colonne, valeur in enumerate(valeurs, start=1):

        cellule = feuille.cell(row=ligne_excel, column=indice_colonne, value=valeur)
        cellule.border = BORDURE_FINE
        cellule.alignment = ALIGNEMENT_CENTRE

        if indice_colonne in colonnes_pourcentage and isinstance(valeur, (int, float)):
            cellule.value = valeur / 100
            cellule.number_format = "0.00%"

        if est_total_general or est_total or indice_colonne == 1:
            cellule.font = Font(bold=True, name="Arial")


def _construire_feuille_statistiques_niveaux(feuille, lignes):
    """Construit la feuille "Statistiques par niveau" en reproduisant

    exactement le tableau du dashboard : en-têtes groupés sur deux
    lignes, ordre de colonnes F/G/T, pourcentages et mise en
    évidence des lignes de total."""

    groupes = [
        ("Effectif classé", ["F", "G", "T"]),
        ("Effectif non classé", ["F", "G", "T"]),
        ("MOY TRIMES >= 10", ["F", "G", "T", "%"]),
        ("08.50 <= MOY TRIMES < 10", ["F", "G", "T", "%"]),
        ("MOY TRIMES < 08.50", ["F", "G", "T", "%"]),
    ]

    _cellule_entete_sombre(feuille, 1, 1, "Niveaux")
    feuille.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)

    _cellule_entete_sombre(feuille, 1, 2, "Nombre de classes")
    feuille.merge_cells(start_row=1, start_column=2, end_row=2, end_column=2)

    colonne = 3
    colonnes_pourcentage = []

    for titre_groupe, sous_colonnes in groupes:

        largeur = len(sous_colonnes)

        _cellule_entete_sombre(feuille, 1, colonne, titre_groupe)
        feuille.merge_cells(
            start_row=1, start_column=colonne,
            end_row=1, end_column=colonne + largeur - 1
        )

        for decalage, sous_titre in enumerate(sous_colonnes):
            _cellule_entete_sombre(feuille, 2, colonne + decalage, sous_titre)
            if sous_titre == "%":
                colonnes_pourcentage.append(colonne + decalage)

        colonne += largeur

    _cellule_entete_sombre(feuille, 1, colonne, "Moyenne par niveau")
    feuille.merge_cells(start_row=1, start_column=colonne, end_row=2, end_column=colonne)

    nombre_colonnes = colonne

    ligne_excel = 3

    for ligne in lignes:

        valeurs = [
            ligne["niveau"],
            ligne["nombre_classes"],
            ligne["classe_f"], ligne["classe_g"], ligne["classe_t"],
            ligne["non_classe_f"], ligne["non_classe_g"], ligne["non_classe_t"],
            ligne["sup10_f"], ligne["sup10_g"], ligne["sup10_t"], ligne["sup10_pct"],
            ligne["moy_f"], ligne["moy_g"], ligne["moy_t"], ligne["moy_pct"],
            ligne["inf_f"], ligne["inf_g"], ligne["inf_t"], ligne["inf_pct"],
            ligne["moyenne_niveau"] if ligne["moyenne_niveau"] is not None else "-",
        ]

        _remplir_ligne_donnees(
            feuille,
            ligne_excel,
            valeurs,
            est_total=ligne.get("is_total", False),
            est_total_general=(ligne["niveau"] == "Total général"),
            colonnes_pourcentage=colonnes_pourcentage,
        )

        ligne_excel += 1

    feuille.freeze_panes = "C3"

    feuille.column_dimensions["A"].width = 22
    for indice_colonne in range(2, nombre_colonnes + 1):
        feuille.column_dimensions[get_column_letter(indice_colonne)].width = 12


def _construire_feuille_effectifs_niveaux(feuille, lignes):
    """Construit la feuille "Effectifs" en reproduisant exactement le
    tableau du dashboard : en-têtes groupés sur deux lignes, ordre
    de colonnes G/F/Total et mise en évidence des lignes de total."""

    groupes = [
        ("Effectif", ["G", "F", "Total"]),
        ("Affectés", ["G", "F", "Total"]),
        ("Redoublants", ["G", "F", "Total"]),
    ]

    _cellule_entete_sombre(feuille, 1, 1, "Niveau")
    feuille.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)

    colonne = 2

    for titre_groupe, sous_colonnes in groupes:

        largeur = len(sous_colonnes)

        _cellule_entete_sombre(feuille, 1, colonne, titre_groupe)
        feuille.merge_cells(
            start_row=1, start_column=colonne,
            end_row=1, end_column=colonne + largeur - 1
        )

        for decalage, sous_titre in enumerate(sous_colonnes):
            _cellule_entete_sombre(feuille, 2, colonne + decalage, sous_titre)

        colonne += largeur

    nombre_colonnes = colonne - 1

    ligne_excel = 3

    for ligne in lignes:

        valeurs = [
            ligne["niveau"],
            ligne["effectif_g"], ligne["effectif_f"], ligne["effectif_t"],
            ligne["affectes_g"], ligne["affectes_f"], ligne["affectes_t"],
            ligne["redoublants_g"], ligne["redoublants_f"], ligne["redoublants_t"],
        ]

        _remplir_ligne_donnees(
            feuille,
            ligne_excel,
            valeurs,
            est_total=ligne.get("is_total", False),
            est_total_general=(ligne["niveau"] == "Total Général"),
        )

        ligne_excel += 1

    feuille.freeze_panes = "B3"

    feuille.column_dimensions["A"].width = 22
    for indice_colonne in range(2, nombre_colonnes + 1):
        feuille.column_dimensions[get_column_letter(indice_colonne)].width = 12


def exporter_eleves_sous_10_excel(request):
    """Exporte, dans un classeur Excel, la liste complète des élèves
    dont la moyenne trimestrielle est inférieure à 10, du plus faible
    au plus élevé, en respectant les filtres actuellement appliqués
    sur le dashboard (année, niveau, classe, trimestre, recherche)."""

    # -----------------------------
    # FILTRES (identiques à la vue dashboard)
    # -----------------------------

    annee_id = request.GET.get("annee")
    classe_id = request.GET.get("classe")
    trimestre = request.GET.get("trimestre")
    niveau = request.GET.get("niveau")
    recherche = request.GET.get("recherche", "").strip()

    resultats = Resultat.objects.select_related(
        "eleve",
        "annee_scolaire",
        "classe"
    )

    if niveau:
        resultats = resultats.filter(classe__niveau=niveau)

    if annee_id:
        resultats = resultats.filter(annee_scolaire_id=annee_id)

    if classe_id:
        resultats = resultats.filter(classe_id=classe_id)

    if trimestre:
        resultats = resultats.filter(trimestre=trimestre)

    if recherche:
        resultats = resultats.filter(
            Q(eleve__nom__icontains=recherche)
            | Q(eleve__prenoms__icontains=recherche)
            | Q(eleve__matricule__icontains=recherche)
        )

    eleves_sous_10 = resultats.filter(
        moyenne_trimestrielle__lt=10
    ).order_by("moyenne_trimestrielle", "eleve__nom", "eleve__prenoms")

    # -----------------------------
    # CONSTRUCTION DU CLASSEUR EXCEL
    # -----------------------------

    classeur = Workbook()

    feuille = classeur.active
    feuille.title = "Eleves sous 10"

    _ecrire_entetes(
        feuille,
        ["Rang", "Matricule", "Nom", "Prénoms", "Classe", "Niveau", "Moyenne"]
    )

    ligne_excel = 2

    for rang, resultat in enumerate(eleves_sous_10, start=1):

        feuille.cell(row=ligne_excel, column=1, value=rang)
        feuille.cell(row=ligne_excel, column=2, value=resultat.eleve.matricule)
        feuille.cell(row=ligne_excel, column=3, value=resultat.eleve.nom)
        feuille.cell(row=ligne_excel, column=4, value=resultat.eleve.prenoms)
        feuille.cell(row=ligne_excel, column=5, value=resultat.classe.nom)
        feuille.cell(row=ligne_excel, column=6, value=resultat.classe.niveau)
        feuille.cell(
            row=ligne_excel,
            column=7,
            value=float(resultat.moyenne_trimestrielle)
        )

        ligne_excel += 1

    _ajuster_largeurs_colonnes(feuille)

    # -----------------------------
    # RÉPONSE HTTP (téléchargement du fichier)
    # -----------------------------

    tampon = io.BytesIO()
    classeur.save(tampon)
    tampon.seek(0)

    horodatage = datetime.now().strftime("%Y%m%d_%H%M")
    nom_fichier = f"eleves_sous_10_{horodatage}.xlsx"

    reponse = HttpResponse(
        tampon.getvalue(),
        content_type=(
            "application/vnd.openxmlformats-officedocument"
            ".spreadsheetml.sheet"
        )
    )
    reponse["Content-Disposition"] = f'attachment; filename="{nom_fichier}"'

    return reponse


def exporter_statistiques_excel(request):
    """Exporte, dans un classeur Excel à plusieurs feuilles, l'ensemble
    des statistiques affichées sur le dashboard (résumé, classement
    complet des élèves, tableau par niveau, tableau des effectifs),
    en respectant les filtres actuellement appliqués (année, niveau,
    classe, trimestre)."""

    # -----------------------------
    # FILTRES (identiques à la vue dashboard)
    # -----------------------------

    annee_id = request.GET.get("annee")
    classe_id = request.GET.get("classe")
    trimestre = request.GET.get("trimestre")
    niveau = request.GET.get("niveau")
    recherche = request.GET.get("recherche", "").strip()

    resultats = Resultat.objects.select_related(
        "eleve",
        "annee_scolaire",
        "classe"
    )

    if niveau:
        resultats = resultats.filter(classe__niveau=niveau)

    if annee_id:
        resultats = resultats.filter(annee_scolaire_id=annee_id)

    if classe_id:
        resultats = resultats.filter(classe_id=classe_id)

    if trimestre:
        resultats = resultats.filter(trimestre=trimestre)

    if recherche:
        resultats = resultats.filter(
            Q(eleve__nom__icontains=recherche)
            | Q(eleve__prenoms__icontains=recherche)
            | Q(eleve__matricule__icontains=recherche)
        )

    # -----------------------------
    # STATISTIQUES (mêmes calculs que la vue dashboard)
    # -----------------------------

    resultats_avec_moyenne = resultats.exclude(
        moyenne_trimestrielle=0
    )

    statistiques = resultats.aggregate(
        nombre_eleves=Count("id"),
    )

    statistiques.update(
        resultats_avec_moyenne.aggregate(
            moyenne_classe=Avg("moyenne_trimestrielle"),
            meilleure_moyenne=Max("moyenne_trimestrielle"),
            plus_faible_moyenne=Min("moyenne_trimestrielle"),
        )
    )

    classement = resultats.order_by(
        "-moyenne_trimestrielle"
    )

    tableau_statistiques_niveaux = calculer_tableau_statistiques_niveaux(
        resultats
    )

    tableau_effectifs_niveaux = calculer_tableau_effectifs_niveaux(
        resultats
    )

    # -----------------------------
    # LIBELLÉS DES FILTRES APPLIQUÉS (pour le rappel dans le fichier)
    # -----------------------------

    annee_nom = "Toutes les années"
    if annee_id:
        annee_obj = AnneeScolaire.objects.filter(id=annee_id).first()
        annee_nom = annee_obj.nom if annee_obj else annee_id

    classe_nom = "Toutes les classes"
    if classe_id:
        classe_obj = Classe.objects.filter(id=classe_id).first()
        classe_nom = classe_obj.nom if classe_obj else classe_id

    niveau_nom = niveau or "Tous les niveaux"
    trimestre_nom = trimestre or "Tous les trimestres"

    # -----------------------------
    # CONSTRUCTION DU CLASSEUR EXCEL
    # -----------------------------

    classeur = Workbook()

    # ---- Feuille "Résumé" ----

    feuille_resume = classeur.active
    feuille_resume.title = "Résumé"

    _ecrire_entetes(feuille_resume, ["Indicateur", "Valeur"])

    moyenne_classe = statistiques.get("moyenne_classe")
    meilleure_moyenne = statistiques.get("meilleure_moyenne")
    plus_faible_moyenne = statistiques.get("plus_faible_moyenne")

    lignes_resume = [
        ("Date d'export", datetime.now().strftime("%d/%m/%Y %H:%M")),
        ("Année scolaire", annee_nom),
        ("Niveau", niveau_nom),
        ("Classe", classe_nom),
        ("Trimestre", trimestre_nom),
        ("Nombre d'élèves", statistiques.get("nombre_eleves") or 0),
        (
            "Moyenne générale",
            round(float(moyenne_classe), 2)
            if moyenne_classe is not None else "-"
        ),
        (
            "Meilleure moyenne",
            round(float(meilleure_moyenne), 2)
            if meilleure_moyenne is not None else "-"
        ),
        (
            "Plus faible moyenne",
            round(float(plus_faible_moyenne), 2)
            if plus_faible_moyenne is not None else "-"
        ),
    ]

    for indice, (libelle, valeur) in enumerate(lignes_resume, start=2):
        feuille_resume.cell(row=indice, column=1, value=libelle).font = Font(
            bold=True, name="Arial"
        )
        feuille_resume.cell(row=indice, column=2, value=valeur)

    _ajuster_largeurs_colonnes(feuille_resume)

    # ---- Feuille "Classement" ----

    feuille_classement = classeur.create_sheet("Classement")

    _ecrire_entetes(
        feuille_classement,
        ["Rang", "Matricule", "Nom", "Prénoms", "Classe", "Niveau", "Moyenne"]
    )

    rang = 0

    for indice, resultat in enumerate(classement, start=2):

        moyenne = resultat.moyenne_trimestrielle
        est_classe = moyenne not in (None, 0)

        if est_classe:
            rang += 1

        feuille_classement.cell(row=indice, column=1, value=rang if est_classe else "-")
        feuille_classement.cell(row=indice, column=2, value=resultat.eleve.matricule)
        feuille_classement.cell(row=indice, column=3, value=resultat.eleve.nom)
        feuille_classement.cell(row=indice, column=4, value=resultat.eleve.prenoms)
        feuille_classement.cell(row=indice, column=5, value=resultat.classe.nom)
        feuille_classement.cell(row=indice, column=6, value=resultat.classe.niveau)
        feuille_classement.cell(
            row=indice,
            column=7,
            value=float(moyenne) if est_classe else None
        )

    _ajuster_largeurs_colonnes(feuille_classement)

    # ---- Feuille "Statistiques par niveau" (identique au dashboard) ----

    feuille_niveaux = classeur.create_sheet("Statistiques par niveau")

    _construire_feuille_statistiques_niveaux(
        feuille_niveaux,
        tableau_statistiques_niveaux
    )

    # ---- Feuille "Effectifs" (identique au dashboard) ----

    feuille_effectifs = classeur.create_sheet("Effectifs")

    _construire_feuille_effectifs_niveaux(
        feuille_effectifs,
        tableau_effectifs_niveaux
    )

    # -----------------------------
    # RÉPONSE HTTP (téléchargement du fichier)
    # -----------------------------

    tampon = io.BytesIO()
    classeur.save(tampon)
    tampon.seek(0)

    horodatage = datetime.now().strftime("%Y%m%d_%H%M")
    nom_fichier = f"statistiques_{horodatage}.xlsx"

    reponse = HttpResponse(
        tampon.getvalue(),
        content_type=(
            "application/vnd.openxmlformats-officedocument"
            ".spreadsheetml.sheet"
        )
    )
    reponse["Content-Disposition"] = f'attachment; filename="{nom_fichier}"'

    return reponse



# =====================================================================
# PAGE "FICHE ÉLÈVE" : SUIVI INDIVIDUEL
# =====================================================================
#
# L'utilisateur recherche un élève (nom, prénoms ou matricule). La page
# affiche ensuite, pour une année scolaire choisie :
#   - les moyennes trimestrielles et leur évolution (hausse / baisse) ;
#   - les notes de chaque matière, trimestre par trimestre, avec
#     l'évolution d'un trimestre au suivant ;
#   - les matières où l'élève doit faire plus d'efforts ;
#   - son profil : scientifique, littéraire ou équilibré.
#
# Réglages (à ajuster selon les usages de l'établissement) :

ORDRE_TRIMESTRES = ["T1", "T2", "T3"]

# Baisse d'une matière (en points) à partir de laquelle on la signale.
SEUIL_BAISSE_NOTABLE = Decimal("1")

# Écart minimal (en points) entre les deux profils pour conclure qu'un
# élève est plutôt scientifique ou plutôt littéraire.
SEUIL_ECART_PROFIL = Decimal("1")

# Noms exacts des matières, tels que définis dans MATIERES.
# Conduite, EPS et EDHC ne sont classées dans aucun des deux profils.
MATIERES_SCIENTIFIQUES = ["Mathématiques", "Physiques-Chimie", "SVT"]

MATIERES_LITTERAIRES = [
    "Composition française",
    "Orthographe / Grammaire",
    "Expression orale",
    "Français",
    "Histoire-Géographie",
    "Anglais",
    "Espagnol",
    "Allemand",
    "Philosophie",
]


def _decimal(valeur):
    return Decimal(str(valeur))


def _evolution(actuelle, precedente):
    """Écart entre deux valeurs. Renvoie (ecart, sens) où sens vaut
    "hausse", "baisse" ou "stable", ou (None, None) si l'une des deux
    valeurs manque."""

    if actuelle is None or precedente is None:
        return None, None

    ecart = (_decimal(actuelle) - _decimal(precedente)).quantize(
        Decimal("0.01")
    )

    if ecart > 0:
        return ecart, "hausse"
    if ecart < 0:
        return ecart, "baisse"
    return ecart, "stable"


def _moyenne_classee(resultat):
    """Moyenne trimestrielle d'un résultat, ou None si l'élève n'est
    pas classé (moyenne absente ou égale à 0)."""

    if resultat is None:
        return None

    moyenne = resultat.moyenne_trimestrielle

    if moyenne is None or moyenne == 0:
        return None

    return _decimal(moyenne)


def _recherche_eleves(recherche):
    """Cherche des élèves : chaque mot saisi doit apparaître dans le
    nom, les prénoms ou le matricule."""

    filtre = Q()

    for mot in recherche.split():
        filtre &= (
            Q(nom__icontains=mot)
            | Q(prenoms__icontains=mot)
            | Q(matricule__icontains=mot)
        )

    return list(Eleve.objects.filter(filtre).order_by("nom", "prenoms")[:30])


def _ajouter_derniere_classe(eleves):
    """Associe à chaque élève sa classe la plus récente (une seule
    requête), pour aider à les distinguer dans la liste."""

    derniere_classe = {}

    lignes = Resultat.objects.filter(
        eleve_id__in=[e.id for e in eleves]
    ).select_related("classe", "annee_scolaire").order_by(
        "annee_scolaire__nom", "trimestre"
    )

    for ligne in lignes:
        derniere_classe[ligne.eleve_id] = ligne

    candidats = []

    for eleve in eleves:
        ligne = derniere_classe.get(eleve.id)
        candidats.append({
            "eleve": eleve,
            "classe": ligne.classe.nom if ligne else "",
            "annee": ligne.annee_scolaire.nom if ligne else "",
        })

    return candidats


def _rang_entier(valeur):
    """Rang sous forme d'entier positif, ou None s'il est absent."""

    try:
        rang = int(valeur)
    except (TypeError, ValueError):
        return None

    return rang if rang > 0 else None


def _effectifs_classes(par_trimestre):
    """Nombre d'élèves de la classe à chaque trimestre, pour afficher
    « 3e sur 45 ». Une seule requête."""

    classes_ids = {r.classe_id for r in par_trimestre.values()}
    annees_ids = {r.annee_scolaire_id for r in par_trimestre.values()}

    lignes = (
        Resultat.objects
        .filter(classe_id__in=classes_ids, annee_scolaire_id__in=annees_ids)
        .order_by()
        .values("classe_id", "trimestre")
        .annotate(total=Count("id"))
    )

    return {
        (ligne["classe_id"], ligne["trimestre"]): ligne["total"]
        for ligne in lignes
    }


def _construire_lignes_trimestres(par_trimestre, effectifs):
    """Une ligne par trimestre : moyenne, rang, classe et évolution
    par rapport au trimestre précédent."""

    lignes = []
    moyenne_precedente = None
    rang_precedent = None

    for trim in ORDRE_TRIMESTRES:

        resultat = par_trimestre.get(trim)
        moyenne = _moyenne_classee(resultat)

        ecart, sens = _evolution(moyenne, moyenne_precedente)

        rang = _rang_entier(resultat.rang) if resultat else None

        # Un rang plus petit est meilleur : gagner des places, c'est
        # passer à un rang numériquement plus bas.
        places_gagnees = None
        sens_rang = None

        if rang is not None and rang_precedent is not None:
            places_gagnees = rang_precedent - rang
            if places_gagnees > 0:
                sens_rang = "hausse"
            elif places_gagnees < 0:
                sens_rang = "baisse"
            else:
                sens_rang = "stable"

        lignes.append({
            "trimestre": trim,
            "libelle": TRIMESTRE_LABELS.get(trim, trim),
            "present": resultat is not None,
            "classe": resultat.classe.nom if resultat else "",
            "moyenne": moyenne,
            "ecart": ecart,
            "sens": sens,
            "rang": rang,
            "effectif": (
                effectifs.get((resultat.classe_id, trim)) if resultat else None
            ),
            "places_gagnees": places_gagnees,
            "sens_rang": sens_rang,
            "reference": ORDRE_TRIMESTRES[ORDRE_TRIMESTRES.index(trim) - 1]
            if trim != "T1" else "",
        })

        # Le trimestre suivant se compare au dernier trimestre classé.
        moyenne_precedente = moyenne
        rang_precedent = rang

    return lignes


def _construire_tableau_matieres(par_trimestre):
    """Une ligne par matière, avec les notes des trois trimestres et
    l'évolution d'un trimestre au suivant. Les matières sans aucune
    note sont ignorées."""

    lignes = []

    for nom, champ in MATIERES.items():

        notes = []

        for trim in ORDRE_TRIMESTRES:
            resultat = par_trimestre.get(trim)
            note = getattr(resultat, champ) if resultat else None
            notes.append(None if note is None else _decimal(note))

        if all(note is None for note in notes):
            continue

        cellules = []

        for indice, trim in enumerate(ORDRE_TRIMESTRES):

            if indice == 0:
                ecart, sens = None, None
            else:
                ecart, sens = _evolution(notes[indice], notes[indice - 1])

            cellules.append({
                "trimestre": trim,
                "note": notes[indice],
                "ecart": ecart,
                "sens": sens,
            })

        lignes.append({
            "nom": nom,
            "notes": notes,
            "cellules": cellules,
        })

    return lignes


def _matieres_a_renforcer(lignes_matieres, dernier, precedent):
    """Matières où l'élève doit faire plus d'efforts, d'après le
    dernier trimestre disponible : note inférieure à 10, ou baisse
    notable par rapport au trimestre précédent."""

    indice_dernier = ORDRE_TRIMESTRES.index(dernier)
    indice_precedent = (
        ORDRE_TRIMESTRES.index(precedent) if precedent else None
    )

    a_renforcer = []

    for ligne in lignes_matieres:

        note = ligne["notes"][indice_dernier]

        if note is None:
            continue

        sous_moyenne = note < SEUIL_REUSSITE
        baisse = None

        if indice_precedent is not None:
            note_precedente = ligne["notes"][indice_precedent]

            if note_precedente is not None:
                ecart = note - note_precedente

                if ecart <= -SEUIL_BAISSE_NOTABLE:
                    baisse = abs(ecart).quantize(Decimal("0.01"))

        if not sous_moyenne and baisse is None:
            continue

        a_renforcer.append({
            "nom": ligne["nom"],
            "note": note,
            "sous_moyenne": sous_moyenne,
            "baisse": baisse,
            "reference": precedent,
            "priorite_haute": sous_moyenne and baisse is not None,
        })

    a_renforcer.sort(key=lambda m: (not m["priorite_haute"], m["note"]))

    return a_renforcer


def _moyenne_groupe(par_trimestre, noms_matieres):
    """Moyenne d'un groupe de matières sur l'année : on calcule d'abord
    la moyenne annuelle de chaque matière, puis la moyenne de ces
    moyennes (chaque matière compte pour un)."""

    details = []

    for nom in noms_matieres:

        champ = MATIERES[nom]

        notes = [
            _decimal(getattr(resultat, champ))
            for resultat in par_trimestre.values()
            if getattr(resultat, champ) is not None
        ]

        if notes:
            details.append({
                "nom": nom,
                "moyenne": (sum(notes) / len(notes)).quantize(
                    Decimal("0.01")
                ),
            })

    if not details:
        return None, []

    moyenne = (
        sum(d["moyenne"] for d in details) / len(details)
    ).quantize(Decimal("0.01"))

    for detail in details:
        detail["pourcentage"] = float(detail["moyenne"]) / 20 * 100

    details.sort(key=lambda d: d["moyenne"], reverse=True)

    return moyenne, details


def _calculer_aptitude(par_trimestre):
    """Profil scientifique / littéraire de l'élève sur l'année."""

    moy_sci, details_sci = _moyenne_groupe(
        par_trimestre, MATIERES_SCIENTIFIQUES
    )
    moy_lit, details_lit = _moyenne_groupe(
        par_trimestre, MATIERES_LITTERAIRES
    )

    aptitude = {
        "scientifique": {
            "moyenne": moy_sci,
            "pourcentage": float(moy_sci) / 20 * 100 if moy_sci else 0,
            "details": details_sci,
        },
        "litteraire": {
            "moyenne": moy_lit,
            "pourcentage": float(moy_lit) / 20 * 100 if moy_lit else 0,
            "details": details_lit,
        },
        "ecart": None,
    }

    if moy_sci is None or moy_lit is None:
        aptitude["profil"] = "indetermine"
        aptitude["libelle"] = "Profil non déterminé"
        aptitude["description"] = (
            "Il manque des moyennes dans les matières scientifiques ou "
            "littéraires pour comparer les deux profils."
        )
        return aptitude

    ecart = (moy_sci - moy_lit).quantize(Decimal("0.01"))
    aptitude["ecart"] = abs(ecart)

    if ecart >= SEUIL_ECART_PROFIL:
        aptitude["profil"] = "scientifique"
        aptitude["libelle"] = "Profil scientifique"
        aptitude["description"] = (
            "Les résultats sont plus solides en mathématiques, "
            "physique-chimie et SVT que dans les matières littéraires."
        )
    elif ecart <= -SEUIL_ECART_PROFIL:
        aptitude["profil"] = "litteraire"
        aptitude["libelle"] = "Profil littéraire"
        aptitude["description"] = (
            "Les résultats sont plus solides en langues, français et "
            "histoire-géographie que dans les matières scientifiques."
        )
    else:
        aptitude["profil"] = "equilibre"
        aptitude["libelle"] = "Profil équilibré"
        aptitude["description"] = (
            "Les résultats sont proches dans les deux domaines : "
            "aucun ne se détache nettement."
        )

    return aptitude


def _donnees_fiche(eleve, annee_id=""):
    """Calcule tout le contenu de la fiche d'un élève pour une année
    scolaire (la plus récente si annee_id est vide ou inconnu). Sert à
    la page web comme à l'export PDF. Renvoie {"aucune_donnee": True}
    si l'élève n'a aucun résultat."""

    tous_resultats = list(
        Resultat.objects
        .filter(eleve=eleve)
        .select_related("annee_scolaire", "classe")
        .order_by("annee_scolaire__nom", "trimestre")
    )

    par_annee = {}

    for resultat in tous_resultats:
        groupe = par_annee.setdefault(
            resultat.annee_scolaire_id,
            {"annee": resultat.annee_scolaire, "trimestres": {}},
        )
        groupe["trimestres"][resultat.trimestre] = resultat

    if not par_annee:
        return {"aucune_donnee": True}

    annees_eleve = [g["annee"] for g in reversed(list(par_annee.values()))]

    if str(annee_id).isdigit() and int(annee_id) in par_annee:
        annee_choisie = par_annee[int(annee_id)]["annee"]
    else:
        annee_choisie = annees_eleve[0]

    par_trimestre = par_annee[annee_choisie.id]["trimestres"]

    # ---- Moyennes trimestrielles, rang et évolution ----

    lignes_trimestres = _construire_lignes_trimestres(
        par_trimestre, _effectifs_classes(par_trimestre)
    )

    moyennes_valides = [
        l["moyenne"] for l in lignes_trimestres if l["moyenne"] is not None
    ]

    # Même formule que le bilan de fin d'année : T1 x 1, T2 x 2, T3 x 2.
    # Tant que les trois trimestres ne sont pas classés, la moyenne est
    # provisoire (calculée sur les trimestres disponibles).
    moyenne_annuelle = _moyenne_ponderee({
        l["trimestre"]: l["moyenne"] for l in lignes_trimestres
    })

    moyenne_annuelle_provisoire = (
        len(moyennes_valides) < len(ORDRE_TRIMESTRES)
    )

    # Dernière évolution connue, pour le résumé en haut de page.
    derniere_evolution = next(
        (l for l in reversed(lignes_trimestres) if l["sens"]),
        None,
    )

    # ---- Matières ----

    lignes_matieres = _construire_tableau_matieres(par_trimestre)

    trimestres_disponibles = [
        t for t in ORDRE_TRIMESTRES if t in par_trimestre
    ]

    dernier = trimestres_disponibles[-1]
    precedent = (
        trimestres_disponibles[-2]
        if len(trimestres_disponibles) > 1 else ""
    )

    matieres_a_renforcer = _matieres_a_renforcer(
        lignes_matieres, dernier, precedent
    )

    ligne_dernier_trimestre = next(
        l for l in lignes_trimestres if l["trimestre"] == dernier
    )

    aptitude = _calculer_aptitude(par_trimestre)

    # ---- Données du graphique (toutes les années de l'élève) ----

    graphique_labels = []
    graphique_valeurs = []

    for groupe in par_annee.values():
        for trim in ORDRE_TRIMESTRES:
            if trim in groupe["trimestres"]:
                moyenne = _moyenne_classee(groupe["trimestres"][trim])
                graphique_labels.append(f"{groupe['annee'].nom} {trim}")
                graphique_valeurs.append(
                    float(moyenne) if moyenne is not None else None
                )

    return {
        "aucune_donnee": False,
        "annees_eleve": annees_eleve,
        "annee_choisie": annee_choisie,
        "lignes_trimestres": lignes_trimestres,
        "moyenne_annuelle": moyenne_annuelle,
        "moyenne_annuelle_provisoire": moyenne_annuelle_provisoire,
        "derniere_evolution": derniere_evolution,
        "lignes_matieres": lignes_matieres,
        "matieres_a_renforcer": matieres_a_renforcer,
        "dernier_trimestre": dernier,
        "ligne_dernier_trimestre": ligne_dernier_trimestre,
        "trimestre_reference": precedent,
        "aptitude": aptitude,
        "seuil_reussite": SEUIL_REUSSITE,
        "donnees_graphique": {
            "labels": graphique_labels,
            "valeurs": graphique_valeurs,
            "seuil": SEUIL_REUSSITE,
        },
    }


def fiche_eleve(request):
    """Fiche de suivi individuel : recherche d'un élève puis affichage
    de ses moyennes, de leur évolution, de son rang, des matières à
    renforcer et de son profil scientifique / littéraire."""

    recherche = request.GET.get("recherche", "").strip()
    eleve_id = request.GET.get("eleve", "")
    annee_id = request.GET.get("annee", "")

    eleve = None
    candidats = []
    recherche_sans_resultat = False

    # ---- Recherche de l'élève ----

    if eleve_id.isdigit():
        eleve = Eleve.objects.filter(id=int(eleve_id)).first()

    if eleve is None and recherche:

        trouves = _recherche_eleves(recherche)

        if len(trouves) == 1:
            eleve = trouves[0]
        elif trouves:
            candidats = _ajouter_derniere_classe(trouves)
        else:
            recherche_sans_resultat = True

    contexte = {
        "recherche": recherche,
        "candidats": candidats,
        "recherche_sans_resultat": recherche_sans_resultat,
        "eleve": eleve,
    }

    if eleve is not None:
        contexte.update(_donnees_fiche(eleve, annee_id))

    return render(request, "analyse/fiche_eleve.html", contexte)


# ---------------------------------------------------------------------
# Export PDF de la fiche élève
# ---------------------------------------------------------------------

def _fmt(valeur, signe=False):
    """Nombre au format français : 12,50 (ou +1,20 avec signe=True)."""

    if valeur is None:
        return "—"

    texte = f"{valeur:+.2f}" if signe else f"{valeur:.2f}"

    return texte.replace(".", ",")


def _rang_texte(rang):
    if not rang:
        return "—"

    return "1er" if rang == 1 else f"{rang}e"


def _generer_pdf_fiche(eleve, donnees):
    """Construit le PDF de la fiche et renvoie ses octets. Utilise la
    bibliothèque reportlab (pip install reportlab)."""

    from xml.sax.saxutils import escape

    from reportlab.graphics.shapes import Drawing, Rect, String
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import (
        KeepTogether,
        Paragraph,
        SimpleDocTemplate,
        Spacer,
        Table,
        TableStyle,
    )

    # Largeur utile des tableaux : la page moins les marges internes du cadre.
    FACTEUR = 16.5 / 17

    ACCENT = colors.HexColor("#2f5bea")
    VIOLET = colors.HexColor("#7a5af8")
    GRIS = colors.HexColor("#667085")
    FOND = colors.HexColor("#eef1f6")
    LIGNE = colors.HexColor("#e4e7ec")

    COULEURS_SENS = {
        "hausse": "#0b6b2f",
        "baisse": "#b42318",
        "stable": "#475467",
    }

    base = ParagraphStyle(
        "base", fontName="Helvetica", fontSize=9.5, leading=13,
        textColor=colors.HexColor("#1d2433"),
    )
    centre = ParagraphStyle("centre", parent=base, alignment=1)
    petit = ParagraphStyle(
        "petit", parent=base, fontSize=8.5, leading=11, textColor=GRIS,
    )
    titre = ParagraphStyle(
        "titre", parent=base, fontName="Helvetica-Bold", fontSize=19,
        leading=23,
    )
    titre2 = ParagraphStyle(
        "titre2", parent=base, fontName="Helvetica-Bold", fontSize=12,
        leading=15, textColor=ACCENT, spaceBefore=16, spaceAfter=6,
    )

    def p(texte, style=base):
        return Paragraph(texte, style)

    def evolution(ecart, sens, unite=""):
        """Écart coloré : +1,20 / -0,50 / = (sans flèches, absentes de
        la police standard du PDF)."""

        if not sens:
            return ""

        couleur = COULEURS_SENS[sens]
        texte = "=" if sens == "stable" else _fmt(ecart, signe=True)

        return f'<font color="{couleur}" size="8">{texte}{unite}</font>'

    def evolution_rang(places, sens):
        if not sens:
            return ""

        couleur = COULEURS_SENS[sens]

        if sens == "stable":
            texte = "même rang"
        else:
            mot = "place" if abs(places) == 1 else "places"
            texte = f"{places:+d} {mot}"

        return f'<font color="{couleur}" size="8">{texte}</font>'

    def note_texte(note):
        couleur = "#b42318" if note < donnees["seuil_reussite"] else "#1d2433"
        return f'<font color="{couleur}"><b>{_fmt(note)}</b></font>'

    def style_tableau(extra=None):
        regles = [
            ("BACKGROUND", (0, 0), (-1, 0), FOND),
            ("LINEBELOW", (0, 0), (-1, -1), 0.5, LIGNE),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ]
        return TableStyle(regles + (extra or []))

    def entete(texte, style=base):
        return p(f"<b>{texte}</b>", ParagraphStyle(
            "th", parent=style, fontSize=8.5, textColor=GRIS,
        ))

    annee = donnees["annee_choisie"]
    dernier = donnees["ligne_dernier_trimestre"]
    aptitude = donnees["aptitude"]

    contenu = []

    # ---- En-tête ----

    contenu.append(p("FICHE DE SUIVI DE L'ÉLÈVE", ParagraphStyle(
        "sur", parent=petit, fontName="Helvetica-Bold", textColor=ACCENT,
    )))
    contenu.append(p(
        f"{escape(eleve.nom)} {escape(eleve.prenoms)}", titre
    ))
    contenu.append(Spacer(1, 4))
    contenu.append(p(
        f"Matricule {escape(str(eleve.matricule))} &nbsp;·&nbsp; "
        f"Année scolaire {escape(str(annee.nom))}"
        + (f" &nbsp;·&nbsp; Classe {escape(dernier['classe'])}"
           if dernier["classe"] else ""),
        petit,
    ))

    # ---- Résumé ----

    evol = donnees["derniere_evolution"]

    cartes = [
        ("Moyenne annuelle (provisoire)"
         if donnees["moyenne_annuelle_provisoire"] else "Moyenne annuelle",
         _fmt(donnees["moyenne_annuelle"]) + " / 20"
         if donnees["moyenne_annuelle"] is not None else "—"),
        (
            f"Rang au {donnees['dernier_trimestre']}",
            _rang_texte(dernier["rang"])
            + (f" sur {dernier['effectif']}" if dernier["effectif"] else ""),
        ),
        (
            "Dernière évolution",
            f"{_fmt(evol['ecart'], signe=True)} ({evol['reference']} à {evol['trimestre']})"
            if evol else "—",
        ),
        ("Profil", aptitude["libelle"].replace("Profil ", "").capitalize()),
    ]

    resume = Table(
        [[
            [p(f'<font color="#667085" size="8">{lib}</font>'),
             p(f"<b>{escape(val)}</b>", ParagraphStyle(
                 "val", parent=base, fontSize=12, leading=15))]
            for lib, val in cartes
        ]],
        colWidths=[largeur * FACTEUR for largeur in [4.25 * cm]] * 4,
    )
    resume.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.5, LIGNE),
        ("INNERGRID", (0, 0), (-1, -1), 0.5, LIGNE),
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#f8f9fb")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
    ]))
    contenu.append(Spacer(1, 12))
    contenu.append(resume)

    # ---- Moyennes trimestrielles ----

    contenu.append(p("Moyennes trimestrielles et rang", titre2))

    lignes = [[
        entete("Trimestre"), entete("Classe"),
        entete("Moyenne", centre), entete("Évolution", centre),
        entete("Rang", centre), entete("Évolution du rang", centre),
    ]]

    for l in donnees["lignes_trimestres"]:
        if l["moyenne"] is not None:
            moyenne = note_texte(l["moyenne"])
        elif l["present"]:
            moyenne = '<font color="#667085">Non classé</font>'
        else:
            moyenne = "—"

        rang = _rang_texte(l["rang"])
        if l["rang"] and l["effectif"]:
            rang += f' <font color="#667085" size="8">/ {l["effectif"]}</font>'

        lignes.append([
            p(escape(l["libelle"])),
            p(escape(l["classe"]) or "—"),
            p(moyenne, centre),
            p(evolution(l["ecart"], l["sens"]) or "—", centre),
            p(rang, centre),
            p(evolution_rang(l["places_gagnees"], l["sens_rang"]) or "—",
              centre),
        ])

    tableau = Table(
        lignes,
        colWidths=[largeur * FACTEUR for largeur in [3 * cm, 2.6 * cm, 2.8 * cm, 2.8 * cm, 2.8 * cm, 3 * cm]],
    )
    tableau.setStyle(style_tableau())
    contenu.append(tableau)
    contenu.append(Spacer(1, 4))
    contenu.append(p(
        "L'évolution compare chaque trimestre au précédent. Pour le rang, "
        "« + » signifie que l'élève a gagné des places.", petit,
    ))

    # ---- Moyennes par matière ----

    contenu.append(p("Moyennes par matière", titre2))

    if donnees["lignes_matieres"]:

        lignes = [[
            entete("Matière"), entete("Trimestre 1", centre),
            entete("Trimestre 2", centre), entete("Trimestre 3", centre),
        ]]

        for ligne in donnees["lignes_matieres"]:
            cellules = [p(escape(ligne["nom"]))]

            for c in ligne["cellules"]:
                if c["note"] is None:
                    cellules.append(p("—", centre))
                    continue

                texte = note_texte(c["note"])
                if c["sens"]:
                    texte += "<br/>" + evolution(c["ecart"], c["sens"])

                cellules.append(p(texte, centre))

            lignes.append(cellules)

        tableau = Table(
            lignes, repeatRows=1,
            colWidths=[largeur * FACTEUR for largeur in [6.8 * cm, 3.4 * cm, 3.4 * cm, 3.4 * cm]],
        )
        tableau.setStyle(style_tableau())
        contenu.append(tableau)
        contenu.append(Spacer(1, 4))
        contenu.append(p(
            "Sous chaque moyenne : évolution par rapport au trimestre "
            "précédent.", petit,
        ))
    else:
        contenu.append(p("Aucune moyenne par matière enregistrée."))

    # ---- Matières à renforcer ----

    bloc = [p("Matières où faire plus d'efforts", titre2)]

    if donnees["trimestre_reference"]:
        critere = (
            f"D'après le {donnees['dernier_trimestre']} : moyenne inférieure "
            f"à {donnees['seuil_reussite']} ou baisse d'au moins "
            f"{_fmt(SEUIL_BAISSE_NOTABLE).rstrip('0').rstrip(',')} point "
            f"depuis le {donnees['trimestre_reference']}."
        )
    else:
        critere = (
            f"D'après le {donnees['dernier_trimestre']} : moyenne "
            f"inférieure à {donnees['seuil_reussite']}."
        )

    bloc.append(p(critere, petit))
    bloc.append(Spacer(1, 6))

    if donnees["matieres_a_renforcer"]:

        lignes = []

        for m in donnees["matieres_a_renforcer"]:
            raisons = []

            if m["priorite_haute"]:
                raisons.append('<font color="#b42318"><b>Priorité</b></font>')
            if m["sous_moyenne"]:
                raisons.append("sous la moyenne")
            if m["baisse"]:
                raisons.append(
                    f"baisse de {_fmt(m['baisse'])} pt "
                    f"depuis {m['reference']}"
                )

            lignes.append([
                p(f"<b>{escape(m['nom'])}</b> — {_fmt(m['note'])}/20"),
                p(" · ".join(raisons)),
            ])

        tableau = Table(lignes, colWidths=[largeur * FACTEUR for largeur in [7 * cm, 10 * cm]])
        tableau.setStyle(TableStyle([
            ("LINEBELOW", (0, 0), (-1, -1), 0.5, LIGNE),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ]))
        bloc.append(tableau)
    else:
        bloc.append(p(
            f'<font color="#0b6b2f"><b>Aucune matière en difficulté au '
            f"{donnees['dernier_trimestre']} : toutes les moyennes sont à "
            f"{donnees['seuil_reussite']} ou plus, sans baisse notable."
            f"</b></font>"
        ))

    contenu.append(KeepTogether(bloc))

    # ---- Profil scientifique / littéraire ----

    sci = aptitude["scientifique"]
    lit = aptitude["litteraire"]

    bloc = [
        p("Capacité scientifique ou littéraire", titre2),
        p(f"<b>{escape(aptitude['libelle'])}</b>", ParagraphStyle(
            "verdict", parent=base, fontSize=11, leading=14)),
        Spacer(1, 2),
        p(escape(aptitude["description"])
          + (f" Écart entre les deux domaines : "
             f"{_fmt(aptitude['ecart'])} points."
             if aptitude["ecart"] else "")),
        Spacer(1, 8),
    ]

    largeur_piste = 9.5 * cm
    dessin = Drawing(17 * cm, 1.7 * cm)

    for indice, (libelle, groupe, couleur) in enumerate([
        ("Matières scientifiques", sci, ACCENT),
        ("Matières littéraires", lit, VIOLET),
    ]):
        y = 0.95 * cm - indice * 0.85 * cm
        dessin.add(String(
            0, y + 0.08 * cm, libelle,
            fontName="Helvetica", fontSize=9, fillColor=colors.black,
        ))
        dessin.add(Rect(
            5 * cm, y, largeur_piste, 0.35 * cm,
            fillColor=FOND, strokeColor=None,
        ))
        if groupe["moyenne"] is not None:
            dessin.add(Rect(
                5 * cm, y, largeur_piste * groupe["pourcentage"] / 100,
                0.35 * cm, fillColor=couleur, strokeColor=None,
            ))
        dessin.add(String(
            5 * cm + largeur_piste + 0.3 * cm, y + 0.08 * cm,
            f"{_fmt(groupe['moyenne'])} / 20"
            if groupe["moyenne"] is not None else "—",
            fontName="Helvetica-Bold", fontSize=9, fillColor=colors.black,
        ))

    bloc.append(dessin)
    bloc.append(Spacer(1, 6))

    def liste(groupe):
        if not groupe["details"]:
            return p("Aucune moyenne disponible.", petit)

        return p("<br/>".join(
            f"{escape(d['nom'])} : <b>{_fmt(d['moyenne'])}</b>"
            for d in groupe["details"]
        ), petit)

    details = Table(
        [[liste(sci), liste(lit)]], colWidths=[largeur * FACTEUR for largeur in [8.5 * cm, 8.5 * cm]],
    )
    details.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
    ]))
    bloc.append(details)
    bloc.append(Spacer(1, 6))
    bloc.append(p(
        "Moyennes calculées sur les trimestres de l'année affichée. "
        "Conduite, EPS et EDHC ne sont pas comptées. Ce profil est un "
        "repère, à croiser avec l'avis des enseignants.", petit,
    ))

    contenu.append(KeepTogether(bloc))

    # ---- Assemblage ----

    def pied_de_page(canvas, doc):
        canvas.saveState()
        canvas.setStrokeColor(LIGNE)
        canvas.line(2 * cm, 1.6 * cm, A4[0] - 2 * cm, 1.6 * cm)
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(GRIS)
        canvas.drawString(
            2 * cm, 1.1 * cm,
            f"Analyse scolaire — généré le {datetime.now():%d/%m/%Y}",
        )
        canvas.drawRightString(A4[0] - 2 * cm, 1.1 * cm, f"Page {doc.page}")
        canvas.restoreState()

    tampon = io.BytesIO()

    document = SimpleDocTemplate(
        tampon, pagesize=A4,
        leftMargin=2 * cm, rightMargin=2 * cm,
        topMargin=1.8 * cm, bottomMargin=2.2 * cm,
        title=f"Fiche élève — {eleve.nom} {eleve.prenoms}",
        author="Analyse scolaire",
    )
    document.build(
        contenu, onFirstPage=pied_de_page, onLaterPages=pied_de_page,
    )

    return tampon.getvalue()


def exporter_fiche_eleve_pdf(request):
    """Télécharge la fiche d'un élève en PDF pour l'année demandée."""

    eleve_id = request.GET.get("eleve", "")
    annee_id = request.GET.get("annee", "")

    if not eleve_id.isdigit():
        raise Http404("Élève introuvable.")

    eleve = get_object_or_404(Eleve, id=int(eleve_id))

    donnees = _donnees_fiche(eleve, annee_id)

    if donnees["aucune_donnee"]:
        return redirect(f"{reverse('fiche_eleve')}?eleve={eleve.id}")

    try:
        pdf = _generer_pdf_fiche(eleve, donnees)
    except ImportError:
        return HttpResponse(
            "La bibliothèque reportlab n'est pas installée. "
            "Installez-la avec : pip install reportlab",
            status=501,
            content_type="text/plain; charset=utf-8",
        )

    nom_fichier = re.sub(
        r"[^A-Za-z0-9_-]+", "_",
        f"fiche_{_normaliser(eleve.nom)}_{eleve.matricule}_{donnees['annee_choisie'].nom}",
    ).strip("_")

    reponse = HttpResponse(pdf, content_type="application/pdf")
    reponse["Content-Disposition"] = (
        f'attachment; filename="{nom_fichier}.pdf"'
    )

    return reponse


# =====================================================================
# BILAN DE FIN D'ANNÉE
# =====================================================================
#
# Moyenne générale annuelle = (T1 x 1 + T2 x 2 + T3 x 2) / 5.
# Moyenne annuelle d'une matière = (T1 + T2 + T3) / 3 (coefficients 1-1-1).
# Décision de fin d'année, d'après la moyenne générale annuelle arrondie
# à 2 décimales :
#   - 10 ou plus          : passe en classe supérieure ;
#   - de 8,50 à 9,99      : reprend la classe (redoublement) ;
#   - moins de 8,50       : exclu.
#
# Le bilan n'est disponible que pour les années dont les moyennes des
# trois trimestres sont enregistrées.

COEFFICIENTS_TRIMESTRES = {"T1": 1, "T2": 2, "T3": 2}

SEUIL_PASSAGE = Decimal("10")
SEUIL_REDOUBLEMENT = Decimal("8.5")

DECISIONS = {
    "passe": "Passe en classe supérieure",
    "redouble": "Reprend la classe",
    "exclu": "Exclu",
}


def _ids_annees_completes():
    """Identifiants des années scolaires dont les trois trimestres ont
    des moyennes enregistrées (une seule requête)."""

    lignes = (
        Resultat.objects
        .filter(trimestre__in=ORDRE_TRIMESTRES, moyenne_trimestrielle__gt=0)
        .order_by()
        .values("annee_scolaire_id")
        .annotate(nb=Count("trimestre", distinct=True))
        .filter(nb=len(ORDRE_TRIMESTRES))
    )

    return [ligne["annee_scolaire_id"] for ligne in lignes]


def annees_bilan_disponibles():
    """Années scolaires pour lesquelles le bilan peut être établi,
    de la plus récente à la plus ancienne."""

    ids = _ids_annees_completes()

    if not ids:
        return []

    return list(AnneeScolaire.objects.filter(id__in=ids).order_by("-nom"))


def _moyenne_ponderee(valeurs_par_trimestre):
    """Moyenne pondérée (T1 x 1, T2 x 2, T3 x 2) arrondie à 2 décimales.
    Si une valeur manque, elle est ignorée et les coefficients restants
    servent de base. Renvoie None s'il n'y a aucune valeur."""

    total = Decimal("0")
    poids = 0

    for trim, coefficient in COEFFICIENTS_TRIMESTRES.items():
        valeur = valeurs_par_trimestre.get(trim)

        if valeur is None:
            continue

        total += _decimal(valeur) * coefficient
        poids += coefficient

    if poids == 0:
        return None

    return (total / poids).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def _decision_fin_annee(moyenne_annuelle):
    if moyenne_annuelle >= SEUIL_PASSAGE:
        return "passe"

    if moyenne_annuelle >= SEUIL_REDOUBLEMENT:
        return "redouble"

    return "exclu"


def _moyenne_simple(valeurs):
    valeurs = [v for v in valeurs if v is not None]

    if not valeurs:
        return None

    return (sum(valeurs) / len(valeurs)).quantize(
        Decimal("0.01"), rounding=ROUND_HALF_UP
    )


def _construire_bilan(annee):
    """Bilan de tous les élèves d'une année scolaire. Renvoie
    (lignes, incomplets, noms_matieres) :
      - lignes : élèves ayant les trois moyennes trimestrielles ;
      - incomplets : élèves à qui il manque au moins un trimestre ;
      - noms_matieres : matières ayant au moins une moyenne annuelle."""

    resultats = (
        Resultat.objects
        .filter(annee_scolaire=annee, trimestre__in=ORDRE_TRIMESTRES)
        .select_related("eleve", "classe")
    )

    par_eleve = {}

    for resultat in resultats:
        groupe = par_eleve.setdefault(
            resultat.eleve_id,
            {"eleve": resultat.eleve, "trimestres": {}},
        )
        groupe["trimestres"][resultat.trimestre] = resultat

    lignes = []
    incomplets = []

    for groupe in par_eleve.values():

        trimestres = groupe["trimestres"]

        moyennes = {
            trim: _moyenne_classee(trimestres.get(trim))
            for trim in ORDRE_TRIMESTRES
        }

        # La classe retenue est celle du trimestre le plus récent.
        dernier = next(
            trimestres[t] for t in reversed(ORDRE_TRIMESTRES)
            if t in trimestres
        )

        manquants = [t for t, m in moyennes.items() if m is None]

        if manquants:
            incomplets.append({
                "eleve": groupe["eleve"],
                "classe": dernier.classe,
                "manquants": manquants,
            })
            continue

        moyenne_annuelle = _moyenne_ponderee(moyennes)

        matieres = {}

        # Matières : moyenne simple des trois trimestres (1-1-1).
        for nom, champ in MATIERES.items():
            matieres[nom] = _moyenne_simple([
                None if getattr(trimestres[trim], champ) is None
                else _decimal(getattr(trimestres[trim], champ))
                for trim in ORDRE_TRIMESTRES
            ])

        decision = _decision_fin_annee(moyenne_annuelle)

        lignes.append({
            "eleve": groupe["eleve"],
            "classe": dernier.classe,
            "moyennes": [moyennes[t] for t in ORDRE_TRIMESTRES],
            "moyenne_annuelle": moyenne_annuelle,
            "matieres_par_nom": matieres,
            "decision": decision,
            "decision_libelle": DECISIONS[decision],
        })

    # ---- Rang annuel dans la classe (ex æquo : même rang) ----

    par_classe = {}

    for ligne in lignes:
        par_classe.setdefault(ligne["classe"].id, []).append(ligne)

    for groupe in par_classe.values():
        groupe.sort(key=lambda l: l["moyenne_annuelle"], reverse=True)

        rang = 0
        precedente = None

        for position, ligne in enumerate(groupe, start=1):
            if ligne["moyenne_annuelle"] != precedente:
                rang = position
                precedente = ligne["moyenne_annuelle"]

            ligne["rang"] = rang
            ligne["effectif"] = len(groupe)

    # ---- Colonnes de matières à afficher ----

    noms_matieres = [
        nom for nom in MATIERES
        if any(l["matieres_par_nom"][nom] is not None for l in lignes)
    ]

    for ligne in lignes:
        ligne["matieres"] = [
            ligne["matieres_par_nom"][nom] for nom in noms_matieres
        ]

        # Détail affiché au clic sur « Voir » : une pastille par matière.
        ligne["details_matieres"] = [
            {"nom": nom, "note": ligne["matieres_par_nom"][nom]}
            for nom in noms_matieres
            if ligne["matieres_par_nom"][nom] is not None
        ]

    lignes.sort(key=lambda l: (l["classe"].nom, -l["moyenne_annuelle"]))

    return lignes, incomplets, noms_matieres


def _donnees_bilan(parametres):
    """Calcule le bilan de fin d'année en appliquant les filtres de la
    page (année, niveau, classe, décision, recherche). Sert à la page,
    à l'export Excel et à l'export PDF : les trois affichent donc
    exactement les mêmes élèves.

    Renvoie {"indisponible": True, ...} tant que les trois trimestres
    ne sont pas enregistrés."""

    annees = annees_bilan_disponibles()

    derniere_annee = AnneeScolaire.objects.order_by("-nom").first()

    # ---- Bilan indisponible : il manque des trimestres ----

    annee_en_attente = None

    if derniere_annee and derniere_annee.id not in [a.id for a in annees]:
        presents = set(
            Resultat.objects
            .filter(annee_scolaire=derniere_annee,
                    moyenne_trimestrielle__gt=0)
            .values_list("trimestre", flat=True)
        )

        annee_en_attente = {
            "annee": derniere_annee,
            "presents": [t for t in ORDRE_TRIMESTRES if t in presents],
            "manquants": [t for t in ORDRE_TRIMESTRES if t not in presents],
        }

    if not annees:
        return {"indisponible": True, "annee_en_attente": annee_en_attente}

    # ---- Filtres ----

    annee_id = parametres.get("annee", "")

    annee = next(
        (a for a in annees if str(a.id) == annee_id), annees[0]
    )

    niveau = parametres.get("niveau", "")
    classe_id = parametres.get("classe", "")
    decision_filtre = parametres.get("decision", "")
    recherche = parametres.get("recherche", "").strip()

    lignes, incomplets, noms_matieres = _construire_bilan(annee)

    # ---- Niveaux disponibles (triés : 6e, 5e, ... Terminale) ----

    def niveau_de(classe):
        return classe.niveau or "Non renseigné"

    niveaux = sorted(
        {niveau_de(l["classe"]) for l in lignes + incomplets},
        key=lambda n: _info_tri_niveau(n)[1],
    )

    if niveau in niveaux:
        lignes = [l for l in lignes if niveau_de(l["classe"]) == niveau]
        incomplets = [
            i for i in incomplets if niveau_de(i["classe"]) == niveau
        ]
    else:
        niveau = ""

    # Les classes proposées suivent le niveau choisi.
    classes = sorted(
        {l["classe"].id: l["classe"] for l in lignes + incomplets}.values(),
        key=lambda c: c.nom,
    )

    if classe_id.isdigit() and int(classe_id) in [c.id for c in classes]:
        classe_id_int = int(classe_id)
        lignes = [l for l in lignes if l["classe"].id == classe_id_int]
        incomplets = [i for i in incomplets if i["classe"].id == classe_id_int]
    else:
        classe_id = ""

    # ---- Synthèse (sur la classe choisie, avant les autres filtres) ----

    total = len(lignes)

    synthese = {"total": total, "incomplets": len(incomplets)}

    for code in DECISIONS:
        nombre = sum(1 for l in lignes if l["decision"] == code)
        synthese[code] = nombre
        synthese[f"{code}_pct"] = (
            round(nombre / total * 100, 1) if total else 0
        )

    pied = {
        "moyenne_annuelle": _moyenne_simple(
            [l["moyenne_annuelle"] for l in lignes]
        ),
    }

    # ---- Filtres d'affichage ----

    if decision_filtre in DECISIONS:
        lignes = [l for l in lignes if l["decision"] == decision_filtre]
    else:
        decision_filtre = ""

    if recherche:
        mots = [_normaliser(m) for m in recherche.split()]

        def correspond(eleve):
            texte = _normaliser(
                f"{eleve.nom} {eleve.prenoms} {eleve.matricule}"
            )
            return all(mot in texte for mot in mots)

        lignes = [l for l in lignes if correspond(l["eleve"])]

    return {
        "indisponible": False,
        "annee_en_attente": annee_en_attente,
        "annees": annees,
        "annee": annee,
        "niveaux": niveaux,
        "niveau": niveau,
        "classes": classes,
        "classe_id": classe_id,
        "decision_filtre": decision_filtre,
        "recherche": recherche,
        "lignes": lignes,
        "incomplets": incomplets,
        "noms_matieres": noms_matieres,
        "synthese": synthese,
        "pied": pied,
    }


def bilan_annuel(request):
    """Bilan de fin d'année : moyennes annuelles par matière, moyenne
    générale annuelle pondérée et décision de fin d'année."""

    donnees = _donnees_bilan(request.GET)

    if donnees["indisponible"]:
        return render(request, "analyse/bilan_annuel.html", donnees)

    page = Paginator(donnees["lignes"], 50).get_page(request.GET.get("page"))

    parametres = request.GET.copy()
    parametres.pop("page", None)

    contexte = dict(donnees)
    del contexte["lignes"]

    contexte.update({
        "decisions": DECISIONS,
        "page": page,
        "nombre_filtre": len(donnees["lignes"]),
        "querystring": parametres.urlencode(),
        "seuil_passage": SEUIL_PASSAGE,
        "seuil_redoublement": SEUIL_REDOUBLEMENT,
    })

    return render(request, "analyse/bilan_annuel.html", contexte)


# ---------------------------------------------------------------------
# Exports du bilan (Excel et PDF), selon les filtres de la page
# ---------------------------------------------------------------------

def _libelle_filtres_bilan(donnees):
    """Filtres actifs sous forme de phrases courtes, pour les en-têtes
    des exports."""

    parties = [f"Année scolaire {donnees['annee'].nom}"]

    if donnees["niveau"]:
        parties.append(f"Niveau {donnees['niveau']}")

    if donnees["classe_id"]:
        classe = next(
            c for c in donnees["classes"]
            if str(c.id) == donnees["classe_id"]
        )
        parties.append(f"Classe {classe.nom}")

    if donnees["decision_filtre"]:
        parties.append(
            f"Décision : {DECISIONS[donnees['decision_filtre']]}"
        )

    if donnees["recherche"]:
        parties.append(f"Recherche : « {donnees['recherche']} »")

    return parties


def _nom_fichier_bilan(donnees, extension):
    morceaux = ["bilan_annuel", str(donnees["annee"].nom)]

    if donnees["niveau"]:
        morceaux.append(donnees["niveau"])

    if donnees["classe_id"]:
        morceaux.append(next(
            c.nom for c in donnees["classes"]
            if str(c.id) == donnees["classe_id"]
        ))

    nom = re.sub(
        r"[^A-Za-z0-9_-]+", "_", _normaliser("_".join(morceaux))
    ).strip("_")

    return f"{nom}_{datetime.now():%Y%m%d_%H%M}.{extension}"


def _construire_classeur_bilan(donnees):
    """Classeur Excel du bilan : une feuille par élève (avec les
    moyennes annuelles de chaque matière), une feuille de synthèse et,
    s'il y en a, la liste des élèves sans bilan."""

    classeur = Workbook()

    # ---- Feuille 1 : bilan par élève ----

    feuille = classeur.active
    feuille.title = "Bilan annuel"

    noms_matieres = donnees["noms_matieres"]

    _ecrire_entetes(feuille, [
        "Matricule", "Nom", "Prénoms", "Classe", "Niveau",
        "Moyenne T1", "Moyenne T2", "Moyenne T3",
        "Moyenne annuelle", "Rang", "Effectif classe", "Décision",
    ] + noms_matieres)

    feuille.freeze_panes = "D2"

    for numero, ligne in enumerate(donnees["lignes"], start=2):

        valeurs = [
            ligne["eleve"].matricule,
            ligne["eleve"].nom,
            ligne["eleve"].prenoms,
            ligne["classe"].nom,
            ligne["classe"].niveau,
        ]
        valeurs += [float(m) for m in ligne["moyennes"]]
        valeurs += [
            float(ligne["moyenne_annuelle"]),
            ligne["rang"],
            ligne["effectif"],
            DECISIONS[ligne["decision"]],
        ]
        valeurs += [
            float(m) if m is not None else None for m in ligne["matieres"]
        ]

        for colonne, valeur in enumerate(valeurs, start=1):
            cellule = feuille.cell(row=numero, column=colonne, value=valeur)

            if isinstance(valeur, float):
                cellule.number_format = "0.00"

            if colonne >= 6:
                cellule.alignment = ALIGNEMENT_CENTRE

        feuille.cell(row=numero, column=9).font = Font(bold=True, name="Arial")

    _ajuster_largeurs_colonnes(feuille)

    # ---- Feuille 2 : synthèse ----

    synthese = donnees["synthese"]
    resume = classeur.create_sheet("Synthèse")

    ligne_excel = 1

    for texte in _libelle_filtres_bilan(donnees):
        resume.cell(row=ligne_excel, column=1, value=texte)
        ligne_excel += 1

    ligne_excel += 1

    resume.cell(
        row=ligne_excel, column=1,
        value="Synthèse sur le niveau ou la classe choisi(e), "
              "avant les filtres Décision et Recherche",
    ).font = Font(italic=True, name="Arial")
    ligne_excel += 2

    for colonne, texte in enumerate(["Indicateur", "Effectif", "Part"], 1):
        _cellule_entete_sombre(resume, ligne_excel, colonne, texte)

    lignes_synthese = [
        ("Élèves avec bilan complet", synthese["total"], None),
        (DECISIONS["passe"], synthese["passe"], synthese["passe_pct"]),
        (DECISIONS["redouble"], synthese["redouble"],
         synthese["redouble_pct"]),
        (DECISIONS["exclu"], synthese["exclu"], synthese["exclu_pct"]),
        ("Élèves sans bilan (trimestre manquant)",
         synthese["incomplets"], None),
    ]

    for libelle, effectif, pourcentage in lignes_synthese:
        ligne_excel += 1

        for colonne, valeur in enumerate(
            [libelle, effectif, pourcentage], start=1
        ):
            cellule = resume.cell(
                row=ligne_excel, column=colonne,
                value=(valeur / 100 if colonne == 3 and valeur is not None
                       else valeur),
            )
            cellule.border = BORDURE_FINE

            if colonne == 3 and valeur is not None:
                cellule.number_format = "0.0%"
            if colonne > 1:
                cellule.alignment = ALIGNEMENT_CENTRE

    moyenne = donnees["pied"]["moyenne_annuelle"]

    if moyenne is not None:
        ligne_excel += 2
        resume.cell(row=ligne_excel, column=1,
                    value="Moyenne générale annuelle").font = Font(
            bold=True, name="Arial")
        cellule = resume.cell(row=ligne_excel, column=2, value=float(moyenne))
        cellule.number_format = "0.00"
        cellule.alignment = ALIGNEMENT_CENTRE

    ligne_excel += 2
    regles = [
        "Règles de calcul",
        "Moyenne générale annuelle = (T1 x 1 + T2 x 2 + T3 x 2) / 5",
        "Moyenne annuelle d'une matière = (T1 + T2 + T3) / 3",
        f"Moyenne >= {SEUIL_PASSAGE} : {DECISIONS['passe']}",
        f"Moyenne de {SEUIL_REDOUBLEMENT} à 9,99 : {DECISIONS['redouble']}",
        f"Moyenne < {SEUIL_REDOUBLEMENT} : {DECISIONS['exclu']}",
    ]
    for indice, texte in enumerate(regles):
        cellule = resume.cell(row=ligne_excel + indice, column=1, value=texte)
        if indice == 0:
            cellule.font = Font(bold=True, name="Arial")

    resume.column_dimensions["A"].width = 52
    resume.column_dimensions["B"].width = 12
    resume.column_dimensions["C"].width = 10

    # ---- Feuille 3 : élèves sans bilan ----

    if donnees["incomplets"]:
        sans_bilan = classeur.create_sheet("Sans bilan")

        _ecrire_entetes(sans_bilan, [
            "Matricule", "Nom", "Prénoms", "Classe", "Trimestres manquants",
        ])

        for numero, item in enumerate(donnees["incomplets"], start=2):
            for colonne, valeur in enumerate([
                item["eleve"].matricule,
                item["eleve"].nom,
                item["eleve"].prenoms,
                item["classe"].nom,
                ", ".join(item["manquants"]),
            ], start=1):
                sans_bilan.cell(row=numero, column=colonne, value=valeur)

        _ajuster_largeurs_colonnes(sans_bilan)

    tampon = io.BytesIO()
    classeur.save(tampon)

    return tampon.getvalue()


def _generer_pdf_bilan(donnees):
    """PDF du bilan (format paysage) : synthèse puis liste des élèves
    avec moyennes trimestrielles, moyenne annuelle, rang et décision.
    Les moyennes par matière figurent dans l'export Excel."""

    from xml.sax.saxutils import escape

    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import (
        Paragraph,
        SimpleDocTemplate,
        Spacer,
        Table,
        TableStyle,
    )

    ACCENT = colors.HexColor("#2f5bea")
    GRIS = colors.HexColor("#667085")
    FOND = colors.HexColor("#eef1f6")
    LIGNE = colors.HexColor("#e4e7ec")

    COULEURS_DECISION = {
        "passe": "#0b6b2f",
        "redouble": "#b54708",
        "exclu": "#b42318",
    }

    base = ParagraphStyle(
        "base", fontName="Helvetica", fontSize=9, leading=12,
        textColor=colors.HexColor("#1d2433"),
    )
    centre = ParagraphStyle("centre", parent=base, alignment=1)
    petit = ParagraphStyle(
        "petit", parent=base, fontSize=8.5, leading=11, textColor=GRIS,
    )

    def p(texte, style=base):
        return Paragraph(texte, style)

    def nombre(valeur):
        return f"{valeur:.2f}".replace(".", ",")

    def moyenne_texte(valeur, gras=False):
        couleur = "#b42318" if valeur < SEUIL_PASSAGE else "#1d2433"
        texte = nombre(valeur)
        if gras:
            texte = f"<b>{texte}</b>"
        return f'<font color="{couleur}">{texte}</font>'

    synthese = donnees["synthese"]
    moyenne = donnees["pied"]["moyenne_annuelle"]

    contenu = [
        p("BILAN DE FIN D'ANNÉE", ParagraphStyle(
            "sur", parent=petit, fontName="Helvetica-Bold",
            textColor=ACCENT)),
        p(f"Année scolaire {escape(str(donnees['annee'].nom))}",
          ParagraphStyle("titre", parent=base, fontName="Helvetica-Bold",
                         fontSize=18, leading=22)),
        Spacer(1, 3),
        p(" &nbsp;·&nbsp; ".join(
            escape(x) for x in _libelle_filtres_bilan(donnees)[1:]
        ) or "Tous les niveaux et toutes les classes", petit),
        Spacer(1, 10),
    ]

    # ---- Synthèse ----

    cartes = [
        ("Élèves avec bilan complet", str(synthese["total"])),
        (DECISIONS["passe"],
         f"{synthese['passe']}  ({synthese['passe_pct']:.1f} %)".replace(
             ".", ",")),
        (DECISIONS["redouble"],
         f"{synthese['redouble']}  ({synthese['redouble_pct']:.1f} %)"
         .replace(".", ",")),
        (DECISIONS["exclu"],
         f"{synthese['exclu']}  ({synthese['exclu_pct']:.1f} %)".replace(
             ".", ",")),
        ("Moyenne générale annuelle",
         nombre(moyenne) if moyenne is not None else "—"),
    ]

    resume = Table(
        [[
            [p(f'<font color="#667085" size="8">{escape(lib)}</font>'),
             p(f"<b>{escape(val)}</b>", ParagraphStyle(
                 "val", parent=base, fontSize=12, leading=15))]
            for lib, val in cartes
        ]],
        colWidths=[5.2 * cm] * 5,
    )
    resume.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.5, LIGNE),
        ("INNERGRID", (0, 0), (-1, -1), 0.5, LIGNE),
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#f8f9fb")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
        ("LEFTPADDING", (0, 0), (-1, -1), 9),
    ]))
    contenu.append(resume)
    contenu.append(Spacer(1, 4))
    contenu.append(p(
        "Moyenne annuelle = (T1 × 1 + T2 × 2 + T3 × 2) ÷ 5. "
        f"Passage à partir de {SEUIL_PASSAGE} ; de "
        f"{str(SEUIL_REDOUBLEMENT).replace('.', ',')} à 9,99 : reprend la "
        f"classe ; en dessous de {str(SEUIL_REDOUBLEMENT).replace('.', ',')} "
        ": exclu.", petit,
    ))
    contenu.append(Spacer(1, 12))

    # ---- Liste des élèves ----

    def entete(texte, style=base):
        return p(f"<b>{texte}</b>", ParagraphStyle(
            "th", parent=style, fontSize=8.5, textColor=GRIS))

    lignes = [[
        entete("N°", centre), entete("Élève"), entete("Classe"),
        entete("T1", centre), entete("T2", centre), entete("T3", centre),
        entete("Moy. annuelle", centre), entete("Rang", centre),
        entete("Décision"),
    ]]

    for numero, ligne in enumerate(donnees["lignes"], start=1):
        rang = "1er" if ligne["rang"] == 1 else f"{ligne['rang']}e"

        lignes.append([
            p(str(numero), centre),
            p(f"<b>{escape(ligne['eleve'].nom)}</b> "
              f"{escape(ligne['eleve'].prenoms)}"),
            p(escape(ligne["classe"].nom)),
            *[p(moyenne_texte(m), centre) for m in ligne["moyennes"]],
            p(moyenne_texte(ligne["moyenne_annuelle"], gras=True), centre),
            p(f'{rang} <font color="#667085" size="8">'
              f'/ {ligne["effectif"]}</font>', centre),
            p(f'<font color="{COULEURS_DECISION[ligne["decision"]]}">'
              f'<b>{escape(DECISIONS[ligne["decision"]])}</b></font>'),
        ])

    if len(lignes) == 1:
        contenu.append(p("Aucun élève ne correspond à ces filtres."))
    else:
        facteur = 26.0 / 26.6
        largeurs = [1.1, 6.8, 2.6, 2.0, 2.0, 2.0, 2.6, 2.6, 4.9]

        tableau = Table(
            lignes, repeatRows=1,
            colWidths=[l * cm * facteur for l in largeurs],
        )
        tableau.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), FOND),
            ("LINEBELOW", (0, 0), (-1, -1), 0.5, LIGNE),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("LEFTPADDING", (0, 0), (-1, -1), 6),
            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ]))
        contenu.append(tableau)

    if donnees["incomplets"]:
        contenu.append(Spacer(1, 8))
        contenu.append(p(
            f"{len(donnees['incomplets'])} élève(s) sans bilan : "
            "au moins un trimestre manque (liste dans l'export Excel).",
            petit,
        ))

    contenu.append(Spacer(1, 4))
    contenu.append(p(
        "Les moyennes annuelles par matière figurent dans l'export Excel.",
        petit,
    ))

    def pied_de_page(canvas, doc):
        largeur, _ = landscape(A4)
        canvas.saveState()
        canvas.setStrokeColor(LIGNE)
        canvas.line(1.5 * cm, 1.4 * cm, largeur - 1.5 * cm, 1.4 * cm)
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(GRIS)
        canvas.drawString(
            1.5 * cm, 0.9 * cm,
            f"Analyse scolaire — bilan {donnees['annee'].nom} — "
            f"généré le {datetime.now():%d/%m/%Y}",
        )
        canvas.drawRightString(largeur - 1.5 * cm, 0.9 * cm,
                               f"Page {doc.page}")
        canvas.restoreState()

    tampon = io.BytesIO()

    document = SimpleDocTemplate(
        tampon, pagesize=landscape(A4),
        leftMargin=1.5 * cm, rightMargin=1.5 * cm,
        topMargin=1.5 * cm, bottomMargin=2 * cm,
        title=f"Bilan de fin d'année {donnees['annee'].nom}",
        author="Analyse scolaire",
    )
    document.build(
        contenu, onFirstPage=pied_de_page, onLaterPages=pied_de_page,
    )

    return tampon.getvalue()


def exporter_bilan_annuel_excel(request):
    """Télécharge le bilan de fin d'année en Excel, avec les mêmes
    filtres que la page."""

    donnees = _donnees_bilan(request.GET)

    if donnees["indisponible"]:
        return redirect("bilan_annuel")

    reponse = HttpResponse(
        _construire_classeur_bilan(donnees),
        content_type=(
            "application/vnd.openxmlformats-officedocument"
            ".spreadsheetml.sheet"
        ),
    )
    reponse["Content-Disposition"] = (
        f'attachment; filename="{_nom_fichier_bilan(donnees, "xlsx")}"'
    )

    return reponse


def exporter_bilan_annuel_pdf(request):
    """Télécharge le bilan de fin d'année en PDF, avec les mêmes
    filtres que la page."""

    donnees = _donnees_bilan(request.GET)

    if donnees["indisponible"]:
        return redirect("bilan_annuel")

    try:
        pdf = _generer_pdf_bilan(donnees)
    except ImportError:
        return HttpResponse(
            "La bibliothèque reportlab n'est pas installée. "
            "Installez-la avec : pip install reportlab",
            status=501,
            content_type="text/plain; charset=utf-8",
        )

    reponse = HttpResponse(pdf, content_type="application/pdf")
    reponse["Content-Disposition"] = (
        f'attachment; filename="{_nom_fichier_bilan(donnees, "pdf")}"'
    )

    return reponse


# =====================================================================
# BILAN DE FIN DE TRIMESTRE
# =====================================================================
#
# Pour une année et un trimestre : moyenne de chaque élève classé, rang,
# évolution par rapport au trimestre précédent, situation d'après la
# moyenne, synthèse par classe et moyennes par matière (bouton « Voir »).
# Les élèves sans moyenne (0 ou absente) sont comptés « non classés ».
#
# Les trois situations reprennent les tranches du tableau « Statistiques » :
#   - 10 ou plus ; de 8,50 à moins de 10 ; moins de 8,50.

SITUATIONS = {
    "reussite": "Moyenne de 10 ou plus",
    "surveiller": "De 8,50 à moins de 10",
    "difficulte": "Moins de 8,50",
}


def _situation_moyenne(moyenne):
    if moyenne >= SEUIL_PASSAGE:
        return "reussite"

    if moyenne >= SEUIL_REDOUBLEMENT:
        return "surveiller"

    return "difficulte"


def _attacher_evolution_par_lots(resultats, taille=400):
    """Comme _attacher_evolution, mais par lots : évite les requêtes
    trop volumineuses quand on traite tout un établissement."""

    resultats = list(resultats)

    for debut in range(0, len(resultats), taille):
        _attacher_evolution(resultats[debut:debut + taille])


def _annees_bilan_trimestre():
    """Années scolaires ayant au moins une moyenne trimestrielle,
    de la plus récente à la plus ancienne."""

    ids = set(
        Resultat.objects
        .filter(moyenne_trimestrielle__gt=0)
        .values_list("annee_scolaire_id", flat=True)
        .distinct()
    )

    if not ids:
        return []

    return list(AnneeScolaire.objects.filter(id__in=ids).order_by("-nom"))


def _construire_bilan_trimestre(annee, trimestre):
    """Renvoie (lignes, non_classes, noms_matieres) pour une année et
    un trimestre :
      - lignes : élèves classés (moyenne > 0), avec rang, évolution,
        situation, performance et détail des matières ;
      - non_classes : élèves de la période sans moyenne ;
      - noms_matieres : matières ayant au moins une note."""

    resultats = list(
        Resultat.objects
        .filter(annee_scolaire=annee, trimestre=trimestre)
        .select_related("eleve", "classe")
    )

    classes_effectifs = {}
    for resultat in resultats:
        classes_effectifs[resultat.classe_id] = (
            classes_effectifs.get(resultat.classe_id, 0) + 1
        )

    classes_resultats = [r for r in resultats if _moyenne_classee(r) is not None]
    non_classes = [
        {"eleve": r.eleve, "classe": r.classe}
        for r in resultats if _moyenne_classee(r) is None
    ]

    for resultat in classes_resultats:
        _attacher_details_matieres(resultat)
        _definir_performance(resultat)

    _attacher_evolution_par_lots(classes_resultats)

    lignes = []

    for resultat in classes_resultats:
        moyenne = _moyenne_classee(resultat)

        lignes.append({
            "eleve": resultat.eleve,
            "classe": resultat.classe,
            "moyenne": moyenne,
            "rang": _rang_entier(resultat.rang),
            "effectif": classes_effectifs[resultat.classe_id],
            "situation": _situation_moyenne(moyenne),
            "situation_libelle": SITUATIONS[_situation_moyenne(moyenne)],
            "evolution": resultat.evolution,
            "evolution_sens": resultat.evolution_sens,
            "evolution_ref": resultat.evolution_ref,
            "evolution_precedente": resultat.evolution_precedente,
            "performance": resultat.performance,
            "performance_css": resultat.performance_css,
            "details_matieres": resultat.details_matieres,
        })

    # Rang manquant dans le fichier : on le calcule dans la classe
    # (ex æquo : même rang).
    par_classe = {}
    for ligne in lignes:
        par_classe.setdefault(ligne["classe"].id, []).append(ligne)

    for groupe in par_classe.values():
        groupe.sort(key=lambda l: l["moyenne"], reverse=True)

        rang = 0
        precedente = None

        for position, ligne in enumerate(groupe, start=1):
            if ligne["moyenne"] != precedente:
                rang = position
                precedente = ligne["moyenne"]

            if ligne["rang"] is None:
                ligne["rang"] = rang

    noms_matieres = [
        nom for nom in MATIERES
        if any(
            any(d["nom"] == nom for d in l["details_matieres"])
            for l in lignes
        )
    ]

    lignes.sort(key=lambda l: (l["classe"].nom, l["rang"] or 9999, -l["moyenne"]))

    return lignes, non_classes, noms_matieres


def _synthese_bilan_trimestre(lignes, non_classes):
    total = len(lignes)

    synthese = {"total": total, "non_classes": len(non_classes)}

    for code in SITUATIONS:
        nombre = sum(1 for l in lignes if l["situation"] == code)
        synthese[code] = nombre
        synthese[f"{code}_pct"] = round(nombre / total * 100, 1) if total else 0

    moyennes = [l["moyenne"] for l in lignes]

    synthese["moyenne_generale"] = _moyenne_simple(moyennes)
    synthese["meilleure"] = max(moyennes) if moyennes else None
    synthese["plus_faible"] = min(moyennes) if moyennes else None

    return synthese


def _synthese_par_classe(lignes):
    """Une ligne par classe : effectif classé, moyenne, taux de réussite,
    meilleure et plus faible moyenne."""

    groupes = {}
    for ligne in lignes:
        groupes.setdefault(ligne["classe"].id, []).append(ligne)

    resultat = []

    for groupe in groupes.values():
        moyennes = [l["moyenne"] for l in groupe]
        reussis = sum(1 for m in moyennes if m >= SEUIL_PASSAGE)
        premier = max(groupe, key=lambda l: l["moyenne"])
        dernier = min(groupe, key=lambda l: l["moyenne"])

        resultat.append({
            "classe": groupe[0]["classe"],
            "effectif": len(groupe),
            "moyenne": _moyenne_simple(moyennes),
            "taux_reussite": round(reussis / len(groupe) * 100, 1),
            "meilleur": premier,
            "plus_faible": dernier,
        })

    resultat.sort(key=lambda g: g["classe"].nom)

    return resultat


def _donnees_bilan_trimestre(parametres):
    """Calcule le bilan de fin de trimestre en appliquant les filtres de
    la page (année, trimestre, niveau, classe, situation, recherche).
    Sert à la page et aux exports Excel et PDF."""

    annees = _annees_bilan_trimestre()

    if not annees:
        return {"indisponible": True}

    # ---- Année et trimestre ----

    annee_id = parametres.get("annee", "")
    annee = next((a for a in annees if str(a.id) == annee_id), annees[0])

    presents = set(
        Resultat.objects
        .filter(annee_scolaire=annee, moyenne_trimestrielle__gt=0)
        .values_list("trimestre", flat=True)
    )

    trimestres = [t for t in ORDRE_TRIMESTRES if t in presents]

    trimestre = parametres.get("trimestre", "")
    if trimestre not in trimestres:
        trimestre = trimestres[-1]

    lignes, non_classes, noms_matieres = _construire_bilan_trimestre(
        annee, trimestre
    )

    # ---- Niveaux et classes ----

    niveau = parametres.get("niveau", "")
    classe_id = parametres.get("classe", "")
    situation_filtre = parametres.get("situation", "")
    recherche = parametres.get("recherche", "").strip()

    def niveau_de(classe):
        return classe.niveau or "Non renseigné"

    tous = lignes + non_classes

    niveaux = sorted(
        {niveau_de(x["classe"]) for x in tous},
        key=lambda n: _info_tri_niveau(n)[1],
    )

    if niveau in niveaux:
        lignes = [l for l in lignes if niveau_de(l["classe"]) == niveau]
        non_classes = [
            n for n in non_classes if niveau_de(n["classe"]) == niveau
        ]
    else:
        niveau = ""

    classes = sorted(
        {x["classe"].id: x["classe"] for x in lignes + non_classes}.values(),
        key=lambda c: c.nom,
    )

    if classe_id.isdigit() and int(classe_id) in [c.id for c in classes]:
        classe_int = int(classe_id)
        lignes = [l for l in lignes if l["classe"].id == classe_int]
        non_classes = [n for n in non_classes if n["classe"].id == classe_int]
    else:
        classe_id = ""

    # ---- Synthèses (sur le niveau ou la classe choisi) ----

    synthese = _synthese_bilan_trimestre(lignes, non_classes)
    par_classe = _synthese_par_classe(lignes)

    # ---- Filtres d'affichage ----

    if situation_filtre in SITUATIONS:
        lignes = [l for l in lignes if l["situation"] == situation_filtre]
    else:
        situation_filtre = ""

    if recherche:
        mots = [_normaliser(m) for m in recherche.split()]

        def correspond(eleve):
            texte = _normaliser(
                f"{eleve.nom} {eleve.prenoms} {eleve.matricule}"
            )
            return all(mot in texte for mot in mots)

        lignes = [l for l in lignes if correspond(l["eleve"])]

    return {
        "indisponible": False,
        "annees": annees,
        "annee": annee,
        "trimestres": [
            {"code": t, "libelle": TRIMESTRE_LABELS[t]} for t in trimestres
        ],
        "trimestre": trimestre,
        "trimestre_libelle": TRIMESTRE_LABELS[trimestre],
        "niveaux": niveaux,
        "niveau": niveau,
        "classes": classes,
        "classe_id": classe_id,
        "situation_filtre": situation_filtre,
        "recherche": recherche,
        "lignes": lignes,
        "non_classes": non_classes,
        "noms_matieres": noms_matieres,
        "synthese": synthese,
        "par_classe": par_classe,
    }


def bilan_trimestre(request):
    """Bilan de fin de trimestre : moyennes, rangs, évolution, situation
    et synthèse par classe."""

    donnees = _donnees_bilan_trimestre(request.GET)

    if donnees["indisponible"]:
        return render(request, "analyse/bilan_trimestre.html", donnees)

    page = Paginator(donnees["lignes"], 50).get_page(request.GET.get("page"))

    parametres = request.GET.copy()
    parametres.pop("page", None)

    contexte = dict(donnees)
    del contexte["lignes"]

    contexte.update({
        "situations": SITUATIONS,
        "page": page,
        "nombre_filtre": len(donnees["lignes"]),
        "querystring": parametres.urlencode(),
        "seuil_passage": SEUIL_PASSAGE,
        "seuil_redoublement": SEUIL_REDOUBLEMENT,
    })

    return render(request, "analyse/bilan_trimestre.html", contexte)


# ---------------------------------------------------------------------
# Exports du bilan de trimestre (Excel et PDF), selon les filtres
# ---------------------------------------------------------------------

def _libelle_filtres_bilan_trimestre(donnees):
    parties = [
        f"Année scolaire {donnees['annee'].nom}",
        donnees["trimestre_libelle"],
    ]

    if donnees["niveau"]:
        parties.append(f"Niveau {donnees['niveau']}")

    if donnees["classe_id"]:
        classe = next(
            c for c in donnees["classes"]
            if str(c.id) == donnees["classe_id"]
        )
        parties.append(f"Classe {classe.nom}")

    if donnees["situation_filtre"]:
        parties.append(f"Moyenne : {SITUATIONS[donnees['situation_filtre']]}")

    if donnees["recherche"]:
        parties.append(f"Recherche : « {donnees['recherche']} »")

    return parties


def _nom_fichier_bilan_trimestre(donnees, extension):
    morceaux = [
        "bilan", donnees["trimestre"], str(donnees["annee"].nom),
    ]

    if donnees["niveau"]:
        morceaux.append(donnees["niveau"])

    if donnees["classe_id"]:
        morceaux.append(next(
            c.nom for c in donnees["classes"]
            if str(c.id) == donnees["classe_id"]
        ))

    nom = re.sub(
        r"[^A-Za-z0-9_-]+", "_", _normaliser("_".join(morceaux))
    ).strip("_")

    return f"{nom}_{datetime.now():%Y%m%d_%H%M}.{extension}"


def _construire_classeur_bilan_trimestre(donnees):
    """Classeur Excel : bilan par élève (avec les moyennes par matière),
    synthèse par classe, synthèse générale et élèves non classés."""

    classeur = Workbook()

    # ---- Feuille 1 : bilan par élève ----
    # Colonnes : identité, moyennes par matière, puis la moyenne
    # trimestrielle et le rang. Élèves classés par ordre alphabétique.

    feuille = classeur.active
    feuille.title = "Bilan trimestre"

    noms_matieres = donnees["noms_matieres"]

    _ecrire_entetes(feuille, [
        "Matricule", "Nom", "Prénoms", "Classe", "Niveau",
    ] + noms_matieres + ["Moyenne trimestrielle", "Rang"])

    feuille.freeze_panes = "D2"

    colonne_moyenne = 5 + len(noms_matieres) + 1

    eleves_alphabetiques = sorted(
        donnees["lignes"],
        key=lambda l: (
            _normaliser(l["eleve"].nom),
            _normaliser(l["eleve"].prenoms),
        ),
    )

    for numero, ligne in enumerate(eleves_alphabetiques, start=2):

        notes = {d["nom"]: d["note"] for d in ligne["details_matieres"]}

        valeurs = [
            ligne["eleve"].matricule,
            ligne["eleve"].nom,
            ligne["eleve"].prenoms,
            ligne["classe"].nom,
            ligne["classe"].niveau,
        ]
        valeurs += [
            float(notes[nom]) if notes.get(nom) is not None else None
            for nom in noms_matieres
        ]
        valeurs += [float(ligne["moyenne"]), ligne["rang"]]

        for colonne, valeur in enumerate(valeurs, start=1):
            cellule = feuille.cell(row=numero, column=colonne, value=valeur)

            if isinstance(valeur, float):
                cellule.number_format = "0.00"

            if colonne >= 6:
                cellule.alignment = ALIGNEMENT_CENTRE

        feuille.cell(
            row=numero, column=colonne_moyenne
        ).font = Font(bold=True, name="Arial")

    _ajuster_largeurs_colonnes(feuille)

    # ---- Feuille 2 : synthèse par classe ----

    par_classe = classeur.create_sheet("Synthèse par classe")

    _ecrire_entetes(par_classe, [
        "Classe", "Niveau", "Effectif classé", "Moyenne de la classe",
        "Taux de réussite", "Meilleure moyenne", "Élève (meilleure)",
        "Plus faible moyenne", "Élève (plus faible)",
    ])

    for numero, g in enumerate(donnees["par_classe"], start=2):
        valeurs = [
            g["classe"].nom,
            g["classe"].niveau,
            g["effectif"],
            float(g["moyenne"]),
            g["taux_reussite"] / 100,
            float(g["meilleur"]["moyenne"]),
            f"{g['meilleur']['eleve'].nom} {g['meilleur']['eleve'].prenoms}",
            float(g["plus_faible"]["moyenne"]),
            f"{g['plus_faible']['eleve'].nom} {g['plus_faible']['eleve'].prenoms}",
        ]

        for colonne, valeur in enumerate(valeurs, start=1):
            cellule = par_classe.cell(row=numero, column=colonne, value=valeur)

            if colonne in (4, 6, 8):
                cellule.number_format = "0.00"
            if colonne == 5:
                cellule.number_format = "0.0%"
            if colonne >= 3 and colonne not in (7, 9):
                cellule.alignment = ALIGNEMENT_CENTRE

    _ajuster_largeurs_colonnes(par_classe)

    # ---- Feuille 3 : synthèse générale ----

    synthese = donnees["synthese"]
    resume = classeur.create_sheet("Synthèse")

    ligne_excel = 1

    for texte in _libelle_filtres_bilan_trimestre(donnees):
        resume.cell(row=ligne_excel, column=1, value=texte)
        ligne_excel += 1

    ligne_excel += 1

    resume.cell(
        row=ligne_excel, column=1,
        value="Synthèse sur le niveau ou la classe choisi(e), "
              "avant les filtres Moyenne et Recherche",
    ).font = Font(italic=True, name="Arial")
    ligne_excel += 2

    for colonne, texte in enumerate(["Indicateur", "Effectif", "Part"], 1):
        _cellule_entete_sombre(resume, ligne_excel, colonne, texte)

    lignes_synthese = [
        ("Élèves classés", synthese["total"], None),
        (SITUATIONS["reussite"], synthese["reussite"], synthese["reussite_pct"]),
        (SITUATIONS["surveiller"], synthese["surveiller"],
         synthese["surveiller_pct"]),
        (SITUATIONS["difficulte"], synthese["difficulte"],
         synthese["difficulte_pct"]),
        ("Élèves non classés", synthese["non_classes"], None),
    ]

    for libelle, effectif, pourcentage in lignes_synthese:
        ligne_excel += 1

        for colonne, valeur in enumerate(
            [libelle, effectif, pourcentage], start=1
        ):
            cellule = resume.cell(
                row=ligne_excel, column=colonne,
                value=(valeur / 100 if colonne == 3 and valeur is not None
                       else valeur),
            )
            cellule.border = BORDURE_FINE

            if colonne == 3 and valeur is not None:
                cellule.number_format = "0.0%"
            if colonne > 1:
                cellule.alignment = ALIGNEMENT_CENTRE

    ligne_excel += 2

    for libelle, valeur in [
        ("Moyenne générale", synthese["moyenne_generale"]),
        ("Meilleure moyenne", synthese["meilleure"]),
        ("Plus faible moyenne", synthese["plus_faible"]),
    ]:
        if valeur is None:
            continue

        resume.cell(row=ligne_excel, column=1, value=libelle).font = Font(
            bold=True, name="Arial")
        cellule = resume.cell(row=ligne_excel, column=2, value=float(valeur))
        cellule.number_format = "0.00"
        cellule.alignment = ALIGNEMENT_CENTRE
        ligne_excel += 1

    resume.column_dimensions["A"].width = 44
    resume.column_dimensions["B"].width = 12
    resume.column_dimensions["C"].width = 10

    # ---- Feuille 4 : non classés ----

    if donnees["non_classes"]:
        feuille_nc = classeur.create_sheet("Non classés")

        _ecrire_entetes(feuille_nc, ["Matricule", "Nom", "Prénoms", "Classe"])

        for numero, item in enumerate(donnees["non_classes"], start=2):
            for colonne, valeur in enumerate([
                item["eleve"].matricule,
                item["eleve"].nom,
                item["eleve"].prenoms,
                item["classe"].nom,
            ], start=1):
                feuille_nc.cell(row=numero, column=colonne, value=valeur)

        _ajuster_largeurs_colonnes(feuille_nc)

    tampon = io.BytesIO()
    classeur.save(tampon)

    return tampon.getvalue()


def _generer_pdf_bilan_trimestre(donnees):
    """PDF du bilan de trimestre (paysage) : synthèse, synthèse par
    classe puis liste des élèves."""

    from xml.sax.saxutils import escape

    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import (
        Paragraph,
        SimpleDocTemplate,
        Spacer,
        Table,
        TableStyle,
    )

    ACCENT = colors.HexColor("#2f5bea")
    GRIS = colors.HexColor("#667085")
    FOND = colors.HexColor("#eef1f6")
    LIGNE = colors.HexColor("#e4e7ec")

    COULEURS = {
        "reussite": "#0b6b2f",
        "surveiller": "#b54708",
        "difficulte": "#b42318",
        "hausse": "#0b6b2f",
        "baisse": "#b42318",
        "stable": "#475467",
    }

    base = ParagraphStyle(
        "base", fontName="Helvetica", fontSize=9, leading=12,
        textColor=colors.HexColor("#1d2433"),
    )
    centre = ParagraphStyle("centre", parent=base, alignment=1)
    petit = ParagraphStyle(
        "petit", parent=base, fontSize=8.5, leading=11, textColor=GRIS,
    )
    titre2 = ParagraphStyle(
        "titre2", parent=base, fontName="Helvetica-Bold", fontSize=12,
        leading=15, textColor=ACCENT, spaceBefore=14, spaceAfter=6,
    )

    def p(texte, style=base):
        return Paragraph(texte, style)

    def nombre(valeur, signe=False):
        if valeur is None:
            return "—"
        texte = f"{valeur:+.2f}" if signe else f"{valeur:.2f}"
        return texte.replace(".", ",")

    def moyenne_texte(valeur, gras=True):
        couleur = "#b42318" if valeur < SEUIL_PASSAGE else "#1d2433"
        texte = nombre(valeur)
        return f'<font color="{couleur}">{"<b>" + texte + "</b>" if gras else texte}</font>'

    def entete(texte, style=base):
        return p(f"<b>{texte}</b>", ParagraphStyle(
            "th", parent=style, fontSize=8.5, textColor=GRIS))

    def style_tableau():
        return TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), FOND),
            ("LINEBELOW", (0, 0), (-1, -1), 0.5, LIGNE),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("LEFTPADDING", (0, 0), (-1, -1), 6),
            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ])

    facteur = 26.0 / 26.6
    synthese = donnees["synthese"]

    contenu = [
        p("BILAN DE FIN DE TRIMESTRE", ParagraphStyle(
            "sur", parent=petit, fontName="Helvetica-Bold", textColor=ACCENT)),
        p(f"{escape(donnees['trimestre_libelle'])} — "
          f"{escape(str(donnees['annee'].nom))}",
          ParagraphStyle("titre", parent=base, fontName="Helvetica-Bold",
                         fontSize=18, leading=22)),
        Spacer(1, 3),
        p(" &nbsp;·&nbsp; ".join(
            escape(x) for x in _libelle_filtres_bilan_trimestre(donnees)[2:]
        ) or "Tous les niveaux et toutes les classes", petit),
        Spacer(1, 10),
    ]

    # ---- Synthèse ----

    def part(code):
        return (f"{synthese[code]}  ({synthese[code + '_pct']:.1f} %)"
                .replace(".", ","))

    cartes = [
        ("Élèves classés", f"{synthese['total']}"
         + (f"  (+ {synthese['non_classes']} non classé(s))"
            if synthese["non_classes"] else "")),
        ("Moyenne générale", nombre(synthese["moyenne_generale"])),
        (SITUATIONS["reussite"], part("reussite")),
        (SITUATIONS["surveiller"], part("surveiller")),
        (SITUATIONS["difficulte"], part("difficulte")),
    ]

    resume = Table(
        [[
            [p(f'<font color="#667085" size="8">{escape(lib)}</font>'),
             p(f"<b>{escape(val)}</b>", ParagraphStyle(
                 "val", parent=base, fontSize=11.5, leading=14))]
            for lib, val in cartes
        ]],
        colWidths=[5.2 * cm * facteur] * 5,
    )
    resume.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.5, LIGNE),
        ("INNERGRID", (0, 0), (-1, -1), 0.5, LIGNE),
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#f8f9fb")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
        ("LEFTPADDING", (0, 0), (-1, -1), 9),
    ]))
    contenu.append(resume)

    # ---- Synthèse par classe ----

    if donnees["par_classe"]:
        contenu.append(p("Synthèse par classe", titre2))

        lignes = [[
            entete("Classe"), entete("Effectif classé", centre),
            entete("Moyenne", centre), entete("Réussite", centre),
            entete("Meilleure moyenne"), entete("Plus faible moyenne"),
        ]]

        for g in donnees["par_classe"]:
            m, f = g["meilleur"], g["plus_faible"]
            lignes.append([
                p(f"<b>{escape(g['classe'].nom)}</b>"),
                p(str(g["effectif"]), centre),
                p(moyenne_texte(g["moyenne"]), centre),
                p(f"{g['taux_reussite']:.1f} %".replace(".", ","), centre),
                p(f"{nombre(m['moyenne'])} — {escape(m['eleve'].nom)}"),
                p(f"{nombre(f['moyenne'])} — {escape(f['eleve'].nom)}"),
            ])

        tableau = Table(
            lignes, repeatRows=1,
            colWidths=[l * cm * facteur for l in [3.2, 3.0, 2.8, 2.8, 7.1, 7.1]],
        )
        tableau.setStyle(style_tableau())
        contenu.append(tableau)

    # ---- Liste des élèves ----

    contenu.append(p("Élèves", titre2))

    lignes = [[
        entete("N°", centre), entete("Élève"), entete("Classe"),
        entete("Moyenne", centre), entete("Évolution", centre),
        entete("Rang", centre), entete("Situation"),
    ]]

    for numero, ligne in enumerate(donnees["lignes"], start=1):
        rang = (f"{'1er' if ligne['rang'] == 1 else str(ligne['rang']) + 'e'}"
                f' <font color="#667085" size="8">/ {ligne["effectif"]}</font>'
                if ligne["rang"] else "—")

        if ligne["evolution_sens"]:
            texte_evo = (
                "=" if ligne["evolution_sens"] == "stable"
                else nombre(ligne["evolution"], signe=True)
            )
            evolution = (
                f'<font color="{COULEURS[ligne["evolution_sens"]]}">'
                f'{texte_evo}</font> '
                f'<font color="#667085" size="8">vs {ligne["evolution_ref"]}</font>'
            )
        else:
            evolution = "—"

        lignes.append([
            p(str(numero), centre),
            p(f"<b>{escape(ligne['eleve'].nom)}</b> "
              f"{escape(ligne['eleve'].prenoms)}"),
            p(escape(ligne["classe"].nom)),
            p(moyenne_texte(ligne["moyenne"]), centre),
            p(evolution, centre),
            p(rang, centre),
            p(f'<font color="{COULEURS[ligne["situation"]]}">'
              f'<b>{escape(ligne["situation_libelle"])}</b></font>'),
        ])

    if len(lignes) == 1:
        contenu.append(p("Aucun élève ne correspond à ces filtres."))
    else:
        tableau = Table(
            lignes, repeatRows=1,
            colWidths=[l * cm * facteur for l in [1.1, 7.2, 2.8, 2.4, 3.2, 2.6, 6.7]],
        )
        tableau.setStyle(style_tableau())
        contenu.append(tableau)

    if donnees["non_classes"]:
        contenu.append(Spacer(1, 8))
        contenu.append(p(
            f"{len(donnees['non_classes'])} élève(s) non classé(s) : pas de "
            "moyenne pour ce trimestre (liste dans l'export Excel).", petit,
        ))

    contenu.append(Spacer(1, 4))
    contenu.append(p(
        "Les moyennes par matière figurent dans l'export Excel. "
        "Évolution : écart avec la moyenne du trimestre précédent.", petit,
    ))

    def pied_de_page(canvas, doc):
        largeur, _ = landscape(A4)
        canvas.saveState()
        canvas.setStrokeColor(LIGNE)
        canvas.line(1.5 * cm, 1.4 * cm, largeur - 1.5 * cm, 1.4 * cm)
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(GRIS)
        canvas.drawString(
            1.5 * cm, 0.9 * cm,
            f"Analyse scolaire — bilan {donnees['trimestre']} "
            f"{donnees['annee'].nom} — généré le {datetime.now():%d/%m/%Y}",
        )
        canvas.drawRightString(largeur - 1.5 * cm, 0.9 * cm, f"Page {doc.page}")
        canvas.restoreState()

    tampon = io.BytesIO()

    document = SimpleDocTemplate(
        tampon, pagesize=landscape(A4),
        leftMargin=1.5 * cm, rightMargin=1.5 * cm,
        topMargin=1.5 * cm, bottomMargin=2 * cm,
        title=f"Bilan {donnees['trimestre']} {donnees['annee'].nom}",
        author="Analyse scolaire",
    )
    document.build(
        contenu, onFirstPage=pied_de_page, onLaterPages=pied_de_page,
    )

    return tampon.getvalue()


def exporter_bilan_trimestre_excel(request):
    """Télécharge le bilan de fin de trimestre en Excel, avec les mêmes
    filtres que la page."""

    donnees = _donnees_bilan_trimestre(request.GET)

    if donnees["indisponible"]:
        return redirect("bilan_trimestre")

    reponse = HttpResponse(
        _construire_classeur_bilan_trimestre(donnees),
        content_type=(
            "application/vnd.openxmlformats-officedocument"
            ".spreadsheetml.sheet"
        ),
    )
    reponse["Content-Disposition"] = (
        f'attachment; filename="{_nom_fichier_bilan_trimestre(donnees, "xlsx")}"'
    )

    return reponse


def exporter_bilan_trimestre_pdf(request):
    """Télécharge le bilan de fin de trimestre en PDF, avec les mêmes
    filtres que la page."""

    donnees = _donnees_bilan_trimestre(request.GET)

    if donnees["indisponible"]:
        return redirect("bilan_trimestre")

    try:
        pdf = _generer_pdf_bilan_trimestre(donnees)
    except ImportError:
        return HttpResponse(
            "La bibliothèque reportlab n'est pas installée. "
            "Installez-la avec : pip install reportlab",
            status=501,
            content_type="text/plain; charset=utf-8",
        )

    reponse = HttpResponse(pdf, content_type="application/pdf")
    reponse["Content-Disposition"] = (
        f'attachment; filename="{_nom_fichier_bilan_trimestre(donnees, "pdf")}"'
    )

    return reponse