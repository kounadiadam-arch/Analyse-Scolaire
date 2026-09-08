import pandas as pd
from decimal import Decimal, InvalidOperation
from django.db import transaction

from analyse.models import (
    AnneeScolaire,
    Classe,
    Eleve,
    Resultat,
)


class ImportExcelError(Exception):
    pass


COLONNES_MATIERES = {
    "COMPOSITION FRANCAISE": "composition_francaise",
    "ORTHOGRAPHE GRAMMAIRE": "orthographe_grammaire",
    "EXPRESSION ORALE": "expression_orale",
    "HIST-GEO": "hist_geo",
    "ESPAGNOL": "espagnol",
    "ALLEMAND": "allemand",
    "ANGLAIS": "anglais",
    "FRANCAIS": "francais",
    "CONDUITE": "conduite",
    "EDHC": "edhc",
    "EPS": "eps",
    "SVT": "svt",
    "PHYSIQUES-CHIMIE": "physiques_chimie",
    "MATHEMATIQUES": "mathematiques",
    "PHILOSOPHIE": "philosophie",
}


def nettoyer_colonne(nom):
    """Nettoie le nom d'une colonne Excel."""
    return " ".join(str(nom).strip().upper().split())


def nettoyer_decimal(valeur, colonne, ligne):
    """Convertit une valeur Excel en Decimal."""
    if pd.isna(valeur) or str(valeur).strip() == "":
        return None

    valeur = str(valeur).strip().replace(",", ".")

    try:
        return Decimal(valeur)
    except InvalidOperation:
        raise ImportExcelError(
            f"Valeur invalide '{valeur}' dans la colonne "
            f"'{colonne}', ligne {ligne}."
        )


def convertir_genre(valeur):
    if pd.isna(valeur):
        return ""

    valeur = str(valeur).strip().upper()

    if valeur in ["M", "MASCULIN","Masculin", "HOMME"]:
        return "M"

    if valeur in ["F", "FEMININ", "FÉMININ", "Féminin", "FEMME"]:
        return "F"

    return ""


def convertir_booleen(valeur):
    if pd.isna(valeur):
        return False

    valeur = str(valeur).strip().upper()

    return valeur in [
        "OUI",
        "Oui",
        "O",
        "NON",
        "Non",
        "N" , 
    ]


def convertir_trimestre(valeur):
    if pd.isna(valeur):
        raise ImportExcelError("Le trimestre est obligatoire.")

    valeur = str(valeur).strip().upper()

    if "1" in valeur or "PREMIER" in valeur or "1ER" in valeur:
        return "T1"

    if "2" in valeur or "DEUXIÈME" in valeur or "DEUXIEME" in valeur or "2E" in valeur:
        return "T2"

    if "3" in valeur or "TROISIÈME" in valeur or "TROISIEME" in valeur or "3E" in valeur:
        return "T3"

    raise ImportExcelError(
        f"Trimestre inconnu : '{valeur}'."
    )


@transaction.atomic
def importer_fichier_excel(fichier):

    try:
        df = pd.read_excel(fichier)
    except Exception as e:
        raise ImportExcelError(
            f"Impossible de lire le fichier Excel : {e}"
        )

    # Nettoyage des noms de colonnes
    df.columns = [
        nettoyer_colonne(colonne)
        for colonne in df.columns
    ]

    colonnes_obligatoires = [
        "ANNEE SCOLAIRE",
        "TRIMESTRE",
        "MATRICULE",
        "NOM",
        "PRENOMS",
        "GENRE",
        "DATE DE NAISSANCE",
        "NATIONALITE",
        "REDOUBLANT",
        "STATUT ELEVE",
        "NIVEAU",
        "CLASSE",
        "MOYENNE TRIMESTRIELLE",
        "RANG",
    ]

    # Ajouter les colonnes des matières
    colonnes_obligatoires += list(COLONNES_MATIERES.keys())

    colonnes_manquantes = [
        colonne
        for colonne in colonnes_obligatoires
        if colonne not in df.columns
    ]

    if colonnes_manquantes:
        raise ImportExcelError(
            "Colonnes manquantes dans le fichier Excel : "
            + ", ".join(colonnes_manquantes)
        )

    nombre_eleves = 0
    nombre_resultats = 0

    for index, ligne in df.iterrows():

        numero_ligne = index + 2

        matricule = str(ligne["MATRICULE"]).strip()

        if not matricule or matricule == "nan":
            raise ImportExcelError(
                f"Matricule manquant à la ligne {numero_ligne}."
            )

        annee_nom = str(ligne["ANNEE SCOLAIRE"]).strip()
        classe_nom = str(ligne["CLASSE"]).strip()
        niveau = str(ligne["NIVEAU"]).strip()

        if not annee_nom or annee_nom == "nan":
            raise ImportExcelError(
                f"Année scolaire manquante à la ligne {numero_ligne}."
            )

        if not classe_nom or classe_nom == "nan":
            raise ImportExcelError(
                f"Classe manquante à la ligne {numero_ligne}."
            )

        # Année scolaire
        annee, _ = AnneeScolaire.objects.get_or_create(
            nom=annee_nom
        )

        # Classe
        classe, created = Classe.objects.get_or_create(
            nom=classe_nom,
            defaults={
                "niveau": niveau
            }
        )

        if not created and niveau and classe.niveau != niveau:
            classe.niveau = niveau
            classe.save(update_fields=["niveau"])

        # Date de naissance
        date_naissance = pd.to_datetime(
            ligne["DATE DE NAISSANCE"],
            errors="coerce"
        )

        if pd.isna(date_naissance):
            date_naissance = None
        else:
            date_naissance = date_naissance.date()

        # Élève
        eleve, created = Eleve.objects.update_or_create(
            matricule=matricule,
            defaults={
                "nom": str(ligne["NOM"]).strip(),
                "prenoms": str(ligne["PRENOMS"]).strip(),
                "genre": convertir_genre(ligne["GENRE"]),
                "date_naissance": date_naissance,
                "nationalite": str(
                    ligne["NATIONALITE"]
                ).strip(),
                "redoublant": convertir_booleen(
                    ligne["REDOUBLANT"]
                ),
                "statut": str(
                    ligne["STATUT ELEVE"]
                ).strip(),
            }
        )

        nombre_eleves += 1

        # Trimestre
        trimestre = convertir_trimestre(
            ligne["TRIMESTRE"]
        )

        # Données du résultat
        donnees_resultat = {}

        for colonne_excel, champ_django in COLONNES_MATIERES.items():
            donnees_resultat[champ_django] = nettoyer_decimal(
                ligne[colonne_excel],
                colonne_excel,
                numero_ligne
            )

        donnees_resultat["moyenne_trimestrielle"] = nettoyer_decimal(
            ligne["MOYENNE TRIMESTRIELLE"],
            "MOYENNE TRIMESTRIELLE",
            numero_ligne
        )

        rang = ligne["RANG"]

        if pd.isna(rang):
            rang = None
        else:
            try:
                rang = int(
                    str(rang)
                    .replace("EME", "")
                    .replace("ème", "")
                    .replace("er", "")
                    .replace("ère", "")
                    .strip()
                )
            except ValueError:
                raise ImportExcelError(
                    f"Rang invalide '{rang}' à la ligne {numero_ligne}."
                )

        donnees_resultat["rang"] = rang

        # Enregistrement du résultat
        Resultat.objects.update_or_create(
            eleve=eleve,
            annee_scolaire=annee,
            trimestre=trimestre,
            defaults={
                "classe": classe,
                **donnees_resultat,
            }
        )

        nombre_resultats += 1

    return {
        "eleves": nombre_eleves,
        "resultats": nombre_resultats,
    }