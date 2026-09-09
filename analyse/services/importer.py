import re
from decimal import Decimal, InvalidOperation
from typing import Any, Optional

import pandas as pd
from django.db import transaction

from analyse.models import (
    AnneeScolaire,
    Classe,
    Eleve,
    Resultat,
)


class ImportExcelError(Exception):
    """Erreur métier lors de l'import Excel."""
    pass


# ---------------------------------------------------------------------------
# Constantes
# ---------------------------------------------------------------------------

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

COLONNES_OBLIGATOIRES = [
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
] + list(COLONNES_MATIERES.keys())


# ---------------------------------------------------------------------------
# Utilitaires de nettoyage / conversion
# ---------------------------------------------------------------------------

def nettoyer_colonne(nom: Any) -> str:
    """Nettoie le nom d'une colonne Excel."""
    return " ".join(str(nom).strip().upper().split())


def nettoyer_texte(valeur: Any) -> str:
    """Retourne une chaîne propre (jamais 'nan')."""
    if pd.isna(valeur):
        return ""
    texte = str(valeur).strip()
    return "" if texte.lower() == "nan" else texte


def nettoyer_decimal(valeur: Any, colonne: str, ligne: int) -> Optional[Decimal]:
    """Convertit une valeur Excel en Decimal (ou None)."""
    if pd.isna(valeur) or str(valeur).strip() == "":
        return None

    valeur_str = str(valeur).strip().replace(",", ".")

    try:
        return Decimal(valeur_str)
    except InvalidOperation:
        raise ImportExcelError(
            f"Valeur invalide '{valeur}' dans la colonne '{colonne}', ligne {ligne}."
        )


def convertir_genre(valeur: Any) -> str:
    """Convertit le genre en 'M' ou 'F' (ou chaîne vide)."""
    texte = nettoyer_texte(valeur).upper()

    if texte in {"M", "MASCULIN", "HOMME", "MALE"}:
        return "M"
    if texte in {"F", "FEMININ", "FÉMININ", "FEMME", "FEMALE"}:
        return "F"
    return ""


def convertir_booleen(valeur: Any) -> bool:
    """Convertit une valeur en booléen (True uniquement pour les affirmations)."""
    texte = nettoyer_texte(valeur).upper()

    if texte in {"OUI", "O", "YES", "1", "TRUE", "VRAI"}:
        return True
    if texte in {"NON", "N", "NO", "0", "FALSE", "FAUX"}:
        return False
    return False  # Valeur inconnue → False (politique permissive)


def convertir_trimestre(valeur: Any) -> str:
    """Convertit le trimestre en 'T1', 'T2' ou 'T3'."""
    texte = nettoyer_texte(valeur).upper()

    if not texte:
        raise ImportExcelError("Le trimestre est obligatoire.")

    if any(x in texte for x in ("1", "PREMIER", "1ER", "PREMIÈRE")):
        return "T1"
    if any(x in texte for x in ("2", "DEUXIEME", "DEUXIÈME", "2E", "2EME")):
        return "T2"
    if any(x in texte for x in ("3", "TROISIEME", "TROISIÈME", "3E", "3EME")):
        return "T3"

    raise ImportExcelError(f"Trimestre inconnu : '{valeur}'.")


def convertir_rang(valeur: Any, ligne: int) -> Optional[int]:
    """Extrait le rang numérique (ex: '3ème' → 3)."""
    if pd.isna(valeur):
        return None

    match = re.search(r"(\d+)", str(valeur))
    if not match:
        raise ImportExcelError(f"Rang invalide '{valeur}' à la ligne {ligne}.")

    try:
        return int(match.group(1))
    except ValueError:
        raise ImportExcelError(f"Rang invalide '{valeur}' à la ligne {ligne}.")


def convertir_date(valeur: Any) -> Optional[Any]:
    """Convertit une date Excel en date Python (ou None)."""
    date = pd.to_datetime(valeur, errors="coerce", dayfirst=True)
    if pd.isna(date):
        return None
    return date.date()


# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------

def valider_colonnes(df: pd.DataFrame) -> None:
    """Vérifie la présence de toutes les colonnes obligatoires."""
    colonnes_manquantes = [
        col for col in COLONNES_OBLIGATOIRES if col not in df.columns
    ]
    if colonnes_manquantes:
        raise ImportExcelError(
            "Colonnes manquantes dans le fichier Excel : "
            + ", ".join(colonnes_manquantes)
        )


# ---------------------------------------------------------------------------
# Parsing d'une ligne
# ---------------------------------------------------------------------------

def parser_ligne(ligne: pd.Series, numero_ligne: int) -> dict:
    """Parse une ligne du DataFrame et retourne un dictionnaire prêt à l'emploi."""
    matricule = nettoyer_texte(ligne["MATRICULE"])
    if not matricule:
        raise ImportExcelError(f"Matricule manquant à la ligne {numero_ligne}.")

    annee_nom = nettoyer_texte(ligne["ANNEE SCOLAIRE"])
    if not annee_nom:
        raise ImportExcelError(f"Année scolaire manquante à la ligne {numero_ligne}.")

    classe_nom = nettoyer_texte(ligne["CLASSE"])
    if not classe_nom:
        raise ImportExcelError(f"Classe manquante à la ligne {numero_ligne}.")

    niveau = nettoyer_texte(ligne["NIVEAU"])

    # Notes des matières
    notes = {}
    for col_excel, champ in COLONNES_MATIERES.items():
        notes[champ] = nettoyer_decimal(ligne[col_excel], col_excel, numero_ligne)

    moyenne = nettoyer_decimal(
        ligne["MOYENNE TRIMESTRIELLE"],
        "MOYENNE TRIMESTRIELLE",
        numero_ligne,
    )

    return {
        "matricule": matricule,
        "annee_nom": annee_nom,
        "classe_nom": classe_nom,
        "niveau": niveau,
        "trimestre": convertir_trimestre(ligne["TRIMESTRE"]),
        "eleve": {
            "nom": nettoyer_texte(ligne["NOM"]),
            "prenoms": nettoyer_texte(ligne["PRENOMS"]),
            "genre": convertir_genre(ligne["GENRE"]),
            "date_naissance": convertir_date(ligne["DATE DE NAISSANCE"]),
            "nationalite": nettoyer_texte(ligne["NATIONALITE"]),
            "redoublant": convertir_booleen(ligne["REDOUBLANT"]),
            "statut": nettoyer_texte(ligne["STATUT ELEVE"]),
        },
        "resultat": {
            **notes,
            "moyenne_trimestrielle": moyenne,
            "rang": convertir_rang(ligne["RANG"], numero_ligne),
        },
    }


# ---------------------------------------------------------------------------
# Import principal
# ---------------------------------------------------------------------------

@transaction.atomic
def importer_fichier_excel(fichier) -> dict:
    """
    Importe un fichier Excel de résultats scolaires.

    Retourne :
        {
            "eleves": int,
            "resultats": int,
        }
    """
    try:
        df = pd.read_excel(fichier)
    except Exception as e:
        raise ImportExcelError(f"Impossible de lire le fichier Excel : {e}")

    # Nettoyage des noms de colonnes
    df.columns = [nettoyer_colonne(c) for c in df.columns]

    valider_colonnes(df)

    # Option de sécurité : limiter la taille
    MAX_LIGNES = 10_000
    if len(df) > MAX_LIGNES:
        raise ImportExcelError(
            f"Le fichier contient trop de lignes ({len(df)}). Maximum autorisé : {MAX_LIGNES}."
        )

    nombre_eleves = 0
    nombre_resultats = 0

    # Cache simple pour éviter les get_or_create répétés
    cache_annees: dict[str, AnneeScolaire] = {}
    cache_classes: dict[str, Classe] = {}

    for index, ligne in df.iterrows():
        numero_ligne = index + 2  # +2 car Excel commence à 1 + header

        data = parser_ligne(ligne, numero_ligne)

        # --- Année scolaire ---
        annee_nom = data["annee_nom"]
        if annee_nom not in cache_annees:
            annee, _ = AnneeScolaire.objects.get_or_create(nom=annee_nom)
            cache_annees[annee_nom] = annee
        annee = cache_annees[annee_nom]

        # --- Classe ---
        classe_nom = data["classe_nom"]
        niveau = data["niveau"]

        if classe_nom not in cache_classes:
            classe, created = Classe.objects.get_or_create(
                nom=classe_nom,
                defaults={"niveau": niveau},
            )
            if not created and niveau and classe.niveau != niveau:
                classe.niveau = niveau
                classe.save(update_fields=["niveau"])
            cache_classes[classe_nom] = classe
        classe = cache_classes[classe_nom]

        # --- Élève ---
        eleve, _ = Eleve.objects.update_or_create(
            matricule=data["matricule"],
            defaults=data["eleve"],
        )
        nombre_eleves += 1

        # --- Résultat ---
        Resultat.objects.update_or_create(
            eleve=eleve,
            annee_scolaire=annee,
            trimestre=data["trimestre"],
            defaults={
                "classe": classe,
                **data["resultat"],
            },
        )
        nombre_resultats += 1

    return {
        "eleves": nombre_eleves,
        "resultats": nombre_resultats,
    }