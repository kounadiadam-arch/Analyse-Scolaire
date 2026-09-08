from django.contrib import admin
from .models import (
    AnneeScolaire,
    Classe,
    Eleve,
    Resultat,
)


@admin.register(AnneeScolaire)
class AnneeScolaireAdmin(admin.ModelAdmin):
    list_display = ("nom",)


@admin.register(Classe)
class ClasseAdmin(admin.ModelAdmin):
    list_display = ("nom", "niveau")


@admin.register(Eleve)
class EleveAdmin(admin.ModelAdmin):
    list_display = (
        "matricule",
        "nom",
        "prenoms",
        "genre",
        "nationalite",
        "redoublant",
    )
    search_fields = (
        "matricule",
        "nom",
        "prenoms",
    )


@admin.register(Resultat)
class ResultatAdmin(admin.ModelAdmin):
    list_display = (
        "eleve",
        "annee_scolaire",
        "classe",
        "trimestre",
        "moyenne_trimestrielle",
        "rang",
    )

    list_filter = (
        "annee_scolaire",
        "classe",
        "trimestre",
    )

    search_fields = (
        "eleve__nom",
        "eleve__prenoms",
        "eleve__matricule",
    )
