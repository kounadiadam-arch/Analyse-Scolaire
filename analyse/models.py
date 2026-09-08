from django.db import models


class AnneeScolaire(models.Model):
    nom = models.CharField(max_length=20, unique=True)

    def __str__(self):
        return self.nom


class Classe(models.Model):
    nom = models.CharField(max_length=100)
    niveau = models.CharField(max_length=100, blank=True)

    def __str__(self):
        return self.nom


class Eleve(models.Model):
    GENRE_CHOICES = [
        ("M", "Masculin"),
        ("F", "Féminin"),
    ]

    matricule = models.CharField(
        max_length=50,
        unique=True
    )
    nom = models.CharField(max_length=100)
    prenoms = models.CharField(max_length=150)

    genre = models.CharField(
        max_length=1,
        choices=GENRE_CHOICES,
        blank=True
    )

    date_naissance = models.DateField(
        null=True,
        blank=True
    )

    nationalite = models.CharField(
        max_length=100,
        blank=True
    )

    redoublant = models.BooleanField(
        default=False
    )

    statut = models.CharField(
        max_length=100,
        blank=True
    )

    def __str__(self):
        return f"{self.nom} {self.prenoms}"


class Resultat(models.Model):

    TRIMESTRE_CHOICES = [
        ("T1", "Trimestre 1"),
        ("T2", "Trimestre 2"),
        ("T3", "Trimestre 3"),
    ]

    eleve = models.ForeignKey(
        Eleve,
        on_delete=models.CASCADE,
        related_name="resultats"
    )

    annee_scolaire = models.ForeignKey(
        AnneeScolaire,
        on_delete=models.CASCADE,
        related_name="resultats"
    )

    classe = models.ForeignKey(
        Classe,
        on_delete=models.CASCADE,
        related_name="resultats"
    )

    trimestre = models.CharField(
        max_length=2,
        choices=TRIMESTRE_CHOICES
    )

    composition_francaise = models.DecimalField(
        max_digits=5,
        decimal_places=2,
        null=True,
        blank=True
    )

    orthographe_grammaire = models.DecimalField(
        max_digits=5,
        decimal_places=2,
        null=True,
        blank=True
    )

    expression_orale = models.DecimalField(
        max_digits=5,
        decimal_places=2,
        null=True,
        blank=True
    )

    hist_geo = models.DecimalField(
        max_digits=5,
        decimal_places=2,
        null=True,
        blank=True
    )

    espagnol = models.DecimalField(
        max_digits=5,
        decimal_places=2,
        null=True,
        blank=True
    )

    allemand = models.DecimalField(
        max_digits=5,
        decimal_places=2,
        null=True,
        blank=True
    )

    anglais = models.DecimalField(
        max_digits=5,
        decimal_places=2,
        null=True,
        blank=True
    )

    francais = models.DecimalField(
        max_digits=5,
        decimal_places=2,
        null=True,
        blank=True
    )

    conduite = models.DecimalField(
        max_digits=5,
        decimal_places=2,
        null=True,
        blank=True
    )

    edhc = models.DecimalField(
        max_digits=5,
        decimal_places=2,
        null=True,
        blank=True
    )

    eps = models.DecimalField(
        max_digits=5,
        decimal_places=2,
        null=True,
        blank=True
    )

    svt = models.DecimalField(
        max_digits=5,
        decimal_places=2,
        null=True,
        blank=True
    )

    physiques_chimie = models.DecimalField(
        max_digits=5,
        decimal_places=2,
        null=True,
        blank=True
    )

    mathematiques = models.DecimalField(
        max_digits=5,
        decimal_places=2,
        null=True,
        blank=True
    )

    philosophie = models.DecimalField(
        max_digits=5,
        decimal_places=2,
        null=True,
        blank=True
    )

    moyenne_trimestrielle = models.DecimalField(
        max_digits=5,
        decimal_places=2,
        null=True,
        blank=True
    )

    rang = models.PositiveIntegerField(
        null=True,
        blank=True
    )

    class Meta:
        constraints = [
            models.UniqueConstraint(
                fields=[
                    "eleve",
                    "annee_scolaire",
                    "trimestre"
                ],
                name="resultat_unique_par_trimestre"
            )
        ]

    def __str__(self):
        return (
            f"{self.eleve} - "
            f"{self.annee_scolaire} - "
            f"{self.trimestre}"
        )


