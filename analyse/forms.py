from django import forms


class ImportExcelForm(forms.Form):
    fichier = forms.FileField(
        label="Fichier Excel",
        help_text="Format accepté : .xlsx"
    )