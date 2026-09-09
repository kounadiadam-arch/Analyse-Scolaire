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
]