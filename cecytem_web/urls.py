from django.contrib import admin
from django.urls import include, path

urlpatterns = [
    path("admin/", admin.site.urls),
    path("alumnos/", include("alumnos.urls")),
    path("usuarios/", include("usuarios.urls")),
    path("reportes/", include("reportes.urls")),
    path("administracion/", include("administracion.urls")),
    path("asistencia/", include("asistencia.urls")),
]
