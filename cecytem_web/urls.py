from django.contrib import admin
from django.urls import include, path
from asistencia import views

urlpatterns = [
    path('admin/', admin.site.urls),
    path('asistencia/', views.lista_asistencia, name='lista_asistencia'),
    path("descargar/", views.descargar_excel, name="descargar_excel"),
    path("json/", views.lista_asistencia_json, name="lista_asistencia_json"),
    path("alumnos/", include("alumnos.urls")),
    path("usuarios/", include("usuarios.urls")),

]
