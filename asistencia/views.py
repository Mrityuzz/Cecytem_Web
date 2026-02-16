from django.shortcuts import render
from django.http import FileResponse
import openpyxl
import os
from django.conf import settings

def lista_asistencia(request):
    ruta_excel = os.path.join(settings.BASE_DIR, "base_datos", "BD_Alumnos.xlsx")
    alumnos = []

    if not os.path.isfile(ruta_excel):
        return render(request, "asistencia/lista.html", {
            "alumnos": [],
            "error": "El archivo BD_Alumnos.xlsx no se encontró en la carpeta base_datos."
        })

    try:
        libro = openpyxl.load_workbook(ruta_excel, data_only=True)
        hoja = libro.active

        # Filtro dinámico por carrera (ejemplo: ?carrera=Electrónica)
        carrera_filtro = request.GET.get("carrera")

        for i in range(2, hoja.max_row + 1):
            alumno = hoja.cell(row=i, column=1).value
            num_control = hoja.cell(row=i, column=2).value
            telefono = hoja.cell(row=i, column=3).value
            carrera = hoja.cell(row=i, column=4).value
            entrada = hoja.cell(row=i, column=5).value
            salida = hoja.cell(row=i, column=6).value

            if carrera_filtro:
                if carrera and carrera_filtro.lower() in str(carrera).lower():
                    alumnos.append({
                        "alumno": alumno,
                        "num_control": num_control,
                        "telefono": telefono,
                        "carrera": carrera,
                        "entrada": entrada,
                        "salida": salida,
                    })
            else:
                alumnos.append({
                    "alumno": alumno,
                    "num_control": num_control,
                    "telefono": telefono,
                    "carrera": carrera,
                    "entrada": entrada,
                    "salida": salida,
                })

    except Exception as e:
        return render(request, "asistencia/lista.html", {
            "alumnos": [],
            "error": f"Ocurrió un error al abrir el archivo: {e}"
        })

    return render(request, "asistencia/lista.html", {
        "alumnos": alumnos,
        "carrera_filtro": carrera_filtro
    })


# Vista para descargar el Excel original
def descargar_excel(request):
    ruta_excel = os.path.join(settings.BASE_DIR, "base_datos", "BD_Alumnos.xlsx")
    return FileResponse(open(ruta_excel, "rb"), as_attachment=True, filename="BD_Alumnos.xlsx")
