import os
import time
from reading import read_excel_vehicular
import numpy as np
import shutil
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, NamedStyle, Border, Side
import re
import docx
from  docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

def vehicular(entregable_path):

    hojas = ["N","S","E","O"]

    ruta = entregable_path
    listas = [[] for _ in range(4)]

    ruta_vehicular = os.path.join(ruta,"7.- Informacion de Campo","Vehicular")
    files = ["Atipico",
            "Atipico (Protransito)",
            "Tipico",
            "Tipico (Protransito)",]

    for file in files:
        ruta_tipicidad = os.path.join(ruta_vehicular,file)
        list_excels = os.listdir(ruta_tipicidad)

        for excel in list_excels:
            path_excel = os.path.join(ruta_tipicidad,excel)

            if not "~$" in excel:
                if file == "Tipico":
                    listas[0].append(path_excel)
                elif file == "Atipico":
                    listas[1].append(path_excel)
                elif file == "Tipico (Protransito)":
                    listas[2].append(path_excel)
                elif file == "Atipico (Protransito)":
                    listas[3].append(path_excel)

    archivos = [[] for _ in range(2)]

    #Comparison among Tipico
    for veh2 in listas[2]: #Supervisor
        for veh0 in listas[0]: #Consultor
            nombre_veh2 = os.path.basename(veh2)
            nombre_veh0 = os.path.basename(veh0)

            if nombre_veh2 != nombre_veh0 and nombre_veh2.replace("_REV","") == nombre_veh0:
                archivos[0].append((veh2,veh0))

    #Comparison among Atipico
    for veh3 in listas[3]: #Supervisor
        for veh1 in listas[1]: #Consultor
            nombre_veh3 = os.path.basename(veh3)
            nombre_veh1 = os.path.basename(veh0)

            if nombre_veh3 != nombre_veh1 and nombre_veh3.replace("_REV","") == nombre_veh1:
                archivos[1].append(veh3,veh1)

    data_row = [[] for _ in range(2)]

    inicio_lectura = time.perf_counter()

    ###################
    # Report Creation #
    ###################

    for index, archivo in enumerate(archivos):
        if index == 0:
            print("#### Reportes Tipicos ####")
        else:
            print("#### Reportes Atipicos ####")

        for l,pair in enumerate(archivo):
            if index == 0:
                print(f"Comparando Tipicos ({l+1}/{len(archivo)})")
            else:
                print(f"Comparando Atipicos ({l+1}/{len(archivo)})")
            data_consultor,fecha,giros = read_excel_vehicular(pair[1])
            data_supervisor,_,_ = read_excel_vehicular(pair[0])

            ruta_destino = pair[0].replace("_REV.xlsm","_REP.xlsx")

            container_file = os.path.dirname(pair[1])
            _,tipicidad = os.path.split(container_file)

            resultados = []

            volumen_supervisor = 0
            volumen_consultor = 0

            # Reading Info #

            lista_primeros_valores = []
            lista_ultimos_valores = []

            for data_con, data_sup in zip(data_consultor, data_supervisor):
                mascara = data_sup != 0
                resultado = np.zeros_like(data_con)
                
                volumen_supervisor += np.sum(data_sup[mascara])
                volumen_consultor += np.sum(data_con[mascara])

                resultado[mascara] = np.abs((data_con[mascara]-data_sup[mascara])) / data_sup[mascara]
                resultados.append(resultado)

                #Searching upper and lower limits
                filled_rows = np.any(mascara,axis=1)
                indices = np.where(filled_rows)[0]

                if indices.size > 0:
                    lista_primeros_valores.append(indices[0])
                    lista_ultimos_valores.append(indices[-1])

            if lista_primeros_valores and lista_ultimos_valores:
                lower_index = max(lista_ultimos_valores)
                upper_index = min(lista_primeros_valores)

            upper_hour = ((upper_index)*15)//60
            lower_hour = ((lower_index+1)*15)//60
            
            upper_min = ((upper_index)*15)%60
            lower_min = ((lower_index+1)*15)%60

            sample_hour = f"{upper_hour:02d}:{upper_min:02d} - {lower_hour:02d}:{lower_min:02d}"

            ###################
            # Getting Indexes #
            ###################

            suma_total = 0
            num_valores = 0
            conteo_mayor_10 = 0
            conteo_menor_10 = 0
            error_minimo = 1
            error_maximo = 0

            for result in resultados:
                mascara = result != 0
                suma_total += np.sum(result[mascara])

                num_valores += np.count_nonzero(mascara)
                conteo_mayor_10 += np.count_nonzero(result[mascara] > 0.1)
                conteo_menor_10 += np.count_nonzero(result[mascara] <= 0.1)

                if result[mascara].size > 0:
                    error_maximo_actual = np.max(result[mascara])
                    error_minimo_actual = np.min(result[mascara])
                    error_maximo = max(error_maximo,error_maximo_actual)
                    error_minimo = min(error_minimo, error_minimo_actual)

            promedio = round(suma_total / num_valores, 2)
            muestra = conteo_mayor_10 + conteo_menor_10
            interseccion_critica = round(conteo_mayor_10/muestra,2)

            path_formato = "./tools/Formato_Vehicular.xlsx"
            current_file, name_excel = os.path.split(ruta_destino)
            destiny_file = os.path.join(current_file,'Reportes')
            final_route = os.path.join(destiny_file,name_excel)
            if not os.path.exists(destiny_file):
                os.makedirs(destiny_file)
            shutil.copyfile(path_formato,final_route)
            wb = load_workbook(final_route)

            # Formats #
            porcentaje_style = NamedStyle(name="porcentaje")
            porcentaje_style.number_format = '0%'
            porcentaje_style.border = Border(left=Side(style='thin'),
                                            right=Side(style='thin'),
                                            top=Side(style='thin'),
                                            bottom=Side(style='thin'),)

            for hoja, giro in zip(hojas,giros):
                ws = wb[hoja]
                for p, valor in enumerate(giro):
                    celda = ws.cell(row=15, column=p+11)
                    celda.value = valor

            for hoja, result in zip(hojas,resultados):
                ws = wb[hoja]
                for k, fila in enumerate(result):
                    for l, valor in enumerate(fila):
                        celda = ws.cell(row=k+16, column=l+11)

                        if valor == 0: celda.value = None
                        else: celda.value = valor

                        celda.style = porcentaje_style

                        if valor >= 0.1: #Error Relativo: 10%
                            relleno = PatternFill(start_color="FFFF0000", end_color="FFFF0000",fill_type="solid")
                            celda.fill = relleno

            wb.save(final_route)

            _, nombre_archivo = os.path.split(ruta_destino)
            patron = r"^([A-Z]+-[0-9]+)"
            coincidencia = re.search(patron, nombre_archivo)

            if coincidencia:
                codigo = coincidencia.group(1)
            else:
                print(f"El archivo de excel no posee Código de intersección:\n{nombre_archivo}")
                codigo = "ERROR"

            data_row[index].append((codigo,fecha[:-9],tipicidad,
                                    sample_hour,muestra,
                                    conteo_menor_10,conteo_mayor_10,
                                    error_maximo,error_minimo,
                                    promedio,interseccion_critica))
            
    fin_lectura = time.perf_counter()
    duracion_lectura = fin_lectura - inicio_lectura
    print(f"Processing Time in Reading: {duracion_lectura:.2f} seconds.")

    #########################
    # Creating Word Report #
    #########################

    inicio_word = time.perf_counter()

    doc = docx.Document()
    doc.add_heading(f"REPORTE FLUJO VEHICULAR")
    table = doc.add_table(rows=1,cols=11,style="Table Grid")
    section = doc.sections[0]
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)

    for cell in table.columns[1].cells:
        cell.width = Inches(2)

    for cell in table.columns[3].cells:
        cell.width = Inches(2)

    titulos = ['ID', 'Fecha Muestra', 'Esc.', 'Hora Muestra', 'Muestra', 
                '<10%', '>10%', 'E.R.\n%(máx)', 'E.R.\n%(mín)', 
                'E.R.\nProm', '>10/mues.\n(%)']

    for i, titulo in enumerate(titulos):
        celda = table.rows[0].cells[i]
        parrafo = celda.paragraphs[0]
        run = parrafo.add_run(titulo)
        run.bold = True
        run.font.size = Pt(10)

    for i, data in enumerate(data_row):
        if len(data) == 0:
            pass
        
        for dt in data:
            row = table.add_row().cells
            row[0].text = dt[0] #Intersección
            row[1].text = dt[1] #Fecha
            row[2].text = dt[2] #Escenario
            row[3].text = dt[3] #Hora
            row[4].text = str(int(dt[4])) #Muestra
            row[5].text = str(int(dt[5])) #<10
            row[6].text = str(int(dt[6])) #>10
            row[7].text = str(int(dt[7]*100)) #ER MAX
            row[8].text = str(int(dt[8]*100)) #ER MIN
            row[9].text = str(int(dt[9]*100)) #ER PROMEDIO
            row[10].text = str(int(dt[10]*100)) #PORCENTAJE <10

        for row in table.rows:
            for cell in row.cells:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                cell.vertical_alignment = WD_ALIGN_PARAGRAPH.CENTER
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(10)

    try:
        direccion_archivo = os.path.join(ruta_vehicular,f"INFORME_GENERAL_VEHICULAR.docx")
        doc.save(direccion_archivo)
        print("¡Informe general guardado con éxito!")
    except IndexError:
        print(f"No hay datos para generar un informe general")

    final_word = time.perf_counter()
    duracion_word = final_word-inicio_word
    print(f"Processing Time - Word: {duracion_word:.2f} seconds.")
    print("#### Comparación de flujogramas vehiculares finalizado ####")