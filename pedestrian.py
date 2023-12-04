from openpyxl import load_workbook
from openpyxl.styles import PatternFill, NamedStyle, Border, Side
import numpy as np
import shutil
import docx
import time
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Inches
import os
import re
from reading import read_excel_peatonal

def peatonal(entregable_path):
    ruta = entregable_path
    #ruta_N1 = os.listdir(ruta)#lista de las rutas 
    #print(f"ruta del entregable: {ruta_N1}")
    
    listas = [[] for _ in range(4)]

    ruta_N2 = os.path.join(ruta,"7.- Informacion de Campo","Peatonal") #SE CAMBIA
    files = ["Atipico",
            "Atipico (Protransito)",
            "Tipico",
            "Tipico (Protransito)",]
    
    #print(f"Lista de carpetas:{files}") # los cuales se van analizar
    #slice("L3","AZ3")
    for file in files:
        ruta_penultima = os.path.join(ruta_N2,file)
        hojasexcel = os.listdir(ruta_penultima)
        for hojaexcel in hojasexcel:
            ruta_ultima = os.path.join(ruta_penultima,hojaexcel)
            #print(f": {ruta_penultima}") 
            if not "~$" in hojaexcel: #Depuración para archivos ocultos "~$..."
                        
                if file == "Tipico": #0
                        listas[0].append(ruta_ultima)
                            
                elif file == "Atipico": #1
                        listas[1].append(ruta_ultima)
                        
                elif file == "Tipico (Protransito)": #2
                        listas[2].append(ruta_ultima)
                        
                elif file == "Atipico (Protransito)":
                    listas[3].append(ruta_ultima)

    archivos = [[] for _ in range(2)] #LISTA DE 2
    
    #tuplas disgregar _REV COMPARA TODA LA RUTA NO SOLO ARCHIVO
    for veh2 in listas[2]: #COMPARACIÓN DE TIPICO CON EL DE PROTRÁNSITO 
        for veh0 in listas[0]:
            nombre_veh2= os.path.basename(veh2)
            nombre_veh0= os.path.basename(veh0)
            #print(f"veh1:{nombre_veh1}",f"veh2: {nombre_veh2}")
            if nombre_veh2 != nombre_veh0 and nombre_veh2.replace("_REV","") == nombre_veh0:
                archivos[0].append((veh2,veh0))

    for veh3 in listas[3]: #COMPARACIÓN DE ATIPICO CON EL DE PROTRÁNSITO 
        for veh1 in listas[1]:
            nombre_veh3= os.path.basename(veh3)
            nombre_veh1= os.path.basename(veh1)
            #print(f"veh1:{nombre_veh1}",f"veh2: {nombre_veh2}")

            if nombre_veh3 != nombre_veh1 and nombre_veh3.replace("_REV","") == nombre_veh1:
                archivos[1].append((veh3,veh1))
    
    #print(f"imprimir archivos : {archivos}")
    #global hojas

    data_row = [[] for _ in range(2)] #LISTA DE 2 CREA UNA LISTA

    inicio_lectura = time.time()
    
    for j in range(2): #[[(veh1,veh2),(veh1,veh2)],[(veh1,veh2),(veh1,veh2)],[(veh1,veh2),(veh1,veh2)],[(veh1,veh2),(veh1,veh2)]]  #LISTA DE 2
        if j == 0:
            print("#### Reportes Tipicos ####")
        else:
            print("#### Reportes Atipicos ####")

        for index,arch in enumerate(archivos[j]): #el archivo es un par "j" entra a tupla
            if j==0:
                print(f"Comparando Tipicos ({index+1}/{len(archivos[j])})")
            else:
                print(f"Comparando Atipicos ({index+1}/{len(archivos[j])})")
            #0: Supervisor / 1: Consultor
            data_consultor,fecha,_ = read_excel_peatonal(arch[1]) #pares de archivos - tupla
            data_supervisor,_,_ = read_excel_peatonal(arch[0]) #"_" para solo requerir dato de la primera entrada
    
            _,_,encabezado= read_excel_peatonal(arch[1])  

            ruta_destino = arch[0].replace("_REV.xlsm","_REP.xlsx")
            
            container_file = os.path.dirname(arch[1])
            _,tipicidad = os.path.split(container_file)

            resultados = []

            volumen_supervisor  = 0
            volumen_consultor   = 0

            ##########################
            # LECTURA DE INFORMACIÓN #
            ##########################
            
            lista_primeros_valores = []
            lista_ultimos_valores =[]
            #LEE LA HOJA Y LUEGO LAS CELDAS CON MASCARAS
            for data_con,data_sup in zip(data_consultor,data_supervisor): #lectura paralela
                mascara = data_sup!=0 #define el verdadero,boleano,verdadero-falso
                resultado = np.zeros_like(data_con)
                volumen_supervisor += np.sum(data_sup[mascara])
                volumen_consultor+= np.sum(data_con[mascara])
                resultado[mascara] = np.abs((data_con[mascara] - data_sup[mascara])) / data_sup[mascara]
                resultados.append(resultado)
                
                filas_con_info = np.any(mascara,axis=1)
                indices=np.where(filas_con_info)[0]
                
                if indices.size> 0:
                    lista_primeros_valores.append(indices[0]) #Para saber el rango
                    lista_ultimos_valores.append(indices[-1])
                #print(lista_primeros_valores,lista_ultimos_valores)

            if lista_primeros_valores and lista_ultimos_valores:
                valor_ultimo = max(lista_ultimos_valores)
                valor_primero = min(lista_primeros_valores)

            #print(f"resultado de la diferencia: {resultados}")  
            #print(f"primer valor: {valor_primero}")
            #print(f"ultimo valor:{valor_ultimo}")

            primer_valor = ((valor_primero+24)*15)//60
            ultimo_valor = (((valor_ultimo+24)+1)*15)//60
            residuo_primero = (((valor_primero+24)*15)) % 60
            residuo_segundo = (((valor_ultimo+24)+1)*15) % 60
            hora_muestra = f"{primer_valor:02d}:{residuo_primero:02d}"+"-"+f"{ultimo_valor:02d}:{residuo_segundo:02d}"
            #print(hora_muestra)
                
            suma_total = 0
            num_valores = 0
            conteo_mayor_10 = 0
            conteo_menor_10 = 0
            error_minimo = 1
            error_maximo = 0

            for result in resultados: #análisis de celda
                mascara = result !=0 #aca s ve la cantidad de errores
                #print(mascara)
                suma_total += np.sum(result[mascara])
                #print(f"suma de la hoja:{suma_total}")
                num_valores += np.count_nonzero(mascara)
                conteo_mayor_10 += np.count_nonzero(result[mascara] > 0.1) #celda
                conteo_menor_10 += np.count_nonzero(result[mascara] <= 0.1)
                                    
                if result[mascara].size > 0:
                    error_maximo_actual = np.max(result[mascara])
                    error_minimo_actual = np.min(result[mascara])
                    error_maximo = max(error_maximo, error_maximo_actual)
                    error_minimo = min(error_minimo, error_minimo_actual)
            
            promedio = round(suma_total / num_valores,2)
            muestra = conteo_mayor_10+conteo_menor_10
            interseccion_critica = round(conteo_mayor_10/muestra,2)
            
            path_formato = "./tools/Formato_Peatonal.xlsx"
            current_file, name_excel = os.path.split(ruta_destino)
            destiny_file = os.path.join(current_file,'Reportes')
            final_route = os.path.join(destiny_file,name_excel)
            if not os.path.exists(destiny_file):
                os.makedirs(destiny_file)
            shutil.copyfile(path_formato,final_route) #copia el contenido del archivo path formato a otro archivo ubicado en la ruta_destino

            wb = load_workbook(final_route) # es la ruta del archivo
            ws = wb["Data Peatonal"]

            #Definición de formatos:
            porcentaje_style = NamedStyle(name="porcentaje")
            porcentaje_style.number_format = '0%'
            porcentaje_style.border = Border(left=Side(style='thin'),
                                            right=Side(style='thin'),
                                            top=Side(style='thin'),
                                            bottom=Side(style='thin'),)

            #Asignación de formatos por hojas: #esto se tiene que cambiar #encabezado
            #for mat in enumerate(encabezado):
            for f, supid in enumerate(encabezado):
                arriba = ws.cell(row=19,column=f+12)
                arriba.value = supid
                #print(arriba)
                    
            for result in resultados:
                for k, fila in enumerate(result):
                    for l, valor in enumerate(fila):
                        celda = ws.cell(row=k+20,column=l+12)
                        if valor == 0:
                            celda.value = None
                        else:
                            celda.value = valor

                        celda.style = porcentaje_style

                        if valor >=0.1:
                            relleno = PatternFill(start_color="FFFF0000", end_color="FFFF0000",fill_type="solid")
                            celda.fill = relleno
            #establece patrones
            wb.save(final_route)
            _,nombre_archivo = os.path.split(ruta_destino)
            patron = r"^([A-Z]+[0-9]+)_" 
            coincidencia = re.search(patron,nombre_archivo)

            if coincidencia:
                codigo = coincidencia.group(1) #coincidencia
            else:
                print(f"NO HAY CÓDIGO, REVISAR ARCHIVO {nombre_archivo}")
                codigo = "ERROR"

            #RESULTADOS FINALES
            data_row[j].append((codigo,fecha[:-9],tipicidad,hora_muestra,muestra,conteo_menor_10,conteo_mayor_10,error_maximo,error_minimo,promedio,interseccion_critica))

    final_lectura = time.time()

    duracion_lectura = final_lectura-inicio_lectura

    print(f"Processing Time - Reading: {duracion_lectura:.2f} seconds.")

    ########################
    # CREACIÓN DEL REPORTE #
    ########################

    inicio_word = time.time()
    
    doc = docx.Document()
    doc.add_heading(f"REPORTE FLUJO PEATONAL") #ENCABEZADO 
    table = doc.add_table(rows=1,cols=11,style='Table Grid')  #SE AGREGA LA TABLA
    section = doc.sections[0] 
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)
    
    for cell in table.columns[1].cells:
        cell.width = Inches(2)  # Ajustar al ancho deseado

    # Ajustar el ancho de la sexta columna (índice 5)
    for cell in table.columns[3].cells:
        cell.width = Inches(2)

    titulos = ['ID', 'Fecha Muestra', 'Esc.', 'Hora Muestra', 'Muestra', 
                    '<10%', '>10%', 'E.R.\n%(máx)', 'E.R.\n%(mín)', 
                    'E.R.\nProm', '>10/mues.%']

    for i, titulo in enumerate(titulos):
        celda = table.rows[0].cells[i]
        parrafo = celda.paragraphs[0]
        run = parrafo.add_run(titulo)
        run.bold = True  # Aplica negrita al run
        run.font.size = Pt(10)

    for i,data in enumerate(data_row): #Son 2
        if len(data)==0: #data igual a cero se passa
            pass
                    
        #Agregación de datos
        for dt in data:
            row = table.add_row().cells #agrega una celda
            row[0].text = dt[0] #Intersección
            row[1].text = dt[1] #Fecha, tiene que entrar com str
            row[2].text = dt[2] #Escenario
            row[3].text = dt[3] #Hora, tiene que entrar com str
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
                        
    #MODIFICAMOS LA UBICACIÓN DEL INFORME 
    try:
        #ruta_word,_ = os.path.split(ruta_N2) # aca cambiar a la ruta origina
        _,nombre_del_archivo = os.path.split(entregable_path)
        direccion_archivo = os.path.join(ruta_N2,f"INFORME_GENERAL_PEATONAL {nombre_del_archivo}.docx") #JUNTO EL NOMBRE DEL DOC CON LA RUTA DE DOCUMENTO , agregar nombre del entregable
        doc.save(direccion_archivo)
        print("Informe general guardado correctamente")    
    except IndexError:
        print(f"No hay datos para generar un informe general.")
    
    final_word = time.time()
    duracion_word = final_word-inicio_word
    print(f"Processing Time - Word: {duracion_word:.2f} segundos.")
    print("#### Comparación de flujogramas peatonales finalizado ####")