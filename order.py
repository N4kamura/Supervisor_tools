import os
import re
import shutil

#Patterns:
patron_atipico = ["_A.xlsm$","A.xlsm$","_A\(.*\)$","_A..xlsm$","_A \(.*\).xlsm$",
                "_A.xlsx$","A.xlsx$","_A\(.*\)$","_A..xlsx$","_A \(.*\).xlsx$"]
patron_atipico = '|'.join(patron_atipico)

patron_tipico = ["_T.xlsm$","T.xlsm$","_T\(.*\)$","_t..xlsm$","_T \(.*\).xlsm$",
                "_T.xlsx$","T.xlsx$","_T\(.*\)$","_t..xlsx$","_T \(.*\).xlsx$"]
patron_tipico = '|'.join(patron_tipico)

def order_files(root_path):
    field_path = os.path.join(root_path,'7.- Informacion de Campo')

    #path = r"C:\Users\dacan\OneDrive\Desktop\PRUEBAS\Supervisor\Entregable Nro 06\7.- Informacion de Campo\Embarque y Desembarque"
    counts = os.listdir(field_path)
    for count_file in counts: #Enter to Embarque y Desembarque for example.
        if count_file == 'Tiempo de Ciclo Semaforico':
            continue

        path_tipico = os.path.join(field_path,count_file,'Tipico')
        path_atipico = os.path.join(field_path,count_file,'Atipico')

        if not os.path.exists(path_tipico):
            os.makedirs(path_tipico)

        if not os.path.exists(path_atipico):
            os.makedirs(path_atipico)

        list_files = os.listdir(os.path.join(field_path,count_file))
        excel_files = [excel for excel in list_files if excel.endswith('xlsx') or excel.endswith('xlsm')]

        for excel in excel_files:
            if re.search(patron_atipico,excel,re.IGNORECASE):
                shutil.move(os.path.join(field_path,count_file,excel), path_atipico)
            elif re.search(patron_tipico,excel,re.IGNORECASE):
                shutil.move(os.path.join(field_path,count_file,excel), path_tipico)
            else:
                print(f"Este archivo de excel no tiene ningún patrón:\n{excel}")