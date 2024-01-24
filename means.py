from openpyxl import load_workbook
import logging
import numpy as np
import os

#Define Logger

LOGGER = logging.getLogger(__name__)
LOGGER.setLevel(logging.DEBUG)
f = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

""" 
10  DEBUG       logging.debug()     Lowest level. Used to record simple details.
20  INFO        logging.info()      Record general information.
30  WARNING     logging.warning()   Potential issues which may not cause errors in the future.
40  ERROR       logging.error()     Record erros which causes a section of the code to fail.
50  CRITICAL    logging.critical()  Highest level. Blockers which fails your whole program.
 """

def read_excel(excel_path
               ) -> type(np.array) and type(np.array) and list:
    """Función que obtiene los conteos por giros de vehículos y motocicletas para todos los sentidos.
    """
    
    #Slices definition
    turn_slice = [slice("G12","G21"),
                  slice("M12","M21"),
                  slice("G24","G33"),
                  slice("M24","M33")]
    
    car_slice = slice("K16","T112")
    moto_slice = slice("U16","AD112")

    #Reading number of turns
    hojas = ['N','S','E','O']
    wb = load_workbook(excel_path, read_only=True, data_only=True)
    LOGGER.info(f"Lectura de excel:\n{excel_path}")
    ws = wb['Inicio']

    try:
        quantities = [
            [row[0].value for row in ws[s]].index(None)
            for s in turn_slice
        ]
    except ValueError:
        LOGGER.error(f"Hay algun dato que no es número en la base de datos")

    #Reading data from sheets
    CAR_LIST = []
    MOTO_LIST = []

    for hoja, num in zip(hojas,quantities):
        ws = wb[hoja]

        car_data = [[elem.value for elem in row][:num] for row in ws[car_slice]]
        try:
            CAR = np.array(car_data, dtype="float")
        except ValueError:
            LOGGER.critical(f"Para autos en la hoja '{hoja}' hay datos que no son números")
        CAR = np.nan_to_num(CAR, nan=0.0)
        CAR_MORNING = CAR[26:38,:]
        CAR_EVENING = CAR[48:60,:]
        CAR_NIGHT = CAR[70:82,:]

        moto_data = [[elem.value for elem in row][:num] for row in ws[moto_slice]]
        try:
            MOTO = np.array(moto_data, dtype="float")
        except ValueError:
            LOGGER.critical(f"Para motos en la hoja '{hoja}' hay datos que no son números")
        MOTO = np.nan_to_num(MOTO, nan=0.0)
        MOTO_MORNING = MOTO[27:39,:]
        MOTO_EVENING = MOTO[49:61,:]
        MOTO_NIGHT = MOTO[71:83,:]

        CAR_LIST.append([CAR_MORNING,CAR_EVENING,CAR_NIGHT]) #POR HOJA
        MOTO_LIST.append([MOTO_MORNING,MOTO_EVENING,MOTO_NIGHT]) #POR HOJA

    wb.close()

    return CAR_LIST, MOTO_LIST

def find_duplicate(CARS, length) -> None:
    count = 0
    repes_list = []
    CAR = np.concatenate((CARS[0],CARS[1]))
    CAR = np.concatenate((CAR,CARS[2]))

    for k in range(CAR.shape[1]):
        GIRO = CAR[:,k]
        for i in range(len(GIRO)-length+1):
            set = GIRO[i:i+length]
            if np.array_equal(set, np.zeros(length)):
                continue
            for j in range(i+1, len(GIRO) - length + 1):
                if np.array_equal(GIRO[j:j+length], set):
                    print(f"Conjunto repetido: {set} en hoja Nro. {k}")
                    repes_list.append(set)
                    count += 1

    return count, repes_list

def jump_single():
    pass

def jump_multiple():
    pass

def main():
    excel_path = r"Pruebas/SS-25_ Av. Circunvalación Golf Los Incas - Av. La Fontana.xlsm"

    directory, _ = os.path.split(excel_path)
    logger_path = os.path.join(directory,"LOGS")
    if not os.path.exists(logger_path):
        os.mkdir(logger_path)
    fh = logging.FileHandler(os.path.join(directory,"LOGS","means.log"))
    fh.setFormatter(f)
    LOGGER.addHandler(fh)

    CARS_LIST, MOTOS_LIST = read_excel(excel_path)
    repes_count = 0
    repes_list = []
    for CARS in CARS_LIST: #CARS = Arreglo por sentido.
        car_result, car_reps = find_duplicate(CARS,2)
        repes_count += car_result
        repes_list.append(car_reps)

    print(f"Número de repeticiones = {repes_count}")
    print(f"Repeticiones = {repes_list}")

    #moto_result, moto_reps = find_duplicate(MOTOS_LIST,2)

if __name__ == '__main__':
    main()