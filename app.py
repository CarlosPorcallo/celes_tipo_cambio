from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.support.select import Select
from time import sleep

import constants as cons
import pandas as pd
import argparse
import shutil
import sys

URL_SITIO: str = cons.URL_SITIO
WAIT: int = cons.WAIT

FILENAME_TIPO_CAMBIO: str = cons.FILENAME_TIPO_CAMBIO
FILENAME_NOMINA_ESPEJO: str = cons.FILENAME_NOMINA_ESPEJO

DWN_DIR: str = cons.DWN_DIR
TMP_DIR: str = cons.TMP_DIR
OUTPUT_DIR: str = cons.OUTPUT_DIR

XPATH_FECHA_INICIAL: str = cons.XPATH_FECHA_INICIAL
XPATH_FECHA_FINAL: str = cons.XPATH_FECHA_FINAL
XPATH_SELECT_FORMATO: str = cons.XPATH_SELECT_FORMATO
XPATH_BTN_SUBMIT: str = cons.XPATH_BTN_SUBMIT

def main():
    # pedir fechas
    parser = argparse.ArgumentParser(description='Fechas para dscarga de reportes.')
    parser.add_argument('--fecha_inicial', type=str, nargs='+',help='Fecha de inicio para el reporte: ddmmyyyy')
    parser.add_argument('--fecha_final', type=str, nargs='+',help='Fecha de inicio para el reporte: ddmmyyyy')
    args = parser.parse_args()

    if (args.fecha_inicial is not None and args.fecha_final is not None):
        fecha_inicial: datetime = datetime(int(args.fecha_inicial[0][4:8]), int(args.fecha_inicial[0][2:4]), int(args.fecha_inicial[0][0:2]))
        fecha_final: datetime = datetime(int(args.fecha_final[0][4:8]), int(args.fecha_final[0][2:4]), int(args.fecha_final[0][0:2]))
    else:
        print("Por favor proporcione una fecha inicial y una final para comenzar.")
        sys.exit()

    ### descargar excel

    # se inicia selenium
    options = Options()
    options.headless = True
    driver = webdriver.Firefox(options = options)

    try:
        driver.get(URL_SITIO)
        
        # se establecen las fechas 
        input_fecha_inicial = WebDriverWait(driver, WAIT).until(EC.element_to_be_clickable((By.XPATH, XPATH_FECHA_INICIAL)))
        input_fecha_final = WebDriverWait(driver, WAIT).until(EC.element_to_be_clickable((By.XPATH, XPATH_FECHA_FINAL)))

        input_fecha_inicial.clear()
        input_fecha_inicial.send_keys(fecha_inicial.strftime("%d/%m/%Y"))
        input_fecha_final.clear()
        input_fecha_final.send_keys(fecha_final.strftime("%d/%m/%Y"))

        # se elige el formato de descarga 
        select_formato = Select(driver.find_element(By.XPATH, XPATH_SELECT_FORMATO))
        select_formato.select_by_value("XLS")

        # se da click en descargar
        btn_descarga = WebDriverWait(driver, WAIT).until(EC.element_to_be_clickable((By.XPATH, XPATH_BTN_SUBMIT)))
        btn_descarga.click()

        sleep(5)
    except Exception as e:
        print(e)
        driver.close()
    finally:
        driver.close()

        # abrir excel y extraer columna
        source_file: str = f"{TMP_DIR}/{FILENAME_TIPO_CAMBIO}"
        shutil.move(f"{DWN_DIR}/{FILENAME_TIPO_CAMBIO}", source_file)

        df_tipo_cambio: pd.DataFrame = pd.read_excel(source_file)
        df_tipo_cambio.drop([0,1,2,3,4,6], inplace = True)
        
        columns = df_tipo_cambio.head(1)
        columns = columns.to_dict()
        new_columns: dict = {}
        
        for k in columns.keys():
            new_columns[k] = columns[k][5]

        df_tipo_cambio.rename(columns = new_columns, inplace = True)
        df_tipo_cambio.drop([5], inplace = True)

        # agregar datos a otro excel
        dest_file: str = f"{OUTPUT_DIR}/{FILENAME_NOMINA_ESPEJO}"
        shutil.copyfile(f"{TMP_DIR}/{FILENAME_NOMINA_ESPEJO}", dest_file)

        # se valida si la hoja existe
        file = pd.ExcelFile(dest_file)  
        sheets: list = file.sheet_names

        # si al hoja existe se lee y se concatena a lo leído desde el archivo
        sheet_name = "Tipo de cambio"
        if sheet_name in sheets:
            df_tipo_cambio_leido = pd.read_excel(dest_file, sheet_name=sheet_name)

            with pd.ExcelWriter(dest_file, engine='openpyxl', mode='a') as writer: 
                workBook = writer.book
                try:
                    workBook.remove(workBook[sheet_name])
                except Exception as e:
                    print(e)
                    sys.exit()

            if len(df_tipo_cambio_leido) != 0:
                df_final: pd.DataFrame = pd.concat([df_tipo_cambio_leido, df_tipo_cambio], join = "inner")
                df_final.drop_duplicates(inplace = True)
            else:
                df_final: pd.DataFrame = df_tipo_cambio
        else:
            df_final: pd.DataFrame = df_tipo_cambio

        df_final["Fecha"] = pd.to_datetime(df_final['Fecha'], format = "%d/%m/%Y")
        df_final.sort_values(by = 'Fecha', ascending = True, inplace = True)

        # se agrega la hoja al nuevo excel
        with pd.ExcelWriter(dest_file, engine='openpyxl', mode='a') as writer:
            df_final.to_excel(writer, sheet_name=sheet_name, index=False)

        # se rota el archivo de entrada
        shutil.copy(f"{TMP_DIR}/{FILENAME_NOMINA_ESPEJO}", f"{TMP_DIR}/{FILENAME_NOMINA_ESPEJO}.bk")
        shutil.copy(dest_file, f"{TMP_DIR}/{FILENAME_NOMINA_ESPEJO}")

if __name__ == "__main__":
    main()