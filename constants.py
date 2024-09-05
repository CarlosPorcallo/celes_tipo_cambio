import os

URL_SITIO: str = "<url_sitio_banxico>"
WAIT: int = 40

FILENAME_TIPO_CAMBIO: str = "tipoCambio.xls"
FILENAME_NOMINA_ESPEJO: str = "nominaEspejo.xlsx"

DWN_DIR: str = "</path/to/downloads>"
TMP_DIR: str = f"{os.getcwd()}/tmp"
OUTPUT_DIR: str = f"{os.getcwd()}/output"

XPATH_FECHA_INICIAL: str = '//*[@id="fechaInicialContainer"]/table/tbody/tr/td[1]/input'
XPATH_FECHA_FINAL: str = '//*[@id="fechaFinalContainer"]/table/tbody/tr/td[1]/input'
XPATH_SELECT_FORMATO: str = '/html/body/form/table/tbody/tr[3]/td/table/tbody/tr/td[2]/table/tbody/tr[2]/td[1]/div/select'
XPATH_BTN_SUBMIT: str = '/html/body/form/table/tbody/tr[3]/td/table/tbody/tr/td[2]/table/tbody/tr[2]/td[2]/input'
