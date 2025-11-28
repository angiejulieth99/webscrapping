import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook
from webdriver_manager.firefox import GeckoDriverManager


def ejecutar_flujo():
    url = "https://apps5.mineco.gob.pe/transparencia/Navegador/Default.aspx"

    options = Options()
    options.headless = False  # si quieres ocultar la ventana, pon True

    # Geckodriver automático
    service = Service(GeckoDriverManager().install())

    driver = webdriver.Firefox(service=service, options=options)
    driver.get(url)
    time.sleep(3)

    driver.switch_to.frame("frame0")

    wait = WebDriverWait(driver, 20)

    boton_nivel = wait.until(
        EC.element_to_be_clickable((By.ID, "ctl00_CPH1_BtnTipoGobierno"))
    )
    boton_nivel.click()

    radio = wait.until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[type="radio"][value^="E/"]'))
    )
    driver.execute_script("arguments[0].click();", radio)

    boton_sector = wait.until(
        EC.element_to_be_clickable((By.ID, "ctl00_CPH1_BtnSector"))
    )
    boton_sector.click()

    tabla = wait.until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "table.Data"))
    )

    filas = tabla.find_elements(By.TAG_NAME, "tr")

    wb = Workbook()
    ws = wb.active
    ws.title = "Sectores"

    ws.append([
        "Sector",
        "PIA",
        "PIM",
        "Certificación",
        "Compromiso Anual",
        "Atención de Compromiso Mensual",
        "Devengado",
        "Girado",
        "Avance %"
    ])

    for fila in filas:
        celdas = fila.find_elements(By.TAG_NAME, "td")

        if not celdas:
            celdas = fila.find_elements(By.TAG_NAME, "th")

        valores = [c.text.strip() for c in celdas]
        valores = valores[1:]  # eliminar primera columna

        # Quitamos los 4 primeros dígitos "Sector"
        if len(valores) > 0:
            valores[0] = valores[0][4:]

        ws.append(valores)

    wb.save("Gobierno_Nacional.xlsx")
    driver.quit()


if __name__ == "__main__":
    ejecutar_flujo()
