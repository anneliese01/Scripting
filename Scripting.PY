import time
import keyboard
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# ================= CONFIGURACIÓN =================
URL = "https://opm-digemid.minsa.gob.pe/#/consulta-producto"
PRODUCTO = "BISOPROLOL"
ARCHIVO = "digemid_ok.xlsx"
DEPARTAMENTO_FIJO = "LIMA"
PROVINCIA_FIJA = "LIMA"
DISTRITO_FIJO = "LOS OLIVOS"

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")

driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=options
)

wait = WebDriverWait(driver, 60)

# ================= CONTROL DE PARADA =================
detener = False
finalizar = False


def detener_scraping():
    global detener
    detener = True
    print("\n F8 presionado → pausa inmediata")


def finalizar_scraping():
    global finalizar
    finalizar = True
    print("\n F9 presionado → finalizando y guardando")


keyboard.add_hotkey("F8", detener_scraping)
keyboard.add_hotkey("F9", finalizar_scraping)

print("ℹ F8 = pausar | F9 = finalizar y guardar")

# ================= FUNCIONES =================
def seleccionar_1000():
    try:
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//select[.//option[contains(text(),'1000')]]")
        )).click()

        wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//option[contains(text(),'1000')]")
        )).click()

        time.sleep(2)
    except Exception:
        print("⚠ No se pudo seleccionar 1000 registros")


def get_by_label(label, intentos=3, delay=1):
    for _ in range(intentos):
        if finalizar:
            return ""
        try:
            el = driver.find_element(
                By.XPATH,
                f"//label[contains(normalize-space(.), '{label}')]/following::input[1]"
            )
            valor = el.get_attribute("value")
            if valor:
                return valor.strip()
        except Exception:
            pass
        time.sleep(delay)
    return ""


def get_nombre_producto(intentos=3, delay=1):
    for _ in range(intentos):
        if finalizar:
            return ""
        try:
            el = driver.find_element(By.XPATH, "//input[@name='nombreProducto']")
            valor = el.get_attribute("value") or el.get_attribute("ng-reflect-model")
            if valor:
                return valor.strip()
        except Exception:
            pass
        time.sleep(delay)
    return ""


def get_datos_fila(boton):
    fila = boton.find_element(By.XPATH, "./ancestor::tr")
    celdas = fila.find_elements(By.TAG_NAME, "td")

    def safe_text(i):
        try:
            return celdas[i].text.strip()
        except Exception:
            return ""

    return {
        "Tipo": safe_text(0),
        "Fecha de Actualización": safe_text(1),
        "Producto (tabla)": safe_text(2),
        "Laboratorio (tabla)": safe_text(3),
        "Farmacia/Botica (tabla)": safe_text(4),
        "Precio Unitario (tabla)": safe_text(5),
    }


def esperar_enter_o_f9(mensaje=" ENTER para continuar | F9 para finalizar"):
    print(mensaje)
    while True:
        if finalizar:
            return False
        if keyboard.is_pressed("enter"):
            time.sleep(0.3)  # anti rebote
            while keyboard.is_pressed("enter"):
                time.sleep(0.05)
            return True
        time.sleep(0.1)


def guardar_excel(datos, columnas, archivo):
    df = pd.DataFrame(datos, columns=columnas)
    df.to_excel(archivo, index=False)
    print(f"\n Excel generado: {archivo}")
    print(f" Total registros: {len(datos)}")


# ================= COLUMNAS =================
columnas = [
    "RUC",
    "Tipo",
    "Fecha de Actualización",
    "Producto (tabla)",
    "Titular",
    "Fabricante",
    "Farmacia/Botica",
    "Correo electronico",
    "Horario de atencion",
    "Teléfono",
    "Precio Unit.",
    "Departamento",
    "Provincia",
    "Distrito",
    "Dirección",
    "Director Técnico - Químico Farmacéutico",
    "Nombre",
    "Presentación",
    "Registro Sanitario",
    "Condición de venta",
    "Nombre del Titular",
    "Nombre del Fabricante",
    "Precio unitario S/.",
    "Precio empaque S/.",
    "Tipo producto",
    "País de fabricación"
]

datos = []

# ================= PROGRAMA PRINCIPAL =================
try:
    # ================= NAVEGACIÓN =================
    driver.get(URL)

    try:
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//button[contains(.,'Cerrar')]"))
        ).click()
    except Exception:
        pass

    input_producto = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//input[@type='text']"))
    )
    input_producto.clear()
    input_producto.send_keys(PRODUCTO)

    print("👉 Selecciona MANUALMENTE el medicamento")
    wait.until(lambda d: input_producto.get_attribute("value") != PRODUCTO)

    # ================= FILTROS =================
    wait.until(EC.element_to_be_clickable((By.XPATH, "(//select)[1]"))).click()
    wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//option[contains(text(),'LIMA')]"))
    ).click()

    print("\n PAUSA OBLIGATORIA")
    print(" Selecciona MANUALMENTE el DISTRITO (LIMA)")
    continuar = esperar_enter_o_f9()
    if not continuar:
        raise SystemExit

    wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//button[contains(.,'Buscar')]"))
    ).click()

    seleccionar_1000()

    # ================= SCRAPEO =================
    pagina = 1

    while True:
        if finalizar:
            break

        print(f"\nScrapeando página {pagina}")

        wait.until(EC.presence_of_all_elements_located(
            (By.XPATH, "//a[@title='Ver detalle']"))
        )

        botones = driver.find_elements(By.XPATH, "//a[@title='Ver detalle']")
        total = len(botones)
        print(f"🔎 Registros en página: {total}")

        detener = False  # resetea pausa en cada página

        for i in range(total):
            if detener or finalizar:
                break

            try:
                botones = driver.find_elements(By.XPATH, "//a[@title='Ver detalle']")
                if i >= len(botones):
                    break

                boton = botones[i]

                datos_tabla = get_datos_fila(boton)

                driver.execute_script("arguments[0].scrollIntoView(true);", boton)
                time.sleep(0.35)
                driver.execute_script("arguments[0].click();", boton)

                wait.until(EC.visibility_of_element_located(
                    (By.CLASS_NAME, "modal-content"))
                )

                nombre_producto = get_nombre_producto()

                fila = {
                    "RUC": get_by_label("RUC"),
                    "Tipo": datos_tabla["Tipo"],
                    "Fecha de Actualización": datos_tabla["Fecha de Actualización"],
                    "Producto (tabla)": datos_tabla["Producto (tabla)"],
                    "Titular": get_by_label("Nombre del Titular"),
                    "Fabricante": get_by_label("Nombre del Fabricante"),
                    "Farmacia/Botica": get_by_label("Nombre Comercial"),
                    "Correo electronico": get_by_label("Correo electrónico"),
                    "Horario de atencion": get_by_label("Horario de atención"),
                    "Teléfono": get_by_label("Teléfono"),
                    "Precio Unit.": get_by_label("Precio unitario"),
                    "Departamento": DEPARTAMENTO_FIJO,
                    "Provincia": PROVINCIA_FIJA,
                    "Distrito": DISTRITO_FIJO,
                    "Dirección": get_by_label("Dirección"),
                    "Director Técnico - Químico Farmacéutico": get_by_label("Director Técnico"),
                    "Nombre": nombre_producto,
                    "Presentación": get_by_label("Presentación"),
                    "Registro Sanitario": get_by_label("Registro Sanitario"),
                    "Condición de venta": get_by_label("Condición de venta"),
                    "Nombre del Titular": get_by_label("Nombre del Titular"),
                    "Nombre del Fabricante": get_by_label("Nombre del Fabricante"),
                    "Precio unitario S/.": get_by_label("Precio unitario"),
                    "Precio empaque S/.": get_by_label("Precio empaque"),
                    "Tipo producto": get_by_label("Tipo producto"),
                    "País de fabricación": get_by_label("País de fabricación")
                }

                datos.append(fila)

                try:
                    wait.until(EC.element_to_be_clickable(
                        (By.XPATH, "//button[contains(.,'Cerrar')]"))
                    ).click()
                except Exception:
                    pass

            except Exception as e:
                print(f"⚠ Error en registro {i + 1}: {e}")
                try:
                    driver.find_element(By.XPATH, "//button[contains(.,'Cerrar')]").click()
                except Exception:
                    pass

        if finalizar:
            break

        if detener:
            print("\n⏸ Pausado con F8")
            continuar = esperar_enter_o_f9(" ENTER para reanudar | F9 para finalizar")
            if not continuar:
                break
            detener = False

        print("\n CAMBIO MANUAL DE PÁGINA")
        print(" Cambia a la SIGUIENTE página")
        print(" Vuelve a seleccionar 1000")

        continuar = esperar_enter_o_f9()
        if not continuar:
            break

        pagina += 1
        seleccionar_1000()

except SystemExit:
    print(" Programa finalizado por el usuario")

except Exception as e:
    print(f" Error general: {e}")

finally:
    try:
        guardar_excel(datos, columnas, ARCHIVO)
    except Exception as e:
        print(f" No se pudo guardar el Excel: {e}")

    try:
        driver.quit()
    except Exception:
        pass

    print(" Programa terminado")