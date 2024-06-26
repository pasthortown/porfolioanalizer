import pandas as pd
import json
import openpyxl
import os 
import requests

archivo_excel = "portafolio.xlsx"
nombre_hoja = "BASE"
columnas_deseadas = ["Producto", "Requerimiento", "Correo Electrónico", "Área", "Aprobación", "Prioridad\nPO", "Progreso"]
areas_excluir = ["DSI", "VE", "PAÍSES", "CL", "AR", "CO", "ES"]
url_servidor = "http://localhost:5555/send"

def cargar_contactos_desde_json():
    with open('salida.json', 'r', encoding='utf-8') as archivo_json:
        contactos_por_area = json.load(archivo_json)
    return contactos_por_area

def cargar_productos_desde_json():
    with open('productos.json', 'r', encoding='utf-8') as archivo_json:
        productos = json.load(archivo_json)
    return productos

def cargar_data_base(filename, sheet, columns):
    datos_excel = pd.read_excel(filename, sheet_name=sheet)
    columnas_filtradas = datos_excel[datos_excel["Prioridad\nPO"].isin(['Q1', 'Q2'])]
    columnas_filtradas = columnas_filtradas[columns]
    matriz_datos = columnas_filtradas.values.tolist()
    return matriz_datos

def get_areas(datos, areas_excluir=None):
    indice_area = columnas_deseadas.index("Área")
    columna_area = [fila[indice_area] for fila in datos]
    columna_area_sin_duplicados = pd.Series(columna_area).drop_duplicates().dropna().tolist()
    if areas_excluir:
        columna_area_sin_duplicados = [area for area in columna_area_sin_duplicados if area not in areas_excluir]
    return columna_area_sin_duplicados

def filtrar_por_area(datos, area):
    indice_area = columnas_deseadas.index("Área")
    indice_contacto = columnas_deseadas.index("Correo Electrónico")
    indices_filtrar = [indice_area, indice_contacto]
    datos_filtrados = []
    for fila in datos:
        if fila[indice_area] == area:
            fila_filtrada = [fila[i] for i in range(len(fila)) if i not in indices_filtrar]
            datos_filtrados.append(fila_filtrada)    
    return datos_filtrados

def obtener_datos_por_areas(base, areas):
    datos_por_areas = {}
    for area in areas:
        datos_filtrados = filtrar_por_area(base, area)
        datos_por_areas[area] = datos_filtrados
    return datos_por_areas

def obtener_productos(datos):
    indice_producto = columnas_deseadas.index("Producto")
    columna_producto = [fila[indice_producto] for fila in datos]
    columna_producto_sin_duplicados = pd.Series(columna_producto).drop_duplicates().dropna().tolist()
    return columna_producto_sin_duplicados

def leer_imagen_base64(nombre_archivo):
    with open(nombre_archivo, "r") as file:
        return file.read()

def construir_correos_enviar(_areas, _contactos_todos_productos, _contactos_por_area, _datos_por_areas):
    correos_enviar = []
    for area in _areas:
        contactos_area = _contactos_por_area[area]
        productos = obtener_productos(_datos_por_areas[area])
        contactos_productos = {producto: _contactos_todos_productos[producto] for producto in productos}
        columnas = _datos_por_areas[area][0]
        datos_ordenados = sorted(_datos_por_areas[area], key=lambda x: (x[-1], x[0]))
        imagen_base64 = leer_imagen_base64("firma.txt")
        toPush = {"area": area, "contacto": contactos_area, "data": datos_ordenados, "productos": contactos_productos, "imagen_pie": imagen_base64}
        correos_enviar.append(toPush)
    return correos_enviar

def generar_productos(datos):
    productos = []
    for producto, info in datos.items():
        nombre = info["nombre"]
        email = info["Correo"]
        productos.append({"producto": producto, "nombre": nombre, "email": email})
    return productos

def send_mail(datos_correo, destinatarios):
    payload = {
        "email": destinatarios,
        "subject": "Estado de Requerimientos - " + datos_correo["area"],
        "template_name": "portfolio.html",
        "attachments": [],
        "params": {
            "destinatario": datos_correo["area"],
            "data": datos_correo["data"],
            "productos": generar_productos(datos_correo["productos"]),
            "imagen_firma": datos_correo["imagen_pie"]
        }
    }
    headers = {"Content-Type": "application/json"}
    response = requests.post(url_servidor, json=payload, headers=headers)
    if response.status_code == 200:
        return response.json()["response"]
    else:
        return None
        
def generar_contactos_area():
    contactos_areas = cargar_data_base(archivo_excel, nombre_hoja, ['Área', 'Correo Electrónico'])
    contactos_areas_df = pd.DataFrame(contactos_areas, columns=['Área', 'Correo Electrónico'])
    contactos_areas_df.dropna(subset=['Correo Electrónico'], inplace=True)
    contactos_areas_df.drop_duplicates(inplace=True)
    contactos_por_area = {}
    for area, contacto in zip(contactos_areas_df['Área'], contactos_areas_df['Correo Electrónico']):
        if area not in contactos_por_area:
            contactos_por_area[area] = []
        contactos_por_area[area].append(contacto)
    with open('salida.json', 'w', encoding='utf-8') as archivo_salida:
        json.dump(contactos_por_area, archivo_salida, ensure_ascii=False)

def crear_carpeta_si_no_existe(carpeta):
    if not os.path.exists(carpeta):
        os.makedirs(carpeta)

def guardar_correos_por_area(correos_enviar):
    for correo in correos_enviar:
        area = correo["area"]
        crear_carpeta_si_no_existe("salida")
        carpeta_area = os.path.join("salida", area)
        crear_carpeta_si_no_existe(carpeta_area)
        correo_sin_data = {
            "area": area,
            "contacto": correo["contacto"],
            "productos": correo["productos"]
        }
        with open(os.path.join(carpeta_area, f'{area}.json'), 'w', encoding='utf-8') as archivo_salida:
            json.dump(correo_sin_data, archivo_salida, ensure_ascii=False)
        guardar_datos(correo["data"], os.path.join(carpeta_area, f'{area}.xlsx'))

def guardar_datos(data, filename):
    column_names = ["Célula Ágil", "Requerimiento", "Aprobación", "Prioridad", "Estado"]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(column_names)
    for row in data:
        ws.append(row)
    wb.save(filename)

def enviar_correos_por_area(correos_enviar):
    for correo in correos_enviar:
        # productos_data = generar_productos(correo["productos"])
        # destinatarios = correo["contacto"]
        destinatarios = []
        # for producto in productos_data:
        #     destinatarios.append(producto['email'])
        destinatarios.append("luis.salazar@kfc.com.ec")
        # destinatarios.append("jaime.rodriguez@kfc.com.ec")
        destinatarios.append("cesar.siguenza@kfc.com.ec")
        destinatarios.append("tatiana.vizcaino@kfc.com.ec")
        destinatarios.append("daira.ona@kfc.com.ec")
        destinatarios.append("mario.molina@kfc.com.ec")
        destinatarios_str = '; '.join(destinatarios)
        send_mail(correo, destinatarios_str)

generar_contactos_area()
contactos_por_area = cargar_contactos_desde_json()    
base = cargar_data_base(archivo_excel, nombre_hoja, columnas_deseadas)
areas = get_areas(base, areas_excluir)
datos_por_areas = obtener_datos_por_areas(base, areas)
productos = cargar_productos_desde_json()
correos_enviar = construir_correos_enviar(areas, productos, contactos_por_area ,datos_por_areas)
guardar_correos_por_area(correos_enviar)
enviar_correos_por_area(correos_enviar)