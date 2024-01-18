import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import data
import json
import difflib
from flask import Flask, render_template

app = Flask(__name__)

with open('config.json', 'r') as config_file:
    constants = json.load(config_file)

SALESMEETINGSDB = constants['SM']

credentials_file = 'client_secret.json'
spreadsheet_id = '1xX6QGLXL_0_5zgurpF6rEBqDLwfszXSiPnDPrm59yOU'
worksheet_name = 'Copy of January 2024'


def compare(str_1, str_2, umbral=0.6):
    comparador = difflib.SequenceMatcher(None, str_1, str_2)
    similitud = comparador.ratio()
    return similitud >= umbral

def update_google_sheets(row_number, tag_content, name_content, tag_column_letter, name_column_letter):
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        credentials = ServiceAccountCredentials.from_json_keyfile_name(credentials_file, scope)
        client = gspread.authorize(credentials)

        spreadsheet = client.open_by_key(spreadsheet_id)
        worksheet = spreadsheet.worksheet(worksheet_name)

        # Update Tag column
        tag_cell = f'{tag_column_letter}{row_number}'
        worksheet.update_acell(tag_cell, tag_content)

        # Update Name column
        name_cell = f'{name_column_letter}{row_number}'
        worksheet.update_acell(name_cell, name_content)

        print(f"Row {row_number} updated successfully.")

    except gspread.exceptions.APIError as e:
        if 'The caller does not have permission' in str(e):
            raise PermissionError("No tienes permisos para actualizar la hoja de cálculo.")
        else:
            raise

def get_google_sheets():

    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        credentials = ServiceAccountCredentials.from_json_keyfile_name(credentials_file, scope)
        client = gspread.authorize(credentials)

        spreadsheet = client.open_by_key(spreadsheet_id)
        worksheet = spreadsheet.worksheet(worksheet_name)

        return worksheet.get_all_values()

    except gspread.exceptions.APIError as e:
        if 'The caller does not have permission' in str(e):
            raise PermissionError("No tienes permisos para acceder a la hoja de cálculo.")
        else:
            raise

def buscar_nombre_en_notion(nombre, all_data):
    for i in all_data:
        if 'properties' in i and 'Name' in i['properties']:
            notion_name_list = i['properties']['Name']['title']
            
            # Verificar si la lista 'title' no está vacía
            if notion_name_list:
                notion_name = notion_name_list[0]['plain_text']
                
                if compare(nombre, notion_name):
                    tags = i['properties']['Tags']['multi_select']
                    note = i['properties']['Text']['rich_text']
                    content = 'empty'
                    tag = 'empty'

                    if note:
                        content = note[0]['text']['content']

                    if tags:
                        tag = tags[0]['name']

                    return content, tag

    return False

def procesar_google_sheets():
    google_sheets_data = get_google_sheets()
    notion_data = data.read(SALESMEETINGSDB)

    # Obtener los títulos de la hoja de cálculo
    titulos = google_sheets_data[4]

    # Obtener índice de la columna 'Nombre'
    indice_nombre = titulos.index('Nombre')

    # Inicializar el contador de celdas vacías
    celdas_vacias_consecutivas = 0

    # Iterar sobre las filas de la hoja de cálculo a partir de la quinta fila (índice 4)
    for row_number, fila in enumerate(google_sheets_data[5:], start=5):
        nombre_fila = fila[indice_nombre]

        # Verificar si la celda de la columna 'Nombre' está vacía
        if not nombre_fila:
            celdas_vacias_consecutivas += 1

            # Terminar el programa si se encuentran tres celdas vacías consecutivas
            if celdas_vacias_consecutivas == 3:
                print("Se encontraron tres celdas vacías consecutivas. Terminando el programa.")
                break
            else:
                continue
        else:
            # Reiniciar el contador de celdas vacías consecutivas
            celdas_vacias_consecutivas = 0

        # Buscar el nombre en Notion
        notion_result = buscar_nombre_en_notion(nombre_fila, notion_data)

        if notion_result:
            print(f"Coincidencia encontrada para el nombre: {nombre_fila}")

            # Extract content and tag from Notion result
            content, tag = notion_result

            # Update Google Sheets
            update_google_sheets(row_number + 1, tag, content, 'M', 'N')
        else:
            print(f"No se encontró coincidencia para el nombre: {nombre_fila}")

@app.route('/')
def index():
    try:
        procesar_google_sheets()
        mensaje = "La hoja de cálculo se ha actualizado exitosamente."
    except PermissionError as e:
        mensaje = str(e)
    except Exception as e:
        mensaje = f"Error durante la actualización: {e}"

    return render_template('index.html', mensaje=mensaje)

# Agrega el resto de tu código aquí

if __name__ == '__main__':
    app.run(debug=True)