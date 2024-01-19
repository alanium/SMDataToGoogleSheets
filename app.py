import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import data
import json
import difflib
from flask import Flask, render_template
from threading import Thread

app = Flask(__name__)

with open('config.json', 'r') as config_file:
    constants = json.load(config_file)

SALESMEETINGSDB = constants['SM']

credentials_file = 'client_secret.json'
spreadsheet_id = '1xX6QGLXL_0_5zgurpF6rEBqDLwfszXSiPnDPrm59yOU'
worksheet_name = 'January 2024'


# Controller

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

    # Obtener índice de las columnas 'Nombre', 'Tags' y 'Notes'
    indice_nombre = titulos.index('Nombre')
    indice_tags = titulos.index('Tags')
    indice_notes = titulos.index('Notes')

    # Inicializar el contador de celdas vacías
    celdas_vacias_consecutivas = 0

    # Iterar sobre las filas de la hoja de cálculo a partir de la quinta fila (índice 4)
    for row_number, fila in enumerate(google_sheets_data[5:], start=5):
        nombre_fila = fila[indice_nombre]
        tags_fila = fila[indice_tags]
        notes_fila = fila[indice_notes]

        # Verificar si alguna de las columnas 'Tags' o 'Notes' contiene algo
        if tags_fila or notes_fila:
            print(f"Skipping row {row_number} because Tags or Notes is not empty.")
            continue

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

        # Buscar el nombre en Notion solo si 'Tags' y 'Notes' están vacías
        notion_result = buscar_nombre_en_notion(nombre_fila, notion_data)

        if notion_result:
            content, tag = notion_result
            update_google_sheets(row_number + 1, tag, content, 'M', 'N')
        else:
            update_google_sheets(row_number + 1, 'not found', 'not found', 'M', 'N')

def parse_google_sheets_data():
    data = get_google_sheets()
    headers = data[4] 
    result = []

    for row in data[5:]:
        item = {}
        for i, header in enumerate(headers):
            if header != '':
                item[header.lower()] = row[i]
        result.append(item)

    return result

def combine_similar_names(dashboard):
    keys = list(dashboard.keys())
    for i in range(len(keys) - 1):
        for j in range(i + 1, len(keys)):
            if compare(keys[i], keys[j]):
                # Combinar las estadísticas de ambos y eliminar la clave antigua
                dashboard[keys[i]] = {
                    metric: dashboard[keys[i]][metric] + dashboard[keys[j]][metric] for metric in dashboard[keys[i]]
                }
                del dashboard[keys[j]]
    return dashboard

def generate_dashboard():
    data = parse_google_sheets_data()
    dashboard = {}
    total_dashboard = {
        'Total Appointments': 0,
        'Total Visited': 0,
        'Total Cancelled': 0,
        'Total Qualified Appointments': 0,
        'Total Not Qualified Appointments': 0,
        'Total Sold Appointments': 0,
        'Total Estimated Money': 0,
        'Total Sold Money': 0
    }

    for item in data:
        sales_person = item.get('sales person ', 'Unknown')

        if sales_person not in dashboard:
            dashboard[sales_person] = {
                'Total Appointments': 0,
                'Total Visited': 0,
                'Total Cancelled': 0,
                'Total Qualified Appointments': 0,
                'Total Not Qualified Appointments': 0,
                'Total Sold Appointments': 0,
                'Total Estimated Money': 0,
                'Total Sold Money': 0
            }

        dashboard[sales_person]['Total Appointments'] += 1

        if item.get('appt_status') == 'Cancelled':
            dashboard[sales_person]['Total Cancelled'] += 1

        if item.get('appt_status') == 'Visited':
            dashboard[sales_person]['Total Visited'] += 1

            if item.get('tags') == 'QUALIFIED':
                dashboard[sales_person]['Total Qualified Appointments'] += 1

            if item.get('tags') != 'QUALIFIED':
                dashboard[sales_person]['Total Not Qualified Appointments'] += 1

        if item.get('status') == 'SOLD':
            dashboard[sales_person]['Total Sold Appointments'] += 1
            sold_money = float(item.get('latest estimate total', '').replace('$', '').replace(',', '')) if item.get('latest estimate total') else 0
            dashboard[sales_person]['Total Sold Money'] += sold_money

        estimated_money = float(item.get('latest estimate total', '').replace('$', '').replace(',', '')) if item.get('latest estimate total') else 0
        dashboard[sales_person]['Total Estimated Money'] += estimated_money

        # Actualizar totales generales
        total_dashboard['Total Appointments'] += 1

        if item.get('appt_status') == 'Cancelled':
            total_dashboard['Total Cancelled'] += 1

        if item.get('appt_status') == 'Visited':
            total_dashboard['Total Visited'] += 1

            if item.get('tags') == 'QUALIFIED':
                total_dashboard['Total Qualified Appointments'] += 1

            if item.get('tags') != 'QUALIFIED':
                total_dashboard['Total Not Qualified Appointments'] += 1

        if item.get('status') == 'SOLD':
            total_dashboard['Total Sold Appointments'] += 1
            total_dashboard['Total Sold Money'] += sold_money

        total_dashboard['Total Estimated Money'] += estimated_money

    # Cambiar la clave vacía a 'noname'
    dashboard['noname'] = dashboard.pop('')

    # Combina nombres similares
    dashboard = combine_similar_names(dashboard)

    # editar total con noname
    noname_tot_appts = dashboard.pop('noname', {}).get('Total Appointments', 0)
    noname_tot_visited = dashboard.pop('noname', {}).get('Total Visited', 0)
    noname_tot_not_q = dashboard.pop('noname', {}).get('Total Not Qualified Appointments', 0)

    total_dashboard['Total Appointments'] -= noname_tot_appts
    total_dashboard['Total Cancelled'] += noname_tot_appts
    total_dashboard['Total Visited'] -= 5
    total_dashboard['Total Not Qualified Appointments'] -= 5
    

    # Agregar totales generales al diccionario final
    dashboard['Total'] = total_dashboard
    
    return dashboard

def update_google_sheets_row(worksheet, row_number, name, values):
    name_column = 'A'
    total_appts_column = 'B'
    total_visited_column = 'C'
    total_cancelled_column = 'D'
    total_qualified_column = 'E'
    total_not_qualified_column = 'F'
    total_solds_column = 'G'
    total_estimated_column = 'H'
    total_sold_money_column = 'I'

    # Define el formato normal para las celdas normales
    cell_format_normal = {
        "textFormat": {
            "bold": False
        }
    }

    # Define el formato en negrita para las celdas de 'Total'
    cell_format_bold = {
        "textFormat": {
            "bold": True
        }
    }

    # Actualiza las celdas según el nombre
    if name == 'Total':
        worksheet.format(f'{name_column}{row_number}', cell_format_bold)
        worksheet.format(f'{total_appts_column}{row_number}', cell_format_bold)
        worksheet.format(f'{total_visited_column}{row_number}', cell_format_bold)
        worksheet.format(f'{total_cancelled_column}{row_number}', cell_format_bold)
        worksheet.format(f'{total_qualified_column}{row_number}', cell_format_bold)
        worksheet.format(f'{total_not_qualified_column}{row_number}', cell_format_bold)
        worksheet.format(f'{total_solds_column}{row_number}', cell_format_bold)
        worksheet.format(f'{total_estimated_column}{row_number}', cell_format_bold)
        worksheet.format(f'{total_sold_money_column}{row_number}', cell_format_bold)
    else:
        # Actualiza las celdas normales con formato normal
        worksheet.format(f'{name_column}{row_number}', cell_format_normal)
        worksheet.format(f'{total_appts_column}{row_number}', cell_format_normal)
        worksheet.format(f'{total_visited_column}{row_number}', cell_format_normal)        
        worksheet.format(f'{total_cancelled_column}{row_number}', cell_format_normal)
        worksheet.format(f'{total_qualified_column}{row_number}', cell_format_normal)
        worksheet.format(f'{total_not_qualified_column}{row_number}', cell_format_normal)
        worksheet.format(f'{total_solds_column}{row_number}', cell_format_normal)
        worksheet.format(f'{total_estimated_column}{row_number}', cell_format_normal)
        worksheet.format(f'{total_sold_money_column}{row_number}', cell_format_normal)

    # Actualiza los datos
    worksheet.update_acell(f'{name_column}{row_number}', name)
    worksheet.update_acell(f'{total_appts_column}{row_number}', values['Total Appointments'])
    worksheet.update_acell(f'{total_visited_column}{row_number}', values['Total Visited'])
    worksheet.update_acell(f'{total_cancelled_column}{row_number}', values['Total Cancelled'])
    worksheet.update_acell(f'{total_qualified_column}{row_number}', values['Total Qualified Appointments'])
    worksheet.update_acell(f'{total_not_qualified_column}{row_number}', values['Total Not Qualified Appointments'])
    worksheet.update_acell(f'{total_solds_column}{row_number}', values['Total Sold Appointments'])
    worksheet.update_acell(f'{total_estimated_column}{row_number}', values['Total Estimated Money'])
    worksheet.update_acell(f'{total_sold_money_column}{row_number}', values['Total Sold Money'])

def update_google_sheets_stats():
    dashboard = generate_dashboard()
    worksheet_name = 'stats'
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        credentials = ServiceAccountCredentials.from_json_keyfile_name(credentials_file, scope)
        client = gspread.authorize(credentials)

        spreadsheet = client.open_by_key(spreadsheet_id)
        worksheet = spreadsheet.worksheet(worksheet_name)

        row_number = 2  # Assuming data starts from row 2
        for name, values in dashboard.items():
            update_google_sheets_row(worksheet, row_number, name, values)
            row_number += 1

        print("Data updated successfully.")

    except gspread.exceptions.APIError as e:
        if 'The caller does not have permission' in str(e):
            raise PermissionError("No tienes permisos para actualizar la hoja de cálculo.")
        else:
            raise


# segundo plano
def ejecutar_proceso_google_sheets():
    try:
        procesar_google_sheets()
        mensaje = "La hoja de cálculo se ha actualizado exitosamente."
    except PermissionError as e:
        mensaje = str(e)
    except Exception as e:
        mensaje = f"Error durante la actualización: {e}"

    print(mensaje)

def ejecutar_proceso_update_stats():
    try:
        update_google_sheets_stats()
        mensaje = "La hoja de cálculo se ha actualizado exitosamente."
    except PermissionError as e:
        mensaje = str(e)
    except Exception as e:
        mensaje = f"Error durante la actualización: {e}"

    print(mensaje)


# Routes

@app.route('/')
def index():
    t = Thread(target=ejecutar_proceso_google_sheets)
    t.start()
    
    return render_template('index.html', mensaje="Ya puede cerrar esta pestaña :)", worksheet=worksheet_name)

@app.route('/update_stats')
def update_stats():
    t = Thread(target=ejecutar_proceso_update_stats)
    t.start()
    
    return render_template('index.html', mensaje="Ya puede cerrar esta pestaña :)", worksheet='stats')


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)

