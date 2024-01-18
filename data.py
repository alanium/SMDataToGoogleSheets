import requests
import json

with open('config.json', 'r') as config_file:
    constants = json.load(config_file)

KEY = constants["KEY"]
TOKEN = constants["TOKEN"]
NOTION_VERSION = '2022-06-28'


def notion_request(url, method, payload=None):
    headers = {
        "Authorization": "Bearer " + TOKEN,
        "Content-Type": "application/json",
        "Notion-Version": NOTION_VERSION,
    }
    
    if method == "GET":
        response = requests.get(url, headers=headers)
    elif method == "POST":
        response = requests.post(url, json=payload, headers=headers)
    elif method == "PATCH":
        response = requests.patch(url, json=payload, headers=headers)
    elif method == "DELETE":
        response = requests.delete(url, headers=headers)
    else:
        raise ValueError("Unsupported HTTP method")

    return response

def read(db_id, num_pages=None):
    url = f"https://api.notion.com/v1/databases/{db_id}/query"
    get_all = num_pages is None
    page_size = 100 if get_all else num_pages
    results = []

    payload = {"page_size": page_size}
    response = notion_request(url, "POST", payload)
    data = response.json()
    results.extend(data["results"])

    while data["has_more"] and get_all:
        payload = {"page_size": page_size, "start_cursor": data["next_cursor"]}
        response = notion_request(url, "POST", payload)
        data = response.json()
        results.extend(data["results"])

    return results

def create(properties, db_id):
    url = "https://api.notion.com/v1/pages"
    payload = {
        "parent": {"database_id": db_id},
        "properties": properties
    }
    response = notion_request(url, "POST", payload)

    if response.status_code == 200:
        print("Entry successfully created")
        data = response.json()
        # Obtener el ID del objeto creado y retornarlo
        return data.get("id")
    else:
        print("Error creating entry")
        try:
            error_message = response.json().get("message", "Unknown error")
            print(error_message)
        except ValueError:
            error_message = response.text  # En caso de que no sea JSON válido
            print(error_message)
        return {"error": error_message}

def update(item_id, properties):
    url = f"https://api.notion.com/v1/pages/{item_id}"
    payload = {"properties": properties}
    response = notion_request(url, "PATCH", payload)
    
    if response.status_code == 200:
        print("Entry updated successfully")
    else:
        print("Error updating entry")
        print(response.text)

    return response

def delete(item_id):
    url = f"https://api.notion.com/v1/blocks/{item_id}"
    response = notion_request(url, "DELETE")
    
    if response.status_code == 200:
        print("Entry deleted successfully")
    else:
        print("Error deleting entry")
        print(response.text)

    return response

def read_item(item_id):
    url = f"https://api.notion.com/v1/pages/{item_id}"
    response = notion_request(url, "GET")

    if response.status_code == 200:
        data = response.json()
        return data
    else:
        print("Error reading item")
        print(response.text)
        return None