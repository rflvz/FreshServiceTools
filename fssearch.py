import requests
import argparse
from colorama import Fore, Style, init
import time  # Importar el módulo time para manejar la espera

# Inicializar colorama
init(autoreset=True)

# Configuración de la API
api_key = ''  # Reemplaza 'your_default_api_key' con tu clave real
subdomain = ''  # Define el subdominio aquí
base_url = f'https://{subdomain}.freshservice.com/api/v2/'

def handle_rate_limit(response):
    """
    Maneja el error de límite de solicitudes (HTTP 429).
    Espera un minuto antes de reintentar la solicitud.

    :param response: Objeto de respuesta de la solicitud.
    """
    if response.status_code == 429:
        print(f"{Fore.YELLOW}Advertencia: Límite de solicitudes alcanzado. Esperando 1 minuto antes de reintentar...")
        time.sleep(60)  # Esperar 60 segundos

# Función para buscar un usuario por nombre y apellido
def search_user(first_name, last_name):
    """
    Busca un usuario en la API de Freshservice por nombre y apellido.

    :param first_name: Nombre del usuario.
    :param last_name: Apellido del usuario.
    :return: Diccionario con los datos del usuario o None si no se encuentra.
    """
    url = f'{base_url}requesters?query="first_name:\'{first_name}\'"&query="last_name:\'{last_name}\'"'
    while True:
        try:
            response = requests.get(url, auth=(api_key, ''))
            if response.status_code == 429:
                handle_rate_limit(response)
                continue  # Reintentar la solicitud
            response.raise_for_status()  # Lanza una excepción si el código de estado no es 200
            data = response.json()
            if 'requesters' in data and len(data['requesters']) > 0:
                return data['requesters'][0]  # Devolver el primer usuario encontrado
            else:
                print(f"{Fore.YELLOW}Advertencia: No se encontró ningún usuario con el nombre '{first_name}' y apellido '{last_name}'.")
                return None
        except requests.exceptions.RequestException as e:
            print(f"{Fore.RED}Error al buscar el usuario: {e}")
            return None

# Función para buscar los activos asociados a un usuario por user_id
def search_assets_by_user(user_id):
    """
    Busca los activos asociados a un usuario en la API de Freshservice.

    :param user_id: ID del usuario.
    :return: Lista de activos asociados o una lista vacía si no se encuentran.
    """
    url = f'{base_url}assets?query="user_id:{user_id}"'
    while True:
        try:
            response = requests.get(url, auth=(api_key, ''))
            if response.status_code == 429:
                handle_rate_limit(response)
                continue  # Reintentar la solicitud
            response.raise_for_status()  # Lanza una excepción si el código de estado no es 200
            data = response.json()
            if 'assets' in data and len(data['assets']) > 0:
                return data['assets']  # Devolver la lista de activos
            else:
                print(f"{Fore.YELLOW}Advertencia: No se encontraron activos asociados al usuario con ID {user_id}.")
                return []
        except requests.exceptions.RequestException as e:
            print(f"{Fore.RED}Error al buscar los activos asociados al usuario: {e}")
            return []

# Función principal
def main(search_name):
    # Unir los argumentos capturados como nombre completo
    full_name = ' '.join(search_name)
    try:
        first_name, last_name = full_name.split(' ', 1)  # Dividir en dos partes: nombre y apellido
    except ValueError:
        print(f"{Fore.RED}Error: Debes proporcionar un nombre y un apellido separados por un espacio.")
        return

    # Buscar el usuario por nombre y apellido
    user = search_user(first_name, last_name)
    if not user:
        return  # Terminar si no se encuentra el usuario

    # Extraer información del usuario
    user_id = user.get('id', 'Unknown')
    user_first_name = user.get('first_name', 'Unknown')
    user_last_name = user.get('last_name', 'Unknown')

    print(f"{Fore.GREEN}Usuario encontrado:")
    print(f"  Nombre: {user_first_name}")
    print(f"  Apellido: {user_last_name}")
    print(f"  ID de Usuario: {user_id}")

    # Buscar los activos asociados al usuario
    assets = search_assets_by_user(user_id)
    if assets:
        print(f"{Fore.CYAN}Activos asociados al usuario:")
        for asset in assets:
            asset_name = asset.get('name', 'Unknown')
            display_id = asset.get('display_id', 'Unknown')
            print(f"  - Nombre del Activo: {asset_name}")
            print(f"    ID de Visualización: {display_id}")
    else:
        print(f"{Fore.YELLOW}No se encontraron activos asociados al usuario.")

if __name__ == "__main__":
    # Configurar argumentos de línea de comandos
    parser = argparse.ArgumentParser(
        description=f"{Fore.CYAN}Buscador de usuarios y activos asociados en Freshservice."
    )
    parser.add_argument(
        '-sn', '--search-name',
        nargs='+',  # Capturar múltiples palabras como una lista
        required=True,
        help=f"{Fore.GREEN}Nombre completo del usuario a buscar (por ejemplo: Rafael Aceituno)."
    )
    args = parser.parse_args()

    # Ejecutar la función principal con los argumentos especificados
    main(args.search_name)
