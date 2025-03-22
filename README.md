# fsmanage y fssearch

## Descripción

Este repositorio contiene dos herramientas para interactuar con la API de Freshservice:  
- `fsmanage.py`: Permite gestionar activos y exportar información en formato Excel.  
- `fssearch.py`: Facilita la búsqueda de usuarios y sus activos asociados en Freshservice.  

Autor: **Rafael Aceituno Álvarez**  

## Requisitos

Antes de ejecutar los scripts, asegúrate de tener instaladas las siguientes dependencias:

```bash
pip install requests pandas openpyxl colorama
```

## Configuración

Ambos scripts requieren una clave de API de Freshservice y un subdominio. Debes editarlos para añadir tu información:

```python
api_key = 'your_api_key'  # Reemplázalo con tu clave de API
subdomain = 'your_subdomain'  # Reemplázalo con tu subdominio
```

## Uso

### `fsmanage.py`

Este script permite obtener información sobre activos de Freshservice y generar reportes en Excel.

#### Ejemplo de uso:

```bash
python fsmanage.py -i 143,150 -o reporte.xlsx -a -d -t -l -u -s -n
```

Parámetros principales:
- `-i` / `--ids`: Lista de IDs de activos a procesar.
- `-o` / `--output`: Nombre del archivo Excel de salida.
- `-a` / `--asset-data`: Obtiene los datos completos de los activos.
- `-d` / `--departments`: Incluye el nombre del departamento.
- `-t` / `--asset-type`: Incluye el tipo de activo.
- `-l` / `--location`: Incluye la ubicación.
- `-u` / `--user`: Incluye la información del usuario.
- `-s` / `--system-os`: Incluye el sistema operativo.
- `-n` / `--machine-ip`: Incluye la dirección IP.

### `fssearch.py`

Este script permite buscar usuarios en Freshservice y obtener sus activos asociados.

#### Ejemplo de uso:

```bash
python fssearch.py -sn "Rafael Aceituno"
```

Parámetro principal:
- `-sn` / `--search-name`: Nombre y apellido del usuario a buscar.

## Notas
- Si la API de Freshservice impone límites de solicitudes (`HTTP 429`), el script esperará automáticamente antes de reintentar.
- Asegúrate de que tu clave API tenga permisos suficientes para acceder a la información requerida.

---

© 2025 Rafael Aceituno Álvarez

