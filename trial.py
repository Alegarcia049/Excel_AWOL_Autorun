import pandas as pd

from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext

# Crea un objeto de contexto de autenticación con tu usuario y contraseña de SharePoint
usuario = "AGARC611"
contrasena = "!BERNARDITO2023"
url_sitio = "https://azureford.sharepoint.com/:f:/r/sites/site-management-arg"
contexto_autenticacion = AuthenticationContext(url_sitio)
contexto_autenticacion.acquire_token_for_user(usuario, contrasena)

# Crea un objeto de contexto de cliente con la URL del sitio y el objeto de autenticación
contexto_cliente = ClientContext(url_sitio, contexto_autenticacion)

# Obtiene la dirección URL del archivo
ruta_archivo = "/ITI%20Documents/INVENTARIO%20MASTER?csf=1&web=1&e=P7PmYx/Inventario_MASTER_Leasing_v137%20del%2010-07-2023.xlsx?d=wdce5c1db858b4c4c88901fff873680e8&csf=1&web=1&e=t1RtJu"
archivo = contexto_cliente.web.get_file_by_server_relative_url(ruta_archivo)
contexto_cliente.load(archivo)
contexto_cliente.execute_query()
direccion_archivo = archivo.serverRelativeUrl

# Imprime la dirección URL del archivo
print(f'La dirección URL del archivo es: {direccion_archivo}')


# Lee el archivo de Excel
df = pd.read_excel(direccion_archivo)

# Busca la primera celda que contiene la palabra "SERIE" en cualquier columna
primera_celda = df.where(df == 'SERIE').first_valid_index()

# Obtiene el nombre de la columna donde se encuentra la celda
columna = df.columns[df.columns.get_loc(primera_celda[1])]

# Encuentra la última celda no vacía en la columna encontrada
ultima_celda = df[columna].last_valid_index()

# Obtiene el índice de la fila donde se encuentra la celda
primera_fila = primera_celda[0] + 1

# Selecciona el rango de celdas de la columna encontrada
rango_celdas = df.loc[primera_fila:ultima_celda, columna]

# Guarda los datos en un archivo de texto
rango_celdas.to_csv('datos.csv', index=False, header=False)

# Imprime un mensaje de confirmación
print('Los datos se han guardado correctamente en el archivo datos.txt')

input('Presiona Enter para salir...')

