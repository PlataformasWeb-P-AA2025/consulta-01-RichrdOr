import pandas as pd                  # Libreria para leer archivos Excel
from pymongo import MongoClient      # Libreria para conectarse a MongoDB
import os                            # Para manejar rutas de archivos

# Conexion a MongoDB local
client = MongoClient("mongodb://localhost:27017/")

# Crear o acceder a la base de datos 'tennis_db'
db = client["tennis_db"]

# Crear o acceder a la coleccion 'matches'
collection = db["matches"]

# Nombres de las columnas del archivo Excel
columns = [
    "ATP", "Location", "Tournament", "Date", "Series", "Court", "Surface", "Round", "Best of",
    "Winner", "Loser", "WRank", "LRank", "WPts", "LPts",
    "W1", "L1", "W2", "L2", "W3", "L3", "W4", "L4", "W5", "L5",
    "Wsets", "Lsets", "Comment",
    "B365W", "B365L", "PSW", "PSL", "MaxW", "MaxL", "AvgW", "AvgL"
]

# Carpeta donde estan los archivos
data_folder = "data"

# Lista de archivos a procesar
files = ["2022.xlsx", "2023.xlsx"]

# Leer cada archivo y procesarlo
for file in files:
    # Ruta completa al archivo
    path = os.path.join(data_folder, file)
    
    # Leer el archivo Excel
    df = pd.read_excel(path, header=None, names=columns, skiprows=1)
    
    # Reemplazar comas por puntos en las columnas decimales
    decimal_cols = ['B365W', 'B365L', 'PSW', 'PSL', 'MaxW', 'MaxL', 'AvgW', 'AvgL']
    for col in decimal_cols:
        df[col] = df[col].astype(str).str.replace(",", ".", regex=False)
        df[col] = pd.to_numeric(df[col], errors='coerce')
    
    # Convertir el dataframe a diccionarios
    records = df.to_dict(orient='records')
    
    # Insertar los datos en la base de datos
    collection.insert_many(records)

print("Datos insertados correctamente en MongoDB")

# Consulta 1: Partidos donde gano 'Kwon S.W.'
print("\n Partidos donde gano 'Kwon S.W.':")
for match in collection.find({"Winner": "Kwon S.W."}):
    print(f"{match['Date']} - {match['Winner']} vencio a {match['Loser']}")

# Consulta 2: Partidos jugados a 3 sets (2-1)
print("\n Partidos que se jugaron a 3 sets exactos:")
query = {"Wsets": 2, "Lsets": 1}
results = collection.find(query)

for match in results:
    print(f"{match['Date'].date()} - {match['Winner']} vencio a {match['Loser']} (Sets: 2-1)")
