import pandas as pd
import sqlalchemy
from sqlalchemy import create_engine

server_name = 'Doga\\SQLEXPRESS'
database = 'prueba'
driver = 'ODBC+Driver+17+for+SQL+Server'

# Construye la cadena de conexión directamente en la URL de SQLAlchemy
connection_url = f'mssql+pyodbc://{server_name}/{database}?driver={driver}&Trusted_Connection=yes;'
engine = create_engine(connection_url)

excel_file = r'contactos.xlsx'

# Lee todas las hojas del archivo Excel en un solo DataFrame
dfs = pd.read_excel(excel_file, sheet_name=None)

# Combina todos los DataFrames en uno solo
combined_df = pd.concat(dfs.values(), ignore_index=True)

# Define el mapeo de tipos de datos
dtype_mapping = {
    'name_u': sqlalchemy.types.VARCHAR(length=255),
    'email': sqlalchemy.types.VARCHAR(length=255),
    'phone': sqlalchemy.types.VARCHAR(length=255),
}

# Nombre de la tabla en la base de datos
table_name = 'empleados'

# Guarda el DataFrame combinado en la tabla de la base de datos
combined_df.to_sql(table_name, con=engine, if_exists='replace', index=False, dtype=dtype_mapping)
