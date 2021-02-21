import pyodbc 
conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=desarrolloit\qax;'
                      'Database=Comunes;'
                      'Trusted_Connection=yes;')

cursor = conn.cursor()
cursor.execute('SELECT * FROM a_CategoriasSoftware')

for row in cursor:
    print(row)