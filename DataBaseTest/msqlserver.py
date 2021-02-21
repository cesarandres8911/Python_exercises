import pyodbc 
# Some other example server values are
# server = 'localhost\sqlexpress' # for a named instance
# server = 'myserver,port' # to specify an alternate port
server = 'desarrolloit\qax' 
database = 'Comunes' 
username = 'CDJPRO\cmeneses' 
aut = 'ActiveDirectoryInteractive'
cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';Authentication='+aut+'TrustServerCertificate=True')
cursor = cnxn.cursor()