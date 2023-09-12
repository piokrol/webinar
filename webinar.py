import adsk.core, adsk.fusion, adsk.cam, traceback
import pyodbc

def run(context):
    ui = None
    try:
        app = adsk.core.Application.get()
        ui  = app.userInterface
        
        # Connect to the database
        conn = pyodbc.connect('DRIVER={SQL Server};'
                              'SERVER=CBR-PKR;'
                              'DATABASE=iLogic;'
                              'Trusted_Connection=yes;')
        cursor = conn.cursor()
        
        # Fetch rows from your table (replace YourTable with your table name)
        cursor.execute('SELECT * FROM tableTop')
        rows = cursor.fetchall()
        
        # Display the content of each row in a MessageBox
        for row in rows:
            ui.messageBox(str(row))
        
        ui.messageBox('Script completed successfully')

    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
