import adsk.core, adsk.fusion, adsk.cam, traceback
import pyodbc

def connect_to_database():
    # Connect to the database
    conn = pyodbc.connect('DRIVER={SQL Server};'
                          'SERVER=CBR-PKR;'
                          'DATABASE=iLogic;'
                          'Trusted_Connection=yes;')
    cursor = conn.cursor()
    
    # Fetch rows from your table
    cursor.execute('SELECT * FROM tableTop')
    rows = cursor.fetchall()
    
    return rows

def run(context):
    try:
        app = adsk.core.Application.get()
        ui = app.userInterface
        logger = UiLogger(True)

        logger.print("Start")

        # Get the rows from the database
        rows = connect_to_database()

        # Get the active document
        doc = adsk.fusion.FusionDocument.cast(app.activeDocument)

        # Get the parameters
        params = doc.design.userParameters

        # Process each row from the database
        for row in rows:
            process_row(row, params, logger, ui)

    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))


def process_row(row, params, logger, ui):
    order_code = row.orderCode
    parts = order_code.split('-')

    # Set the parameters
    set_parameter(params, "height_user", float(parts[1]))
    set_parameter(params, "width_user", float(parts[2]))
    set_parameter(params, "thickness_user", float(parts[3]))
    set_parameter(params, "radius_user", round(float(parts[4][1:])))
    set_parameter(params, "chamfer_user", round(float(parts[5][2:])))

    # Log the parameter values
    log_parameter_values(row, parts, logger, ui)

    # Save the document
    app = adsk.core.Application.get()
    doc = app.activeDocument

    # Get the current document's project and folder
    folder = doc.dataFile.parentFolder

    # Save the document to the folder
    doc.saveAs(row.name + "_modified", folder, "Modified parameters", "")


    ui.messageBox(row.name)



def set_parameter(params, param_name, value):
    param = params.itemByName(param_name)
    if param is not None:
        param.expression = str(value)
    else:
        logger.print(f"{param_name} parameter not found!")


def log_parameter_values(row, parts, logger, ui):
    logger.print("Name:" + row.name)
    logger.print("Description: " + row.description)
    logger.print("Order Code: " + row.orderCode)
    logger.print("Height: " + str(float(parts[1])))
    logger.print("Width: " + str(float(parts[2])))
    logger.print("Thickness: " + str(float(parts[3])))
    logger.print("Radius: " + str(round(float(parts[4][1:]))))
    logger.print("ID: " + str(row.id))
    logger.print("Chamfer: " + str(round(float(parts[5][2:]))))
    ui.messageBox(row.name)



class UiLogger:
    def __init__(self, forceUpdate):  
        app = adsk.core.Application.get()
        ui  = app.userInterface
        palettes = ui.palettes
        self.textPalette = palettes.itemById("TextCommands")
        self.forceUpdate = forceUpdate
        self.textPalette.isVisible = True 
    
    def print(self, text):       
        self.textPalette.writeText(text)
        if (self.forceUpdate):
            adsk.doEvents() 

class FileLogger:
    def __init__(self, filePath): 
        try:
            open(filePath, 'a').close()
        
            self.filePath = filePath
        except:
            raise Exception("Could not open/create file = " + filePath)

    def print(self, text):
        with open(self.filePath, 'a') as txtFile:
            txtFile.writelines(text + '\r\n')