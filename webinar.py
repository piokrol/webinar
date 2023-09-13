import adsk.core, adsk.fusion, adsk.cam, traceback
import pyodbc
import adsk.cam


programName = 100


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

    doc.saveAs(row.name + "_modified", folder, "Modified parameters", "")

    create_cam_setup_and_set_wcs(doc, ui, row.name)

    # Save the document to the folder
    doc.save("cam update")

    
    ui.messageBox(row.name + " sekces")




def create_cam_setup_and_set_wcs(doc, ui, name):
    global programName
    # switch to manufacturing space
    camWS = ui.workspaces.itemById('CAMEnvironment') 
    camWS.activate()
    # get the CAM product
    products = doc.products
      
    # get the tool libraries from the library manager
    camManager = adsk.cam.CAMManager.get()
    libraryManager = camManager.libraryManager
    toolLibraries = libraryManager.toolLibraries
    # we can use a library URl directly if we know its address (here we use Fusion's Metric sample library)
    url = adsk.core.URL.create('systemlibraryroot://Samples/Milling Tools (Metric).json')
    
    # load tool library
    toolLibrary = toolLibraries.toolLibraryAtURL(url)
    # create some variables for the milling tools which will be used in the operations
    faceTool = None
    adaptiveTool = None
    
    # searching the face mill and the bull nose using a loop for the roughing operations
    for tool in toolLibrary:
        # read the tool type
        toolType = tool.parameters.itemByName('tool_type').value.value 
        
        # select the first face tool found
        if toolType == 'face mill' and not faceTool:
            faceTool = tool  
        
        # search the roughing tool
        elif toolType == 'bull nose end mill' and not adaptiveTool:
            # we look for a bull nose end mill tool larger or equal to 10mm but less than 14mm
            diameter = tool.parameters.itemByName('tool_diameter').value.value
            if diameter >= 1.0 and diameter < 1.4: 
                adaptiveTool = tool
        # exit when the 2 tools are found
        if faceTool and adaptiveTool:
            break
    
    #################### create setup ####################
    cam = adsk.cam.CAM.cast(products.itemByProductType("CAMProductType"))
    setups = cam.setups

    # delete all setups
    for setup in setups:
        setup.deleteMe()

    setupInput = setups.createInput(adsk.cam.OperationTypes.MillingOperation)
    # create a list for the models to add to the setup Input
    models = [] 
    part = cam.designRootOccurrence.bRepBodies.item(0)
    # add the part to the model list
    models.append(part) 
    # pass the model list to the setup input
    setupInput.models = models 
    # create the setup
    setup = setups.add(setupInput) 
    # change some properties of the setup
    setup.name = 'CAM Basic Script Sample'  
    setup.stockMode = adsk.cam.SetupStockModes.RelativeBoxStock
    # set offset mode
    setup.parameters.itemByName('job_stockOffsetMode').expression = "'simple'"
    # set offset stock side
    setup.parameters.itemByName('job_stockOffsetSides').expression = '0 mm'
    # set offset stock top
    setup.parameters.itemByName('job_stockOffsetTop').expression = '2 mm'
    # set setup origin
    setup.parameters.itemByName('wcs_origin_boxPoint').value.value = 'top 1'
    
    #################### face operation ####################
    # create a face operation input
    input = setup.operations.createInput('face')
    input.tool = faceTool
    input.displayName = 'Face Operation'       
    input.parameters.itemByName('tolerance').expression = '0.01 mm'
    input.parameters.itemByName('stepover').expression = '0.75 * tool_diameter'
    input.parameters.itemByName('direction').expression = "'climb'"
    # add the operation to the setup
    faceOp = setup.operations.add(input)



    # generate toolpaths for the operations in the setup
    cam.generateAllToolpaths(True)

    
    # post-process the operations in the setup
    # specify the program name, post configuration to use and a folder destination for the nc file
    programName = programName + 1
    outputFolder = 'C:\\Users\\pkrol\\Desktop\\Szkoleniowe\\webinary\\Fusion API'
    # set the post configuration to use based on Operation Type of the first Setup
    firstSetupOperationType = cam.setups.item(0).operationType
    postConfig = 'C:\\Users\\pkrol\\Desktop\\Szkoleniowe\\webinary\\Fusion API\\haas.cps'
    # prompt the user with an option to view the resulting NC file.
    viewResults = ui.messageBox('Pokazać kod po wygenerowaniu?', 'Post NC Files',
                                adsk.core.MessageBoxButtonTypes.YesNoButtonType,
                                adsk.core.MessageBoxIconTypes.QuestionIconType)
    if viewResults == adsk.core.DialogResults.DialogNo:
        viewResult = False
    else:
        viewResult = True
    # specify the NC file output units
    units = adsk.cam.PostOutputUnitOptions.DocumentUnitsOutput
#       units = adsk.cam.PostOutputUnitOptions.InchesOutput
#       units = adsk.cam.PostOutputUnitOptions.MillimetersOutput
    # create the postInput object
    postInput = adsk.cam.PostProcessInput.create(str(programName), postConfig, outputFolder, units)
    postInput.isOpenInEditor = viewResult
    # create the post properties
    postProperties = adsk.core.NamedValues.create()
    # create the disable sequence number property
    disableSequenceNumbers = adsk.core.ValueInput.createByBoolean(False)
    postProperties.add("showSequenceNumbers", disableSequenceNumbers)
    # add the post properties to the post process input
    postInput.postProperties = postProperties
    # set the value of scenario to 1, 2 or 3 to post all, post the first setup, or post only the first operation of the first setup.
    scenario = 3
    if scenario == 1:
        ui.messageBox('Cała scieżka narzedzia została wygenerowana')
        cam.postProcessAll(postInput)
    elif scenario == 2:
        ui.messageBox('Toolpaths in the first Setup will be posted')
        setups = cam.setups
        setup = setups.item(0)
        cam.postProcess(setup, postInput)
    elif scenario == 3:
        ui.messageBox('The first Toolpath in the first Setup will be posted')
        setups = cam.setups
        setup = setups.item(0)
        operations = setup.allOperations
        operation = operations.item(0)
        if operation.hasToolpath == True:
            cam.postProcess(operation, postInput)
        else:
            ui.messageBox('Operation has no toolpath to post')
            return
    # switch back to design space
    designWS = ui.workspaces.itemById("FusionSolidEnvironment")
    designWS.activate()
        
    ui.messageBox('Post processing is complete. The results have been written to:\n"' + 'C:\\Users\\pkrol\\Desktop\\Szkoleniowe\\webinary\\Fusion API' + '.nc"') 


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
    ui.messageBox(row.name + row.orderCode)



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