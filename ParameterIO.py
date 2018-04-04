#Author-Wayne Brill
#Description-Allows you to select a CSV (comma seperated values) file and then edits existing Attributes. Also allows you to write parameters to a file

import adsk.core, adsk.fusion, traceback, csv

commandId = 'ParamsFromCSV'
workspaceToUse = 'FusionSolidEnvironment'
panelToUse = 'SolidModifyPanel'

# global set of event handlers to keep them referenced for the duration of the command
handlers = []

def commandDefinitionById(id):
    app = adsk.core.Application.get()
    ui = app.userInterface
    if not id:
        ui.messageBox('commandDefinition id is not specified')
        return None
    commandDefinitions_ = ui.commandDefinitions
    commandDefinition_ = commandDefinitions_.itemById(id)
    return commandDefinition_

def commandControlByIdForPanel(id):
    app = adsk.core.Application.get()
    ui = app.userInterface
    if not id:
        ui.messageBox('commandControl id is not specified')
        return None
    workspaces_ = ui.workspaces
    modelingWorkspace_ = workspaces_.itemById(workspaceToUse)
    toolbarPanels_ = modelingWorkspace_.toolbarPanels
    toolbarPanel_ = toolbarPanels_.itemById(panelToUse)
    toolbarControls_ = toolbarPanel_.controls
    toolbarControl_ = toolbarControls_.itemById(id)
    return toolbarControl_

def destroyObject(uiObj, tobeDeleteObj):
    if uiObj and tobeDeleteObj:
        if tobeDeleteObj.isValid:
            tobeDeleteObj.deleteMe()
        else:
            uiObj.messageBox('tobeDeleteObj is not a valid object')

def run(context):
    ui = None
    try:
        commandName = 'Import/Export Parameters (CSV)'
        commandDescription = 'Import parameters from or export them to a CSV (Comma Separated Values) file\n'
        commandResources = './resources/command'

        app = adsk.core.Application.get()
        ui = app.userInterface

        class CommandExecuteHandler(adsk.core.CommandEventHandler):
            def __init__(self):
                super().__init__()
            def notify(self, args):
                try:
                  updateParamsFromCSV()
                except:
                    if ui:
                        ui.messageBox('command executed failed:\n{}'.format(traceback.format_exc()))

        class CommandCreatedEventHandlerPanel(adsk.core.CommandCreatedEventHandler):
            def __init__(self):
                super().__init__() 
            def notify(self, args):
                try:
                    cmd = args.command
                    onExecute = CommandExecuteHandler()
                    cmd.execute.add(onExecute)
                    # keep the handler referenced beyond this function
                    handlers.append(onExecute)

                except:
                    if ui:
                        ui.messageBox('Panel command created failed:\n{}'.format(traceback.format_exc()))

        class CommandCreatedEventHandlerQAT(adsk.core.CommandCreatedEventHandler):
            def __init__(self):
                super().__init__()
            def notify(self, args):
                try:
                    command = args.command
                    onExecute = CommandExecuteHandler()
                    command.execute.add(onExecute)
                    # keep the handler referenced beyond this function
                    handlers.append(onExecute)

                except:
                    ui.messageBox('QAT command created failed:\n{}'.format(traceback.format_exc()))

        commandDefinitions_ = ui.commandDefinitions

		# check if we have the command definition
        commandDefinition_ = commandDefinitions_.itemById(commandId)
        if not commandDefinition_:
            commandDefinition_ = commandDefinitions_.addButtonDefinition(commandId, commandName, commandDescription, commandResources)		 

        onCommandCreated = CommandCreatedEventHandlerPanel()
        commandDefinition_.commandCreated.add(onCommandCreated)
        # keep the handler referenced beyond this function
        handlers.append(onCommandCreated)
        
        # add a command on create panel in modeling workspace
        workspaces_ = ui.workspaces
        modelingWorkspace_ = workspaces_.itemById(workspaceToUse)
        toolbarPanels_ = modelingWorkspace_.toolbarPanels
        toolbarPanel_ = toolbarPanels_.itemById(panelToUse) 
        toolbarControlsPanel_ = toolbarPanel_.controls
        toolbarControlPanel_ = toolbarControlsPanel_.itemById(commandId)
        if not toolbarControlPanel_:
            toolbarControlPanel_ = toolbarControlsPanel_.addCommand(commandDefinition_, '')
            toolbarControlPanel_.isVisible = True
            #ui.messageBox('A CSV command is successfully added to the create panel in modeling workspace')

    except:
        if ui:
            ui.messageBox('AddIn Start Failed:\n{}'.format(traceback.format_exc()))

def stop(context):
    ui = None
    try:
        app = adsk.core.Application.get()
        ui = app.userInterface
        objArray = []

        commandControlPanel_ = commandControlByIdForPanel(commandId)
        if commandControlPanel_:
            objArray.append(commandControlPanel_)
            
        commandDefinition_ = commandDefinitionById(commandId)
        if commandDefinition_:
            objArray.append(commandDefinition_)

        for obj in objArray:
            destroyObject(ui, obj)

    except:
        if ui:
            ui.messageBox('AddIn Stop Failed:\n{}'.format(traceback.format_exc()))

def updateParamsFromCSV():
     app = adsk.core.Application.get()
     ui  = app.userInterface
     
     try:
         #Ask if reading or writing parameters
         dialogResult = ui.messageBox('Importing/Updating parameters from file or Exporting them to file?\n' \
         'Import = Yes, Export = No', 'Import or Export Parameters', \
         adsk.core.MessageBoxButtonTypes.YesNoCancelButtonType, \
         adsk.core.MessageBoxIconTypes.QuestionIconType) 
         if dialogResult == adsk.core.DialogResults.DialogYes:
             readParameters = True
         elif dialogResult == adsk.core.DialogResults.DialogNo:
             readParameters = False
         else:
             return
            
         fileDialog = ui.createFileDialog()
         fileDialog.isMultiSelectEnabled = False
         fileDialog.title = "Get the file to read from or the file to save the parameters to"
         fileDialog.filter = 'Text files (*.csv)'
         fileDialog.filterIndex = 0
         if readParameters:
             dialogResult = fileDialog.showOpen()
         else:
             dialogResult = fileDialog.showSave()
             
         if dialogResult == adsk.core.DialogResults.DialogOK:
             filename = fileDialog.filename
         else:
             return

         #if readParameters is true read the parameters from a file
         if readParameters:
             readTheParameters(filename)
         else:
             writeTheParameters(filename)

     except:
         if ui:
             ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

            
def writeTheParameters(theFileName):
    app = adsk.core.Application.get()
    design = app.activeProduct
      
    result = ""
    for _param in design.allParameters:
        try:
            paramUnit = _param.unit
        except:
            paramUnit = ""
            
        result = result + "\"" + _param.name +  "\",\"" + paramUnit +  "\",\"" + _param.expression.replace('"', '""') + "\",\"" + _param.comment + "\"\n"    
                      
    with open(theFileName, 'w') as outputFile:
        outputFile.writelines(result)
    
    #get the name of the file without the path    
    pathsInTheFileName = theFileName.split("/")
    ui  = app.userInterface
    ui.messageBox('Parameters written to ' + pathsInTheFileName[-1])   
   
def readTheParameters(theFileName):
    app = adsk.core.Application.get()
    design = app.activeProduct
    ui  = app.userInterface
    try:
        paramsList = []
        for oParam in design.allParameters:
            paramsList.append(oParam.name)           
        
        # Read the csv file.
        with open(theFileName) as csvFile:
            csvReader = csv.reader(csvFile, dialect=csv.excel)
            for row in csvReader:
                # Get the values from the csv file.
                nameOfParam = row[0]
                unitOfParam = row[1]
                expressionOfParam = row[2]
                # userParameters.add does not like empty string as comment
                # so we make it a space
                commentOfParam = ' ' 
                if len(row) > 3:
                    commentOfParam = row[3]
                    
                # comment might be empty
                if commentOfParam == '':
                    commentOfParam = ' ' 
                    
                print(expressionOfParam)
                    
                # if the name of the paremeter is not an existing parameter add it
                if nameOfParam not in paramsList:
                    valInput_Param = adsk.core.ValueInput.createByString(expressionOfParam) 
                    design.userParameters.add(nameOfParam, valInput_Param, unitOfParam, commentOfParam)
                # update the values of existing parameters            
                else:
                    paramInModel = design.allParameters.itemByName(nameOfParam)
                    paramInModel.unit = unitOfParam
                    paramInModel.expression = expressionOfParam
                    paramInModel.comment = commentOfParam
                
        ui.messageBox('Finished reading and updating parameters')
    except:
        if ui:
            ui.messageBox('AddIn Stop Failed:\n{}'.format(traceback.format_exc()))