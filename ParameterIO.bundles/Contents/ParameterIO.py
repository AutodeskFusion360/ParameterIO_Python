#Author-Wayne Brill
#Description-Allows you to select a CSV (comma seperated values) file and then edits existing Attributes. Also allows you to write parameters to a file

import adsk.core, adsk.fusion, traceback, csv

_commandId = 'ParamsFromCSV'
_workspaceToUse = 'FusionSolidEnvironment'
_panelToUse = 'SolidModifyPanel'

# global set of event handlers to keep them referenced for the duration of the command
_handlers = []

def commandDefinitionById(id):
    app = adsk.core.Application.get()
    ui = app.userInterface
    if not id:
        ui.messageBox('commandDefinition id is not specified')
        return None
    commandDefinitions = ui.commandDefinitions
    commandDefinition = commandDefinitions.itemById(id)
    return commandDefinition

def commandControlByIdForPanel(id):
    app = adsk.core.Application.get()
    ui = app.userInterface
    if not id:
        ui.messageBox('commandControl id is not specified')
        return None
    workspaces = ui.workspaces
    modelingWorkspace = workspaces.itemById(_workspaceToUse)
    toolbarPanels = modelingWorkspace.toolbarPanels
    toolbarPanel = toolbarPanels.itemById(_panelToUse)
    toolbarControls = toolbarPanel.controls
    toolbarControl = toolbarControls.itemById(id)
    return toolbarControl

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
                    cmd = args.command
                    inputs = cmd.commandInputs
                    radioButtonGroup = inputs.tableInput = inputs.itemById('radioImportExport') 
                    doImportExport(radioButtonGroup.selectedItem.name == 'Import')
                except:
                    if ui:
                        ui.messageBox('command executed failed:\n{}'.format(traceback.format_exc()))

        class CommandCreatedHandler(adsk.core.CommandCreatedEventHandler):
            def __init__(self):
                super().__init__() 
            def notify(self, args):
                try:
                    cmd = args.command
                    onExecute = CommandExecuteHandler()
                    cmd.execute.add(onExecute)
                    # keep the handler referenced beyond this function
                    _handlers.append(onExecute)

                    inputs = cmd.commandInputs

                    radioButtonGroup = inputs.addRadioButtonGroupCommandInput('radioImportExport', ' Import or Export ')
                    radioButtonGroup.isFullWidth = True
                    radioButtonItems = radioButtonGroup.listItems
                    radioButtonItems.add('Import', True)
                    radioButtonItems.add('Export', False)

                except:
                    if ui:
                        ui.messageBox('Panel command created failed:\n{}'.format(traceback.format_exc()))

        commandDefinitions = ui.commandDefinitions

		# check if we have the command definition
        commandDefinition = commandDefinitions.itemById(_commandId)
        if not commandDefinition:
            commandDefinition = commandDefinitions.addButtonDefinition(_commandId, commandName, commandDescription, commandResources)		 

        onCommandCreated = CommandCreatedHandler()
        commandDefinition.commandCreated.add(onCommandCreated)
        # keep the handler referenced beyond this function
        _handlers.append(onCommandCreated)
        
        # add a command on create panel in modeling workspace
        workspaces = ui.workspaces
        modelingWorkspace = workspaces.itemById(_workspaceToUse)
        toolbarPanels = modelingWorkspace.toolbarPanels
        toolbarPanel = toolbarPanels.itemById(_panelToUse) 
        toolbarControlsPanel = toolbarPanel.controls
        toolbarControlPanel = toolbarControlsPanel.itemById(_commandId)
        if not toolbarControlPanel:
            toolbarControlPanel = toolbarControlsPanel.addCommand(commandDefinition, '')
            toolbarControlPanel.isVisible = True
            print('The Parameter I/O command was successfully added to the create panel in modeling workspace')

    except:
        if ui:
            ui.messageBox('AddIn Start Failed:\n{}'.format(traceback.format_exc()))

def stop(context):
    ui = None
    try:
        app = adsk.core.Application.get()
        ui = app.userInterface
        objArray = []

        commandControlPanel = commandControlByIdForPanel(_commandId)
        if commandControlPanel:
            objArray.append(commandControlPanel)
            
        commandDefinition = commandDefinitionById(_commandId)
        if commandDefinition:
            objArray.append(commandDefinition)

        for obj in objArray:
            destroyObject(ui, obj)

    except:
        if ui:
            ui.messageBox('AddIn Stop Failed:\n{}'.format(traceback.format_exc()))

def doImportExport(isImport):
     app = adsk.core.Application.get()
     ui  = app.userInterface
     
     try:   
         fileDialog = ui.createFileDialog()
         fileDialog.isMultiSelectEnabled = False
         fileDialog.title = "Get the file to read from or the file to save the parameters to"
         fileDialog.filter = 'Text files (*.csv)'
         fileDialog.filterIndex = 0
         if isImport:
             dialogResult = fileDialog.showOpen()
         else:
             dialogResult = fileDialog.showSave()
             
         if dialogResult == adsk.core.DialogResults.DialogOK:
             filename = fileDialog.filename
         else:
             return

         # if isImport is true read the parameters from a file
         if isImport:
             readParametersFromFile(filename)
         else:
             writeParametersToFile(filename)

     except:
         if ui:
             ui.messageBox('Failed:\n{}'.format(traceback.format_exc())) 

def writeParametersToFile(filePath):
    app = adsk.core.Application.get()
    design = app.activeProduct
                      
    with open(filePath, 'w', newline='') as csvFile:
        csvWriter = csv.writer(csvFile, dialect=csv.excel)
        for param in design.allParameters:
            try:
                paramUnit = param.unit
            except:
                paramUnit = ""
            
            csvWriter.writerow([param.name, paramUnit, param.expression, param.comment]) 
    
    # get the name of the file without the path    
    partsOfFilePath = filePath.split("/")
    ui  = app.userInterface
    ui.messageBox('Parameters written to ' + partsOfFilePath[-1])       
   
def updateParameter(design, paramsList, row):
    # get the values from the csv file.
    try:
        nameOfParam = row[0]
        unitOfParam = row[1]
        expressionOfParam = row[2]
        try:
            commentOfParam = row[3]
        except:
            commentOfParam = ''
    except Exception as e:
        print(str(e))
        # no plint to retry
        return True

    # userParameters.add did not used to like empty string as comment
    # so we make it a space
    # comment might be missing
    #if commentOfParam == '':
    #    commentOfParam = ' ' 

    try: 
        # if the name of the paremeter is not an existing parameter add it
        if nameOfParam not in paramsList:
            valInputparam = adsk.core.ValueInput.createByString(expressionOfParam) 
            design.userParameters.add(nameOfParam, valInputparam, unitOfParam, commentOfParam)
            print("Added {}".format(nameOfParam))
            
        # update the values of existing parameters            
        else:
            paramInModel = design.allParameters.itemByName(nameOfParam)
            paramInModel.unit = unitOfParam
            paramInModel.expression = expressionOfParam
            paramInModel.comment = commentOfParam
            print("Updated {}".format(nameOfParam))
        
        return True

    except Exception as e:
        print(str(e))
        print("Failed to update {}".format(nameOfParam))
        return False

def readParametersFromFile(filePath):
    app = adsk.core.Application.get()
    design = app.activeProduct
    ui  = app.userInterface
    try:
        paramsList = []
        for oParam in design.allParameters:
            paramsList.append(oParam.name) 

        retryList = []            
        
        # read the csv file.
        with open(filePath) as csvFile:
            csvReader = csv.reader(csvFile, dialect=csv.excel)
            for row in csvReader:
                # if the parameter is referencing a non-existent
                # parameter then  this will fail
                # so let's store those params and try to add them in the next round 
                if not updateParameter(design, paramsList, row):
                    retryList.append(row)

        # let's keep going through the list until all is done
        count = 0
        while len(retryList) + 1 > count:
            count = count + 1
            for row in retryList:
                if updateParameter(design, paramsList, row):
                    retryList.remove(row)
                
        if len(retryList) > 0:
            params = ""
            for row in retryList:
                params = params + '\n' + row[0]

            ui.messageBox('Could not set the following parameters:' + params)
        else:
            ui.messageBox('Finished reading and updating parameters')
    except:
        if ui:
            ui.messageBox('AddIn Stop Failed:\n{}'.format(traceback.format_exc()))