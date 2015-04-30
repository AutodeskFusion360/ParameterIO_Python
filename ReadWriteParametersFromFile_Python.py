#Author-Wayne Brill
#Description-Allows you to select a CSV (comma seperated values) file and then edits existing Attributes. Also allows you to write parameters to a file

import adsk.core, adsk.fusion, traceback

commandIdOnQAT = 'ParamsFromCSVOnQAT'
commandIdOnPanel = 'ParamsFromCSVOnPanel'

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

def commandControlByIdForQAT(id):
    app = adsk.core.Application.get()
    ui = app.userInterface
    if not id:
        ui.messageBox('commandControl id is not specified')
        return None
    toolbars_ = ui.toolbars
    toolbarQAT_ = toolbars_.itemById('QAT')
    toolbarControls_ = toolbarQAT_.controls
    toolbarControl_ = toolbarControls_.itemById(id)
    return toolbarControl_

def commandControlByIdForPanel(id):
    app = adsk.core.Application.get()
    ui = app.userInterface
    if not id:
        ui.messageBox('commandControl id is not specified')
        return None
    workspaces_ = ui.workspaces
    modelingWorkspace_ = workspaces_.itemById('FusionSolidEnvironment')
    toolbarPanels_ = modelingWorkspace_.toolbarPanels
    toolbarPanel_ = toolbarPanels_.item(0)
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
        commandName = 'CSV Parameters'
        commandDescription = 'Parameters from CSV File'
        commandResources = './resources'

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

        # add a command button on Quick Access Toolbar
        toolbars_ = ui.toolbars
        toolbarQAT_ = toolbars_.itemById('QAT')
        toolbarControlsQAT_ = toolbarQAT_.controls
        toolbarControlQAT_ = toolbarControlsQAT_.itemById(commandIdOnQAT)
        if not toolbarControlQAT_:
            commandDefinitionQAT_ = commandDefinitions_.itemById(commandIdOnQAT)
            if not commandDefinitionQAT_:
                commandDefinitionQAT_ = commandDefinitions_.addButtonDefinition(commandIdOnQAT, commandName, commandDescription, commandResources)
            onCommandCreated = CommandCreatedEventHandlerQAT()
            commandDefinitionQAT_.commandCreated.add(onCommandCreated)
            # keep the handler referenced beyond this function
            handlers.append(onCommandCreated)
            toolbarControlQAT_ = toolbarControlsQAT_.addCommand(commandDefinitionQAT_, commandIdOnQAT)
            toolbarControlQAT_.isVisible = True
            ui.messageBox('A CSV command button is successfully added to the Quick Access Toolbar')

        # add a command on create panel in modeling workspace
        workspaces_ = ui.workspaces
        modelingWorkspace_ = workspaces_.itemById('FusionSolidEnvironment')
        toolbarPanels_ = modelingWorkspace_.toolbarPanels
        toolbarPanel_ = toolbarPanels_.item(0) # add the new command under the first panel
        toolbarControlsPanel_ = toolbarPanel_.controls
        toolbarControlPanel_ = toolbarControlsPanel_.itemById(commandIdOnPanel)
        if not toolbarControlPanel_:
            commandDefinitionPanel_ = commandDefinitions_.itemById(commandIdOnPanel)
            if not commandDefinitionPanel_:
                commandDefinitionPanel_ = commandDefinitions_.addButtonDefinition(commandIdOnPanel, commandName, commandDescription, commandResources)
            onCommandCreated = CommandCreatedEventHandlerPanel()
            commandDefinitionPanel_.commandCreated.add(onCommandCreated)
            # keep the handler referenced beyond this function
            handlers.append(onCommandCreated)
            toolbarControlPanel_ = toolbarControlsPanel_.addCommand(commandDefinitionPanel_, commandIdOnPanel)
            toolbarControlPanel_.isVisible = True
            ui.messageBox('A CSV command is successfully added to the create panel in modeling workspace')

    except:
        if ui:
            ui.messageBox('AddIn Start Failed:\n{}'.format(traceback.format_exc()))


def stop(context):
    ui = None
    try:
        app = adsk.core.Application.get()
        ui = app.userInterface
        objArrayQAT = []
        objArrayPanel = []

        commandControlQAT_ = commandControlByIdForQAT(commandIdOnQAT)
        if commandControlQAT_:
            objArrayQAT.append(commandControlQAT_)

        commandDefinitionQAT_ = commandDefinitionById(commandIdOnQAT)
        if commandDefinitionQAT_:
            objArrayQAT.append(commandDefinitionQAT_)

        commandControlPanel_ = commandControlByIdForPanel(commandIdOnPanel)
        if commandControlPanel_:
            objArrayPanel.append(commandControlPanel_)

        commandDefinitionPanel_ = commandDefinitionById(commandIdOnPanel)
        if commandDefinitionPanel_:
            objArrayPanel.append(commandDefinitionPanel_)

        for obj in objArrayQAT:
            destroyObject(ui, obj)

        for obj in objArrayPanel:
            destroyObject(ui, obj)

    except:
        if ui:
            ui.messageBox('AddIn Stop Failed:\n{}'.format(traceback.format_exc()))

def updateParamsFromCSV():
    ui = None
    
    try:
        app = adsk.core.Application.get()
        design = app.activeProduct
        ui  = app.userInterface
        
        fileDialog = ui.createFileDialog()
        fileDialog.isMultiSelectEnabled = False
        fileDialog.title = "Select CSV file the with parameters"
        fileDialog.filter = 'Text files (*.csv)'
        fileDialog.filterIndex = 0
        dialogResult = fileDialog.showOpen()
        if dialogResult == adsk.core.DialogResults.DialogOK:
            filename = fileDialog.filename
        else:
            return
        
        
        # get the names of current parameters and put them in a list
        paramsList = []
        for oParam in design.allParameters:
           paramsList.append(oParam.name)           
        
        # Read the csv file.
        csvFile = open(filename)
        for line in csvFile:
            # Get the values from the csv file.
            valsInTheLine = line.split(',')
            nameOfParam = valsInTheLine[0]
            valueOfParam = valsInTheLine[1]
            # if the name of the paremeter is not an existing parameter
            if nameOfParam not in paramsList:
                valInput_Param = adsk.core.ValueInput.createByString(valueOfParam)                
                design.userParameters.add(nameOfParam, valInput_Param, 'mm', 'Comment')
                #create a new parameter            
            else:
                paramInModel = design.userParameters.itemByName(nameOfParam)
                paramInModel.expression = valueOfParam
     
        ui.messageBox('Finished adding parameters')

    except Exception as error:
        if ui:
           ui.messageBox('Failed : ' + str(error))            
           #ui.messageBox('Failed:\n{}'.format(traceback.forma​t_exc()))           