import tkinter as tk
import webbrowser, json

from scripts.utilities.log import log
import scripts.gvar as gvar

#create default config
defaultConfig = {
            "projectName": "KS-ChatAnalyzer",
            "hourDivisions": 96,
            "smoothingFactorCounts": 3,
            "smoothingFactorHours": 2,

            "dd_mmFormat": True,
            "defaultFilePath": "sensible/ks-chat.txt",
            "logPath": "log.txt",
            "outputDirPath": "output/",
            "outputJsonFileName": "json_output.json",
            "outputExcelFileName": "output.xlsx"
}

def readConfig(configPath: str) -> dict:
    """reads the config file and returns it as a dict

    Args:
        configPath (str): path to the config file

    Returns:
        dict: config file as a dict
    """
    with open(configPath, "r") as f:
        config = json.load(f)
    
    return config
    
def saveConfig(configPath: str, config: dict) -> None:
    """saves the config file

    Args:
        configPath (str): path to the config file
        config (dict): config file as a dict
    """
    with open(configPath, "w") as f:
        json.dump(config, f, indent=2)

def configurationMenu() -> None:
    """opens the configuration menu"""
    log("configuration menu opened")
    #open new window
    configWindow = tk.Toplevel()
    configWindow.title('Configuration')
    configWindow.geometry("700x400")
    configWindow.config(background = "turquoise2")
    
    #option project name, hour divisions, smoothing factor, ouput dir path, output excel file name
    projectName = tk.StringVar(value=gvar.config["projectName"])
    hourDivisions = tk.IntVar(value=gvar.config["hourDivisions"])
    smoothingFactorCounts = tk.IntVar(value=gvar.config["smoothingFactorCounts"])
    smoothingFactorHours = tk.IntVar(value=gvar.config["smoothingFactorHours"])
    outputDirPath = tk.StringVar(value=gvar.config["outputDirPath"])
    outputExcelFileName = tk.StringVar(value=gvar.config["outputExcelFileName"])
    
    #project name
    projectNameLabel = tk.Label(configWindow, text="Project Name", width=20, height = 2, fg="white", bg="black")
    projectNameEntry = tk.Entry(configWindow, textvariable=projectName, width=50, font=("Arial", 15), border=5)
    projectNameLabel.grid(column = 1, row = 1, sticky="w")
    projectNameEntry.grid(column = 2, row = 1, sticky="w")
    
    #hour divisions
    hourDivisionsLabel = tk.Label(configWindow, text="Hour Divisions", width=20, height = 2, fg="white", bg="black")
    hourDivisionsEntry = tk.Entry(configWindow, textvariable=hourDivisions, width=50, font=("Arial", 15), border=5)
    hourDivisionsLabel.grid(column = 1, row = 2, sticky="w")
    hourDivisionsEntry.grid(column = 2, row = 2, sticky="w")
    
    #smoothing factor counts
    smoothingFactorLabel = tk.Label(configWindow, text="Smoothing Factor Counts", width=20, height = 2, fg="white", bg="black")
    smoothingFactorEntry = tk.Entry(configWindow, textvariable=smoothingFactorCounts, width=50, font=("Arial", 15), border=5)
    smoothingFactorLabel.grid(column = 1, row = 3, sticky="w")
    smoothingFactorEntry.grid(column = 2, row = 3, sticky="w")
    
    #smoothing factor hours
    smoothingFactorLabel = tk.Label(configWindow, text="Smoothing Factor Hours", width=20, height = 2, fg="white", bg="black")
    smoothingFactorEntry = tk.Entry(configWindow, textvariable=smoothingFactorHours, width=50, font=("Arial", 15), border=5)
    smoothingFactorLabel.grid(column = 1, row = 4, sticky="w")
    smoothingFactorEntry.grid(column = 2, row = 4, sticky="w")
    
    #output dir path
    outputDirPathLabel = tk.Label(configWindow, text="Output Dir Path", width=20, height = 2, fg="white", bg="black")
    outputDirPathEntry = tk.Entry(configWindow, textvariable=outputDirPath, width=50, font=("Arial", 15), border=5)
    outputDirPathLabel.grid(column = 1, row = 5, sticky="w")
    outputDirPathEntry.grid(column = 2, row = 5, sticky="w")
    
    #output excel file name
    outputExcelFileNameLabel = tk.Label(configWindow, text="Output Excel File Name", width=20, height = 2, fg="white", bg="black")
    outputExcelFileNameEntry = tk.Entry(configWindow, textvariable=outputExcelFileName, width=50, font=("Arial", 15), border=5)
    outputExcelFileNameLabel.grid(column = 1, row = 6, sticky="w")
    outputExcelFileNameEntry.grid(column = 2, row = 6, sticky="w")
    
    #save button
    def saveConfigButton() -> None:
        """saves the config file and closes the window"""
        gvar.config["projectName"]=projectName.get()
        gvar.config["hourDivisions"]=hourDivisions.get()
        gvar.config["smoothingFactorCounts"]=smoothingFactorCounts.get()
        gvar.config["smoothingFactorHours"]=smoothingFactorHours.get()
        gvar.config["outputDirPath"]=outputDirPath.get()
        gvar.config["outputExcelFileName"]=outputExcelFileName.get()
        
        saveConfig(gvar.configPath, gvar.config)
        
        configWindow.destroy()
        
    saveButton = tk.Button(configWindow, text="Save", command=saveConfigButton, width=20, height = 2)
    saveButton.grid(column = 1, row = 7, columnspan=2)
    
    #cancel button
    def cancelConfig() -> None:
        """closes the window without saving the config file"""
        gvar.config = readConfig(gvar.configPath)
        configWindow.destroy()
    
    cancelButton = tk.Button(configWindow, text="Cancel", command=cancelConfig, width=20, height = 2)
    cancelButton.grid(column = 1, row = 8, columnspan=2)
    
    #retore default values button
    def restoreDefaultValues() -> None:
        """restores the default values"""
        projectName.set(defaultConfig["projectName"])
        hourDivisions.set(defaultConfig["hourDivisions"])
        smoothingFactorCounts.set(defaultConfig["smoothingFactorCounts"])
        smoothingFactorHours.set(defaultConfig["smoothingFactorHours"])
        outputDirPath.set(defaultConfig["outputDirPath"])
        outputExcelFileName.set(defaultConfig["outputExcelFileName"])
        
    restoreDefaultValuesButton = tk.Button(configWindow, text="Restore Default Values", command=restoreDefaultValues, width=20, height = 2)
    restoreDefaultValuesButton.grid(column = 1, row = 9, columnspan=2)
    
    #wiki button
    def wiki() -> None:
        """opens the wiki in the browser"""
        webbrowser.open('https://github.com/roccat1/KS-ChatAnalyzer/wiki')
    
    wikiButton = tk.Button(configWindow, text="Wiki", command=wiki, width=20, height = 2)
    wikiButton.grid(column = 1, row = 10, columnspan=2)
    
    #let the window wait for any events
    configWindow.mainloop()