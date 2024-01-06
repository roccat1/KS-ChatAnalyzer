import tkinter as tk
import webbrowser, json

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

def readConfig(configPath):
    with open(configPath, "r") as f:
        config = json.load(f)
    
    return config
    
def saveConfig(configPath, config):
    with open(configPath, "w") as f:
        json.dump(config, f, indent=2)

def configuration(configPath, config):
    #open new window
    configWindow = tk.Toplevel()
    configWindow.title('Configuration')
    configWindow.geometry("700x400")
    configWindow.config(background = "turquoise2")
    
    #option project name, hour divisions, smoothing factor, ouput dir path, output excel file name
    projectName = tk.StringVar(value=config["projectName"])
    hourDivisions = tk.IntVar(value=config["hourDivisions"])
    smoothingFactorCounts = tk.IntVar(value=config["smoothingFactorCounts"])
    smoothingFactorHours = tk.IntVar(value=config["smoothingFactorHours"])
    outputDirPath = tk.StringVar(value=config["outputDirPath"])
    outputExcelFileName = tk.StringVar(value=config["outputExcelFileName"])
    
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
    def saveConfigButton():
        config["projectName"]=projectName.get()
        config["hourDivisions"]=hourDivisions.get()
        config["smoothingFactorCounts"]=smoothingFactorCounts.get()
        config["smoothingFactorHours"]=smoothingFactorHours.get()
        config["outputDirPath"]=outputDirPath.get()
        config["outputExcelFileName"]=outputExcelFileName.get()
        
        saveConfig(configPath, config)
        
        configWindow.destroy()
        
        return config
        
    saveButton = tk.Button(configWindow, text="Save", command=saveConfigButton, width=20, height = 2)
    saveButton.grid(column = 1, row = 7, columnspan=2)
    
    #cancel button
    def cancelConfig():
        configWindow.destroy()
        
        return config
    
    cancelButton = tk.Button(configWindow, text="Cancel", command=cancelConfig, width=20, height = 2)
    cancelButton.grid(column = 1, row = 8, columnspan=2)
    
    #retore default values button
    def restoreDefaultValues():
        projectName.set(defaultConfig["projectName"])
        hourDivisions.set(defaultConfig["hourDivisions"])
        smoothingFactorCounts.set(defaultConfig["smoothingFactorCounts"])
        smoothingFactorHours.set(defaultConfig["smoothingFactorHours"])
        outputDirPath.set(defaultConfig["outputDirPath"])
        outputExcelFileName.set(defaultConfig["outputExcelFileName"])
        
        saveConfig(configPath, defaultConfig)
        
    restoreDefaultValuesButton = tk.Button(configWindow, text="Restore Default Values", command=restoreDefaultValues, width=20, height = 2)
    restoreDefaultValuesButton.grid(column = 1, row = 9, columnspan=2)
    
    #wiki button
    def wiki():
        webbrowser.open('https://github.com/roccat1/KS-ChatAnalyzer/wiki')
    
    wikiButton = tk.Button(configWindow, text="Wiki", command=wiki, width=20, height = 2)
    wikiButton.grid(column = 1, row = 10, columnspan=2)
    
    #let the window wait for any events
    configWindow.mainloop()