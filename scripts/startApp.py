import datetime, os, appdirs

import scripts.configuration.configuration as configuration
import scripts.gvar as gvar
from scripts.utilities.log import log
import scripts.UI.mainWindowUI as mainWindowUI

def startApp() -> None:
    """starts the app, creates the log and loads the config"""
    
    #create log path
    gvar.logPath = os.path.join(appdirs.user_log_dir(gvar.appName, gvar.appAuthor), datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S") + ".log")
    log("program started, log at "+str(gvar.logPath))
    
    #create log file
    if not os.path.exists(os.path.dirname(gvar.logPath)):
        os.makedirs(os.path.dirname(gvar.logPath), exist_ok=True)
        log("log.log created")
    
    #create/load config
    gvar.configPath = os.path.join(appdirs.user_config_dir(gvar.appName, gvar.appAuthor), "config", "config.json")
    if not os.path.exists(gvar.configPath):
        os.makedirs(os.path.dirname(gvar.configPath), exist_ok=True)
        configuration.saveConfig(gvar.configPath, configuration.defaultConfig)
        gvar.config = configuration.defaultConfig
        log("config.json created")
    else:
        gvar.config = configuration.readConfig(gvar.configPath)
        log("config.json loaded")
    
    #set default file path
    gvar.filename = gvar.config["defaultFilePath"]

    #run UI
    mainWindowUI.mainWindowUI()