import scripts.configuration.configuration as configuration
from scripts.utilities.log import log
import scripts.gvar as gvar

def setDDMM() -> None:
    """Sets the dd_mmFormat to True in the config file"""

    configuration.config["dd_mmFormat"]=True
    configuration.saveConfig(gvar.configPath, configuration.config)
    log("dd_mmFormat set to True")

def setMMDD() -> None:
    """Sets the dd_mmFormat to False in the config file"""

    configuration.config["dd_mmFormat"]=False
    configuration.saveConfig(gvar.configPath, configuration.config)
    log("dd_mmFormat set to False")