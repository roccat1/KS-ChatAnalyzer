import json

import scripts.configuration.configuration as configuration

def saveJson(list: dict) -> None:
    """Saves a dictionary with json format to a json file

    Args:
        list (dict): Dictionary with json format
    """
    with open(configuration.config["outputDirPath"]+configuration.config["outputJsonFileName"], "w") as fp:
        json.dump(list, fp, indent=2)

def readJson() -> dict:
    """Reads a json file and returns a dictionary with json format

    Returns:
        dict: Dictionary with json format
    """
    with open("output/output.json", 'r') as f:
        list = json.load(f)
    return list