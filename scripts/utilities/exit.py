import os

import scripts.gvar as gvar
from scripts.utilities.log import log



def exit() -> None:
    """Closes the program"""
    log("program closed")
    gvar.window.destroy()
    print(f"Thank you for using {gvar.appName}!")
    print(f"Log file: {gvar.logPath}")
    os._exit(0)