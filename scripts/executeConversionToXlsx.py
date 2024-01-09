import os, subprocess
from tkinter import messagebox

import scripts.gvar as gvar
from scripts.utilities.log import log
import scripts.whatsappChatUtilites.readWhatsappChatDates as readWhatsappChatDates
import scripts.whatsappChatUtilites.datesToJson as datesToJson
import scripts.utilities.manageJson as manageJson
import scripts.xlsxUtilites.jsonToXlsx as writeJsonToXls


def runProgram() -> None:
    """Runs the program with the selected file on the global variable filename"""
    log("program executed... ")
    try:
        datesRaw = readWhatsappChatDates.readWAChatDates(gvar.filename)
        jsonResult = datesToJson.datesToJson(datesRaw)
        if not os.path.exists(gvar.config["outputDirPath"]): os.makedirs(gvar.config["outputDirPath"])
        manageJson.saveJson(jsonResult)
        writeJsonToXls.writeJsonToXls(jsonResult)

        gvar.label_file_explorer.configure(text="Program executed")
        messagebox.showinfo("Done!", str(os.path.basename(gvar.filename))+" has been analyzed successfully!")
        subprocess.run([os.path.join(os.getenv('WINDIR'), 'explorer.exe'), os.getcwd()+"\\output"])
        log("successfully")
    except Exception as e:
        log("with an error :(")
        if str(e)=="[Errno 13] Permission denied: 'output/output.xlsx'":
            gvar.label_file_explorer.configure(text="Close output.xlsx before running the program again")
            messagebox.showerror("Error!", "Close output.xlsx before running the program again")
            log("ERROR: Close output.xlsx before running the program again")
        elif str(e)==f"[Errno 2] No such file or directory: '{gvar.config['defaultFilePath']}'":
            gvar.label_file_explorer.configure(text="ERROR: whatsapp chat file not found")
            messagebox.showerror("Error!", "ERROR: whatsapp chat file not found")
            log("ERROR: whatsapp chat file not found")
        else:
            gvar.label_file_explorer.configure(text="ERROR(is the format correct?/check log/terminal)")
        log("ERROR: "+str(e))