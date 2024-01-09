import datetime

import scripts.gvar as gvar
import scripts.utilities.log as log

def readWAChatDates(fileName: str) -> list:
    """Reads a whatsapp chat and returns a list of dates

    Args:
        fileName (str): Path to the file

    Returns:
        list: List of dates of the messages in the chat
    """
    #llegir doc
    f = open(fileName, "r", encoding='UTF-8')
    chatRaw = f.readlines()
    f.close()

    result=[]
    for line in chatRaw:
        try:
            #comprovar si es missatge nou
            completeDate=line.split(" - ")[0]
            dateAndTime=completeDate.split(", ")
            date=dateAndTime[0].split("/")
            time=dateAndTime[1].split(":")

            if gvar.config["dd_mmFormat"]:
                fullDate = datetime.datetime(int("20"+ date[2]), int(date[1]), int(date[0]), int(time[0]), int(time[1]))
            else:
                fullDate = datetime.datetime(int("20"+ date[2]), int(date[0]), int(date[1]), int(time[0]), int(time[1]))
            
            result.append(fullDate)
        except:
            log("error reading line: "+line)
    return result