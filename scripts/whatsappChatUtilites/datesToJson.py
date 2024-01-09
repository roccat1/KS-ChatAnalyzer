import datetime

import scripts.gvar as gvar
import scripts.configuration.configuration as configuration

def datesToJson(dates: list) -> dict:
    """Converts a list of dates to a dictionary with the format of the json result

    Args:
        dates (list): List of dates

    Returns:
        dict: Dictionary with the format of the json result
    """
    #create canvas for json result
    #primera data
    date = datetime.datetime(dates[0].year, dates[0].month, dates[0].day)
    prevDate = datetime.datetime(1950,1,1)
    result={
        "metadata": {
            "name": configuration.config["projectName"],
            "creation date": datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        },
        "data": {}
    }

    while date<=dates[-1]:
        change=0
        #canvi any
        if date.year!=prevDate.year:
            change=1
            result["data"][date.year]={}
        #canvi mes
        if date.month!=prevDate.month or change>=1:
            result["data"][date.year][gvar.months[date.month]]={}

        result["data"][date.year][gvar.months[date.month]][date.day]=[0,[]]

        prevDate=date
        date+=datetime.timedelta(days=1)
    
    #restructure dates ['15/02/2023', 2, ['08:45', '18:17']]
    datesResturcture=[[dates[0], 1, [dates[0].strftime("%H:%M")]]]
    prevDate=dates[0]
    for date in dates:
        if date.strftime("%d/%m/%Y") == prevDate.strftime("%d/%m/%Y"):
            datesResturcture[-1][1]+=1
            datesResturcture[-1][2].append(date.strftime("%H:%M"))
        else:
            datesResturcture.append([date, 1, [date.strftime("%H:%M")]])

        prevDate=date

    #arreglar primera data
    datesResturcture[0][1]-=1
    datesResturcture[0][2].pop(0)
    
    #merge restructured data to json result
    for dateData in datesResturcture:
        result["data"][dateData[0].year][gvar.months[dateData[0].month]][dateData[0].day]=[dateData[1], dateData[2]]

    return result