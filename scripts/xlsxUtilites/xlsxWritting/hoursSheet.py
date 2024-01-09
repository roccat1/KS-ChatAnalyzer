import datetime
from xlsxwriter.utility import xl_col_to_name

import scripts.gvar as gvar
import scripts.xlsxUtilites.xlsxGvar as xlsxGvar
import scripts.configuration.configuration as configuration
from scripts.xlsxUtilites.createChart import createChart
import scripts.xlsxUtilites.xlsxWritting.cellFormat as cellFormat

def createHoursSheet(jsonFile: dict) -> None:
    ############################################ HOURS SHEET #########################################################
    hoursSheet = xlsxGvar.workbook.add_worksheet("hours")
    hoursSheet.write(0, 0, 'Hours')

    #upper row
    row = 2
    hour=datetime.datetime(2000,1,1,hour=0, minute=0)
    for i in range(0,configuration.config["hourDivisions"]):
        nextHour=hour+datetime.timedelta(hours=24/configuration.config["hourDivisions"])
        hoursSheet.write(1, 3+i, hour.strftime("%H:%M")+"-"+nextHour.strftime("%H:%M"), cellFormat.cellFormatUp)
        hour=nextHour
    
    totalHourResults=[0]*configuration.config["hourDivisions"]
    
    #data per month
    for year in jsonFile["data"]:
        hoursSheet.write(row, 1, year, cellFormat.cellFormatDown)
        #months
        for month in jsonFile["data"][year]:
            hoursSheet.write(row, 2, month, cellFormat.cellFormatDown)

            #create list of 0s
            hourResults=[0]*configuration.config["hourDivisions"]

            #days
            for day in jsonFile["data"][year][month]:
                #each hour
                for hour in jsonFile["data"][year][month][day][1]:
                    hourResults[int((int(hour.split(":")[0])*60+int(hour.split(":")[1]))/(1440/configuration.config["hourDivisions"]))]+=1
            
            col=3
            for i in hourResults:
                hoursSheet.write(row, col, i, cellFormat.cellFormatDown)
                totalHourResults[col-3]+=i
                col+=1
            
            row+=1
            
    #total data
    #upper row
    row+=1
    hour=datetime.datetime(2000,1,1,hour=0, minute=0)
    for i in range(0,configuration.config["hourDivisions"]):
        nextHour=hour+datetime.timedelta(hours=24/configuration.config["hourDivisions"])
        hoursSheet.write(row, 3+i, hour.strftime("%H:%M")+"-"+nextHour.strftime("%H:%M"), cellFormat.cellFormatUp)
        hour=nextHour
    #data
    col=3
    row+=1
    hoursSheet.write(row, 2, "Total", cellFormat.cellFormatDown)
    for i in totalHourResults:
        hoursSheet.write(row, col, i, cellFormat.cellFormatDown)
        col+=1
    
    #smoothed data
    #upper row
    row+=2
    hour=datetime.datetime(2000,1,1,hour=0, minute=0)
    for i in range(0,configuration.config["hourDivisions"]):
        nextHour=hour+datetime.timedelta(hours=24/configuration.config["hourDivisions"])
        hoursSheet.write(row, 3+i, hour.strftime("%H:%M")+"-"+nextHour.strftime("%H:%M"), cellFormat.cellFormatUp)
        hour=nextHour
    #data
    col=3
    row+=1
    hoursSheet.write(row, 2, "Smoothed", cellFormat.cellFormatDown)
    for i in range(0,configuration.config["hourDivisions"]):
        if i>=configuration.config["smoothingFactorHours"] and i<configuration.config["hourDivisions"]-configuration.config["smoothingFactorHours"]:
            hoursSheet.write(row, col, round(sum(totalHourResults[i-configuration.config["smoothingFactorHours"]:i+configuration.config["smoothingFactorHours"]+1])/(configuration.config["smoothingFactorHours"]*2+1), 2), cellFormat.cellFormatDown)
        else:
            hoursSheet.write(row, col, 0, cellFormat.cellFormatDown)
        col+=1
    
    hoursSheet.autofit()

    createChart("Hours Distribution", "hours", ["D", 2, xl_col_to_name(2+configuration.config["hourDivisions"]), 2], ["D", row-2, xl_col_to_name(2+configuration.config["hourDivisions"]), row-2], "Hours Chart", "Hours Count", "Hours")

    createChart("Hours Distribution Smoothed", "hours", ["D", 2, xl_col_to_name(2+configuration.config["hourDivisions"]), 2], ["D", row+1, xl_col_to_name(2+configuration.config["hourDivisions"]), row+1], "Hours Chart Smoothed", "Hours Count Smoothed", "Hours")
    