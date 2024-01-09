import xlsxwriter, datetime

import scripts.gvar as gvar
from scripts.xlsxUtilites.createChart import createChart

def writeJsonToXls(jsonFile: dict) -> None:
    """Writes the json file to an excel file processing all the data and creating charts and sheets with statistics

    Args:
        jsonFile (dict): Dictionary with json format to be written to the excel file
    """
    gvar.workbook = xlsxwriter.Workbook(gvar.config["outputDirPath"]+gvar.config["outputExcelFileName"])

    ############################################ COUNTS SHEET #########################################################
    countsSheet = gvar.workbook.add_worksheet("counts")
    countsSheet.write(0, 0, 'Counts')

    rowDays = 2
    rowMonths = 5
    rowYears = 5
    
    countsThisYear=0
    
    #cell format for upper part
    cellFormatUp = gvar.workbook.add_format()
    cellFormatUp.set_border(1)
    cellFormatUp.set_bg_color('#C6C6C6')
    
    #cell format for lower part
    cellFormatDown = gvar.workbook.add_format()
    cellFormatDown.set_border(1)
    cellFormatDown.set_bg_color('#E1E1E1')
    
    countsSheet.write(1, 1, 'Date', cellFormatUp)
    countsSheet.write(1, 2, 'Counts', cellFormatUp)
    
    countsSheet.write(4, 4, 'Month', cellFormatUp)
    countsSheet.write(4, 5, 'Counts', cellFormatUp)
    countsSheet.write(4, 6, 'Average', cellFormatUp)
    
    countsSheet.write(4, 8, 'Year', cellFormatUp)
    countsSheet.write(4, 9, 'Counts', cellFormatUp)
    countsSheet.write(4, 10, 'Average', cellFormatUp)
    
    counts = []
    
    #per day counts
    for year in jsonFile["data"]:
        for month in jsonFile["data"][year]:
            for day in jsonFile["data"][year][month]:
                #write date
                countsSheet.write(rowDays, 1, f"{year}-{gvar.months.index(month):02}-{day:02}", cellFormatDown)
                #write counts day
                countsSheet.write(rowDays, 2, jsonFile["data"][year][month][day][0], cellFormatDown)
                counts.append(jsonFile["data"][year][month][day][0])
                
                rowDays+=1
                
            #write month
            countsSheet.write(rowMonths, 4, f"{month}-{year}", cellFormatDown)
            #write counts month
            countsSheet.write(rowMonths, 5, sum([jsonFile['data'][year][month][day][0] for day in jsonFile['data'][year][month]]), cellFormatDown)
            countsThisYear+=sum([jsonFile['data'][year][month][day][0] for day in jsonFile['data'][year][month]])
            #write average counts month
            countsSheet.write(rowMonths, 6, round(sum([jsonFile['data'][year][month][day][0] for day in jsonFile['data'][year][month]])/len(jsonFile['data'][year][month]), 2), cellFormatDown)
            
            rowMonths+=1
            
        #write year
        countsSheet.write(rowYears, 8, year, cellFormatDown)
        #write counts year
        countsSheet.write(rowYears, 9, countsThisYear, cellFormatDown)
        #write average counts year
        countsSheet.write(rowYears, 10, round(countsThisYear/sum([len(jsonFile['data'][year][month]) for month in jsonFile['data'][year]]), 2), cellFormatDown)
        
        countsThisYear=0
        rowYears+=1
    
    countsSheet.write(4, 12, 'Date', cellFormatUp)
    countsSheet.write(4, 13, 'Smoothed Counts', cellFormatUp)
    
    rowDays = 5
    countsLocation=0
    
    #smoothed counts
    for year in jsonFile["data"]:
        for month in jsonFile["data"][year]:
            for day in jsonFile["data"][year][month]:
                if countsLocation>=gvar.config["smoothingFactorCounts"]:
                    #write date
                    countsSheet.write(rowDays, 12, f"{year}-{gvar.months.index(month):02}-{day:02}", cellFormatDown)
                    #write smoothed counts
                    countsSheet.write(rowDays, 13, round(sum(counts[countsLocation-gvar.config["smoothingFactorCounts"]:countsLocation+gvar.config["smoothingFactorCounts"]+1])/(gvar.config["smoothingFactorCounts"]*2+1), 2), cellFormatDown)
                    
                    rowDays+=1
                countsLocation+=1
    
    
    countsSheet.write(1, 4, 'Total Average', cellFormatUp)
    countsSheet.write(2, 4, round(sum(counts)/len(counts), 2), cellFormatDown)
    
    countsSheet.write(1, 5, 'Total Sum', cellFormatUp)
    countsSheet.write(2, 5, sum(counts), cellFormatDown)
    
    countsSheet.write(1, 6, 'Total Days', cellFormatUp)
    countsSheet.write(2, 6, len(counts), cellFormatDown)
    
    countsSheet.write(1, 7, 'Median', cellFormatUp)
    countsSheet.write(2, 7, sorted(counts)[len(counts)//2], cellFormatDown)
    
    countsSheet.write(1, 8, 'Max', cellFormatUp)
    countsSheet.write(2, 8, max(counts), cellFormatDown)
    
    countsSheet.write(1, 9, 'Standard Deviation', cellFormatUp)
    countsSheet.write(2, 9, round((sum([(i-sum(counts)/len(counts))**2 for i in counts])/(len(counts)-1))**0.5, 4), cellFormatDown)
    
    countsSheet.write(1, 10, 'Variance', cellFormatUp)
    countsSheet.write(2, 10, round((sum([(i-sum(counts)/len(counts))**2 for i in counts])/(len(counts)-1)), 4), cellFormatDown)
    
    countsSheet.write(1, 11, 'Coefficient of Variation', cellFormatUp)
    countsSheet.write(2, 11, round((sum([(i-sum(counts)/len(counts))**2 for i in counts])/(len(counts)-1))**0.5/(sum(counts)/len(counts)), 4), cellFormatDown)
    
    countsSheet.write(1, 12, 'Skewness', cellFormatUp)
    countsSheet.write(2, 12, round((sum([(i-sum(counts)/len(counts))**3 for i in counts])/(len(counts)-1))/(round((sum([(i-sum(counts)/len(counts))**2 for i in counts])/(len(counts)-1)), 4))**(3/2), 4), cellFormatDown)
    
    countsSheet.write(1, 13, 'Kurtosis', cellFormatUp)
    countsSheet.write(2, 13, round((sum([(i-sum(counts)/len(counts))**4 for i in counts])/(len(counts)-1))/(round((sum([(i-sum(counts)/len(counts))**2 for i in counts])/(len(counts)-1)), 4))**2, 4), cellFormatDown)
    
    
    countsSheet.autofit()
    
    #total counts
    createChart("Day Counts", "counts", ["B", 3, "B", rowDays], ["C", 3, "C", rowDays], "Day Counts Chart", "Counts", "Date")
    
    #smoothed counts
    createChart("Smoothed Counts", "counts", ["M", 6, "M", rowDays], ["N", 6, "N", rowDays], "Smoothed Counts Chart", "Smoothed Counts", "Date")
    
    #month averages
    createChart("Month Averages", "counts", ["E", 6, "E", rowMonths], ["G", 6, "G", rowMonths], "Month Averages Chart", "Month Average Counts", "Month")
    
    #year averages
    createChart("Year Averages", "counts", ["I", 6, "I", rowYears], ["K", 6, "K", rowYears], "Year Averages Chart", "Year Average Counts", "Year")

    ############################################ HOURS SHEET #########################################################
    hoursSheet = gvar.workbook.add_worksheet("hours")
    hoursSheet.write(0, 0, 'Hours')

    #upper row
    row = 2
    hour=datetime.datetime(2000,1,1,hour=0, minute=0)
    for i in range(0,gvar.config["hourDivisions"]):
        nextHour=hour+datetime.timedelta(hours=24/gvar.config["hourDivisions"])
        hoursSheet.write(1, 3+i, hour.strftime("%H:%M")+"-"+nextHour.strftime("%H:%M"), cellFormatUp)
        hour=nextHour
    
    totalHourResults=[0]*gvar.config["hourDivisions"]
    
    #data per month
    for year in jsonFile["data"]:
        hoursSheet.write(row, 1, year, cellFormatDown)
        #months
        for month in jsonFile["data"][year]:
            hoursSheet.write(row, 2, month, cellFormatDown)

            #create list of 0s
            hourResults=[0]*gvar.config["hourDivisions"]

            #days
            for day in jsonFile["data"][year][month]:
                #each hour
                for hour in jsonFile["data"][year][month][day][1]:
                    hourResults[int((int(hour.split(":")[0])*60+int(hour.split(":")[1]))/(1440/gvar.config["hourDivisions"]))]+=1
            
            col=3
            for i in hourResults:
                hoursSheet.write(row, col, i, cellFormatDown)
                totalHourResults[col-3]+=i
                col+=1
            
            row+=1
            
    #total data
    #upper row
    row+=1
    hour=datetime.datetime(2000,1,1,hour=0, minute=0)
    for i in range(0,gvar.config["hourDivisions"]):
        nextHour=hour+datetime.timedelta(hours=24/gvar.config["hourDivisions"])
        hoursSheet.write(row, 3+i, hour.strftime("%H:%M")+"-"+nextHour.strftime("%H:%M"), cellFormatUp)
        hour=nextHour
    #data
    col=3
    row+=1
    hoursSheet.write(row, 2, "Total", cellFormatDown)
    for i in totalHourResults:
        hoursSheet.write(row, col, i, cellFormatDown)
        col+=1
    
    #smoothed data
    #upper row
    row+=2
    hour=datetime.datetime(2000,1,1,hour=0, minute=0)
    for i in range(0,gvar.config["hourDivisions"]):
        nextHour=hour+datetime.timedelta(hours=24/gvar.config["hourDivisions"])
        hoursSheet.write(row, 3+i, hour.strftime("%H:%M")+"-"+nextHour.strftime("%H:%M"), cellFormatUp)
        hour=nextHour
    #data
    col=3
    row+=1
    hoursSheet.write(row, 2, "Smoothed", cellFormatDown)
    for i in range(0,gvar.config["hourDivisions"]):
        if i>=gvar.config["smoothingFactorHours"] and i<gvar.config["hourDivisions"]-gvar.config["smoothingFactorHours"]:
            hoursSheet.write(row, col, round(sum(totalHourResults[i-gvar.config["smoothingFactorHours"]:i+gvar.config["smoothingFactorHours"]+1])/(gvar.config["smoothingFactorHours"]*2+1), 2), cellFormatDown)
        else:
            hoursSheet.write(row, col, 0, cellFormatDown)
        col+=1
    
    hoursSheet.autofit()

    createChart("Hours Distribution", "hours", ["D", 2, xlsxwriter.utility.xl_col_to_name(2+gvar.config["hourDivisions"]), 2], ["D", row-2, xlsxwriter.utility.xl_col_to_name(2+gvar.config["hourDivisions"]), row-2], "Hours Chart", "Hours Count", "Hours")

    createChart("Hours Distribution Smoothed", "hours", ["D", 2, xlsxwriter.utility.xl_col_to_name(2+gvar.config["hourDivisions"]), 2], ["D", row+1, xlsxwriter.utility.xl_col_to_name(2+gvar.config["hourDivisions"]), row+1], "Hours Chart Smoothed", "Hours Count Smoothed", "Hours")
    
    gvar.workbook.close()