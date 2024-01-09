import scripts.gvar as gvar
import scripts.xlsxUtilites.xlsxGvar as xlsxGvar
import scripts.configuration.configuration as configuration
from scripts.xlsxUtilites.createChart import createChart
import scripts.xlsxUtilites.xlsxWritting.cellFormat as cellFormat

def createCountsSheet(jsonFile: dict) -> None:
    """Creates the counts sheet in the excel file with all the data and charts

    Args:
        jsonFile (dict): Dictionary with json format to be written to the excel file
    """

    countsSheet = xlsxGvar.workbook.add_worksheet("counts")

    countsSheet.write(0, 0, 'Counts')
    
    countsSheet.write(1, 1, 'Date', cellFormat.cellFormatUp)
    countsSheet.write(1, 2, 'Counts', cellFormat.cellFormatUp)
    
    countsSheet.write(4, 4, 'Month', cellFormat.cellFormatUp)
    countsSheet.write(4, 5, 'Counts', cellFormat.cellFormatUp)
    countsSheet.write(4, 6, 'Average', cellFormat.cellFormatUp)
    
    countsSheet.write(4, 8, 'Year', cellFormat.cellFormatUp)
    countsSheet.write(4, 9, 'Counts', cellFormat.cellFormatUp)
    countsSheet.write(4, 10, 'Average', cellFormat.cellFormatUp)

    rowDays = 2
    rowMonths = 5
    rowYears = 5
    
    countsThisYear=0
    
    counts = []
    
    #per day counts
    for year in jsonFile["data"]:
        for month in jsonFile["data"][year]:
            for day in jsonFile["data"][year][month]:
                #write date
                countsSheet.write(rowDays, 1, f"{year}-{gvar.months.index(month):02}-{day:02}", cellFormat.cellFormatDown)
                #write counts day
                countsSheet.write(rowDays, 2, jsonFile["data"][year][month][day][0], cellFormat.cellFormatDown)
                counts.append(jsonFile["data"][year][month][day][0])
                
                rowDays+=1
                
            #write month
            countsSheet.write(rowMonths, 4, f"{month}-{year}", cellFormat.cellFormatDown)
            #write counts month
            countsSheet.write(rowMonths, 5, sum([jsonFile['data'][year][month][day][0] for day in jsonFile['data'][year][month]]), cellFormat.cellFormatDown)
            countsThisYear+=sum([jsonFile['data'][year][month][day][0] for day in jsonFile['data'][year][month]])
            #write average counts month
            countsSheet.write(rowMonths, 6, round(sum([jsonFile['data'][year][month][day][0] for day in jsonFile['data'][year][month]])/len(jsonFile['data'][year][month]), 2), cellFormat.cellFormatDown)
            
            rowMonths+=1
            
        #write year
        countsSheet.write(rowYears, 8, year, cellFormat.cellFormatDown)
        #write counts year
        countsSheet.write(rowYears, 9, countsThisYear, cellFormat.cellFormatDown)
        #write average counts year
        countsSheet.write(rowYears, 10, round(countsThisYear/sum([len(jsonFile['data'][year][month]) for month in jsonFile['data'][year]]), 2), cellFormat.cellFormatDown)
        
        countsThisYear=0
        rowYears+=1
    
    countsSheet.write(4, 12, 'Date', cellFormat.cellFormatUp)
    countsSheet.write(4, 13, 'Smoothed Counts', cellFormat.cellFormatUp)
    
    rowDays = 5
    countsLocation=0
    
    #smoothed counts
    for year in jsonFile["data"]:
        for month in jsonFile["data"][year]:
            for day in jsonFile["data"][year][month]:
                if countsLocation>=configuration.config["smoothingFactorCounts"]:
                    #write date
                    countsSheet.write(rowDays, 12, f"{year}-{gvar.months.index(month):02}-{day:02}", cellFormat.cellFormatDown)
                    #write smoothed counts
                    countsSheet.write(rowDays, 13, round(sum(counts[countsLocation-configuration.config["smoothingFactorCounts"]:countsLocation+configuration.config["smoothingFactorCounts"]+1])/(configuration.config["smoothingFactorCounts"]*2+1), 2), cellFormat.cellFormatDown)
                    
                    rowDays+=1
                countsLocation+=1
    
    
    countsSheet.write(1, 4, 'Total Average', cellFormat.cellFormatUp)
    countsSheet.write(2, 4, round(sum(counts)/len(counts), 2), cellFormat.cellFormatDown)
    
    countsSheet.write(1, 5, 'Total Sum', cellFormat.cellFormatUp)
    countsSheet.write(2, 5, sum(counts), cellFormat.cellFormatDown)
    
    countsSheet.write(1, 6, 'Total Days', cellFormat.cellFormatUp)
    countsSheet.write(2, 6, len(counts), cellFormat.cellFormatDown)
    
    countsSheet.write(1, 7, 'Median', cellFormat.cellFormatUp)
    countsSheet.write(2, 7, sorted(counts)[len(counts)//2], cellFormat.cellFormatDown)
    
    countsSheet.write(1, 8, 'Max', cellFormat.cellFormatUp)
    countsSheet.write(2, 8, max(counts), cellFormat.cellFormatDown)
    
    countsSheet.write(1, 9, 'Standard Deviation', cellFormat.cellFormatUp)
    countsSheet.write(2, 9, round((sum([(i-sum(counts)/len(counts))**2 for i in counts])/(len(counts)-1))**0.5, 4), cellFormat.cellFormatDown)
    
    countsSheet.write(1, 10, 'Variance', cellFormat.cellFormatUp)
    countsSheet.write(2, 10, round((sum([(i-sum(counts)/len(counts))**2 for i in counts])/(len(counts)-1)), 4), cellFormat.cellFormatDown)
    
    countsSheet.write(1, 11, 'Coefficient of Variation', cellFormat.cellFormatUp)
    countsSheet.write(2, 11, round((sum([(i-sum(counts)/len(counts))**2 for i in counts])/(len(counts)-1))**0.5/(sum(counts)/len(counts)), 4), cellFormat.cellFormatDown)
    
    countsSheet.write(1, 12, 'Skewness', cellFormat.cellFormatUp)
    countsSheet.write(2, 12, round((sum([(i-sum(counts)/len(counts))**3 for i in counts])/(len(counts)-1))/(round((sum([(i-sum(counts)/len(counts))**2 for i in counts])/(len(counts)-1)), 4))**(3/2), 4), cellFormat.cellFormatDown)
    
    countsSheet.write(1, 13, 'Kurtosis', cellFormat.cellFormatUp)
    countsSheet.write(2, 13, round((sum([(i-sum(counts)/len(counts))**4 for i in counts])/(len(counts)-1))/(round((sum([(i-sum(counts)/len(counts))**2 for i in counts])/(len(counts)-1)), 4))**2, 4), cellFormat.cellFormatDown)
    
    
    countsSheet.autofit()
    
    #total counts
    createChart("Day Counts", "counts", ["B", 3, "B", rowDays], ["C", 3, "C", rowDays], "Day Counts Chart", "Counts", "Date")
    
    #smoothed counts
    createChart("Smoothed Counts", "counts", ["M", 6, "M", rowDays], ["N", 6, "N", rowDays], "Smoothed Counts Chart", "Smoothed Counts", "Date")
    
    #month averages
    createChart("Month Averages", "counts", ["E", 6, "E", rowMonths], ["G", 6, "G", rowMonths], "Month Averages Chart", "Month Average Counts", "Month")
    
    #year averages
    createChart("Year Averages", "counts", ["I", 6, "I", rowYears], ["K", 6, "K", rowYears], "Year Averages Chart", "Year Average Counts", "Year")
