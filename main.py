import datetime, json, xlsxwriter, os, subprocess
import tkinter as tk
from tkinter import filedialog, messagebox

with open("config.json", 'r') as f: config = json.load(f)
filename = config["defaultFilePath"]

months = ["","January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]

def log(msg):
    print(msg)
    with open(config["outputDirPath"]+config["logPath"], "a", encoding="utf8") as f:
        f.write(msg+"\n")

if not os.path.exists(config["outputDirPath"]): os.makedirs(config["outputDirPath"])
with open(config["outputDirPath"]+config["logPath"], "w", encoding="utf8") as f: f.write("Program started\n")

def browseFiles():
    global filename 
    filename = filedialog.askopenfilename(initialdir = os.getcwd(),
                                            title = "Select a File",
                                            filetypes = (("Text files",
                                            "*.txt*"),
                                            ("all files",
                                            "*.*")))
	
	# Change label contents
    label_file_explorer.configure(text="File Opened: "+os.path.basename(filename))

#creates a list of dates from a whatsapp chat
def readWAChatDates(fileName):
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

            if dd_mmFormat.get():
                fullDate = datetime.datetime(int("20"+ date[2]), int(date[1]), int(date[0]), int(time[0]), int(time[1]))
            else:
                fullDate = datetime.datetime(int("20"+ date[2]), int(date[0]), int(date[1]), int(time[0]), int(time[1]))
            
            result.append(fullDate)
        except:
            log("error reading line: "+line)
    return result

#creates a json from a list of dates
def datesToJson(dates):
    #create canvas for json result
    #primera data
    date = datetime.datetime(dates[0].year, dates[0].month, dates[0].day)
    prevDate = datetime.datetime(1950,1,1)
    result={
        "metadata": {
            "name": config["projectName"],
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
            result["data"][date.year][months[date.month]]={}

        result["data"][date.year][months[date.month]][date.day]=[0,[]]

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
        result["data"][dateData[0].year][months[dateData[0].month]][dateData[0].day]=[dateData[1], dateData[2]]

    return result

def saveJson(list):
    with open(config["outputDirPath"]+config["outputJsonFileName"], "w") as fp:
        json.dump(list, fp, indent=2)

def readJson():
    with open("output/output.json", 'r') as f:
        list = json.load(f)
    return list

def writeJsonToXls(jsonFile):
    workbook = xlsxwriter.Workbook(config["outputDirPath"]+config["outputExcelFileName"])

    ############################################ COUNTS SHEET #########################################################
    countsSheet = workbook.add_worksheet("counts")
    countsSheet.write(0, 0, 'Counts')

    rowDays = 2
    rowMonths = 5
    rowYears = 5
    
    countsThisYear=0
    
    #cell foramat for upper part
    cellFormatUp = workbook.add_format()
    cellFormatUp.set_border(1)
    cellFormatUp.set_bg_color('#C6C6C6')
    
    #cell foramat for lower part
    cellFormatDown = workbook.add_format()
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
                countsSheet.write(rowDays, 1, f"{year}-{months.index(month):02}-{day:02}", cellFormatDown)
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
                if countsLocation>=config["smoothingFactor"]:
                    #write date
                    countsSheet.write(rowDays, 12, f"{year}-{months.index(month):02}-{day:02}", cellFormatDown)
                    #write smoothed counts
                    countsSheet.write(rowDays, 13, round(sum(counts[countsLocation-config["smoothingFactor"]:countsLocation+config["smoothingFactor"]+1])/(config["smoothingFactor"]*2+1), 2), cellFormatDown)
                    
                    rowDays+=1
                countsLocation+=1
    
    
    countsSheet.write(1, 4, 'Total Average', cellFormatUp)
    countsSheet.write(2, 4, round(sum(counts)/len(counts), 4), cellFormatDown)
    
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
    dayCountsChart = workbook.add_chart({'type': 'line'})
    dayCountsChart.add_series({
        'name': 'Day Counts',
        'categories': '=counts!$B$3:$B$'+str(rowDays),
        'values': '=counts!$C$3:$C$'+str(rowDays),
        'trendline': {
            'type': 'moving_average',
            'period': 23,
        }
    })
    dayCountsChart.set_x_axis({'date_axis': True})
    dayCountsChartSheet = workbook.add_chartsheet("Day Counts Chart")
    dayCountsChartSheet.set_chart(dayCountsChart)
    dayCountsChart.set_title({'name': 'Counts'})
    dayCountsChart.set_x_axis({'name': 'Date'})
    
    #smoothed counts
    smoothedCountsChart = workbook.add_chart({'type': 'line'})
    smoothedCountsChart.add_series({
        'name': 'Smoothed Counts',
        'categories': '=counts!$M$6:$M$'+str(rowDays),
        'values': '=counts!$N$6:$N$'+str(rowDays),
        'trendline': {
            'type': 'moving_average',
            'period': 23,
        }
    })
    smoothedCountsChart.set_x_axis({'date_axis': True})
    smoothedCountsChartSheet = workbook.add_chartsheet("Smoothed Counts Chart")
    smoothedCountsChartSheet.set_chart(smoothedCountsChart)
    smoothedCountsChart.set_title({'name': 'Smoothed Counts'})
    smoothedCountsChart.set_x_axis({'name': 'Date'})
    
    #month averages
    monthCountsChart = workbook.add_chart({'type': 'line'})
    monthCountsChart.add_series({
        'name': 'Month Averages',
        'categories': '=counts!$E$6:$E$'+str(rowMonths),
        'values': '=counts!$G$6:$G$'+str(rowMonths),
        'trendline': {
            'type': 'moving_average',
            'period': 23,
        }
    })
    monthCountsChart.set_x_axis({'date_axis': True})
    monthCountsChartSheet = workbook.add_chartsheet("Month Averages Chart")
    monthCountsChartSheet.set_chart(monthCountsChart)
    monthCountsChart.set_title({'name': 'Month Average Counts'})
    monthCountsChart.set_x_axis({'name': 'Month'})
    
    #year averages
    yearCountsChart = workbook.add_chart({'type': 'line'})
    yearCountsChart.add_series({
        'name': 'Year Averages',
        'categories': '=counts!$I$6:$I$'+str(rowYears),
        'values': '=counts!$K$6:$K$'+str(rowYears),
        'trendline': {
            'type': 'moving_average',
            'period': 23,
        }
    })
    yearCountsChart.set_x_axis({'date_axis': True})
    yearCountsChartSheet = workbook.add_chartsheet("Year Averages Chart")
    yearCountsChartSheet.set_chart(yearCountsChart)
    yearCountsChart.set_title({'name': 'Year Average counts'})
    yearCountsChart.set_x_axis({'name': 'Year'})




    ############################################ HOURS SHEET #########################################################
    hoursSheet = workbook.add_worksheet("hours")
    hoursSheet.write(0, 0, 'Hours')

    row = 2
    hour=datetime.datetime(2000,1,1,hour=0, minute=0)
    for i in range(0,config["hourDivisions"]):
        nextHour=hour+datetime.timedelta(hours=24/config["hourDivisions"])
        hoursSheet.write(1, 3+i, hour.strftime("%H:%M")+"-"+nextHour.strftime("%H:%M"))
        hour=nextHour
    #years
    for year in jsonFile["data"]:
        hoursSheet.write(row, 1, year)
        #months
        for month in jsonFile["data"][year]:
            hoursSheet.write(row, 2, month)

            hourResults=[]
            for i in range(0,config["hourDivisions"]):
                hourResults.append(0)

            #days
            for day in jsonFile["data"][year][month]:
                #each hour
                for hour in jsonFile["data"][year][month][day][1]:
                    hourResults[int((int(hour.split(":")[0])*60+int(hour.split(":")[1]))/(1440/config["hourDivisions"]))]+=1
            
            col=3
            for i in hourResults:
                hoursSheet.write(row, col, i)
                col+=1
            
            row+=1
    for i in range(0,config["hourDivisions"]): 
        hoursSheet.write_formula(row, 3+i, "=SUM("+xlsxwriter.utility.xl_col_to_name(i+3)+"3:"+xlsxwriter.utility.xl_col_to_name(i+3)+str(row)+")")


    hoursChart = workbook.add_chart({'type': 'line'})
    hoursChart.add_series({
        'name': 'Hours Distribution',
        'categories': '=hours!$D$2:$'+xlsxwriter.utility.xl_col_to_name(2+config["hourDivisions"])+'$2',
        'values': '=hours!$D$'+str(row+1)+':$'+xlsxwriter.utility.xl_col_to_name(2+config["hourDivisions"])+'$'+str(row+1),
        'trendline': {
            'type': 'moving_average',
            'period': 23,
        }
    })
    hoursChartSheet = workbook.add_chartsheet("Hours Chart")
    hoursChartSheet.set_chart(hoursChart)

    
    workbook.close()

def runProgram():
    log("program executed... ")
    try:
        datesRaw = readWAChatDates(filename)
        jsonResult = datesToJson(datesRaw)
        saveJson(jsonResult)
        writeJsonToXls(jsonResult)

        label_file_explorer.configure(text="Program executed")
        messagebox.showinfo("Done!", str(os.path.basename(filename))+" has been analyzed successfully!")
        subprocess.run([os.path.join(os.getenv('WINDIR'), 'explorer.exe'), os.getcwd()+"\\output"])
        log("successfully")
    except Exception as e:
        log("with an error :(")
        label_file_explorer.configure(text="ERROR(is the format correct?/check log/terminal)")
        log("ERROR: "+str(e))

if __name__=='__main__':
    window = tk.Tk()
    window.title('KS Project')
    window.geometry("700x300")
    window.config(background = "turquoise2")
    window.iconbitmap("assets/icon.ico")

    label_file_explorer = tk.Label(window, 
							text = "KS-ChatAnalyzer",
							width = 44, height = 2, 
							fg = "black",
                            background="pale green",
                            font=("Arial", 20)
        )

    button_explore = tk.Button(window, 
                        text = "Browse Files",
                        command = browseFiles,
                        width = 40, height = 2
                        )

    button_run = tk.Button(window, 
						text = "Run program",
						command = runProgram,
                        width = 40, height = 2) 
    
    button_exit = tk.Button(window, 
					text = "Exit",
					command = exit,
                    width = 40, height = 2) 

    label_file_explorer.grid(column = 1, row = 1, columnspan=2)

    button_explore.grid(column = 1, row = 2, columnspan=2)

    button_run.grid(column = 1, row = 4, columnspan=2)

    button_exit.grid(column = 1,row = 5, columnspan=2)

    dd_mmFormat = tk.BooleanVar(value=config["dd_mmFormat"])
    DDMM_Button = tk.Radiobutton(window, text="DD_MM_YY Format", variable=dd_mmFormat,
                                indicatoron=False, value=True, width=19, height = 2)
    MMDD_Button = tk.Radiobutton(window, text="MM_DD_YY Format", variable=dd_mmFormat,
                                indicatoron=False, value=False, width=19, height = 2)
    
    DDMM_Button.grid(column = 1,row = 3, sticky="e")
    MMDD_Button.grid(column = 2,row = 3, sticky="w")
    # Let the window wait for any events
    window.mainloop()

'''
git status

git fetch  //comprovar
git pull

git add .
git commit <-m msg>
git push

https://bluuweb.github.io/tutorial-github/

ctr+ç #
ctr+'¡

xlsxwriter.utility.xl_col_to_name(27)
'''

'''
def popup_bonus():
    win = tk.Toplevel()
    win.geometry("300x50")
    win.wm_title("Popup")

    l = tk.Entry(win, text="Input", width=50)
    l.grid(row=0, column=1)

    b = ttk.Button(win, text="Okay", command=win.destroy,width=50)
    b.grid(row=1, column=1)
'''