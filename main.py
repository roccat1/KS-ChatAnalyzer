import datetime, json, xlsxwriter, os, subprocess
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog

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
    label_file_explorer.configure(text="File Opened: "+filename)

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

            if config["formatDay-Month"]:
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

    row = 1


    #counts sheet
    countsSheet = workbook.add_worksheet("counts")
    countsSheet.write(0, 0, 'KS Counts')

    for year in jsonFile["data"]:
        countsSheet.write(row, 1, year)
        #months
        for month in jsonFile["data"][year]:
            countsSheet.write(row, 2, month)
            #days
            for day in jsonFile["data"][year][month]:
                countsSheet.write(row, 3, str(year)+"-"+str(months.index(month))+"-"+str(day))
                countsSheet.write(row, 4, jsonFile["data"][year][month][day][0])
                row+=1

    countsChart = workbook.add_chart({'type': 'line'})
    countsChart.add_series({
        'name': 'Total Counts',
        'categories': '=counts!$D$2:$D$'+str(row),
        'values': '=counts!$E$2:$E$'+str(row),
        'trendline': {
            'type': 'moving_average',
            'period': 23,
        }
    })
    countsChart.set_x_axis({'date_axis': True})
    countsChartSheet = workbook.add_chartsheet("Counts Chart")
    countsChartSheet.set_chart(countsChart)
    #fer fulla de grafics

    #hours sheet

    hoursSheet = workbook.add_worksheet("hours")
    hoursSheet.write(0, 0, 'KS Hours')

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
    datesRaw = readWAChatDates(filename)
    print("25% done")

    jsonResult = datesToJson(datesRaw)
    print("50% done")

    saveJson(jsonResult)
    print("75% done")

    writeJsonToXls(jsonResult)
    print("100% done")

    label_file_explorer.configure(text="Program executed")
    subprocess.run([os.path.join(os.getenv('WINDIR'), 'explorer.exe'), os.getcwd()+"\\output"])



if __name__=='__main__':
    window = tk.Tk()
    window.title('KS Project')
    window.geometry("700x300")
    window.config(background = "white")

    label_file_explorer = tk.Label(window, 
							text = "Check config.json to make sure it runs correctly",
							width = 100, height = 4, 
							fg = "blue")

    button_explore = tk.Button(window, 
						text = "Browse Files",
						command = browseFiles) 
    
    button_run = tk.Button(window, 
						text = "Run program",
						command = runProgram) 
    
    button_exit = tk.Button(window, 
					text = "Exit",
					command = exit) 

    label_file_explorer.grid(column = 1, row = 1)

    button_explore.grid(column = 1, row = 2)

    button_run.grid(column = 1, row = 3)

    button_exit.grid(column = 1,row = 4)

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