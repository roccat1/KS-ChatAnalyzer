# KS-ChatAnalyzer
This program analyzes a whatsapp chat and converts it to an excel with charts.

## Read before using
To use this program you have to check config.json, important configurations are:
- formatDay-Month, set to true if your chat in in DD/MM/YY

## Code Guide
### main.py
#### browseFiles()
sets filname as the path of the chosen file

#### readWAChatDates(fileName)
file -> [datetime.datetime(2023, 2, 11, 10, 13), ...]
converts a file to a list of dates

#### datesToJson
[datetime.datetime(2023, 2, 11, 10, 13), ...] -> "2023": {"February": {"11": [2,["10:13","16:07"]]}}
converts list of dates to json file

#### saveJson(list)
saves json file

#### readJson()
returns the "json" list

#### writeJsonToXls(jsonFile)
"2023": {"February": {"11": [2,["10:13","16:07"]]}} -> excel
converts json file to excel

#### runProgram()
converts chat to excel