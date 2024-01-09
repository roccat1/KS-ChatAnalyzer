import xlsxwriter

import scripts.xlsxUtilites.xlsxGvar as xlsxGvar
import scripts.configuration.configuration as configuration
import scripts.xlsxUtilites.xlsxWritting.countsSheet as countsSheet
import scripts.xlsxUtilites.xlsxWritting.hoursSheet as hoursSheet
import scripts.xlsxUtilites.xlsxWritting.cellFormat as cellFormat

def writeJsonToXls(jsonFile: dict) -> None:
    """Writes the json file to an excel file processing all the data and creating charts and sheets with statistics

    Args:
        jsonFile (dict): Dictionary with json format to be written to the excel file
    """
    xlsxGvar.workbook = xlsxwriter.Workbook(configuration.config["outputDirPath"]+configuration.config["outputExcelFileName"])

    cellFormat.createFormats()

    countsSheet.createCountsSheet(jsonFile)
    
    hoursSheet.createHoursSheet(jsonFile)

    xlsxGvar.workbook.close()