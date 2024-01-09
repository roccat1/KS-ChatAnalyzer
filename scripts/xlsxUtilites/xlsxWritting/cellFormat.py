import scripts.xlsxUtilites.xlsxGvar as xlsxGvar

cellFormatDown = None
cellFormatUp = None

def createFormats() -> None:
    """Creates the cell formats for the excel file"""
    global cellFormatDown
    global cellFormatUp
    
    #cell format for upper part
    cellFormatUp = xlsxGvar.workbook.add_format()
    cellFormatUp.set_border(1)
    cellFormatUp.set_bg_color('#C6C6C6')
    
    #cell format for lower part
    cellFormatDown = xlsxGvar.workbook.add_format()
    cellFormatDown.set_border(1)
    cellFormatDown.set_bg_color('#E1E1E1')