import scripts.gvar as gvar

def createChart(dataName: str, sheet: str, categories: list, values: list, chartsheetTitle: str, chartTitle: str, xAxisName: str) -> None:
    """Creates a chart in the excel file on the workbook

    Args:
        dataName (str): Name of the data to be displayed in the chart
        sheet (str): Name of the sheet where the data is
        categories (list): [startColumn, startRow, endColumn, endRow]
        values (list): [startColumn, startRow, endColumn, endRow]
        chartsheetTitle (str): Title of the chart sheet
        chartTitle (str): Title of the chart
        xAxisName (str): Name of the x axis
    """
    chart = gvar.workbook.add_chart({'type': 'line'})
    chart.add_series({
        'name': dataName,
        'categories': f'={sheet}!${categories[0]}${categories[1]}:${categories[2]}${categories[3]}',
        'values': f'={sheet}!${values[0]}${values[1]}:${values[2]}${values[3]}',
        'trendline': {
            'type': 'moving_average',
            'period': 23,
        }
    })
    chart.set_x_axis({'date_axis': True})
    chartSheet = gvar.workbook.add_chartsheet(chartsheetTitle)
    chartSheet.set_chart(chart)
    chart.set_title({'name': chartTitle})
    chart.set_x_axis({'name': xAxisName})