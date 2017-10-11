import openpyxl
import datetime
from openpyxl.chart import ScatterChart, Reference, Series


def milestone_swimlane(col, start_row, project_number, newwb):

    newsheet = newwb.active

    wb = openpyxl.load_workbook('/home/lemon/Downloads/compiled_master_2017-07-18_Q1 Apr - Jun 2017 FOR Q2 COMMISSION DO NOT CHANGE.xlsx')
    sheet = wb.active

    # print project title
    newsheet.cell(row=1, column=1, value=sheet.cell(row=1, column=col).value)

    x = start_row
    for i in range(90, 269, 6):
        val = sheet.cell(row=i, column=col).value
        newsheet.cell(row=x, column=1, value=val)
        x += 1
    x = start_row
    for i in range(91, 269, 6):
        val = sheet.cell(row=i, column=col).value
        newsheet.cell(row=x, column=2, value=val)
        x += 1

    today = datetime.datetime.today()
    counter = 2
    for i in range(91, 269, 6):
        time_line_date = sheet.cell(row=i, column=col).value
        try:
            difference = (time_line_date - today).days
            print(difference)
            if difference in range(1, 5000):
                newsheet.cell(row=counter, column=3, value=difference)
        except TypeError:
                pass
        finally:
            counter += 1

    for i in range(2, 32):
        newsheet.cell(row=i, column=4, value=project_number)


#       chart = ScatterChart()
#       chart.title = "Scatter Chart"
#       chart.style = 1
#       chart.x_axis.title = 'Date'
#       chart.y_axis.title = 'Project No'
#
#       xvalues = Reference(newsheet, min_col=3, min_row=1, max_row=30)
#       for i in range(1, 31):
#           values = Reference(newsheet, min_col=4, min_row=1, max_row=30)
#           series = Series(values, xvalues, title_from_data=True)
#           chart.series.append(series)
#
#       newsheet.add_chart(chart, "E10")

    return newwb


wb = openpyxl.Workbook()
proj_num = 1
st_row = 2
wb = milestone_swimlane(3, st_row, proj_num, wb)
wb.save('output.xlsx')



