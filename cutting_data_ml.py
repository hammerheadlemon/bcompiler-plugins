import openpyxl
import datetime
from openpyxl.chart import ScatterChart, Reference, Series


def milestone_swimlane(col, start_row, project_number, newwb):

    newsheet = newwb.active

    wb = openpyxl.load_workbook('/home/lemon/Downloads/compiled_master_2017-07-18_Q1 Apr - Jun 2017 FOR Q2 COMMISSION DO NOT CHANGE.xlsx')
    sheet = wb.active

    x = 1
    for i in range(90, 269, 6):
        val = sheet.cell(row=i, column=col).value
        newsheet.cell(row=x, column=1, value=val)
        x += 1
    x = 1
    for i in range(91, 269, 6):
        val = sheet.cell(row=i, column=col).value
        newsheet.cell(row=x, column=2, value=val)
        x += 1

    today = datetime.datetime.today()
    import pdb; pdb.set_trace()  # XXX BREAKPOINT
    for i in range(91, 269, 6):
        time_line_date = sheet.cell(row=i, column=2).value
        try:
            difference = (time_line_date - today).days
            print(difference)
            if difference in range(1, 5000):
                newsheet.cell(row=i, column=3, value=difference)
        except TypeError:
                pass

    for i in range(1, 30):
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

st_row = 90
proj_num = 1
newwb = openpyxl.Workbook()
for column in range(2, 33):
        print("Doing project {}".format(proj_num))
        wb = milestone_swimlane(column, st_row, proj_num, newwb)
        st_row = st_row + wb[1]
        proj_num = proj_num + 1
        newwb = wb[0]
wb[0].save('output.xlsx')



