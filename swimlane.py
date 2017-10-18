import os
import openpyxl
import datetime
from openpyxl.chart import ScatterChart, Reference, Series

HOME = os.path.abspath(os.path.expanduser('~'))
DESKTOP = os.path.join(HOME, 'Desktop')

NUMBER_OF_PROJECTS = 10


def milestone_swimlane(start_row, project_number, newwb, block_start_row=90,
                       interested_range=365):
    newsheet = newwb.active
    col = project_number + 1
    start_row = start_row + 1

    wb = openpyxl.load_workbook(
        os.path.join(
            DESKTOP, 'Q2_1718_master.xlsx'))
    sheet = wb.active

    # print project title
    newsheet.cell(row=start_row - 1, column=1, value=sheet.cell(row=1, column=col).value)
    print(sheet.cell(row=1, column=col).value)

    x = start_row
    for i in range(block_start_row, 269, 6):
        val = sheet.cell(row=i, column=col).value
        newsheet.cell(row=x, column=1, value=val)
        x += 1
    x = start_row
    for i in range(block_start_row + 1, 270, 6):
        val = sheet.cell(row=i, column=col).value
        newsheet.cell(row=x, column=2, value=val)
        x += 1

    today = datetime.datetime.today()
    current_row = start_row
    for i in range(91, 269, 6):
        time_line_date = sheet.cell(row=i, column=col).value
        try:
            difference = (time_line_date - today).days
            if difference in range(1, interested_range):
                newsheet.cell(row=current_row, column=3, value=difference)
        except TypeError:
            pass
        finally:
            current_row += 1

    for i in range(start_row, start_row + 30):
        newsheet.cell(row=i, column=4, value=project_number)

    return newwb, start_row


def _build_chart(sheet, start_row):
    xvalues = Reference(sheet, min_col=3, min_row=start_row, max_row=start_row + 29)
    values = Reference(sheet, min_col=4, min_row=start_row, max_row=start_row + 29)
    series = Series(values, xvalues, title_from_data=True)
    return series


def _start_cells():
    total_rows = NUMBER_OF_PROJECTS * 30
    for i in range(2, total_rows, 31):
        yield i


def _row_calc_chart(start_row):
    return start_row + 29


def _row_calc(project_number):
    if project_number == 1:
        return (1, 1)
    if project_number == 2:
        return (2, 32)
    else:
        return (project_number, (project_number + 30) + ((project_number - 2) * 30))


def main():
    wb = openpyxl.Workbook()
    start_generator = _start_cells()
    for p in range(1, 31):
        proj_num, st_row = _row_calc(p)
        wb = milestone_swimlane(st_row, proj_num, wb, block_start_row=90, interested_range=365)[0]

    chart = ScatterChart()
    chart.title = "Swimlane Chart"
    chart.style = 1
    chart.x_axis.title = 'Date'
    chart.y_axis.title = 'Project No'

    for p in range(1, 5):
        start_row = next(start_generator)
        series = _build_chart(wb.active, start_row)
        series.marker.symbol = "triangle"
        series.marker.graphicalProperties.solidFill = "FF0000"
        series.marker.graphicalProperties.line.solidFill = "FF0000"  # Marker outline
        chart.series.append(series)

    wb.active.add_chart(chart, "E1")
    wb.save(os.path.join(DESKTOP, 'output.xlsx'))


if __name__ == "__main__":
    main()
