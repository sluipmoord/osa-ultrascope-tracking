import datetime
import xlwt
import random

# generates random data in one second intervals

workbook = xlwt.Workbook()
sheet = workbook.add_sheet("Sheet1")
start = datetime.datetime(2015, 10, 1, 1, 0, 0)
end = datetime.datetime(2015, 10, 5, 5, 20, 0)

row = 0
sheet_num = 1
while start <= end:
    if row == 65536:
        row = 0
        sheet_num += 1
        sheet = workbook.add_sheet("Sheet{0}".format(sheet_num))

    sheet.write(row, 0, start.year)
    sheet.write(row, 1, start.month)
    sheet.write(row, 2, start.day)
    sheet.write(row, 3, start.hour)
    sheet.write(row, 4, start.minute)
    sheet.write(row, 5, start.second)
    sheet.write(row, 6, random.randrange(34, 37, 1))
    sheet.write(row, 7, random.randrange(10, 20, 1))
    sheet.write(row, 8, -9)
    sheet.write(row, 9, random.randrange(20, 25, 1))
    sheet.write(row, 10, random.randrange(0, 60, 1))
    sheet.write(row, 11, random.randrange(0, 60, 1))
    sheet.write(row, 12, random.randrange(0, 60, 1))
    sheet.write(row, 13, random.randrange(0, 60, 1))
    row += 1
    start = start + datetime.timedelta(seconds=1)

workbook.save('random_data.xls')
