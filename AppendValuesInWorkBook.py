from openpyxl import Workbook, load_workbook

wb = Workbook()
ws = wb.active
ws.title = "Marks"

ws.append(['Name', 'English', 'Tamil', 'Match', 'Science', 'Social', 'Total', 'Grade'])
ws.append(['Mari', 98, 86, 95, 97, 98, 474, 'A+'])
ws.append(['Shelby', 90, 80, 72, 83, 68, 393, 'A'])
ws.append(['Thomas', 70, 58, 64, 97, 45, 334, 'B'])
ws.append(['Logesh', 65, 54, 67, 45, 58, 289, 'C'])
ws.append(['Raja', 60, 84, 73, 80, 76, 373, 'A'])

wb.save('result.xlsx')
