import win32com.client

# import os
# os.path.exists('T:\\')

files = ['C:/Users/Jonathan/Desktop/Book1.xlsx', 'C:/Users/Jonathan/Desktop/Book2.xlsx',
         'C:/Users/Jonathan/Desktop/Book3.xlsx']

counter = 1


def daily_refresh(workbook):

    # Start an instance of Excel
    xlapp = win32com.client.DispatchEx("Excel.Application")

    # Open the workbook in said instance of Excel
    wb = xlapp.workbooks.open(workbook)

    # Refresh all data connections.
    wb.RefreshAll()
    wb.Save()

    # Quit
    xlapp.Quit()

for file in files:
    daily_refresh(file)
    print('File ' + str(counter) + ' of ' + str(len(files)) + ' completed.')
    counter += 1


