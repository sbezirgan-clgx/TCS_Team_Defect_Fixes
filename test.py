import my_date
import xlsxwriter
hello = my_date.My_date()
print(hello.get_query_Start())
print(hello.get_query_end())


print(hello.get_query_Start2())
print(hello.get_query_end2())



# Workbook() takes one, non-optional, argument
# which is the filename that we want to create.
workbook = xlsxwriter.Workbook(r'\\filer-diablo-prd\data_acquisition\QA\EagleQC\TCS_defect_fixes\Week_08-21-2023_to_08-27-2023\TCS_Defect_Fixes_QCPI_Blank.xlsx')

# The workbook object is then used to add new
# worksheet via the add_worksheet() method.
worksheet = workbook.add_worksheet()

# Use the worksheet object to write
# data via the write() method.
worksheet.write('A1', 'No data has been found for the given week')


# Finally, close the Excel file
# via the close() method.
workbook.close()