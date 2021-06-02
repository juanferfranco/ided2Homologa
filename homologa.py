import xlsxwriter
import pandas as pd
from bs4 import BeautifulSoup

student_name = ""
student_id = ""
# "Experiencias Interactivas" "Videojuegos" "Animación"
student_line = "Videojuegos"

excel_file = "".join([student_id , student_name, student_line,".xlsx"])

workbook = xlsxwriter.Workbook(excel_file)
worksheet = workbook.add_worksheet()

title_format = workbook.add_format({
    'bold': 0,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': 'gray',
    'color': 'white',
    'font_name': 'Verdana',
    'font_size': 12})

cell_format = workbook.add_format({
    'bold': 0,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': 'white',
    'color': 'black',
    'text_wrap': 1,
    'font_name': 'Verdana',
    'font_size': 8})

cell_yes_format = workbook.add_format({
    'bold': 0,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': '#00C800',
    'color': 'black',
    'text_wrap': 1,
    'font_name': 'Verdana',
    'font_size': 8})

student_courses = ""
if student_line == "Experiencias Interactivas":
    student_courses = r'./ided2_experiencias.xlsx'
elif student_line == "Videojuegos":
    student_courses = r'./ided2_videojuegos.xlsx'
elif student_line == "Animación":
    student_courses = r'./ided2_animacion.xlsx'
else:
    print("".join(["Line: ",student_line, " is not supported"]))
    quit()

data = pd.read_excel (student_courses)
courses = data.values

page = open('capp.html',encoding='UTF-8').read()

#encoding='UTF-8'

soup = BeautifulSoup(page, 'html.parser')
tables = soup.find_all('table', attrs={'class': 'datadisplaytable'})

capp = {}

for table in tables:

   output_rows = []
   for table_row in table.findAll('tr'):
      columns = table_row.findAll('td')
      output_row = []
      for column in columns:
         output_row.append(column.text.strip())
      if output_row and (output_row[0] == 'Si' or output_row[0] == 'No'):
         output_rows.append(output_row)
         capp[output_row[1]] = output_row[0]

worksheet.merge_range('B1:I2', 'Ingeniería en Diseño de Entretenimiento Digital - LINEA '+ student_line.upper(), title_format)
worksheet.set_column('A:J', 15)

# Generate worksheet title

for i in range(8):
    worksheet.write(3, i+1, 'Sem ' + str(i+1), cell_format)

# row,column 0,0 (first course)
r0 = 4
c0 = 1

#print(courses)

credits = [0]*12
for course in courses:
    #print(course)
    if course[0] in capp:
        if capp[course[0]] == "Si":
            worksheet.write(course[2] + r0, course[3] + c0, str(course[0]), cell_yes_format)
            worksheet.set_row(course[2] + r0,40)
        else:
            worksheet.write(course[2] + r0, course[3] + c0, str(course[0]), cell_format)
            worksheet.set_row(course[2] + r0, 40)
    else:
        print("".join(["Course: " , course[0], " isn't in capp"]))
    
    worksheet.write(course[2] + r0 + 1, course[3] + c0, str(course[1]), cell_format)
    credits[course[3]] = credits[course[3]] + course[1]

workbook.close()    