import xlsxwriter
import pandas as pd
from bs4 import BeautifulSoup

student_name = "JHONY A. ATEHORTUA MEJIA"
student_id = "000318647"

# "Experiencias Interactivas" "Videojuegos" "Animación"
student_line = "Videojuegos"

infoAddDict = {}

def printInfoAdd():
    global infoAddDict

    infoAddPage = open('infoAdd.html',encoding='UTF-8').read()
    infoAddSoup = BeautifulSoup(infoAddPage, 'html.parser')
    infoAddtables = infoAddSoup.find_all('table', attrs={'class': 'datadisplaytable'})

    infoAddOutput_rows = []

    for table in infoAddtables:
        for table_row in table.findAll('tr'):
            columns = table_row.findAll('td')
            output_row = []
            for column in columns:
                output_row.append(column.text.strip())
                if output_row and len(output_row) == 6:
                    infoAddOutput_rows.append(output_row)

    for item in infoAddOutput_rows:
        try:
            nota = float(item[5])
            creditos = int(float(item[4]))
            if "pregrado" in item[3].lower() and nota >= 3.0 and creditos > 0:    
                infoAddDict[item[2]] = creditos
        except:
            pass

    print("#################################")        
    print("Información Adicional")
    print("#################################")

    for pair in infoAddDict.items():
        print(pair)



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

courses = {}
for item in data.values:
    courses[item[0]] = list(item[1:])

page = open('capp.html',encoding='UTF-8').read()
soup = BeautifulSoup(page, 'html.parser')
tables = soup.find_all('table', attrs={'class': 'datadisplaytable'})

output_rows = []

for table in tables:
   for table_row in table.findAll('tr'):
      columns = table_row.findAll('td')
      output_row = []
      for column in columns:
         output_row.append(column.text.strip())
      if output_row and (output_row[0] == 'Si' or output_row[0] == 'No' \
         or (output_row[0] == '' and len(output_row) >= 4 and len(output_row) <= 5)):
            output_rows.append(output_row)
      elif len(output_row) == 3:
        # rare case when homologate a course with 3 courses. The
        # course 3 is not saved in html table with first column
        # empty.
        output_row.insert(0,'')
        output_rows.append(output_row)
        print(output_row)

capp = {}
last_key = 'none'
capp[last_key] =[]

for row in output_rows:
    if row[0] == 'Si' or (row[0] == 'No' and len(row) >= 4 and len(row) <= 5):
        last_key = 'none'
        if row[1] in courses:
            last_key = row[1]
            capp[last_key] = [row[0], courses[row[1]][0], [ [row[2], row[3]] ] ]
    elif  row[0] == 'No':
        last_key = 'none'
        if row[1] in courses:
            capp[row[1]] = [row[0], courses[row[1]][0], [['none']]]    
    elif row[0] == '':
        if last_key != 'none':
            capp[last_key][2].append([row[1],row[2]])


capp = dict( filter( lambda elem: elem[0] in courses , capp.items()))




# https://stackabuse.com/how-to-print-colored-text-in-pythons
def printHom(text): 
    print("\033[48;5;2m{}\033[00m" .format(text))

def printHomCreditsOverflow(text): 
    print("\033[48;5;4m{}\033[00m" .format(text))

def printHomCreditsUnderflow(text): 
    print("\033[48;5;6m{}\033[00m" .format(text))

def printNoHomCredits(text): 
    print("\033[48;5;3m{}\033[00m" .format(text))

def printNoHom(text): 
    print("\033[48;5;1m{}\033[00m" .format(text))

def printChangos(text): 
    print("\033[48;5;7m{}\033[00m" .format(text))

totalCredits = 0

for pair in capp.items():
    creditsIded2 = pair[1][1]
    creditsIded1 = 0
    
    for item in pair[1][2]:
        if item[0] != 'none':
             creditsIded1 += int(float(item[1]))
    
    totalCredits += creditsIded1

    if pair[1][0] == 'Si' and creditsIded1 == creditsIded2:
       print(str(pair))
    elif pair[1][0] == 'Si' and creditsIded1 > creditsIded2:
        printHomCreditsOverflow(str(pair))
    elif pair[1][0] == 'Si' and creditsIded1 < creditsIded2:
        printHomCreditsUnderflow(str(pair))
    elif pair[1][0] == 'No' and creditsIded1 == 0:
        printNoHom(str(pair))
    elif pair[1][0] == 'No' and creditsIded1 > 0:
        printNoHomCredits(str(pair))
    else:
        printChangos(str(pair))

print("#################################")
print("Total approved credits: {}".format(totalCredits))

printInfoAdd()

outputExcel = "".join([student_id ,"-", student_name, "-",student_line,".xlsx"])
workbook = xlsxwriter.Workbook(outputExcel)
worksheet = workbook.add_worksheet()

title_format = workbook.add_format({
    'bold': 0,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': '#36393f', #gray discord
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
    'fg_color': '#2ecc71', # green discord
    'color': 'black',
    'text_wrap': 1,
    'font_name': 'Verdana',
    'font_size': 8})


worksheet.merge_range('B1:I2', 'Ingeniería en Diseño de Entretenimiento Digital - LINEA '+ student_line.upper(), title_format)
worksheet.set_column('A:J', 15)

# Generate worksheet title

for i in range(8):
    worksheet.write(3, i+1, 'Sem ' + str(i+1), cell_format)

# row,column 0,0 (first course) and table Height
r0 = 4
c0 = 1
tableHeight = 17

for course in courses:
    if course in capp:
        if capp[course][0] == "Si":
            worksheet.write( courses[course][1] + r0, courses[course][2] + c0, str(course), cell_yes_format)
            worksheet.set_row(courses[course][1] + r0,40)
        else:
            worksheet.write(courses[course][1] + r0, courses[course][2] + c0, str(course), cell_format)
            worksheet.set_row(courses[course][1] + r0, 40)
    else:
        print("Course {} isn't in capp".format(course))

    worksheet.write(courses[course][1] + r0 + 1, courses[course][2] + c0, str(courses[course][0]), cell_format)

worksheet.merge_range("B22:C22", "Créditos no utilizados", title_format)
i = 1
worksheet.write(r0 + tableHeight +i , c0, "Curso",cell_format )
worksheet.write(r0 + tableHeight +i , c0 + 1, "Créditos",cell_format )
i +=1
for pair in infoAddDict.items():
    worksheet.set_row(r0 + tableHeight +i,40)
    worksheet.write(r0 + tableHeight +i , c0, pair[0], cell_format)
    worksheet.write(r0 + tableHeight +i , c0 + 1, pair[1], cell_format)
    i += 1
    
workbook.close()

