# importing all the required libraraies or modules
import csv    
import os                                 
from openpyxl.workbook import Workbook

# Dictionary that will be used to access pointer from grade(integer)
grade_dict = {'AA' : 10, 'AB' : 9, 'BB': 8, 'BC': 7,'CC' : 6,  'CD' : 5, 'DD' : 4,'DD*' : 4, 'F' : 0, 'F*' : 0, ' BB' : 8}

# Column Names for Semester Wise sheets
headings = ['Sl No.','Subject Code','Subject Name','L-T-P','Credit','Subject Type','Grade']

# Dictionary that will be used to access name from roll number
Dict_Name_RollNo = {}
file_pointer_nr = open("./names-roll.csv", 'r')
File_Name_RollNumber = csv.DictReader(file_pointer_nr)
for x in File_Name_RollNumber:
  Dict_Name_RollNo[x['Roll']] = x['Name']

# Dictionary that will be used to access subject name,ltp,credits from subject number
Dict_Subj_Master = {}
file_pointer_sm = open("./subjects_master.csv", 'r')
File_Subj_Master = csv.DictReader(file_pointer_sm)
for x in File_Subj_Master:
  Dict_Subj_Master[x['subno']] = [x['subname'],x['ltp'],x['crd']]
  file_grades = open("./grades.csv", 'r')

# Dictionary that with keys roll number and semester, so we can access details
# corresponding to each roll number and semester
Dicti_Grades = {}
File_Grades = csv.DictReader(file_grades)
for x in File_Grades:
  # if roll number is not present then add a dictionary
  # with the key as rollnumber and add info
  if not x['Roll'] in Dicti_Grades.keys():
    
    Dicti_Grades[x['Roll']] = {}
    # if semester is present then append the info
    # in the key as rollnumber and semester
    if x['Sem'] in Dicti_Grades[x['Roll']].keys():
      Dicti_Grades[x['Roll']][x['Sem']].append([ x['SubCode'],Dict_Subj_Master[x['SubCode']][0],Dict_Subj_Master[x['SubCode']][1],
                                          Dict_Subj_Master[x['SubCode']][2],x['Sub_Type'],x['Grade']])
     # if semester is not present then intialise info
    # in the key as rollnumber and semester  
    else:
      Dicti_Grades[x['Roll']][x['Sem']] = [[ x['SubCode'],Dict_Subj_Master[x['SubCode']][0],Dict_Subj_Master[x['SubCode']][1],
                                          Dict_Subj_Master[x['SubCode']][2],x['Sub_Type'],x['Grade']]]
  # if roll number is present then add the info
  # to the dict with key as roll no   
  else:
    # if semester is present then append the info
    # in the key as rollnumber and semester
    if x['Sem'] in Dicti_Grades[x['Roll']].keys():
      Dicti_Grades[x['Roll']][x['Sem']].append([ x['SubCode'], Dict_Subj_Master[x['SubCode']][0],Dict_Subj_Master[x['SubCode']][1],
                                            Dict_Subj_Master[x['SubCode']][2],x['Sub_Type'],x['Grade']])
    # if semester is not present then intialise info
    # in the key as rollnumber and semester  
    else:
       Dicti_Grades[x['Roll']][x['Sem']] = [[ x['SubCode'],Dict_Subj_Master[x['SubCode']][0],Dict_Subj_Master[x['SubCode']][1],
                                          Dict_Subj_Master[x['SubCode']][2],x['Sub_Type'], x['Grade']]]

# Loop for each roll number in the dictionary 
for x in Dicti_Grades.keys():

  New_Workbook = Workbook()

  Sem_Wise_Credits = ['Semester wise Credit Taken']
  Semester_Number = ['Semester No.']
  
  # Make output directory if it is not already present
  Name_Of_File = './output/' + x + '.xlsx'
  Folder_Output = os.path.dirname(Name_Of_File)
  if not os.path.exists(Folder_Output):
    os.makedirs(Folder_Output)
  
  Total_Credits_Taken = ['Total Credits Taken']
  SPI = ['SPI']
  CPI = ['CPI']

  # Adding Rollno, name from dictionary, and Discpline to the overall sheet
  overall = []
  overall.append(['Roll No.', x])
  overall.append(['Name of the Student',Dict_Name_RollNo[x]])
  overall.append(['Discipline',x[4:6]])

  
  for Semester in Dicti_Grades[x].keys():

    Semester_Number.append(Semester)
    Serial_Number = 1 
    
    # Create separate sheets for all semesters
    # Name of the sheet must be Sem (Semester_number) 
    New_Worksheet = New_Workbook.create_sheet()
    New_Worksheet.title = 'Sem' + Semester
    New_Worksheet.append(headings)

    creditss = 0
    spii = 0 

  # Calculating spi for the semester using grade pointer and credits
  # SPI = (C1 ∗ G1 + C2 ∗ G2 + C3 ∗ G3 + C4 ∗ G4)/(C1 + C2 + C3 + C4) 
    for Details in Dicti_Grades[x][Semester]:
      spii += int(Details[3]) * grade_dict[Details[5]]
      creditss += int(Details[3]) 
      Details.insert(0,Serial_Number)
      New_Worksheet.append(Details)
      Serial_Number += 1
  
    Sem_Wise_Credits.append(creditss)
    spi=spii/creditss
    # Adding spi 
    SPI.append(round(spi,2))

    # Calculating cpi for the semester using spi and credits
    # CPI = (SPI1 ∗ Credits in semester1 + SPI2 ∗ Credits in semester2 + ...)/(Total credits) 
    if type(Total_Credits_Taken[-1]) == str:

      Total_Credits_Taken.append(creditss)
      cpi=spii/creditss
      CPI.append(round(cpi,2))

    else:

      Total_Cred = Total_Credits_Taken[-1] + creditss
      calc_spi=round((CPI[-1]*Total_Credits_Taken[-1]+ spii)/Total_Cred,2)
      CPI.append(calc_spi)
      Total_Credits_Taken.append(Total_Cred)
  # Creating new worksheet for overall details
  New_Worksheet = New_Workbook['Sheet'] 
  New_Worksheet.title='Overall'
  # Adding semester number, credits, SPI,credits and CPI
  overall.append(Semester_Number),overall.append(Sem_Wise_Credits),overall.append(SPI)
  overall.append(Total_Credits_Taken),overall.append(CPI)

  for x in overall:
    New_Worksheet.append(x)
  # Save the changes in the workbook
  New_Workbook.save(filename=Name_Of_File)