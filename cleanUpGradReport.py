import xlrd
import xlwt
import sys
import re

countOfNonNumberGPAs = 0
abreviationToString = {

  "MA"    : "Master of Arts",
  "MAT"   : "Master of Arts in Teaching",
  "MALS"  : "Master of Arts in Liberal Studies",
  "MPS"   : "Master of Professional Studies",
  "MED"   : "Master of Education",
  "MPH"   : "Master of Public Health",
  "MS"    : "Master of Science",
  "MEH"   : "Master of Environmental Health",
  "MSEH"  : "Master of Science in Environmental Health",
  "MPA"   : "Master of Public Administration",
  "MCM"   : "Master of City Management",
  "MBA"   : "Master of Business Administration",
  "MACC"  : "Master of Accountancy",
  "MFA"   : "Master of Fine Arts",
  "MSN"   : "Master of Science in Nursing",
  "MSW"   : "Master of Social Work",
  "MSAH"  : "Master of Science in Allied Health",
  "EDS"   : "Education Specialist",
  "AUD"   : "Doctor of Audiology",
  "DPT"   : "Doctor of Physical Therapy",
  "EDD"   : "Doctor of Education",
  "DNP"   : "Doctor of Nursing Practice",
  "PHD"   : "Doctor of Philosophy",
  "DRPH"  : "Doctor of Public Health",
  "C4"    : "Certificate",
  "C3"    : "Certificate",
  "BA"    : "Bachelor of Arts",
  "BAS"   : "Bachelor of Applied Science",
  "BBA"   : "Bachelor of Business Administration",
  "BFA"   : "Bachelor of Fine Arts",
  "BGS"   : "Bachelor of General Studies",
  "BM"    : "Bachelor of Music",
  "BS"    : "Bachelor of Science",
  "BSDH"  : "Bachelor of Dental Hygiene",
  "BSED"  : "Bachelor of Science in Education",
  "BSEH"  : "Bachelor of Science in Environmental Health",
  "BSN"   : "Bachelor of Science in Nursing",
  "BSW"   : "Bachelor of Social Work"

}

print("Opening Workbook" + str(sys.argv[1]) + " and sheet " + str(sys.argv[2]))

workbook = xlrd.open_workbook(sys.argv[1])
worksheet = workbook.sheet_by_name(str(sys.argv[2]))

#print "Please enter the following"

#IDColumnIdentifier = raw_input("ID Column Identifier: ")
#FNameColumnIdentifier = raw_input("First Name Column Identifier: ")
#LNameColumnIdentifier = raw_input("Last Name Column Identifier: ")

IDColumnIdentifier       = "ID"
FNameColumnIdentifier    = "FNAME"
LNameColumnIdentifier    = "LNAME"
CollegeColumnIdentifier  = "COLLEGE"
DipNameColumnIdentifier  = "DIPLOMA_NAME"
DegreeColumnIdentifier   = "G_DEGREE"
MajorColumnIdentifier    = "G_MAJOR"
PrevDegreeCodeIdentifier = "PREVIOUS_DEGC_CODE"
PrevSchoolIdentifier     = "PREVIOUS_SCHOOL"
ThesisChairIdentifier    = "THESIS_CHAIR"
ThesisTitleIdentifier    = "THESIS_TITLE_LINE1"
ThesisTitleIdentifier2   = "THESIS_TITLE_LINE2"
ThesisTitleIdentifier3   = "THESIS_TITLE_LINE3"
ThesisTitleIdentifier4   = "THESIS_TITLE_LINE4"
CityIdentifier = "CITY"
ProgramIdentifier = "G_PROGRAM"
GPAIdentifier = "UG_GPA"

IDIndex = -1
FNameIndex = -1
LNameIndex = -1
CollegeIndex = -1
DipNameIndex = -1
DegreeColumnIndex = -1
MajorColumnIndex = -1
PrevDegreeCodeIndex = -1
PrevSchoolIndex = -1
ThesisChairIndex = -1
ThesisTitleIndex = -1
CityIndex = -1
ProgramIndex = -1
GPAIndex = -1

num_cells = worksheet.ncols - 1

curr_cell = 0

headerDictionary = {}

while curr_cell < num_cells:
  cell_value = worksheet.cell_value(0, curr_cell)
  headerDictionary[cell_value] = curr_cell
  curr_cell+=1

IDIndex = headerDictionary[IDColumnIdentifier]
LNameIndex = headerDictionary[LNameColumnIdentifier]
FNameIndex = headerDictionary[FNameColumnIdentifier]
CollegeIndex = headerDictionary[CollegeColumnIdentifier]
DipNameIndex = headerDictionary[DipNameColumnIdentifier]
DegreeColumnIndex = headerDictionary[DegreeColumnIdentifier]
MajorColumnIndex = headerDictionary[MajorColumnIdentifier]
PrevDegreeCodeIndex = headerDictionary[PrevDegreeCodeIdentifier]
PrevSchoolIndex = headerDictionary[PrevSchoolIdentifier]
ThesisChairIndex = headerDictionary[ThesisChairIdentifier]
ThesisTitleIndex = headerDictionary[ThesisTitleIdentifier]
ThesisTitleIndex2 = headerDictionary[ThesisTitleIdentifier2]
ThesisTitleIndex3 = headerDictionary[ThesisTitleIdentifier3]
ThesisTitleIndex4 = headerDictionary[ThesisTitleIdentifier4]
CityIndex = headerDictionary[CityIdentifier]
ProgramIndex = headerDictionary[ProgramIdentifier]
if GPAIdentifier in headerDictionary:
  GPAIndex = headerDictionary[GPAIdentifier]

curr_cell = 1
num_cells = worksheet.nrows - 1
listOfRecords = []

while curr_cell < num_cells:
  record = {}

  line = str(worksheet.cell_value(curr_cell, IDIndex)) + "\t"
  #record.append(str(worksheet.cell_value(curr_cell, IDIndex)))
  record["ID"] = str(worksheet.cell_value(curr_cell, IDIndex))

  line += str(worksheet.cell_value(curr_cell, LNameIndex)) + "    \t"
  #record.append(str(worksheet.cell_value(curr_cell, LNameIndex)))
  record["LNAME"] = str(worksheet.cell_value(curr_cell, LNameIndex))

  line += str(worksheet.cell_value(curr_cell, FNameIndex)) + "    \t"
  #record.append(str(worksheet.cell_value(curr_cell, FNameIndex)))
  record["FNAME"] = str(worksheet.cell_value(curr_cell, FNameIndex))

  record["CITY"] = str(worksheet.cell_value(curr_cell, CityIndex))

  college = str(worksheet.cell_value(curr_cell, CollegeIndex))

  if college.find("\v") != -1:
    collegeChoice = college.split("\v")
    college = collegeChoice[0]

  #record.append(college)
  record["COLLEGE"] = college
  line += college + " \t"

  program = str(worksheet.cell_value(curr_cell, ProgramIndex))

  if program.find("\v") != -1:
    programChoice = program.split("\v")
    program = programChoice[len(programChoice)-1]

  #record.append(program)
  record["PROGRAM"] = program
  line += program + " \t"

  dipName = str(worksheet.cell_value(curr_cell, DipNameIndex))

  if dipName.find("\v") != -1:
    dipNameChoice = dipName.split("\v")
    dipName = dipNameChoice[len(dipNameChoice)-1]

  #record.append(dipName)
  record["DIPLOMA_NAME"] = dipName
  line += dipName + " \t"

  #print(dipName)

  degreeName = str(worksheet.cell_value(curr_cell, DegreeColumnIndex))

  if degreeName.find("\v") != -1:
    degreeChoice = degreeName.split("\v")
    degreeName = degreeChoice[len(degreeChoice)-1]


  degreeName = abreviationToString[degreeName]

  #record.append(degreeName)
  record["DEGREE_NAME"] = degreeName
  line += degreeName + " \t"

  major = str(worksheet.cell_value(curr_cell, MajorColumnIndex))

  if major.find("\v") != -1:
    majorChoice = major.split("\v")
    major = majorChoice[len(majorChoice)-1]

  #record.append(major)
  record["MAJOR"] = major
  line += major + "\t"

  #program goes here idk where to get this from

  degreeCode = str(worksheet.cell_value(curr_cell, PrevDegreeCodeIndex))
  prevSchool = str(worksheet.cell_value(curr_cell, PrevSchoolIndex))
  listOfDegreeAndSchool = ""

  degreeCodeChoice = degreeCode.split("\v")
  prevSchoolChoice = prevSchool.split("\v")

  x = 0

  while x < len(degreeCodeChoice):

    degreeCodeEle = degreeCodeChoice[x]

    parrenIndex = degreeCodeEle.find(")")
    degreeCodeEle = degreeCodeEle[int(parrenIndex)+1: len(degreeCodeEle)-1]

    if degreeCodeEle != "AS" and degreeCodeEle != "CRT4":

      if degreeCodeEle == "MED":
        degreeCodeEle = "M.ED"
      else:
        degreeCodeEle = re.sub("(.{1})", "\\1.", degreeCodeEle, 0, re.DOTALL)

      prevSchoolEle = prevSchoolChoice[x]

      parrenIndex = prevSchoolEle.find(")")

      prevSchoolEle = prevSchoolEle[int(parrenIndex)+1: len(prevSchoolEle)-1]

      univRegex = re.compile('Univ\\b')
      prevSchoolEle = univRegex.sub("University", prevSchoolEle)

      prevSchoolEle = prevSchoolEle.replace("Cmty Coll", "Community College")
      #print prevSchoolEle

      degreeAndSchool = degreeCodeEle + ", " + prevSchoolEle
      listOfDegreeAndSchool += degreeAndSchool + ", "

    x += 1

  if listOfDegreeAndSchool == ", , ":
    listOfDegreeAndSchool = ""

  listOfDegreeAndSchool = listOfDegreeAndSchool.strip()

  if listOfDegreeAndSchool[-1:] == ",":
    listOfDegreeAndSchool = listOfDegreeAndSchool[0:len(listOfDegreeAndSchool)-1]

  #record.append(listOfDegreeAndSchool)
  record["DEG_AND_SCHOOL"] = listOfDegreeAndSchool
  line += listOfDegreeAndSchool + " \t"

  thesisChair = str(worksheet.cell_value(curr_cell, ThesisChairIndex))

  if thesisChair.find("\v") != -1:
    chairChoice = thesisChair.split("\v")
    thesisChair = chairChoice[len(chairChoice)-1]

  #record.append(thesisChair)
  record["THESIS_CHAIR"] = thesisChair
  line += thesisChair + "\t"

  thesisTitle = str(worksheet.cell_value(curr_cell, ThesisTitleIndex))
  thesisTitle2 = str(worksheet.cell_value(curr_cell, ThesisTitleIndex2))
  thesisTitle3 = str(worksheet.cell_value(curr_cell, ThesisTitleIndex3))
  thesisTitle4 = str(worksheet.cell_value(curr_cell, ThesisTitleIndex4))

  thesisTitle += " " + thesisTitle2 + " " + thesisTitle3 + " " + thesisTitle4

  #if thesisTitle.find("\v") != -1:
  #  titleChoice = thesisTitle.split("\v")
  #  thesisTitle = titleChoice[len(titleChoice)-1]

  thesisTitle = thesisTitle.replace("\v", " ")
  thesisTitle = thesisTitle.strip()

  if len(thesisChair) !=0 :
    if len(thesisTitle) != 0:
      #print len(thesisTitle)
      if degreeName[0] == 'M':
        thesisTitle = "Thesis: \"" + thesisTitle + "\""
      elif degreeName[0] == 'D':
        thesisTitle = "Dissertation : \"" + thesisTitle + "\""

    record["THESIS_TITLE"] = thesisTitle
    line += thesisTitle + "\t"

  listOfRecords.append(record)

  #get the GPA
  if GPAIndex != -1:
    try:
      gpa = float(worksheet.cell_value(curr_cell, GPAIndex))
    except ValueError:
      gpa = 0
      countOfNonNumberGPAs = countOfNonNumberGPAs + 1

    distinction = ""

    if gpa >= 3.85:
      distinction = "Summa Cum Laude"
    elif gpa >= 3.65 and gpa < 3.85:
      distinction = "Magna Cum Laude"
    elif gpa >= 3.50 and gpa < 3.65:
      distinction = "Cum Laude"

    record["DISTINCTION"] = distinction
    record["GPA"] = gpa
    #print "Distinction: " + distinction


  #print line
  curr_cell+=1

outputBook = xlwt.Workbook()
sheet1 = outputBook.add_sheet('Sheet 1')

sheet1.write(0,0,'ID')
sheet1.write(0,1,'LNAME')
sheet1.write(0,2,'FNAME')
sheet1.write(0,3,'CITY')
sheet1.write(0,4,'COLLEGE')
sheet1.write(0,5,'DIPLOMA_NAME')
sheet1.write(0,6,'DEGREE')
sheet1.write(0,7,'MAJOR')
sheet1.write(0,8,'PROGRAM')
sheet1.write(0,9,'PREVIOUS_SCHOOL')
sheet1.write(0,10,'THESIS_CHAIR')
sheet1.write(0,11,'THESIS_TITLE')

recordCount = 0
currentRow = 1

while recordCount < len(listOfRecords):

  oneRecord = listOfRecords[recordCount]

  sheet1.write(currentRow,0,  oneRecord["ID"])
  sheet1.write(currentRow,1,  oneRecord["LNAME"])
  sheet1.write(currentRow,2,  oneRecord["FNAME"])
  if "CITY" in oneRecord:
    sheet1.write(currentRow,3,  oneRecord["CITY"])
  sheet1.write(currentRow,4,  oneRecord["COLLEGE"])
  sheet1.write(currentRow,5,  oneRecord["DIPLOMA_NAME"])
  #print(oneRecord["DIPLOMA_NAME"])
  if "DEGREE_NAME" in oneRecord:
    sheet1.write(currentRow,6,  oneRecord["DEGREE_NAME"])
  sheet1.write(currentRow,7,  oneRecord["MAJOR"])
  if "PROGRAM" in oneRecord:
    sheet1.write(currentRow,8,  oneRecord["PROGRAM"])
  if "DEG_AND_SCHOOL" in oneRecord:
    sheet1.write(currentRow,9,  oneRecord["DEG_AND_SCHOOL"])
  if "THESIS_CHAIR" in oneRecord:
    sheet1.write(currentRow,10, oneRecord["THESIS_CHAIR"])
  if "THESIS_TITLE" in oneRecord:
    sheet1.write(currentRow,11, oneRecord["THESIS_TITLE"])
  if "DISTINCTION" in oneRecord:
    sheet1.write(currentRow,12, oneRecord["DISTINCTION"])
  if "GPA" in oneRecord:
    sheet1.write(currentRow,13, oneRecord["GPA"])

  currentRow += 1
  recordCount += 1

outputBook.save(str(sys.argv[3]))

print("Done. With " + str(countOfNonNumberGPAs) + " blank GPAs")
#outputBook.save(TemporaryFile())
