# Script devloped by Alfredo Ramirez Izaguirre EDA Engineer Intern, June 2019
# Things to do ✓✓✓:
#     (  ) Refactor main function into a couple smaller functions
#     (  ) Get ITAR Access to search through those files. Give option to user no narrow down search?
#     (  ) Test Linux version once modules are correctly installed
#     ( ✓ ) Added Simple error handling
#     ( ✓ ) Open log file after program is done running 

from xml.etree import ElementTree
import xlrd
import os

def ID_Index(sheetA):
    listColumns = []

    for columns in range(sheetA.ncols):
        listColumns.append(sheetA.cell_value(1,columns))

    Column_Index = listColumns.index("DIE_ID")

    return(Column_Index)

def CKT_Index(sheetA):
    listColumns = []
    
    for columns in range(sheetA.ncols):
        listColumns.append(sheetA.cell_value(1,columns))

    Column_Index = listColumns.index("CKT Name")

    return(Column_Index)

def XML_Finder(file_A, EG_Code):
    folderName = EG_Code
    newDir = '' 
    
    for root, dirs, files in os.walk('//x/x/x/x/x'):
        for directory in dirs:
            if directory.endswith(folderName):
                newDir = (os.path.join(root, directory))
        break

    if newDir == '':
        input("No XML file Found! Hit Enter to close")
        raise SystemExit()

    for root, dirs, files in os.walk(newDir):
        for file in files :
            if file.endswith('Assignment.xml'):
                xmlFile = os.path.join(root, file)
    
    return(xmlFile)

def Input_Workbook_Finder(EG_Code):
    file_A = ''
    newDir = ''

    for root, dirs, files in os.walk('//x/x/x/x/x/x/x'):
        for directory in dirs:
            if EG_Code in directory:
                newDir = (os.path.join(root, directory))
        break
    
    if newDir == '':
        input("No Workbook file Found! Hit Enter to close")
        raise SystemExit()
        
    for root, dirs, files in os.walk(newDir):
        for file in files :
            if 'INPUT' in file:
                file_A = os.path.join(root, file)
            elif 'input' in file:
                file_A = os.path.join(root, file)
            elif 'Input' in file:
                file_A = os.path.join(root, file)

    newDir = newDir.replace("/", "\\") 
    
    return(file_A, newDir)

def Custom_input():
    EG_Number = input("Enter Input Workbook EG number: ")
    EG_Code = ''

    if 'EG' not in EG_Number:
        EG_Code = 'EG' + EG_Number
    else:
        EG_Code = EG_Number

    return(EG_Code)

def Text_File(file_A_Dir):
    Log_File = open(file_A_Dir + "/compare_log.txt","w+")
    return(Log_File)
    

EG_Code = Custom_input()

file_A = Input_Workbook_Finder(EG_Code)[0]
file_B = XML_Finder(file_A, EG_Code)

file_A_Dir = Input_Workbook_Finder(EG_Code)[1]
f = Text_File(file_A_Dir)

workbookA = xlrd.open_workbook(file_A)
sheetA = workbookA.sheet_by_index(1)

Index_ID = ID_Index(sheetA)
#Index_CKT = CKT_Index(sheetA)


listA = []
listB = []
listC = []
listD = []
counter = 0

#Filling up listB: cell values of XML
tree = ElementTree.parse(file_B)   
DIE_ID = tree.findall("//DIE_ID")
CKT_NAME = tree.findall("//CKT_NAME")
listB = [t.text for t in DIE_ID]
listD = [a.text for a in CKT_NAME]

#Filling up listC: CKT names from xls
for rowsC in range(2,sheetA.nrows):
    listC.append(sheetA.cell_value(rowsC,0))  

#Filling up listA: DIE_ID values from xls
for rowsA in range(2,sheetA.nrows):
    listA.append(sheetA.cell_value(rowsA,Index_ID))

#Show path of files being used
print("\nInput Workbook: %s \n" % (file_A))
f.write("\nInput Workbook: %s \n" % (file_A))
print("XML File: %s \n" % (file_B))
f.write("XML File: %s \n" % (file_B))
print("Log file: %s \n" % (file_A_Dir))
f.write("Log file: %s \n" % (file_A_Dir))

print("\n{1:<20}{0:^20}{2:>20}\n".format('XML', 'Message', 'Workbook'))
f.write("\n{1:<20}{0:^20}{2:>20}\n".format('XML', 'Message', 'Workbook'))

#Comparing lists 
for i in range(len(listB)):
    if listA[i] != listB[i] and listA[i] != '':
        if listB[i] == None:
            listB[i] = 'None'
        print("\n{1:<20}{0:^20}{2:>20}\n".format(listB[i],"Mistmatch at %s:" %(listC[i]) , listA[i]))
        f.write("\n{1:<20}{0:^20}{2:>20}\n".format(listB[i],"Mistmatch at %s:" %(listC[i]) , listA[i]))    
        counter += 1
    elif listA[i] == '' and listB[i] != None:
        print("\n{1:<20}{0:^20}\n".format(listB[i],"Mistmatch at %s:" %(listD[i])))
        f.write("\n{1:<20}{0:^20}\n".format(listB[i],"Mistmatch at %s:" %(listD[i])))    
        counter += 1
if counter == 0:
    print("No mismatches")
    f.write("No mismatches") 

f.close()
input("\nHit Enter to close and open log...")
os.startfile(file_A_Dir + '\compare_log.txt')
