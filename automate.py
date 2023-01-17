import xml.etree.ElementTree as ET
import openpyxl
from os.path import join
# Read an Excel file
str=input("Enter name of the excel file to read data from (also write .xlsx): ")
try:
 workbook = openpyxl.load_workbook(str)
 worksheet = workbook.active
except Exception as e:
     print("Error: ", e)
     print("Please enter a valid Excel filename")
     
str1=input("Give absolute path of the base .ptpx file for modifications: ")
str2=input("Give absolute path of folder where you want to save all .ptpx files: ")

print("")
print("Creating files...")
print("")

# Print the data
for row in worksheet.iter_rows():
    name=row[0].value
    
    try:
     tree = ET.parse(str1)
    except Exception as e:
     print("Error: ", e)
     print("Please enter a valid .ptpx file path")
     break
   
    root = tree.getroot()
    jobProp=root.find("./JOB_PROPERTIES/JOB_NAME")
    mergeDes=root.find("./PRINTING/MERGE_DISC_DESC")
    mergeTit=root.find("./PRINTING/MERGE_DISC_TITLE")
    files = root.findall('./RECORDING/SOURCE/FILES/FILE')
   
    jobProp.text="C:\\Users\\james\\OneDrive\\Desktop\\New folder\\{}.ptpx".format(name)
    mergeDes.text=name
    mergeTit.text=name

    count=0

    for file in files:
        if count==1:  
          file.text=row[1].value
        count+=1
   
  
    finalstr=join(str2,"{}.ptpx".format(name))
    try:
     tree.write(finalstr)
    except Exception as e:
     print("Error: ", e)
     print("Please enter a valid output folder path")
     break


