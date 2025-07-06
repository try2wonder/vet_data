import pdfplumber
import csv
import re
from openpyxl import Workbook

#############Checking pets and labs. Adding new pets and labs
pets = {}
    #'pet1':['lab1','lab2'],#dict with key=pet_name: value=labs 
    #'pet2':['lab1','lab2']
 

with open('petsAndLabs.txt','r') as file:
    data_string = file.read()
    pets = eval(data_string)
   
print(pets)


pet_name = input("Enter pet's name: ")
lab_name = input("Enter lab's name: ")



def check_pet_lab(pet_name,lab_name):
    if pet_name in pets:
        if lab_name in pets[pet_name]:
            return "pet lab found"
        else:
            pets[pet_name].append(lab_name)
            print(pets)
            return "pet found, lab was added to the list"
    else:
        pets[pet_name] = [lab_name]
        print(pets)
        return "pet not found, added it adn it's lab to the list"

check_pet_lab(pet_name, lab_name)


###########Creating MetaData

with open('petsAndLabs.txt','w', newline='') as file:
    file.write(str(pets))




##################Creating xlsx file with sections
def create_excel_with_sheets():
    wb = Workbook()

    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])
    for pet, labs in pets.items():
        for lab in labs:
            sheet = wb.create_sheet(title=f"{pet}_{lab}")
            #Add headers and sample data and more rows
            sheet.append(["The date"])
            sheet.append(["keys"])

    wb.save("pet_lab_data.xlsx")
    print("Created Excel file with multiple sheets")

choice1 = input("do you want to create a new .xlsx file(1), or use an existing one(2): ")
if int(choice1) == 1:
    create_excel_with_sheets()
elif int(choice1) == 2:
    print("checking existing one")


########Adding Columns to Specific Sheets
wb = load_workbook('pet_lab_data.xlsx')

sheet = wb[f"{pet_name}_{lab_name}"]

last_col = sheet.max_column #Finds the last one with data

new_columns = {
        last_col+1: "New Column 1"
        }

for col_num, header in new_columns.items():
    sheet.cell(row=1, column=col_num, value=header)

for row in range(2, sheet.max_row + 1):
    sheet.cell(row=row, column=last_col+1, value=f"Data {row-1}")

wb.save('pet_lab_data.xlsx')

#########Workign with pdf
pdf_name = "хорси.pdf"
with pdfplumber.open(pdf_name) as pdf:
    page = pdf.pages[0]

    full_text = page.extract_text()
    tables = page.extract_tables()
    #print(type(text))   #string
    #print(type(tables)) #list


