import pdfplumber
import csv
import re
from openpyxl import Workbook, load_workbook

#############Checking pets and labs. Adding new pets and labs
pets = {
        "pet1": "lab1"
        }
pet_name = str()
lab_name = str()


with open('petsAndLabs.txt','r') as file:
    data_string = file.read()
    pets = eval(data_string)
   
print(pet_name, lab_name)
print(pets)


def pet_choice_dialog():
    
    pet_name = input("Enter pet's name: ")
    lab_name = input("Enter lab's name: ")
    


pet_choice_dialog()

print(pet_name)
print(pets)

def check_pet_lab(pet_name,lab_name):
    if pet_name in pets.keys():
        if lab_name in pets[pet_name]:
            return "pet lab found" #choosing the right Sheet
        else:
            choi=input("Pet is found, but not the lab. Do you want to add another pet_lab combination?(y/n): ")
            if choi == "y":
                pets[pet_name].append(lab_name)
                print(pets)
            else:
                print(pets)
                choi = input("Do you want to return to pet_lab choice?(y/n): ")
                if choi == "y":
                    pet_choice_dialog()
                    check_pet_lab(pet_name,lab_name)
    else:
        choi = input("no pet_lab combination found. Do you want to create a new one?(y/n): ")
        if choi == "y":
            pets[pet_name] = [lab_name]
        else:
            choi = input("Do you want to return to pet_lab choice?(y/n): ")
            if choi == "y":
                pet_choice_dialog()
                check_pet_lab(pet_name,lab_name)
                 
            print(pets)

check_pet_lab(pet_name, lab_name)


###########Creating MetaData

with open('petsAndLabs.txt','w', newline='') as file:
    file.write(str(pets))

#########Workign with pdf
pdf_name = "хорси.pdf"
temp_list_var = []
def work_with_pdf(name):
    with pdfplumber.open(name) as pdf:
        page = pdf.pages[0]
    
        full_text = page.extract_text()
        tables = page.extract_tables()
        #print(type(text))   #string
        #print(type(tables)) #list
        return tables
temp_list_var = work_with_pdf(pdf_name)[0]
clean_list = temp_list_var[1:]
clean_dict = {testd[0] : testd[2] for testd in clean_list}
print(clean_dict)




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

    wb.save("pet_lab_data.xlsx")
    print("Created Excel file with multiple sheets")

choice1 = input("do you want to create a new .xlsx file(1), or use an existing one(2): ")
if int(choice1) == 1:
    create_excel_with_sheets()
elif int(choice1) == 2:
    print("checking existing one")


########Adding Columns to Specific Sheets
def add_data_column(pet_name,lab_name):
    wb = load_workbook('pet_lab_data.xlsx')

    sheet = wb[f"{pet_name}_{lab_name}"]

    last_col = sheet.max_column #Finds the last one with data. Need to check if the date already exists there

    new_columns = {
            last_col+1: "New Column 1" #This is supposed to be the date of analysis. mby make a temp var for column number, in case its not the last one
            }

    for col_num, header in new_columns.items():
        sheet.cell(row=1, column=col_num, value=header)

    for row in range(2, sheet.max_row + 1):
        sheet.cell(row=row, column=last_col+1, value=f"Data {row-1}")#Data row 1 is the values i need to get from PDF

    wb.save('pet_lab_data.xlsx')

add_data_column(pet_name,lab_name)


