from openpyxl import Workbook, load_workbook


test_list = ['an1','an2','an3','an4']
test_val_list = ['val1','val2','val3','val4']

wb = load_workbook('test.xlsx')

sheet = wb['Sheet1']

temp_val_list = []
temp2 = []

for cell in sheet['A']:
    temp_val_list.append(cell.value)

print(temp_val_list)

if temp_val_list[1:] == test_list:
    print('yeeeeah')

for cell in sheet['B']:
    temp2.append(cell.value)
print(temp2)

if temp2[1:] == test_val_list:
    print('fuc yeah!')


#print(sheet['A1'].value)
