import PySimpleGUI as sg
from openpyxl import load_workbook

# set file path
filepath = "KS_punkter.xlsx"
# load demo.xlsx
wb = load_workbook(filepath)
# select demo.xlsx
sheet = wb.active
# get max row count
max_row = sheet.max_row
# get max column count
#print (max_row)
max_column = sheet.max_column
#print (max_column)
# iterate over all cells
# iterate over all rows
fagfelt=[]
for f in range(2, max_column+1):
    #print (f)
    fag = sheet.cell(row=1,column=f)
    fagfelt.append(fag.value)

layout = [[sg.Text('Fagfelt: ', size=(10, 1)),sg.InputCombo(fagfelt, change_submits=True),sg.Submit("Go"), sg.Cancel("Exit")]]
buildingElement = sg.Window('IFC Property Updater')

#Loop for programmet og knapper
while True:  # Event Loop<
    event, values = buildingElement.Layout(layout).Read()
    if event in (None, 'Exit'):
        print("Skriptet ble stoppet")
        sys.exit()
    if event == "Go":
        break
    print(values[0])

nedtrekk = values[0]







klar =[]
sjekkliste = []
index =2#Rad
fagnr = fagfelt.index(nedtrekk)
print (fagnr)


for i in range(index, max_row + 1):

    cell_obj = sheet.cell(row=i, column=int(fagnr+2))
    if not cell_obj.value == None:
        sjekk= cell_obj.value
        klar.append(sjekk)

print (fagfelt)
print (klar)
