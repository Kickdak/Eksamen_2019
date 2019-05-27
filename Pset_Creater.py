import PySimpleGUI as sg
import sys
import ifcopenshell
import uuid
import os

txtfil = "filbane.txt"

def filbane():




        window_rows = [[sg.Text('Velg IFC Lokasjon')],
                       [sg.InputText(), sg.FileBrowse()],
                       [sg.Submit("Kjør"), sg.Cancel("Exit")]]
    #
        window = sg.Window('IFC')

        # Loop for programmet og knapper
        while True:  # Event Loop<
            event, values = window.Layout(window_rows).Read()
            if event in (None, 'Exit'):
                sys.exit("aa! errors!")
            if event == 'Kjør':
                window.Close()
                break
        global source_filename
        source_filename = values[0]
        with open(txtfil,"w") as writefile:
            writefile.write(source_filename+"\n")

if not os.path.exists(txtfil):
    filbane()

with open(txtfil,"r") as txt:
    global fil
    fil =txt.read().replace("\n","")
ifcfile = ifcopenshell.open(fil) #set til file selection
rooms = ifcfile.by_type("IfcSpace")
create_guid = lambda: ifcopenshell.guid.compress(uuid.uuid1().hex)
owner_history = ifcfile.by_type("IfcOwnerHistory")[0]
spaces = ifcfile.by_type("IfcSpace")
sets= ifcfile.by_type("IfcPropertySet")
list = []
for sett in sets:
    list.append(sett.Name)
#print (list)
myset = set(list)
#print (myset)



# Properties
def create_pset():
    file = "Parameter_template.txt"
    global antall_linjer
    antall_linjer = 0
    global index
    index = 0
    #åpner fil
    with open(file, "r+") as f:
        data = f.readlines()
    print (data)
    #teller antall linjer i txt
    for parametere in data:
        antall_linjer += len(parametere[0])



    if "Pset_KS" not in myset:
        print("Pset lages")

        property_values = []

        print(property_values)
        antall_properties = len(property_values)
        print (antall_properties)
        for space in spaces:
            while index < antall_linjer:
                property_values.append(ifcfile.createIfcPropertySingleValue(data[index][:-1], data[index][:-1],
                                                                            ifcfile.create_entity("IfcText", "Ikke klart"),
                                                                            None), )

                index += 1
            #print(space.Name)
            property_set = ifcfile.createIfcPropertySet(create_guid(), owner_history, "Pset_KS", str(space.Name), property_values)
            ifcfile.createIfcRelDefinesByProperties(create_guid(), owner_history, None, None, [space], property_set)
            #print (index)
            #print (antall_linjer)
            ifcfile.write(fil)

            property_values =[]
            index = 0
            fil

            #print (fil)


    else:
        print("Pset finnes")
        sg.PopupError("Pset er allerede oprettet i modell.")


def ferdig_gui():
    layout = [[sg.Text("Pset_KS ferdig.\nStart IFC_KS programmet?",justification='center')],
            [sg.Submit("Start"),sg.Text("                   "), sg.Cancel("Lukk")]]
    window = sg.Window("Pset_KS",size=(230,90))
    while True:
        event, values = window.Layout(layout).Read()
        if event=="Start":
            os.startfile("IFC_KS.exe")
            sys.exit()

        if event=="Lukk":
            sys.exit("Program lukkes")

create_pset()
#ferdig_gui()

#180=IFCPROPERTYSINGLEVALUE('MAL fug',$,IFCBOOLEAN(.F.),$);
#5621=IFCPROPERTYSINGLEVALUE('Elektriker sterkstrom\X4\0000000A\X0\','Elektriker sterkstrom\X4\0000000A\X0\',IFCBOOLEAN(.F.),$);