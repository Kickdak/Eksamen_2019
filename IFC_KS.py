import PySimpleGUI as sg
import sys
import ifcopenshell
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

def ifcOrganizer ():
    global fil
    with open(txtfil,"r") as txt:
        fil =txt.read().replace("\n","")

    ifcfile = ifcopenshell.open(fil)  # set til file selection
    print(fil)
    spaces = ifcfile.by_type("IfcSpace")
    sets = ifcfile.by_type("IfcPropertySet")
    liste = []
    romliste = []
    propertyliste = []
    print(spaces)
    for room in spaces:
        romliste.append(room.Name)

    for sett in sets:
        liste.append(sett.Name)


    for sett in sets:
        if sett.Name == "Pset_KS":
            for job in sett.__getitem__(4):
                #print(job.__getitem__(0))
                propertyliste.append(job.__getitem__(0))


    mylist = list(dict.fromkeys(propertyliste))
    mylist.sort()
    print(mylist)
    if mylist == []:
        print("hey")
        sg.PopupError('Pset finnes ikke i modell.\nOpprett parametere og kjør programmet på nytt.')
        os.startfile("Pset_Creater.exe")
        sys.exit()
    print(romliste)
    #print(room.Name +' - '+room.LongName)print(job.__getitem__(0))
    #UI starter
    layout = [[sg.Text('Rom nummer: ', size=(10, 1)),sg.InputCombo(romliste),sg.Text('Property Value: -->', size=(14, 1)),sg.InputCombo(mylist),sg.Checkbox('Utført'),sg.Checkbox('Avvik'),sg.Submit("Kjør"), sg.Cancel("Exit")]]
    window = sg.Window('IFC Property Updater')

    #Loop for programmet og knapper
    while True:  # Event Loop<
        event, values = window.Layout(layout).Read()
        if event in (None, 'Exit'):
            sys.exit("aa! errors!")
        if event == 'Kjør':
            if values[2] == False and values[3] == False:
                values[2] = "Ikke klart med Avvik"
            if values[2] == True and values[3] == True:
                values[2] = "Ferdig med Avvik"
            if values[2] == False:
                values[2] = "Ikke klart"
            if values[2] == True:
                values[2] = "Ferdig"

            oppdater_job(values[1], values[0], values[2], ifcfile, sets, fil)


def oppdater_job(yrkesgruppe, romnummer, ferdig, ifc, sets, fil):
    
    #sets = ifc.by_type("IfcPropertySet")
    setlist = []
    #testtall = romnummer
    for sett in sets:
        #ifc.by_type("Pset_isFinished")

        if sett.Name == "Pset_KS" and sett.__getitem__(3) == romnummer:
            print(sett)
            for job in sett.__getitem__(4):
                if job.Name == yrkesgruppe:
                    print(job)
                    job.NominalValue.__setitem__(0, ferdig)
                    print(job.NominalValue)
    ifc.write(fil)
if __name__ == '__main__':
    ifcOrganizer()
