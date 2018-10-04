from openpyxl import Workbook
from openpyxl import load_workbook

wbOpsum = load_workbook(filename = "Opsummering.xlsx")
wsSamlet = wbOpsum.active

def opstart():
    nameCell = ""
    ORCell = ""
    for x in range(70):
        #adding formular and link to nameCells
        nameCell = "B" + str(x + 6)
        nametxt = "='Template (" + str(x + 1) + ")'!B2"
        wsSamlet[nameCell] = nametxt
        #adding formular to the Orginator cells
        ORCell = "C" + str(x + 6)
        ORtxt = "='Template (" + str(x + 1) + ")'!B5"
        wsSamlet[ORCell] = ORtxt
        #adding formular to coordinator cells
        KOCell = "D" + str(x + 6)
        KOtxt = "=HVIS(IKKE('Template (" + str(x + 1) + ")'!C5='None');'Template (" + str(x + 1) + ")'!C5;"")"
        #=HVIS(IKKE('Template (1)'!C5="None");'Template (1)'!C5;"")
        print(KOtxt)
        wsSamlet[KOCell] = KOtxt
    wbOpsum.save("Opsummering2.xlsx")
