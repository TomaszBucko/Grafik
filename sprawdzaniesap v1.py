import os.path
import xlrd
import xlwt

pliksap = "a.txt"
miesiac = "Czerwiec2019.xls"
dodruku = "dodruku.txt"
def odczytajsap():
    sfile = pliksap  # nazwa pliku z estymecjami
    slownik = {}
    licznik = 1
    urlopy = 0
    if os.path.isfile(sfile):  # czy istnieje plik słownika?
        with open(sfile, "r") as sTxt:  # otwórz plik do odczytu
            for line in sTxt:  # przeglądamy kolejne linie
                skok = 1
                lista = []
                line = line.replace("\n", "")
                t = line.split(";")  # rozbijamy linię
                nazwiskoimie = str(t[1]).upper()
                if len(t[3]) > 21:
                    pomocnicza = t[3].split(",", skok)
                    t.remove(t[3])
                    t.insert(3, pomocnicza[0])
                    t.insert(4, pomocnicza[1])
                pracowal = 1
                for i in range(4, len(t), 1):
                    if t[i] == "000000,000000":
                        #print("Wolne")
                        lista.append(0)
                    elif t[i] == "060000,140000" or t[i] == "070000,150000":
                        #print("Pierwsza zamiana")
                        lista.append(1)
                    elif t[i] == "140000,220000":
                        #print("Druga zamiana")
                        lista.append(2)
                    elif t[i] == "220000,060000":
                        #print("Trzecia zamiana")
                        lista.append(3)
                    elif t[i] == "Urloprodzicielski":
                        #print("Urlop rodzicielski")
                        lista.append(5)
                    elif t[i] == "Urlopmacierzyński":
                        #print("Urlop macierzyński")
                        lista.append(5)
                    elif t[i] == "Rehabilitacja" or t[i] == "Chorobazwykła" or t[i] == "Opiekanaddzieckiem<14" or t[i] == "Wypadekwpracy/chor.zaw.":
                        #print("Urlop macierzyński")
                        lista.append(5)
                    elif t[i] == "Urlopwychowawczy0-4":
                        #print("Urlop macierzyński")
                        lista.append(5)
                    elif t[i] == "ERROR":
                        #nic nie robimy
                        pracowal = 0
                        #print("Nie pracuje u nas")
                    elif t[i] == "Urlopwypoczynkowy" or t[i] == "Urlopwypocz.nażądanie":
                        #print("Urlop Wypoczynkowy")
                        lista.append(4)
                        urlopy += 1
                    elif t[i] == "Urlopokolicznościowy":
                        lista.append(4)
                    elif len(t[i]) < 287:
                        if len(t[i]) != 0:
                            print(len(t[i]))
                            print("Brakuje zapisu: " + t[i])
                if pracowal == 1:
                    slownik[nazwiskoimie] = lista
                    licznik += 1
    else:
        print("Nie ma pliku z sap z grafikiem:" + pliksap)
    return slownik

def ileludzijestwarkuszu(arkusz):
    len = 78
    start = arkusz.row_values(5)[0]
    koniec = 0
    for lp in range(5, len, 2):
        if type(arkusz.row_values(lp)[0]) == float:
            koniec = arkusz.row_values(lp)[0]
        else:
            break
    liczba = koniec - start + 1
    return liczba

def odczytajarkusz(nazwaarkusza, grafik):
    arkusz = grafik.sheet_by_name(nazwaarkusza)
    slownikzmian = {}
    ileludzi = int(ileludzijestwarkuszu(arkusz))
    zasieg = int(ileludzi * 2 + 4)
    ilednimamiesiac = 30
    nr = 0
    for lp in range(5, zasieg, 2):
        listazmian = []
        nazwiskoimie = arkusz.row_values(lp)[1].upper()
        imieinazwisko = nazwiskoimie.split(" ")

        if nazwiskoimie == "PUSTE" or nazwiskoimie == "":
            continue
        nazwiskoimie = imieinazwisko[0] + imieinazwisko[1]
        pomoc = 0
        for zmiana in range(3, ilednimamiesiac + 3, 1):
            numerekzmiany = arkusz.row_values(lp)[zmiana]
            numerekzmiany = str(numerekzmiany).strip().upper()
            if numerekzmiany == "W" or numerekzmiany == "M":
                listazmian.append(int(0))
                pomoc += 1
            elif numerekzmiany == "UW" or numerekzmiany == "UT" or numerekzmiany == "UO":
                listazmian.append(int(4))
            elif numerekzmiany == "UN":
                listazmian.append(int(5))
            else:
                try:
                    if numerekzmiany == "I" or numerekzmiany == "1.0":
                        listazmian.append(int(1))
                    elif numerekzmiany == "II" or numerekzmiany == "2.0":
                        listazmian.append(int(2))
                    elif numerekzmiany == "III" or numerekzmiany == "3.0":
                        listazmian.append(int(3))
                    else:
                        #listazmian.append(numerekzmiany)
                        a = 1
                except ValueError:
                    #print(nazwaarkusza + " " + nazwiskoimie + "Error !!!!!!!!!!!!!!!!!!!!")
                    pass
        slownikzmian[nazwiskoimie] = listazmian
        nr += 1
    return slownikzmian


slownik = odczytajsap()

sfile = dodruku
file1 = open(sfile, "w", encoding="utf-8")


grafik = xlrd.open_workbook(miesiac)
listaarkuszy = ["Kierownictwo", "Operatorzy", "Magazynierzy", "Rozbieracze", "Mieszana", "Wkładacze", "Wagowi",  "Stażyści", "Nowi"]

exel = {}
for arkusz in listaarkuszy:
    pomoc = odczytajarkusz(arkusz, grafik)
    for i in pomoc:
        exel[i] = pomoc[i]

nazwiska = 35
raport = "raport.txt"
rfile = raport
file2 = open(rfile, "w", encoding="utf-8")
nr = 0
for i in exel:
    test = []
    roznice = False
    nr += 1
    if i != "":
        file1.write(("Sap    - " + i).ljust(nazwiska))
        if i in slownik:
            if nr - 11 > 0:
                tekst = str(nr-11) + "."
            else:
                tekst = " "
            tekst += i.ljust(nazwiska) + " - "
            tekst += "Różnice sap/exel, dzień: "
            for j in range(0, len(exel[i]), 1):
                if slownik[i][j] == exel[i][j]:
                    test.append(" ")
                else:
                    test.append("-")
                    tekst += str(j + 1) + ", "
                    roznice = True
            tekst += "\n\n"
            file1.write(str(slownik[i]))
        else:
            file1.write("Brak Nazwiska w pliku Sap")

        file1.write("\n")
        file1.write(("Grafik - " + i).ljust(nazwiska))
        file1.write(str(exel[i]))
        file1.write("\n         Różnice:".ljust(nazwiska+1))
        linia = "["
        for l in range(0 , len(test), 1):
            linia += test[l]
            if l != len(test) - 1:
                linia += "  "
        linia += "]"
        file1.write(linia)
        file1.write("\n")
        file1.write("\n")
        if roznice:
            file2.write(tekst)

