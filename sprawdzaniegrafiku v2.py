import xlrd


def zapisz(tekst):
    sfile = "spr.txt"
    file1 = open(sfile, "w")  # otwieramy plik do zapisu, istniejący plik zostanie nadpisany(!)
    file1.write(tekst)
    file1.close()  # zamykamy plik


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
    ilednimamiesiac = 31
    nr = 0
    for lp in range(5, zasieg, 2):
        listazmian = []
        nazwiskoimie = arkusz.row_values(lp)[1].upper()
        if nazwiskoimie == "PUSTE":
            continue
        pomoc = 0
        for zmiana in range(3, ilednimamiesiac + 3, 1):
            numerekzmiany = arkusz.row_values(lp)[zmiana]
            numerekzmiany = str(numerekzmiany).strip().upper()
            if numerekzmiany == "UN" or numerekzmiany == "M" or numerekzmiany == "OP" or numerekzmiany == "UT" or numerekzmiany == "UO":
                numerekzmiany = arkusz.row_values(lp + 1)[zmiana]
                numerekzmiany = str(numerekzmiany).strip().upper()
            if numerekzmiany == "W":
                listazmian.append(0)
                pomoc += 1
            elif numerekzmiany == "UW":
                listazmian.append(4)
            else:
                try:
                    if numerekzmiany == "I":
                        listazmian.append(1.0)
                    elif numerekzmiany == "II":
                        listazmian.append(2.0)
                    elif numerekzmiany == "III":
                        listazmian.append(3.0)
                    else:
                        listazmian.append(float(numerekzmiany))
                except ValueError:
                    # print(nazwaarkusza + " " + nazwiskoimie + "Error !!!!!!!!!!!!!!!!!!!!")
                    pass
        slownikzmian[nazwiskoimie] = listazmian
        nr += 1
    return slownikzmian


def czymaszybkieprzejscie(lista, dnipierwszy):
    czyma = False
    przelom = False
    nocprzedurlopem = False
    for i in range(1, len(lista)):
        if lista[i - 1] > lista[i]:
            if lista[i - 1] != 4 and lista[i] != 0:
                czyma = True
                if i == dnipierwszy:
                    przelom = True
        if lista[i - 1] == 3 and lista[i] == 4:
            nocprzedurlopem = True
    return czyma, przelom, nocprzedurlopem


def najwiecejdniwciagu(lista):
    najwiecej = 0
    pomoc = 0
    for i in lista:
        if i != 0:
            pomoc += 1
        else:
            if pomoc > najwiecej:
                najwiecej = pomoc
            pomoc = 0
    if pomoc > najwiecej:
        najwiecej = pomoc
    return najwiecej


def niedziela(nazwaarkusza, grafik):
    arkusz = grafik.sheet_by_name(nazwaarkusza)
    for dni in range(3, 34, 1):
        dzientygodnia = arkusz.row_values(3)[dni]
        if dzientygodnia == "N":
            return dni - 2


def ciagniedziel(lista, pierwszaniedziela):
    najwiecej = 0
    pomoc = 0
    dzien = 1
    for i in lista:
        if (dzien - pierwszaniedziela) % 7 == 0:
            if i != 0:
                pomoc += 1
            else:
                if pomoc > najwiecej:
                    najwiecej = pomoc
                pomoc = 0
        dzien += 1
    if pomoc > najwiecej:
        najwiecej = pomoc
    return najwiecej


# miesiac1 = "Czerwiec2019.xls"
miesiac1 = "Lipiec2019.xls"
miesiac2 = "Sierpien2019.xls"

grafik1 = xlrd.open_workbook(miesiac1)
grafik2 = xlrd.open_workbook(miesiac2)

niedziela = niedziela("Kierownictwo", grafik1)

listaarkuszy = ["Kierownictwo", "Operatorzy", "Magazynierzy", "Rozbieracze", "Mieszana", "Wkładacze", "Wagowi",
                "Stażyści", "Nowi"]

pierwszy = {}
drugi = {}
for arkusz in listaarkuszy:
    pomoc1 = odczytajarkusz(arkusz, grafik1)
    pomoc2 = odczytajarkusz(arkusz, grafik2)
    for i in pomoc1:
        pierwszy[i] = pomoc1[i]
    for j in pomoc2:
        drugi[j] = pomoc2[j]
dnipierwszy = len(pierwszy['BUĆKO TOMASZ'])

for j in pierwszy:
    if j not in drugi:
        print(j + " nie ma grafiku na nowy miesiąc")

for j in drugi:
    if j in pierwszy:
        pierwszy[j].extend(drugi[j])
    else:
        lista = []
        for lp in range(0, dnipierwszy):
            lista.append(0)
        pierwszy[j] = lista
        pierwszy[j].extend(drugi[j])
        print(j + " nie pracował/ła w poprzednim miesiącu")

tekst = ""
for i in pierwszy:
    test, przelom, nocprzedurlopem = czymaszybkieprzejscie(pierwszy[i], dnipierwszy)

    if test:
        tekst += i + " ma szybkie przejscie"
        if przelom:
            tekst += " na przełomie miesiąca"
        tekst += ".\n"
    if nocprzedurlopem:
        tekst += i + " ma noc przed urlopem.\n"
    test = najwiecejdniwciagu(pierwszy[i])
    if test > 6:
        tekst += i + " ma " + str(test) + " dni w ciągu do pracy.\n"

    test = ciagniedziel(pierwszy[i], niedziela)
    if test > 3:
        niedziele = "niedziele"
        if test > 4:
            niedziele = "niedziel"

        tekst += i + " pracuje " + str(test) + " " + niedziele + " bez przerwy.\n"
zapisz(tekst)
