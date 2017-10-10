import xlrd


def column_len(sheet, index):
    col_values = sheet.col_values(index)
    col_len = len(col_values)
    return col_len


def czy_lista(indeks, lista):
    wynik = ""
    col1 = [row[0] for row in lista]
    col2 = [row[1] for row in lista]
    licz = 0
    for ll in col1:
        if ll == indeks:
            wynik = str(col2[licz])
            break
        else:
            wynik = ""
        licz += 1
    return wynik


def wykluczenie(nazwa1, nazwaw):
    wyklucz = ""
    for nw in nazwaw:
        if nw == nazwa1[:len(nw)]:
            wyklucz = "ok"
            break
        else:
            wyklucz = ""
    return wyklucz


def przesuniecie(indeks, lista, w_d):
    if lista[indeks] != "":
        w_d = int(w_d) + 1
        lista.insert(indeks, w_d)
        lista.pop(indeks + 1)
    else:
        w_d = float(w_d) + 0.00001
        lista.insert(indeks, w_d)
        lista.pop(indeks + 1)
    return w_d


def przesun_up(indeks, indeks_up, lista, w_d):
    if lista[indeks] != "":
        w_d = int(w_d) + 1
        lista.insert(indeks_up, w_d)
        lista.pop(indeks + 1)
    else:
        w_d = float(w_d) + 0.00001
        lista.insert(indeks_up, w_d)
        lista.pop(indeks + 1)
    return w_d


def nowy_wybor(w, lista1, lista2):
    if w == "":
        if lista1 == "":
            w = ""
        else:
            if lista1 != lista2:
                w = "ok"
            else:
                w = ""
    return w


print("\nPobieranie zawartości pliku towary.xls ")
towar_n = xlrd.open_workbook("towary.xls")
arkusz_n = towar_n.sheet_by_index(0)
granica_n = column_len(arkusz_n, 0)
baza_nowa = {}
baza_indeks = {}
c_jedn = c_opak = ""
id_n = 0
i = 0
for i in range(granica_n - 1):
    linia1 = int(arkusz_n.cell_value(i + 1, 0))
    linia2 = arkusz_n.cell_value(i + 1, 1)
    linia3 = arkusz_n.cell_value(i + 1, 2)
    linia4 = arkusz_n.cell_value(i + 1, 3)
    linia6 = arkusz_n.cell_value(i + 1, 6)
    linia7 = arkusz_n.cell_value(i + 1, 10)
    linia8 = arkusz_n.cell_value(i + 1, 11)
    linia9 = arkusz_n.cell_value(i + 1, 12)
    linia10 = arkusz_n.cell_value(i + 1, 13)
    if linia8 != "":
        if linia9 != "":
            p_war = '{0:.2f}'.format(linia6 / float(linia9))
            c_jedn = "cena " + p_war + "zł/" + linia8.lower()
        else:
            c_jedn = ""
    else:
        c_jedn = ""
    if linia10 != "":
        p_opak = '{0:.2f}'.format(linia6 * float(linia10))
        c_opak = "cena " + p_opak + "zł/" + str(linia10) + linia7.lower()
    else:
        c_opak = ""
    id_n = linia1
    baza_nowa[id_n] = {"Indeks": linia2, "Identyfikator": linia3, "Asortyment": linia4, "Cena brutto": linia6,
                       "jm": linia7, "Cena jednostki": c_jedn, "Cena opakowania": c_opak}
    baza_indeks[linia2] = linia1
csv_tekst = {}
print("\nPobieranie zawartości pliku towarys.csv ")
csv_file = open("towarys.csv", "r")
id_kar = 0
licznik = 0
s_in = ""
s_id = ""
s_as = ""
cb = ""
cena = ""
for line in csv_file:
    dok_nazwa = line.strip("\n").split("\t")
    if dok_nazwa != "":
        try:
            id_kar = int(dok_nazwa[0])
        except ValueError:
            print("")
        try:
            s_in = dok_nazwa[1]
            s_id = dok_nazwa[2]
            s_as = dok_nazwa[3]
            cb = dok_nazwa[4]
        except IndexError:
            s_in = ""
            s_id = ""
            s_as = ""
            cb = ""
        try:
            cena = float(cb.replace(",", "."))
        except ValueError:
            cena = ""
    csv_tekst[id_kar] = {"S_Indeks": s_in, "S_Identyfikator": s_id,
                         "S_Asortyment": s_as, "S_Cena brutto": cena}
    licznik += 1
csv_file.close()
id_kar = licznik + 10002
print("Pobieranie zawartości pliku cennik.xls")
cennik = xlrd.open_workbook("cennik.xls")
arkusz_c = cennik.sheet_by_index(0)
granica_c = column_len(arkusz_c, 0)
baza_cennik = []
wynik_c = []
for i in range(granica_c - 1):
    liniac = arkusz_c.cell_value(i + 1, 0)
    liniacc = arkusz_c.cell_value(i + 1, 4)
    wynik_c = [baza_indeks[liniac], i+1]
    baza_cennik.append(wynik_c)
    b = wynik_c[0]
    if liniacc != baza_nowa[b]["Cena brutto"]:
        print("\nUWAGA!!! Błędna cena, uaktualnij plik TOWARY.xls")
        input("\n\nAby kontynuować program, naciśnij klawisz Enter.")
dok_tekst = {}
print("\nPobieranie zawartości pliku dokument.txt ")
dok_file = open("dokument.txt", "r")
baza_dok = []
ii = 1
for line in dok_file:
    tab = int(line.find("\t"))
    wynik_d = [baza_indeks[line[:tab]], ii]
    baza_dok.append(wynik_d)
    ii += 1
dok_file.close()
print("\nPobieranie zawartości pliku nazwy00.txt ")
dok_file00 = open("nazwy00.txt", "r")
dok_nazwa = dok_file00.readline().strip("\n").split(";")
dok_grupa = dok_file00.readline().strip("\n").split(";")
dok_file00.close()
if id_n < id_kar:
    granica = id_kar + 1
else:
    granica = id_n + 1
baza_dane = {}
for i in range(10004, granica):
    wybor = ""
    wybor_w = ""
    cena_w = ""
    try:
        wybor = nowy_wybor(wybor, baza_nowa[i]["Cena brutto"], csv_tekst[i]["S_Cena brutto"])
        if csv_tekst[i]["S_Cena brutto"] != "":  #  wybor == "ok" and
            cena_n = float(baza_nowa[i]["Cena brutto"])
            cena_s = float(csv_tekst[i]["S_Cena brutto"])
            if cena_n > cena_s:
                cena_w = "."
            elif cena_n == cena_s:
                cena_w = "="
            elif cena_n < 0.9 * cena_s:
                cena_w = "obniżka -" + str(int(100*(cena_s - cena_n)/cena_s)) + "%"
            elif cena_n >= 0.9 * cena_s and cena_n < cena_s:
                cena_w = "stara cena "  + str(format(cena_s, '0.2f'))    

    except KeyError:
        wybor = ""
    try:
        wybor = nowy_wybor(wybor, baza_nowa[i]["Indeks"], csv_tekst[i]["S_Indeks"])
    except KeyError:
        wybor = ""
    try:
        wybor = nowy_wybor(wybor, baza_nowa[i]["Identyfikator"], csv_tekst[i]["S_Identyfikator"])
    except KeyError:
        wybor = ""
    try:
        if csv_tekst[i]["S_Identyfikator"]:
            a = wybor
    except KeyError:
        try:
            if baza_nowa[i]["Identyfikator"]:
                wybor = "ok"
        except KeyError:
            wybor = ""
    try:
        grupa = baza_nowa[i]["Asortyment"]
        czy_ug = wykluczenie(grupa, dok_grupa)
    except KeyError:
        czy_ug = ""
    try:
        nazwa = baza_nowa[i]["Identyfikator"]
        czy_un = wykluczenie(nazwa, dok_nazwa)
    except KeyError:
        czy_un = ""
    if czy_ug == "ok":
        czy_u = "ok"
    elif czy_un == "ok":
        czy_u = "ok"
    else:
        czy_u = ""
    if wybor == "ok":
        if czy_u == "":
            wybor_w = "ok"
        else:
            wybor_w = ""
    else:
        wybor_w = ""
    c_cennik = czy_lista(i, baza_cennik)
    c_dok = czy_lista(i, baza_dok)
    if csv_tekst.get(i):
        if baza_nowa.get(i):
            baza_dane[i] = (baza_nowa[i]["Identyfikator"], baza_nowa[i]["Indeks"],
                            baza_nowa[i]["Asortyment"], str(baza_nowa[i]["Cena brutto"]).replace(".", ","),
                            baza_nowa[i]["jm"], baza_nowa[i]["Cena jednostki"],
                            baza_nowa[i]["Cena opakowania"], cena_w, i, csv_tekst[i]["S_Indeks"],
                            csv_tekst[i]["S_Identyfikator"], csv_tekst[i]["S_Asortyment"],
                            str(csv_tekst[i]["S_Cena brutto"]).replace(".", ","), "", wybor_w, c_cennik, c_dok)
        else:
            baza_dane[i] = ("", "", "", "", "", "", "", "", i, csv_tekst[i]["S_Indeks"],
                            csv_tekst[i]["S_Identyfikator"], csv_tekst[i]["S_Asortyment"],
                            str(csv_tekst[i]["S_Cena brutto"]).replace(".", ","), "", "", "", "")
    else:
        baza_dane[i] = ("", "", "", "", "", "", "", "", i, "", "", "", "", "", "", "", "")
        if baza_nowa.get(i):
            baza_dane[i] = (baza_nowa[i]["Identyfikator"], baza_nowa[i]["Indeks"],
                            baza_nowa[i]["Asortyment"], str(baza_nowa[i]["Cena brutto"]).replace(".", ","),
                            baza_nowa[i]["jm"], baza_nowa[i]["Cena jednostki"], baza_nowa[i]["Cena opakowania"],
                            "", i, "", "", "", "", "", wybor_w, c_cennik, c_dok)
        else:
            baza_dane[i] = ("", "", "", "", "", "", "", "", i, "", "", "", "", "", "", "", "")
prices_sorted = sorted(zip(baza_dane.values()), reverse=True)
plik = open('dane_towar.txt', 'w')
linia = ""
do_zapisu = ""
wybor_dane1 = wybor_dane2 = wybor_dane3 = 0
naglowek1 = "WYBÓR \tWYBÓRC \tWYBÓRD \tIdentyfikator \tIndeks \tAsortyment \tCena brutto "
naglowek2 = "\tjm \tCena jednostki \tCena opakowania \tPUSTY0 \tIdent. \tS_Indeks "
naglowek3 = "\tS_Identyfikator \tS_Asortyment \tS_Cena brutto \tPUSTY1 \n"
naglowek = naglowek1 + naglowek2 + naglowek3
plik.write(naglowek)
for j in prices_sorted:
    test = list(j[0])
    wybor_dane1 = przesun_up(14, 0, test, wybor_dane1)
    wybor_dane2 = przesun_up(15, 1, test, wybor_dane2)
    wybor_dane3 = przesun_up(16, 2, test, wybor_dane3)
    if type(wybor_dane1) == int or type(wybor_dane2) == int or type(wybor_dane3) == int:
        do_zapisu = '\t'.join(str(d) for d in test) + '\n'
        plik.write(do_zapisu)
plik.close()
plik2 = open('towaryss.csv', 'w')
do_zapisu = ""
nowy = []
for j in prices_sorted:
    test = list(j[0])
    nowy.append(test[8])
    nowy.append(test[1])
    nowy.append(test[0])
    nowy.append(test[2])
    nowy.append(test[3])
    tekst1 = nowy[:]
    nowy.clear()
    do_zapisu = '\t'.join(str(d) for d in tekst1) + '\n'
    plik2.write(do_zapisu)
plik2.close()
