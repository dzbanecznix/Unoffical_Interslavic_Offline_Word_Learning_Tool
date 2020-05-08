import msvcrt, random, os, xlrd, urllib, gspread, json
from time import sleep
from urllib.request import urlopen
from oauth2client.service_account import ServiceAccountCredentials
clear = lambda: os.system('cls')
lines = [1, 5, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20]
langs = ["medžuslovjansky", "english", "русский", "Беларускa","українськa","polski","čeština","slovački","Бугарски","Македонски","српски","hrvatski","slovenski","staroslovjansky","deutsch"]
ll = []
for lang in langs:
    ll.append(len(lang))
l = max(ll)
l = (l//2)*2+1
p1, p2 = [0, 1, 0], 0
l3 = len("NEWS&UPDATES")
version, localisation = "", ""
def OpenGoogleSheet(filename, jsonpath):
    scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(jsonpath, scope)
    client = gspread.authorize(creds)
    sheet = client.open(filename).sheet1  # Open the spreadhseet  # Get a list of all records
    return(sheet)
def doths():
    global version, localisation
    version = open("C:\\InterslavicDictionary\\version.txt", "r", encoding = "utf-8").read()
    localisation = open("C:\\InterslavicDictionary\\localisation.txt", "r", encoding = "utf-8").read()
    names0, nums0 = ["Evrething"], [-1]
    themes, tns = getThemes(open("C:\\InterslavicDictionary\\themes(en).txt", "r", encoding = "utf-8").readlines(), open("C:\\InterslavicDictionary\\themes(isv).txt", "r", encoding = "utf-8").readlines(), open("C:\\InterslavicDictionary\\medžuslovjansky.txt", "r", encoding = "utf-8").readlines(), open("C:\\InterslavicDictionary\\english.txt", "r", encoding = "utf-8").readlines(), open("C:\\InterslavicDictionary\\things_that_cant_be_in_themes.txt", "r", encoding = "utf-8").readlines())
    themes, tns = names0 + themes, nums0 + tns
    ll2 = []
    for theme in themes:
        ll2.append(len(theme))
    l2 = max(ll2)
    l2 = (l2//2)*2+1
    return(themes, tns, l2)
def getnum(what2, what1, where1, where2, wherecantitbe, whatcantitbe):
    i = 0
    while i < len(where1):
        thing1, thing2 = where1[i], where2[i]
        if what1 in thing1 and what2 in thing2 and not i in wherecantitbe:
            can = True
            for thing in whatcantitbe:
                if thing in thing2 or thing2 in thing:
                    can = False
            if can:
                return(i)
        i+=1
    return(-1)
def getThemes(f1, f2, lang1, lang2, f3):
    themes = []        
    isname = False
    for i in range(len(f1)-2):
        line1 = f1[i]
        line2 = f2[i]
        line3 = f3[i]
        isname = not(isname)
        if isname:
            themes.append([])
            themes[-1].append(line1[:-1])
        else:
            strs1, strs2, wherecantitbe = [], [], []
            s = ""
            line1 += ","
            line2 += ","
            line3 += ","
            for zn in line1:
                if zn == ",":
                    if s[0] == " ":
                        s = s[1:]
                    if s[-1] == "\n":
                        s = s[:-1]
                    if s[-1] == " ":
                        s = s[:-1]
                    strs1.append(s)
                    s = ""
                else:
                    s += zn
            s = ""
            for zn in line2:
                if zn == ",":
                    if s[0] == " ":
                        s = s[1:]
                    if s[-1] == "\n":
                        s = s[:-1]
                    if s[-1] == " ":
                        s = s[:-1]
                    strs2.append(s)
                    s = ""
                else:
                    s += zn
            s = ""
            for zn in line3:
                if zn == ",":
                    if s[0] == " ":
                        s = s[1:]
                    if s[-1] == "\n":
                        s = s[:-1]
                    if s[-1] == " ":
                        s = s[:-1]
                    wherecantitbe.append(s)
                    s = ""
                else:
                    s += zn
            for i in range(len(strs1)):
                s1, s2 = strs1[i], strs2[i]
                n = getnum(s1, s2,lang1, lang2, themes[-1], wherecantitbe)
                if n != -1:
                    themes[-1].append(n)
    themes0, themes1 = [], []
    for t in themes:
        themes0.append(t[0])
    for t in themes:
        themes1.append(t[1:])
    return(themes0, themes1)
def czekajnaklawisz(klawisze):
    while not msvcrt.kbhit():
        sleep(0)
    key = msvcrt.getch()
    i = 0
    for kl in klawisze:
        if key == kl:
            return(i)
        i += 1
    return(czekajnaklawisz(klawisze))
def zgraj(update):
    if not update:
        print("INTERLAVIC OFFLINE LEARNING TOOL.EXE\nSeems like you're using it first time...")
        input("watch this tutorial: https://youtu.be/hsv8TjJhtxA  and type 'done' when you will know what to do.\n")
    else:
        clear()
        print("Updating: INTERLAVIC OFFLINE LEARNING TOOL.EXE")
    loc = input("Enter localisation of the folder:\n")
    print("Wait...")
    Idict = xlrd.open_workbook(loc+"\\Unoffical_Interslavic_Offline_Word_Learning_Tool\\new_interslavic_words_list.xlsx")
    sl = Idict.sheet_by_index(0)
    r = sl.nrows
    c = sl.ncols
    if not update:
        os.chdir("C:\\")
        os.mkdir("InterslavicDictionary")
    for lang in langs:
        txt = open("C:\\InterslavicDictionary\\"+lang+".txt", "w")
        txt.close()
    for l in range(len(langs)):
        lang = langs[l]
        txt = open("C:\\InterslavicDictionary\\"+lang+".txt", "a", encoding = "utf-8")
        i = 0
        print("Installing language:", lang)
        while i < r:
            print((sl.cell_value(i, lines[l])), file = txt)
            i += 1
        txt.close()
    print("Installing themes...")
    sheet1 = OpenGoogleSheet("test arkusza", loc+"\\Unoffical_Interslavic_Offline_Word_Learning_Tool\\test projektu-9f91d648512d.json")
    cell = "nothing"
    i = 0
    str1, str2, str3 = "", "", ""
    while cell != "":
        i+=1
        cell= sheet1.cell(i, 1).value
        str1+= cell
        str1+= "\n"
        str1+=sheet1.cell(i, 2).value
        str1+="\n"
        str2+=cell
        str2+= "\n"
        str2+=sheet1.cell(i, 3).value
        str2+="\n"
        str3+=cell
        str3+= "\n"
        str3+=sheet1.cell(i, 4).value
        str3+="\n"
    txt = open("C:\\InterslavicDictionary\\themes(en).txt", "w", encoding = "utf-8")
    txt.write(str1)
    txt.close()
    txt = open("C:\\InterslavicDictionary\\themes(isv).txt", "w", encoding = "utf-8")
    txt.write(str2)
    txt.close()
    txt = open("C:\\InterslavicDictionary\\things_that_cant_be_in_themes.txt", "w", encoding = "utf-8")
    txt.write(str3)
    txt.close()
    txt = open("C:\\InterslavicDictionary\\version.txt", "w", encoding = "utf-8")
    txt.write(sheet1.cell(1, 5).value)
    txt.close()
    print("Saving localisation...")
    txt = open("C:\\InterslavicDictionary\\localisation.txt", "w", encoding = "utf-8")
    txt.write(loc+"\\Unoffical_Interslavic_Offline_Word_Learning_Tool\\test projektu-9f91d648512d.json")
    txt.close()
    print("Done :). Press space to continue.")
    czekajnaklawisz([b' '])
def przepytaj(n, f, s2, s3):
    i = 0
    nn, nf = [], []
    while i < len(n):
        clear()
        print(s2, "\n", n[i], sep = "")
        sleep(0.2)
        czekajnaklawisz([b' '])
        print("\n", f[i], "\n", s3, sep = "")
        pn, pf = n[i], f[i]
        sleep(0)
        keynum = czekajnaklawisz([b'K', b'M'])
        if keynum == 0:
            nn.append(pn)
            nf.append(pf)
        i += 1
    if nn != []:
        przepytaj(nn, nf, s2, s3)
def run(l2, l1, s1, s2, s3, thm):
    foreign = open("C:\\InterslavicDictionary\\"+l1+".txt", "r", encoding = "utf-8").readlines()
    native = open("C:\\InterslavicDictionary\\"+l2+".txt", "r", encoding = "utf-8").readlines()
    Pl, Is = [], []
    if thm == -1:
        Pl, Is = native[1:], foreign[1:]
    else:
        for n in thm:
            Pl.append(native[n])
            Is.append(foreign[n])
    maxx = len(Pl)
    i = 0
    nf, nn = [], []
    nmx = 0
    while i < maxx:
        if not(Is[i] == "!\n" or Pl[i] == "!\n"):
            nmx += 1
            nf.append(Is[i])
            nn.append(Pl[i])
        i+=1
    ran = input(s1+str(nmx)+"):\n")
    ms = ""
    tr = []
    for zn in ran:
        if zn != "-":
            ms += zn
        else:
            tr.append(int(ms))
            ms = ""
    tr.append(int(ms))
    mn = tr[0]-1
    mx = tr[1]-1
    key = b''
    i = 0
    newis = []
    newnt = []
    while i <= mx-mn:
        rand = random.randint(mn, mx-i)
        newis.append(nf.pop(rand))
        newnt.append(nn.pop(rand))
        i += 1
    przepytaj(newnt, newis, s2, s3)
    print("You already know all the words!\nPress space to back to menu.")
    czekajnaklawisz([b' '])
    while True:
        menu()
def plus1(var, minv, maxv):
    var += 1
    if var > maxv:
        var = minv
    return(var)
def minus1(var, minv, maxv):
    var -= 1
    if var < minv:
        var = maxv
    return(var)
def printmenu():
    ous = "INTERLAVIC OFFLINE LEARNING TOOL.EXE (v "
    ous += str(version)
    ous += ")\nArrows to change settings, space to run.\n"
    ous += ((" ") * (l//2+1))
    if p2 == 0:
        ous += "v"
    else:
        ous += " "
    ous += (((l//2+1)*2)*" "+"    ")
    if p2 == 1:
        ous += "v"
    else:
        ous += " "
    ous += (((" ") * (l//2+1))+"    "+((" ") * (l2//2+1)))
    if p2 == 2:
        ous += "v"
    else:
        ous += " "
    ous += (((" ") * (l2//2+1))+"    "+((" ") * (l3//2+1)))
    if p2 == 3:
        ous += "v"
    else:
        ous += " "
    ous += ((" ") * (l3//2+1))
    ous += "\n"
    for i in range(max(len(langs), len(themes))):
        lang, theme = "", ""
        if i < len(themes):
            theme = themes[i]
        if i < len(langs):
            lang = langs[i]
        if p1[0] == i:
            ous += (">" + lang + (" "*(l-len(lang))) + "<")
        else:
            ous += (" " + lang + (" "*(l-len(lang))) + " ")
        if i == len(langs)//2:
            ous += " to "
        else:
            ous += "    "
        if p1[1] == i:
            ous += (">" + lang + (" "*(l-len(lang))) + "<")
        else:
            ous += (" " + lang + (" "*(l-len(lang))) + " ")
        ous += "    "
        if p1[2] == i:
            ous += (">" + theme + (" "*(l2-len(theme))) + "<")
        else:
            ous += (" " + theme + (" "*(l2-len(theme))) + " ")
        if i == 0:
            ous += "    NEWS&UPDATES\n"
        else:
            ous += "\n"
    ous += ((" ") * (l//2+1))
    if p2 == 0:
        ous += "^"
    else:
        ous += " "
    ous += (((l//2)*2)*" "+"    ")
    if p2 == 1:
        ous += "^"
    else:
        ous += " "
    ous += (((" ") * (l//2+1))+"    "+((" ") * (l2//2+1)))
    if p2 == 2:
        ous += "^"
    else:
        ous += " "
    ous += (((" ") * (l2//2+1))+"    "+((" ") * (l3//2+1)))
    if p2 == 3:
        ous += "^"
    else:
        ous += " "
    ous += ((" ") * (l3//2+1))
    ous += "\n"
    print(ous)
def printnews():
    clear()
    try:
        updating = False
        urlopen("https://www.google.com", timeout = 1)
        print("Checking...")
        sheet1 = OpenGoogleSheet("test arkusza", localisation)
        s = sheet1.cell(2, 5).value
        ous = ""
        for zn in s:
            if zn == "\\":
                ous += "\n"
            else:
                ous += zn
        clear()
        print(ous)
        newversion = sheet1.cell(1, 5).value
        if version != newversion:
            print("\nThere's a newer version! press 'u' to update, space to go back to menu.")
        else:
            print("\nThere are no updates.\nPress space to go back to menu.")
        key = czekajnaklawisz([b' ', b'u', b'U'])
        if key == 1 or key == 2:
            updating = True
            zgraj(True)
    except:
        if updating:
            print("Failed to update.\nPress space to go back to menu.")
        else:
            print("No internet connection.\nPress space to go back to menu.")
        czekajnaklawisz([b' '])
def menu():
    global p1, p2
    clear()
    printmenu()
    key = czekajnaklawisz([b'K',b'M',b'P',b'H',b' '])
    for i in range(len(langs)):
        lang = langs[i]
    if key == 3:
        if p2 != 2:
            p1[p2] = minus1(p1[p2], 0, len(langs)-1)
        else:
            p1[2] = minus1(p1[p2], 0, len(themes)-1)
    elif key == 2:
        if p2 != 2:
            p1[p2] = plus1(p1[p2], 0, len(langs)-1)
        else:
            p1[2] = plus1(p1[p2], 0, len(themes)-1)
    elif key == 1:
        p2 = plus1(p2, 0, 3)
    elif key == 0:
        p2 = minus1(p2, 0, 3)
    elif key == 4:
        if p2 != 3:
            getready(langs[p1[0]],langs[p1[1]], tns[p1[2]])
        else:
            printnews()
    sleep(0.1)
def getready(l1, l2, thm):
    st1 = "Enter the range of words you want to learn (for example '1-50', '100-140'; total range 1-"
    st2 = "Press space to see the answer!"
    st3 = "Have you guessed?\nYes - press right arrow, No - press left arrow."
    run(l1, l2, st1, st2, st3, thm)
#########################################################################################       
try:
    Is = open("C:\\InterslavicDictionary\\"+langs[0]+".txt", "r", encoding = "utf-8").read()
    Is = ""
except FileNotFoundError:
    zgraj(False)
themes, tns, l2 = doths()
while True:
    menu()
