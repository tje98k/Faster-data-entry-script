import webbrowser
import sys
import gspread
from oauth2client.service_account import ServiceAccountCredentials
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('driveapi_secret.json', scope)
client = gspread.authorize(creds)
sheet = client.open("VA IRRRL 500k-03.11.20").worksheet("8927_First Financial Mortgage S")
pp = pprint.PrettyPrinter()


#for first financial. my attempt at using scripts to speed up job
#starting point, enter info and have it auto open
#while loop so it keeps looping process until user exits
print('Enter row number to start on now:')
rownumber = input()

while True:

    nameval = str(('A' + str(rownumber) + ':' + 'B' + str(rownumber)))
    csval = str(('H' + str(rownumber) + ':' + 'I' + str(rownumber)))
    adval = str(('G' + str(rownumber) + ':' + 'I' + str(rownumber)))
    namelist = []
    cslist = []
    adlist = []
   
    for cell in sheet.range(nameval):
        namelist.append(cell.value)
    for cell in sheet.range(csval):
        cslist.append(cell.value)
    for cell in sheet.range(adval):
        adlist.append(cell.value)
   
#gets cell ranges with names, city, state, address

    namedash = "-".join(namelist)
    locationdash = "-".join(cslist)
    addressplus = "+".join(adlist)
    #gets a version of name and city/state that the url can read
    urlname = "https://premium.whitepages.com/name/" + namedash +"/" + locationdash +"?type=person_query&button="
    urladdress = "https://premium.whitepages.com/results/address/?type=person_address_query&address=" + addressplus + "&button="
    webbrowser.open(urlname)
    webbrowser.open(urladdress)
#opens two tabs in web browser, with name and address searches

    print("When you're ready, enter the info to split to two sections, then hit enter, then Ctrl D")
    numbers = str(sys.stdin.readlines())
    numbers2 = numbers.replace("(", "\n(")
    numbers3 = numbers2
    letters = []
    numberlist = []
    for char in numbers3:
        if char.isalpha() == True:
            letters.append(char)
        else:
            numberlist.append(char)
    letters = ''.join(letters)
    numberlist = ''.join(numberlist)
    #splits phone type and phone numbers
               
    letters = letters.replace('nnn', '\n')
    letters = letters.replace('n', '')

    numval = str(('D' + str(rownumber)))
    typeval = str(('E' + str(rownumber)))
    sheet.update_acell(typeval, letters)
    numberlist = (str(numberlist).replace("'", ""))
    numberlist = numberlist.replace(',', '')
    numberlist = numberlist.replace('\\' , '')
    numberlist = numberlist.replace("'", "")
    numberlist = numberlist.replace("]", "")
    numberlist = numberlist.replace("[", "")
    #cleans up lists, adds proper breaks etc
    sheet.update_acell(numval, str(numberlist))
   
    rownumber = str(int(rownumber) + 1)
