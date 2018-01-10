import pandas as pd
import json
import requests
#import schedule
import time
import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from goto import with_goto
from splinter import Browser

#@with_goto
def job():
    init = 0
    #label .begin
    print("Waiting for a line to change...")

    sheet = open_spreadsheet('Feuille 1')
    old_value = sheet.cell(2, 1).value

    # Detecteur de mouvement sur le spreadsheet
    while(1):
        today = str(datetime.datetime.now())
        print("date :", today)
        sheet = open_spreadsheet('Feuille 1')
        time.sleep(10)
        print("old_value", old_value)
        new_value = sheet.cell(2, 1).value
        print("new_value: ", new_value)
        if (new_value != old_value or init == 0) and new_value != "":
            break
        else:
            old_value = new_value

    init = 1
    print("Changes, go!go!go!")
    # Ouvre le spreadsheet
    sheet = open_spreadsheet('Feuille 1')

    # Telecharge le spreadsheet en csv
    response = requests.get(
        'https://docs.google.com/spreadsheets/d/11EroRI36SeERHisr50ZE0OE9Nxg9OZ9FnUh1X1BTu3M/gviz/tq?tqx=out:csv&sheet=Germinal_Trumail')
    response.raise_for_status()

    # Converti le format <'bytes'> en <'str'> avant de séparer les lignes
    #print(type(response.content))
    liste_mail = response.content.decode('utf8')
    liste_mail = liste_mail.splitlines()

    # Ne garde que les mails
    liste_mail_verifiee = []
    for i in range(1, len(liste_mail)):
        liste_mail[i] = liste_mail[i].split(',', 1)[0]
        #print('liste_mail[', i, ']: ', liste_mail[i])


    # Verifie chaque mail via l'API de Trumail et inscrit les valides et non valides dans le spreadsheet
    starting_index = int(get_index())
    for i in range(starting_index, len(liste_mail)):

        index_backup(i)
        print('ligne numéro: ', i)
        mail = liste_mail[i]
        mail = mail.replace('"', "")
        print(mail)
        #verification = get_verification(mail)
        re_check(mail, sheet, i+1, liste_mail_verifiee)
        print(liste_mail_verifiee[len(liste_mail_verifiee) - 1])
        """
        if verification == None:
            #write_into_spreadsheet(sheet, i+1, "error", "error", "error", "error", "error")
            liste_mail_verifiee.append((mail, "error", "error", "error", "error", "error"))
            print("error")
        else:
            print(i, verification)
            if verification['deliverable'] == True and verification['catchAll'] == False:
                #write_into_spreadsheet(sheet, i+1, mail, "", "", "", "")
                liste_mail_verifiee.append((mail, mail, "", "", "", ""))
                print("valid mail")

            if verification['catchAll'] == True:
                if(re_check(mail, sheet, i+1)):
                    #write_into_spreadsheet(sheet, i+1, "", mail, "", "", "")
                    liste_mail_verifiee.append((mail, "", mail, "", "", ""))
                    print("catchAll True")

            if verification['deliverable'] == False and verification['catchAll'] == False:
                if verification['hostExists'] == False:
                    #write_into_spreadsheet(sheet, i+1, "", "", "", "", mail)
                    liste_mail_verifiee.append((mail, "", "", "", "", mail))
                    print("hostExists False")
                else:
                    try:
                        if verification['error'] == "Timeout connecting to mail-exchanger":
                            if(re_check(mail, sheet, i+1)):
                                #write_into_spreadsheet(sheet, i+1, "", "", mail, "", "")
                                liste_mail_verifiee.append((mail, "", "", mail, "", ""))
                                print("Error Time Out")
                        else:
                            if(re_check(mail, sheet, i+1)):
                                #write_into_spreadsheet(sheet, i + 1, "", "", "", mail, "")
                                liste_mail_verifiee.append((mail, "", "", "", mail, ""))
                                print("invalide mail")

                    except KeyError:
                        pass
        """

    df = pd.DataFrame(data=liste_mail_verifiee, columns=["Mails à vérifier ", "Mails valides", "Catch all = true",
                                                         "Error Time Out", "Mails invalides", "Serveurs Mails invalides"])
    df.to_csv("Mails_Valides.csv")
    print("")
    print("valid_adresses: ", valid_adress, " ; invalide_mail: ", invalide_mail, " ; catchAll: ", catchAll)

    index_backup(1) # Reset index
    #goto .begin



def open_spreadsheet(name_of_spreadsheet):
    print("Ouverture du Spreadsheet en cours...")
    scope = ['https://spreadsheets.google.com/feeds']
    credentials = ServiceAccountCredentials.from_json_keyfile_name('Apps Script Execution API Quic-0d5369b6757b.json', scope)
    try:
        gc = gspread.authorize(credentials)
        sheet = gc.open('Germinal_Trumail').get_worksheet(1)
        print("Ouverture du Spreadsheet: OK")
    except gspread.exceptions.RequestError:
        print("Cannot open Google Spreadsheet. Retry one more time in 10 seconds to access it.")
        time.sleep(10)
        gc = gspread.authorize(credentials)
        sheet = gc.open('Germinal_Trumail').get_worksheet(1)
    return sheet

def login_spreadsheet():
    scope = ['https://spreadsheets.google.com/feeds']
    credentials = ServiceAccountCredentials.from_json_keyfile_name('Apps Script Execution API Quic-0d5369b6757b.json', scope)
    gc = gspread.authorize(credentials)
    gc.login()
    print("Successfully logged-in !")

def get_verification(mail):
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json"
    }
    params = {
        "limit": 250,
        "page": 1
    }
    url = "https://trumail.io/json/" + mail

    request_result = requests.get(url, params=params, headers=headers)

    try:
        request_result_json = request_result.json()
        pass
    except:
        return None

    return request_result_json


def write_into_spreadsheet(sheet, index_ligne, arg1, arg2, arg3, arg4, arg5):
    try:
        sheet.update_cell(index_ligne, 2, arg1)
        sheet.update_cell(index_ligne, 3, arg2)
        sheet.update_cell(index_ligne, 4, arg3)
        sheet.update_cell(index_ligne, 5, arg4)
        sheet.update_cell(index_ligne, 6, arg5)
    except: #(gspread.exceptions.HTTPError, gspread.exceptions.requests, gspread.exceptions.RequestError) as e:
        print("Cannot access Google Drive. Retry one more time in 10 seconds to access it.")
        time.sleep(10)
        login_spreadsheet()
        sheet.update_cell(index_ligne, 2, arg1)
        sheet.update_cell(index_ligne, 3, arg2)
        sheet.update_cell(index_ligne, 4, arg3)
        sheet.update_cell(index_ligne, 5, arg4)
        sheet.update_cell(index_ligne, 6, arg5)


def create():
    data = [1]
    index = pd.DataFrame(data=data, columns=['Index'])
    index.to_csv('index.csv')


def index_backup(i):
    liste = []
    liste.append(i)
    index_file = pd.DataFrame(data=liste, columns=['Index'])
    index_file.to_csv('index.csv')
    print("")
    print("liste:: ", liste)

def get_index():
    try:
        df = pd.read_csv('index.csv')
        return df['Index'][0]
    except:
        return 1


def re_check(mail, sheet, index_ligne, liste_mail_verifiee):
    # open a browser
    #try:
    print('Re_check protocol launched')
    browser = Browser('chrome')

    # Width, Height
    browser.driver.set_window_size(640, 680)

    print('Browser loaded: OK')

    browser.visit('https://nojunk.io/')

    search_bar_xpath= '//*[@id="demo"]/div/div[1]/div[1]/div/input'
    search_bar = browser.find_by_xpath(search_bar_xpath)[0]
    search_bar.fill(mail)

    search_button_xpath = '//*[@id="demo"]/div/div[1]/div[2]/button'
    search_button = browser.find_by_xpath(search_button_xpath)[0]
    search_button.click()

    print('Request entered: OK')

    time.sleep(10)

    response_xpath = '//*[@id="demo"]/div/div[2]/div'
    response = browser.find_by_xpath(response_xpath)

    print('Get response: OK')


    if response.text == "Adresse is valid" or response.text == "Adresse valide":
        print("Address is valid")
        #write_into_spreadsheet(sheet, index_ligne, mail, "", "", "", "")
        liste_mail_verifiee.append((mail,mail,"","","",""))
        browser.quit()
        return 0
    if response.text == "Address maybe valid" or response.text == "Adresse peut-être valide":
        print("Adresse is maybe valid")
        #write_into_spreadsheet(sheet, index_ligne, "", mail, "", "", "")
        liste_mail_verifiee.append((mail, "", mail, "", "", ""))
        browser.quit()
        return 1

    else:
        print("Adresse invalide")
        #write_into_spreadsheet(sheet, index_ligne, "", "", mail, "", "")
        liste_mail_verifiee.append((mail, "", "", "", mail, ""))
        browser.quit()
    """
    except:
        print("Re_check failed")
        try:
            driver.quit()
        except:
            pass
        return 1
    """




job()
#create()


# json.decoder.JSONDecodeError:
        #return None
    #except simplejson.scanner.JSONDecodeError:
