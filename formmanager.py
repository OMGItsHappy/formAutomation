from readResponses import *
import json
from classFormManager import *

def updateClient(sessionFormId, collaberativeId, sheetName, cId, sId, client):

    #session form

    test = readForm()
    test.getForm(sessionFormId)
    test.parseForm()
    test.readResponses(sessionFormId)
    test.parseResults()
    test.writeSpreadsheet(sheetName, sId)

    #collaberativeId

    test = readForm()
    test.getForm(collaberativeId)
    test.parseForm()
    test.readResponses(collaberativeId)
    test.parseResults()
    test.writeSpreadsheet(sheetName, cId)


SHEETNAME = 'Client Forms Spreadsheet'
SHEETLINK = 'https://docs.google.com/spreadsheets/d/1EiDgWoVjuS2a7VF7sSjyCTGe48P8ZI_pPsVhMwGKhns/edit'
SESSIONFORMCOPYID = "1DMBe4JMalTKdhNOFjvLIqBArG5mR4eoZJYhLIjbDBXk"
COLLABORATIVEFORMCOPYID = "1XHWrvwd8csH0__3z_M3n_8MOmy80sNXfMF_4pZnIU90"


sheet = gspread.oauth(credentials_filename = r"client_secrets.json", authorized_user_filename=r"getSheet.json")

sheet = sheet.open('Client Forms Spreadsheet')

sheet = sheet.get_worksheet(0)

clients = sheet.get_all_values(value_render_option='FORMULA')

data = {
    "sheetName": "Client Forms Spreadsheet",
    "sessionFormCopyId": "1DMBe4JMalTKdhNOFjvLIqBArG5mR4eoZJYhLIjbDBXk",
    "collaborativeFormCopyId": "1XHWrvwd8csH0__3z_M3n_8MOmy80sNXfMF_4pZnIU90",
    "clients": {
    }
}

def getForm(link : str):
    return link[link.find("d/") + 2:link.find("/viewform")]

def getSheet(link : str):
    return link[link.find('#gid=')+5:link.find('",')]

for client in clients[1:]:
    name = client[0]
    sessionFormLink = getForm(client[2])
    sessionFormSheet = getSheet(client[4])
    collaborativeFormLink = getForm(client[6])
    collaborativeFormSheet = getSheet(client[8])

    data["clients"].update({name : 
                                {"sessionSheet" : int(sessionFormSheet),
                                 "collaborativeSheet" : int(collaborativeFormSheet),
                                 "sessionFormId" : sessionFormLink,
                                 "collaborativeFormId" : collaborativeFormLink
                                }})

answer = input("Do you want to create a new client or update current ones (c/u):")

while answer.lower() not in ['c', 'u']:
    answer = input("Invalid input please try again:")

if answer.lower() == 'u':
    for client in list(data['clients'].keys()):
        clientData = data['clients'][client]

        try:

            form = Wrapper(clientData['sessionFormId'], clientData['sessionSheet'])

            form.writeToSpreadsheet(data['sheetName'])
        
        except KeyError:

            print(f"Skipping {client}'s session form as it is empty")

        try:

            form = Wrapper(clientData["collaborativeFormId"], clientData['collaborativeSheet'])

            form.writeToSpreadsheet(data['sheetName'])

        except KeyError:

            print(f"Skipping {client}'s collaborative form as it is empty")

        #updateClient(clientData['sessionFormId'], clientData["collaborativeFormId"], data['sheetName'], clientData['collaborativeSheet'], clientData['sessionSheet'],  client)

elif answer.lower() == "c":
    a = True
    answer = ""
    while a:
        answer = input("What is the client name? :")
        if input(f"Is {answer} right? (y/n)").lower() == "y":
            a = False

    test = Wrapper(None, None)
    tmp = test.createUser(answer, data["sessionFormCopyId"], data["collaborativeFormCopyId"], data['sheetName'])

    data['clients'].update({answer :         
        {
            "sessionSheet" : tmp[3],
            "collaborativeSheet" :tmp[2],
            "sessionFormId" : tmp[0],
            "collaborativeFormId" : tmp[1]
        }})

    test.updateInfoSheet(data['sheetName'], data)



f = open('actualClients.json', 'w')
json.dump(data, f)
f.close()
    
    