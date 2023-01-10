import datetime
import json
from asyncore import read
from collections import defaultdict
from datetime import datetime
from math import fabs
import os
from re import A

from apiclient import discovery
#from httplib2 import Credentials, Http
import httplib2
from oauth2client import client, file, tools
from requests import JSONDecodeError
from settings import templateId
import json
from oauth2client.service_account import ServiceAccountCredentials
import gspread
from gspread_formatting import *
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials


class Question():
    def __init__(self, question : json) -> None:

        #all items have these data fields

        self.itemId = question['itemId']

        try:
            self.title = question['title']
        except KeyError:
            self.title = "Date"

        self.format = False
        self.answers = []
        self.answerIndexes = []
        self.name = "Responder's Name" in self.title

        self.descriptionAndNotes = self.title.lower().replace(" ", '') in [x.lower().replace(" ", '') for x in ['Language & Communication Skills', 'Attention and Working Memory Skills',
                                                 'Emotion- and Self-Regulation Skills', 'Cognitive Flexibility Skills', 'Social Thinking Skills']]

        self.toColor = False

        #checking type of item

        if "textItem" in list(question.keys()) or "pageBreakItem" in list(question.keys()):
            self.type = "Desciption"

        else:
            qD = question['questionItem']['question']

            self.questionId = qD['questionId']

            if 'textQuestion' in list(qD.keys()):
                self.type = 'shortAnswerQuestion'

            elif "dateQuestion" in list(qD.keys()):
                self.type = "Date"

            else:
                self.type = 'choiceQuestion'

                self.options = [value[list(value.keys())[0]] for value in qD[self.type]['options']]
                self.options.reverse()

                self.toColor = True not in self.options or self.title == "Session Type"

    def loadQuestion(self, answer : json):

        try:

            tmpAnswer = answer['textAnswers']['answers'][0]['value']
            self.answers.append(tmpAnswer)

            if self.toColor:
                self.answerIndexes.append(self.options.index(tmpAnswer))

        except KeyError:
            self.answers.append("")

    def averageChange(self) -> str:

        inital = self.answerIndexes[0]

        final = self.answerIndexes[-1]

        if final - inital > 0:
            return "+" + final - inital

        elif final - inital < 0:
            return "-" + final - inital

        else:
            return 0


class Response():

    def __init__(self, formTemplate : json) -> None:

        self.questions : Question = []
        self.names : str = []
        self.dates : str = []

        for question in formTemplate['items']:

            self.questions.append(Question(question))

    def readResponses(self, responses : json) -> None:

        for response in responses['responses']:

            date = response['lastSubmittedTime'].replace("T", " ")
            date = date[:date.find(" ") + 9]

            self.dates.append(datetime.strptime(date, '%Y-%m-%d %H:%M:%S'))

            for i, question in enumerate(self.questions):


                try:
                    self.questions[i].loadQuestion(response['answers'][question.questionId])
                except (KeyError, AttributeError):
                    self.questions[i].loadQuestion({})

                if question.name == True:
                    self.names.append(question.answers[-1])

    def sortResponses(self):

        self.order = list(range(len(self.names)))

        tmp = sorted(zip(self.names, self.order, self.dates), key = lambda name : name[0].lower())

        self.order, self.names, self.dates = [], [], []

        tmp2 = []

        i = 0

        while i < len(tmp):

            for z, group in enumerate(tmp[i:]):
                if group[0] != tmp[i][0]:
                    break

            else:
                tmp2.append([x for x in tmp[i:]])
                break

            tmp2.append([x for x in tmp[i:z+i]])


            i += z

        for person in tmp2:

            tmp = sorted(person, key = lambda p : p[2])

            for time in tmp:
                self.names.append(time[0])
                self.order.append(time[1])
                self.dates.append(time[2])

            self.names.append(None)
            self.order.append(None)
            self.dates.append(None)

    def prepareToWrite(self) -> None:

        self.sortResponses()

        self.toWrite = [[title.title for title in self.questions]] #titles and the intial two empty spaces
        self.toWrite[0].insert(0, '')
        self.toWrite[0].insert(0, '')

        self.formatting = []
        
        for i, thing in enumerate(zip(self.order, self.names, self.dates)):

            index = thing[0]

            if index == None:
                self.toWrite.append(["" for x in range(len(self.questions) + 2)])
                continue

            tmpToWrite = [thing[1], thing[2].strftime("%Y-%m-%d %H:%M:%S")]

            #Column starts two over due to name and date, row is one down
            for z, question in enumerate(self.questions):
                tmpToWrite.append(question.answers[index])

                if question.toColor and question.answers[index] != "":
                    self.formatting.append((self.column(z+3) + str(i+2), CellFormat(backgroundColor = self.formatQuestion(question, index))))

            self.toWrite.append(tmpToWrite)

        tmpIndex = []

        for i, question in enumerate(self.questions):
            if question.descriptionAndNotes:
                tmpIndex.append(i)

        self.toWrite = ("A1:" + self.column(z+3) + str(i+2), self.toWrite)

    def column(self, i) -> str:
        a = ord('A')

        i -= 1

        if i + a > ord("Z"):

            return "" + chr((i//26) + ord("A") - 1) + chr(((i)%26) + ord("A"))

        return "A" if i+a < 65 else chr(i + a)

    def formatQuestion(self, question : Question, index : int) -> json:

        
        i = question.options.index(question.answers[index])

        options = len(question.options)

        red = 1 - 1*(i/(-1+options))
        green = 1*(i/(-1+options))

        return {
            'red' : red,
            'green' : green,
            'blue' : 0.0,
            'alpha' : 0.6
        }


class Wrapper():

    def __init__(self, formId : str, sheetIndex : int) -> None:
        self.formId = formId
        self.sheetIndex = sheetIndex

    def getForm(self):

        SCOPES = "https://www.googleapis.com/auth/forms.body.readonly"

        """Shows basic usage of the Admin SDK Directory API.
        Prints the emails and names of the first 10 users in the domain.
        """
        creds = None
        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        try:

            if os.path.exists('getForm.json'):
                creds = Credentials.from_authorized_user_file('getForm.json', SCOPES)

            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    creds.refresh(Request())
                else:
                    flow = InstalledAppFlow.from_client_secrets_file(
                        'client_secrets.json', SCOPES)
                    creds = flow.run_local_server(port=0)
                # Save the credentials for the next run
                with open('getForm.json', 'w') as token:
                    token.write(creds.to_json())

        except:

            os.remove("getForm.json")

            creds = None

            if os.path.exists('getForm.json'):
                creds = Credentials.from_authorized_user_file('getForm.json', SCOPES)

            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    creds.refresh(Request())
                else:
                    flow = InstalledAppFlow.from_client_secrets_file(
                        'client_secrets.json', SCOPES)
                    creds = flow.run_local_server(port=0)
                # Save the credentials for the next run
                with open('getForm.json', 'w') as token:
                    token.write(creds.to_json())

        service = build('forms', 'v1', credentials=creds)

        result = service.forms().get(formId=self.formId).execute()

        #print(json.dumps(result))

        self.formClass = Response(result)

    def getResponses(self):

        
        SCOPES = "https://www.googleapis.com/auth/forms.responses.readonly"
        DISCOVERY_DOC = "https://forms.googleapis.com/$discovery/rest?version=v1"

        """Shows basic usage of the Admin SDK Directory API.
        Prints the emails and names of the first 10 users in the domain.
        """
        creds = None
        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        try:

            if os.path.exists('readResponse.json'):
                creds = Credentials.from_authorized_user_file('readResponse.json', SCOPES)

            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    creds.refresh(Request())
                else:
                    flow = InstalledAppFlow.from_client_secrets_file(
                        'client_secrets.json', SCOPES)
                    creds = flow.run_local_server(port=0)
                # Save the credentials for the next run
                with open('readResponse.json', 'w') as token:
                    token.write(creds.to_json())

        except:

            os.remove("readResponse.json")

            creds = None

            if os.path.exists('readResponse.json'):
                creds = Credentials.from_authorized_user_file('readResponse.json', SCOPES)

            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    creds.refresh(Request())
                else:
                    flow = InstalledAppFlow.from_client_secrets_file(
                        'client_secrets.json', SCOPES)
                    creds = flow.run_local_server(port=0)
                # Save the credentials for the next run
                with open('readResponse.json', 'w') as token:
                    token.write(creds.to_json())

        service = build('forms', 'v1', credentials=creds)


        # Prints the title of the sample form:

        result = service.forms().responses().list(formId=self.formId).execute()

        json.dump(result, open('test.json', 'w'), indent = 1)

        self.formClass.readResponses(result)

    def writeToSpreadsheet(self, sheetName : str):

        self.getForm()
        self.getResponses()

        self.formClass.prepareToWrite()

        sheet = gspread.oauth(credentials_filename = r"E:\Coding\formAutomation\client_secrets.json", authorized_user_filename=r"E:\Coding\formAutomation\token1.json")

        sheet = sheet.open(sheetName)

        sheet = sheet.get_worksheet_by_id(self.sheetIndex)

        format_cell_range(sheet, self.formClass.toWrite[0], CellFormat(backgroundColor={"red": 1, "green": 1, "blue": 1}))

        sheet.update(self.formClass.toWrite[0], self.formClass.toWrite[1])
        format_cell_ranges(sheet, self.formClass.formatting)
        
        #not logging answers

    def createUser(self, name : str, sessionFormId : str, collaborativeFormId : str, sheetName : str):

        SCOPES = ['https://www.googleapis.com/auth/drive']

        creds = None
        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists('create.json'):
            creds = Credentials.from_authorized_user_file('create.json', SCOPES)
        # If there are no (valid) credentials available, let the user log in.

        try:

            if os.path.exists('create.json'):
                creds = Credentials.from_authorized_user_file('create.json', SCOPES)

            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    creds.refresh(Request())
                else:
                    flow = InstalledAppFlow.from_client_secrets_file(
                        'client_secrets.json', SCOPES)
                    creds = flow.run_local_server(port=0)
                # Save the credentials for the next run
                with open('create.json', 'w') as token:
                    token.write(creds.to_json())

        except Exception as e:


            os.remove("create.json")

            creds = None

            if os.path.exists('create.json'):
                creds = Credentials.from_authorized_user_file('create.json', SCOPES)

            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    creds.refresh(Request())
                else:
                    flow = InstalledAppFlow.from_client_secrets_file(
                        'client_secrets.json', SCOPES)
                    creds = flow.run_local_server(port=0)
                # Save the credentials for the next run
                with open('create.json', 'w') as token:
                    token.write(creds.to_json())


        service = build('drive', 'v3', credentials=creds)

        copiedFile = {"name" : "Session Form For " + name}

        sessionFormId = service.files().copy(fileId=sessionFormId, body = copiedFile).execute()

        print(sessionFormId)

        copiedFile = {"name" : "Collaborative Form For " + name}

        collaborativeFormId = service.files().copy(fileId = collaborativeFormId, body = copiedFile).execute()

        print(sheetName)

        try:

            sheet = gspread.oauth(credentials_filename = r"client_secret_321225505918-fjrig7rr6f6n9bb2guimhr16qg1o6noj.apps.googleusercontent.com.json", authorized_user_filename=r"E:\Coding\formAutomation\token1.json")

            sheet = sheet.open(sheetName)

        except:

            os.remove(r"token1.json")

            sheet = gspread.oauth(credentials_filename = r"client_secret_321225505918-fjrig7rr6f6n9bb2guimhr16qg1o6noj.apps.googleusercontent.com.json", authorized_user_filename=r"E:\Coding\formAutomation\token1.json")

            sheet = sheet.open(sheetName)

            print('here')

        cId = sheet.add_worksheet(title = "Collaborative Sheet For " + name, rows=1000, cols=50).id

        sId = sheet.add_worksheet(title = "Session Sheet For " + name, rows=1000, cols=50).id

        return sessionFormId['id'], collaborativeFormId['id'], cId, sId

    def updateInfoSheet(self, sheetName, clientData):

        sheet = gspread.oauth(credentials_filename = r"E:\Coding\formAutomation\client_secret_321225505918-fjrig7rr6f6n9bb2guimhr16qg1o6noj.apps.googleusercontent.com.json", authorized_user_filename=r"E:\Coding\formAutomation\token1.json")

        sheet = sheet.open(sheetName)

        toSend = [['Name', '', 'Session form link', '', 'sesson form Sheet', '', 'collaborative form link', '', 'collaborative form sheet']]

        for client in list(clientData['clients'].keys()):
            toSend.append([client, 
            '',
            f'=HYPERLINK("https://docs.google.com/forms/d/{clientData["clients"][client]["sessionFormId"]}/viewform", "Link")',
            '',
            f'=HYPERLINK("{sheet.url}/edit#gid={clientData["clients"][client]["sessionSheet"]}", "Link")',
            '',
            f'=HYPERLINK("https://docs.google.com/forms/d/{clientData["clients"][client]["collaborativeFormId"]}/viewform", "Link")',
            '',
            f'=HYPERLINK("{sheet.url}/edit#gid={clientData["clients"][client]["collaborativeSheet"]}", "Link")',]
            )

        sheet = sheet.get_worksheet(0)

        sheet.clear()

        sheet.update("A1:ZZZ300", toSend, raw = False)



