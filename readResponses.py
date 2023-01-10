#use pipreqs

from __future__ import print_function
from asyncore import read
from collections import defaultdict
from datetime import datetime
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



class readForm():

    def __init__(self) -> None:
        self.formAnswers = []

    def createClientFormsAndSheet(self, clientName, sessionId, collaborativeId, sheetName):

        SCOPES = ['https://www.googleapis.com/auth/drive']

        creds = None
        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists('create.json'):
            creds = Credentials.from_authorized_user_file('create.json', SCOPES)
        # If there are no (valid) credentials available, let the user log in.
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

        copiedFile = {"name" : "Session Form For " + clientName}

        sessionFormId = service.files().copy(fileId=sessionId, body = copiedFile).execute()

        copiedFile = {"name" : "Collaborative Form For " + clientName}

        collaborativeFormId = service.files().copy(fileId = collaborativeId, body = copiedFile).execute()

        sheet = gspread.oauth(credentials_filename = r"E:\Coding\formAutomation\client_secrets.json", authorized_user_filename=r"E:\Coding\formAutomation\token1.json")

        sheet = sheet.open(sheetName)

        cId = sheet.add_worksheet(title = "Collaborative Sheet For " + clientName, rows=1000, cols=50).id

        sId = sheet.add_worksheet(title = "Session Sheet For " + clientName, rows=1000, cols=50).id

        return sessionFormId['id'], collaborativeFormId['id'], cId, sId


    def readResponses(self, formId):

        SCOPES = "https://www.googleapis.com/auth/forms.responses.readonly"
        DISCOVERY_DOC = "https://forms.googleapis.com/$discovery/rest?version=v1"

        """Shows basic usage of the Admin SDK Directory API.
        Prints the emails and names of the first 10 users in the domain.
        """
        creds = None
        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists('readResponse.json'):
            creds = Credentials.from_authorized_user_file('readResponse.json', SCOPES)
        # If there are no (valid) credentials available, let the user log in.
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

        result = service.forms().responses().list(formId=formId).execute()

        self.rawResponses = result

    def getForm(self, formId):
        SCOPES = "https://www.googleapis.com/auth/forms.body.readonly"

        """Shows basic usage of the Admin SDK Directory API.
        Prints the emails and names of the first 10 users in the domain.
        """
        creds = None
        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists('getForm.json'):
            creds = Credentials.from_authorized_user_file('getForm.json', SCOPES)
        # If there are no (valid) credentials available, let the user log in.
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

        result = service.forms().get(formId=formId).execute()
        self.rawForm = result

    def findDicInList(self, lst, key, value):
        for i, dic in enumerate(lst):
            try:
                if dic[key] == value:
                    return i
            except:
                pass
        return -1

    #this is run for each person individually
    """
    
    we need dates
    we need each response to each question for each response
    
    """
    def parseResults(self):
        formDataCopy = self.formAnswers
        try:
            toParse = self.rawResponses['responses']
        except KeyError:
            toParse = []
        nameId = 1

        for i, _ in enumerate(formDataCopy):
            formDataCopy[i].update({'dates' : []})
            formDataCopy[i].update({'responses' : []})
            formDataCopy[i].update({'responsesRaw' : []})
            formDataCopy[i].update({'name' : []})
            if _['questionTitle'] == "Responder's Name":
                nameId = _['questionId']

        #json.dumps(formDataCopy)

        for response in toParse:
            date = response['lastSubmittedTime']
            for answer in list(response['answers'].keys()):

                value = response['answers'][answer]['textAnswers']['answers'][0]['value']

                dicToIndex = self.findDicInList(formDataCopy, 'questionId', answer)

                if answer == nameId:
                    for i, _ in enumerate(formDataCopy):
                        formDataCopy[i]['name'].append(value)

                formDataCopy[dicToIndex]['dates'].append(date)

                formDataCopy[dicToIndex]['responsesRaw'].append(value)


                if formDataCopy[dicToIndex]['toColor'] == True:
                    formDataCopy[dicToIndex]['responses'].append(formDataCopy[dicToIndex]["answers"].index(value))

        self.responses = formDataCopy

        



    def parseForm(self): #returns the title of the question and each possible answer for each multiple choice question in the form
        formList = []
        inData = self.rawForm
        inData = inData['items']

        for i, key in enumerate(inData):
            dic = defaultdict()

            try:
                key['questionItem']
            except:
                continue



            dic['i'] = i
            dic['questionTitle'] = key['title']
            dic['answers'] = []

            dic['toColor'] = 'choiceQuestion' in list(key["questionItem"]["question"].keys())

            dic['questionId'] = key["questionItem"]["question"]['questionId']
            
            if dic['toColor'] == True: 
                if key['title'] == 'Session Type':
                    dic['toColor'] = False

                for dics in key["questionItem"]["question"]['choiceQuestion']['options']:
                    if 'isOther' in list(dics.keys()):
                        dic['toColor'] = False

            if dic['toColor']:
                for question in key["questionItem"]["question"]["choiceQuestion"]["options"]:
                    dic['answers'].append(question["value"])


            formList.append(dic)


        self.formAnswers = formList

    def writeSpreadsheet(self, sheetName, sheetId):

        #2d [[], []], inside is row and outside is column
        #group and sort by name then sort by date

        colors = defaultdict()

        l = 0
        r = 0

        for i, x in enumerate(self.responses):
            if x['toColor'] == True:
                l = len(x['answers'])
                r = len(x['responses'])
                break

        for x in range(l):

            red = 1 - 1*(x/(-1+l))
            green = 1*(x/(-1+l))

            colors[x] = {
                'red' : red,
                'green' : green,
                'blue' : 0.0,
                'alpha' : 0.6
            }

        rows = r

        sheet = gspread.oauth(credentials_filename = r"E:\Coding\formAutomation\client_secrets.json", authorized_user_filename=r"E:\Coding\formAutomation\token1.json")

        sheet = sheet.open(sheetName)
        
        sheet = sheet.get_worksheet_by_id(sheetId)

        sheet.clear()

        answers = [[] for x in range(rows)]
        formatting = []
        toSkip = [0, 1]
        toMatch = []

        for i, question in enumerate(self.responses):
            if not question['toColor']: toSkip.append(i+2)
            C = ord('C')
            column = chr(i + C)
            for z, response in enumerate(question['responsesRaw']):
                answers[z].append(response)
            for x in question['answers']:
                if x not in toMatch: toMatch.append(x)
                
        toMatch.reverse()

        for i, name in enumerate(self.responses[0]['name']):
            answers[i].insert(0, self.responses[0]['dates'][i][:10] + " " + self.responses[0]['dates'][i][11:19])
            answers[i].insert(0, name)

        eachPerson = defaultdict()

        for person in answers:
            try:
                eachPerson[person[0]]
            except:
                eachPerson[person[0]] = []

            eachPerson[person[0]].append(person)

        answers = []

        for key in sorted(list(eachPerson.keys())):
            #eachPerson[key] = sorted(eachPerson[key], key=lambda t: datetime.strptime(t[1], '%Y-%m-%d %H:%M:%S'))
            for x in sorted(eachPerson[key], key=lambda t: datetime.strptime(t[1], '%Y-%m-%d %H:%M:%S')):
                answers.append(x)
            answers.append([])

        for i, question in enumerate(answers):

            for p, answer in enumerate(question):
                if p not in toSkip:
                    fmt = CellFormat(
                        backgroundColor = colors[toMatch.index(answer)]
                    )
                    formatting.append((self.column(p+1) + str(i+2), fmt))

        try:
            tmp = len(answers[0])
        except IndexError:
            tmp = 0

        sheet.update("A2:"+ self.column(tmp) + str(len(answers) + 2), answers)
        format_cell_range(sheet, 'A:' + self.column(tmp), CellFormat(backgroundColor={"red": 1, "green": 1, "blue": 1}))
        try:
            format_cell_ranges(sheet, formatting)
        except gspread.exceptions.APIError:
            pass
        

    def column(self, i):
        a = ord('A')

        i -= 1

        if i + a > ord("Z"):

            return "" + chr((i//26) + ord("A") - 1) + chr((i%26) + ord("A"))

        return "A" if i+a < 65 else chr(i + a)





