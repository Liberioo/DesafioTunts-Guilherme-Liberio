from googleapiclient.discovery import build
from google.oauth2 import service_account
from math import ceil

SERVICE_ACCOUNT_FILE = 'keys.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

creds = None
creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# The ID spreadsheet.
SAMPLE_SPREADSHEET_ID = '1yqNrWGScqbnVtA799WJbEgkmruqBJtK1ZW7O3DV1y6I'

service = build('sheets', 'v4', credentials=creds)

# Call the Sheets API
sheet = service.spreadsheets()


students = [] #empty list that will be used to store lists elements with info of each student
i = 0 #iterator starting at 0
r = "engenharia_de_software!C4:F4" #initializes our range with the first line with possible values
info = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=r).execute() #initializes the info accordint to initial range


#loop that will continue until there are no more registered values in the spreadsheet
while(info.get('values', [])):
    current = info.get('values', [])[0] #since it's getting a single list out of this, accesses the 0 position to not get a list containing a single list element
    students.append(current) #adds the current information list to the end of the students list
    i+=1 #increments iterator
    r = "engenharia_de_software!C" + str(4+i) + ":F" + str(4+i) #updates range to the next line of values
    info = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=r).execute() #updates the info with the new range


results = [] #list that will have the result information to be registered on the spreadsheet

total_classes = (sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range="engenharia_de_software!A2").execute()) #gets the field with total classes from the spreadsheet
total_classes = total_classes.get('values', [])[0][0] #converts the list of lists into the type of the single element it gets (string)
total_classes = int(total_classes[len(total_classes)-3:len(total_classes)]) #gets the last 3 digits of the string where the number of total classes is located and converts it to int


for s in students:
    attendance = (total_classes - int(s[0]))/total_classes #calculates the attendance for each element of the students list
    avg = ceil((int(s[1]) + int(s[2]) + int(s[3]))/3) #calculates the average grade between 0 and 100 and rounds up the division
    feg = 0 #initializes the grande in the final exam as 0
    if attendance >= 0.75:
        if avg >= 70:
            results.append(['Aprovado', feg]) #adds a list with situation and required final exam grade to the end of the results list
        elif avg >= 50:
            feg = float(100 - avg)/10 #calculates the grade required for the final exam and converts to a float from 1 to 10
            feg = "{:.1f}".format(feg) #formats the required grade to have only one decimal place 
            results.append(['Exame Final', feg])
        else:
            results.append(['Reprovado por Nota', feg])
    else:
        results.append(['Reprovado por Falta', feg])


#updates the spreadsheets with the results list
request = service.spreadsheets().values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range="engenharia_de_software!G4", valueInputOption="USER_ENTERED", body={'values':results})
response = request.execute()