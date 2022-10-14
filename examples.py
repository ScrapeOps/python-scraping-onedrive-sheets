
from pydoc import pager
import requests
import os
from ms_graph import generate_access_token, GRAPH_API_ENDPOINT
from microsoftgraph.client import Client


# Get the following details by create an app in your Microsoft Azure Portal
# 1. Go to App registrations page
# 2. Register your app

# Your Application (client) ID
CLIENT_ID = 'YOUR_CLIENT_ID'

# Your Application (client) secret value
CLIENT_SECRET = 'YOUR_CLIENT_SECRET'

# Your OneDrive Excel Sheet Id
WORKBOOK_ID = 'YOUR_SHEET_ID';


# The scopes you wish your auth token to have
SCOPES = ['files.readwrite.all']


# LIST OF DIFFERENT SCOPES
# files.read 
# files.read.all 
# files.readwrite 
# files.readwrite.all 
# offline_access




# AUTHENTICATION SECTION 
CLIENT = '9cef1616-a075-4596-8b2b-90cf9fa06579'
SCOPES = ['files.readwrite.all']


# accessToken = find_access_token(CLIENT, SCOPES)
# headers = {
#     'Authorization' : 'Bearer ' + accessToken['access_token']
# }



# POST TO SHEET TABLE EXAMPLE
tableName = 'Table1';

# url = 'https://graph.microsoft.com/v1.0/me/drive/items/'+  sheetId + '/workbook/tables/'+ tableName +'/rows/add'


# data = {
#         "values": [
#             ["John", "Doe", "+1 305 1234567"],
#         ]
# }

# x = requests.post(url, json = data, headers = headers)
# print(x.json())




# UPLOAD FILE PYTHON REQUEST EXAMPLE
# filePath = '/Applications/MAMP/htdocs/video-guides/python-scrapy-beginner-series/scrapy-indeed/basic-scrapy-project/text.txt'
# fileName = os.path.basename(filePath)
# with open(filePath, 'rb') as upload:
#     media_content = upload.read()

# response = requests.put(
#     GRAPH_API_ENDPOINT + '/me/drive/items/root:/{fileName}:/content',
#     headers = headers,
#     data = media_content
# )
# print(response)








# UPDATE SPREADSHEET RANGE EXAMPLE

# client = Client(CLIENT, KEY, account_type='common') # by default common, thus account_type is optional parameter.

# accessToken = generate_access_token(CLIENT, SCOPES)
# token = accessToken
# client.set_token(token)

# https://learn.microsoft.com/en-us/graph/api/range-update?view=graph-rest-1.0&tabs=javascript
# # response = client.users.get_me()
# # response = client.files.drive_root_items()
# # print(response.data)

# # workbook_id = '77C597EFB8B5B71B'


# #get the workbook_id from the sheet url: resid=XXXXXXXXXX
# workbook_id = '77C597EFB8B5B71B!840'
# response = client.workbooks.list_worksheets(workbook_id)
# print(response.data)


# workbook_id = '77C597EFB8B5B71B!840'
# worksheet_id = '{A41E3F32-A5CD-448C-A2C9-F5445D791EA6}'


# response1 = client.workbooks.create_session(workbook_id)
# workbook_session_id = response1.data["id"]

# client.set_workbook_session_id(workbook_session_id)


# response = client.workbooks.get_used_range(workbook_id, worksheet_id)
# print(response.data)

# range_address = ":"
# data = {
#         "values": [
#             ["John", "Doe", "+1 305 1234567", "Miami, FL"],
#             ["Bill", "Gates", "+1 305 1234567", "St. Redmond, WA"],
#         ]
# }
# response2 = client.workbooks.update_range(workbook_id, worksheet_id, range_address, json=data)
# response3 = client.workbooks.close_session(workbook_id)










# MANUALLY INSERTING A ROW
# UPDATE RANGE EXAMPLE 2
# https://learn.microsoft.com/en-us/graph/api/range-update?view=graph-rest-1.0&tabs=javascript


# await client.api('/me/drive/items/{id}/workbook/worksheets/{sheet-id}/range(address='A1:B2')')


# url = 'https://graph.microsoft.com/v1.0/me/drive/items/'+  sheetId + '/workbook/tables/'+ tableName +'/rows/add'


# workbookId = '77C597EFB8B5B71B!840'
# worksheetId = '{A41E3F32-A5CD-448C-A2C9-F5445D791EA6}'

# url = "https://graph.microsoft.com/v1.0/me/drive/items/"+  workbookId + "/workbook/worksheets/"+ worksheetId + "/range(address='A1:B1')"


# data = {
#         "values": [
#             ["John", "Doe"],
#         ],
#         "formulas": [
#             [None, None]
#         ],
#         "numberFormat": [
#             [None, None]
#         ],
# }

# print(url)

# x = requests.patch(url, json = data, headers = headers)
# print(x.json())



accessToken = generate_access_token(CLIENT, SCOPES)
headers = {
    'Authorization' : 'Bearer ' + accessToken['access_token']
}


# Add new table 

tableData = {
    "address": "Sheet1!A1:D5",
    "hasHeaders": True
}


url = GRAPH_API_ENDPOINT + '/me/drive/items/'+  sheetId + '/workbook/tables/'+ tableName

headers = {
        'Authorization' : 'Bearer ' + accessToken
    }

returnedRequest = requests.post(url, json=tableData, headers = headers)

print(returnedRequest.json())
