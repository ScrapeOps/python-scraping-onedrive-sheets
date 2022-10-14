from urllib import request
import webbrowser
from datetime import datetime
import json
import os
import msal
import requests

GRAPH_API_ENDPOINT = 'https://graph.microsoft.com/v1.0'


####
# Sheet id can be found in the URL of your OneDrive excel sheet  
# It is the resid: resid=XXXXXXX
#
# Table name is Table1 for the first table created on the sheet, 
# Table2 for the second table on the sheet, etc
###

def post_to_sheet_table(accessToken, sheetId, tableName, data):

    url = GRAPH_API_ENDPOINT + '/me/drive/items/'+  sheetId + '/workbook/tables/'+ tableName +'/rows/add'

    headers = {
            'Authorization' : 'Bearer ' + accessToken
        }

    returnedRequest = requests.post(url, json = data, headers = headers)
    return returnedRequest


def create_table(accessToken, sheetId, data):
    # POST https://graph.microsoft.com/v1.0/me/drive/items/{id}/workbook/tables/add
    # Content-type: application/json

    # {
    #   "address": "Sheet1!A1:D5",
    #   "hasHeaders": true
    # }

    url = GRAPH_API_ENDPOINT + '/me/drive/items/'+  sheetId + '/workbook/tables/add'

    headers = {
            'Authorization' : 'Bearer ' + accessToken
        }

    returnedRequest = requests.post(url, json=data, headers = headers)
    return returnedRequest



def delete_table(accessToken, sheetId, tableName):
    # DELETE https://graph.microsoft.com/v1.0/me/drive/items/{id}/workbook/tables/{id|name}

    url = GRAPH_API_ENDPOINT + '/me/drive/items/'+  sheetId + '/workbook/tables/'+ tableName

    headers = {
            'Authorization' : 'Bearer ' + accessToken
        }

    returnedRequest = requests.delete(url, headers = headers)
    return returnedRequest





def open_access_token(app_id, scopes):
    # Save Session Token as a token file
    access_token_cache = msal.SerializableTokenCache()

    # read the token file
    if os.path.exists('ms_graph_api_token.json'):
        access_token_cache.deserialize(open("ms_graph_api_token.json", "r").read())
        token_detail = json.load(open('ms_graph_api_token.json',))
        token_detail_key = list(token_detail['AccessToken'].keys())[0]
        token_expiration = datetime.fromtimestamp(int(token_detail['AccessToken'][token_detail_key]['expires_on']))
        print('TOKEN EXPIRY: ')
        print(token_expiration)
        if datetime.now() > token_expiration:
            os.remove('ms_graph_api_token.json')
            access_token_cache = msal.SerializableTokenCache()

    # assign a SerializableTokenCache object to the client instance
    client = msal.PublicClientApplication(client_id=app_id, token_cache=access_token_cache)

    accounts = client.get_accounts()
    if accounts:
        # load the session
        token_response = client.acquire_token_silent(scopes, accounts[0])
    else:
        # authenticate your account 
        print('************ GRAPH TOKEN NOT FOUND *************')
        raise Exception('Error: Graph API token not found! Please re-generate your graph API token!')

    with open('ms_graph_api_token.json', 'w') as _f:
        _f.write(access_token_cache.serialize())

    return token_response['access_token']
