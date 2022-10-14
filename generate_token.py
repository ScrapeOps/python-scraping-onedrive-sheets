from urllib import request
import webbrowser
from datetime import datetime
import json
import os
import msal
import sys

#####
# To generate the token: 
# 1. First run python generate_token.py 'YOUR_APP_ID_HERE' 'YOUR_SCOPE_HERE'
#    Example: python generate_token.py '9cef1616-a875-4596-4b2b-20cf9fa065d' 'files.readwrite.all'
# 2. Then insert the user_code displayed in your terminal into the page that opens in your browser
# 3. Next give access to the app to your onedrive account by following the instructions displayed
# 4. An access token will be generated and saved locally called ms_graph_api_token
####

app_id = sys.argv[1] 
scopes = [sys.argv[2]] 

# Save Session Token as a token file
access_token_cache = msal.SerializableTokenCache()

# read the token file
if os.path.exists('ms_graph_api_token.json'):
    access_token_cache.deserialize(open("ms_graph_api_token.json", "r").read())
    token_detail = json.load(open('ms_graph_api_token.json',))
    token_detail_key = list(token_detail['AccessToken'].keys())[0]
    token_expiration = datetime.fromtimestamp(int(token_detail['AccessToken'][token_detail_key]['expires_on']))
    print('\nTOKEN EXPIRY: ')
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
    # authenticate your account as usual
    flow = client.initiate_device_flow(scopes=scopes)
    print('user_code: ' + flow['user_code'])
    webbrowser.open('https://microsoft.com/devicelogin')
    token_response = client.acquire_token_by_device_flow(flow)

with open('ms_graph_api_token.json', 'w') as _f:
    _f.write(access_token_cache.serialize())

print('\nGraph API token:')
print(token_response['access_token'])
