# Python Scraping OneDrive Sheets
We have tried to make it easy for anyone scraping with Python to authenticate with 
and then save data to their OneDrive Excel Sheets 

# Getting Started
pip install this module


# Authentication
Authentication works via OAuth

## Create your "app" in Azure
You first need to create an app in your Microsoft azure portal - don't worry, its free!

## Generate your token
The first time your generate your token you will be taken to OneDrive to accept the scope and permissions that your app will need to communicate with your OneDrive account. This token will then be saved locally. 

## Interact with OneDrive - Add/Remove from your excel sheet etc
Now that you have the token locally available to your python code you can 
run any of the available functions we have or create your own by looking at the microsoft graph API. 

# Examples


## Add a table 


## Delete a table 


## Add row to a table 


## Delete a row from a table  



# Other resources
https://learn.microsoft.com/en-us/graph/api/table-post-rows?view=graph-rest-1.0&tabs=http