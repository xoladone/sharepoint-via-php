# Sharepoint-via-php

Short description: Just a small reference how to reach sharepoint online via php

This project is a small reference implementation to access sharepoint online via the MS Graph API V1.0.

## Preliminary work
  - you habe to register an app on https://portal.azure.com and give the needed permissions in my example: (MS Graph, Sites.Manage.All, Application. Be sure that the status of the permissions is granted. If not, you must contact your ms admin. Create a client secret. Note down client ID, client secret, hostname of your sharepoint (eg: yourcompany.sharepoint.com).
  - Go to submenu Properties of Azure Active Directory and note down the Directory ID.

## Implemented till now:
  - Get an access-token for further actions.
  - Some easys site actions to get ids and list items.
  
## Help
You have to change some constants. You get the values from the Azure Application Registration. Have a look at the upper paragraph "Preliminary work"
   - SP_CLIENTID     = 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx';
   - SP_CLIENTSECRET = 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx';
   - SP_DIRECTORYID  = 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx';
   - SP_HOSTNAME     = 'yourcompany.sharepoint.com';
