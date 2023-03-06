# ms-graph-app

The ms-graph-app allows the user to create text files, share them with the desired people (Microsoft account holders) and view the files created with the application. To use the application, you must have a Microsoft account.

# To run locally

### Deploy Microsoft Graph Toolkit

- Set up your Microsoft 365 tenant: https://learn.microsoft.com/en-us/graph/toolkit/get-started/overview?tabs=html (Note! Select "Instant sandbox")
- Create an Azure Active Directory app, instructions at the following link: https://learn.microsoft.com/en-us/graph/toolkit/get-started/add-aad-app-registration (Note! Use the account created above)
- In the Azure Portal, go to your application registration.
  Verify that you are on the Overview page.
  From the Essentials section, copy the values of the Application (client) ID and Directory (tenant) ID properties
- Create an .env file at the root of the folder and add the IDs to that file as follows:

```
REACT_APP_CLIENT_ID=ADD_YOUR_OWN_CLIENT_ID
REACT_APP_AUTHORITY=https://login.microsoftonline.com/ADD_YOUR_OWN_TENANT_ID
```

- `npm i`
- `npm start`
- Open the app on http://localhost:3000/
