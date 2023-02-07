# ms-graph-app

Coming later...

# To run locally

- `npm i`
- Create an Azure Active Directory app, instructions at the following link: https://learn.microsoft.com/en-us/graph/toolkit/get-started/add-aad-app-registration
- In the Azure Portal, go to your application registration.
  Verify that you are on the Overview page.
  From the Essentials section, copy the values of the Application (client) ID and Directory (tenant) ID properties and add them to a file called /src/index.tsx. Note! Do not add them directly into the index.txs file. Create an .env-file for them to the root of the ms-graph-app -folder and add the client_id and authority value there.
- `npm start`
- Open the app on http://localhost:3000/
