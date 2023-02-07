import React from "react";
import ReactDOM from "react-dom";
import App from "./App";
import "./index.css";
import { Providers } from "@microsoft/mgt-element";
import { Msal2Provider } from "@microsoft/mgt-msal2-provider";

let client_id: string = process.env.REACT_APP_CLIENT_ID || "";
let authority: string = process.env.REACT_APP_AUTHORITY || "";

Providers.globalProvider = new Msal2Provider({
  // replace the value in the following line with: CLIENT_ID
  clientId: client_id,
  // replace the value in the following line with: https://login.microsoftonline.com/ADD_YOUR_OWN_TENANT_ID
  authority: authority,
  scopes: [
    // "calendars.read",
    // "user.read",
    // "openid",
    // "profile",
    // "people.read",
    // "user.readbasic.all",
    "files.read",
    "Files.ReadWrite",
    "Files.ReadWrite.All",
  ],
});

ReactDOM.render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
  document.getElementById("root")
);
