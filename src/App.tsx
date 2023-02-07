import { Providers, ProviderState } from "@microsoft/mgt-element";
import { Agenda, Login } from "@microsoft/mgt-react";

import React, { useState, useEffect } from "react";
import "./App.css";
import FileView from "./components/FileView";

import FormFilling from "./components/FormFilling";

function useIsSignedIn(): [boolean] {
  const [isSignedIn, setIsSignedIn] = useState(false);

  useEffect(() => {
    const updateState = () => {
      const provider = Providers.globalProvider;
      setIsSignedIn(provider && provider.state === ProviderState.SignedIn);
    };

    Providers.onProviderUpdated(updateState);
    updateState();

    return () => {
      Providers.removeProviderUpdatedListener(updateState);
    };
  }, []);

  return [isSignedIn];
}

/** test function, useful */
async function test() {
  let provider = Providers.globalProvider;
  if (provider) {
    let graphClient = provider.graph.client;
    let userDetails = await graphClient.api("/me/").get();
    console.log(userDetails, "user info");

    await graphClient
      .api("/me/drive/root/search(q='student-app')")
      .get()
      .then((result) => {
        console.log(result.value[0], "how many files?");
      });
  }
}
// --------------------

function App() {
  const [isSignedIn] = useIsSignedIn();

  if (isSignedIn) {
    // test()
  }

  return (
    <div className="App">
      <header>
        <Login />
      </header>
      <div>
        {isSignedIn && (
          <div>
            <FileView />
            {/* <Agenda /> */}

            <FormFilling />
          </div>
        )}
      </div>
    </div>
  );
}

export default App;
