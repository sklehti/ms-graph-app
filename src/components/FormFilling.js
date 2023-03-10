import React, { useState, useRef, useEffect } from "react";
import { Providers } from "@microsoft/mgt-element";
import { prepScopes } from "@microsoft/mgt-element";
import { PeoplePicker } from "@microsoft/mgt-react";

function FormFilling() {
  const [formText, setFormText] = useState("");
  const [peopleEmails, setPeopleEmails] = useState([]);
  const [userInfo, setUserInfo] = useState(null);
  const [folderInfo, setFolderInfo] = useState({});
  const [fileName, setFileName] = useState("");

  const refPeople = useRef(null);

  useEffect(() => {
    getUserInfo();
  }, []);

  const handleText = (e) => {
    setFormText(e.target.value);
  };

  const handleFileName = (e) => {
    setFileName(e.target.value);
  };

  /**
   * Functio get user informations
   */
  const getUserInfo = async () => {
    let provider = Providers.globalProvider;
    if (provider) {
      let graphClient = provider.graph.client;
      let userDetails = await graphClient.api("/me").get();
      setUserInfo(userDetails);

      const children = await graphClient.api("/me/drive/root/children").get();
      let rightFolder = children.value.filter(
        (file) => file.name === "Sovellukset"
      );

      if (rightFolder.length > 0) {
        setFolderInfo(rightFolder[0].webUrl);
      }
    }
  };

  /**
   * open your OneDrive files
   */
  const handleOpenFileLink = async (e) => {
    window.open(folderInfo, "_blank");
  };

  /**
   * Search and display selected users
   * @param {users} e
   */
  const handlePeople = () => {
    const people = refPeople.current.selectedPeople;

    people.map((p) => {
      setPeopleEmails([...peopleEmails, { email: p.userPrincipalName }]);
    });
  };

  const permission = {
    recipients: peopleEmails,
    message: "Here's the file that we're collaborating on.",
    requireSignIn: true,
    sendInvitation: true,
    roles: ["write"],
    // TODO: add password
    //password: "password123",
    // TODO: remove hardcoding
    expirationDateTime: "2023-07-15T14:00:00.000Z",
  };

  /**
   * Creates a new folder and a file with text inside it. In addition, share the selected folder with the desired users and  send the share link
   */
  async function handleForm(e) {
    let provider = Providers.globalProvider;
    if (provider) {
      let graphClient = provider.graph.client;

      let searchFile = await graphClient
        .api(
          `/me/drive/root/search(q='${fileName}_${userInfo.userPrincipalName}')?select=name`
        )
        .get();

      if (searchFile.value.length === 0 && userInfo !== null) {
        let children = await graphClient
          .api(
            `/me/drive/special/approot/children/${fileName}_${userInfo.userPrincipalName}.txt/content`
          )
          .middlewareOptions(
            prepScopes("APIConnectors.Read.All", "APIConnectors.ReadWrite.All")
          )
          .put(formText);

        if (peopleEmails.length > 0) {
          await graphClient
            .api(`/me/drive/items/${children.id}/invite`)
            .post(permission);
        }

        setFormText("");
        setFileName("");
        setPeopleEmails([]);

        refPeople.current.selectedPeople = [];
        console.log("Tiedosto on luotu ja t??ytetty.");

        getUserInfo();
      } else {
        console.log(
          "Saman niminen tiedosto l??ytyy jo. Vaihda tiedoston nime??."
        );
      }
    }
  }

  return (
    <div>
      <h2>Luo tiedosto</h2>

      <div>
        <label>Tiedoston nimi:</label>
        <input
          placeholder="Tiedoston nimi"
          value={fileName}
          onChange={handleFileName}
        />
      </div>

      <br />
      <textarea
        value={formText}
        onChange={handleText}
        rows="10"
        style={{ width: "99%" }}
        placeholder="Teksti??..."
      ></textarea>
      <h2>Valitse henkil??t, joille haluat jakaa tiedoston</h2>
      <PeoplePicker ref={refPeople} selectionChanged={handlePeople} />
      <div>
        <br />
        <button onClick={handleForm}>Talleta ja jaa tiedosto</button>

        {folderInfo.length > 0 ? (
          <div>
            <h2>T??ss?? sovelluksessa tallentamiasi tiedostoja</h2>
            <button onClick={handleOpenFileLink}>
              Tarkastele tiedostojasi
            </button>
          </div>
        ) : (
          <div></div>
        )}
        <div></div>
      </div>
    </div>
  );
}

export default FormFilling;
