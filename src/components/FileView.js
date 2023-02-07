import React, { useState, useRef } from "react";
import { FileList, File } from "@microsoft/mgt-react";

function FileView() {
  const [selectedFile, setSelectedFile] = useState(null);

  const refFileItem = useRef(null);

  /**
   * Displays the contents of the selected folder
   * @param {selected folder } e
   */
  const handleFolder = (e) => {
    const allFolders = e.target.__files;
    const folderIndex = e.target._focusedItemIndex;

    if (
      allFolders[folderIndex].folder &&
      allFolders[folderIndex].id &&
      allFolders[folderIndex].folder.childCount > 0
    ) {
      const id = allFolders[folderIndex].id;
      refFileItem.current.itemId = id;
    }

    if (!allFolders[folderIndex].folder && allFolders[folderIndex].id) {
      setSelectedFile(allFolders[folderIndex]);
    }
  };

  const handleFile = (e) => {
    console.log(e, "test", selectedFile);
  };

  return (
    <div>
      {/* TODO: FileList displays all folders and files for the user */}
      {/* <FileList
        ref={refFileItem}
        onClick={handleFolder}
        // fileListQuery={`me/drive/items/{file_id_here}/children`}
      /> */}

      {selectedFile !== null ? (
        <div>
          <h2>Valittu asiakirja</h2>
          <div className="selected-file-style">
            <File
              onClick={handleFile}
              fileQuery={`/drives/${selectedFile.parentReference.driveId}/items/${selectedFile.id}`}
            ></File>
          </div>
        </div>
      ) : (
        <div></div>
      )}
    </div>
  );
}

export default FileView;
