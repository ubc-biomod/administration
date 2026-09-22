# NOTES: replace rangs U2:U100 and V2:V100 as necessary
# ALSO FOR ACTIVE SPREADSHEET MAKE SURE YOU ATTACH THE CORRECT SPREADSHEET

function getPDFLinks() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const folderId = "YOUR_FOLDER_ID";

  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();

  const fileMap = [];

  while (files.hasNext()) {
    const file = files.next();
    fileMap.push({
      name: file.getName(),
      url: file.getUrl()
    });
  }

  const ids = sheet.getRange("U2:U100").getValues();

  const output = ids.map(([id]) => {
    if (!id) return [""];

    const match = fileMap.find(file => file.name.includes(id));

    return [match ? match.url : ""];
  });

  sheet.getRange("V2:V100").setValues(output);
}
