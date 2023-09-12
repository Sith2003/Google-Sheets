const express = require('express');
const { GoogleSpreadsheet } = require('google-spreadsheet')
const { google } = require('googleapis')

const app = express();
const port = process.env.PORT || 3001
app.use(express.json())

async function googleAuth() {
    const auth = new google.auth.GoogleAuth({
      keyFile: "credentials.json",
      scopes: ["https://www.googleapis.com/auth/spreadsheets"],
    });
    const authClient = await auth.getClient();
    const sheets = google.sheets({ version: "v4", auth: authClient });
    return {
      auth,
      authClient,
      sheets,
    };
}

app.get('/add-text', async(req, res) => {  
    try {
        const { sheets } = await googleAuth()
        const doc = new GoogleSpreadsheet('1CHAuWnmLT7DFhSBQH0EGrrIkF-tWVV3uwU6HHF70hH0' , client)
        await doc.loadInfo()
        await doc.updateProperties({ title: 'Google Spreadsheet 101'})
    
        const sheet = doc.sheetsByIndex[0]
        const headers = ['Name', 'Age']
    
        await sheet.setHeaderRow(headers)
        let dataArray = [
            { 'Name': "Test", 'Age': 20 },
            { 'Name': "Test2", 'Age': 21}
        ]
        await sheet.addRows(dataArray)
    
        res.status(200).json({ message: 'Successfully' })
    } catch (error) {
        res.status(500).json({ message: error.message })
    }
})
app.get("/get-sheets", async (req, res) => {
    const { sheets } = await googleAuth();
  
    // Read rows from spreadsheet
    const getRows = await sheets.spreadsheets.values.get({
      spreadsheetId: "1M1uVRHMCK_7NZLZpZ7HJRV84KWI2bCRPcltW_EZKCeQ",
      range: "ສັງລວມ",
    });
  
    res.send(getRows.data);
  });



app.post('/create-sheet', async (req, res) => {
    try {
        const { sheets } = await googleAuth()

        // Load the source and destination Google Sheets by their IDs
        const sourceSheetId = '1M1uVRHMCK_7NZLZpZ7HJRV84KWI2bCRPcltW_EZKCeQ';
        const destinationSheetId = '1uANZKaNqe0AS3nwom3xeAHgBOilkqw8ivL22aaNMtwM';
        const sourceSheetName = 'ສັງລວມ';
        
        const sheetsMetadata = await sheets.spreadsheets.get({
            spreadsheetId: sourceSheetId,
            fields: 'sheets.properties',
            // range: sourceSheetName,
            // valueRenderOption: 'FORMATTED_VALUE'
        });
        const sourceSheet = sheetsMetadata.data.sheets.find(sheet => sheet.properties.title === sourceSheetName);
        if (!sourceSheet) throw new Error('Source sheet not found.');
        // Copy the specific sheet to the destination document without data validation
        const copyRequest = {
            spreadsheetId: sourceSheetId,
            sheetId: sourceSheet.properties.sheetId,
            resource: {
                destinationSpreadsheetId: destinationSheetId,
            },
        };
        const response = (await sheets.spreadsheets.sheets.copyTo(copyRequest)).data;
        
        res.status(201).json({ response, message: 'Data copied successfully.' });
    } catch (error) {
        res.status(500).json({ message: error.message });
    }
});



app.listen(port, () => console.log(`Server listening on ${port}`));