const express = require('express');
const { google } = require('googleapis')
const fs = require('fs')
require('dotenv').config()

const app = express();
const port = process.env.PORT || 3001
app.use(express.json())

// Sheets Authentication
async function googleAuth() {
    const auth = new google.auth.GoogleAuth({
      keyFile: "credentials.json",
      scopes: ["https://www.googleapis.com/auth/spreadsheets"],
    });
    const authClient = await auth.getClient();
    return authClient
}

// Drive Authentication
const CLIENT_ID = process.env.CLIENT_ID
const CLIENT_SECRET = process.env.CLIENT_SECRET
const REDIRECT_URI = process.env.REDIRECT_URI

const oauth2Client = new google.auth.OAuth2(
    CLIENT_ID,
    CLIENT_SECRET,
    REDIRECT_URI
);
try {
    const creds = fs.readFileSync("token.json");
    oauth2Client.setCredentials(JSON.parse(creds));
} catch (error) {
    console.log("No token found");
}


app.get('/auth/google', (req, res) => {
    try {
        const url = oauth2Client.generateAuthUrl({
            access_type: 'offline',
            scope: [
              'https://www.googleapis.com/auth/userinfo.profile',
              'https://www.googleapis.com/auth/drive',
            ],
        });
        res.redirect(url);
    } catch (error) {
        return res.status(500).json({ message: error.message });
    }
});
  
app.get('/google/redirect', async (req, res) => {
    try {
      const { code } = req.query;
      const { tokens } = await oauth2Client.getToken(code);
      oauth2Client.setCredentials(tokens);
      fs.writeFileSync('token.json', JSON.stringify(tokens));
      res.send('Success');
    } catch (error) {
      return res.status(500).json({ message: error.message });
    }
});

app.get('/all-files', async (req, res) => {
    try {
        const drive = google.drive({ version: "v3", auth: oauth2Client });
        const { data } = await drive.files.list();
        res.json({ message: 'Successfully', files: data.files });
    } catch (error) {
      res.status(500).json({ error: error.message });
    }
});

app.get('/folder', async (req, res) => {
    try {
        const drive = google.drive({ version: "v3", auth: oauth2Client });
        const { folderId } = req.body;
        const { data } = await drive.files.list({
            q: `'${folderId}' in parents`, // Filter by the specified folder ID
        });
        res.json({ message: 'Successfully', files: data.files });
    } catch (error) {
      res.status(500).json({ error: error.message });
    }
});

app.get('/search', async (req, res) => {
    try {
        const drive = google.drive({ version: "v3", auth: oauth2Client });
        const { mimeTypeFilter } = req.body;
        const { data } = await drive.files.list({
            q: mimeTypeFilter
        });
        res.json({ message: 'Successfully', files: data.files });
    } catch (error) {
      res.status(500).json({ error: error.message });
    }
})

app.get('/sheets/:spreadsheetId', async (req, res) => {
    try {
      const { spreadsheetId } = req.params;
      const oauth2Client = await googleAuth();
      // Function to get all sheets within the specified spreadsheetId
      const getAllSheets = async (spreadsheetId, oauth2Client) => {
        const sheets = google.sheets({ version: 'v4', auth: oauth2Client });
        const { data } = await sheets.spreadsheets.get({
          spreadsheetId,
        });
        // Extract the sheets information from the response
        const sheetsInfo = data.sheets.map(sheet => {
          return {
            sheetId: sheet.properties.sheetId,
            sheetTitle: sheet.properties.title,
          };
        });
        return sheetsInfo;
      };
      const sheetsInfo = await getAllSheets(spreadsheetId, oauth2Client);
  
      res.json({ total: sheetsInfo.length, sheets: sheetsInfo});
    } catch (error) {
      console.error('Error:', error);
      res.status(500).json({ error: 'Internal Server Error' });
    }
});

app.post('/create-all-sheet', async (req, res) => {
    try {
        const auth = await googleAuth()
        const sheets = google.sheets({ version: "v4", auth: auth})
        // Load the source and destination Google Sheets by their IDs
        const { sourceSheetId, destinationSheetId } = req.body;

        const sourceSheets = await sheets.spreadsheets.get({
            spreadsheetId: sourceSheetId,
            fields: 'sheets(properties(title,sheetId))', // Retrieve sheet properties
        });
        // Iterate through each sheet and copy it to the destination spreadsheet
        for (const sourceSheet of sourceSheets.data.sheets) {
            const copyRequest = {
                spreadsheetId: sourceSheetId,
                sheetId: sourceSheet.properties.sheetId,
                resource: {
                    destinationSpreadsheetId: destinationSheetId,
                },
            };
            const sheetCopied = await sheets.spreadsheets.sheets.copyTo(copyRequest);

            const copiedSheetId = sheetCopied.data.sheetId;
            const space = ""
            const uniqueSheetTitle = `${sourceSheet.properties.title}${space}`;
            // Optionally, you can set the sheet title here to match the original sheet's title
            const updateSheetRequest = {
                spreadsheetId: destinationSheetId,
                resource: {
                    requests: [
                        {
                            updateSheetProperties: {
                                properties: {
                                    sheetId: copiedSheetId,
                                    title: uniqueSheetTitle, // Set the same title as the source sheet
                                },
                                fields: 'title',
                            },
                        },
                    ],
                },
            };
            await sheets.spreadsheets.batchUpdate(updateSheetRequest);
        }

        res.status(201).json({ message: 'All sheets copied successfully.' });
    } catch (error) {
        res.status(500).json({ message: error.message });
    }
});

app.post('/copy-sheet', async (req, res) => {
    try {
        const auth = await googleAuth()
        const sheets = google.sheets({ version: "v4", auth: auth})

        // Load the source and destination Google Sheets by their IDs
        const {
            sourceSheetId,
            destinationSheetId,
            sourceSheetName,
            gid
        } = req.body;
        // const sourceSheetName = 'ລາຍງານ 3. ຄ່າອາຫານ ແລະ ເດີນທາງ'; // your sheet name
        // const gid = 284768657 //your gid
        
        const copyRequest = {
            spreadsheetId: sourceSheetId,
            sheetId: gid,
            resource: {
                destinationSpreadsheetId: destinationSheetId,
            },
        };
        const sheetCopied = await sheets.spreadsheets.sheets.copyTo(copyRequest);

        const copiedSheetId = sheetCopied.data.sheetId;
        const updateSheetRequest = {
            spreadsheetId: destinationSheetId,
            resource: {
                requests: [
                    {
                        updateSheetProperties: {
                            properties: {
                                sheetId: copiedSheetId,
                                title: sourceSheetName, // Set the desired title here
                            },
                            fields: 'title',
                        },
                    },
                ],
            },
        };

        await sheets.spreadsheets.batchUpdate(updateSheetRequest);
        
        res.status(201).json({ message: 'Data copied successfully.' });
    } catch (error) {
        res.status(500).json({ message: error.message });
    }
});


app.listen(port, () => console.log(`Server listening on ${port}`));