const express = require("express");
const bodyParser = require("body-parser");
const { getSheets } = require("./googleSheetsHelper");
const path = require("path");
const cors = require("cors");

const app = express();
app.use(bodyParser.json());

// Enable CORS for all domains
app.use(cors());

// Middleware to parse JSON request bodies
app.use(express.json());

// Serve static HTML files from the public directory
app.use(express.static(path.join(__dirname, "../public")));

// Function to find column index based on column name
const findColumnIndex = (row, columnName) => row.indexOf(columnName);

const getLastColumnWithData = (values) => {
  for (let i = values[0].length - 1; i >= 0; i--) {
    if (values[0][i]) {
      return i + 1;
    }
  }
  return -1;
};

// Update Rankings and Notes
app.post("/update-ranking-and-notes", async (req, res) => {
  const { sheetNames, spreadsheetId } = req.body;

  if (!sheetNames || !spreadsheetId) {
    return res.status(400).send({ message: "Missing required fields" });
  }

  try {
    const sheets = await getSheets();
    for (const sheetName of sheetNames) {
      // Get the sheet's metadata to get row and column count
      const sheetData = await sheets.spreadsheets.get({
        spreadsheetId: spreadsheetId,
      });

      // Find the sheet object by name
      const sheet = sheetData.data.sheets.find(
        (sheet) => sheet.properties.title === sheetName
      );

      // Get row and column counts
      const maxRows = sheet.properties.gridProperties.rowCount;
      const maxColumns = sheet.properties.gridProperties.columnCount;
      console.log(`maxColumns: ${maxColumns}`);

      // Retrieve data from the sheet based on actual row/column limits
      const range = `${sheetName}!A1:${String.fromCharCode(
        64 + maxColumns
      )}${maxRows}`;
      const values = await sheets.spreadsheets.values.get({
        spreadsheetId: spreadsheetId,
        range: range,
      });

      const lastColumnWithData = getLastColumnWithData(values.data.values);
      console.log(`lastColumnWithData: ${lastColumnWithData}`);

      // Add a new column (for Ranking and Notes)
      if (lastColumnWithData == maxColumns) {
        await sheets.spreadsheets.batchUpdate({
          spreadsheetId: spreadsheetId,
          resource: {
            requests: [
              {
                insertDimension: {
                  range: {
                    sheetId: sheet.properties.sheetId,
                    dimension: "COLUMNS",
                    startIndex: lastColumnWithData, // Add a new column after the last column with data
                    endIndex: lastColumnWithData + 1,
                  },
                  inheritFromBefore: true,
                },
              },
            ],
          },
        });
      }

      const headerRow = values.data.values[10]; // Header row
      const dataRows = values.data.values.slice(11); // Data rows starting from row 12

      const rankingColIndex = findColumnIndex(headerRow, "Today Ranking");
      const urlToolColIndex = findColumnIndex(headerRow, "URL Tool");
      const checkURLTargetColIndex = findColumnIndex(
        headerRow,
        "Check URL Target"
      );

      const updates = [];
      const notes = [];

      for (const row of dataRows) {
        updates.push([row[rankingColIndex]]);
        notes.push(
          row[checkURLTargetColIndex] === "NO"
            ? `Đọc sai URL ${row[urlToolColIndex]}`
            : null
        );
      }

      // Calculate the range dynamically based on the number of rows that were updateds
      const updateRange = `${sheetName}!${String.fromCharCode(
        64 + lastColumnWithData + 1
      )}12:${String.fromCharCode(64 + lastColumnWithData + 1)}${
        11 + updates.length
      }`;
      console.log(`updateRange: ${updateRange}`);

      // Write updates back to the sheet (Ranking and Notes column combined)
      await sheets.spreadsheets.values.update({
        spreadsheetId: spreadsheetId,
        range: updateRange,
        valueInputOption: "USER_ENTERED",
        resource: {
          values: updates.map((update, index) => [update[0], notes[index]]),
        },
      });
    }

    res.send({ message: "Update completed" });
  } catch (error) {
    console.error(error);
    res.status(500).send({
      message: `Error updating sheet: ${error}`,
    });
  }
});

// Start Server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});
