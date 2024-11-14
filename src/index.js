const express = require("express");
const bodyParser = require("body-parser");
const { getSheets } = require("./googleSheetsHelper");
const path = require("path");
const cors = require("cors");
const moment = require("moment");

const app = express();
app.use(bodyParser.json());

// Enable CORS for all domains
app.use(cors());

// Middleware to parse JSON request bodies
app.use(express.json());

// Serve static HTML files from the public directory
app.use(express.static(path.join(__dirname, "../public")));

// Convert column index (1-based) to column name (AA, AB,...)
const columnNumberToLetter = (col) => {
  let letter = "";
  while (col > 0) {
    const remainder = (col - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
};

// Function to find column index based on column name
const findColumnIndex = (row, columnName) => row.indexOf(columnName) + 1;

// Find the last column with data
const getLastColumnWithData = (values) => {
  if (!values || values.length === 0) return -1; // No data

  let lastColumn = -1;

  for (let row of values) {
    if (row && row.length > 0) {
      lastColumn = Math.max(lastColumn, row.length);
    }
  }

  return lastColumn; // Return 1-based column index
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

      const maxColumns = sheet.properties.gridProperties.columnCount;
      const range = `${sheetName}!A1:${columnNumberToLetter(maxColumns)}`;
      const valuesResponse = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range,
      });

      const values = valuesResponse.data.values || [];
      const lastColumnWithData = getLastColumnWithData(values);

      console.log(`***Sheet: ${sheetName}`);
      console.log(`- Last column with data: ${lastColumnWithData}`);

      if (lastColumnWithData === maxColumns) {
        await sheets.spreadsheets.batchUpdate({
          spreadsheetId,
          resource: {
            requests: [
              {
                insertDimension: {
                  range: {
                    sheetId: sheet.properties.sheetId,
                    dimension: "COLUMNS",
                    startIndex: maxColumns,
                    endIndex: maxColumns + 1,
                  },
                  inheritFromBefore: true,
                },
              },
            ],
          },
        });
      }

      const headerRow = values[10]; // Header row (row 11 in 1-based index)
      const dataRows = values.slice(11); // Data from row 12

      // Add today's date (dd/mm) to the header row
      const today = moment().format("DD/MM");
      headerRow[lastColumnWithData + 1] = today;
      const newColumnLetter = columnNumberToLetter(lastColumnWithData + 1);

      // Update header row in the sheet
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `${sheetName}!${newColumnLetter}11`,
        valueInputOption: "USER_ENTERED",
        resource: {
          values: [[today]],
        },
      });

      const rankingColIndex = findColumnIndex(headerRow, "Today Ranking");
      const urlToolColIndex = findColumnIndex(headerRow, "URL Tool");
      const checkURLTargetColIndex = findColumnIndex(
        headerRow,
        "Check URL Target"
      );

      if (
        rankingColIndex <= 0 ||
        urlToolColIndex <= 0 ||
        checkURLTargetColIndex <= 0
      ) {
        console.log(`- Columns not found: 
          +) RankingColumnIndex: ${rankingColIndex}
          +) urlToolColIndex: ${urlToolColIndex}
          +) checkURLTargetColumnIndex: ${checkURLTargetColIndex}`);
        continue;
      }

      const updates = [];
      const notes = [];
      const requests = []; // For batchUpdate notes

      for (let i = 0; i < dataRows.length; i++) {
        const row = dataRows[i];
        const rowIndex = i + 11; // Row index in sheet (1-based index)
        updates.push([row[rankingColIndex - 1]]);
        if (row[checkURLTargetColIndex - 1] === "NO") {
          const note = `Đọc sai URL ${row[urlToolColIndex - 1]}`;
          notes.push(note);

          // Add a request to add a note
          requests.push({
            updateCells: {
              range: {
                sheetId: sheet.properties.sheetId,
                startRowIndex: rowIndex,
                endRowIndex: rowIndex + 1,
                startColumnIndex: lastColumnWithData, // updateColumn (0-based index)
                endColumnIndex: lastColumnWithData + 1,
              },
              rows: [
                {
                  values: [
                    {
                      note: note,
                    },
                  ],
                },
              ],
              fields: "note",
            },
          });
        } else {
          notes.push(null);
        }
      }

      const updateRange = `${sheetName}!${newColumnLetter}12:${newColumnLetter}${
        11 + updates.length
      }`;

      console.log(`- Update range: ${updateRange}`);
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: updateRange,
        valueInputOption: "USER_ENTERED",
        resource: {
          values: updates.map((update) => [update[0]]),
        },
      });

      // Batch update notes
      if (requests.length > 0) {
        await sheets.spreadsheets.batchUpdate({
          spreadsheetId,
          resource: {
            requests,
          },
        });
        console.log("- Notes added successfully.");
      } else {
        console.log("- No notes to add.");
      }

      // Remove existing filters
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        resource: {
          requests: [
            {
              clearBasicFilter: {
                sheetId: sheet.properties.sheetId,
              },
            },
          ],
        },
      });

      // Add a new filter from row 11 downwards
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        resource: {
          requests: [
            {
              setBasicFilter: {
                filter: {
                  range: {
                    sheetId: sheet.properties.sheetId,
                    startRowIndex: 10, // Row 11 (0-based index)
                    startColumnIndex: 0, // Column A
                    endColumnIndex: lastColumnWithData + 1, // Include new column
                  },
                },
              },
            },
          ],
        },
      });
    }

    res.send({ message: "Update completed with notes and filter." });
  } catch (error) {
    console.error(error);
    res.status(500).send({
      message: `Error updating sheet: ${error}`,
    });
  }
});

// Update Performance
app.post("/update-performance", async (req, res) => {
  const { sheetNames, spreadsheetId } = req.body;

  if (!sheetNames || !spreadsheetId) {
    return res.status(400).send({ message: "Missing required fields" });
  }

  try {
    const sheets = await getSheets();

    for (const sheetName of sheetNames) {
      const sheetData = await sheets.spreadsheets.get({
        spreadsheetId,
      });

      const sheet = sheetData.data.sheets.find(
        (sheet) => sheet.properties.title === sheetName
      );

      const maxColumns = sheet.properties.gridProperties.columnCount;
      const maxRows = sheet.properties.gridProperties.rowCount;
      const range = `${sheetName}!A1:${columnNumberToLetter(maxColumns)}`;
      const valuesResponse = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range,
      });

      const values = valuesResponse.data.values || [];
      const headerRow = values[10]; // Row 11
      const dataRows = values.slice(11); // Rows 12 onwards
      const today = moment().format("DD/MM");

      // Find the column with today's date
      const updateColumnIndex = headerRow.indexOf(today) + 1; // 1-based index
      if (updateColumnIndex <= 0) {
        return res.status(400).send({
          message: `Today's date (${today}) not found in header row.`,
        });
      }

      const updateColumnLetter = columnNumberToLetter(updateColumnIndex);

      // Clear previous values in updateColumn (row 1 to 10)
      const clearRange = `${sheetName}!${updateColumnLetter}1:${updateColumnLetter}10`;
      await sheets.spreadsheets.values.clear({
        spreadsheetId,
        range: clearRange,
      });

      // Count rows for each condition
      const counts = {
        "<=3": 0,
        "4-5": 0,
        "6-10": 0,
        "11-20": 0,
        "21-50": 0,
        "51-100": 0,
        "n/a": 0,
        "Check URL Target = NO": 0,
      };

      const checkURLTargetColIndex = findColumnIndex(
        headerRow,
        "Check URL Target"
      );
      console.log(`***Sheet: ${sheetName}`);
      console.log(`- checkURLTargetColIndex: ${checkURLTargetColIndex}`);
      console.log(`- maxRows: ${maxRows}`);

      for (const row of dataRows) {
        const value = row[updateColumnIndex - 1]; // Value in updateColumn
        if (value === "n/a") {
          counts["n/a"]++;
        } else if (!isNaN(value)) {
          const numericValue = parseFloat(value);
          if (numericValue <= 3) counts["<=3"]++;
          else if (numericValue >= 4 && numericValue <= 5) counts["4-5"]++;
          else if (numericValue >= 6 && numericValue <= 10) counts["6-10"]++;
          else if (numericValue >= 11 && numericValue <= 20) counts["11-20"]++;
          else if (numericValue >= 21 && numericValue <= 50) counts["21-50"]++;
          else if (numericValue >= 51 && numericValue <= 100)
            counts["51-100"]++;
        }
        if (
          checkURLTargetColIndex > 0 &&
          row[checkURLTargetColIndex - 1] === "NO"
        ) {
          counts["Check URL Target = NO"]++;
        }
      }

      // Calculate total for row 8
      const totalValue =
        counts["<=3"] +
        counts["4-5"] +
        counts["6-10"] +
        counts["11-20"] +
        counts["21-50"] +
        counts["51-100"] +
        counts["n/a"];

      // Prepare data to write to the sheet
      const performanceData = [
        ["KW (TĐ)"], // Row 1
        [counts["<=3"]], // Row 2
        [counts["4-5"]], // Row 3
        [counts["6-10"]], // Row 4
        [counts["11-20"]], // Row 5
        [counts["21-50"]], // Row 6
        [counts["51-100"]], // Row 7
        [counts["n/a"]], // Row 8
        [counts["Check URL Target = NO"]], // Row 9
        [totalValue], // Row 10
      ];

      const updateRange = `${sheetName}!${updateColumnLetter}1:${updateColumnLetter}10`;

      // Update sheet with performance data
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: updateRange,
        valueInputOption: "USER_ENTERED",
        resource: {
          values: performanceData,
        },
      });

      // Set borders for the updateColumn
      // await sheets.spreadsheets.batchUpdate({
      //   spreadsheetId,
      //   resource: {
      //     requests: [
      //       {
      //         updateBorders: {
      //           range: {
      //             sheetId: sheet.properties.sheetId,
      //             startRowIndex: 0, // Row 1 (0-based index)
      //             endRowIndex: maxRows, // Last row
      //             startColumnIndex: updateColumnIndex - 1, // Update column (0-based index)
      //             endColumnIndex: updateColumnIndex, // Next column
      //           },
      //           top: {
      //             style: "SOLID",
      //             width: 1,
      //             color: { red: 0, green: 0, blue: 0 },
      //           },
      //           bottom: {
      //             style: "SOLID",
      //             width: 1,
      //             color: { red: 0, green: 0, blue: 0 },
      //           },
      //           left: {
      //             style: "SOLID",
      //             width: 1,
      //             color: { red: 0, green: 0, blue: 0 },
      //           },
      //           right: {
      //             style: "SOLID",
      //             width: 1,
      //             color: { red: 0, green: 0, blue: 0 },
      //           },
      //         },
      //       },
      //     ],
      //   },
      // });
    }

    res.send({ message: "Performance data updated successfully." });
  } catch (error) {
    console.error(error);
    res.status(500).send({
      message: `Error updating performance data: ${error.message}`,
    });
  }
});

// Start Server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});
