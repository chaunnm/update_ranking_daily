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

const splitIntoBatches = (array, batchSize) => {
  const batches = [];
  for (let i = 0; i < array.length; i += batchSize) {
    batches.push(array.slice(i, i + batchSize));
  }
  return batches;
};

const processBatch = async (sheets, spreadsheetId, batch, updateFunc) => {
  for (const sheetName of batch) {
    await updateFunc(sheets, spreadsheetId, sheetName);
  }
};

const updateRankingAndNotes = async (sheets, spreadsheetId, sheetName) => {
  try {
    const sheetData = await sheets.spreadsheets.get({
      spreadsheetId,
    });

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

    const today = moment().format("DD/MM");
    headerRow[lastColumnWithData + 1] = today;
    const newColumnLetter = columnNumberToLetter(lastColumnWithData + 1);

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

    const updates = [];
    const notes = [];
    const requests = [];

    for (let i = 0; i < dataRows.length; i++) {
      const row = dataRows[i];
      const rowIndex = i + 11; // Row index in sheet (1-based index)
      updates.push([row[rankingColIndex - 1]]);
      if (row[checkURLTargetColIndex - 1] === "NO") {
        const note = `Đọc sai URL ${row[urlToolColIndex - 1]}`;
        notes.push(note);

        requests.push({
          updateCells: {
            range: {
              sheetId: sheet.properties.sheetId,
              startRowIndex: rowIndex,
              endRowIndex: rowIndex + 1,
              startColumnIndex: lastColumnWithData,
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

    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: updateRange,
      valueInputOption: "USER_ENTERED",
      resource: {
        values: updates.map((update) => [update[0]]),
      },
    });

    if (requests.length > 0) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        resource: {
          requests,
        },
      });
    }
  } catch (error) {
    console.error(
      `Error in updateRankingAndNotes for sheet ${sheetName}:`,
      error.message
    );
    throw error;
  }
};

const updatePerformance = async (sheets, spreadsheetId, sheetName) => {
  try {
    const sheetData = await sheets.spreadsheets.get({
      spreadsheetId,
    });

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
    const headerRow = values[10]; // Row 11
    const dataRows = values.slice(11); // Rows 12 onwards
    const today = moment().format("DD/MM");

    const updateColumnIndex = headerRow.indexOf(today) + 1;
    const updateColumnLetter = columnNumberToLetter(updateColumnIndex);

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

    const checkURLTargetColIndex = headerRow.indexOf("Check URL Target") + 1;

    for (const row of dataRows) {
      const value = row[updateColumnIndex - 1];
      if (value === "n/a") {
        counts["n/a"]++;
      } else if (!isNaN(value)) {
        const numericValue = parseFloat(value);
        if (numericValue <= 3) counts["<=3"]++;
        else if (numericValue >= 4 && numericValue <= 5) counts["4-5"]++;
        else if (numericValue >= 6 && numericValue <= 10) counts["6-10"]++;
        else if (numericValue >= 11 && numericValue <= 20) counts["11-20"]++;
        else if (numericValue >= 21 && numericValue <= 50) counts["21-50"]++;
        else if (numericValue >= 51 && numericValue <= 100) counts["51-100"]++;
      }
      if (
        checkURLTargetColIndex > 0 &&
        row[checkURLTargetColIndex - 1] === "NO"
      ) {
        counts["Check URL Target = NO"]++;
      }
    }

    const performanceData = [
      ["KW (TĐ)"],
      [counts["<=3"]],
      [counts["4-5"]],
      [counts["6-10"]],
      [counts["11-20"]],
      [counts["21-50"]],
      [counts["51-100"]],
      [counts["n/a"]],
      [counts["Check URL Target = NO"]],
      [
        counts["<=3"] +
          counts["4-5"] +
          counts["6-10"] +
          counts["11-20"] +
          counts["21-50"] +
          counts["51-100"] +
          counts["n/a"],
      ],
    ];

    const updateRange = `${sheetName}!${updateColumnLetter}1:${updateColumnLetter}10`;

    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: updateRange,
      valueInputOption: "USER_ENTERED",
      resource: {
        values: performanceData,
      },
    });
  } catch (error) {
    console.error(
      `Error in updatePerformance for sheet ${sheetName}:`,
      error.message
    );
    throw error;
  }
};

// Update Rankings and Notes
app.post("/update-ranking-and-notes", async (req, res) => {
  const startTime = Date.now();

  // const { sheetNames, spreadsheetId } = req.body;
  const { sheetName, spreadsheetId } = req.body;

  // if (!sheetNames || !spreadsheetId) {
  if (!sheetName || !spreadsheetId) {
    return res.status(400).send({ message: "Missing required fields" });
  }

  try {
    const sheets = await getSheets();
    // const batchSize = 2;
    // const batches = splitIntoBatches(sheetNames, batchSize);

    // for (const batch of batches) {
    // await processBatch(sheets, spreadsheetId, batch, updateRankingAndNotes);
    // }
    await updateRankingAndNotes(sheets, spreadsheetId, sheetName);

    const endTime = Date.now(); // Ghi lại thời gian kết thúc
    const duration = (endTime - startTime) / 1000; // Tính thời gian thực thi (giây)

    res.send({
      message: `Ranking and notes updated successfully, ${duration} seconds`,
    });
  } catch (error) {
    console.error(error);
    const endTime = Date.now(); // Ghi lại thời gian kết thúc
    const duration = (endTime - startTime) / 1000; // Tính thời gian thực thi (giây)
    res.status(500).send({
      message: `Error updating ranking and notes: ${error.message} - ${duration} seconds`,
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
    const batchSize = 2; // Số lượng sheet xử lý mỗi batch
    const batches = splitIntoBatches(sheetNames, batchSize);

    for (const batch of batches) {
      await processBatch(sheets, spreadsheetId, batch, updatePerformance);
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
