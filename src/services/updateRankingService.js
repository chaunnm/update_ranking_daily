const updateRanking = async (sheets, spreadsheetId, sheetName) => {
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
