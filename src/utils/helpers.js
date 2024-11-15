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

module.exports = {
  columnNumberToLetter,
  findColumnIndex,
  getLastColumnWithData,
};
