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

module.exports = { splitIntoBatches, processBatch };
