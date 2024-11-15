const { splitIntoBatches, processBatch } = require("../utils/batching");
const { getSheets } = require("../services/googleSheetsService");
const updatePerformance = require("../services/updatePerformanceService");

const updatePerformanceHandler = async (req, res) => {
  const { sheetNames, spreadsheetId } = req.body;

  if (!sheetNames || !spreadsheetId) {
    return res.status(400).send({ message: "Missing required fields" });
  }

  try {
    const sheets = await getSheets();
    const batchSize = 2;
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
};

module.exports = { updatePerformanceHandler };
