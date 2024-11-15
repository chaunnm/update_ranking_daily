const { google } = require("googleapis");
// const fs = require("fs");
const credentials = JSON.parse(process.env.GOOGLE_SHEETS_CREDENTIALS);

const auth = new google.auth.GoogleAuth({
  credentials,
  scopes: ["https://www.googleapis.com/auth/spreadsheets"],
});

const getSheets = async () => {
  const authClient = await auth.getClient();
  return google.sheets({ version: "v4", auth: authClient });
};

module.exports = { getSheets };
