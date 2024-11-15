const express = require("express");
const bodyParser = require("body-parser");
const cors = require("cors");
const path = require("path");

const rankingRoute = require("./routes/rankingRoute");
const notesRoute = require("./routes/notesRoute");
const performanceRoute = require("./routes/performanceRoute");

const app = express();
app.use(cors());
app.use(bodyParser.json());
app.use(express.json());
app.use(express.static(path.join(__dirname, "../public")));

// Register routes
app.use("/api", rankingRoute);
app.use("/api", notesRoute);
app.use("/api", performanceRoute);

module.exports = app;
