const express = require("express");
const { updateNotesHandler } = require("../controllers/notesController");

const router = express.Router();

router.post("/update-notes", updateNotesHandler);

module.exports = router;
