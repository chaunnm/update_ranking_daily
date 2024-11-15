const express = require("express");
const { updateRankingHandler } = require("../controllers/rankingController");

const router = express.Router();

router.post("/update-ranking", updateRankingHandler);

module.exports = router;
