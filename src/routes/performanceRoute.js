const express = require("express");
const {
  updatePerformanceHandler,
} = require("../controllers/performanceController");

const router = express.Router();

router.post("/update-performance", updatePerformanceHandler);

module.exports = router;
