const express = require("express");
const multer = require("multer");
const excelToJson = require("convert-excel-to-json");
const fs = require("fs-extra");

const app = express();
const port = 3000;

var upload = multer({ dest: "uploads/" });

app.post("/read", upload.single("file"), (req, res) => {
  try {
    if (!req.file || !req.file.filename) {
      res.status(400).json("No File");
    } else {
      var filepath = "uploads/" + req.file.filename;

      const excelData = excelToJson({
        sourceFile: filepath,
        header: {
          rows: 1,
        },
        columnToKey: {
          "*": "{{columnHeader}}",
        },
      });

      fs.remove(filepath, (err) => {
        if (err) {
          console.error("Error removing file:", err);
        }
      });

      res.status(200).json(excelData);
    }
  } catch (error) {
    res.status(500).json({ error: "Internal Server Error" });
  }
});

app.listen(port, () => {
  console.log(`Node.js app listening on PORT ${port}`);
});
