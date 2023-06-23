const express = require("express");
const multer = require("multer");
const excelToJson = require("convert-excel-to-json");

const fs = require("fs-extra");
const path = require("path");

const app = express();
const port = 8000;
const jsonFolder = "json-data"; // Folder to store JSON files

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
        raw: false,
      });

      const transformedData = transformExcelData(excelData);

      // Generate a unique filename for the JSON file
      const timestamp = Date.now();
      const jsonFilename = `data_${timestamp}.json`;
      const jsonFilePath = path.join(jsonFolder, jsonFilename);

      fs.ensureDirSync(jsonFolder); // Create the JSON folder if it doesn't exist

      fs.writeFile(jsonFilePath, JSON.stringify(transformedData), (err) => {
        if (err) {
          console.error("Error writing JSON file:", err);
          res.status(500).json({ error: "Internal Server Error" });
        } else {
          fs.remove(filepath, (err) => {
            if (err) {
              console.error("Error removing file:", err);
            }
          });

          res.status(200).json({ filename: jsonFilename });
        }
      });
    }
  } catch (error) {
    res.status(500).json({ error: "Internal Server Error" });
  }
});

app.listen(port, () => {
  console.log(`Node.js app listening on PORT ${port}`);
});

function transformExcelData(excelData) {
  const sheets = Object.keys(excelData);
  const transformedData = [];

  for (const sheet of sheets) {
    const rows = excelData[sheet];

    for (const row of rows) {
      const entry = {
        studentId: row["studentId"],
        courseName: row["courseName"],
        date: new Date(row["date"]),
        week: parseInt(row["week"]),
        isPresent: row["isPresent"].toUpperCase() === "TRUE",
        scores: {
          aptitude: {
            maxScore: parseInt(row["Scores - Aptitude (Max)"]),
            minScore: parseInt(row["Scores - Aptitude (Min)"]),
            score: parseInt(row["Scores - Aptitude"]),
          },
          verbal: {
            maxScore: parseInt(row["Scores - Verbal (Max)"]),
            minScore: parseInt(row["Scores - Verbal (Min)"]),
            score: parseInt(row["Scores - Verbal"]),
          },
          technical: {
            maxScore: parseInt(row["Scores - Technical (Max)"]),
            minScore: parseInt(row["Scores - Technical (Min)"]),
            score: parseInt(row["Scores - Technical"]),
          },
        },
      };

      transformedData.push(entry);
    }
  }

  return transformedData;
}
