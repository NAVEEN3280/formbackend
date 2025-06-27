import express from "express";
import cors from "cors";
import bodyParser from "body-parser";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import async from "async";
import ExcelJS from "exceljs";

// Setup path
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const FILE_PATH = path.join(__dirname, "waitlist.xlsx");

const app = express();
const PORT = 5000;

app.use(cors());
app.use(bodyParser.json());

// Create a write queue (concurrency = 1)
const queue = async.queue(async (task, done) => {
  try {
    await task();
  } finally {
    done();
  }
}, 1);

// Helper to check if file exists
const fileExists = (file) => {
  return fs.existsSync(file) && fs.statSync(file).size > 0;
};

// Submit route
app.post("/submit", (req, res) => {
  const { email, whatsapp, businessType, challenge } = req.body;

  console.log("📥 Submission received:", email);

  queue.push(async () => {
    const newRow = {
      email,
      whatsapp,
      businessType,
      challenge,
      timestamp: new Date().toLocaleString("en-IN", {
        timeZone: "Asia/Kolkata",
      }),
    };

    // If file doesn't exist, create and write headers + row
    if (!fileExists(FILE_PATH)) {
      const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
        filename: FILE_PATH,
      });
      const worksheet = workbook.addWorksheet("Waitlist");

      worksheet.columns = [
        { header: "Email", key: "email", width: 30 },
        { header: "WhatsApp", key: "whatsapp", width: 20 },
        { header: "Business Type", key: "businessType", width: 25 },
        { header: "Challenge", key: "challenge", width: 40 },
        { header: "Timestamp", key: "timestamp", width: 30 },
      ];

      worksheet.addRow(newRow).commit();
      await workbook.commit();
      console.log("✅ First row written to new Excel file");
    } else {
      // Append to existing file using a workaround
      const tempPath = FILE_PATH + ".tmp";

      // Read old data
      const oldWorkbook = new ExcelJS.Workbook();
      await oldWorkbook.xlsx.readFile(FILE_PATH);
      const oldSheet = oldWorkbook.getWorksheet("Waitlist");

      // Create new streaming workbook
      const newWorkbook = new ExcelJS.stream.xlsx.WorkbookWriter({
        filename: tempPath,
      });
      const newSheet = newWorkbook.addWorksheet("Waitlist");

      newSheet.columns = [
        { header: "Email", key: "email", width: 30 },
        { header: "WhatsApp", key: "whatsapp", width: 20 },
        { header: "Business Type", key: "businessType", width: 25 },
        { header: "Challenge", key: "challenge", width: 40 },
        { header: "Timestamp", key: "timestamp", width: 30 },
      ];

      // Copy existing rows
      oldSheet.eachRow({ includeEmpty: false }, (row) => {
        newSheet.addRow(row.values.slice(1)).commit();
      });

      // Add new row
      newSheet.addRow(newRow).commit();

      await newWorkbook.commit();

      // Replace original with temp
      fs.renameSync(tempPath, FILE_PATH);
      console.log("✅ New row appended to existing Excel");
    }
  });

  res.json({ success: true });
});

// Optional: Excel download
app.get("/download", (req, res) => {
  if (fs.existsSync(FILE_PATH)) {
    res.download(FILE_PATH, "waitlist.xlsx");
  } else {
    res.status(404).send("Excel file not found.");
  }
});

app.listen(PORT, () => {
  console.log(`🚀 Server running at http://localhost:${PORT}`);
});
