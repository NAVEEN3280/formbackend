import express from "express";
import cors from "cors";
import bodyParser from "body-parser";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import async from "async";
import ExcelJS from "exceljs";

// Setup __dirname for ES modules
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Path to Excel file
const FILE_PATH = path.join(__dirname, "waitlist.xlsx");

const app = express();
const PORT = process.env.PORT || 5000;

// ✅ Allow CORS from your frontend domain (Hostinger)
app.use(
  cors({
    origin: ["https://getchris.vallaham.com", "http://localhost:5173"],
  })
);

app.use(bodyParser.json());

// Create write queue with concurrency = 1
const queue = async.queue(async (task, done) => {
  try {
    await task();
  } finally {
    done();
  }
}, 1);

// Helper to check if file exists and has content
const fileExists = (file) => {
  return fs.existsSync(file) && fs.statSync(file).size > 0;
};

// POST route to collect waitlist data
app.post("/submit", (req, res) => {
  const { email, whatsapp, businessType, challenge } = req.body;

  console.log("📥 New submission from:", email);

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

    if (!fileExists(FILE_PATH)) {
      // First-time creation
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
      console.log("✅ Created and saved first entry.");
    } else {
      // Append by reading old + writing new
      const tempPath = FILE_PATH + ".tmp";

      const oldWorkbook = new ExcelJS.Workbook();
      await oldWorkbook.xlsx.readFile(FILE_PATH);
      const oldSheet = oldWorkbook.getWorksheet("Waitlist");

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

      // Copy old rows
      oldSheet.eachRow({ includeEmpty: false }, (row) => {
        newSheet.addRow(row.values.slice(1)).commit();
      });

      // Add new row
      newSheet.addRow(newRow).commit();

      await newWorkbook.commit();
      fs.renameSync(tempPath, FILE_PATH);
      console.log(`✅ Appended new row for ${email}`);
    }
  });

  res.json({ success: true });
});

// Optional download endpoint
app.get("/download", (req, res) => {
  if (fs.existsSync(FILE_PATH)) {
    res.download(FILE_PATH, "waitlist.xlsx");
  } else {
    res.status(404).send("Excel file not found.");
  }
});

// Start server
app.listen(PORT, () => {
  console.log(`🚀 Server running on port ${PORT}`);
});
