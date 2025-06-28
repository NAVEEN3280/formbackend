import express from "express";
import cors from "cors";
import bodyParser from "body-parser";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import async from "async";
import ExcelJS from "exceljs";
import axios from "axios";
import dotenv from "dotenv";
dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const FILE_PATH = path.join(__dirname, "waitlist.xlsx");

const app = express();
const PORT = process.env.PORT || 5000;

// ✅ Allow frontend origin (Render + Local)
app.use(
  cors({
    origin: ["https://getchris.in", "http://localhost:5173"],
  })
);

app.use(bodyParser.json());
app.set("trust proxy", true); // Needed for real IPs on Render

const queue = async.queue(async (task, done) => {
  try {
    await task();
  } finally {
    done();
  }
}, 1);

const fileExists = (file) => fs.existsSync(file) && fs.statSync(file).size > 0;

app.post("/submit", async (req, res) => {
  const { email, whatsapp, businessType, challenge } = req.body;

  // ✅ Get real IP
  let userIP =
    req.headers["x-forwarded-for"]?.split(",")[0]?.trim() ||
    req.socket.remoteAddress;

  // If local or unknown IP, use fallback for testing
  if (
    !userIP ||
    userIP === "::1" ||
    userIP.startsWith("::ffff:127.") ||
    userIP === "127.0.0.1"
  ) {
    userIP = "8.8.8.8"; // fallback IP for development
  }

  console.log("📡 Detected IP:", userIP);

  // ✅ Get location from ipinfo.io
  let location = {
    ip: userIP,
    city: "Unknown",
    region: "Unknown",
    country: "Unknown",
  };

  try {
    const geo = await axios.get(
      `https://ipinfo.io/${userIP}?token=${process.env.IPINFO_TOKEN}`
    );

    location = {
      ip: userIP,
      city: geo.data.city || "Unknown",
      region: geo.data.region || "Unknown",
      country: geo.data.country || "Unknown",
    };

    console.log("🌍 Location:", location);
  } catch (err) {
    console.warn("⚠️ IP lookup failed:", err.message);
  }

  const newRow = {
    email,
    whatsapp,
    businessType,
    challenge,
    timestamp: new Date().toLocaleString("en-IN", {
      timeZone: "Asia/Kolkata",
    }),
    ip: location.ip,
    city: location.city,
    region: location.region,
    country: location.country,
  };

  queue.push(async () => {
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
        { header: "IP", key: "ip", width: 20 },
        { header: "City", key: "city", width: 20 },
        { header: "Region", key: "region", width: 20 },
        { header: "Country", key: "country", width: 20 },
      ];

      worksheet.addRow(newRow).commit();
      await workbook.commit();
      console.log("✅ First row saved.");
    } else {
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
        { header: "IP", key: "ip", width: 20 },
        { header: "City", key: "city", width: 20 },
        { header: "Region", key: "region", width: 20 },
        { header: "Country", key: "country", width: 20 },
      ];

      oldSheet.eachRow({ includeEmpty: false }, (row) => {
        newSheet.addRow(row.values.slice(1)).commit();
      });

      newSheet.addRow(newRow).commit();
      await newWorkbook.commit();

      fs.renameSync(tempPath, FILE_PATH);
      console.log(`✅ New row appended for ${email}`);
    }
  });

  res.json({ success: true });
});

app.get("/download", (req, res) => {
  if (fs.existsSync(FILE_PATH)) {
    res.download(FILE_PATH, "waitlist.xlsx");
  } else {
    res.status(404).send("Excel file not found.");
  }
});

app.listen(PORT, () => {
  console.log(`🚀 Server running on port ${PORT}`);
});
