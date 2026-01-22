const express = require("express");
const cors = require("cors");
const fs = require("fs");
const ExcelJS = require("exceljs");
const sgMail = require("@sendgrid/mail");

const app = express();

/* =========================
   MIDDLEWARE
========================= */
app.use(cors({
  origin: "*", // allow local + deployed frontend
  methods: ["GET", "POST"],
  allowedHeaders: ["Content-Type"]
}));
app.use(express.json());

/* =========================
   SENDGRID CONFIG
========================= */
sgMail.setApiKey(process.env.SENDGRID_API_KEY);

/* =========================
   EXCEL CONFIG
========================= */
const EXCEL_FILE = "gst_submissions.xlsx";

async function ensureExcel() {
  if (!fs.existsSync(EXCEL_FILE)) {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("GST Data");

    sheet.columns = [
      { header: "Date", key: "date", width: 22 },
      { header: "GSTIN", key: "gstin", width: 25 },
      { header: "Name", key: "name", width: 30 },
      { header: "Email", key: "email", width: 30 },
      { header: "Amount", key: "amount", width: 15 }
    ];

    await workbook.xlsx.writeFile(EXCEL_FILE);
  }
}

/* =========================
   API ROUTE
========================= */
app.post("/submit-gst", async (req, res) => {
  try {
    const { gstin, name, email, amount } = req.body;

    if (!gstin || !name || !email || !amount) {
      return res.status(400).json({ message: "Missing required fields" });
    }

    await ensureExcel();

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(EXCEL_FILE);
    const sheet = workbook.getWorksheet("GST Data");

    sheet.addRow({
      date: new Date().toLocaleString(),
      gstin,
      name,
      email,
      amount
    });

    await workbook.xlsx.writeFile(EXCEL_FILE);

    /* =========================
       SEND EMAIL (SENDGRID)
    ========================= */
    await sgMail.send({
      to: "satish090490@kashiit.ac.in",
      from: "no-reply@gstreturn.app", // must be verified in SendGrid
      subject: "New GST Form Submission",
      text: `New GST form submitted by ${name} (${email})`
    });

    res.json({
      message: "GST submitted, Excel updated & email sent successfully!"
    });

  } catch (err) {
    console.error("SERVER ERROR:", err);
    res.status(500).json({ message: "Internal server error" });
  }
});

/* =========================
   SERVER START
========================= */
const PORT = process.env.PORT || 5000;
app.listen(PORT, () => {
  console.log(`🚀 Server running on port ${PORT}`);
});
