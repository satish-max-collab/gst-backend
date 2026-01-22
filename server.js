const express = require("express");
const cors = require("cors");
const fs = require("fs");
const ExcelJS = require("exceljs");
const nodemailer = require("nodemailer");

const app = express();
app.use(cors());
app.use(express.json());

const EXCEL_FILE = "gst_submissions.xlsx";

/* Create Excel if not exists */
async function ensureExcel() {
  if (!fs.existsSync(EXCEL_FILE)) {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("GST Data");

    sheet.columns = [
      { header: "Date", key: "date", width: 20 },
      { header: "GSTIN", key: "gstin", width: 25 },
      { header: "Name", key: "name", width: 30 },
      { header: "Email", key: "email", width: 30 },
      { header: "Amount", key: "amount", width: 15 }
    ];

    await workbook.xlsx.writeFile(EXCEL_FILE);
  }
}

app.post("/submit-gst", async (req, res) => {
  try {
    await ensureExcel();

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(EXCEL_FILE);
    const sheet = workbook.getWorksheet("GST Data");

    sheet.addRow({
      date: new Date().toLocaleString(),
      gstin: req.body.gstin,
      name: req.body.name,
      email: req.body.email,
      amount: req.body.amount
    });

    await workbook.xlsx.writeFile(EXCEL_FILE);

    /* Send Email */
    const transporter = nodemailer.createTransport({
      service: "gmail",
      auth: {
        user: "satish090490@kashiit.ac.in",
        pass: "skv009kit"
      }
    });

    await transporter.sendMail({
      from: "satish090490@kashiit.ac.in",
      to: "satish090490@kashiit.ac.in",
      subject: "New GST Form Submission",
      text: "GST form submitted. Excel attached.",
      attachments: [
        {
          filename: "gst_submissions.xlsx",
          path: EXCEL_FILE
        }
      ]
    });

    res.json({ message: "Form submitted & email sent successfully!" });

  } catch (err) {
    console.error(err);
    res.status(500).json({ message: "Server error" });
  }
});

app.listen(5000, () => {
  console.log("Server running on http://localhost:5000");
});
