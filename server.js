const express = require("express");
const ExcelJS = require("exceljs");
const cors = require("cors");
const path = require("path");
const fs = require("fs");

const app = express();
app.use(express.json());
app.use(cors());

app.get("/", (req, res) => {
    res.sendFile(path.join(__dirname, "index.html"));
});

const FILE_NAME = path.join(process.cwd(), "data.xlsx");

const SHEETS = ["TOKOA","TOKOB","TOKOC","TOKOD","TOKOE","SEMUA"];

// AUTO CREATE FILE
async function initFile() {
  if (!fs.existsSync(FILE_NAME)) {
    const workbook = new ExcelJS.Workbook();

    const headers = [
      "SPX","ANTERAJA","JNT","JNE",
      "SICEPAT","GTL","CARGO","INSTAN"
    ];

    SHEETS.forEach(name => {
      const sheet = workbook.addWorksheet(name);
      sheet.getRow(1).values = headers;
    });

    await workbook.xlsx.writeFile(FILE_NAME);
  }
}

// DETECT KURIR
function detectKurir(resi) {
  if (resi.startsWith("SPX")) return { col: 1 };
  if (resi.startsWith("TSA")) return { col: 2 };
  if (resi.startsWith("110035")) return { col: 2 };
  if (resi.startsWith("JX")) return { col: 3 };
  if (resi.startsWith("CM")) return { col: 4 };
  if (resi.startsWith("JY")) return { col: 4 };
  if (resi.startsWith("TG")) return { col: 4 };
  if (resi.startsWith("00")) return { col: 5 };
  if (resi.startsWith("GTL")) return { col: 6 };
  if (resi.startsWith("570")) return { col: 7 };
  return { col: 8 };
}

// ADD RESI
app.post("/add", async (req, res) => {
  const { resi, toko } = req.body;

  await initFile();

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(FILE_NAME);

  // CEK DUPLICATE
  let duplicate = false;

  workbook.eachSheet(sheet => {
    sheet.eachRow(row => {
      row.eachCell(cell => {
        if (cell.value === resi) duplicate = true;
      });
    });
  });

  if (duplicate) {
    return res.json({ status: "duplicate" });
  }

  const { col } = detectKurir(resi);

  const targets = [toko, "SEMUA"];

  for (let name of targets) {
    let sheet = workbook.getWorksheet(name);

    if (!sheet) {
      sheet = workbook.addWorksheet(name);
      sheet.getRow(1).values = [
        "SPX","ANTERAJA","JNT","JNE",
        "SICEPAT","GTL","CARGO","INSTAN"
      ];
    }

    let row = 2;
    while (sheet.getCell(row, col).value) row++;

    sheet.getCell(row, col).value = resi;
  }

  await workbook.xlsx.writeFile(FILE_NAME);

  res.json({ status: "success" });
});

// GET DATA
app.get("/sheet/:name", async (req, res) => {
  await initFile();

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(FILE_NAME);

  const sheet = workbook.getWorksheet(req.params.name);

  if (!sheet) return res.json([]);

  let data = [];

  sheet.eachRow((row, i) => {
    if (i === 1) return;

    data.push({
      SPX: row.getCell(1).value || "",
      ANTERAJA: row.getCell(2).value || "",
      JNT: row.getCell(3).value || "",
      JNE: row.getCell(4).value || "",
      SICEPAT: row.getCell(5).value || "",
      GTL: row.getCell(6).value || "",
      CARGO: row.getCell(7).value || "",
      INSTAN: row.getCell(8).value || ""
    });
  });

  res.json(data);
});

// EXPORT
app.get("/export", (req, res) => {
  const fileName = `scan-${new Date().toISOString().slice(0,10)}.xlsx`;
  res.download(FILE_NAME, fileName);
});

app.listen(5000, () => {
  console.log("Server jalan di http://localhost:5000");
});