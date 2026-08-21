const express = require("express");
const ExcelJS = require("exceljs");
const cors = require("cors");
const path = require("path");
const fs = require("fs");
const crypto = require("crypto");

const app = express();

app.use(express.json());
app.use(cors());
require("dotenv").config();

// ======================================================
// KONFIGURASI
// ======================================================

const PORT = 5000;

const FILE_NAME = path.join(process.cwd(), "data.xlsx");

const SHEETS = [
  "TOKOA",
  "TOKOB",
  "TOKOC",
  "TOKOD",
  "TOKOE",
  "SEMUA"
];

const HEADERS = [
  "SPX",
  "ANTERAJA",
  "JNT",
  "JNE",
  "SICEPAT",
  "GTL",
  "CARGO",
  "INSTAN"
];

// ======================================================
// USER SEDERHANA
// ======================================================
//
// ROLE 1 = hanya ekspedisi
// ROLE 2 = ekspedisi + toko
//
// Nanti kalau aplikasi online, bagian ini sebaiknya
// dipindahkan ke database dan password di-hash.
//

// const USERS = [
//   {
//     username: "operator",
//     password: "operator123",
//     role: 1
//   },
//   {
//     username: "admin",
//     password: "admin123",
//     role: 2
//   }
// ];

const USERS = [
  {
    username: process.env.ROLE1_USERNAME,
    password: process.env.ROLE1_PASSWORD,
    role: 1
  },
  {
    username: process.env.ROLE2_USERNAME,
    password: process.env.ROLE2_PASSWORD,
    role: 2
  }
];

// ======================================================
// SESSION
// ======================================================

const sessions = new Map();

function createSession(user) {
  const token = crypto.randomBytes(32).toString("hex");

  sessions.set(token, {
    username: user.username,
    role: user.role
  });

  return token;
}

function getSession(req) {
  const token = req.headers.authorization;

  if (!token) {
    return null;
  }

  return sessions.get(token) || null;
}

function requireLogin(req, res, next) {
  const session = getSession(req);

  if (!session) {
    return res.status(401).json({
      status: "unauthorized",
      message: "Silakan login terlebih dahulu."
    });
  }

  req.user = session;

  next();
}

function requireRole2(req, res, next) {
  if (!req.user || req.user.role !== 2) {
    return res.status(403).json({
      status: "forbidden",
      message: "Akses hanya untuk Role 2."
    });
  }

  next();
}

// ======================================================
// INDEX.HTML
// ======================================================

app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "index.html"));
});

// ======================================================
// AUTO CREATE EXCEL
// ======================================================

async function initFile() {
  if (!fs.existsSync(FILE_NAME)) {
    const workbook = new ExcelJS.Workbook();

    SHEETS.forEach(name => {
      const sheet = workbook.addWorksheet(name);

      sheet.getRow(1).values = HEADERS;

      sheet.getRow(1).font = {
        bold: true
      };
    });

    await workbook.xlsx.writeFile(FILE_NAME);

    console.log("data.xlsx berhasil dibuat.");
  }
}

// ======================================================
// DETECT KURIR
// ======================================================

function detectKurir(resi) {

  resi = resi.toUpperCase();

  if (resi.startsWith("SPX"))
    return {
      kurir: "SPX",
      col: 1
    };

  if (resi.startsWith("TSA"))
    return {
      kurir: "ANTERAJA",
      col: 2
    };

  if (resi.startsWith("110035"))
    return {
      kurir: "ANTERAJA",
      col: 2
    };

  if (resi.startsWith("JX"))
    return {
      kurir: "JNT",
      col: 3
    };

  if (resi.startsWith("CM"))
    return {
      kurir: "JNE",
      col: 4
    };

  if (resi.startsWith("JY"))
    return {
      kurir: "JNE",
      col: 4
    };

  if (resi.startsWith("TG"))
    return {
      kurir: "JNE",
      col: 4
    };

  if (resi.startsWith("00"))
    return {
      kurir: "SICEPAT",
      col: 5
    };

  if (resi.startsWith("GTL"))
    return {
      kurir: "GTL",
      col: 6
    };

  if (resi.startsWith("570"))
    return {
      kurir: "CARGO",
      col: 7
    };

  return {
    kurir: "INSTAN/SAMEDAY",
    col: 8
  };
}

// ======================================================
// LOGIN
// ======================================================

app.post("/login", (req, res) => {

  const {
    username,
    password
  } = req.body;

  const user = USERS.find(u =>
    u.username === username &&
    u.password === password
  );

  if (!user) {

    return res.status(401).json({
      status: "failed",
      message: "Username atau password salah."
    });

  }

  const token = createSession(user);

  res.json({
    status: "success",
    token: token,
    username: user.username,
    role: user.role
  });
});

// ======================================================
// LOGOUT
// ======================================================

app.post("/logout", requireLogin, (req, res) => {

  const token = req.headers.authorization;

  sessions.delete(token);

  res.json({
    status: "success"
  });
});

// ======================================================
// CEK SESSION
// ======================================================

app.get("/me", requireLogin, (req, res) => {

  res.json({
    status: "success",
    username: req.user.username,
    role: req.user.role
  });
});

// ======================================================
// ADD RESI
// ======================================================

app.post("/add", requireLogin, async (req, res) => {

  try {

    let {
      resi,
      toko
    } = req.body;

    if (!resi) {

      return res.status(400).json({
        status: "failed",
        message: "Resi kosong."
      });

    }

    // ==================================================
    // BERSIHKAN HASIL SCANNER
    // ==================================================
    //
    // Contoh:
    //
    // SPX123456 BUDI 14:30
    //
    // menjadi:
    //
    // SPX123456
    //

    resi = resi
      .trim()
      .split(/\s+/)[0]
      .toUpperCase();

    // ==================================================
    // ROLE 1
    // ==================================================
    //
    // Role 1 tidak boleh menentukan toko.
    //

    if (req.user.role === 1) {
      toko = null;
    }

    // ==================================================
    // ROLE 2
    // ==================================================

    if (req.user.role === 2) {

      if (!SHEETS.includes(toko) || toko === "SEMUA") {

        return res.status(400).json({
          status: "failed",
          message: "Toko tidak valid."
        });

      }

    }

    await initFile();

    const workbook = new ExcelJS.Workbook();

    await workbook.xlsx.readFile(FILE_NAME);

    // ==================================================
    // CEK DUPLICATE
    // ==================================================

    let duplicate = false;

    workbook.eachSheet(sheet => {

      sheet.eachRow(row => {

        row.eachCell(cell => {

          if (
            String(cell.value || "").trim().toUpperCase() === resi
          ) {
            duplicate = true;
          }

        });

      });

    });

    if (duplicate) {

      return res.json({
        status: "duplicate",
        resi: resi
      });

    }

    // ==================================================
    // DETEKSI KURIR
    // ==================================================

    const {
      kurir,
      col
    } = detectKurir(resi);

    // ==================================================
    // TENTUKAN TARGET SHEET
    // ==================================================

    let targets = [];

    if (req.user.role === 1) {

      // Role 1 hanya masuk SEMUA

      targets = ["SEMUA"];

    } else {

      // Role 2:
      //
      // masuk toko
      // +
      // SEMUA

      targets = [
        toko,
        "SEMUA"
      ];

    }

    // ==================================================
    // MASUKKAN DATA
    // ==================================================

    for (const name of targets) {

      let sheet = workbook.getWorksheet(name);

      if (!sheet) {

        sheet = workbook.addWorksheet(name);

        sheet.getRow(1).values = HEADERS;

        sheet.getRow(1).font = {
          bold: true
        };

      }

      let row = 2;

      while (
        sheet.getCell(row, col).value
      ) {
        row++;
      }

      sheet.getCell(row, col).value = resi;

    }

    // ==================================================
    // SIMPAN
    // ==================================================

    await workbook.xlsx.writeFile(FILE_NAME);

    res.json({

      status: "success",

      resi: resi,

      kurir: kurir,

      toko: toko || "SEMUA"

    });

  } catch (error) {

    console.error(error);

    res.status(500).json({

      status: "error",

      message: error.message

    });

  }

});

// ======================================================
// GET DATA SHEET
// ======================================================

app.get(
  "/sheet/:name",
  requireLogin,
  async (req, res) => {

    try {

      const sheetName = req.params.name;

      // ================================================
      // ROLE 1
      // ================================================

      if (
        req.user.role === 1 &&
        sheetName !== "SEMUA"
      ) {

        return res.status(403).json({
          status: "forbidden",
          message: "Role 1 hanya dapat melihat SEMUA."
        });

      }

      // ================================================
      // VALIDASI SHEET
      // ================================================

      if (!SHEETS.includes(sheetName)) {

        return res.status(400).json({
          status: "failed",
          message: "Sheet tidak valid."
        });

      }

      await initFile();

      const workbook = new ExcelJS.Workbook();

      await workbook.xlsx.readFile(FILE_NAME);

      const sheet =
        workbook.getWorksheet(sheetName);

      if (!sheet) {

        return res.json([]);

      }

      const data = [];

      sheet.eachRow((row, i) => {

        if (i === 1) return;

        data.push({

          SPX:
            row.getCell(1).value || "",

          ANTERAJA:
            row.getCell(2).value || "",

          JNT:
            row.getCell(3).value || "",

          JNE:
            row.getCell(4).value || "",

          SICEPAT:
            row.getCell(5).value || "",

          GTL:
            row.getCell(6).value || "",

          CARGO:
            row.getCell(7).value || "",

          INSTAN:
            row.getCell(8).value || ""

        });

      });

      res.json(data);

    } catch (error) {

      console.error(error);

      res.status(500).json({
        status: "error",
        message: error.message
      });

    }

  }
);

// ======================================================
// EXPORT EXCEL
// ======================================================

app.get(
  "/export",
  requireLogin,
  async (req, res) => {

    await initFile();

    const fileName =
      `scan-${new Date()
        .toISOString()
        .slice(0, 10)}.xlsx`;

    res.download(
      FILE_NAME,
      fileName
    );

  }
);

// ======================================================
// SERVER
// ======================================================

app.listen(PORT, () => {

  console.log(
    `Server jalan di http://localhost:${PORT}`
  );

});