import express from "express";
import Excel from "exceljs";
import path from "path";
import cors from "cors";
import fs from "fs";
import { dirname } from "path";
import https from "https";

const COLUMNS_COUNT = 11;

const app = express();
app.use(cors());
app.use(express.json());

const httpsOptions = {
  cert: fs.readFileSync("ssl/apate.crt"),
  key: fs.readFileSync("ssl/apate.key"),
  ca: fs.readFileSync("ssl/apate.ca-bundle"),
};

app.post("/add-users", async (req, res) => {
  const filePath = "Participants.xlsx";
  const workbook = new Excel.Workbook();
  const { username, phone, messenger, answers } = req.body;

  console.log(req.body, "req.body");

  try {
    await workbook.xlsx.readFile(filePath);
    const worksheet =
      workbook.getWorksheet("Participants") ||
      workbook.addWorksheet("Participants");
    worksheet
      .addRow([username, phone, messenger, ...Object.values(answers)])
      .commit();

    for (let i = 4; i <= 11; i++) {
      worksheet
        .getColumn(i)
        .eachCell({ includeEmpty: true }, (cell, rowNumber) => {
          cell.alignment = { wrapText: true };
        });
    }

    for (let i = 0; i < COLUMNS_COUNT; i++) {
      worksheet.getColumn(i + 1).width = 25;
    }

    await workbook.xlsx.writeFile(filePath);
    res.status(200).json({ message: "Пользователь успешно добавлен" });
  } catch (e) {
    console.log(e, "error");
  }
});

app.get("/download-sheet", (req, res) => {
  const filePath = "Participants.xlsx";
  res.status(200).sendFile(filePath, { root: path.resolve() });
});

https.createServer(httpsOptions, app).listen(3000, () => {
  console.log("listen2");
});

export const viteNodeApp = app;
