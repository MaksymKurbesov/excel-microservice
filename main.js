import express from "express";
import Excel from "exceljs";
import path from "path";
import cors from "cors";
import fs from "fs";
import { dirname } from "path";
import https from "https";

const app = express();
app.use(cors());
app.use(express.json());

const httpsOptions = {
  cert: fs.readFileSync("ssl/apate.crt"),
  key: fs.readFileSync("ssl/apate.key"),
  ca: fs.readFileSync("ssl/apate.ca-bundle"),
};

app.get("/add-users", async (req, res) => {
  const workbook = new Excel.Workbook();
  const sheet = workbook.addWorksheet("Participants");

  sheet.getCell(`A1`).value = `test`;

  const fileName = "Participants.xlsx";
  await workbook.xlsx.writeFile(fileName);

  res.sendFile(fileName, { root: path.resolve() });
  // res.send("Hello World123");
});

https.createServer(httpsOptions, app).listen(3000, () => {
  console.log("listen2");
});

export const viteNodeApp = app;
