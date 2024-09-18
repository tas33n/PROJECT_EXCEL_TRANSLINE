const express = require("express");
const multer = require("multer");
const path = require("path");
const fs = require("fs");
const archiver = require("archiver");
const { processFile, saveAsExcel } = require("./excelProcessor");

const app = express();
const upload = multer({ dest: "uploads/" });

app.use(express.static("public"));

app.post("/upload", upload.array("files"), (req, res) => {
  const files = req.files;
  const outputDirs = [];
  const logs = [];

  files.forEach((file) => {
    const { refValue, containerData } = processFile(file.path);
    const sanitizedRefValue = refValue.replace(/[^a-zA-Z0-9]/g, "_");
    const folderPath = path.join(__dirname, "output", sanitizedRefValue);

    if (!fs.existsSync(folderPath)) {
      fs.mkdirSync(folderPath, { recursive: true });
    }

    const savedFiles = saveAsExcel({ refValue, containerData });
    savedFiles.forEach((file) => {
      const destPath = path.join(folderPath, path.basename(file));
      fs.renameSync(file, destPath);
    });

    outputDirs.push({ name: sanitizedRefValue, path: folderPath });
    logs.push(`Processed ${file.originalname}`);
  });

  // Convert folder names to be used in the frontend
  const foldersForResponse = outputDirs.map((dir) => ({
    name: dir.name,
  }));

  res.json({ success: true, folders: foldersForResponse, logs });
});

app.get("/download-zip/:folderName", (req, res) => {
  const folderName = req.params.folderName;
  const folderPath = path.join(__dirname, "output", folderName);

  if (!fs.existsSync(folderPath)) {
    return res.status(404).send({ error: "Folder not found" });
  }

  const output = fs.createWriteStream(`output-${folderName}.zip`);
  const archive = archiver("zip", {
    zlib: { level: 9 },
  });

  output.on("close", function () {
    res.download(`output-${folderName}.zip`, `${folderName}.zip`, (err) => {
      if (err) {
        console.error("Error downloading zip:", err);
      }
      // Delete the zip file after download
      fs.unlinkSync(`output-${folderName}.zip`);
    });
  });

  archive.on("error", function (err) {
    res.status(500).send({ error: err.message });
  });

  archive.pipe(output);

  fs.readdir(folderPath, (err, files) => {
    if (err) {
      return res.status(500).send({ error: "Error reading folder" });
    }

    files.forEach((file) => {
      const filePath = path.join(folderPath, file);
      if (fs.statSync(filePath).isFile()) {
        archive.file(filePath, { name: file });
      }
    });

    archive.finalize();
  });
});

app.get("/open-folder", (req, res) => {
  const folderPath = path.join(__dirname, "output");
  require("child_process").exec(`start "" "${folderPath}"`);
  res.json({ success: true });
});

const PORT = process.env.PORT || 3001;
const BASE_URL = `http://localhost:${PORT}`;

app.listen(PORT, async () => {
  const timestamp = new Date().toLocaleString();
  console.log(`[${timestamp}] Excel Mew Mew Server is running on ${BASE_URL}`);

  const { default: open } = await import("open");

  // Open the URL in the default browser after 3 seconds
  setTimeout(() => {
    open(BASE_URL).catch((err) => console.error("Error opening browser:", err));
  }, 3000);
});
