const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");
const chalk = require("chalk");

const INPUT_DIR = "./input";
const BASE_OUTPUT_DIR = "./output";
const CONTAINER_NUMBER_REGEX = /^[A-Z]{4}\d{7}$/;

// Helper Functions
const isValidContainerNumber = (number) => CONTAINER_NUMBER_REGEX.test(number);

const sanitizeFolderName = (name) => {
  return name
    .replace(/[<>:"/\\|?*\x00-\x1F]/g, "")
    .replace(/\s+/g, "_")
    .replace(/\//g, "-")
    .trim();
};

const createOutputDirectory = (refValue) => {
  const sanitizedRefValue = sanitizeFolderName(refValue);
  const refDir = path.join(BASE_OUTPUT_DIR, sanitizedRefValue);
  if (!fs.existsSync(refDir)) {
    fs.mkdirSync(refDir, { recursive: true });
  }
  return refDir;
};

const extractReferenceValue = (sheet) => {
  const data = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  const refRow = data.find(
    (row) => row["__EMPTY_4"] && row["__EMPTY_4"].startsWith("REF: ")
  );
  return refRow ? refRow["__EMPTY_4"].trim() : "Unknown_Reference";
};

const extractContainerData = (sheet) => {
  const data = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  const containerData = {};
  let currentContainer = null;

  data.forEach((row) => {
    const potentialContainer = row["__EMPTY_5"];
    if (row["__EMPTY_4"] && isValidContainerNumber(potentialContainer)) {
      currentContainer = potentialContainer;
      containerData[currentContainer] = [];
    }

    if (currentContainer && row["__EMPTY_9"]) {
      containerData[currentContainer].push({
        Description: row["__EMPTY_9"],
        Total: row["__EMPTY_10"],
      });
    }
  });

  return containerData;
};

const filterValidContainers = (containerData) => {
  return Object.fromEntries(
    Object.entries(containerData).filter(
      ([_, data]) =>
        data.length > 0 &&
        !data.every((item) => item.Description === "Checked & Found.")
    )
  );
};

const createExcelWorkbook = (containerRows) => {
  const header = [
    "Description",
    "Man Hrs",
    "Labour Cost",
    "Material Cost",
    "Total",
    "TAX",
    "CGST",
    "SGST",
    "IGST",
    "Currency",
  ];
  const ws = XLSX.utils.json_to_sheet(containerRows, { header });

  // Set column widths
  const maxWidth = getMaxWidth(containerRows, "Description");
  ws["!cols"] = [
    { wch: maxWidth }, // Set width for Description column
    { wch: 10 }, // Set default width for other columns
    { wch: 10 },
    { wch: 15 },
    { wch: 10 },
    { wch: 10 },
    { wch: 10 },
    { wch: 10 },
    { wch: 10 },
    { wch: 10 },
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Repair Charges Details");
  return wb;
};

const getMaxWidth = (data, key) => {
  let maxLength = key.length; // Start with the header length
  data.forEach((row) => {
    const cellLength = row[key] ? row[key].toString().length : 0;
    if (cellLength > maxLength) {
      maxLength = cellLength;
    }
  });
  return Math.min(maxLength + 2, 100); // Add some padding, cap at 100
};

const processFile = (filePath) => {
  console.log(
    chalk.blue(`[${new Date().toISOString()}] Processing file: ${filePath}`)
  );

  const workbook = XLSX.readFile(filePath);
  let allContainerData = {};
  let refValue = "Unknown_Reference";

  workbook.SheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    const sheetRefValue = extractReferenceValue(sheet);
    if (sheetRefValue !== "Unknown_Reference") {
      refValue = sheetRefValue;
    }
    const sheetContainerData = extractContainerData(sheet);
    allContainerData = { ...allContainerData, ...sheetContainerData };
  });

  return { refValue, containerData: filterValidContainers(allContainerData) };
};

const saveAsExcel = ({ refValue, containerData }) => {
  if (Object.keys(containerData).length === 0) {
    console.warn(
      chalk.yellow(
        `[${new Date().toISOString()}] No valid container data found.`
      )
    );
    return [];
  }

  const refDir = createOutputDirectory(refValue);
  const savedFiles = [];

  Object.entries(containerData).forEach(([containerNo, data]) => {
    if (!isValidContainerNumber(containerNo)) {
      console.warn(
        chalk.yellow(
          `[${new Date().toISOString()}] Skipping invalid container number: ${containerNo}`
        )
      );
      return;
    }

    if (data.length === 0) {
      console.warn(
        chalk.yellow(
          `[${new Date().toISOString()}] Skipping empty container data for: ${containerNo}`
        )
      );
      return;
    }

    const containerRows = data.map((item) => ({
      Description: item.Description,
      "Man Hrs": "",
      "Labour Cost": "",
      "Material Cost": item.Total,
      Total: item.Total,
      TAX: "",
      CGST: "",
      SGST: "",
      IGST: "",
      Currency: "",
    }));

    const workbook = createExcelWorkbook(containerRows);
    const filePath = path.join(
      refDir,
      `RepairChargesUpload-${containerNo}.xlsx`
    );
    XLSX.writeFile(workbook, filePath);
    console.log(
      chalk.green(
        `[${new Date().toISOString()}] Excel file written: ${filePath}`
      )
    );
    savedFiles.push(filePath);
  });

  return savedFiles;
};

const main = () => {
  try {
    const start = Date.now();
    let fileCount = 0;
    let createdFileCount = 0;

    fs.readdirSync(INPUT_DIR)
      .filter((file) => path.extname(file) === ".xlsx")
      .forEach((file) => {
        const filePath = path.join(INPUT_DIR, file);
        fileCount++;
        const { refValue, containerData } = processFile(filePath);
        if (Object.keys(containerData).length > 0) {
          createdFileCount += Object.keys(containerData).length;
        }
        saveAsExcel({ refValue, containerData });
      });

    const end = Date.now();
    const duration = ((end - start) / 1000).toFixed(2); // Duration in seconds

    console.log(chalk.cyan(`\n[${new Date().toISOString()}] Summary:`));
    console.log(chalk.cyan(`- Total input files processed: ${fileCount}`));
    console.log(
      chalk.cyan(`- Total output files created: ${createdFileCount}`)
    );
    console.log(chalk.cyan(`- Total processing time: ${duration} seconds`));
  } catch (error) {
    console.error(
      chalk.red(
        `[${new Date().toISOString()}] An error occurred during processing:`
      ),
      error
    );
  }
};

// main();

module.exports = { processFile, saveAsExcel };
