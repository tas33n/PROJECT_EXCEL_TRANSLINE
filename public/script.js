// public/script.js
const dropZone = document.getElementById("dropZone");
const fileInput = document.getElementById("fileInput");
const consoleOutput = document.getElementById("console");
const outputFiles = document.getElementById("outputFiles");

dropZone.addEventListener("click", () => fileInput.click());
dropZone.addEventListener("dragover", (e) => {
  e.preventDefault();
  dropZone.classList.add("drag-over");
});
dropZone.addEventListener("dragleave", () =>
  dropZone.classList.remove("drag-over")
);
dropZone.addEventListener("drop", handleDrop);
fileInput.addEventListener("change", handleFileSelect);

function handleDrop(e) {
  e.preventDefault();
  dropZone.classList.remove("drag-over");
  const files = e.dataTransfer.files;
  processFiles(files);
}

function handleFileSelect(e) {
  const files = e.target.files;
  processFiles(files);
}

function processFiles(files) {
  const formData = new FormData();
  for (let file of files) {
    formData.append("files", file);
  }

  fetch("/upload", {
    method: "POST",
    body: formData,
  })
    .then((response) => response.json())
    .then((data) => {
      updateConsole(data.logs);
      displayOutputFiles(data.files);
    })
    .catch((error) => console.error("Error:", error));
}

function updateConsole(logs) {
  logs.forEach((log) => {
    const logElement = document.createElement("div");
    logElement.textContent = log;
    consoleOutput.appendChild(logElement);
  });
  consoleOutput.scrollTop = consoleOutput.scrollHeight;
}

function displayOutputFiles(files) {
  outputFiles.innerHTML = "";
  files.forEach((file) => {
    const fileCard = document.createElement("div");
    fileCard.className = "file-card";
    fileCard.innerHTML = `
            <span>${file}</span>
            <button class="btn btn-sm btn-primary ms-2" onclick="downloadFile('${file}')">Download</button>
        `;
    outputFiles.appendChild(fileCard);
  });
}

function downloadFile(filename) {
  window.location.href = `/download/${filename}`;
}

function openOutputFolder() {
  fetch("/open-folder")
    .then((response) => response.json())
    .then((data) => {
      if (data.success) {
        console.log("Output folder opened");
      }
    })
    .catch((error) => console.error("Error:", error));
}
