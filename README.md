# PROJECT_EXCEL_TRANSLINE

## Overview

`PROJECT_EXCEL_TRANSLINE` is a Node.js application designed to process Excel files, extract data, and generate output in various formats. The application supports file uploads, processes the files, and provides options for downloading results and opening output folders.

## Features

- **File Upload**: Upload multiple Excel files for processing.
- **Data Extraction**: Process uploaded files and extract relevant data.
- **Excel Generation**: Save processed data into new Excel files.
- **Zip Download**: Download processed files as a zip archive.
- **Open Folder**: Open the output folder in the file explorer (only available in localhost).

## Installation

1. **Clone the Repository**

   ```bash
   git clone https://github.com/tas33n/PROJECT_EXCEL_TRANSLINE.git
   cd PROJECT_EXCEL_TRANSLINE
   ```

2. **Install Dependencies**

   ```bash
   npm install
   ```

## Usage

1. **Start the Server**

   ```bash
   npm start
   ```

   The server will run on port `3001` by default. You can change the port by setting the `PORT` environment variable.

2. **Open the Application**

   Open your browser and go to `http://localhost:3001`.

3. **Upload Files**

   - Drag and drop files into the drop zone or select files using the file input.
   - Click "Process Files" to upload and process the files.

4. **Download Results**

   - After processing, you can download individual files or a zip archive containing all processed files.

5. **Open Output Folder** (Only available in localhost)

   - Click "Open Output Folder" to view the output folder in your file explorer.

## Development

To contribute to the development of this project, follow the instructions below:

1. **Create a Branch**

   ```bash
   git checkout -b feature/your-feature-name
   ```

2. **Make Changes**

3. **Commit Changes**

   ```bash
   git add .
   git commit -m "Add your commit message here"
   ```

4. **Push Changes**

   ```bash
   git push origin feature/your-feature-name
   ```

5. **Create a Pull Request**

   - Navigate to the repository on GitHub and create a pull request from your branch.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.