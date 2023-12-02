# Story Upload Page

Welcome to the Story Upload Page! This web application allows you to upload a story in .docx format and a spreadsheet of names in .xlsx format. The program then processes the story by finding and replacing the names, creating a new story. You can view the old and new stories using Quill editors and download the new story in .docx format.

## Table of Contents

- [Getting Started](#getting-started)
  - [Prerequisites](#prerequisites)
  - [Installation](#installation)
- [Usage](#usage)
  - [Uploading Story](#uploading-story)
  - [Uploading Names](#uploading-names)
  - [Processing Story](#processing-story)
  - [Downloading New Story](#downloading-new-story)
- [Dependencies](#dependencies)
- [License](#license)

## Getting Started

### Prerequisites

Make sure you have the following installed on your machine:

- Node.js
- A modern web browser (e.g., Chrome, Firefox)

### Installation

1. Clone this repository to your local machine:

    ```bash
    git clone https://github.com/your-username/story-upload-page.git
    ```

2. Navigate to the project directory:

    ```bash
    cd story-upload-page
    ```

3. Open the index.html file in your preferred web browser.

## Usage

### Uploading Story

1. Click the "Upload Story" button.
2. Select a valid .docx file when prompted.
3. The old story content will be displayed in the "Old Story" Quill editor.

### Uploading Names

1. Click the "Upload Names" button.
2. Select a valid .xlsx file containing a spreadsheet of names when prompted.
3. The names data will be logged to the console.

### Processing Story

1. After uploading both the story and names, click the "Create New Story" button.
2. The program will find and replace names in the story, and the updated content will be displayed in the "New Story" Quill editor.

### Downloading New Story

1. Click the "Download New Story" button.
2. A new .docx file containing the updated story will be downloaded to your machine.

## Dependencies

- Quill: A modern WYSIWYG editor.
- FileSaver.js: A library for saving files on the client-side.
- JSZip: A JavaScript library for creating, reading, and editing .zip files.
- docxtemplater: A library for generating .docx documents.
- ExcelJS: A library for reading, writing, and manipulating Excel files.
- SheetJS: A library for reading and writing spreadsheet files.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
