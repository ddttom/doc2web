# doc2web

A Node.js application that extracts content from Microsoft Word (.docx) documents and converts them to detailed markdown and HTML, preserving styling and structure.

## New Features

- **Flexible Input Handling**:
  - Process a single DOCX file
  - Process all DOCX files in a directory (including subdirectories)
  - Process a list of files from a text file (e.g., output from `find` command)

- **Organized Output Structure**:
  - Outputs are automatically organized in a directory structure mirroring the input
  - All files are placed under an `./output` base directory
  - Directory structure is preserved and created if it doesn't exist

## Installation

1. Clone this repository or download the files
2. Install dependencies:

```bash
npm install
```

## Usage

### Process a Single File

```bash
node doc2web.js /path/to/your/document.docx
```

### Process a Directory (including subdirectories)

```bash
node doc2web.js /path/to/your/documents/folder
```

### Process a List of Files (e.g., from `find` command)

You can capture the output of a `find` command into a file:

```bash
find /path/to/search -name "*.docx" > docx-files.txt
```

Then process all those files:

```bash
node doc2web.js docx-files.txt --list
```

Alternatively, use the provided helper script:

```bash
./process-find.sh docx-files.txt
```

### Options

- `--html-only`: Generate only HTML output, skip markdown conversion
- `--list`: Explicitly treat the input file as a list of files to process

## Output Structure

For each processed DOCX file, the following outputs are generated:

- `output/original/path/filename.md`: Markdown version
- `output/original/path/filename.html`: HTML version with styles
- `output/original/path/filename.css`: Extracted CSS styles
- `output/original/path/images/`: Folder containing extracted images

For example, if you process `/home/user/docs/report.docx`, the outputs will be:

```
output/home/user/docs/report.md
output/home/user/docs/report.html
output/home/user/docs/report.css
output/home/user/docs/images/
```

## Features

- Converts DOCX documents to markdown and HTML formats
- Preserves document structure, formatting, and styles
- Extracts and applies document styles for faithful HTML rendering
- Full Unicode support for international languages
- Handles mixed language content within documents
- Extracts and saves images from documents

## Example Use Cases

- **Batch Processing Documentation**: Convert entire documentation libraries
- **Website Content Migration**: Extract content from Word documents for website publishing
- **Knowledge Base Conversion**: Transform corporate knowledge bases to markdown
- **Archive Processing**: Convert document archives to more accessible formats
