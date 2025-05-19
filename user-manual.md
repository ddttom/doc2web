# doc2web User Manual

## Table of Contents

1. [Introduction](#introduction)
2. [Installation](#installation)
3. [Basic Usage](#basic-usage)
4. [Advanced Features](#advanced-features)
5. [Command Reference](#command-reference)
6. [Output Files](#output-files)
7. [Formatting Features](#formatting-features)
8. [Multilingual Support](#multilingual-support)
9. [Batch Processing](#batch-processing)
10. [Interactive Mode](#interactive-mode)
11. [Troubleshooting](#troubleshooting)
12. [FAQ](#faq)

## Introduction

doc2web is a powerful tool for converting Microsoft Word (.docx) documents to web-friendly formats, specifically Markdown and HTML. Unlike simple converters, doc2web preserves document styling, structure, images, and supports multilingual content including mixed scripts.

### Key Features

- Converts DOCX to both Markdown and HTML formats
- Preserves document styling in the HTML output
- Extracts and properly organizes images
- Maintains document structure (headings, lists, tables)
- Supports Unicode and multilingual content
- Processes single files, directories, or file lists
- Preserves original directory structures in output

## Installation

### Prerequisites

- Node.js (version 14.x or higher)
- npm (usually comes with Node.js)

### Automatic Installation

Download all project files to a directory

Make the installation script executable:

```bash
chmod +x doc2web-install.sh
```

Run the installation script:

```bash
./doc2web-install.sh
```

Make helper scripts executable:

```bash
chmod +x doc2web-run.js
chmod +x process-find.sh
```

## Basic Usage

### Converting a Single Document

To convert a single DOCX file:

```bash
node doc2web.js path/to/your/document.docx
```

This will create in the `output` directory:

- A markdown version (.md)
- An HTML version with styles (.html)
- A CSS file with extracted styles (.css)
- An images folder with any extracted images

### Output Location

All output files are stored under the `./output` directory, preserving the original directory structure. For example:

- Input: `/home/user/documents/report.docx`
- Output:
  - `./output/home/user/documents/report.md`
  - `./output/home/user/documents/report.html`
  - `./output/home/user/documents/report.css`
  - `./output/home/user/documents/images/`

## Advanced Features

### HTML-Only Mode

If you only need the HTML version (without Markdown):

```bash
node doc2web.js document.docx --html-only
```

### Processing a Directory

To process all DOCX files in a directory (including subdirectories):

```bash
node doc2web.js /path/to/documents/folder
```

### Processing a List of Files

To process multiple specific files:

Create a text file with one filepath per line:

```bash
/path/to/first/document.docx
/path/to/second/document.docx
/another/path/report.docx
```

Process the list:

```bash
node doc2web.js file-list.txt --list
```

### Using Find Command Output

You can process results from a find command:

```bash
# Find all DOCX files in a location
find /search/path -name "*.docx" > docx-files.txt

# Process those files with the helper script
./process-find.sh docx-files.txt
```

## Command Reference

### doc2web.js Options

```bash
node doc2web.js <input> [options]
```

- `<input>`: Can be:
  - Path to a DOCX file
  - Path to a directory
  - Path to a text file containing a list of files

- Options:
  - `--html-only`: Skip markdown generation
  - `--list`: Treat the input file as a list of files

### process-find.sh

```bash
./process-find.sh <find-output-file>
```

Processes a file containing the output of a find command.

### doc2web-run.js

```bash
node doc2web-run.js
```

Launches the interactive command-line interface.

## Output Files

For each processed DOCX file, doc2web generates:

### Markdown File (.md)

- Clean, formatted Markdown
- Preserves document structure
- Includes references to images
- Suitable for use in static site generators, documentation systems, etc.

### HTML File (.html)

- Complete HTML document with styles
- Closely resembles the original document appearance
- Includes references to images
- Can be viewed directly in a browser

### CSS File (.css)

- Contains styles extracted from the document
- Implements DOCX styling in web-friendly format
- Linked from the HTML file

### Images Folder

- Contains all images extracted from the document
- Images are referenced from both Markdown and HTML

## Formatting Features

doc2web preserves the following elements and provides special handling for navigation elements:

### Document Structure

- Headings (H1-H6)
- Paragraphs with proper spacing
- Lists (ordered and unordered)
- Tables with formatting
- Images with captions

### Text Formatting

- Bold, italic, underline
- Strikethrough
- Superscript and subscript
- Font styles and sizes (in HTML output)
- Text colors (in HTML output)

### Advanced Elements

- Hyperlinks
- Blockquotes
- Code blocks
- Horizontal rules

### Table of Contents and Index Handling

doc2web intelligently detects and handles navigation elements:

- Automatically identifies Table of Contents (TOC) sections
- Detects Index sections at the end of documents
- Replaces these elements with clear placeholders instead of converting their full content
- In HTML output: `<p class="docx-placeholder">** TOC HERE **</p>` or `<p class="docx-placeholder">** INDEX HERE **</p>`
- In Markdown output: `**TOC HERE**` or `**INDEX HERE**`

This feature prevents unnecessary duplication of navigation elements that aren't needed in web formats, improving readability and reducing file size.

## Multilingual Support

doc2web provides robust support for international content:

- Full Unicode support
- Mixed language content in the same document
- Right-to-left languages (Arabic, Hebrew)
- CJK (Chinese, Japanese, Korean) characters
- Special characters and symbols

Examples of supported content:

- Traditional Chinese: 澳門 (Macao)
- Cyrillic: ею (European Union)
- Mixed content: "Resolved (2016.02.03.06), ICANN has reviewed the ею domain"

## Batch Processing

The batch processing capabilities make doc2web suitable for:

### Mass Document Conversion

Convert entire document libraries:

```bash
node doc2web.js /path/to/document/library
```

### Selective Processing

For more control, create a list of specific files to process:

```bash
# Create a list of files
echo "/path/to/doc1.docx" > to-process.txt
echo "/path/to/doc2.docx" >> to-process.txt
echo "/path/to/doc3.docx" >> to-process.txt

# Process the list
node doc2web.js to-process.txt --list
```

### Processing Find Results

Filter documents with complex criteria:

```bash
# Find documents modified in the last 7 days
find /docs -name "*.docx" -mtime -7 > recent-docs.txt

# Process those documents
./process-find.sh recent-docs.txt
```

## Interactive Mode

For a more user-friendly experience, use the interactive mode:

```bash
node doc2web-run.js
```

This provides a menu-driven interface with options to:

1. Process a single DOCX file
2. Process a directory of DOCX files
3. Process a list of files from a text file
4. Search and process DOCX files (using find command)

The interactive mode guides you through each step, including selecting input files and choosing whether to generate HTML only or both HTML and Markdown.

## Troubleshooting

### Common Issues

#### Missing Dependencies

If you encounter errors about missing modules:

```bash
./doc2web-install.sh
```

#### Permission Denied

If you get "Permission denied" when running scripts:

```bash
chmod +x doc2web-run.js
chmod +x process-find.sh
chmod +x doc2web-install.sh
```

#### Output Directory Issues

If output files aren't being created:

```bash
mkdir -p ./output
```

#### Large Document Processing

For very large documents that cause memory issues:

```bash
# Increase Node.js memory limit to 4GB
NODE_OPTIONS=--max-old-space-size=4096 node doc2web.js large-document.docx
```

### Error Messages

- **"Error: File not found"**: Check the file path and ensure the file exists.
- **"Error: Not a .docx document"**: The file must be in DOCX format.
- **"Error processing document"**: The document may be corrupted or password-protected.

## FAQ

### General Questions

**Q: Will doc2web work with older .doc format?**  
A: No, doc2web only supports the .docx format.

**Q: Does doc2web support password-protected documents?**  
A: No, documents must be unprotected.

**Q: What happens to macros in the document?**  
A: Macros are ignored in the conversion process.

### Output Questions

**Q: Can I customize the output directory?**  
A: Not yet, all outputs go to the `./output` directory.

**Q: How closely does the HTML match the original document?**  
A: The HTML output attempts to preserve styling, but complex layouts may differ slightly.

**Q: Why are Table of Contents and Index sections replaced with placeholders?**  
A: These navigation elements are typically not needed in web formats and can create unnecessary duplication. The placeholders make it clear where these elements were in the original document without cluttering the output.

**Q: Can doc2web convert to PDF?**  
A: No, only Markdown and HTML outputs are currently supported.

### Technical Questions

**Q: What Node.js version is required?**  
A: Node.js 14.x or higher is recommended.

**Q: Does doc2web support DOCX files created by applications other than Microsoft Word?**  
A: Yes, it should work with any standard DOCX files, including those created by LibreOffice, Google Docs, etc.

**Q: Is doc2web available as an npm package?**  
A: Not yet, but you can use it directly from the source code.

---

For further assistance or to report issues, please open an issue on the GitHub repository.
