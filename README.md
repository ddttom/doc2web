# doc2web

doc2web is a powerful tool for converting Microsoft Word (.docx) documents to web-friendly formats, including Markdown and HTML with preserved styling.

## Overview

This tool extracts content from DOCX files while maintaining:

- Document structure and formatting
- Images and tables
- Styles and layout
- Unicode and multilingual content

The app must not make any assumptions from test documents, the app must treat creatred css and html as ephemeral, they will be destroyed on every run.
The css and HTML are individual to each document created, they wiol be named after the docx inourt, with folder pattern matched.
The ap must deeply inspect the docx and obtain css and formatting styles, fonr sizes, item prefixes etc.  Nothing will be hard coded with the app. All generated CSS will be in an external stylesheet

## Project Structure

```bash
GitHub/doc2web
├── .gitignore
├── .vscode
├── README.md
├── doc2web-install.sh
├── doc2web-run.js
├── doc2web.js
├── docs
│   └── prd.md
├── docx-style-parser.js
├── init-doc2web.sh
├── input
├── markdownify.js
├── node_modules
├── output
├── package-lock.json
├── package.json
├── process-find.sh
├── style-extractor.js
└── user-manual.md
```

When making code edits update prd.md user-manual.md and README.md

## Getting Started

### Initial Setup

1. Clone or download all project files to a directory:

   ```bash
   git clone https://github.com/ddttom/doc2web.git
   cd doc2web
   ```

2. Install dependencies:

   ```bash
   npm install
   ```

3. Make the helper scripts executable (Unix/Linux/macOS):

   ```bash
   chmod +x *.sh
   ```

## Usage

### Command Line

Process a single DOCX file:

```bash
node doc2web.js path/to/document.docx
```

Process all DOCX files in a directory:

```bash
node doc2web.js path/to/directory/
```

Process a list of files:

```bash
node doc2web.js file-list.txt --list
```

### Interactive Mode

For a user-friendly interface:

```bash
node doc2web-run.js
```

### Processing Find Command Output

```bash
find . -name "*.docx" > docx-files.txt
./process-find.sh docx-files.txt
```

## Output

All processed files are stored in the `./output` directory, preserving the original directory structure:

```bash
output/
└── original/
    └── path/
        ├── filename.md      # Markdown version
        ├── filename.html    # HTML version with styles
        ├── filename.css     # Extracted CSS styles
        └── images/          # Extracted images
```

## Options

- `--html-only`: Generate only HTML output, skip markdown
- `--list`: Treat the input file as a list of files to process

## Recent Fixes

### v1.0.5 (2025-05-19)

- Fixed hierarchical list numbering in document conversion:
  - Properly maintains outline numbering structure (1., a., b., c., 2., etc.)
  - Correctly nests sub-items within parent items in HTML output
  - Preserves the original document's hierarchical list structure
  - Ensures consistent numbering throughout converted documents
  - Handles special cases like "Rationale for Resolution" paragraphs between list items
  - Verified with test file `input/TEST-TWO.docx` containing complex hierarchical lists

### v1.0.4 (2025-05-19)

- Enhanced TOC and index handling:
  - Automatically detects table of contents and index elements in DOCX files
  - Properly decorates these elements in the output with appropriate styling
  - Prevents unnecessary content duplication in web output
  - Improves readability by properly formatting navigation elements in web formats

### v1.0.3 (2025-05-16)

- Refactored CSS handling to improve separation of concerns:
  - Moved all CSS from inline `<style>` tags to external CSS files
  - Updated HTML files to use `<link>` tags to reference external CSS
  - Improved file organization and reduced HTML file size
  - Added 20px margin to body element for better readability
- Fixed URL handling in markdown generation:
  - Properly escaped special characters in URLs (parentheses, ampersands, etc.)
  - Improved handling of empty link text
  - Enhanced compatibility with markdown parsers and linters
- Fixed list formatting in markdown generation:
  - Added proper blank lines before and after lists (MD032)
  - Ensured consistent spacing in list items
  - Improved overall markdown structure and readability

### v1.0.2 (2025-05-16)

- Enhanced `markdownify.js` to fix markdown linting issues in generated files:
  - Fixed hard tabs, trailing spaces, and trailing punctuation in headings
  - Ensured proper spacing around list markers and correct ordered list numbering
  - Added proper blank lines around headings
  - Fixed spaces inside emphasis markers
  - Ensured first line is a top-level heading and prevented multiple top-level headings
  - Improved handling of empty links

### v1.0.1 (2025-05-16)

- Fixed an issue in `docx-style-parser.js` where undefined border values could cause errors
- Improved XML namespace handling for better compatibility with different DOCX file formats
- Added proper error handling for XPath queries to prevent crashes during style extraction

## Troubleshooting

### XML Namespace Errors

If you encounter errors like `Cannot resolve QName w` or `Cannot resolve QName a`, this indicates an issue with XML namespace resolution in the DOCX file. These errors have been fixed in v1.0.1.

### Border Style Errors

If you see errors related to border values or styles, ensure you're using the latest version which includes fixes for handling undefined or missing border properties.

### General Issues

1. Make sure your DOCX file is not corrupted or password-protected
2. Check that you have all required dependencies installed
3. For large files, consider processing them individually rather than in batch mode

## Contributing

1. Fork the repository
2. Create your feature branch
3. Commit your changes
4. Push to the branch
5. Create a new Pull Request

## License

This project is licensed under the MIT License.
