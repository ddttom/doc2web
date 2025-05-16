# doc2web

doc2web is a powerful tool for converting Microsoft Word (.docx) documents to web-friendly formats, including Markdown and HTML with preserved styling.

## Overview

This tool extracts content from DOCX files while maintaining:

- Document structure and formatting
- Images and tables
- Styles and layout
- Unicode and multilingual content

## Project Structure

```bash
doc2web/
├── doc2web.js               # Main converter script
├── markdownify.js           # HTML to Markdown converter
├── style-extractor.js       # Document style extraction
├── docx-style-parser.js     # DOCX XML structure parser
├── process-find.sh          # Helper for processing file lists
├── doc2web-run.js           # Interactive CLI interface
├── doc2web-install.sh       # Installation script
├── init-doc2web.sh          # Project initialization script
└── README.md                # Documentation
```

## Getting Started

### Initial Setup

1. Clone or download all project files to a directory

 ```bash

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

## Contributing

1. Fork the repository
2. Create your feature branch
3. Commit your changes
4. Push to the branch
5. Create a new Pull Request

## License

This project is licensed under the MIT License.
