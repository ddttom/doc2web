# doc2web

doc2web is a powerful tool for converting Microsoft Word (.docx) documents to web-friendly formats, including Markdown and HTML with preserved styling.

## Overview

This tool extracts content from DOCX files while maintaining:

- Document structure and formatting
- Images and tables
- Styles and layout
- Unicode and multilingual content
- Table of Contents (TOC) with proper styling
- Hierarchical list numbering and structure

The app must not make any assumptions from test documents, the app must treat created css and html as ephemeral, they will be destroyed on every run.
The css and HTML are individual to each document created, they will be named after the docx input, with folder pattern matched.
The app must deeply inspect the docx and obtain css and formatting styles, font sizes, item prefixes etc. Nothing will be hard coded with the app. All generated CSS will be in an external stylesheet.

## Project Structure

The codebase has been refactored for better organization and maintainability:

```
doc2web/
├── .gitignore
├── .vscode
├── README.md
├── doc2web-install.sh
├── doc2web-run.js
├── doc2web.js
├── docs/
│   ├── prd.md                  # Product Requirements Document
│   ├── refactoring.md          # Detailed refactoring documentation
│   └── user-guide.md           # User guide for the refactored code
├── init-doc2web.sh
├── input/
├── lib/                        # Refactored library code
│   ├── index.js                # Main entry point
│   ├── xml/                    # XML parsing utilities
│   │   └── xpath-utils.js      # XPath utilities for XML processing
│   ├── parsers/                # DOCX parsing modules
│   │   ├── style-parser.js     # Style parsing functions
│   │   ├── theme-parser.js     # Theme parsing functions
│   │   ├── toc-parser.js       # TOC parsing functions
│   │   ├── numbering-parser.js # Numbering definition parsing
│   │   └── document-parser.js  # Document structure parsing
│   ├── html/                   # HTML processing modules
│   │   ├── html-generator.js   # Main HTML generation
│   │   ├── structure-processor.js # HTML structure handling
│   │   ├── content-processors.js  # Content element processing
│   │   └── element-processors.js  # HTML element processing
│   ├── css/                    # CSS generation modules
│   │   ├── css-generator.js    # CSS generation functions
│   │   └── style-mapper.js     # Style mapping functions
│   └── utils/                  # Utility functions
│       ├── unit-converter.js   # Unit conversion utilities
│       └── common-utils.js     # Common utility functions
├── markdownify.js
├── node_modules/
├── output/
├── package-lock.json
├── package.json
├── process-find.sh
├── style-extractor.js          # Backward compatibility wrapper
└── docx-style-parser.js        # Backward compatibility wrapper
```

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

## Code Organization

The codebase has been refactored into a modular structure:

- **XML Utilities** (`lib/xml/`): Functions for working with XML and XPath
  - `xpath-utils.js`: Provides utilities for XPath queries with proper namespace handling

- **Parsers** (`lib/parsers/`): 
  - `style-parser.js`: Parses DOCX styles and extracts formatting information
  - `theme-parser.js`: Extracts theme information (colors, fonts)
  - `toc-parser.js`: Parses Table of Contents styles and structure
  - `numbering-parser.js`: Extracts numbering definitions for lists
  - `document-parser.js`: Analyzes document structure and settings

- **HTML Processing** (`lib/html/`):
  - `html-generator.js`: Main module for generating HTML from DOCX
  - `structure-processor.js`: Ensures proper HTML document structure
  - `content-processors.js`: Processes headings, TOC, and lists
  - `element-processors.js`: Processes tables, images, and language elements

- **CSS Generation** (`lib/css/`):
  - `css-generator.js`: Generates CSS from extracted style information
  - `style-mapper.js`: Maps DOCX styles to CSS classes

- **Utilities** (`lib/utils/`):
  - `unit-converter.js`: Converts between different units (twips, points)
  - `common-utils.js`: Common utility functions shared across modules

## Documentation

- **README.md**: This file, providing an overview of the project
- **docs/prd.md**: Product Requirements Document with detailed specifications
- **docs/refactoring.md**: Detailed documentation of the refactoring process
- **docs/user-guide.md**: User guide for working with the refactored code

## Recent Fixes

### v1.0.7 (2025-05-20)

- Refactored codebase for better organization and maintainability:
  - Split large files into smaller, focused modules
  - Improved code organization by logical function groups
  - Enhanced documentation with comprehensive comments
  - Maintained backward compatibility with existing code
  - Fixed circular dependencies and improved error handling
  - Added detailed documentation for the refactored code

### v1.0.6 (2025-05-20)

- Enhanced DOCX style extraction and processing:
  - Improved parsing of Table of Contents (TOC) styles with better leader line handling
  - Added robust numbering definition extraction for complex hierarchical lists
  - Implemented comprehensive document structure analysis
  - Enhanced error handling for XML parsing and style extraction
  - Improved CSS generation for TOC and list styling
  - Added better detection and styling of special document sections

- Enhanced TOC (Table of Contents) Processing:
  - Proper leader line rendering with dots connecting entries to page numbers
  - Right-aligned page numbers for better readability
  - Correct indentation for different TOC levels
  - Preservation of TOC structure through advanced DOM manipulation
  - Better visual fidelity to the original document's TOC appearance

- Improved List Handling:
  - Hierarchical list numbering (1., a., b., etc.) maintained with proper nesting
  - Proper indentation for different list levels
  - Special handling for "Rationale for Resolution" paragraphs within lists
  - Better recognition of list structures through content pattern analysis
  - Consistent numbering throughout converted documents

- More Robust Style Extraction:
  - Enhanced DOCX style parser that captures more details from the document
  - Better handling of special paragraph types
  - Improved detection of document structure patterns
  - More accurate conversion of DOCX styles to CSS
  - Deeper inspection of document formatting attributes

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

### TOC and List Structure Issues

If you notice problems with Table of Contents formatting or hierarchical list structures, make sure you're using v1.0.6 or later, which includes significant improvements to TOC style extraction and list numbering handling.

### General Issues

1. Make sure your DOCX file is not corrupted or password-protected
2. Check that you have all required dependencies installed
3. For large files, consider processing them individually rather than in batch mode
4. For complex documents with many styles, ensure you have sufficient memory allocated

## Contributing

1. Fork the repository
2. Create your feature branch
3. Commit your changes
4. Push to the branch
5. Create a new Pull Request

## License

This project is licensed under the MIT License.
