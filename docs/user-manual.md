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
11. [Accessibility Features](#accessibility-features)
12. [Metadata Preservation](#metadata-preservation)
13. [Track Changes Support](#track-changes-support)
14. [DOCX Introspection](#docx-introspection)
15. [API Usage](#api-usage)
16. [Architecture Overview](#architecture-overview)
17. [Troubleshooting](#troubleshooting)
18. [FAQ](#faq) 

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
- Accurately extracts and renders Table of Contents (TOC) elements with proper leader lines and page number alignment
- Maintains hierarchical list structures with proper nesting, indentation, and numbering
- Provides special handling for complex document structures and formatting patterns
- Ensures WCAG 2.1 Level AA accessibility compliance
- Preserves document metadata (title, author, creation date, etc.)
- Supports track changes visualization and handling
- **Extracts exact numbering and formatting from DOCX XML structure**
- **Modular architecture for improved maintainability and extensibility**
- **Generates properly formatted HTML with indentation and line breaks for improved readability and debugging**
- **Enhanced table formatting with professional styling, semantic structure, and accessibility features**
- **Bullet point enhancement with proper display and indentation from DOCX documents**

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
  - `--accessibility=<level>`: Set accessibility compliance level ('A', 'AA', or 'AAA', default: 'AA')
  - `--preserve-metadata`: Enable metadata preservation (default: enabled)
  - `--no-metadata`: Disable metadata preservation
  - `--track-changes=<mode>`: Set track changes mode ('show', 'hide', 'accept', or 'reject', default: 'show')
  - `--show-author`: Show change author information (default: enabled)
  - `--show-date`: Show change date information (default: enabled)

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
- **Properly formatted with indentation and line breaks for improved readability and debugging**
- **Professional HTML structure following web development best practices**

### CSS File (.css)

- Contains styles extracted from the document
- Implements DOCX styling in web-friendly format
- Linked from the HTML file
- Includes specific styles for TOC and hierarchical lists
- Contains detailed styling for leader lines and page number alignment
- Preserves indentation and formatting for different list levels
- **Contains CSS counters that exactly match DOCX numbering definitions**
- **Includes enhanced bullet point styles with high CSS specificity for proper display and indentation**

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

### Enhanced Table Formatting

doc2web now provides comprehensive table formatting improvements:

- **Professional Styling**: Clean borders, consistent padding, and alternating row colors
- **Semantic Structure**: Proper `<thead>` and `<tbody>` sections with `<th>` elements for headers
- **Accessibility Features**: Header cells with `scope="col"` attributes and table captions
- **Responsive Design**: Tables adapt to different screen sizes with horizontal scrolling
- **Hover Effects**: Visual feedback when hovering over table rows
- **Automatic Header Detection**: Intelligently converts first row to proper table headers
- **Modern CSS**: Professional appearance with borders, padding, and visual hierarchy

### Table of Contents and Index Handling

doc2web intelligently detects and handles navigation elements:

- Automatically identifies Table of Contents (TOC) sections
- Detects Index sections at the end of documents
- Properly decorates these elements with appropriate styling in the output
- In HTML output: Elements are styled with the appropriate CSS classes
- In Markdown output: Elements are formatted with proper structure
- Preserves TOC leader lines (dots, dashes, or underscores) connecting entries to page numbers
- Maintains proper right alignment of page numbers for better readability
- Applies appropriate indentation for different TOC levels
- Preserves TOC structure through advanced DOM manipulation
- Ensures visual fidelity to the original document's TOC appearance

This feature prevents unnecessary duplication of navigation elements in web formats while maintaining their visual structure, improving readability and organization.

### Hierarchical List Handling

doc2web provides advanced handling of hierarchical lists:

- Maintains proper nesting of multi-level lists (1., a., b., c., 2., etc.)
- Preserves numbering formats (decimal, alphabetic, roman numerals)
- Correctly indents sub-items based on their level
- Handles special cases like "Rationale for Resolution" paragraphs between list items
- Ensures consistent numbering throughout the document
- Applies appropriate styling for different list levels
- **Extracts exact numbering definitions from DOCX XML structure**
- **Resolves actual sequential numbers based on document position**
- **Maintains hierarchical relationships from original document**
- **Generates CSS counters that precisely match DOCX numbering**
- **Ensures proper bullet point display with enhanced CSS specificity to override inline styles**

## Accessibility Features

doc2web ensures that generated HTML meets WCAG 2.1 Level AA accessibility standards:

### Semantic Structure

- Proper HTML5 sectioning elements (header, main, nav, aside, footer)
- ARIA landmark roles (banner, navigation, main, contentinfo)
- Proper heading hierarchy (h1-h6) with no skipped levels

### Tables

- Table captions
- Header cells with scope attribute
- Row and column headers
- Complex tables with proper headers and ids
- ARIA labels for table purpose

### Images

- Alt text for all images
- Descriptive alt text based on image context
- Empty alt for decorative images
- Figure and figcaption for images with captions

### Navigation

- Skip navigation links for keyboard users
- Keyboard focus indicators
- Logical tab order
- ARIA navigation landmarks

### Color and Contrast

- Sufficient color contrast (4.5:1 for normal text, 3:1 for large text)
- Color not used as the only means of conveying information
- High contrast mode support
- Reduced motion support

## Metadata Preservation

doc2web extracts and preserves document metadata in the generated HTML:

### Standard Metadata

- Title, description, and keywords
- Creation and modification dates
- Document statistics (pages, words, characters)
- Author is always set to "doc2web" (regardless of original document author)

### Enhanced Metadata

- Dublin Core metadata
- Open Graph and Twitter Card metadata for social sharing
- JSON-LD structured data for search engines

### Metadata Usage

The preserved metadata improves:
- Search engine optimization (SEO)
- Social media sharing
- Document organization and cataloging
- Content discovery

## Track Changes Support

doc2web handles tracked changes in documents:

### Visualization Modes

- Show changes mode (display all insertions, deletions, and formatting changes)
- Hide changes mode (show the document without any change indicators)
- Accept all changes mode (incorporate all insertions, remove all deletions)
- Reject all changes mode (remove all insertions, keep all deletions)

### Change Indicators

- Insertions shown with underline or highlighting
- Deletions shown with strikethrough
- Formatting changes shown with highlight
- Moves shown with special indicators

### Change Metadata

- Author information
- Date and time of change
- Original and modified content
- Change type indicators

### Accessibility Considerations

- ARIA attributes for screen readers
- Non-visual indicators for changes
- Keyboard shortcut (Alt+T) to toggle track changes visibility

## DOCX Introspection

doc2web now uses comprehensive DOCX introspection to extract exact numbering and formatting information directly from the DOCX XML structure:

### Numbering Extraction

- **Parses complete numbering definitions from `numbering.xml`**
- **Extracts abstract numbering definitions with all level properties**
- **Captures level text formats (e.g., "%1.", "%1.%2.", "(%1)")**
- **Extracts indentation, alignment, and formatting for each level**
- **Handles level overrides and start value modifications**
- **Parses run properties (font, size, color) for numbering text**

### Numbering Resolution

- **Extracts paragraph numbering context from `document.xml`**
- **Maps paragraphs to their numbering definitions**
- **Resolves actual sequential numbers based on document position**
- **Handles restart logic and level overrides**
- **Tracks numbering sequences across the document**

### HTML Generation

- **Matches HTML elements to DOCX paragraph contexts**
- **Applies exact numbering from DOCX to headings**
- **Creates structured lists based on DOCX numbering definitions**
- **Maintains hierarchical relationships from original document**

### CSS Generation

- **Generates CSS counters from DOCX numbering definitions**
- **Creates level-specific styling for numbered elements**
- **Supports hierarchical numbering with proper resets**
- **Maintains visual fidelity to original DOCX formatting**

### Benefits

- **Exact Numbering Preservation**: Extracts the exact numbering format defined in the DOCX
- **Content-Agnostic Processing**: No reliance on specific content words or patterns
- **Complex Hierarchy Support**: Handles multi-level numbering with proper nesting
- **Visual Fidelity**: Generated HTML closely matches original DOCX appearance

## API Usage

For developers who want to integrate doc2web into their own applications, the tool provides a JavaScript API:

### Basic Usage

```javascript
const { extractAndApplyStyles } = require('./lib');

async function convertDocument(docxPath) {
  const result = await extractAndApplyStyles(docxPath);
  console.log('HTML generated:', result.html);
  console.log('CSS generated:', result.styles);
}

convertDocument('document.docx').catch(console.error);
```

### Advanced Options

```javascript
const { extractAndApplyStyles } = require('./lib');

async function convertDocument(docxPath) {
  // Configure options
  const options = {
    enhanceAccessibility: true,     // Enable accessibility features
    preserveMetadata: true,         // Enable metadata preservation
    trackChangesMode: 'show',       // 'show', 'hide', 'accept', or 'reject'
    showAuthor: true,               // Show change author information
    showDate: true                  // Show change date information
  };

  const result = await extractAndApplyStyles(docxPath, null, options);
  
  // Access conversion results
  console.log('HTML generated:', result.html);
  console.log('CSS generated:', result.styles);
  console.log('Metadata extracted:', result.metadata); // Note: Author will always be "doc2web"
  console.log('Track changes detected:', result.trackChanges.hasTrackedChanges);
}

convertDocument('document.docx').catch(console.error);
```

## Architecture Overview

doc2web features a modular architecture designed for maintainability and extensibility:

### Modular Design (v1.3.0)

The application has been refactored into focused, single-responsibility modules:

#### CSS Generation Modules (`lib/css/generators/`)
- **base-styles.js**: Document defaults and font utilities
- **paragraph-styles.js**: Word-like paragraph formatting
- **character-styles.js**: Inline text styling
- **table-styles.js**: Table layout and borders
- **numbering-styles.js**: CSS counters, hierarchical numbering, and bullet point display with enhanced CSS specificity
- **toc-styles.js**: Table of Contents with flex layout
- **utility-styles.js**: General utility classes
- **specialized-styles.js**: Accessibility and track changes styles

#### HTML Processing Modules (`lib/html/`)
- **generators/**: Style mapping, image processing, HTML formatting with proper indentation
- **processors/**: Heading, TOC, and numbering processors
- **Main orchestrator**: Coordinates all HTML generation with enhanced formatting

#### Benefits of Modular Architecture
- **Improved Maintainability**: Each module has a focused purpose
- **Better Testing**: Components can be tested in isolation
- **Enhanced Readability**: Smaller, focused files
- **API Preservation**: Backward compatibility maintained
- **Scalability**: Easy to add new features

For detailed technical information, see [`docs/architecture.md`](docs/architecture.md).

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
- **"Error parsing DOCX styles"**: The document may have non-standard or corrupted style definitions.
- **"Error selecting nodes with expression"**: There may be issues with the XML structure in the DOCX file.
- **"Error parsing numbering definitions"**: The document may have corrupted numbering.xml or non-standard numbering definitions.

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

**Q: How does doc2web handle Table of Contents and Index sections?**  
A: These navigation elements are automatically detected and properly decorated with appropriate styling in the output. This maintains their visual structure while preventing unnecessary duplication in web formats.

**Q: How does doc2web handle hierarchical lists?**  
A: With the new DOCX introspection feature, doc2web extracts exact numbering definitions from the DOCX XML structure, resolves actual sequential numbers based on document position, and generates CSS counters that precisely match the original formatting.

**Q: Can doc2web convert to PDF?**  
A: No, only Markdown and HTML outputs are currently supported.

### Technical Questions

**Q: What Node.js version is required?**  
A: Node.js 14.x or higher is recommended.

**Q: Does doc2web support DOCX files created by applications other than Microsoft Word?**  
A: Yes, it should work with any standard DOCX files, including those created by LibreOffice, Google Docs, etc.

**Q: Is doc2web available as an npm package?**  
A: Not yet, but you can use it directly from the source code.

**Q: How does the modular architecture benefit users?**  
A: The modular architecture improves maintainability, makes the code easier to understand, and allows for better testing and debugging. While users don't directly interact with the modules, they benefit from more reliable and maintainable software.

**Q: Is the generated HTML properly formatted for debugging?**  
A: Yes, doc2web now generates properly formatted HTML with indentation and line breaks, making it easy to read, debug, and inspect. The HTML follows web development best practices while preserving all document content and structure.

**Q: How does doc2web handle table formatting from DOCX documents?**  
A: doc2web now provides comprehensive table formatting with professional styling, semantic HTML structure, and accessibility features. Tables automatically get proper `<thead>` and `<tbody>` sections, header detection, responsive design, and modern CSS styling with borders, padding, and hover effects.

---

For further assistance or to report issues, please open an issue on the GitHub repository.
