# doc2web User Manual

## Table of Contents

1. [Introduction](#introduction)
2. [Installation](#installation)
3. [Basic Usage](#basic-usage)
4. [Advanced Features](#advanced-features)
5. [Table of Contents Navigation](#table-of-contents-navigation)
6. [Command Reference](#command-reference)
7. [Output Files](#output-files)
8. [Formatting Features](#formatting-features)
9. [Accessibility Features](#accessibility-features)
10. [Metadata Preservation](#metadata-preservation)
11. [Track Changes Support](#track-changes-support)
12. [DOCX Introspection](#docx-introspection)
13. [API Usage](#api-usage)
14. [Architecture Overview](#architecture-overview)
15. [Multilingual Support](#multilingual-support)
16. [Batch Processing](#batch-processing)
17. [Interactive Mode](#interactive-mode)
18. [Troubleshooting](#troubleshooting)
19. [FAQ](#faq)

## Introduction

doc2web is a powerful tool for converting Microsoft Word (.docx) documents to web-friendly HTML format. Unlike simple converters, doc2web preserves document styling, structure, images, and supports multilingual content including mixed scripts.

### Key Features

- Converts DOCX to HTML format
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
- **Comprehensive text formatting preservation including italic, bold, and mixed formatting**
- **Advanced CSS generation with rule conflict resolution for proper display**
- **TOC page number removal for web-appropriate navigation where page numbers are irrelevant**
- **Clickable TOC navigation with intelligent section linking and accessibility compliance**
- **Return to top button for improved navigation and user experience**
- **Proper bullet point margin alignment with document margins and table cell exclusion**
- **Enhanced bullet point formatting with correct hanging indent behavior and round bullet characters**

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

- An HTML version with styles (.html)
- A CSS file with extracted styles (.css)
- An images folder with any extracted images

### Output Location

All output files are stored under the `./output` directory, preserving the original directory structure. For example:

- Input: `/home/user/documents/report.docx`
- Output:
  - `./output/home/user/documents/report.html`
  - `./output/home/user/documents/report.css`
  - `./output/home/user/documents/images/`

## Advanced Features



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
- **Includes enhanced styling with high CSS specificity for proper display**
- **Contains comprehensive formatting styles with fallback rules and enhanced specificity**
- **Includes CSS rule conflict resolution with !important declarations for proper display**

### Images Folder

- Contains all images extracted from the document
- Images are referenced from the HTML

## Table of Contents Navigation

doc2web automatically converts Table of Contents sections into interactive navigation elements with the following features:

### Automatic TOC Detection and Processing

- **Smart Recognition**: Automatically identifies Table of Contents sections in DOCX documents
- **Page Number Removal**: Removes page numbers from TOC entries as they are irrelevant for web navigation
- **Semantic Marking**: Adds CSS classes and data attributes to TOC entries for reliable identification

### Intelligent Section Linking

- **Hierarchical Matching**: Uses multiple strategies to match TOC entries to document sections:
  - Section ID matching (e.g., `section-1`, `section-1-a`)
  - Text content comparison
  - Numbering pattern recognition
  - Fallback matching for edge cases
- **Context-Aware**: Understands parent-child relationships in hierarchical documents
- **Generic Implementation**: Works with any language or content domain without hardcoded patterns

### User Experience Features

- **Clickable Navigation**: All TOC entries become clickable links to their corresponding sections
- **Smooth Scrolling**: Animated scrolling to target sections for better user experience
- **Visual Feedback**: Hover and focus states provide clear interaction cues
- **Target Highlighting**: Briefly highlights the target section when navigated to

### Accessibility Compliance

- **WCAG 2.1 Level AA**: Fully compliant with web accessibility standards
- **ARIA Attributes**: Proper semantic markup for screen readers
- **Keyboard Navigation**: Full keyboard accessibility support
- **Focus Management**: Proper focus handling for assistive technologies

### Technical Implementation

- **Post-Processing**: TOC linking occurs after all HTML generation is complete
- **Multi-Strategy Matching**: Robust algorithm handles various document structures
- **Performance Optimized**: Efficient processing with minimal impact on conversion time
- **Error Resilient**: Graceful handling of malformed or unusual TOC structures

## Return to Top Button

doc2web automatically adds a return to top button to all generated HTML documents for improved navigation:

### Features

- **Always Visible**: The button remains visible in the bottom left corner of the viewport for easy access
- **Theme Integration**: Automatically styled to match the document's color scheme and fonts
- **Smooth Scrolling**: Provides animated scrolling to the top of the document for better user experience
- **Accessibility Compliant**: Full keyboard navigation support with proper ARIA labels and screen reader compatibility

### User Experience

- **Visual Feedback**: Hover and focus states provide clear interaction cues
- **Responsive Design**: Adapts to different screen sizes and devices
- **Consistent Positioning**: Fixed position ensures the button is always accessible regardless of document length
- **Professional Appearance**: Circular design with up arrow icon matches modern web standards

### Technical Implementation

- **Theme-Aware Styling**: Uses CSS variables derived from the document's theme colors
- **Performance Optimized**: Lightweight implementation with minimal impact on page load
- **Cross-Browser Compatible**: Works consistently across all modern browsers
- **Accessibility Standards**: Meets WCAG 2.1 Level AA compliance requirements

## Formatting Features

doc2web preserves the following elements and provides special handling for navigation elements:

### Document Structure

- Headings (H1-H6)
- Paragraphs with proper spacing
- Lists (ordered and unordered)
- Tables with formatting
- Images with captions

### Text Formatting

- Bold, italic, underline with comprehensive preservation from DOCX
- Strikethrough
- Superscript and subscript
- Font styles and sizes (in HTML output)
- Text colors (in HTML output)
- Mixed formatting combinations (bold + italic, underline + italic, etc.)

### Advanced Elements

- Hyperlinks
- Blockquotes
- Code blocks
- Horizontal rules

### Enhanced Table Formatting

doc2web provides comprehensive table formatting improvements:

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
- Preserves TOC leader lines (dots, dashes, or underscores) connecting entries to page numbers
- Maintains proper right alignment of page numbers for better readability
- Applies appropriate indentation for different TOC levels
- Preserves TOC structure through advanced DOM manipulation
- Ensures visual fidelity to the original document's TOC appearance
- **Automatically removes page numbers from TOC entries** since page numbers are irrelevant for web-based navigation
- **Targeted processing** that only scans TOC sections rather than the entire document for efficient operation
- **Smart boundary detection** to identify where TOC sections end and main content begins

This feature prevents unnecessary duplication of navigation elements in web formats while maintaining their visual structure, improving readability and organization. The automatic removal of page numbers makes the TOC more appropriate for web consumption where users navigate via hyperlinks rather than page references.

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
- **Ensures proper display with enhanced CSS specificity to override inline styles**
- **Fixes display issues with CSS rule conflict resolution and margin reset techniques**

### Bullet Point Formatting

doc2web provides comprehensive bullet point formatting that accurately reflects the original DOCX document:

#### Margin Alignment
- **Document Margin Integration**: Bullet points align with the document's standard margins extracted from DOCX introspection
- **Dynamic Calculation**: Uses the same margin calculation as regular paragraphs to ensure consistent alignment
- **Base Margin Respect**: Bullet lists use the document's base margin plus appropriate bullet indentation

#### Table Cell Exclusion
- **Smart Detection**: Automatically detects when content is inside table cells
- **Selective Application**: Only applies bullet formatting to content outside of tables
- **DOCX Fidelity**: Ensures table cell content appears as plain text when the original DOCX has no bullets in tables

#### Proper Hanging Indent
- **Word-Style Formatting**: Implements proper hanging indent behavior matching Microsoft Word's bullet formatting
- **Round Bullet Characters**: Uses appropriate round bullet characters (•, ○) instead of square characters
- **Text Readability**: Prevents aggressive word-breaking that can make bullet text difficult to read
- **CSS Specificity**: Uses high-specificity CSS rules to override conflicting DOCX abstract numbering styles

#### Technical Implementation
- **DOCX Introspection**: Analyzes the original DOCX structure to determine appropriate bullet formatting
- **Context-Aware Processing**: Distinguishes between actual bullet lists and table cell content
- **Enhanced CSS Generation**: Creates CSS with proper specificity to ensure correct display
- **Accessibility Compliance**: Maintains proper semantic structure for screen readers and assistive technologies

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
- Author metadata is configurable (preserves original author by default, can be overridden)

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

doc2web uses comprehensive DOCX introspection to extract exact numbering and formatting information directly from the DOCX XML structure:

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

### Basic API Usage

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
  console.log('Metadata extracted:', result.metadata); // Note: Author is configurable
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
- **character-styles.js**: Inline text styling including comprehensive italic formatting with fallback rules
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

The interactive mode guides you through each step, including selecting input files for HTML generation.

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
A: These navigation elements are automatically detected and properly decorated with appropriate styling in the output. This maintains their visual structure while preventing unnecessary duplication in web formats. Additionally, page numbers are automatically removed from TOC entries since they are irrelevant for web-based navigation where users click hyperlinks instead of referencing page numbers.

**Q: How does doc2web handle hierarchical lists?**  
A: With the DOCX introspection feature, doc2web extracts exact numbering definitions from the DOCX XML structure, resolves actual sequential numbers based on document position, and generates CSS counters that precisely match the original formatting.

**Q: Can doc2web convert to PDF?**  
A: No, only HTML output is currently supported.

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
A: Yes, doc2web generates properly formatted HTML with indentation and line breaks, making it easy to read, debug, and inspect. The HTML follows web development best practices while preserving all document content and structure.

**Q: Does doc2web add navigation aids to the generated HTML?**  
A: Yes, doc2web automatically adds a return to top button that appears in the bottom left corner of all generated HTML documents. The button is always visible, uses theme-aware styling to match the document's appearance, provides smooth scrolling to the top, and includes full accessibility support with keyboard navigation and ARIA labels.

**Q: How does doc2web handle formatting from DOCX documents?**  
A: doc2web provides comprehensive formatting preservation including table styling with professional appearance, semantic HTML structure, and accessibility features. Text formatting (italic, bold, mixed styles) is preserved through enhanced style mapping and CSS generation with proper specificity and conflict resolution.

---

For further assistance or to report issues, please open an issue on the GitHub repository.
