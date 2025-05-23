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
- **WCAG 2.1 Level AA accessibility compliance**
- **Document metadata preservation**
- **Track changes visualization and handling**
- **Exact DOCX numbering preservation through XML introspection**
- **Reliable DOM serialization with content preservation**
- **Section IDs for direct navigation to numbered headings and paragraphs**

The app must not make any assumptions from test documents, the app must treat created css and html as ephemeral, they will be destroyed on every run.
The css and HTML are individual to each document created, they will be named after the docx input, with folder pattern matched.
The app must deeply inspect the docx and obtain css and formatting styles, font sizes, item prefixes etc. Nothing will be hard coded with the app. All generated CSS will be in an external stylesheet.

## Project Structure

The codebase has been refactored for better organization and maintainability:

```bash
doc2web/
├── .gitignore
├── README.md
├── doc2web-install.sh
├── doc2web-run.js
├── doc2web.js
├── debug-test.js           # Diagnostic tool for troubleshooting
├── docs/
│   ├── prd.md                  # Product Requirements Document
│   ├── refactoring.md          # Detailed refactoring documentation
│   ├── user-guide.md           # User guide for the refactored code
│   └── architecture.md         # Technical architecture documentation
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
│   │   ├── numbering-resolver.js # Numbering resolution engine
│   │   ├── document-parser.js  # Document structure parsing
│   │   ├── metadata-parser.js  # Document metadata parsing
│   │   └── track-changes-parser.js # Track changes extraction
│   ├── html/                   # HTML processing modules
│   │   ├── html-generator.js   # Main HTML generation
│   │   ├── structure-processor.js # HTML structure handling
│   │   ├── content-processors.js  # Content element processing
│   │   └── element-processors.js  # HTML element processing
│   ├── css/                    # CSS generation modules
│   │   ├── css-generator.js    # CSS generation functions
│   │   └── style-mapper.js     # Style mapping functions
│   ├── accessibility/          # Accessibility enhancement modules
│   │   └── wcag-processor.js   # WCAG 2.1 compliance processor
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

### Diagnostic Testing

For troubleshooting conversion issues:

```bash
node debug-test.js path/to/document.docx
```

This will create a debug-output/ directory with test files and diagnostic information.

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

## Enhanced Features

### Section IDs and Navigation

doc2web automatically generates section IDs for numbered headings based on their hierarchical position:

- `1. Introduction` becomes `id="section-1"`
- `1.2 Overview` becomes `id="section-1-2"`
- `1.2.a Details` becomes `id="section-1-2-a"`

This enables:

- Direct linking to sections (e.g., `document.html#section-1-2-a`)
- Smooth scrolling navigation
- Accessibility improvements for screen readers
- Visual highlighting of targeted sections

Section IDs are derived directly from the DOCX numbering definitions, ensuring they match the exact hierarchical structure of the document regardless of language or content domain.

### DOCX Introspection for Exact Numbering

doc2web now extracts exact numbering and formatting information directly from the DOCX XML structure:

- Parses complete numbering definitions from `numbering.xml`
- Extracts level text formats (e.g., "%1.", "%1.%2.", "(%1)")
- Captures indentation, alignment, and formatting for each level
- Resolves actual sequential numbers based on document position
- Handles restart logic and level overrides
- Generates CSS counters that precisely match DOCX numbering
- Maintains hierarchical relationships from the original document

This ensures that complex numbered lists and headings appear exactly as they do in the original document, regardless of language or content domain.

### DOM Serialization and Content Preservation

doc2web now implements comprehensive DOM serialization verification to ensure content preservation during HTML processing:

- Verifies document body content before serialization
- Implements fallback mechanisms for empty body issues
- Preserves all document structure during DOM manipulation
- Logs serialization metrics for debugging purposes
- Handles browser-specific DOM serialization differences
- Ensures proper nesting and hierarchical relationships
- Implements error handling for serialization failures

This ensures that all document content is properly preserved during the conversion process, preventing content loss or corruption that can occur during complex DOM manipulations.

### Accessibility Compliance (WCAG 2.1 Level AA)

doc2web now ensures that generated HTML meets WCAG 2.1 Level AA accessibility standards:

- Proper semantic structure with HTML5 sectioning elements and ARIA landmarks
- Accessible tables with captions, header cells, and proper scope attributes
- Images with appropriate alt text and figure/figcaption elements
- Proper heading hierarchy with no skipped levels
- Skip navigation links for keyboard users
- Keyboard focus indicators and proper tab order
- High contrast mode support
- Reduced motion support
- Screen reader compatibility

To enable accessibility features (enabled by default):

```javascript
// When using the API
const { extractAndApplyStyles } = require('./lib');
const options = { enhanceAccessibility: true };
const result = await extractAndApplyStyles('document.docx', null, options);
```

### Metadata Preservation

doc2web now extracts and preserves document metadata in the generated HTML:

- Title, description, and keywords
- Creation and modification dates
- Document statistics (pages, words, characters)
- Author is always set to "doc2web" (regardless of original document author)
- Dublin Core metadata
- Open Graph and Twitter Card metadata for social sharing
- JSON-LD structured data for search engines

To enable metadata preservation (enabled by default):

```javascript
// When using the API
const { extractAndApplyStyles } = require('./lib');
const options = { preserveMetadata: true };
const result = await extractAndApplyStyles('document.docx', null, options);
```

### Track Changes Support

doc2web now handles tracked changes in documents:

- Visual representation of insertions, deletions, moves, and formatting changes
- Multiple view modes (show changes, hide changes, accept all, reject all)
- Author and date information for each change
- Track changes legend with toggle functionality
- Keyboard shortcut (Alt+T) to toggle track changes visibility

To configure track changes handling:

```javascript
// When using the API
const { extractAndApplyStyles } = require('./lib');
const options = {
  trackChangesMode: 'show', // 'show', 'hide', 'accept', or 'reject'
  showAuthor: true,
  showDate: true
};
const result = await extractAndApplyStyles('document.docx', null, options);
```

## API Usage

```javascript
const { extractAndApplyStyles } = require('./lib');

async function convertDocument(docxPath) {
  // Default options
  const options = {
    enhanceAccessibility: true,
    preserveMetadata: true,
    trackChangesMode: 'show',
    showAuthor: true,
    showDate: true,
    verifyDomSerialization: true,
    logSerializationMetrics: false,
    enableSerializationFallbacks: true
  };

  const result = await extractAndApplyStyles(docxPath, null, options);
  console.log('HTML generated:', result.html);
  console.log('CSS generated:', result.styles);
  console.log('Metadata extracted:', result.metadata);
  console.log('Track changes detected:', result.trackChanges.hasTrackedChanges);
}

convertDocument('document.docx').catch(console.error);
```

## Recent Fixes and Enhancements

### v1.2.6 (2025-05-23)

- Fixed missing paragraph numbers and subheader letters in TOC and document:
  - Resolved issue where heading numbers were not displaying in the document
  - Fixed missing paragraph numbers and subheader letters in the Table of Contents
  - Added code to populate the heading-number span with actual numbering content
  - Enhanced TOC entry processing to include numbering in the text content
  - Improved anchor creation for better navigation between TOC and document sections
  - Added debug test script to verify numbering display without rebuilding documents
  - Maintained accessibility attributes for screen readers
  - Ensured consistent display of numbering across all heading levels and TOC entries

### v1.2.5 (2025-05-23)

- Fixed character overlap and numbering display issues:
  - Resolved issues with paragraph numbers or letters from IDs not displaying correctly
  - Fixed character overlap in paragraphs with indentation
  - Improved spacing and positioning for numbered elements
  - Added special handling for Roman numerals to ensure proper display
  - Enhanced box model handling to prevent content overflow
  - Improved spacing between numbering and paragraph content
  - Fixed line wrapping issues with indented paragraphs
  - Added specific CSS rules for Roman numeral sections
  - Implemented consistent box model across all numbered elements

### v1.2.4 (2025-05-23)

- Added section IDs for direct navigation to numbered headings and paragraphs:
  - Implemented automatic generation of section IDs based on hierarchical numbering
  - Added CSS styling for section navigation and highlighting
  - Enhanced heading accessibility with proper ID attributes
  - Improved TOC linking to section IDs
  - Added diagnostic tools for section ID validation
  - Updated documentation with section ID usage examples

### v1.2.3 (2025-05-23)

- Comprehensive TOC implementation fixes:
  - Added CSS specificity and importance improvements with `!important` declarations for critical flex properties
  - Implemented DOM structure validation to ensure all TOC entries have the complete three-part structure
  - Enhanced leader dots implementation with consistent sizing and better baseline alignment
  - Improved layout and container handling with box-sizing and max-width properties
  - Added browser compatibility fallbacks for gradient approaches
  - Fixed paragraph display with specific overrides for TOC entries
  - Added print-specific styles for better TOC appearance in printed documents
- Improved debugging capabilities:
  - Added detailed logging for TOC processing
  - Enhanced validation for TOC entry structure
  - Implemented more robust error handling for TOC-related operations
- Documentation updates:
  - Added comprehensive troubleshooting guidance for TOC issues
  - Updated PRD with detailed TOC implementation requirements
  - Enhanced code comments for TOC-related functions

### v1.2.2 (2025-05-23)

- Enhanced TOC and paragraph numbering display:
  - Improved Table of Contents formatting with proper leader dots and right-aligned page numbers
  - Enhanced paragraph numbering display that exactly matches the original document
  - Fixed "ragged" appearance of TOC entries for a cleaner, more professional look
  - Ensured proper alignment and spacing in hierarchical document elements
  - Improved visual fidelity to original Word documents
- Implemented advanced CSS techniques for document formatting:
  - Used flex-based layout for TOC entries to ensure proper alignment
  - Applied CSS ::before pseudo-elements for paragraph numbering
  - Created leader dots using CSS background-image with radial gradients
  - Positioned numbering elements with absolute positioning for exact placement
- Added detailed troubleshooting guidance for TOC and numbering issues

### v1.2.1 (2025-05-22)

- Fixed critical issues in the HTML generation pipeline:
  - Fixed content loss during DOM manipulation that was causing incomplete HTML output
  - Improved error handling with detailed, actionable error messages
  - Fixed DOM serialization problems that were causing content corruption
  - Re-enabled ARIA landmarks functionality that was disabled due to DOM manipulation errors
  - Added comprehensive validation at key processing steps
- Enhanced main application (doc2web.js):
  - Added better input file validation to prevent processing invalid files
  - Improved error reporting with context-specific messages
  - Enhanced logging to help identify conversion issues
  - Added performance timing and summary statistics
- Fixed accessibility processor (lib/accessibility/wcag-processor.js):
  - Implemented safer DOM manipulation that preserves content
  - Fixed ARIA landmarks functionality
  - Enhanced keyboard navigation features
- Added diagnostic tool (debug-test.js):
  - Provides comprehensive testing for diagnosing issues
  - Tests each component individually
  - Generates detailed diagnostic output
  - Creates test files for inspection
- Improved CSS-based heading numbering implementation:
  - Modified content-processors.js to use data attributes instead of direct HTML manipulation
  - Enhanced css-generator.js to generate precise CSS ::before pseudo-elements for numbering
  - Improved handling of indentation, margins, and text flow based on DOCX properties
  - Ensured proper positioning of numbering elements with absolute positioning
  - Prevented potential conflicts or duplicate numbering in headings
  - Fixed TOC display issues by explicitly preventing multi-column layout
  - Improved visual fidelity to original Word document numbering

### v1.2.0 (2025-05-22)

- Added comprehensive DOCX introspection for exact numbering:
  - Implemented enhanced numbering parser to extract complete definitions from numbering.xml
  - Created numbering resolution engine to calculate actual sequential numbers
  - Enhanced content processors to apply DOCX-derived numbering to HTML elements
  - Improved style parser to integrate numbering context into style extraction
  - Enhanced HTML generator to maintain numbering context through conversion
  - Added CSS generator support for DOCX numbering formats
  - Ensured content-agnostic processing that works with any language or domain
- Implemented reliable DOM serialization with content preservation:
  - Added verification of document body content before serialization
  - Implemented fallback mechanisms for empty body issues
  - Added serialization metrics logging for debugging purposes
  - Enhanced error handling for serialization failures
  - Ensured proper content preservation during DOM manipulation
  - Added support for handling browser-specific serialization differences

### v1.1.0 (2025-05-21)

- Added support for WCAG 2.1 Level AA accessibility compliance
- Implemented document metadata preservation
- Added track changes support with multiple viewing modes
- Enhanced API with configuration options for new features

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

### v1.0.5 (2025-05-19)

- Fixed hierarchical list numbering in document conversion:
  - Properly maintains outline numbering structure (1., a., b., c., 2., etc.)
  - Correctly nests sub-items within parent items in HTML output
  - Preserves the original document's hierarchical list structure

### v1.0.4 (2025-05-19)

- Enhanced TOC and index handling:
  - Automatically detects table of contents and index elements in DOCX files
  - Properly decorates these elements in the output with appropriate styling
  - Prevents unnecessary content duplication in web output

### v1.0.3 (2025-05-16)

- Refactored CSS handling to improve separation of concerns:
  - Moved all CSS from inline `<style>` tags to external CSS files
  - Updated HTML files to use `<link>` tags to reference external CSS
  - Improved file organization and reduced HTML file size
  - Added 20px margin to body element for better readability

### v1.0.2 (2025-05-16)

- Enhanced `markdownify.js` to fix markdown linting issues in generated files:
  - Fixed hard tabs, trailing spaces, and trailing punctuation in headings
  - Ensured proper spacing around list markers and correct ordered list numbering
  - Added proper blank lines around headings

### v1.0.1 (2025-05-16)

- Fixed an issue in `docx-style-parser.js` where undefined border values could cause errors
- Improved XML namespace handling for better compatibility with different DOCX file formats
- Added proper error handling for XPath queries to prevent crashes during style extraction

## Troubleshooting

### Using the Debug Script

If you encounter issues with document conversion, use the debug script first:

```bash
node debug-test.js path/to/your/document.docx
```

This will create a debug-output/ directory with:

- Test HTML and CSS files
- Detailed diagnostic information
- Component test results

### Checking for Successful Conversion

Look for these indicators of successful conversion:

- HTML files > 1000 characters for typical documents
- CSS files with generated styles
- Proper directory structure maintained
- Images extracted to images/ subdirectory

### XML Namespace Errors

If you encounter errors like `Cannot resolve QName w` or `Cannot resolve QName a`, this indicates an issue with XML namespace resolution in the DOCX file. These errors have been fixed in v1.0.1.

### Border Style Errors

If you see errors related to border values or styles, ensure you're using the latest version which includes fixes for handling undefined or missing border properties.

### TOC and List Structure Issues

If you notice problems with Table of Contents formatting or hierarchical list structures, make sure you're using v1.2.0 or later, which includes comprehensive DOCX introspection for exact numbering and formatting preservation.

### Numbering Format Issues

If you encounter issues with complex numbering formats (multi-level, mixed formats, etc.), ensure you're using v1.2.0 or later, which extracts exact numbering definitions from the DOCX XML structure rather than inferring them from text patterns.

### DOM Serialization Issues

If you encounter missing content, empty sections, or corrupted HTML output:

1. Check the console for serialization metrics and error messages
2. Ensure you're using v1.2.1 or later, which includes fixes for content preservation during DOM manipulation
3. For complex documents with extensive DOM manipulation, consider enabling detailed logging
4. If issues persist with specific documents, try processing them with the debug script

### Content Loss Issues

If you notice content missing from the generated HTML:

1. Update to v1.2.1 which fixes content loss during DOM manipulation
2. Check the error logs for specific issues
3. Use the debug script to identify exactly where content is being lost
4. Verify that the original DOCX file is not corrupted

### Accessibility Processing Errors

If accessibility features aren't working correctly:

1. Update to v1.2.1 which fixes issues in the accessibility processor
2. Check that the enhanceAccessibility option is enabled
3. Use the debug script to test the accessibility processor independently

### TOC and Numbering Display Issues

If you encounter issues with Table of Contents formatting or paragraph numbering display:

1. **TOC Formatting Problems**
   - **Symptom**: Missing leader dots or misaligned page numbers
   - **Solution**: Update to v1.2.2 which implements a flex-based layout for TOC entries
   - **Diagnostic**: Check the HTML structure for proper TOC entry elements with text, dots, and page number spans
   - **Fix**: Verify that the CSS contains proper TOC styles with flex layout and background-image for leader dots

2. **Paragraph Numbering Issues**
   - **Symptom**: Missing numbers or incorrect formatting in numbered paragraphs
   - **Solution**: Update to v1.2.2 which uses CSS ::before pseudo-elements for numbering
   - **Diagnostic**: Inspect the HTML for data-numbering-id and data-numbering-level attributes
   - **Fix**: Ensure the CSS contains proper counter-reset and counter-increment rules for numbering

3. **Hierarchical Numbering Problems**
   - **Symptom**: Incorrect hierarchical numbering (e.g., 1.1, 1.2, etc.)
   - **Solution**: Update to v1.2.2 which refines the counter reset strategy
   - **Diagnostic**: Run the debug script and examine the numbering definitions in debug-output/debug-info.json
   - **Fix**: Check that the CSS properly implements the hierarchical counter structure

4. **Section ID Issues**
   - **Symptom**: Missing or incorrect section IDs for headings and numbered paragraphs
   - **Solution**: Update to v1.2.4 which implements automatic section ID generation
   - **Diagnostic**: Run the debug script and check for section IDs in the HTML output
   - **Fix**: Verify that the HTML contains proper id attributes with the pattern "section-X-Y-Z"

5. **Character Overlap Issues**
   - **Symptom**: Characters overlapping in paragraphs with indentation or numbering
   - **Solution**: Update to v1.2.5 which fixes spacing and positioning for numbered elements
   - **Diagnostic**: Inspect the HTML for proper padding and margin values
   - **Fix**: Verify that the CSS contains proper box model properties for numbered elements

6. **Roman Numeral Display Issues**
   - **Symptom**: Roman numerals (I, II, V, VI, etc.) not displaying correctly in section headings
   - **Solution**: Update to v1.2.5 which adds special handling for Roman numerals
   - **Diagnostic**: Check the HTML for the roman-numeral-heading class
   - **Fix**: Ensure the CSS contains specific rules for Roman numeral sections

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
