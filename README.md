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
- **Document header extraction and processing with formatting preservation**

The app must not make any assumptions from test documents, the app must treat created css and html as ephemeral, they will be destroyed on every run.
The css and HTML are individual to each document created, they will be named after the docx input, with folder pattern matched.
The app must deeply inspect the docx and obtain css and formatting styles, font sizes, item prefixes etc. Nothing will be hard coded with the app. All generated CSS will be in an external stylesheet.

## Project Structure

The codebase follows a modular architecture with organized library modules:

```bash
doc2web/
├── doc2web.js             # Main entry point and orchestrator
├── markdownify.js         # HTML to Markdown converter
├── lib/                   # Modular library code
│   ├── index.js           # Main entry point that re-exports public API
│   ├── xml/               # XML parsing utilities
│   ├── parsers/           # DOCX parsing modules
│   ├── html/              # HTML processing modules
│   │   ├── generators/    # HTML generation utilities
│   │   └── processors/    # Content processing modules
│   ├── css/               # CSS generation modules
│   │   └── generators/    # CSS generation utilities
│   ├── accessibility/     # Accessibility enhancement modules
│   └── utils/             # Utility functions
```

For detailed technical architecture information, see [`docs/architecture.md`](docs/architecture.md).

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

### Document Header Extraction and Processing

doc2web automatically extracts and processes document headers from DOCX files, placing them before the Table of Contents in the HTML output:

#### Header Extraction Strategies

1. **Header XML Files**: Extracts content from actual Word header files (`header1.xml`, `header2.xml`, `header3.xml`)
2. **Document Analysis**: Falls back to analyzing document content for header-like material at the beginning
3. **Header Types**: Supports different header types (first page, even pages, default headers)

#### Features

- **Formatting Preservation**: Maintains original fonts, sizes, colors, alignment, and styling
- **Graphics Inclusion**: Preserves images and graphics within headers with proper positioning
- **Image Positioning**: Honors original DOCX image positioning including inline placement, margins, and transform offsets
- **Responsive Design**: Headers adapt to different screen sizes and print media
- **Accessibility**: Proper semantic HTML with ARIA roles for screen readers
- **Content Detection**: Uses intelligent scoring to identify header content based on:
  - Document position and formatting characteristics
  - Style names and paragraph properties
  - Presence of images or special formatting</search>
</search_and_replace>

#### Technical Implementation

- **DOCX Introspection**: Analyzes actual Word document XML structure including image positioning data
- **Style Preservation**: Converts DOCX formatting to equivalent CSS
- **Image Processing**: Extracts images from DOCX archives and processes positioning information from XML
- **Positioning Fidelity**: Honors inline vs anchored positioning, distance attributes, and transform offsets
- **Semantic HTML**: Creates proper `<header>` elements with accessibility attributes and image containers
- **CSS Generation**: Generates responsive styles for different header types and image positioning
- **Content-Agnostic**: Works with any document regardless of language or domain</search>
</search_and_replace>

#### Recent Fixes (v1.2.8)

- **Correct Placement**: Headers now appear at the very beginning of the HTML document, before the TOC and any other content
- **Page Numbering Filter**: Automatically filters out page numbering text like "Page 1 of 5", "Page xx of xx", etc.
- **Clean Content**: Only meaningful header content is displayed, excluding automatic page numbering
- **Professional Structure**: Document order is now: Header → TOC → Document Content

This ensures that document headers appear correctly in the converted HTML, maintaining the professional appearance and context of the original Word document.

### Hanging Margins and Text Indentation

doc2web now properly handles hanging margins (also known as hanging indents) from DOCX documents, replicating Microsoft Word's text formatting behavior:

#### Features

- **TOC Hanging Indents**: Table of Contents entries display with proper hanging margins where long entries wrap with subsequent lines indented under the first line of text
- **Numbered Content Hanging Margins**: Numbered headings and paragraphs display with proper hanging indents where the number appears at the left margin and text content is indented
- **DOCX XML Introspection**: Extracts precise hanging indent measurements from DOCX XML structure
- **CSS Implementation**: Converts DOCX hanging indents to appropriate CSS using `text-indent` and `padding-left` properties
- **Enhanced Text Wrapping**: Comprehensive word wrapping support with `overflow-wrap`, `word-wrap`, and `hyphens` for proper line breaking

#### Technical Implementation

- **Block Layout for TOC**: Uses block layout instead of flex for proper hanging indent behavior
- **Negative Text Indent**: Implements hanging indents using negative `text-indent` values with corresponding `padding-left`
- **Number Positioning**: Positions numbers using CSS `::before` pseudo-elements with precise `left` positioning
- **Adaptive Logic**: Automatically adjusts hanging indent values when DOCX values are too small for proper display
- **Clean Display**: Optional removal of TOC dots and page numbers for cleaner appearance while maintaining hanging indents

#### Example CSS Output

```css
/* TOC entries with hanging indents */
.docx-toc-entry {
  text-indent: -1.5em;
  padding-left: 1.5em;
  overflow-wrap: break-word;
  word-wrap: break-word;
  hyphens: auto;
}

/* Numbered content with hanging margins */
.docx-numbering-level-0 {
  text-indent: -36pt;
  padding-left: 36pt;
}

.docx-numbering-level-0::before {
  left: -36pt;
  position: absolute;
}
```

This implementation ensures that converted HTML documents maintain the same professional text formatting and readability as the original Word documents.

### Technical Features

doc2web implements advanced DOCX introspection and content preservation:

- **Exact Numbering**: Extracts precise numbering definitions from DOCX XML structure
- **Content Preservation**: Comprehensive DOM serialization with fallback mechanisms
- **Structure Fidelity**: Maintains hierarchical relationships from original documents
- **Modular Architecture**: Clean, focused modules for improved maintainability

For detailed technical implementation information, see [`docs/architecture.md`](docs/architecture.md).

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
  const result = await extractAndApplyStyles(docxPath);
  console.log('Conversion completed:', result.html ? 'Success' : 'Failed');
}

convertDocument('document.docx').catch(console.error);
```

For detailed API options and configuration, see [`docs/architecture.md`](docs/architecture.md).

## Recent Fixes and Enhancements

### Recent Updates

**v1.3.1 (2025-05-31)**

- **Hanging Margins Fix**: Fixed hanging indent implementation for both TOC entries and main content
- **TOC Layout Improvements**: Converted TOC from flex to block layout for proper text wrapping with hanging indents
- **Enhanced Text Wrapping**: Added comprehensive word wrapping support with `overflow-wrap`, `word-wrap`, and `hyphens`
- **Clean TOC Display**: Removed dots and page numbers from TOC for cleaner appearance while maintaining hanging indents
- **Numbering Alignment**: Fixed hanging indent logic to ensure proper negative text-indent values for numbered content
- **Header Image Extraction**: Implemented comprehensive header image extraction and positioning functionality
- **Image Positioning**: Added DOCX XML introspection to extract and honor image positioning information from original documents
- **Semantic HTML**: Enhanced image output with proper container structure and accessibility attributes

**v1.3.0 (2025-05-26)**

- **Modular Refactoring**: Split large files into focused, maintainable modules
- **Enhanced Architecture**: Improved code organization with clear separation of concerns
- **API Preservation**: Maintained backward compatibility while improving internal structure
- **Better Maintainability**: Smaller, focused modules for easier development and testing

**v1.2.8 (2025-05-23)**

- Added comprehensive document header extraction and processing
- Fixed TOC duplicate numbering issue using CSS-only approach
- Enhanced page numbering filtering for cleaner header content

**v1.2.7 (2025-05-23)**

- Enhanced document parser to extract page margins and settings
- Improved document layout fidelity

**v1.2.6 (2025-05-23)**

- Fixed missing paragraph numbers and subheader letters in TOC and document
- Enhanced numbering display across all heading levels

**v1.2.5 (2025-05-23)**

- Fixed character overlap and numbering display issues
- Added special handling for Roman numerals

**v1.2.4 (2025-05-23)**

- Added section IDs for direct navigation to numbered headings
- Enhanced accessibility and navigation capabilities

**v1.2.0-1.2.3 (2025-05-22/23)**

- Comprehensive DOCX introspection for exact numbering preservation
- Enhanced TOC implementation with improved visual fidelity
- Fixed critical DOM serialization and content preservation issues

For detailed technical implementation information, see [`docs/architecture.md`](docs/architecture.md).

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

7. **Hanging Margins and Text Indentation Issues**
   - **Symptom**: TOC entries truncating instead of wrapping with hanging indents, or numbered content not displaying proper hanging margins
   - **Solution**: Update to v1.3.1 which fixes hanging indent implementation for both TOC and main content
   - **Diagnostic**: Check the HTML for proper `text-indent` and `padding-left` CSS properties
   - **Fix**: Verify that TOC uses block layout and numbered content has proper negative text-indent values
   - **TOC Specific**: Ensure TOC entries use `text-indent: -1.5em` and `padding-left: 1.5em` with block display
   - **Numbered Content**: Verify numbered elements have adequate hanging indents (typically `-36pt` text-indent with `36pt` padding-left)

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
