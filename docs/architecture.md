# doc2web Architecture

## 1. Overview

doc2web is a Node.js application that converts Microsoft Word documents (.docx) to web-friendly formats (HTML and Markdown) while preserving the original document's styling and structure. This document provides a detailed technical explanation of how the application parses DOCX files and generates the corresponding web content.

The application is designed to be generic, content-agnostic, and document-structure-driven. It analyzes DOCX files based on their XML structure rather than making assumptions about specific content patterns or terms. All formatting decisions are based on document styles, not text content.

## 2. Application Architecture

The application follows a modular architecture organized into logical function groups:

```bash
doc2web/
├── doc2web.js             # Main entry point and orchestrator
├── markdownify.js         # HTML to Markdown converter
├── style-extractor.js     # Backward compatibility wrapper
├── docx-style-parser.js   # Backward compatibility wrapper
├── lib/                   # Refactored library code
│   ├── index.js           # Main entry point that re-exports public API
│   ├── xml/               # XML parsing utilities
│   │   └── xpath-utils.js # XPath utilities for XML processing
│   ├── parsers/           # DOCX parsing modules
│   │   ├── style-parser.js      # Style parsing functions
│   │   ├── theme-parser.js      # Theme parsing functions
│   │   ├── toc-parser.js        # TOC parsing functions
│   │   ├── numbering-parser.js  # Numbering definition parsing
│   │   ├── document-parser.js   # Document structure parsing
│   │   ├── metadata-parser.js   # Document metadata parsing
│   │   └── track-changes-parser.js # Track changes extraction
│   ├── html/              # HTML processing modules
│   │   ├── html-generator.js    # Main HTML generation
│   │   ├── structure-processor.js # HTML structure handling
│   │   ├── content-processors.js  # Content element processing
│   │   └── element-processors.js  # HTML element processing
│   ├── css/               # CSS generation modules
│   │   ├── css-generator.js     # CSS generation functions
│   │   └── style-mapper.js      # Style mapping functions
│   ├── accessibility/     # Accessibility enhancement modules
│   │   └── wcag-processor.js    # WCAG 2.1 compliance processor
│   └── utils/             # Utility functions
│       ├── unit-converter.js    # Unit conversion utilities
│       └── common-utils.js      # Common utility functions

```

## 3. DOCX File Structure

A DOCX file is essentially a ZIP archive containing various XML files that describe the document's content, styles, and metadata. Understanding this structure is crucial for parsing DOCX files effectively.

### 3.1 Key XML Files

The application extracts and processes the following XML files from the DOCX archive:

1. **word/document.xml**
   - Contains the main content of the document
   - Includes paragraphs, runs (text segments), tables, and other content elements
   - References styles and numbering definitions by ID
   - Contains tracked changes information (insertions, deletions, formatting changes)

2. **word/styles.xml**
   - Contains style definitions
   - Includes paragraph styles, character styles, table styles, and list styles
   - Defines properties like fonts, sizes, colors, indentation, etc.

3. **word/numbering.xml**
   - Contains numbering definitions for lists
   - Defines the format, indentation, and text of each level in a list
   - Referenced by paragraphs that are part of lists

4. **word/theme/theme1.xml**
   - Contains theme information (colors, fonts)
   - Referenced by styles to apply consistent theming

5. **word/settings.xml**
   - Contains document-wide settings
   - Includes default tab stops, character spacing, etc.
   - Contains track changes settings and revision information

6. **docProps/core.xml**
   - Contains core document metadata
   - Includes title, subject, author, keywords, created date, modified date, etc.
   - Follows Dublin Core metadata standard

7. **docProps/app.xml**
   - Contains application-specific metadata
   - Includes application name, company, revision count, etc.

### 3.2 XML Namespaces

The XML files in DOCX use several namespaces to organize elements and attributes. The application handles these namespaces using a custom XPath utility that registers the following namespaces:

```javascript
const NAMESPACES = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  m: 'http://schemas.openxmlformats.org/officeDocument/2006/math',
  v: 'urn:schemas-microsoft-com:vml',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006'
};
```

## 4. Core Processing Flow

The main processing flow of the application consists of the following steps:

1. **Input Processing**
   - Determine input type (single file, directory, or list file)
   - Validate input files (check if they are .docx files)
   - Create output directories mirroring the input structure

2. **DOCX Parsing**
   - Unpack the DOCX file (which is a ZIP archive)
   - Extract and parse key XML files (document.xml, styles.xml, etc.)
   - Extract style information, theme, numbering definitions, etc.

3. **HTML Generation**
   - Convert DOCX content to HTML using mammoth.js with custom style mapping
   - Enhance the HTML with proper structure and styling
   - Process special elements like TOC and hierarchical lists

4. **CSS Generation**
   - Generate CSS from the extracted style information
   - Create external CSS file linked from the HTML

5. **Image Extraction**
   - Extract images from the DOCX file
   - Save them in an images directory
   - Update image references in the HTML

6. **Markdown Conversion** (optional)
   - Convert the HTML to well-structured Markdown
   - Fix common Markdown formatting issues

7. **Output Organization**
   - Save all generated files in the appropriate output directories
   - Maintain the original directory structure

## 5. DOCX Parsing Process

### 5.1 Document XML Parsing

The application uses JSZip to extract XML files from the DOCX archive and xmldom's DOMParser to parse the XML into a DOM structure. The parsed XML is then processed using XPath queries to extract relevant information.

```javascript
// Extract key files
const styleXml = await zip.file('word/styles.xml')?.async('string');
const documentXml = await zip.file('word/document.xml')?.async('string');
const themeXml = await zip.file('word/theme/theme1.xml')?.async('string');
const settingsXml = await zip.file('word/settings.xml')?.async('string');
const numberingXml = await zip.file('word/numbering.xml')?.async('string');
const corePropsXml = await zip.file('docProps/core.xml')?.async('string');
const appPropsXml = await zip.file('docProps/app.xml')?.async('string');

// Parse XML content
const styleDoc = new DOMParser().parseFromString(styleXml);
const documentDoc = new DOMParser().parseFromString(documentXml);
const themeDoc = themeXml ? new DOMParser().parseFromString(themeXml) : null;
const settingsDoc = settingsXml ? new DOMParser().parseFromString(settingsXml) : null;
const numberingDoc = numberingXml ? new DOMParser().parseFromString(numberingXml) : null;
const corePropsDoc = corePropsXml ? new DOMParser().parseFromString(corePropsXml) : null;
const appPropsDoc = appPropsXml ? new DOMParser().parseFromString(appPropsXml) : null;
```

### 5.2 Style Extraction

The application extracts style information from the parsed XML using XPath queries and custom parsing functions. The extracted information includes:

1. **Paragraph Styles**
   - Font family, size, weight, style
   - Alignment, indentation, spacing
   - Borders, shading, tabs

2. **Character Styles**
   - Font family, size, weight, style
   - Color, underline, highlight
   - Special text effects

3. **Table Styles**
   - Borders, spacing, alignment
   - Cell padding, background colors

4. **Theme Information**
   - Color schemes
   - Font schemes (major and minor fonts)

5. **Document Settings**
   - Default tab stops
   - Character spacing
   - Right-to-left settings

### 5.3 Numbering Definition Extraction

The application extracts numbering definitions from numbering.xml to handle hierarchical lists correctly. This includes:

1. **Abstract Numbering Definitions**
   - Base definitions for numbering formats
   - Level-specific properties (format, text, alignment, indentation)

2. **Numbering Instances**
   - Instances that reference abstract numbering definitions
   - Level overrides for specific instances

3. **Level Properties**
   - Numbering format (decimal, alpha, roman, etc.)
   - Level text (how the number appears, e.g., "%1.")
   - Alignment and indentation
   - Run properties (font, size, etc.)

### 5.4 Document Structure Analysis

The application analyzes the document structure to identify special elements and patterns without relying on specific content words. This includes:

1. **TOC Detection**
   - Identify TOC fields in document.xml
   - Extract TOC properties (leader character, heading levels)
   - Find TOC styles in style definitions

2. **List Structure Analysis**
   - Identify hierarchical relationships between list items
   - Detect non-numbered paragraphs within lists
   - Analyze paragraph properties for list formatting

3. **Paragraph Pattern Analysis**
   - Identify structural patterns (NOT content-specific patterns)
   - Look for formatting patterns in the document
   - Analyze style usage across the document

## 6. HTML Generation Process

### 6.1 Base HTML Generation

The application uses mammoth.js to convert DOCX to HTML with custom style mapping. The style mapping is generated based on the extracted style information:

```javascript
// Create a custom style map based on extracted styles
const styleMap = createStyleMap(styleInfo);

// Use mammoth to convert with the custom style map
const result = await mammoth.convertToHtml({
  path: docxPath,
  styleMap: styleMap,
  transformDocument: transformDocument,
  includeDefaultStyleMap: true,
  ...imageOptions,
});
```

### 6.2 HTML Enhancement

The generated HTML is enhanced with proper structure and styling:

1. **Ensure HTML Structure**
   - Add proper html, head, and body elements
   - Add meta charset tag and language attribute
   - Add link to external CSS file
   - Incorporate document metadata as meta tags

2. **Process Elements**
   - Enhance tables with responsive wrappers and accessibility attributes
   - Process images with proper attributes and alt text
   - Handle language-specific elements (RTL text, etc.)
   - Add ARIA roles and attributes for accessibility

3. **Process Headings**
   - Add hierarchical numbering
   - Maintain proper heading structure
   - Ensure accessible heading hierarchy

4. **Process TOC**
   - Create properly structured TOC
   - Add text, dots, and page number spans
   - Maintain hierarchical levels
   - Ensure keyboard navigability

5. **Process Lists**
   - Handle hierarchical list structures
   - Maintain proper nesting and numbering
   - Process special paragraphs within lists
   - Ensure proper list semantics for screen readers

6. **Process Track Changes**
   - Represent insertions, deletions, and formatting changes visually
   - Include change metadata (author, date) as data attributes
   - Add CSS for displaying changes
   - Provide option to view document with or without changes

## 7. CSS Generation Process

The application generates CSS from the extracted style information:

1. **Document Defaults**
   - Body font family, size, line height
   - Margin and padding
   - Focus state styles for keyboard navigation

2. **Paragraph Styles**
   - Font properties
   - Alignment, indentation, spacing
   - Borders and backgrounds

3. **Character Styles**
   - Font properties
   - Text decoration and effects

4. **Table Styles**
   - Border collapse, width, margins
   - Cell padding, borders, backgrounds
   - Table captions and headers
   - Accessibility enhancements (e.g., zebra striping)

5. **TOC Styles**
   - Container styles
   - Entry styles for different levels
   - Leader dot styling
   - Focus styles for keyboard navigation

6. **List Styles**
   - List containers and items
   - Counter reset and increment
   - Level-specific indentation and formatting

7. **Track Changes Styles**
   - Insertion styles (typically underlined or highlighted)
   - Deletion styles (typically strikethrough)
   - Formatting change indicators
   - Change metadata display

8. **Accessibility Styles**
   - High contrast mode support
   - Focus indication
   - Skip navigation links
   - Print-specific styles

9. **Utility Styles**
   - Underline, strike-through, etc.
   - RTL text direction
   - Image and table defaults
   - Screen reader-only text

## 8. Special Element Handling

### 8.1 Table of Contents (TOC)

The application carefully processes TOC elements to maintain their structure and styling:

1. **TOC Detection**
   - Identify TOC entries based on style or content pattern
   - Create a TOC container if not already present
   - Group TOC entries by level

2. **TOC Structure**
   - Create spans for text, dots, and page numbers
   - Apply appropriate classes for styling
   - Maintain hierarchical structure with proper indentation

3. **TOC Styling**
   - Style the TOC container and heading
   - Style TOC entries with proper indentation
   - Create leader dots using CSS background images
   - Align page numbers to the right

### 8.2 Hierarchical Lists

The application processes hierarchical lists to maintain their structure and numbering:

1. **List Detection**
   - Identify list items based on numbering properties
   - Group list items into logical lists
   - Detect nested lists and their levels

2. **List Structure**
   - Create ordered or unordered lists as appropriate
   - Maintain proper nesting for hierarchical lists
   - Create list items with proper attributes

3. **List Styling**
   - Apply CSS counters for automatic numbering
   - Style different list levels appropriately
   - Handle special cases like different numbering formats

### 8.3 Language-Specific Elements

The application handles language-specific elements:

1. **RTL Text**
   - Detect right-to-left text direction
   - Apply appropriate directionality attributes
   - Style RTL elements with proper CSS

2. **Language Detection**
   - Detect script types (Latin, Cyrillic, Chinese, Arabic, etc.)
   - Add appropriate language attributes
   - Apply language-specific styling

## 9. Markdown Conversion

The application converts the generated HTML to Markdown using the markdownify.js module:

1. **HTML Parsing**
   - Parse the HTML using jsdom
   - Create a DOM structure for processing

2. **Element Processing**
   - Process different HTML elements (headings, paragraphs, lists, etc.)
   - Convert them to the appropriate Markdown syntax
   - Maintain document structure

3. **Special Element Handling**
   - Handle tables with proper column alignment
   - Process images with alt text
   - Handle code blocks and blockquotes

4. **Markdown Linting Fixes**
   - Fix common Markdown linting issues
   - Ensure proper spacing around headings
   - Fix list marker spacing
   - Fix emphasis spacing

## 10. Output Organization

The application organizes output files in a structured manner:

1. **Directory Structure**
   - Mirror the input directory structure
   - Create output directories as needed

2. **Output Files**
   - HTML file with the same base name as the input
   - CSS file with the same base name
   - Markdown file with the same base name
   - Images directory for extracted images

3. **File Naming**
   - Maintain the original file name (replacing the extension)
   - Use consistent naming for related files

## 11. Error Handling and Recovery

The application includes comprehensive error handling:

1. **Input Validation**
   - Check if files exist
   - Validate file types (must be .docx)
   - Create output directories if they don't exist

2. **DOCX Parsing Errors**
   - Handle malformed XML
   - Provide fallback style information if parsing fails
   - Continue processing other files in batch mode

3. **HTML Generation Errors**
   - Handle mammoth.js conversion errors
   - Recover from DOM manipulation errors
   - Provide fallback HTML if necessary

4. **File System Errors**
   - Handle errors reading from or writing to the file system
   - Create directories recursively
   - Check file permissions

## 12. Processing Units and Data Flow

The following diagram illustrates the data flow through the application:

```bash
Input DOCX File
    │
    ▼
┌───────────────┐
│ DOCX Unpacker │ (JSZip)
└───────┬───────┘
        │
        ▼
┌───────────────────┐
│ XML Parser        │ (xmldom)
└────────┬──────────┘
         │
         ▼
┌────────────────────────┐
│ Style Extractor        │
│  - Paragraph Styles    │
│  - Character Styles    │
│  - Table Styles        │
│  - Theme Information   │
│  - Numbering Definitions│
└────────────┬───────────┘
             │
             ▼
┌─────────────────────┐    ┌───────────────────┐
│ HTML Generator      │    │ CSS Generator     │
│ (mammoth.js +       │───▶│  - Document CSS   │
│  custom processing) │    │  - TOC Styles     │
└────────┬────────────┘    │  - List Styles    │
         │                 └─────────┬─────────┘
         ▼                           │
┌────────────────────┐               │
│ HTML Enhancer      │               │
│  - Structure       │               │
│  - TOC Processing  │               │
│  - List Processing │               │
└────────┬───────────┘               │
         │                           │
         ▼                           ▼
┌────────────────────┐    ┌────────────────────┐
│ Markdown Converter │    │ External CSS File  │
│ (markdownify.js)   │    │                    │
└────────┬───────────┘    └────────────────────┘
         │
         ▼
┌────────────────────┐
│ Output Files       │
│  - HTML            │
│  - Markdown        │
│  - CSS             │
│  - Images          │
└────────────────────┘
```

## 13. Accessibility Implementation

The application includes comprehensive accessibility features to ensure WCAG 2.1 Level AA compliance:

### 13.1 Accessibility Processing

The `wcag-processor.js` module enhances HTML output for accessibility:

```javascript
function processForAccessibility(document, styleInfo, metadata) {
  // Add language attribute to HTML element
  const htmlElement = document.documentElement;
  const documentLang = styleInfo.settings?.language || 'en';
  htmlElement.setAttribute('lang', documentLang);
  
  // Add document title
  const title = document.querySelector('title');
  if (title) {
    title.textContent = metadata?.core?.title || 'Document';
  }
  
  // Process tables for accessibility
  processTables(document);
  
  // Process images for accessibility
  processImages(document);
  
  // Ensure proper heading hierarchy
  ensureHeadingHierarchy(document);
  
  // Add ARIA landmarks
  addAriaLandmarks(document);
  
  // Add skip navigation link
  addSkipNavigation(document);
  
  // Enhance keyboard navigability
  enhanceKeyboardNavigation(document);
  
  // Check and enhance color contrast
  enhanceColorContrast(document, styleInfo);
  
  return document;
}
```

### 13.2 Key Accessibility Enhancements

1. **Semantic Structure**
   - Proper HTML5 sectioning elements (header, main, nav, aside, footer)
   - ARIA landmark roles (banner, navigation, main, contentinfo)
   - Proper heading hierarchy (h1-h6) with no skipped levels

2. **Tables**
   - Table captions
   - Header cells with scope attribute
   - Row and column headers
   - Complex tables with proper headers and ids
   - ARIA labels for table purpose

3. **Images**
   - Alt text for all images
   - Descriptive alt text based on image context
   - Empty alt for decorative images
   - Figure and figcaption for images with captions

4. **Forms**
   - Labels for all form controls
   - Error messages and validation
   - Fieldset and legend for form groups
   - ARIA roles and states

5. **Navigation**
   - Skip navigation links
   - Keyboard focus indicators
   - Logical tab order
   - ARIA navigation landmarks

6. **Color and Contrast**
   - Sufficient color contrast (4.5:1 for normal text, 3:1 for large text)
   - Color not used as the only means of conveying information
   - High contrast mode support

7. **List and TOC Enhancements**
   - Proper list semantics (ul, ol, li)
   - Proper nesting and hierarchy
   - ARIA attributes for custom lists
   - Keyboard navigable TOC

## 14. Metadata Implementation

The application includes comprehensive metadata processing:

### 14.1 Metadata Extraction and Application

The `metadata-parser.js` module extracts and applies document metadata:

```javascript
function applyMetadataToHtml(document, metadata) {
  // Add standard meta tags
  const head = document.head;
  
  // Add title
  let titleElement = head.querySelector('title');
  if (!titleElement) {
    titleElement = document.createElement('title');
    head.appendChild(titleElement);
  }
  titleElement.textContent = metadata.core.title || 'Document';
  
  // Add description
  if (metadata.core.description) {
    const metaDesc = document.createElement('meta');
    metaDesc.setAttribute('name', 'description');
    metaDesc.setAttribute('content', metadata.core.description);
    head.appendChild(metaDesc);
  }
  
  // Add author (always set to "doc2web" regardless of original document author)
  const metaAuthor = document.createElement('meta');
  metaAuthor.setAttribute('name', 'author');
  metaAuthor.setAttribute('content', 'doc2web');
  head.appendChild(metaAuthor);
  
  // Add keywords
  if (metadata.core.keywords) {
    const metaKeywords = document.createElement('meta');
    metaKeywords.setAttribute('name', 'keywords');
    metaKeywords.setAttribute('content', metadata.core.keywords);
    head.appendChild(metaKeywords);
  }
  
  // Add Dublin Core metadata
  addDublinCoreMetadata(head, metadata.core);
  
  // Add Open Graph metadata
  addOpenGraphMetadata(head, metadata.core);
  
  // Add JSON-LD structured data
  addJsonLdStructuredData(head, metadata);
  
  return document;
}
```

### 14.2 Key Metadata Features

1. **HTML Meta Tags**
   - Standard meta tags (title, description, author, keywords)
   - Dublin Core metadata
   - Open Graph metadata
   - Twitter Card metadata

2. **JSON-LD Structured Data**
   - Schema.org Article or Document type
   - Author, publisher, and date information
   - Keywords and categories
   - Document statistics (word count, page count)

3. **Application-Specific Metadata**
   - Original document information
   - Application and template information
   - Conversion timestamp
   - Document statistics preservation

## 15. Track Changes Implementation

The application includes comprehensive track changes processing:

### 15.1 Track Changes Processing

The `track-changes-parser.js` module extracts and processes tracked changes:

```javascript
function processTrackChanges(document, changes, options) {
  const { mode } = options; // 'show', 'hide', 'accept', or 'reject'
  
  // If no tracked changes, return the document unchanged
  if (!changes.hasTrackedChanges) {
    return document;
  }
  
  // Add track changes mode class to body
  document.body.classList.add(`docx-track-changes-${mode}`);
  
  // Process insertions
  if (mode === 'show' || mode === 'accept') {
    changes.insertions.forEach(insertion => {
      const elements = document.querySelectorAll(`[data-change-id="${insertion.id}"]`);
      elements.forEach(element => {
        if (mode === 'show') {
          element.classList.add('docx-insertion');
          element.setAttribute('data-author', insertion.author);
          element.setAttribute('data-date', insertion.date);
        } else {
          // For 'accept' mode, remove the track changes marking but keep the content
          element.classList.remove('docx-insertion');
          element.removeAttribute('data-author');
          element.removeAttribute('data-date');
          element.removeAttribute('data-change-id');
        }
      });
    });
  }
  
  // Process deletions
  if (mode === 'show' || mode === 'reject') {
    changes.deletions.forEach(deletion => {
      const elements = document.querySelectorAll(`[data-change-id="${deletion.id}"]`);
      elements.forEach(element => {
        if (mode === 'show') {
          element.classList.add('docx-deletion');
          element.setAttribute('data-author', deletion.author);
          element.setAttribute('data-date', deletion.date);
        } else {
          // For 'reject' mode, remove the track changes marking but keep the content
          element.classList.remove('docx-deletion');
          element.removeAttribute('data-author');
          element.removeAttribute('data-date');
          element.removeAttribute('data-change-id');
        }
      });
    });
  }
  
  // Similar processing for moves and formatting changes
  // ...
  
  return document;
}
```

### 15.2 Key Track Changes Features

1. **Visualization Modes**
   - Show changes mode (display all insertions, deletions, and formatting changes)
   - Hide changes mode (show the document without any change indicators)
   - Accept all changes mode (incorporate all insertions, remove all deletions)
   - Reject all changes mode (remove all insertions, keep all deletions)

2. **Change Indicators**
   - Insertions shown with underline or highlighting
   - Deletions shown with strikethrough
   - Formatting changes shown with highlight
   - Moves shown with special indicators

3. **Change Metadata**
   - Author information
   - Date and time of change
   - Original and modified content
   - Change type indicators

4. **Accessibility Considerations**
   - ARIA attributes for screen readers
   - Non-visual indicators for changes
   - Keyboard shortcuts for navigating changes

## 16. Key Algorithms

### 16.1 Style Extraction Algorithm

```
function parseDocxStyles(docxPath):
    1. Read DOCX file and extract key XML files
    2. Parse XML files into DOM structures
    3. Extract styles from styles.xml
    4. Extract theme information from theme/theme1.xml
    5. Extract document defaults from styles.xml
    6. Extract document settings from settings.xml
    7. Extract TOC styles by analyzing styles and document structure
    8. Extract numbering definitions from numbering.xml
    9. Extract metadata from docProps/core.xml and docProps/app.xml
    10. Extract tracked changes information from document.xml
    11. Analyze document structure to identify patterns and special elements
    12. Return combined style information
```

### 16.2 CSS Generation Algorithm

```
function generateCssFromStyleInfo(styleInfo):
    1. Generate document defaults CSS (body styles)
    2. For each paragraph style:
        a. Generate a CSS class with appropriate properties
    3. For each character style:
        a. Generate a CSS class with appropriate properties
    4. For each table style:
        a. Generate CSS classes for tables and cells
    5. Generate TOC styles:
        a. Style TOC container and heading
        b. Style TOC entries with proper levels and indentation
        c. Create leader dot styles
    6. Generate list styles:
        a. Create counter reset and increment rules
        b. Style list items for each level
        c. Create pseudo-elements for numbering
    7. Generate track changes styles:
        a. Style insertion, deletion, and formatting change indicators
        b. Style change metadata display
    8. Generate accessibility styles:
        a. Focus indicators and skip links
        b. High contrast mode support
        c. Screen reader enhancements
    9. Add utility styles for common elements
    10. Return combined CSS
```

### 16.3 TOC Processing Algorithm

```
function processTOC(document, styleInfo):
    1. Find TOC entries (by class or content pattern)
    2. If TOC entries found:
        a. Create or find TOC container
        b. Add ARIA role="navigation" and proper landmarks
        c. For each TOC entry:
            i. Determine TOC level
            ii. Extract text content and page number
            iii. Create structured elements (text, dots, page number)
            iv. Apply appropriate classes and ARIA attributes
            v. Ensure keyboard navigability
            vi. Move to TOC container
    3. Return enhanced document
```

### 16.4 Hierarchical List Processing Algorithm

```
function processNestedNumberedParagraphs(document, styleInfo):
    1. Identify list patterns in the document
    2. Process paragraphs that look like list items:
        a. Match with different list patterns (main numbers, alpha, roman, etc.)
        b. Determine list type, prefix, and level
        c. Create/continue appropriate list structure
        d. Create list items with proper attributes and ARIA roles
        e. Handle special paragraphs within lists
        f. Ensure proper list semantics for accessibility
    3. Return enhanced document with structured lists
```

### 16.5 Metadata Processing Algorithm

```
function processMetadata(document, metadata):
    1. Extract metadata from docProps/core.xml and docProps/app.xml
    2. Add standard HTML meta tags (title, description, keywords, etc.)
    3. Add Dublin Core metadata
    4. Add Open Graph metadata
    5. Add Twitter Card metadata
    6. Create JSON-LD structured data
    7. Add metadata to HTML head
    8. Return enhanced document
```

### 16.6 Track Changes Processing Algorithm

```
function processTrackChanges(document, changes, options):
    1. Determine track changes mode (show, hide, accept, reject)
    2. Add mode-specific class to document body
    3. Process insertions:
        a. Find insertion elements by change ID
        b. Apply appropriate styling based on mode
        c. Add metadata (author, date) as attributes
    4. Process deletions:
        a. Find deletion elements by change ID
        b. Apply appropriate styling based on mode
        c. Add metadata (author, date) as attributes
    5. Process moves:
        a. Find move elements by change ID
        b. Apply appropriate styling based on mode
        c. Add metadata (source, destination) as attributes
    6. Process formatting changes:
        a. Find formatting change elements by change ID
        b. Apply appropriate styling based on mode
        c. Add metadata (before, after) as attributes
    7. Return enhanced document
```

### 16.7 Accessibility Processing Algorithm

```
function processForAccessibility(document, styleInfo, metadata):
    1. Add language attribute to HTML element
    2. Add title and metadata
    3. Process headings:
        a. Ensure proper heading hierarchy
        b. Add ARIA attributes for complex headings
    4. Process tables:
        a. Add table captions
        b. Add header cells with scope
        c. Add ARIA labels for complex tables
    5. Process images:
        a. Add appropriate alt text
        b. Add figure and figcaption elements
    6. Add ARIA landmarks:
        a. Banner (header)
        b. Navigation (TOC, menus)
        c. Main (document content)
        d. Complementary (asides)
        e. Contentinfo (footer)
    7. Add skip navigation link
    8. Enhance keyboard navigability:
        a. Add tabindex where needed
        b. Ensure logical tab order
        c. Add focus styles
    9. Enhance color contrast and non-visual indicators
    10. Return enhanced document
```

## 17. Conclusion

The doc2web application provides a robust solution for converting DOCX documents to web-friendly formats while preserving styling and structure. By analyzing the document's XML structure rather than its content, the application remains generic and content-agnostic, working effectively with any document regardless of its domain or purpose.

The enhanced implementation now includes comprehensive accessibility features ensuring WCAG 2.1 Level AA compliance, detailed metadata extraction and preservation, and robust track changes handling. These enhancements make the application suitable for professional and enterprise environments where accessibility, metadata, and revision tracking are critical requirements.

The modular architecture and clear separation of concerns make the codebase maintainable and extensible, while the comprehensive error handling ensures reliable operation even with complex or problematic documents.
