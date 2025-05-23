# doc2web Product Requirements Document

**Document Version:** 3.0  
**Last Updated:** May 23, 2025  
**Status:** Draft  
**Authors:** Technical Team  
**Approved By:** [Pending]  

## Document Revision History

| Version | Date | Author | Description of Changes |
|---------|------|--------|------------------------|
| 1.0 | May 16, 2025 | Initial Team | Original document |
| 1.5 | May 19, 2025 | Technical Team | Enhanced TOC and list handling requirements |
| 2.0 | May 20, 2025 | Technical Team | Added document introspection rules, removed content pattern matching |
| 2.1 | May 20, 2025 | Technical Team | Updated architecture to reflect modular refactoring |
| 2.2 | May 22, 2025 | Technical Team | Added comprehensive DOCX numbering introspection requirements |
| 2.3 | May 22, 2025 | Technical Team | Added DOM serialization requirements for numbering preservation |
| 2.4 | May 22, 2025 | Technical Team | Added critical fixes for DOM manipulation, error handling, and validation |
| 2.5 | May 22, 2025 | Technical Team | Added document statistics calculation and metadata improvements |
| 2.6 | May 22, 2025 | Technical Team | Added HTML formatting for improved debugging and readability |
| 2.7 | May 22, 2025 | Technical Team | Improved CSS-based numbering implementation for headings |
| 2.8 | May 22, 2025 | Technical Team | Fixed TOC layout and body content numbering issues |
| 2.9 | May 23, 2025 | Technical Team | Enhanced TOC formatting and paragraph numbering implementation (v1.2.2) |
| 3.0 | May 23, 2025 | Technical Team | Added section IDs for direct navigation to numbered headings (v1.2.4) |

## Table of Contents

- [doc2web Product Requirements Document](#doc2web-product-requirements-document)
  - [Document Revision History](#document-revision-history)
  - [Table of Contents](#table-of-contents)
  - [1. Product Overview](#1-product-overview)
    - [1.1 Vision Statement](#11-vision-statement)
    - [1.2 Product Goals](#12-product-goals)
    - [1.3 Project Scope](#13-project-scope)
  - [2. Target Users](#2-target-users)
    - [2.1 Primary User Personas](#21-primary-user-personas)
    - [2.2 Use Cases](#22-use-cases)
  - [3. Features and Requirements](#3-features-and-requirements)
    - [3.1 Core Features](#31-core-features)
      - [3.1.1 Document Conversion](#311-document-conversion)
      - [3.1.2 Style Preservation](#312-style-preservation)
      - [3.1.3 Processing Capabilities](#313-processing-capabilities)
      - [3.1.4 User Interface](#314-user-interface)
    - [3.2 Technical Requirements](#32-technical-requirements)
      - [3.2.1 Platform Support](#321-platform-support)
      - [3.2.2 Performance](#322-performance)
      - [3.2.3 Output Quality](#323-output-quality)
      - [3.2.4 Document Analysis and Style Extraction](#324-document-analysis-and-style-extraction)
  - [4. Technical Specifications](#4-technical-specifications)
    - [4.1 Architecture](#41-architecture)
    - [4.2 Dependencies](#42-dependencies)
    - [4.3 Output Specifications](#43-output-specifications)
    - [4.4 Document Analysis Rules](#44-document-analysis-rules)
    - [4.5 API and Integration](#45-api-and-integration)
    - [4.6 DOM Serialization Requirements](#46-dom-serialization-requirements)
    - [4.7 Error Handling and Validation](#47-error-handling-and-validation)
    - [4.8 Recent Implementation Changes](#48-recent-implementation-changes)
      - [4.8.1 CSS-Based Heading Numbering](#481-css-based-heading-numbering)
      - [4.8.2 TOC and Body Content Numbering Fixes](#482-toc-and-body-content-numbering-fixes)
      - [4.8.3 Enhanced TOC and Paragraph Numbering (v1.2.2)](#483-enhanced-toc-and-paragraph-numbering-v122)
      - [4.8.4 Comprehensive TOC Implementation Fixes (v1.2.3)](#484-comprehensive-toc-implementation-fixes-v123)
      - [4.8.5 Section IDs for Navigation (v1.2.4)](#485-section-ids-for-navigation-v124)
      - [4.8.6 Character Overlap and Numbering Display Fixes (v1.2.5)](#486-character-overlap-and-numbering-display-fixes-v125)
      - [4.8.7 Paragraph Numbering and TOC Display Fix (v1.2.6)](#487-paragraph-numbering-and-toc-display-fix-v126)

## 1. Product Overview

### 1.1 Vision Statement

doc2web transforms Microsoft Word documents (.docx) into web-friendly formats (Markdown and HTML) while preserving document styling, structure, and multilingual content. The tool aims to bridge the gap between traditional document authoring and modern web publishing by providing a reliable, efficient conversion process that maintains document fidelity.

### 1.2 Product Goals

- Enable seamless conversion of DOCX documents to both Markdown and HTML
- Preserve the visual appearance and style elements from the original document
- Support international and multilingual content
- Provide an intuitive interface for both individual and batch processing
- Create organized output that mirrors the original directory structure
- Accurately extract and render complex document elements like TOC and hierarchical lists
- Accessibility Compliance: Ensure HTML output meets WCAG 2.1 Level AA standards
- Metadata Preservation: Better extraction and preservation of document metadata, including accurate document statistics
- Track Changes Support: Handle documents with tracked changes appropriately
- **DOCX Introspection**: Extract exact numbering and formatting from DOCX XML structure
- **DOM Serialization**: Ensure proper preservation of document content during DOM manipulation
- **Robust Error Handling**: Provide comprehensive error reporting and recovery mechanisms
- **Content Validation**: Implement validation at key processing steps to ensure data integrity
- **Section Navigation**: Generate section IDs for direct navigation to numbered headings and paragraphs

### 1.3 Project Scope

doc2web is focused on the conversion of DOCX documents to HTML and Markdown formats. The scope includes:

- Processing individual files, directories of files, and file lists
- Extracting and preserving styles, images, and document structure
- Maintaining fidelity to the original document layout
- Providing both interactive and command-line interfaces
- **Extracting exact numbering definitions from DOCX XML structure**
- **Resolving actual sequential numbers based on document position**
- **Ensuring proper DOM serialization to preserve document content**
- **Implementing comprehensive error handling and validation**
- **Providing diagnostic tools for troubleshooting conversion issues**
- **Extracting and calculating document statistics for metadata**
- **Generating section IDs for direct navigation to numbered headings and paragraphs**

The scope explicitly excludes:

- Editing capabilities for the input or output documents
- Conversion of other document formats (PDF, RTF, etc.)
- Web hosting or content management features
- Real-time conversion via web services (unless specified in future versions)

## 2. Target Users

### 2.1 Primary User Personas

1. **Content Managers**
   - Need to convert existing Word-based content for web publishing
   - Manage large document libraries that need web-friendly versions
   - Value preservation of document structure and styles

2. **Technical Writers**
   - Create documentation in Word that needs to be published online
   - Work with multilingual content and complex formatting
   - Require batch processing capabilities for large documentation sets

3. **Web Developers**
   - Import client-provided Word documents into web platforms
   - Need clean, well-structured HTML or Markdown output
   - Value the ability to extract and apply document styles

4. **Knowledge Base Administrators**
   - Maintain organizational documentation in Word format
   - Need to convert documents for internal content management systems
   - Require preservation of document structure and formatting

### 2.2 Use Cases

1. **Content Migration**
   - Convert existing document libraries for web publishing
   - Preserve document organization through directory structure mirroring
   - Process thousands of files in batches

2. **Documentation Publishing**
   - Transform technical documentation from Word to Markdown/HTML
   - Maintain accurate rendering of tables, lists, and code blocks
   - Support international character sets and mixed language content

3. **Archival and Accessibility**
   - Convert proprietary DOCX format to more accessible web formats
   - Create searchable, lightweight representations of document content
   - Preserve document styling for historical accuracy

4. **Technical Specification Conversion**
   - Convert complex technical documents with specialized formatting
   - Preserve hierarchical numbered sections and references
   - Maintain table formats and specialized notation

## 3. Features and Requirements

### 3.1 Core Features

#### 3.1.1 Document Conversion

- Convert DOCX files to Markdown format
- Generate HTML with preserved styling
- Extract and save document images
- Process documents with mixed language content
- Support right-to-left text and international characters
- Automatically detect and properly decorate table of contents and index elements
- Maintain hierarchical list structures with proper nesting and numbering
- Analyze document structure to identify special sections and formatting patterns
- **Extract exact numbering formats from DOCX XML structure**
- **Resolve actual sequential numbers based on document position**
- **Ensure proper DOM serialization to preserve document content**
- **Implement comprehensive error handling and validation**
- **Provide detailed logging for troubleshooting conversion issues**
- **Extract and calculate document statistics for metadata**
- **Generate section IDs for direct navigation to numbered headings and paragraphs**

#### 3.1.2 Style Preservation

- Extract and apply document styles to HTML output
- Generate CSS from original document styling
- Preserve text formatting (bold, italic, underline, etc.)
- Maintain document structure (headings, lists, tables)
- Preserve document layout and spacing
- Ensure proper nesting of multi-level lists with correct numbering
- Accurately render TOC styles including leader lines and page numbers
- Extract and apply numbering definitions for complex hierarchical lists
- **Parse abstract numbering definitions with all level properties**
- **Extract level text formats (e.g., "%1.", "%1.%2.", "(%1)")**
- **Capture indentation, alignment, and formatting for each level**
- **Handle level overrides and start value modifications**
- **Parse run properties (font, size, color) for numbering text**
- **Ensure content preservation during DOM manipulation and serialization**
- **Implement fallback mechanisms for DOM manipulation errors**
- **Generate section IDs based on hierarchical numbering structure**

#### 3.1.3 Processing Capabilities

- Process individual DOCX files
- Process entire directories (including subdirectories)
- Process lists of files from text input
- Search for files using criteria (find command integration)
- Maintain original directory structure in output
- Support for batch processing with progress reporting
- Resume capability for interrupted batch operations
- **Provide diagnostic tools for troubleshooting conversion issues**
- **Generate detailed logs for each processing step**
- **Implement validation at key processing stages**

#### 3.1.4 User Interface

- Command-line interface for direct usage
- Interactive menu-driven interface for guided operation
- Helper scripts for common batch operations
- Clear, actionable error messages
- Progress indicators for longer operations
- **Detailed diagnostic output for troubleshooting**
- **Performance metrics and summary statistics**

### 3.2 Technical Requirements

#### 3.2.1 Platform Support

- Node.js-based application (version 14.x or higher)
- Cross-platform compatibility (Windows, macOS, Linux)
- Minimal external dependencies
- Support for scripting and automation

#### 3.2.2 Performance

- Process standard documents (< 10MB) in under 5 seconds
- Support batch processing of at least 100 files
- Handle documents up to 100MB in size
- Memory usage under 1GB for standard operations
- CPU utilization optimization for multi-core systems
- **Performance timing and metrics for optimization**

#### 3.2.3 Output Quality

- Generated HTML should closely match original document appearance
- Markdown output should maintain document structure
- Markdown output should comply with standard linting rules
- HTML output should follow web best practices with external CSS files
- Images should be properly extracted and referenced
- Unicode and international character support
- Hierarchical lists should maintain proper nesting and numbering
- TOC elements should be properly styled with appropriate leader lines and formatting
- HTML should be valid according to W3C standards
- **Numbering should exactly match the original DOCX document**
- **CSS should accurately reflect DOCX numbering formats**
- **DOM serialization must preserve all document content and structure**
- **Output files should be validated before saving**
- **HTML files should be > 1000 characters for typical documents**
- **CSS files should contain generated styles appropriate for the document**
- **Document statistics (pages, words, characters, paragraphs, lines) must be accurately calculated and included in metadata**
- **Document statistics must be calculated from document content when not available in metadata**
- **HTML output must be properly formatted with indentation and line breaks for improved debugging**
- **HTML formatting must preserve all content and structure while enhancing readability**

#### 3.2.4 Document Analysis and Style Extraction

- **MUST NOT pattern match against specific content words** (e.g., "Rationale", "Whereas", etc.)
- **MUST use DOCX introspection** to determine appropriate styling and structure
- Extract styling information from DOCX XML structure, not from text content
- Determine formatting based on document's style definitions, not content patterns
- Generate styles based on document's structure, not assumptions about text meaning
- All CSS must be generated dynamically from the document's style information
- Do not hardcode formatting based on specific text content patterns
- All structure relationships (headings, lists, TOC) must be determined by analyzing document structure
- **Extract complete numbering definitions from numbering.xml**
- **Resolve actual sequential numbers based on document position**
- **Map paragraphs to their numbering definitions**
- **Handle restart logic and level overrides**
- **Ensure proper DOM serialization to preserve document content**
- **Validate document structure at each processing step**

## 4. Technical Specifications

### 4.1 Architecture

doc2web follows a modular architecture with these primary components:

1. **Main Converter (doc2web.js)**
   - Handles input processing and orchestrates conversion
   - Manages file system operations and output organization
   - Coordinates the overall conversion process
   - Implements error handling and recovery
   - **Provides detailed logging and diagnostics**
   - **Implements input file validation**
   - **Includes performance timing and metrics**

2. **Markdown Generator (markdownify.js)**
   - Converts HTML to well-structured Markdown
   - Preserves document organization and formatting
   - Ensures generated markdown is compliant with linting standards
   - Fixes common markdown formatting issues automatically

3. **Library Modules (lib/)**
   - Organized into logical function groups:
     - **XML Utilities (lib/xml/)**
       - `xpath-utils.js`: Provides utilities for working with XML and XPath
       - Handles namespace resolution for DOCX XML structure
       - Implements error handling for XPath queries

     - **Parsers (lib/parsers/)**
       - `style-parser.js`: Parses DOCX styles and extracts formatting information
       - `theme-parser.js`: Extracts theme information (colors, fonts)
       - `toc-parser.js`: Parses Table of Contents styles and structure
       - `numbering-parser.js`: Extracts numbering definitions for lists
       - `numbering-resolver.js`: Resolves actual numbers based on document position
       - `document-parser.js`: Analyzes document structure and settings
       - `metadata-parser.js`: Extracts and processes document metadata, calculates document statistics

     - **HTML Processing (lib/html/)**
       - `html-generator.js`: Main module for generating HTML from DOCX
       - `structure-processor.js`: Ensures proper HTML document structure
       - `content-processors.js`: Processes headings, TOC, and lists
       - `element-processors.js`: Processes tables, images, and language elements

     - **CSS Generation (lib/css/)**
       - `css-generator.js`: Generates CSS from extracted style information
       - `style-mapper.js`: Maps DOCX styles to CSS classes

     - **Utilities (lib/utils/)**
       - `unit-converter.js`: Converts between different units (twips, points)
       - `common-utils.js`: Common utility functions shared across modules

4. **Style Processing**
   - **Style Parser (lib/parsers/style-parser.js)**
     - Parses DOCX XML structure
     - Extracts detailed style information
     - **Must determine document structure based on XML, not content pattern matching**
     - **Integrates numbering context into style extraction**
     - **Detects embedded numbering in paragraph styles**
     - **Identifies heading styles automatically**

   - **Numbering Parser (lib/parsers/numbering-parser.js)**
     - **Parses abstract numbering definitions with all level properties**
     - **Extracts level text formats (e.g., "%1.", "%1.%2.", "(%1)")**
     - **Captures indentation, alignment, and formatting for each level**
     - **Handles level overrides and start value modifications**
     - **Parses run properties (font, size, color) for numbering text**
     - **Converts DOCX patterns to CSS counter content**

   - **Numbering Resolver (lib/parsers/numbering-resolver.js)**
     - **Extracts paragraph numbering context from document.xml**
     - **Maps paragraphs to their numbering definitions**
     - **Resolves actual sequential numbers based on document position**
     - **Handles restart logic and level overrides**
     - **Tracks numbering sequences across the document**
     - **Generates section IDs based on hierarchical numbering structure**

   - **Metadata Parser (lib/parsers/metadata-parser.js)**
     - Extracts document metadata from core.xml and app.xml
     - Processes document statistics (pages, words, characters, paragraphs, lines)
     - Calculates statistics from document content when not available in metadata
     - Applies metadata to HTML output as meta tags and comments
     - Generates structured data (JSON-LD) with document information
     - **Ensures accurate document statistics in output**

   - **Style Extractor (lib/html/html-generator.js)**
     - Extracts styling information from DOCX
     - Applies styles to HTML output
     - Processes document structure including hierarchical lists
     - **Must analyze document structure without assumptions about content**
     - **Passes numbering context through the conversion pipeline**
     - **Creates enhanced style mappings with numbering attributes**
     - **Applies DOCX-derived numbering during HTML processing**
     - **Ensures proper DOM serialization to preserve document content**
     - **Implements fallback mechanisms for DOM serialization issues**
     - **Verifies body content before serialization**
     - **Formats HTML output with proper indentation and line breaks**
     - **Preserves HTML structure while enhancing readability for debugging**
     - **Implements comprehensive error handling throughout the pipeline**
     - **Validates content at each processing step**

   - **CSS Generator (lib/css/css-generator.js)**
     - Generates clean, readable CSS with proper margins and spacing
     - Creates styles based on document's structure, not content
     - **Generates CSS counters from DOCX numbering definitions**
     - **Creates level-specific styling for numbered elements**
     - **Supports hierarchical numbering with proper resets**
     - **Maintains visual fidelity to original DOCX formatting**
     - **Adds section ID styling for navigation and highlighting**

   - **Content Processors (lib/html/content-processors.js)**
     - Processes headings, TOC, and lists
     - Detects and styles TOC and index elements
     - Handles special document sections with appropriate formatting
     - **Matches HTML elements to DOCX paragraph contexts**
     - **Applies exact numbering from DOCX to headings**
     - **Creates structured lists based on DOCX numbering definitions**
     - **Maintains hierarchical relationships from original document**
     - **Ensures proper DOM manipulation to preserve content**
     - **Applies section IDs to headings and numbered paragraphs**

5. **User Interface (doc2web-run.js)**
   - Provides interactive command-line interface
   - Guides users through conversion options
   - Displays progress and results
   - Handles user input validation

6. **Utility Scripts**
   - Installation and initialization (doc2web-install.sh, init-doc2web.sh)
   - Batch processing helpers (process-find.sh)
   - System compatibility checks
   - Error logging and reporting
   - **Diagnostic tools (debug-test.js)**

7. **Documentation**
   - `README.md`: Project overview and quick start guide
   - `docs/prd.md`: Product Requirements Document
   - `docs/refactoring.md`: Detailed documentation of the refactoring process
   - `docs/user-guide.md`: User guide for working with the refactored code
   - `docs/troubleshooting_guide.md`: Guide for diagnosing and fixing issues

### 4.2 Dependencies

- Node.js ecosystem
- mammoth.js (DOCX parsing and basic conversion)
- jsdom (HTML manipulation)
- jszip (DOCX unpacking)
- xmldom and xpath (XML processing)

### 4.3 Output Specifications

- Output structure mirrors input directory structure
- Consistent naming convention for output files
- HTML output includes external CSS files with proper styling
- HTML documents have a 20px margin for improved readability
- Image files properly extracted and referenced
- Hierarchical lists properly nested with correct numbering
- TOC elements styled with appropriate leader lines and formatting
- Special document sections identified and styled appropriately based on style attributes, not content
- CSS and HTML must be ephemeral and regenerated for each document based on its unique structure
- **Numbering in HTML output exactly matches the original DOCX document**
- **CSS counters accurately reflect DOCX numbering definitions**
- **Hierarchical relationships maintained from original document**
- **DOM serialization must preserve all document content and structure**
- **Fallback mechanisms must be in place to handle DOM serialization issues**
- **HTML files should be > 1000 characters for typical documents**
- **CSS files should contain generated styles appropriate for the document**
- **Document statistics (pages, words, characters, paragraphs, lines) must be accurately calculated and included in metadata**
- **Document statistics must be calculated from document content when not available in metadata**
- **HTML output must be properly formatted with indentation and line breaks for improved debugging**
- **HTML formatting must preserve all content and structure while enhancing readability**
- **TOC elements must have proper leader dots and right-aligned page numbers**
- **TOC entries must maintain proper alignment and spacing using flex-based layout**
- **Paragraph numbering must be implemented using CSS ::before pseudo-elements for exact positioning**
- **Section IDs must be generated for all numbered headings and paragraphs**
- **Section IDs must follow a consistent pattern (e.g., "section-1-2-a")**
- **CSS must include styling for section navigation and highlighting**

### 4.4 Document Analysis Rules

- All styling decisions must be based on document structure, not text content
- Paragraphs must be categorized by their style attributes (indentation, font, etc.)
- Hierarchical relationships must be determined from DOCX style and numbering definitions
- When text patterns are used, they must be generic structural patterns (e.g., numbered format "1." or lettered format "a.") rather than specific word matches
- Special formatting must be determined by style analysis, not by searching for specific words
- The application must work for documents in any language or domain without assuming particular content patterns
- **Numbering must be extracted directly from numbering.xml, not inferred from text patterns**
- **Actual numbers must be resolved based on document position and numbering definitions**
- **Numbering formats must be converted to appropriate CSS counter styles**
- **DOM serialization must be verified to ensure content preservation**
- **Document structure must be validated at each processing step**
- **Section IDs must be derived from numbering structure, not from text content**

### 4.5 API and Integration

- File system API for reading inputs and writing outputs
- Logging API for error reporting and diagnostics
- Optional module exports for integration with other Node.js applications
- Command-line interface for script integration

### 4.6 DOM Serialization Requirements

- **Verify document body content before serialization**
- **Implement fallback mechanisms for empty body issues**
- **Log serialization metrics for debugging purposes**
- **Preserve document structure during DOM manipulation**
- **Ensure all content is properly serialized in the final HTML output**
- **Handle browser-specific DOM serialization differences**
- **Maintain proper nesting and hierarchical relationships**
- **Preserve attributes and data attributes during serialization**
- **Implement error handling for serialization failures**
- **Avoid unnecessary DOM manipulation that could lose content**
- **Validate serialized output before saving to file**
- **Implement content preservation strategies during processing**

### 4.7 Error Handling and Validation

- **Implement comprehensive error handling throughout the pipeline**
- **Provide detailed, actionable error messages**
- **Log errors with context information for debugging**
- **Validate input files before processing**
- **Validate output files before saving**
- **Implement fallback mechanisms for common error conditions**
- **Provide diagnostic tools for troubleshooting issues**
- **Include performance metrics and summary statistics**
- **Verify content preservation at key processing steps**
- **Implement safe DOM manipulation practices**
- **Add validation for accessibility features**

### 4.8 Recent Implementation Changes

#### 4.8.1 CSS-Based Heading Numbering

The implementation of heading numbering has been improved to rely on CSS ::before pseudo-elements rather than directly inserting numbers into the HTML. This change affects two key files:

1. **lib/html/content-processors.js (processHeadings function)**:
   - The function `applyNumberingToHeading` was effectively removed by commenting out its primary logic of inserting a `<span>`.
   - The `processHeadings` function now directly sets `data-numbering-id`, `data-numbering-level`, and `data-format` attributes on heading elements if `context.resolvedNumbering` (numbering information derived from DOCX) exists.
   - This ensures that headings intended to be numbered via DOCX definitions will rely solely on the CSS ::before pseudo-element mechanism, driven by these data attributes, rather than having numbers inserted directly into the HTML. This prevents potential conflicts or duplicate numbering.
   - `ensureHeadingAccessibility` is still called to manage IDs and focusability.

2. **lib/css/css-generator.js (generateDOCXNumberingStyles function)**:
   - The core logic for calculating `marginLeft`, `paddingLeft` (for the number area), and `textIndent` (for the text flow, especially handling hanging indents) based on `levelDef.indentation.left`, `levelDef.indentation.hanging`, and `levelDef.indentation.firstLine` is maintained and emphasized. These DOCX properties are crucial for accurate layout.
   - CSS for `[data-num-id][data-num-level]` (the numbered element itself):
     - `display: block;` is added to ensure consistent block-level behavior, which is important for p, h1-h6, and li that might receive these attributes.
     - `position: relative;` is crucial for the absolute positioning of the ::before pseudo-element.
     - `margin-left`, `padding-left`, and `text-indent` are applied based on the parsed DOCX values. Default to 0 if not specified to avoid unexpected browser defaults interfering.
   - CSS for `::before` (the number/bullet):
     - `content: ${getCSSCounterContent(levelDef, abstractNum.id, numberingDefs)};` generates the actual number string using CSS counters.
     - `position: absolute;`
     - `left: ${textIndent < 0 ? textIndent : 0}pt;`: This is a key adjustment. If there's a hanging indent (textIndent is negative), the number should start at that negative offset from the padding-left edge of the parent. If textIndent is zero or positive (standard first-line indent), the number starts at 0 (the padding edge).
     - `width: ${paddingLeft > 0 ? paddingLeft : (levelDef.indentation?.hanging || 24)}pt;`: The width of the number container is explicitly set, typically to the hanging indent amount. A fallback width (e.g., 24pt) is provided if paddingLeft isn't available from a hanging indent (though it should be if textIndent is negative).
     - `box-sizing: border-box;` is added to ensure the width includes any padding or border of the pseudo-element itself, which is good practice.
   - The `generateEnhancedListStyles` function is simplified to primarily style the list containers (ol, ul) and basic li properties, assuming the numbering and precise indentation of li items (if they also get data-num-id attributes) will be handled by `generateDOCXNumberingStyles`.
   - General improvements to default styles in `generateUtilityStyles` (e.g., centering images, better heading margins, basic print styles) and `generateTOCStyles` for robustness.
   - The `getCSSCounterFormat` was ensured to be imported for use in `getCSSCounterContent`.

#### 4.8.2 TOC and Body Content Numbering Fixes

Several issues with the Table of Contents layout and body content numbering have been addressed:

1. **TOC Layout Fixes**:
   - Fixed the "ragged" or two-column appearance of the TOC by ensuring proper CSS styling in the `generateTOCStyles` function
   - Added explicit `column-count: 1` to the `.docx-toc` selector to prevent unintended column layouts
   - Improved the display and flex properties of TOC entries to ensure proper vertical stacking
   - Enhanced TOC entry styling to maintain proper alignment of text, dots, and page numbers

2. **Body Content Numbering Fixes**:
   - Resolved issues with missing or "zero" numbering in the body content
   - Fixed the hierarchical numbering string construction in the `buildHierarchicalNumbering` function
   - Ensured proper counter initialization and incrementation in the CSS
   - Improved the data attribute application in `processHeadings` and `processTOCEntry` functions
   - Enhanced the CSS counter content generation in `getCSSCounterContent` to properly handle all numbering formats

3. **CSS Counter Implementation**:
   - Refined the counter reset strategy to ensure proper numbering hierarchy
   - Improved the positioning and styling of numbering elements using ::before pseudo-elements
   - Enhanced the handling of indentation and spacing for numbered elements
   - Fixed the interaction between level counters and formatting segments

These changes ensure that both the TOC and body content numbering now display correctly, maintaining the visual fidelity of the original DOCX document.

#### 4.8.3 Enhanced TOC and Paragraph Numbering (v1.2.2)

#### 4.8.4 Comprehensive TOC Implementation Fixes (v1.2.3)

The Table of Contents implementation has been further enhanced with comprehensive fixes to ensure consistent rendering across all browsers and document types:

1. **CSS Specificity and Importance Improvements**:
   - Added `!important` declarations to critical flex properties to ensure they override any conflicting styles
   - Enhanced specificity of TOC-related selectors to prevent style conflicts
   - Fixed CSS inheritance issues that could affect TOC layout
   - Ensured consistent styling across different contexts and parent containers

2. **DOM Structure Validation**:
   - Implemented a comprehensive validation function that ensures all TOC entries have the complete three-part structure
   - Added automatic restructuring for entries that are missing any components (text, dots, or page numbers)
   - Ensured proper ARIA attributes for accessibility
   - Added validation for entries both with and without page numbers

3. **Leader Dots Implementation Enhancements**:
   - Refined the background-image pattern for dots with specific values for consistent visibility
   - Positioned dots closer to text baseline for better alignment
   - Implemented consistent sizing (1px height, 6px spacing) for dots
   - Added browser compatibility fallbacks for gradient approaches

4. **Layout and Container Improvements**:
   - Enhanced single-column layout enforcement with multiple CSS properties
   - Added box-sizing and max-width properties to prevent overflow issues
   - Implemented page-break-inside: avoid to keep TOC entries together when printing
   - Fixed paragraph display with specific overrides for TOC entries that are paragraphs

5. **Browser Compatibility Enhancements**:
   - Added fallback styles for browsers that don't support gradient approaches
   - Ensured consistent rendering across different browser engines
   - Implemented defensive CSS to handle edge cases in older browsers
   - Added print-specific styles for better TOC appearance in printed documents

These comprehensive fixes ensure that the Table of Contents displays correctly with proper alignment, consistent dots, and right-aligned page numbers across all browsers and document types, regardless of the complexity of the original document structure.

#### 4.8.5 Section IDs for Navigation (v1.2.4)

The implementation of section IDs for direct navigation to numbered headings and paragraphs provides significant improvements to document navigation and accessibility:

1. **Section ID Generation**:
   - Automatically generates section IDs based on hierarchical numbering structure
   - Creates IDs that match the exact numbering pattern (e.g., "section-1-2-a")
   - Handles all numbering formats (decimal, alphabetic, roman numerals)
   - Ensures IDs are valid HTML identifiers by removing invalid characters
   - Maintains consistent ID pattern regardless of document language or content domain

2. **Implementation Components**:
   - **Numbering Resolver (lib/parsers/numbering-resolver.js)**:
     - Added `generateSectionId` method to `NumberingSequenceTracker` class
     - Enhanced `formatNumbering` method to include section ID in returned object
     - Converts full numbering string to valid HTML ID format
     - Handles special characters and formatting in numbering strings
     - Ensures consistent ID generation across all document types

   - **Content Processors (lib/html/content-processors.js)**:
     - Updated `processHeadings` function to apply section IDs from resolved numbering
     - Enhanced `ensureHeadingAccessibility` to preserve existing section IDs
     - Extended `processNestedNumberedParagraphs` to apply section IDs to non-heading elements
     - Maintains proper ID hierarchy matching document structure
     - Preserves accessibility attributes while adding navigation capabilities

   - **CSS Generator (lib/css/css-generator.js)**:
     - Added section ID styling in `generateDOCXNumberingStyles` function
     - Implemented scroll-margin-top for smooth scrolling to sections
     - Added target highlighting in `generateUtilityStyles` function
     - Created visual feedback for navigation to specific sections
     - Ensured consistent styling across different browsers

3. **Navigation Benefits**:
   - Enables direct linking to specific sections via fragment identifiers (e.g., `#section-1-2-a`)
   - Improves accessibility for screen readers through proper document structure
   - Enhances TOC functionality with precise section targeting
   - Provides visual feedback when navigating to sections
   - Supports smooth scrolling to targeted sections

4. **Technical Implementation Details**:
   - Section IDs are derived directly from DOCX numbering definitions
   - IDs match the exact hierarchical structure of the document
   - Implementation is content-agnostic, working with any language or domain
   - Diagnostic tools validate section ID generation
   - Documentation provides clear usage examples

This enhancement significantly improves document navigation and accessibility while maintaining the existing functionality and keeping the TOC implementation intact. The section IDs provide a direct mapping between the document's hierarchical structure and the HTML navigation system, enabling precise targeting of specific sections regardless of document complexity.

#### 4.8.6 Character Overlap and Numbering Display Fixes (v1.2.5)

#### 4.8.7 Paragraph Numbering and TOC Display Fix (v1.2.6)

The implementation of paragraph numbering and TOC display fixes addresses issues with missing paragraph numbers and subheader letters in the TOC and throughout the document:

1. **Heading Numbering Display Fixes**:
   - Fixed the issue where heading numbers were not displaying in the document
   - Added code to populate the `<span class="heading-number">` element with the actual numbering content
   - Ensured that heading numbers display correctly for all heading levels
   - Maintained proper spacing between heading numbers and content
   - Preserved accessibility attributes for screen readers

2. **TOC Numbering Display Fixes**:
   - Fixed the issue where TOC entries were missing their paragraph numbers and subheader letters
   - Added code to prepend the numbering directly to the text content of TOC entries
   - Improved the anchor creation logic to handle entries with section IDs from numbering
   - Enhanced the TOC entry structure to maintain proper alignment with numbering
   - Ensured consistent display of numbering across all TOC levels

3. **Implementation Details**:
   - **Heading Number Population**:
     - Modified `processHeadings` function to add the actual numbering content to the heading-number span
     - Used `context.resolvedNumbering.fullNumbering` to get the complete numbering string
     - Maintained the existing HTML structure while adding the missing content
   
   - **TOC Entry Enhancement**:
     - Updated `processTOCEntry` function to add numbering to TOC entry text
     - Improved anchor creation to use section IDs for better navigation
     - Enhanced the TOC entry structure to handle numbering properly
     - Maintained compatibility with existing TOC styling

4. **Diagnostic Tools**:
   - Added a debug test script to verify numbering display
   - Implemented validation for heading and TOC numbering
   - Added logging for numbering-related operations
   - Provided tools to test the fix without rebuilding the entire document

These fixes ensure that paragraph numbers and subheader letters display correctly in both the TOC and throughout the document, improving the visual fidelity and usability of the generated HTML output.

The implementation of character overlap and numbering display fixes addresses issues with paragraph numbers or letters from IDs not displaying correctly and characters overlapping:

1. **CSS Spacing and Positioning Improvements**:
   - Increased padding and width for numbering elements to prevent overlap with text
   - Added proper box-sizing and overflow handling to prevent text from overflowing
   - Added specific CSS rules for Roman numerals to ensure proper spacing
   - Improved whitespace handling with `white-space: nowrap` for numbering elements
   - Enhanced spacing between numbering and paragraph content

2. **Section ID Display Enhancements**:
   - Improved handling of Roman numerals in section IDs
   - Added special CSS rules for Roman numeral sections
   - Enhanced spacing for section headings with numbering
   - Fixed character overlap in section IDs with proper padding

3. **Box Model Fixes**:
   - Added `overflow-wrap: break-word` and `word-wrap: break-word` to prevent text overflow
   - Added extra padding for elements with Roman numerals
   - Fixed line wrapping issues with indented paragraphs
   - Implemented consistent box model across all numbered elements

4. **Numbering Display Improvements**:
   - Enhanced the `buildDisplayNumbering` method to add proper spacing between numbering segments
   - Added special handling for Roman numerals to ensure they display correctly
   - Fixed spacing issues between numbering and content
   - Improved handling of complex numbering formats

5. **HTML Structure Improvements**:
   - Added wrapper elements for heading numbers to help with styling
   - Added new CSS classes for Roman numeral sections and headings
   - Applied inline styles for critical spacing properties
   - Enhanced the structure of numbered elements for better display

These fixes ensure that paragraph numbers and section IDs display correctly without character overlap, improving the readability and visual fidelity of the generated HTML output. The implementation is content-agnostic and works with all numbering formats, including Roman numerals, letters, and decimal numbers.

The implementation of TOC formatting and paragraph numbering has been significantly improved to enhance visual fidelity and user experience:

1. **TOC Formatting Enhancements**:
   - Implemented a flex-based layout for TOC entries that ensures proper alignment of text, dots, and page numbers
   - Created leader dots using CSS background-image with radial gradients for precise control over appearance
   - Applied right-alignment to page numbers for professional document appearance
   - Fixed the "ragged" or multi-column appearance of the TOC by explicitly setting column-count: 1
   - Enhanced the display properties to ensure proper vertical stacking of TOC entries

2. **Paragraph Numbering Improvements**:
   - Implemented CSS ::before pseudo-elements for displaying paragraph numbers
   - Used data attributes to track numbering context without modifying document content
   - Applied absolute positioning for precise placement of numbering elements
   - Enhanced the CSS counter implementation to properly handle all numbering formats
   - Fixed issues with missing or "zero" numbering in the body content
   - Improved the hierarchical numbering string construction

3. **Technical Implementation Details**:
   - **Flex-based TOC Layout**:
     - Each TOC entry is a flex container with `display: flex`
     - Text content has `flex-grow: 0` to maintain its natural size
     - Dots section has `flex-grow: 1` to fill available space
     - Page numbers have `flex-grow: 0` with `text-align: right`
     - Leader dots are created using CSS `background-image` with a radial gradient pattern

   - **CSS ::before Implementation for Numbering**:
     - Numbered elements receive `data-numbering-id`, `data-numbering-level`, and `data-format` attributes
     - Parent elements have `position: relative` for proper positioning context
     - The `::before` pseudo-element uses `position: absolute` with precise left positioning
     - Width is explicitly set to ensure proper text flow
     - Content is generated using CSS counters that match DOCX numbering definitions

   - **Counter Reset Strategy**:
     - Counter reset is applied at appropriate levels to maintain hierarchical structure
     - Each level properly increments its own counter
     - Hierarchical numbering strings are constructed using the appropriate format
     - Level-specific styling ensures visual consistency with the original document

These enhancements ensure that the generated HTML closely resembles the original DOCX document's appearance, providing end users with a more accurate and professional representation of their content. The implementation is completely generic and content-agnostic, working with any document regardless of language or domain.
