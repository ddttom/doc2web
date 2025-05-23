# doc2web Product Requirements Document

**Document Version:** 3.4  
**Last Updated:** May 23, 2025  
**Status:** Draft  
**Authors:** Technical Team  
**Approved By:** [Pending]  

## Document Revision History

| Version | Date | Author | Description of Changes |
|---------|------|--------|------------------------|
| 1.0 | May 16, 2025 | Initial Team | Original document |
| 2.0 | May 20, 2025 | Technical Team | Added document introspection rules, removed content pattern matching |
| 3.0 | May 23, 2025 | Technical Team | Added section IDs for direct navigation to numbered headings |
| 3.4 | May 23, 2025 | Technical Team | Added document header extraction and processing functionality |

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
    - [4.8 Implementation Status](#48-implementation-status)

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
   - **Diagnostic tools (debug_test_script.js)**

7. **Documentation**
   - `README.md`: Project overview and quick start guide
   - `docs/prd.md`: Product Requirements Document
   - `docs/architecture.md`: Technical architecture documentation

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

### 4.8 Implementation Status

The application has undergone significant enhancements to improve document conversion accuracy and user experience. Key improvements include:

- **CSS-Based Numbering**: Implemented precise numbering using CSS ::before pseudo-elements
- **TOC Enhancement**: Fixed layout issues and improved visual fidelity
- **Section Navigation**: Added section IDs for direct navigation to numbered headings
- **DOM Serialization**: Enhanced content preservation during HTML processing
- **Header Processing**: Added comprehensive document header extraction and processing
- **Diagnostic Tools**: Enhanced debugging capabilities for troubleshooting

For detailed technical implementation information, see [`docs/architecture.md`](docs/architecture.md).
