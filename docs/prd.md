# doc2web Product Requirements Document

**Document Version:** 4.6  
**Last Updated:** June 3, 2025  
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
| 3.5 | May 26, 2025 | Technical Team | Updated architecture to reflect modular refactoring |
| 3.6 | May 31, 2025 | Technical Team | Added hanging margins implementation for TOC and numbered content |
| 3.7 | May 31, 2025 | Technical Team | Added header image extraction and positioning functionality |
| 3.8 | June 1, 2025 | Technical Team | Enhanced HTML formatting with proper indentation and line breaks |
| 3.9 | June 1, 2025 | Technical Team | Improved table formatting with enhanced styling and semantic structure |
| 4.0 | June 2, 2025 | Technical Team | Fixed bullet point display and indentation issues with enhanced CSS specificity |
| 4.1 | June 2, 2025 | Technical Team | Fixed italic formatting conversion from DOCX to HTML with enhanced mammoth.js configuration |
| 4.2 | June 2, 2025 | Technical Team | Fixed hanging indentation for numbered headings and paragraphs with enhanced CSS specificity and margin reset |
| 4.3 | June 2, 2025 | Technical Team | Added TOC page number removal functionality for web-appropriate navigation |
| 4.4 | June 2, 2025 | Technical Team | Optimized HTML formatting to remove whitespace between elements while preserving newlines for compact output |
| 4.5 | June 3, 2025 | Technical Team | Added TOC linking functionality with hierarchical navigation from TOC entries to document sections |
| 4.6 | June 3, 2025 | Technical Team | Added return to top button functionality with theme-aware styling and accessibility support |

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
      - [Document Conversion](#document-conversion)
      - [Style Preservation](#style-preservation)
      - [Processing Capabilities](#processing-capabilities)
      - [User Interface](#user-interface)
    - [3.2 Technical Requirements](#32-technical-requirements)
      - [Platform Support](#platform-support)
      - [Performance](#performance)
      - [Output Quality](#output-quality)
      - [Document Analysis and Style Extraction](#document-analysis-and-style-extraction)
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
- Ensure WCAG 2.1 Level AA accessibility compliance
- Preserve document metadata with accurate statistics
- Handle documents with tracked changes appropriately
- Extract exact numbering and formatting from DOCX XML structure
- Ensure proper preservation of document content during DOM manipulation
- Provide comprehensive error reporting and recovery mechanisms
- Implement validation at key processing steps to ensure data integrity
- Generate section IDs for direct navigation to numbered headings and paragraphs
- Maintain clean, focused modules for improved maintainability and extensibility
- Implement proper hanging indent behavior for TOC entries and numbered content
- Extract and display header images with proper positioning
- Generate optimized HTML output with proper formatting
- Enhance table presentation with professional styling and accessibility
- Ensure proper text formatting conversion including italic, bold, and mixed styles
- Remove page numbers from Table of Contents entries for web-appropriate navigation
- Implement clickable TOC navigation with hierarchical section mapping and accessibility compliance
- Provide return to top button with theme-aware styling, smooth scrolling, and full accessibility support
- Provide return to top button functionality for improved user navigation and accessibility

### 1.3 Project Scope

doc2web is focused on the conversion of DOCX documents to HTML and Markdown formats. The scope includes:

- Processing individual files, directories of files, and file lists
- Extracting and preserving styles, images, and document structure
- Maintaining fidelity to the original document layout
- Providing both interactive and command-line interfaces
- Extracting exact numbering definitions from DOCX XML structure
- Resolving actual sequential numbers based on document position
- Ensuring proper DOM serialization to preserve document content
- Implementing comprehensive error handling and validation
- Providing diagnostic tools for troubleshooting conversion issues
- Extracting and calculating document statistics for metadata
- Generating section IDs for direct navigation to numbered headings and paragraphs
- Maintaining modular, focused codebase with clear separation of concerns
- Implementing proper hanging indent behavior for TOC entries and numbered content
- Extracting and positioning header images with fidelity to original DOCX layout
- Generating properly formatted HTML with optimized output
- Enhancing table presentation with professional styling and accessibility features
- Ensuring proper text formatting conversion with comprehensive style mapping
- Implementing clickable TOC navigation with hierarchical section mapping and accessibility compliance

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

#### Document Conversion

- Convert DOCX files to Markdown format
- Generate HTML with preserved styling
- Extract and save document images
- Process documents with mixed language content
- Support right-to-left text and international characters
- Automatically detect and properly decorate table of contents and index elements
- Generate clickable navigation links from TOC entries to corresponding document sections
- Maintain hierarchical list structures with proper nesting and numbering
- Analyze document structure to identify special sections and formatting patterns
- Extract exact numbering formats from DOCX XML structure
- Resolve actual sequential numbers based on document position
- Ensure proper DOM serialization to preserve document content
- Implement comprehensive error handling and validation
- Provide detailed logging for troubleshooting conversion issues
- Extract and calculate document statistics for metadata
- Generate section IDs for direct navigation to numbered headings and paragraphs
- Implement proper hanging indent behavior for TOC entries and numbered content
- Extract and position header images with fidelity to original DOCX positioning
- Generate optimized HTML output with proper formatting
- Enhance table presentation with professional styling and accessibility features
- Ensure proper text formatting conversion with comprehensive style mapping
- Remove page numbers from Table of Contents entries for web-appropriate navigation
- Implement clickable TOC navigation with hierarchical section mapping and accessibility compliance
- Provide return to top button with theme-aware styling, smooth scrolling, and full accessibility support
- Provide return to top button functionality for improved user navigation and accessibility

#### Style Preservation

- Extract and apply document styles to HTML output
- Generate CSS from original document styling
- Preserve text formatting (bold, italic, underline, etc.)
- Maintain document structure (headings, lists, tables)
- Preserve document layout and spacing
- Ensure proper nesting of multi-level lists with correct numbering
- Accurately render TOC styles including leader lines and page numbers
- Extract and apply numbering definitions for complex hierarchical lists
- Parse abstract numbering definitions with all level properties
- Extract level text formats (e.g., "%1.", "%1.%2.", "(%1)")
- Capture indentation, alignment, and formatting for each level
- Handle level overrides and start value modifications
- Parse run properties (font, size, color) for numbering text
- Ensure content preservation during DOM manipulation and serialization
- Implement fallback mechanisms for DOM manipulation errors
- Generate section IDs based on hierarchical numbering structure
- Implement proper hanging indent behavior using CSS text-indent and padding-left properties
- Resolve CSS rule conflicts that cause margin compounding in numbered elements
- Use enhanced CSS specificity to override conflicting styles

#### Processing Capabilities

- Process individual DOCX files
- Process entire directories (including subdirectories)
- Process lists of files from text input
- Search for files using criteria (find command integration)
- Maintain original directory structure in output
- Support for batch processing with progress reporting
- Resume capability for interrupted batch operations
- Provide diagnostic tools for troubleshooting conversion issues
- Generate detailed logs for each processing step
- Implement validation at key processing stages

#### User Interface

- Command-line interface for direct usage
- Interactive menu-driven interface for guided operation
- Helper scripts for common batch operations
- Clear, actionable error messages
- Progress indicators for longer operations
- Detailed diagnostic output for troubleshooting
- Performance metrics and summary statistics

### 3.2 Technical Requirements

#### Platform Support

- Node.js-based application (version 14.x or higher)
- Cross-platform compatibility (Windows, macOS, Linux)
- Minimal external dependencies
- Support for scripting and automation

#### Performance

- Process standard documents (< 10MB) in under 5 seconds
- Support batch processing of at least 100 files
- Handle documents up to 100MB in size
- Memory usage under 1GB for standard operations
- CPU utilization optimization for multi-core systems
- Performance timing and metrics for optimization

#### Output Quality

- Generated HTML should closely match original document appearance
- Markdown output should maintain document structure
- Markdown output should comply with standard linting rules
- HTML output should follow web best practices with external CSS files
- Images should be properly extracted and referenced
- Unicode and international character support
- Hierarchical lists should maintain proper nesting and numbering
- TOC elements should be properly styled with appropriate leader lines and formatting
- HTML should be valid according to W3C standards
- Numbering should exactly match the original DOCX document
- CSS should accurately reflect DOCX numbering formats
- DOM serialization must preserve all document content and structure
- Output files should be validated before saving
- HTML files should be > 1000 characters for typical documents
- CSS files should contain generated styles appropriate for the document
- Document statistics must be accurately calculated and included in metadata
- HTML output must be optimized with proper formatting

#### Document Analysis and Style Extraction

- **MUST NOT pattern match against specific content words** (e.g., "Rationale", "Whereas", etc.)
- **MUST use DOCX introspection** to determine appropriate styling and structure
- Extract styling information from DOCX XML structure, not from text content
- Determine formatting based on document's style definitions, not content patterns
- Generate styles based on document's structure, not assumptions about text meaning
- All CSS must be generated dynamically from the document's style information
- Do not hardcode formatting based on specific text content patterns
- All structure relationships (headings, lists, TOC) must be determined by analyzing document structure
- Extract complete numbering definitions from numbering.xml
- Resolve actual sequential numbers based on document position
- Map paragraphs to their numbering definitions
- Handle restart logic and level overrides
- Ensure proper DOM serialization to preserve document content
- Validate document structure at each processing step

## 4. Technical Specifications

### 4.1 Architecture

doc2web follows a modular architecture with these primary components:

1. **Main Converter (doc2web.js)**
   - Handles input processing and orchestrates conversion
   - Manages file system operations and output organization
   - Coordinates the overall conversion process
   - Implements error handling and recovery
   - Provides detailed logging and diagnostics
   - Implements input file validation
   - Includes performance timing and metrics

2. **Markdown Generator (markdownify.js)**
   - Converts HTML to well-structured Markdown
   - Preserves document organization and formatting
   - Ensures generated markdown is compliant with linting standards
   - Fixes common markdown formatting issues automatically

3. **Library Modules (lib/)**
   - Organized into logical function groups:
     - **XML Utilities (lib/xml/)**: XPath utilities and XML processing
     - **Parsers (lib/parsers/)**: Style, theme, TOC, numbering, document, and metadata parsers
     - **HTML Processing (lib/html/)**: HTML generation, processing, and formatting
     - **CSS Generation (lib/css/)**: Style generation and mapping
     - **Utilities (lib/utils/)**: Common utility functions

4. **Style Processing**
   - **Style Parser**: Parses DOCX XML structure and extracts detailed style information
   - **Numbering Parser**: Extracts numbering definitions and level properties
   - **Numbering Resolver**: Resolves actual sequential numbers based on document position
   - **Metadata Parser**: Extracts and processes document metadata and statistics
   - **Style Extractor**: Applies styles to HTML output with proper DOM serialization
   - **CSS Generator**: Generates clean, readable CSS with proper formatting
   - **Content Processors**: Orchestrates content processing through specialized modules

5. **User Interface (doc2web-run.js)**
   - Provides interactive command-line interface
   - Guides users through conversion options
   - Displays progress and results
   - Handles user input validation

6. **Utility Scripts**
   - Installation and initialization scripts
   - Batch processing helpers
   - System compatibility checks
   - Error logging and reporting
   - Diagnostic tools

7. **Documentation**
   - Project overview and quick start guide
   - Product Requirements Document
   - Technical architecture documentation

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
- Numbering in HTML output exactly matches the original DOCX document
- CSS counters accurately reflect DOCX numbering definitions
- Hierarchical relationships maintained from original document
- DOM serialization must preserve all document content and structure
- Fallback mechanisms must be in place to handle DOM serialization issues
- HTML files should be > 1000 characters for typical documents
- CSS files should contain generated styles appropriate for the document
- Document statistics must be accurately calculated and included in metadata
- HTML output must be optimized with proper formatting
- TOC elements must have proper hanging indents with appropriate styling
- Section IDs must be generated for all numbered headings and paragraphs
- CSS must include styling for section navigation and highlighting

### 4.4 Document Analysis Rules

- All styling decisions must be based on document structure, not text content
- Paragraphs must be categorized by their style attributes (indentation, font, etc.)
- Hierarchical relationships must be determined from DOCX style and numbering definitions
- When text patterns are used, they must be generic structural patterns (e.g., numbered format "1." or lettered format "a.") rather than specific word matches
- Special formatting must be determined by style analysis, not by searching for specific words
- The application must work for documents in any language or domain without assuming particular content patterns
- Numbering must be extracted directly from numbering.xml, not inferred from text patterns
- Actual numbers must be resolved based on document position and numbering definitions
- Numbering formats must be converted to appropriate CSS counter styles
- DOM serialization must be verified to ensure content preservation
- Document structure must be validated at each processing step
- Section IDs must be derived from numbering structure, not from text content

### 4.5 API and Integration

- File system API for reading inputs and writing outputs
- Logging API for error reporting and diagnostics
- Optional module exports for integration with other Node.js applications
- Command-line interface for script integration

### 4.6 DOM Serialization Requirements

- Verify document body content before serialization
- Implement fallback mechanisms for empty body issues
- Log serialization metrics for debugging purposes
- Preserve document structure during DOM manipulation
- Ensure all content is properly serialized in the final HTML output
- Handle browser-specific DOM serialization differences
- Maintain proper nesting and hierarchical relationships
- Preserve attributes and data attributes during serialization
- Implement error handling for serialization failures
- Avoid unnecessary DOM manipulation that could lose content
- Validate serialized output before saving to file
- Implement content preservation strategies during processing

### 4.7 Error Handling and Validation

- Implement comprehensive error handling throughout the pipeline
- Provide detailed, actionable error messages
- Log errors with context information for debugging
- Validate input files before processing
- Validate output files before saving
- Implement fallback mechanisms for common error conditions
- Provide diagnostic tools for troubleshooting issues
- Include performance metrics and summary statistics
- Verify content preservation at key processing steps
- Implement safe DOM manipulation practices
- Add validation for accessibility features

### 4.8 Implementation Status

The application has undergone significant enhancements to improve document conversion accuracy and user experience. Key improvements include:

- **CSS-Based Numbering**: Implemented precise numbering using CSS ::before pseudo-elements
- **TOC Enhancement**: Fixed layout issues and improved visual fidelity
- **Section Navigation**: Added section IDs for direct navigation to numbered headings
- **DOM Serialization**: Enhanced content preservation during HTML processing
- **Header Processing**: Added comprehensive document header extraction and processing
- **Diagnostic Tools**: Enhanced debugging capabilities for troubleshooting
- **Modular Refactoring**: Split large files into focused, maintainable modules while preserving API compatibility
- **Hanging Margins Implementation**: Fixed hanging indent behavior for TOC entries and numbered content
- **Header Image Processing**: Implemented comprehensive header image extraction and positioning functionality
- **HTML Formatting Optimization**: Optimized HTML output with proper formatting
- **Table Formatting Enhancement**: Improved table presentation with professional styling and accessibility
- **Text Formatting Enhancement**: Fixed text formatting conversion with comprehensive style mapping
- **TOC Page Number Removal**: Implemented targeted removal of page numbers from Table of Contents entries
- **Return to Top Button**: Added always-visible return to top button with theme-aware styling, smooth scrolling, and accessibility support

For detailed technical implementation information, see [`docs/architecture.md`](docs/architecture.md).
