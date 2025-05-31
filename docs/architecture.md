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
├── debug-test.js          # Diagnostic tool for troubleshooting
├── lib/                   # Refactored library code
│   ├── index.js           # Main entry point that re-exports public API
│   ├── xml/               # XML parsing utilities
│   │   └── xpath-utils.js # XPath utilities for XML processing
│   ├── parsers/           # DOCX parsing modules
│   │   ├── style-parser.js      # Style parsing functions
│   │   ├── theme-parser.js      # Theme parsing functions
│   │   ├── toc-parser.js        # TOC parsing functions
│   │   ├── numbering-parser.js  # Numbering definition parsing
│   │   ├── numbering-resolver.js # Numbering sequence resolution
│   │   ├── document-parser.js   # Document structure parsing
│   │   ├── metadata-parser.js   # Document metadata parsing
│   │   ├── header-parser.js     # Document header extraction
│   │   └── track-changes-parser.js # Track changes extraction
│   ├── html/              # HTML processing modules
│   │   ├── html-generator.js    # Main HTML generation orchestrator
│   │   ├── generators/          # HTML generation utilities
│   │   │   ├── style-mapping.js     # Style mapping for mammoth conversion
│   │   │   ├── image-processing.js  # Image extraction and processing
│   │   │   ├── html-formatting.js   # HTML formatting and indentation
│   │   │   └── html-processing.js   # Main HTML processing and content manipulation
│   │   ├── processors/          # Content processing modules
│   │   │   ├── heading-processor.js   # Heading processing, numbering, and accessibility
│   │   │   ├── toc-processor.js       # Table of Contents detection, structuring, and linking
│   │   │   └── numbering-processor.js # Numbering and list processing utilities
│   │   ├── structure-processor.js # HTML structure handling
│   │   ├── content-processors.js  # Content element processing orchestrator
│   │   └── element-processors.js  # HTML element processing
│   ├── css/               # CSS generation modules
│   │   ├── css-generator.js     # CSS generation orchestrator
│   │   ├── generators/          # CSS generation utilities
│   │   │   ├── base-styles.js         # Base document styles and utilities
│   │   │   ├── paragraph-styles.js    # Paragraph style generation
│   │   │   ├── character-styles.js    # Character and inline text styles
│   │   │   ├── table-styles.js        # Table styling and border handling
│   │   │   ├── numbering-styles.js    # DOCX numbering and list styles
│   │   │   ├── toc-styles.js          # Table of Contents styling
│   │   │   ├── utility-styles.js      # Utility classes and general styles
│   │   │   └── specialized-styles.js  # Accessibility, track changes, and header styles
│   │   └── style-mapper.js      # Style mapping functions
│   ├── accessibility/     # Accessibility enhancement modules
│   │   └── wcag-processor.js    # WCAG 2.1 compliance processor
│   └── utils/             # Utility functions
│       ├── unit-converter.js    # Unit conversion utilities
│       └── common-utils.js      # Common utility functions
```

### 2.1 Modular Architecture Benefits

The refactored architecture provides several key benefits:

1. **Single Responsibility**: Each module has a focused, well-defined purpose
2. **Improved Maintainability**: Smaller files are easier to understand and modify
3. **Better Testing**: Individual components can be tested in isolation
4. **Enhanced Readability**: Clear separation of concerns makes the codebase more navigable
5. **API Preservation**: All existing exports are maintained for backward compatibility
6. **Scalability**: New features can be added without affecting existing modules

### 2.2 Module Organization

#### CSS Generators (`lib/css/generators/`)

- **base-styles.js**: Document defaults, font utilities, border styles, and fallback CSS
- **paragraph-styles.js**: Word-like paragraph indentation and formatting with enhanced text wrapping
- **character-styles.js**: Inline text styling and character formatting
- **table-styles.js**: Table layout, borders, and cell styling
- **numbering-styles.js**: CSS counters, hierarchical numbering, and hanging indent implementation
- **toc-styles.js**: Table of Contents with block layout, hanging indents, and optional dots/page numbers
- **utility-styles.js**: General utility classes and section navigation
- **specialized-styles.js**: Accessibility, track changes, and header styles

#### HTML Generators (`lib/html/generators/`)

- **style-mapping.js**: Mammoth style mapping configuration
- **image-processing.js**: Image extraction and processing utilities
- **html-formatting.js**: HTML indentation and formatting
- **html-processing.js**: Main HTML processing and DOM manipulation

#### HTML Processors (`lib/html/processors/`)

- **heading-processor.js**: Heading numbering, accessibility, and structure
- **toc-processor.js**: TOC detection, structuring, and navigation links
- **numbering-processor.js**: List processing and numbering application

## 3. Recent Modular Refactoring (v1.3.0)

The v1.3.0 release includes a comprehensive modular refactoring that significantly improves code organization:

### 3.1 CSS Generator Refactoring

**Before**: Single large file (`lib/css/css-generator.js` - 1058 lines)
**After**: 8 focused modules in `lib/css/generators/`

- **base-styles.js**: Base document styles, font utilities, border styles, and fallback CSS
- **paragraph-styles.js**: Paragraph style generation with Word-like indentation
- **character-styles.js**: Character and inline text styles
- **table-styles.js**: Table styling and border handling
- **numbering-styles.js**: DOCX numbering and list styles with CSS counters
- **toc-styles.js**: Table of Contents styling with flex layout
- **utility-styles.js**: Utility classes, section navigation, and heading styles
- **specialized-styles.js**: Accessibility, track changes, and header styles

### 3.2 HTML Generator Refactoring

**Before**: Single large file (`lib/html/html-generator.js` - 832 lines)
**After**: 5 focused modules

- **Main orchestrator**: `lib/html/html-generator.js` (now focused on coordination)
- **Generators**: `lib/html/generators/` (style mapping, image processing, formatting, processing)

### 3.3 Content Processor Refactoring

**Before**: Single large file (`lib/html/content-processors.js` - 832 lines)
**After**: 4 focused modules

- **Main orchestrator**: `lib/html/content-processors.js` (now delegates to processors)
- **Processors**: `lib/html/processors/` (heading, TOC, numbering processors)

### 3.4 Benefits Achieved

1. **Improved Maintainability**: Each module has a single, focused responsibility
2. **Better Code Organization**: Related functionality is grouped together logically
3. **Enhanced Readability**: Smaller files are easier to understand and navigate
4. **Preserved API**: All existing exports are maintained for backward compatibility
5. **Modular Architecture**: Individual components can be tested and modified independently
6. **Clear Separation of Concerns**: CSS generation, HTML processing, and content manipulation are cleanly separated

## 4. Hanging Margins Implementation (v1.3.1)

### 4.1 Overview

The hanging margins feature replicates Microsoft Word's hanging indent behavior in HTML/CSS output. This implementation addresses two key areas: Table of Contents (TOC) entries and numbered content elements.

### 4.2 Technical Implementation

#### 4.2.1 TOC Hanging Indents (`lib/css/generators/toc-styles.js`)

**Problem**: TOC entries were truncating instead of wrapping with proper hanging indents.

**Solution**:

- Converted from flex layout to block layout for proper text wrapping
- Implemented hanging indents using `text-indent: -1.5em` and `padding-left: 1.5em`
- Added comprehensive text wrapping with `overflow-wrap`, `word-wrap`, and `hyphens`
- Optional removal of dots and page numbers for cleaner appearance

**Key CSS Implementation**:

```css
.docx-toc-entry {
  display: block !important;
  text-indent: -1.5em !important;
  padding-left: 1.5em !important;
  overflow-wrap: break-word !important;
  word-wrap: break-word !important;
  hyphens: auto !important;
}
```

#### 4.2.2 Numbered Content Hanging Margins (`lib/css/generators/numbering-styles.js`)

**Problem**: Numbered headings were not displaying with proper hanging margins where numbers should appear at the left margin.

**Solution**:

- Enhanced hanging indent logic to ensure adequate negative text-indent values
- Added adaptive logic: when DOCX hanging indent is too small, use full number region width
- Improved number positioning with CSS `::before` pseudo-elements

**Key Logic**:

```javascript
// Ensure adequate hanging indent for numbered content
if (numRegionWidth > 0 && Math.abs(pTextIndent) < numRegionWidth / 2) {
  pTextIndent = -numRegionWidth;
}
```

#### 4.2.3 Enhanced Text Wrapping (`lib/css/generators/paragraph-styles.js`)

**Enhancement**: Added comprehensive word wrapping support to heading styles for better text flow.

**Implementation**:

```css
word-wrap: break-word;
overflow-wrap: break-word;
hyphens: auto;
```

### 4.3 DOCX XML Introspection

The implementation extracts precise hanging indent measurements from DOCX XML structure:

1. **Paragraph Properties**: Analyzes `w:pPr` elements for indentation values
2. **Numbering Definitions**: Extracts number region widths from `w:numPr` elements
3. **Style Inheritance**: Follows DOCX style hierarchy for accurate measurements
4. **Unit Conversion**: Converts DOCX units (twips, points) to CSS units

### 4.4 CSS Generation Strategy

1. **Negative Text Indent**: Uses negative `text-indent` to pull first line left
2. **Compensating Padding**: Adds equivalent `padding-left` to maintain content alignment
3. **Absolute Positioning**: Positions numbers using `::before` pseudo-elements with precise `left` values
4. **Responsive Design**: Ensures hanging indents work across different screen sizes

### 4.5 Benefits

1. **Word-like Behavior**: Replicates Microsoft Word's hanging indent formatting exactly
2. **Proper Text Wrapping**: Long entries wrap correctly with subsequent lines properly indented
3. **Clean Layout**: Optional removal of TOC decorations for modern, clean appearance
4. **Accessibility**: Maintains proper document structure and reading flow
5. **Content Agnostic**: Works with any document regardless of language or content domain

## 6. Header Image Extraction and Positioning (v1.3.1)

### 6.1 Overview

The header image extraction functionality enables doc2web to extract images from DOCX header files and display them in the HTML output with proper positioning that honors the original document layout.

### 6.2 Technical Implementation

#### 6.2.1 Image Detection and Extraction (`lib/parsers/header-parser.js`)

**DOCX XML Introspection**: The system analyzes DOCX header XML files to detect and extract image information:

1. **Image Element Detection**: Searches for `w:drawing` and `w:pict` elements in header paragraphs
2. **Relationship Resolution**: Parses relationship files (`word/_rels/header*.xml.rels`) to map image IDs to file paths
3. **Image Extraction**: Uses JSZip 3.0 async API to extract image data from the DOCX archive
4. **File System Operations**: Saves images to the output directory with sanitized filenames

#### 6.2.2 Positioning Information Extraction

**Comprehensive Positioning Data**: The system extracts detailed positioning information from DOCX XML:

```javascript
// Positioning data structure
{
  wrapType: 'inline',           // inline vs anchored positioning
  xOffset: '0',                 // horizontal transform offset
  yOffset: '0',                 // vertical transform offset
  distT: '0',                   // top margin distance
  distB: '0',                   // bottom margin distance
  distL: '0',                   // left margin distance
  distR: '0',                   // right margin distance
  horizontal: null,             // alignment (center, left, right)
  vertical: null                // vertical alignment
}
```

#### 6.2.3 HTML Generation with Positioning

**Semantic HTML Structure**: Creates proper HTML structure with positioning applied:

```html
<div class="docx-header-paragraph">
  <div class="docx-image-container">
    <img src="./images/image_name.png" 
         alt="Image Alt Text" 
         class="docx-header-image"
         style="width: 156px; height: 120px; max-width: 100%;">
  </div>
</div>
```

**CSS Positioning Logic**:

1. **Inline Images**: Applied margins and transforms based on DOCX distance attributes
2. **Anchored Images**: Applied absolute positioning with proper offset calculations
3. **Paragraph Alignment**: Applied text-align to container for center/left/right alignment
4. **EMU Conversion**: Proper conversion from DOCX units (1 EMU = 1/914400 inch = 96 pixels)

### 6.3 Key Features

1. **Position Fidelity**: Honors original DOCX image positioning including inline placement, margins, and transform offsets
2. **Multiple Header Support**: Processes images from different header types (first page, even pages, default)
3. **Relationship Resolution**: Correctly resolves image references through DOCX relationship files
4. **Async Processing**: Uses modern async/await patterns for efficient image extraction
5. **Error Handling**: Comprehensive error handling for missing files and processing failures
6. **Semantic HTML**: Creates accessible HTML structure with proper image containers

### 6.4 Integration with Modular Architecture

The header image functionality demonstrates the benefits of the modular architecture:

- **Focused Implementation**: Image processing logic is contained within `header-parser.js`
- **Clean Separation**: Image extraction, positioning, and HTML generation are clearly separated
- **Reusable Components**: Image processing utilities can be reused for other image types
- **Maintainable Code**: Each aspect of image processing can be tested and modified independently

### 6.5 Benefits

1. **Visual Fidelity**: Maintains the exact appearance of header images from the original DOCX
2. **Professional Output**: Preserves corporate logos and header graphics with proper positioning
3. **Responsive Design**: Images adapt to different screen sizes while maintaining proportions
4. **Accessibility**: Proper alt text and semantic HTML structure for screen readers
5. **Content Preservation**: Ensures no header content is lost during conversion

## 7. Conclusion

The doc2web application now features a robust, modular architecture that provides excellent maintainability while preserving all existing functionality. The refactoring into focused, single-responsibility modules makes the codebase much more approachable for developers while maintaining the same powerful conversion capabilities.

The recent hanging margins implementation and header image extraction functionality (v1.3.1) demonstrate the benefits of this modular architecture:

- **Focused Changes**: Hanging indent fixes and image processing were implemented in specific modules without affecting other functionality
- **Clear Separation**: TOC styling, numbering styles, text wrapping, and image positioning were enhanced independently
- **Maintainable Code**: Each module could be modified and tested in isolation
- **Preserved Functionality**: All existing features remained intact while adding new capabilities

The modular design ensures that:

- New features can be added without affecting existing code
- Individual components can be tested and debugged in isolation
- The codebase remains maintainable as it grows
- API compatibility is preserved for existing users
- Code quality and organization continue to improve
- Complex features like hanging margins and header image processing can be implemented systematically across multiple modules

This architecture provides a solid foundation for future enhancements while maintaining the application's core strengths in DOCX conversion, style preservation, and content fidelity. The hanging margins implementation and header image extraction functionality showcase how the modular structure enables sophisticated document formatting features that closely replicate Microsoft Word's behavior, including precise positioning and visual fidelity.
