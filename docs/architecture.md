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
- **paragraph-styles.js**: Word-like paragraph indentation and formatting
- **character-styles.js**: Inline text styling and character formatting
- **table-styles.js**: Table layout, borders, and cell styling
- **numbering-styles.js**: CSS counters and hierarchical numbering
- **toc-styles.js**: Table of Contents with flex layout and leader dots
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

## 4. Conclusion

The doc2web application now features a robust, modular architecture that provides excellent maintainability while preserving all existing functionality. The refactoring into focused, single-responsibility modules makes the codebase much more approachable for developers while maintaining the same powerful conversion capabilities.

The modular design ensures that:
- New features can be added without affecting existing code
- Individual components can be tested and debugged in isolation
- The codebase remains maintainable as it grows
- API compatibility is preserved for existing users
- Code quality and organization continue to improve

This architecture provides a solid foundation for future enhancements while maintaining the application's core strengths in DOCX conversion, style preservation, and content fidelity.
