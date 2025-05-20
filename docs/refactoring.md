# Code Refactoring Documentation

## Overview

This document describes the refactoring of the `style-extractor.js` and `docx-style-parser.js` files into a more modular and maintainable structure. The refactoring was done to improve code organization, reduce file sizes, and make the codebase easier to understand and maintain.

## Directory Structure

The refactored code is organized into the following directory structure:

```
lib/
├── xml/              # XML and XPath utilities
├── parsers/          # DOCX parsing modules
├── html/             # HTML processing modules
├── css/              # CSS generation modules
└── utils/            # Utility functions
```

## Module Descriptions

### XML Utilities (`lib/xml/`)

- **xpath-utils.js**: Provides utilities for working with XML and XPath in DOCX files.
  - `NAMESPACES`: XML namespaces used in DOCX files
  - `createXPathSelector`: Creates an XPath selector with namespaces registered
  - `selectNodes`: Selects nodes using XPath with namespaces
  - `selectSingleNode`: Selects a single node using XPath with namespaces

### Parsers (`lib/parsers/`)

- **style-parser.js**: Main module for parsing DOCX styles.
  - `parseDocxStyles`: Main entry point for extracting styles from DOCX files
  - `parseStyles`: Parses styles.xml to extract style definitions
  - `parseStyleNode`: Parses individual style nodes
  - `parseRunningProperties`: Parses text formatting properties
  - `parseParagraphProperties`: Parses paragraph formatting properties
  - `parseBorders`: Parses border information

- **theme-parser.js**: Parses theme information from DOCX files.
  - `parseTheme`: Extracts theme information (colors, fonts)
  - `getColorValue`: Gets color values from theme color nodes

- **toc-parser.js**: Parses Table of Contents styles.
  - `parseTocStyles`: Extracts TOC styles, heading styles, and leader line information

- **numbering-parser.js**: Parses numbering definitions.
  - `parseNumberingDefinitions`: Extracts numbering definitions and levels
  - `getCSSCounterFormat`: Maps Word numbering formats to CSS counter styles
  - `getCSSCounterContent`: Generates CSS counter content for numbering

- **document-parser.js**: Parses document structure and settings.
  - `parseDocumentDefaults`: Extracts document default styles
  - `parseSettings`: Extracts document settings
  - `analyzeDocumentStructure`: Analyzes document structure for special elements
  - `getDefaultStyleInfo`: Provides default style information if parsing fails

### HTML Processing (`lib/html/`)

- **html-generator.js**: Main module for generating HTML from DOCX content.
  - `extractAndApplyStyles`: Main entry point for extracting and applying styles
  - `convertToStyledHtml`: Converts DOCX to HTML while preserving styles
  - `applyStylesToHtml`: Applies generated CSS to HTML

- **structure-processor.js**: Handles HTML document structure.
  - `ensureHtmlStructure`: Ensures HTML has proper structure (html, head, body)
  - `addDocumentMetadata`: Adds document metadata based on style information

- **content-processors.js**: Processes document content elements.
  - `processHeadings`: Processes headings with numbering
  - `processTOC`: Processes Table of Contents with proper structure
  - `processNestedNumberedParagraphs`: Processes hierarchical lists
  - `identifyListPatterns`: Identifies list patterns in the document
  - `isSpecialParagraph`: Checks if a paragraph is a special type
  - `processSpecialParagraphs`: Processes special paragraph types

- **element-processors.js**: Processes HTML elements.
  - `processTables`: Enhances tables with proper structure
  - `processImages`: Enhances images with proper attributes
  - `processLanguageElements`: Handles language-specific elements

### CSS Generation (`lib/css/`)

- **css-generator.js**: Generates CSS from style information.
  - `generateCssFromStyleInfo`: Main function for generating CSS
  - `getFontFamily`: Gets font family based on style and theme
  - `getBorderStyle`: Gets border style based on style information

- **style-mapper.js**: Maps DOCX styles to CSS classes.
  - `createStyleMap`: Creates style map for mammoth.js
  - `createDocumentTransformer`: Creates document transformer for style preservation

### Utilities (`lib/utils/`)

- **unit-converter.js**: Provides unit conversion utilities.
  - `convertTwipToPt`: Converts twips to points
  - `convertBorderSizeToPt`: Converts border size to points
  - `getBorderTypeValue`: Maps Word border types to CSS border styles

- **common-utils.js**: Common utility functions.
  - `getLeaderChar`: Gets leader character based on leader type

## Main Entry Points

- **lib/index.js**: Re-exports the public API for backward compatibility.
- **style-extractor.js**: Thin wrapper that imports from the new structure.
- **docx-style-parser.js**: Thin wrapper that imports from the new structure.

## Data Flow

1. The process starts with `extractAndApplyStyles` in `html-generator.js`.
2. It calls `parseDocxStyles` from `style-parser.js` to extract style information.
3. The style information is used to generate CSS with `generateCssFromStyleInfo` from `css-generator.js`.
4. The DOCX is converted to HTML using mammoth.js with custom style mapping from `style-mapper.js`.
5. The HTML is enhanced with proper structure using `structure-processor.js`.
6. The content is processed using various processors from `content-processors.js` and `element-processors.js`.
7. The final HTML with external CSS is returned.

## Circular Dependencies

To avoid circular dependencies, the following strategies were used:

1. Moving shared functions to utility modules (e.g., `getLeaderChar` to `common-utils.js`).
2. Using dynamic imports for functions that would create circular dependencies.
3. Organizing modules in a hierarchical manner to minimize cross-dependencies.

## Backward Compatibility

Backward compatibility is maintained through:

1. The `lib/index.js` file that re-exports the public API.
2. The original `style-extractor.js` and `docx-style-parser.js` files that now import from the new structure.
3. Preserving the same function signatures and return values for public API functions.

## Future Improvements

Potential future improvements include:

1. Converting to ES modules for better tree-shaking and module resolution.
2. Adding TypeScript type definitions for better IDE support.
3. Implementing more comprehensive error handling and recovery.
4. Adding unit tests for each module.
5. Further optimizing performance for large documents.