# User Guide for Refactored Code

## Introduction

This guide explains how to use the refactored doc2web code. The codebase has been reorganized into a more modular structure, but maintains backward compatibility with existing code.

## Basic Usage

The main entry points remain the same, so you can continue to use the code as before:

```javascript
const { extractAndApplyStyles } = require('./style-extractor');
const { parseDocxStyles, generateCssFromStyleInfo } = require('./docx-style-parser');

// Extract and apply styles from a DOCX file
async function processDocument(docxPath) {
  const result = await extractAndApplyStyles(docxPath);
  console.log('HTML generated:', result.html);
  console.log('CSS generated:', result.styles);
}
```

## Advanced Usage

If you want to use the new modular structure directly, you can import from the specific modules:

```javascript
// Import from the main index
const { extractAndApplyStyles, parseDocxStyles, generateCssFromStyleInfo } = require('./lib');

// Or import from specific modules
const { extractAndApplyStyles } = require('./lib/html/html-generator');
const { parseDocxStyles } = require('./lib/parsers/style-parser');
const { generateCssFromStyleInfo } = require('./lib/css/css-generator');
```

## Working with Specific Modules

### XML Utilities

```javascript
const { selectNodes, selectSingleNode } = require('./lib/xml/xpath-utils');

// Select nodes using XPath
const nodes = selectNodes("//w:p", xmlDocument);

// Select a single node
const node = selectSingleNode("//w:style[@w:styleId='Heading1']", xmlDocument);
```

### Style Parsing

```javascript
const { parseDocxStyles } = require('./lib/parsers/style-parser');

// Parse styles from a DOCX file
async function parseStyles(docxPath) {
  const styleInfo = await parseDocxStyles(docxPath);
  console.log('Paragraph styles:', Object.keys(styleInfo.styles.paragraph));
  console.log('Character styles:', Object.keys(styleInfo.styles.character));
  console.log('Table styles:', Object.keys(styleInfo.styles.table));
}
```

### CSS Generation

```javascript
const { generateCssFromStyleInfo } = require('./lib/css/css-generator');

// Generate CSS from style information
function generateCSS(styleInfo) {
  const css = generateCssFromStyleInfo(styleInfo);
  return css;
}
```

### HTML Processing

```javascript
const { processHeadings, processTOC } = require('./lib/html/content-processors');
const { processTables, processImages } = require('./lib/html/element-processors');

// Process HTML content
function enhanceHTML(document, styleInfo) {
  // Process headings
  processHeadings(document, styleInfo);
  
  // Process TOC
  processTOC(document, styleInfo);
  
  // Process tables
  processTables(document);
  
  // Process images
  processImages(document);
}
```

## Customizing the Output

### Custom Style Mapping

You can customize the style mapping by modifying the `createStyleMap` function in `lib/css/style-mapper.js`:

```javascript
const { createStyleMap } = require('./lib/css/style-mapper');

// Create a custom style map
function createCustomStyleMap(styleInfo) {
  const styleMap = createStyleMap(styleInfo);
  
  // Add custom mappings
  styleMap.push("p[style-name='MyCustomStyle'] => p.my-custom-class");
  
  return styleMap;
}
```

### Custom CSS Generation

You can customize the CSS generation by extending the `generateCssFromStyleInfo` function:

```javascript
const { generateCssFromStyleInfo } = require('./lib/css/css-generator');

// Generate custom CSS
function generateCustomCSS(styleInfo) {
  let css = generateCssFromStyleInfo(styleInfo);
  
  // Add custom CSS
  css += `
  .my-custom-class {
    color: blue;
    font-weight: bold;
  }
  `;
  
  return css;
}
```

## Error Handling

The refactored code includes improved error handling. You can catch and handle errors as follows:

```javascript
const { extractAndApplyStyles } = require('./style-extractor');

async function processDocumentWithErrorHandling(docxPath) {
  try {
    const result = await extractAndApplyStyles(docxPath);
    return result;
  } catch (error) {
    console.error('Error processing document:', error.message);
    // Handle specific error types
    if (error.message.includes('missing core XML files')) {
      console.error('The DOCX file is invalid or corrupted.');
    }
    // Return a default result or rethrow
    throw error;
  }
}
```

## Performance Considerations

The refactored code maintains the same performance characteristics as the original code. For large documents, consider the following:

1. Use the `convertToStyledHtml` function with custom options to optimize memory usage.
2. Process documents in batches rather than all at once.
3. Use streams for reading and writing files when processing multiple documents.

## Extending the Code

To extend the functionality, you can add new modules or enhance existing ones:

1. Add new parsers in the `lib/parsers/` directory.
2. Add new content processors in the `lib/html/content-processors.js` file.
3. Add new utility functions in the `lib/utils/` directory.
4. Enhance the CSS generation in the `lib/css/css-generator.js` file.

Make sure to update the exports in `lib/index.js` if you add new public API functions.

## Troubleshooting

### Common Issues

1. **DOMParser errors**: Make sure you're using the correct DOMParser from xmldom.
2. **Circular dependencies**: If you encounter circular dependency issues, consider moving shared functions to utility modules.
3. **Missing styles**: Check that the DOCX file contains the expected styles and that the style extraction is working correctly.
4. **Memory issues**: For large documents, consider processing them in chunks or optimizing memory usage.

### Debugging

For debugging, you can add console logs at key points:

```javascript
const { parseDocxStyles } = require('./lib/parsers/style-parser');

async function debugStyles(docxPath) {
  console.log('Parsing styles from:', docxPath);
  try {
    const styleInfo = await parseDocxStyles(docxPath);
    console.log('Style info:', JSON.stringify(styleInfo, null, 2));
    return styleInfo;
  } catch (error) {
    console.error('Error parsing styles:', error);
    throw error;
  }
}
```

## Conclusion

The refactored code provides a more modular and maintainable structure while maintaining backward compatibility. You can continue to use the code as before or take advantage of the new modular structure for more advanced usage.