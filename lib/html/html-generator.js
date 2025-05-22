// lib/html/html-generator.js - Fixed HTML generation with better content preservation
const mammoth = require("mammoth");
const { JSDOM } = require("jsdom");
const path = require("path");
const fs = require("fs");
const JSZip = require("jszip");
const { DOMParser } = require("xmldom");

// Import required modules
const { parseDocxStyles } = require("../parsers/style-parser");
const { generateCssFromStyleInfo } = require("../css/css-generator");
const { createStyleMap, createDocumentTransformer } = require("../css/style-mapper");
const { ensureHtmlStructure, addDocumentMetadata } = require("./structure-processor");
const { processHeadings, processTOC, processNestedNumberedParagraphs, processSpecialParagraphs } = require("./content-processors");
const { processTables, processImages, processLanguageElements } = require("./element-processors");
const { parseDocumentMetadata } = require("../parsers/metadata-parser");
const { parseTrackChanges } = require("../parsers/track-changes-parser");
const { selectNodes, selectSingleNode } = require("../xml/xpath-utils");

/**
 * Extract and apply styles from a DOCX document to HTML with numbering context
 * Enhanced main entry point that includes DOCX introspection for numbering
 * 
 * @param {string} docxPath - Path to the DOCX file
 * @param {string} cssFilename - Filename for the CSS file (without path)
 * @param {Object} options - Processing options
 * @returns {Promise<{html: string, styles: string, numberingContext: Array}>} - HTML with embedded styles and numbering context
 */
async function extractAndApplyStyles(docxPath, cssFilename = null, options = {}) {
  try {
    console.log(`Starting DOCX processing for: ${docxPath}`);
    
    // Validate input file
    if (!fs.existsSync(docxPath)) {
      throw new Error(`DOCX file not found: ${docxPath}`);
    }
    
    // Read and unpack the DOCX file
    const data = fs.readFileSync(docxPath);
    const zip = await JSZip.loadAsync(data);
    
    // Extract core XML files for content and styles
    const styleXml = await zip.file('word/styles.xml')?.async('string');
    const documentXml = await zip.file('word/document.xml')?.async('string');
    const themeXml = await zip.file('word/theme/theme1.xml')?.async('string');
    const settingsXml = await zip.file('word/settings.xml')?.async('string');
    const numberingXml = await zip.file('word/numbering.xml')?.async('string');
    
    // Extract metadata XML files
    const corePropsXml = await zip.file('docProps/core.xml')?.async('string');
    const appPropsXml = await zip.file('docProps/app.xml')?.async('string');
    
    // Validate required files
    if (!styleXml || !documentXml) {
      throw new Error('Invalid DOCX file: missing required XML files (styles.xml or document.xml)');
    }
    
    // Parse XML content
    const styleDoc = new DOMParser().parseFromString(styleXml);
    const documentDoc = new DOMParser().parseFromString(documentXml);
    const themeDoc = themeXml ? new DOMParser().parseFromString(themeXml) : null;
    const settingsDoc = settingsXml ? new DOMParser().parseFromString(settingsXml) : null;
    const numberingDoc = numberingXml ? new DOMParser().parseFromString(numberingXml) : null;
    const corePropsDoc = corePropsXml ? new DOMParser().parseFromString(corePropsXml) : null;
    const appPropsDoc = appPropsXml ? new DOMParser().parseFromString(appPropsXml) : null;
    
    // Extract the raw styles from the document with enhanced numbering context
    console.log('Extracting styles and numbering context from DOCX...');
    const styleInfo = await parseDocxStyles(docxPath);
    console.log(`Extracted ${Object.keys(styleInfo.styles.paragraph || {}).length} paragraph styles`);

    // Extract document metadata
    console.log('Extracting document metadata...');
    const metadata = parseDocumentMetadata(corePropsDoc, appPropsDoc, documentDoc);
    
    // Extract track changes information
    console.log('Extracting track changes information...');
    const trackChanges = parseTrackChanges(documentDoc);

    // Generate CSS from the extracted styles with numbering support
    console.log('Generating CSS with numbering support...');
    const css = generateCssFromStyleInfo(styleInfo);
    console.log(`Generated CSS length: ${css.length} characters`);

    // Convert DOCX to HTML with enhanced style preservation and numbering context
    console.log('Converting DOCX to HTML with numbering preservation...');
    const htmlResult = await convertToStyledHtml(docxPath, styleInfo);
    console.log(`Initial HTML conversion result length: ${htmlResult.value.length} characters`);

    // Get the CSS filename
    const cssFile = cssFilename || path.basename(docxPath, path.extname(docxPath)) + ".css";

    // Combine HTML and CSS with enhanced features and numbering context
    console.log('Applying styles and processing document structure...');
    const styledHtml = applyStylesToHtml(
      htmlResult.value,
      css,
      styleInfo,
      cssFile,
      metadata,
      trackChanges,
      options
    );
    
    console.log(`Final HTML length: ${styledHtml.length} characters`);

    return {
      html: styledHtml,
      styles: css,
      messages: htmlResult.messages,
      metadata,
      trackChanges,
      // Include numbering context in the result
      numberingContext: styleInfo.numberingContext || []
    };
  } catch (error) {
    console.error("Error extracting styles:", error);
    console.error("Stack trace:", error.stack);
    throw error;
  }
}

/**
 * Convert DOCX to HTML while preserving style information with numbering context
 * Enhanced to create style mapping that preserves numbering attributes
 * 
 * @param {string} docxPath - Path to the DOCX file
 * @param {Object} styleInfo - Extracted style information with numbering context
 * @returns {Promise<{value: string, messages: Array}>} - HTML with style information
 */
async function convertToStyledHtml(docxPath, styleInfo) {
  try {
    console.log('Creating enhanced style map...');
    // Create a custom style map based on extracted styles with numbering support
    const styleMap = createEnhancedStyleMap(styleInfo);
    console.log(`Created style map with ${styleMap.length} entries`);

    // Custom document transformer to enhance style preservation with numbering
    const transformDocument = createEnhancedDocumentTransformer(styleInfo);

    // Configure image handling
    const imageOptions = {
      convertImage: mammoth.images.imgElement(function (image) {
        return image.read("base64").then(function (imageBuffer) {
          try {
            const extension = image.contentType.split("/")[1] || 'png';
            const filename = `image-${image.altText ? image.altText.replace(/[^a-zA-Z0-9]/g, '_') : Date.now()}.${extension}`;

            return {
              src: `./images/${filename}`,
              alt: image.altText || "Document image",
              className: "docx-image",
            };
          } catch (error) {
            console.error('Error processing image:', error);
            return {
              src: `./images/image-${Date.now()}.png`,
              alt: "Document image",
              className: "docx-image",
            };
          }
        });
      }),
    };

    console.log('Starting mammoth conversion...');
    // Use mammoth to convert with the enhanced style map
    const result = await mammoth.convertToHtml({
      path: docxPath,
      styleMap: styleMap,
      transformDocument: transformDocument,
      includeDefaultStyleMap: true,
      ...imageOptions,
    });

    console.log(`Mammoth conversion completed. HTML length: ${result.value.length} characters`);
    
    // Log any conversion messages
    if (result.messages && result.messages.length > 0) {
      console.log('Mammoth conversion messages:');
      result.messages.forEach(msg => console.log(`  ${msg.type}: ${msg.message}`));
    }
    
    // Validate the result
    if (!result.value || result.value.length < 10) {
      console.warn('WARNING: Mammoth conversion result is very short or empty!');
      console.log('Result value:', result.value);
      
      // Try alternative conversion without style map
      console.log('Attempting fallback conversion without style map...');
      const fallbackResult = await mammoth.convertToHtml({
        path: docxPath,
        includeDefaultStyleMap: true,
        ...imageOptions,
      });
      
      if (fallbackResult.value && fallbackResult.value.length > result.value.length) {
        console.log('Using fallback conversion result');
        return fallbackResult;
      }
    }

    return result;
  } catch (error) {
    console.error('Error in convertToStyledHtml:', error);
    
    // Fallback: try basic conversion
    try {
      console.log('Attempting basic fallback conversion...');
      const fallbackResult = await mammoth.convertToHtml({ path: docxPath });
      console.log(`Fallback conversion completed. HTML length: ${fallbackResult.value.length} characters`);
      return fallbackResult;
    } catch (fallbackError) {
      console.error('Fallback conversion also failed:', fallbackError);
      throw error;
    }
  }
}

/**
 * Create enhanced style map that preserves numbering attributes
 * @param {Object} styleInfo - Style information with numbering context
 * @returns {Array} - Enhanced style map entries
 */
function createEnhancedStyleMap(styleInfo) {
  const styleMap = [];
  
  try {
    // Enhanced paragraph style mapping with numbering preservation
    Object.entries(styleInfo.styles.paragraph || {}).forEach(([id, style]) => {
      let mapping = `p[style-name='${style.name}'] => p.docx-p-${id.toLowerCase().replace(/[^a-z0-9]/g, '-')}`;
      
      // Add numbering attributes if the style has numbering
      if (style.numbering && style.numbering.hasNumbering) {
        mapping += `[data-num-id='${style.numbering.id}'][data-num-level='${style.numbering.level}']`;
      }
      
      styleMap.push(mapping);
    });
    
    // Enhanced character style mapping
    Object.entries(styleInfo.styles.character || {}).forEach(([id, style]) => {
      styleMap.push(
        `r[style-name='${style.name}'] => span.docx-c-${id.toLowerCase().replace(/[^a-z0-9]/g, '-')}`
      );
    });
    
    // Enhanced table style mapping
    Object.entries(styleInfo.styles.table || {}).forEach(([id, style]) => {
      styleMap.push(
        `table[style-name='${style.name}'] => table.docx-t-${id.toLowerCase().replace(/[^a-z0-9]/g, '-')}`
      );
    });
    
    // Enhanced heading mappings with numbering support
    styleMap.push("p[style-name='heading 1'] => h1.docx-heading1");
    styleMap.push("p[style-name='heading 2'] => h2.docx-heading2");
    styleMap.push("p[style-name='heading 3'] => h3.docx-heading3");
    styleMap.push("p[style-name='heading 4'] => h4.docx-heading4");
    styleMap.push("p[style-name='heading 5'] => h5.docx-heading5");
    styleMap.push("p[style-name='heading 6'] => h6.docx-heading6");
    
    // Handle headings with styles like "Heading1"
    styleMap.push("p[style-name='Heading1'] => h1.docx-heading1");
    styleMap.push("p[style-name='Heading2'] => h2.docx-heading2");
    styleMap.push("p[style-name='Heading3'] => h3.docx-heading3");
    styleMap.push("p[style-name='Heading4'] => h4.docx-heading4");
    styleMap.push("p[style-name='Heading5'] => h5.docx-heading5");
    styleMap.push("p[style-name='Heading6'] => h6.docx-heading6");
    
    // Enhanced TOC style mappings
    styleMap.push("p[style-name='TOC Heading'] => h2.docx-toc-heading");
    styleMap.push("p[style-name='TOC1'] => p.docx-toc-entry.docx-toc-level-1");
    styleMap.push("p[style-name='TOC2'] => p.docx-toc-entry.docx-toc-level-2");
    styleMap.push("p[style-name='TOC3'] => p.docx-toc-entry.docx-toc-level-3");
    styleMap.push("p[style-name='toc 1'] => p.docx-toc-entry.docx-toc-level-1");
    styleMap.push("p[style-name='toc 2'] => p.docx-toc-entry.docx-toc-level-2");
    styleMap.push("p[style-name='toc 3'] => p.docx-toc-entry.docx-toc-level-3");
    
    // Standard text formatting
    styleMap.push("p:fresh => p");
    styleMap.push("r[bold] => strong");
    styleMap.push("r[italic] => em");
    styleMap.push("r[underline] => span.docx-underline");
    styleMap.push("r[strikethrough] => span.docx-strike");
    styleMap.push("r[subscript] => sub");
    styleMap.push("r[superscript] => sup");
    
    // General paragraph styles
    styleMap.push("p[style-name='Normal Web'] => p.docx-normalweb"); 
    styleMap.push("p[style-name='Body Text'] => p.docx-bodytext");
    styleMap.push("p[style-name='Normal'] => p.docx-normal");
    
  } catch (error) {
    console.error('Error creating enhanced style map:', error);
  }
  
  return styleMap;
}

/**
 * Create enhanced document transformer with numbering preservation
 * @param {Object} styleInfo - Style information with numbering context
 * @returns {Function} - Enhanced document transformer function
 */
function createEnhancedDocumentTransformer(styleInfo) {
  return function (document) {
    try {
      // Enhanced document transformation that preserves numbering context
      // This function can walk the document tree and add numbering attributes
      // based on the numbering context extracted from the DOCX
      
      if (styleInfo.numberingContext && styleInfo.numberingContext.length > 0) {
        console.log(`Processing ${styleInfo.numberingContext.length} numbered paragraphs`);
        
        // Add data attributes to preserve numbering context for later processing
        if (document.paragraphs) {
          document.paragraphs.forEach((paragraph, index) => {
            // Find corresponding numbering context
            const context = styleInfo.numberingContext.find(ctx => 
              ctx.paragraphIndex === index
            );
            
            if (context && context.numberingId) {
              // This would require extending mammoth.js to support custom attributes
              // For now, we'll handle this in the post-processing step
              console.log(`Found numbered paragraph at index ${index}: ${context.textContent?.substring(0, 50)}...`);
            }
          });
        }
      }
      
      return document;
    } catch (error) {
      console.error('Error in document transformer:', error);
      return document;
    }
  };
}

/**
 * Apply generated CSS to HTML with enhanced numbering processing
 * Enhanced to include DOCX-derived numbering context
 * 
 * @param {string} html - HTML content
 * @param {string} css - CSS styles
 * @param {Object} styleInfo - Style information with numbering context
 * @param {string} cssFilename - Filename for the CSS file
 * @param {Object} metadata - Document metadata
 * @param {Object} trackChanges - Track changes information
 * @param {Object} options - Processing options
 * @returns {string} - HTML with link to external CSS and enhanced numbering
 */
function applyStylesToHtml(html, css, styleInfo, cssFilename, metadata, trackChanges, options = {}) {
  try {
    console.log(`Starting HTML enhancement. Input HTML length: ${html.length} characters`);
    
    // Default options
    const defaultOptions = {
      trackChangesMode: 'show', // 'show', 'hide', 'accept', or 'reject'
      enhanceAccessibility: true, // Apply WCAG 2.1 Level AA enhancements
      preserveMetadata: true, // Add metadata to HTML
      showAuthor: true, // Show author in track changes
      showDate: true // Show date in track changes
    };
    
    const opts = { ...defaultOptions, ...options };
    
    // Validate input HTML
    if (!html || html.trim().length === 0) {
      console.error('ERROR: Input HTML is empty or invalid');
      return '<html><head><meta charset="utf-8"></head><body><p>Error: Document conversion failed - no content generated</p></body></html>';
    }
    
    // Create a proper HTML document structure
    let fullHtml = html;
    
    // Only wrap in full HTML structure if it's not already a complete document
    if (!html.includes('<!DOCTYPE html>') && !html.includes('<html')) {
      fullHtml = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Document</title>
  <link rel="stylesheet" href="./${cssFilename}">
</head>
<body>
${html}
</body>
</html>`;
    } else {
      // If it's already a complete document, just ensure CSS link is present
      if (!html.includes(cssFilename)) {
        fullHtml = html.replace(
          '</head>',
          `  <link rel="stylesheet" href="./${cssFilename}">\n</head>`
        );
      }
    }
    
    console.log(`Created full HTML structure. Length: ${fullHtml.length} characters`);
    
    // Create a DOM to manipulate the HTML
    const dom = new JSDOM(fullHtml, {
      contentType: "text/html",
      includeNodeLocations: false,
      storageQuota: 10000000
    });
    const document = dom.window.document;

    // Verify the document was created successfully
    if (!document || !document.body) {
      console.error('ERROR: Failed to create DOM document');
      return fullHtml; // Return the original HTML if DOM creation fails
    }
    
    console.log(`DOM created successfully. Body children count: ${document.body.children.length}`);

    // Ensure we have a proper HTML structure
    try {
      ensureHtmlStructure(document);
    } catch (error) {
      console.error("Error ensuring HTML structure:", error);
    }

    // Add metadata if requested
    if (opts.preserveMetadata && metadata) {
      try {
        const { applyMetadataToHtml } = require('../parsers/metadata-parser');
        applyMetadataToHtml(document, metadata);
        console.log('Applied document metadata');
      } catch (error) {
        console.error("Error applying metadata:", error);
      }
    }
    
    // Add body class for RTL if needed
    if (styleInfo.settings?.rtlGutter) {
      document.body.classList.add("docx-rtl");
    }

    // Process document elements with error handling for each step
    const processingSteps = [
      { name: 'tables', fn: () => processTables(document) },
      { name: 'images', fn: () => processImages(document) },
      { name: 'language elements', fn: () => processLanguageElements(document) },
      { name: 'headings', fn: () => processHeadings(document, styleInfo, styleInfo.numberingContext) },
      { name: 'TOC', fn: () => processTOC(document, styleInfo) },
      { name: 'numbered paragraphs', fn: () => processNestedNumberedParagraphs(document, styleInfo, styleInfo.numberingContext) },
      { name: 'special paragraphs', fn: () => processSpecialParagraphs(document, styleInfo) }
    ];

    for (const step of processingSteps) {
      try {
        console.log(`Processing ${step.name}...`);
        step.fn();
        console.log(`Completed processing ${step.name}`);
      } catch (error) {
        console.error(`Error processing ${step.name}:`, error);
        // Continue with other processing steps
      }
    }
    
    // Process track changes if present
    if (trackChanges && trackChanges.hasTrackedChanges) {
      try {
        console.log('Processing track changes...');
        const { processTrackChanges } = require('../parsers/track-changes-parser');
        processTrackChanges(document, trackChanges, {
          mode: opts.trackChangesMode,
          showAuthor: opts.showAuthor,
          showDate: opts.showDate
        });
        console.log('Completed track changes processing');
      } catch (error) {
        console.error("Error processing track changes:", error);
      }
    }
    
    // Enhance accessibility if requested
    if (opts.enhanceAccessibility) {
      try {
        console.log('Enhancing accessibility...');
        const { processForAccessibility } = require('../accessibility/wcag-processor');
        processForAccessibility(document, styleInfo, metadata);
        console.log('Completed accessibility enhancement');
      } catch (error) {
        console.error("Error enhancing accessibility:", error);
      }
    }

    // Log numbering context information for debugging
    if (styleInfo.numberingContext && styleInfo.numberingContext.length > 0) {
      const numberedContexts = styleInfo.numberingContext.filter(ctx => ctx.resolvedNumbering);
      console.log(`Applied DOCX numbering to ${numberedContexts.length} paragraphs from ${styleInfo.numberingContext.length} total contexts`);
    }

    // Verify content is still present before serialization
    const bodyContent = document.body.innerHTML.trim();
    if (!bodyContent || bodyContent.length < 10) {
      console.error('ERROR: Body content lost during processing!');
      // Try to restore original content
      document.body.innerHTML = html;
      console.log('Restored original HTML content to body');
    }
    
    console.log(`Final body content length: ${document.body.innerHTML.length} characters`);
    
    // Serialize back to HTML string
    const serialized = dom.serialize();
    
    console.log(`Serialization completed. Final HTML length: ${serialized.length} characters`);
    
    // Final validation
    if (serialized.length < fullHtml.length / 2) {
      console.warn('WARNING: Serialized HTML is significantly shorter than input. Possible content loss.');
      // Return the enhanced fullHtml instead
      return fullHtml;
    }
    
    // Format the HTML for better readability and debugging
    const formattedHtml = formatHtml(serialized);
    console.log(`HTML formatting completed. Final formatted HTML length: ${formattedHtml.length} characters`);
    
    return formattedHtml;
  } catch (error) {
    console.error("Error in applyStylesToHtml:", error);
    console.error("Stack trace:", error.stack);
    
    // Return a basic HTML structure with the original content
    const fallbackHtml = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Document</title>
  <link rel="stylesheet" href="./${cssFilename}">
</head>
<body>
${html}
</body>
</html>`;
    
    console.log('Returning fallback HTML due to processing error');
    // Format the fallback HTML for better readability
    return formatHtml(fallbackHtml);
  }
}

/**
 * Extract images from a DOCX file
 * Extracts images and saves them to the output directory
 * 
 * @param {string} docxPath - Path to the DOCX file
 * @param {string} outputDir - Directory to save images
 * @returns {Promise<void>}
 */
async function extractImagesFromDocx(docxPath, outputDir) {
  try {
    console.log(`Extracting images from ${docxPath} to ${outputDir}`);
    
    // Ensure output directory exists
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }
    
    const options = {
      path: docxPath,
      encoding: 'utf8',
      convertImage: mammoth.images.imgElement(function(image) {
        return image.read("base64").then(function(imageBuffer) {
          try {
            const extension = image.contentType.split('/')[1] || 'png';
            const sanitizedAltText = image.altText ? 
              image.altText.replace(/[^a-zA-Z0-9]/g, '_').substring(0, 20) : 
              'image';
            const filename = `${sanitizedAltText}_${Date.now()}.${extension}`;
            
            // Save the image to output directory
            const imagePath = path.join(outputDir, filename);
            fs.writeFileSync(imagePath, Buffer.from(imageBuffer, 'base64'));
            console.log(`Saved image: ${filename}`);
            
            return {
              src: `./images/${filename}`,
              alt: image.altText || 'Document image'
            };
          } catch (error) {
            console.error('Error saving image:', error);
            return {
              src: `./images/image_${Date.now()}.png`,
              alt: 'Document image'
            };
          }
        });
      })
    };
    
    // Convert to HTML with images
    const result = await mammoth.convertToHtml(options);
    console.log('Images extracted successfully.');
    
    return result;
  } catch (error) {
    console.error('Error extracting images:', error);
    throw error;
  }
}

/**
 * Format HTML with proper indentation and line breaks for better readability
 * @param {string} html - The HTML string to format
 * @returns {string} - Formatted HTML string
 */
function formatHtml(html) {
  try {
    console.log('Formatting HTML for better readability...');
    
    // Replace all self-closing tags to have a space before the closing bracket
    html = html.replace(/<([^>]+?)([^\s])\/>/g, '<$1$2 />');
    
    // Initial formatting - add line breaks for major elements
    const majorElements = ['<!DOCTYPE', '<html', '</html>', '<head', '</head>', '<body', '</body>', 
                          '<div', '</div>', '<nav', '</nav>', '<section', '</section>', 
                          '<article', '</article>', '<h1', '</h1>', '<h2', '</h2>', 
                          '<h3', '</h3>', '<h4', '</h4>', '<h5', '</h5>', '<h6', '</h6>',
                          '<p', '</p>', '<ul', '</ul>', '<ol', '</ol>', '<li', '</li>',
                          '<table', '</table>', '<tr', '</tr>', '<td', '</td>', '<th', '</th>',
                          '<thead', '</thead>', '<tbody', '</tbody>', '<tfoot', '</tfoot>',
                          '<script', '</script>', '<style', '</style>', '<link', '<meta'];
    
    // Add newlines before major elements
    majorElements.forEach(element => {
      const regex = new RegExp(element, 'g');
      html = html.replace(regex, '\n' + element);
    });
    
    // Split the HTML into lines
    let lines = html.split('\n');
    let formattedLines = [];
    let indentLevel = 0;
    
    // Process each line
    for (let i = 0; i < lines.length; i++) {
      let line = lines[i].trim();
      if (!line) continue; // Skip empty lines
      
      // Check if this line is a closing tag or self-closing tag
      const isClosingTag = line.startsWith('</');
      const isSelfClosingTag = !isClosingTag && (line.endsWith('/>') || 
                               line.startsWith('<meta') || 
                               line.startsWith('<link') || 
                               line.startsWith('<img') ||
                               line.startsWith('<br') ||
                               line.startsWith('<hr'));
      
      // Adjust indent level for closing tags
      if (isClosingTag) {
        indentLevel = Math.max(0, indentLevel - 1);
      }
      
      // Add the line with proper indentation
      const indent = '  '.repeat(indentLevel);
      formattedLines.push(indent + line);
      
      // Increase indent level for opening tags (not self-closing)
      if (!isClosingTag && !isSelfClosingTag && line.startsWith('<')) {
        indentLevel++;
      }
    }
    
    // Join the lines back together
    const formattedHtml = formattedLines.join('\n');
    
    // Add a final newline
    return formattedHtml + '\n';
  } catch (error) {
    console.error('Error formatting HTML:', error);
    // Return the original HTML if formatting fails
    return html;
  }
}

module.exports = {
  extractAndApplyStyles,
  convertToStyledHtml,
  applyStylesToHtml,
  extractImagesFromDocx
};