// lib/html/html-generator.js - Enhanced HTML generation with numbering context
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

    // Extract document metadata
    const metadata = parseDocumentMetadata(corePropsDoc, appPropsDoc);
    
    // Extract track changes information
    const trackChanges = parseTrackChanges(documentDoc);

    // Generate CSS from the extracted styles with numbering support
    console.log('Generating CSS with numbering support...');
    const css = generateCssFromStyleInfo(styleInfo);

    // Convert DOCX to HTML with enhanced style preservation and numbering context
    console.log('Converting DOCX to HTML with numbering preservation...');
    const htmlResult = await convertToStyledHtml(docxPath, styleInfo);

    // Get the CSS filename
    const cssFile =
      cssFilename || path.basename(docxPath, path.extname(docxPath)) + ".css";

    // Combine HTML and CSS with enhanced features and numbering context
    console.log('Applying styles and processing numbering...');
    const styledHtml = applyStylesToHtml(
      htmlResult.value,
      css,
      styleInfo,
      cssFile,
      metadata,
      trackChanges,
      options
    );

    return {
      html: styledHtml,
      styles: css,
      messages: htmlResult.messages,
      metadata,
      trackChanges,
      // NEW: Include numbering context in the result
      numberingContext: styleInfo.numberingContext || []
    };
  } catch (error) {
    console.error("Error extracting styles:", error);
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
  // Create a custom style map based on extracted styles with numbering support
  const styleMap = createEnhancedStyleMap(styleInfo);

  // Custom document transformer to enhance style preservation with numbering
  const transformDocument = createEnhancedDocumentTransformer(styleInfo);

  // Configure image handling
  const imageOptions = {
    convertImage: mammoth.images.imgElement(function (image) {
      return image.read("base64").then(function (imageBuffer) {
        const extension = image.contentType.split("/")[1];
        const filename = `image-${image.altText || Date.now()}.${extension}`;

        return {
          src: `./images/${filename}`,
          alt: image.altText || "Image",
          className: "docx-image",
        };
      });
    }),
  };

  // Use mammoth to convert with the enhanced style map
  const result = await mammoth.convertToHtml({
    path: docxPath,
    styleMap: styleMap,
    transformDocument: transformDocument,
    includeDefaultStyleMap: true,
    ...imageOptions,
  });

  return result;
}

/**
 * Create enhanced style map that preserves numbering attributes
 * @param {Object} styleInfo - Style information with numbering context
 * @returns {Array} - Enhanced style map entries
 */
function createEnhancedStyleMap(styleInfo) {
  const styleMap = [];
  
  // Enhanced paragraph style mapping with numbering preservation
  Object.entries(styleInfo.styles.paragraph || {}).forEach(([id, style]) => {
    let mapping = `p[style-name='${style.name}'] => p.docx-p-${id.toLowerCase()}`;
    
    // Add numbering attributes if the style has numbering
    if (style.numbering && style.numbering.hasNumbering) {
      mapping += `[data-num-id='${style.numbering.id}'][data-num-level='${style.numbering.level}']`;
    }
    
    styleMap.push(mapping);
  });
  
  // Enhanced character style mapping
  Object.entries(styleInfo.styles.character || {}).forEach(([id, style]) => {
    styleMap.push(
      `r[style-name='${style.name}'] => span.docx-c-${id.toLowerCase()}`
    );
  });
  
  // Enhanced table style mapping
  Object.entries(styleInfo.styles.table || {}).forEach(([id, style]) => {
    styleMap.push(
      `table[style-name='${style.name}'] => table.docx-t-${id.toLowerCase()}`
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
  
  return styleMap;
}

/**
 * Create enhanced document transformer with numbering preservation
 * @param {Object} styleInfo - Style information with numbering context
 * @returns {Function} - Enhanced document transformer function
 */
function createEnhancedDocumentTransformer(styleInfo) {
  return function (document) {
    // Enhanced document transformation that preserves numbering context
    // This function can walk the document tree and add numbering attributes
    // based on the numbering context extracted from the DOCX
    
    if (styleInfo.numberingContext && styleInfo.numberingContext.length > 0) {
      // Add data attributes to preserve numbering context for later processing
      document.paragraphs.forEach((paragraph, index) => {
        // Find corresponding numbering context
        const context = styleInfo.numberingContext.find(ctx => 
          ctx.paragraphIndex === index
        );
        
        if (context && context.numberingId) {
          // This would require extending mammoth.js to support custom attributes
          // For now, we'll handle this in the post-processing step
        }
      });
    }
    
    return document;
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
    // Default options
    const defaultOptions = {
      trackChangesMode: 'show', // 'show', 'hide', 'accept', or 'reject'
      enhanceAccessibility: true, // Apply WCAG 2.1 Level AA enhancements
      preserveMetadata: true, // Add metadata to HTML
      showAuthor: true, // Show author in track changes
      showDate: true // Show date in track changes
    };
    
    const opts = { ...defaultOptions, ...options };
    
    // Create a DOM to manipulate the HTML
    const dom = new JSDOM(html);
    const document = dom.window.document;

    // Ensure we have a proper HTML structure
    ensureHtmlStructure(document);

    // Add link to the external CSS file
    const linkElement = document.createElement("link");
    linkElement.rel = "stylesheet";
    linkElement.href = `./${cssFilename}`; // Use relative path to the CSS file
    document.head.appendChild(linkElement);

    // Add metadata if requested
    if (opts.preserveMetadata && metadata) {
      const { applyMetadataToHtml } = require('../parsers/metadata-parser');
      applyMetadataToHtml(document, metadata);
    }
    
    // Add body class for RTL if needed
    if (styleInfo.settings?.rtlGutter) {
      document.body.classList.add("docx-rtl");
    }

    // Process table elements to match word styling better
    try {
      processTables(document);
    } catch (error) {
      console.error("Error processing tables:", error.message);
    }

    // Process images to maintain aspect ratio and positioning
    try {
      processImages(document);
    } catch (error) {
      console.error("Error processing images:", error.message);
    }

    // Handle language-specific elements
    try {
      processLanguageElements(document);
    } catch (error) {
      console.error("Error processing language elements:", error.message);
    }

    // ENHANCED: Process heading structure with DOCX-derived numbering context
    try {
      console.log('Processing headings with DOCX numbering context...');
      processHeadings(document, styleInfo, styleInfo.numberingContext);
    } catch (error) {
      console.error("Error processing headings:", error.message);
    }

    // Process the Table of Contents with better structure and styling
    try {
      processTOC(document, styleInfo);
    } catch (error) {
      console.error("Error processing TOC:", error.message);
    }

    // ENHANCED: Process numbered paragraphs with DOCX-derived numbering context
    try {
      console.log('Processing numbered paragraphs with DOCX introspection...');
      processNestedNumberedParagraphs(document, styleInfo, styleInfo.numberingContext);
    } catch (error) {
      console.error("Error processing numbered paragraphs:", error.message);
    }

    // Process special paragraph types based on document analysis
    try {
      processSpecialParagraphs(document, styleInfo);
    } catch (error) {
      console.error("Error processing special paragraphs:", error.message);
    }
    
    // Process track changes if present
    if (trackChanges && trackChanges.hasTrackedChanges) {
      try {
        const { processTrackChanges } = require('../parsers/track-changes-parser');
        processTrackChanges(document, trackChanges, {
          mode: opts.trackChangesMode,
          showAuthor: opts.showAuthor,
          showDate: opts.showDate
        });
      } catch (error) {
        console.error("Error processing track changes:", error.message);
      }
    }
    
    // Enhance accessibility if requested
    if (opts.enhanceAccessibility) {
      try {
        const { processForAccessibility } = require('../accessibility/wcag-processor');
        processForAccessibility(document, styleInfo, metadata);
      } catch (error) {
        console.error("Error enhancing accessibility:", error.message);
      }
    }

    // Log numbering context information for debugging
    if (styleInfo.numberingContext && styleInfo.numberingContext.length > 0) {
      const numberedContexts = styleInfo.numberingContext.filter(ctx => ctx.resolvedNumbering);
      console.log(`Applied DOCX numbering to ${numberedContexts.length} paragraphs from ${styleInfo.numberingContext.length} total contexts`);
    }

    // Serialize back to HTML string
    return dom.serialize();
  } catch (error) {
    console.error("Error in applyStylesToHtml:", error.message);
    // Return the original HTML if there was an error
    return html;
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
    const options = {
      path: docxPath,
      encoding: 'utf8',
      convertImage: mammoth.images.imgElement(function(image) {
        return image.read("base64").then(function(imageBuffer) {
          const extension = image.contentType.split('/')[1];
          const filename = `image-${image.altText || Date.now()}.${extension}`;
          
          // Save the image to output directory
          const imagePath = path.join(outputDir, filename);
          fs.writeFileSync(imagePath, Buffer.from(imageBuffer, 'base64'));
          
          return {
            src: `./images/${filename}`,
            alt: image.altText || 'Image'
          };
        });
      })
    };
    
    // Convert to HTML with images
    await mammoth.convertToHtml(options);
    console.log('Images extracted successfully.');
  } catch (error) {
    console.error('Error extracting images:', error);
  }
}

module.exports = {
  extractAndApplyStyles,
  convertToStyledHtml,
  applyStylesToHtml,
  extractImagesFromDocx
};