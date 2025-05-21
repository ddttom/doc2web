// lib/html/html-generator.js - HTML generation functions
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
 * Extract and apply styles from a DOCX document to HTML
 * Main entry point for converting DOCX to styled HTML
 * 
 * @param {string} docxPath - Path to the DOCX file
 * @param {string} cssFilename - Filename for the CSS file (without path)
 * @param {Object} options - Processing options
 * @returns {Promise<{html: string, styles: string}>} - HTML with embedded styles
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
    
    // Extract the raw styles from the document
    const styleInfo = await parseDocxStyles(docxPath);

    // Extract document metadata
    const metadata = parseDocumentMetadata(corePropsDoc, appPropsDoc);
    
    // Extract track changes information
    const trackChanges = parseTrackChanges(documentDoc);

    // Generate CSS from the extracted styles
    const css = generateCssFromStyleInfo(styleInfo);

    // Convert DOCX to HTML with style preservation
    const htmlResult = await convertToStyledHtml(docxPath, styleInfo);

    // Get the CSS filename
    const cssFile =
      cssFilename || path.basename(docxPath, path.extname(docxPath)) + ".css";

    // Combine HTML and CSS with enhanced features
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
      trackChanges
    };
  } catch (error) {
    console.error("Error extracting styles:", error);
    throw error;
  }
}

/**
 * Convert DOCX to HTML while preserving style information
 * Uses mammoth.js with custom style mapping and transformations
 * 
 * @param {string} docxPath - Path to the DOCX file
 * @param {Object} styleInfo - Extracted style information
 * @returns {Promise<{value: string, messages: Array}>} - HTML with style information
 */
async function convertToStyledHtml(docxPath, styleInfo) {
  // Create a custom style map based on extracted styles
  const styleMap = createStyleMap(styleInfo);

  // Custom document transformer to enhance style preservation
  const transformDocument = createDocumentTransformer(styleInfo);

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

  // Use mammoth to convert with the custom style map
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
 * Apply generated CSS to HTML
 * Enhances the HTML with proper structure and styling
 * 
 * @param {string} html - HTML content
 * @param {string} css - CSS styles
 * @param {Object} styleInfo - Style information
 * @param {string} cssFilename - Filename for the CSS file
 * @param {Object} metadata - Document metadata
 * @param {Object} trackChanges - Track changes information
 * @param {Object} options - Processing options
 * @returns {string} - HTML with link to external CSS
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

    // Process heading structure to match TOC
    try {
      processHeadings(document, styleInfo);
    } catch (error) {
      console.error("Error processing headings:", error.message);
    }

    // Process the Table of Contents with better structure and styling
    try {
      processTOC(document, styleInfo);
    } catch (error) {
      console.error("Error processing TOC:", error.message);
    }

    // Process numbered paragraphs with proper nesting for hierarchical lists
    try {
      processNestedNumberedParagraphs(document, styleInfo);
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
