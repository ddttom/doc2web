// lib/html/html-generator.js - HTML generation functions
const mammoth = require("mammoth");
const { JSDOM } = require("jsdom");
const path = require("path");

// Import required modules
const { parseDocxStyles } = require("../parsers/style-parser");
const { generateCssFromStyleInfo } = require("../css/css-generator");
const { createStyleMap, createDocumentTransformer } = require("../css/style-mapper");
const { ensureHtmlStructure, addDocumentMetadata } = require("./structure-processor");
const { processHeadings, processTOC, processNestedNumberedParagraphs, processSpecialParagraphs } = require("./content-processors");
const { processTables, processImages, processLanguageElements } = require("./element-processors");

/**
 * Extract and apply styles from a DOCX document to HTML
 * Main entry point for converting DOCX to styled HTML
 * 
 * @param {string} docxPath - Path to the DOCX file
 * @param {string} cssFilename - Filename for the CSS file (without path)
 * @returns {Promise<{html: string, styles: string}>} - HTML with embedded styles
 */
async function extractAndApplyStyles(docxPath, cssFilename = null) {
  try {
    // First, extract the raw styles from the document
    const styleInfo = await parseDocxStyles(docxPath);

    // Generate CSS from the extracted styles
    const css = generateCssFromStyleInfo(styleInfo);

    // Convert DOCX to HTML with style preservation
    const htmlResult = await convertToStyledHtml(docxPath, styleInfo);

    // Get the CSS filename
    const cssFile =
      cssFilename || path.basename(docxPath, path.extname(docxPath)) + ".css";

    // Combine HTML and CSS
    const styledHtml = applyStylesToHtml(
      htmlResult.value,
      css,
      styleInfo,
      cssFile
    );

    return {
      html: styledHtml,
      styles: css,
      messages: htmlResult.messages,
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
 * @param {Object} styleInfo - Style information for reference
 * @param {string} cssFilename - Filename for the CSS file
 * @returns {string} - HTML with link to external CSS
 */
function applyStylesToHtml(html, css, styleInfo, cssFilename) {
  try {
    // Create a DOM to manipulate the HTML
    const dom = new JSDOM(html);
    const document = dom.window.document;

    // Ensure we have a proper HTML structure
    ensureHtmlStructure(document);

    // Instead of embedding CSS, add a link to the external CSS file
    const linkElement = document.createElement("link");
    linkElement.rel = "stylesheet";
    linkElement.href = `./${cssFilename}`; // Use relative path to the CSS file

    // Add to document head
    document.head.appendChild(linkElement);

    // Add metadata
    addDocumentMetadata(document, styleInfo);

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

    // Serialize back to HTML string
    return dom.serialize();
  } catch (error) {
    console.error("Error in applyStylesToHtml:", error.message);
    // Return the original HTML if there was an error
    return html;
  }
}

module.exports = {
  extractAndApplyStyles,
  convertToStyledHtml,
  applyStylesToHtml
};