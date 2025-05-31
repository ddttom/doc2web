// File: lib/html/generators/html-processing.js
// HTML processing and content manipulation

const { JSDOM } = require("jsdom");
const { ensureHtmlStructure } = require("../structure-processor");
const {
  processHeadings,
  processTOC,
  processNestedNumberedParagraphs,
  processSpecialParagraphs,
} = require("../content-processors");
const {
  processTables,
  processImages,
  processLanguageElements,
} = require("../element-processors");
const {
  parseDocumentMetadata,
  applyMetadataToHtml,
} = require("../../parsers/metadata-parser");
const {
  parseTrackChanges,
  processTrackChanges: applyTrackChangesToHtml,
} = require("../../parsers/track-changes-parser");
const {
  extractDocumentHeader,
  processHeaderForHtml,
} = require("../../parsers/header-parser");

/**
 * Apply styles and process HTML content
 */
async function applyStylesAndProcessHtml(
  html,
  cssString,
  styleInfo,
  cssFilename,
  metadata,
  trackChanges,
  headerInfo,
  options,
  zip = null,
  outputDir = null
) {
  try {
    console.log(`Applying styles. Initial HTML length: ${html.length}`);
    
    // Create DOM with content validation
    if (!html || html.trim().length < 10) {
      console.warn("HTML input is very short or empty. Using fallback.");
      html = `<!DOCTYPE html><html><head><meta charset="utf-8"></head><body>Empty document</body></html>`;
    }
    
    const dom = new JSDOM(html, {
      contentType: "text/html",
      includeNodeLocations: true,
    });
    
    const document = dom.window.document;
    
    // Validate document structure
    if (!document || !document.body) {
      throw new Error("Failed to create DOM from HTML input.");
    }
    
    // Check content before processing
    const initialBodyContent = document.body.innerHTML;
    const initialBodyLength = initialBodyContent.length;
    console.log(`Initial body content length: ${initialBodyLength}`);
    
    // Ensure proper HTML structure
    ensureHtmlStructure(document);
    
    // Add CSS link
    const link = document.createElement("link");
    link.setAttribute("rel", "stylesheet");
    link.setAttribute("href", `./${cssFilename}`);
    document.head.appendChild(link);
    
    // Apply metadata if enabled
    if (options.preserveMetadata && metadata) {
      applyMetadataToHtml(document, metadata);
    }
    
    // Add RTL class if needed
    if (styleInfo.settings?.rtlGutter) {
      document.body.classList.add("docx-rtl");
    }
    
    // Apply abstract numbering IDs to elements with numbering
    const elementsToNumber = document.querySelectorAll("p, h1, h2, h3, h4, h5, h6, li");
    elementsToNumber.forEach((el) => {
      const numId = el.getAttribute("data-num-id");
      if (numId && styleInfo.numberingDefs?.nums[numId]) {
        el.setAttribute(
          "data-abstract-num",
          styleInfo.numberingDefs.nums[numId].abstractNumId
        );
      }
    });
    
    // Process document structure
    processHeadings(document, styleInfo, styleInfo.numberingContext);
    
    // Insert header before TOC if header content was found
    if (headerInfo && headerInfo.hasHeaderContent) {
      await insertHeaderBeforeTOC(document, headerInfo, styleInfo, zip, outputDir);
    }
    
    processTOC(document, styleInfo, styleInfo.numberingContext);
    processNestedNumberedParagraphs(document, styleInfo, styleInfo.numberingContext);
    processSpecialParagraphs(document, styleInfo);
    processTables(document);
    processImages(document);
    processLanguageElements(document);
    
    // Apply track changes if enabled
    if (
      trackChanges &&
      trackChanges.hasTrackedChanges &&
      options.trackChangesMode !== "hide"
    ) {
      applyTrackChangesToHtml(document, trackChanges, {
        mode: options.trackChangesMode,
        showAuthor: options.showAuthor,
        showDate: options.showDate,
      });
    }
    
    // Apply accessibility enhancements if enabled
    if (options.enhanceAccessibility) {
      const { processForAccessibility } = require("../../accessibility/wcag-processor");
      processForAccessibility(document, styleInfo, metadata);
    }
    
    // Validate content after processing
    const finalBodyContent = document.body.innerHTML;
    const finalBodyLength = finalBodyContent.length;
    console.log(`Final body content length: ${finalBodyLength}`);
    
    // Check for significant content loss
    if (finalBodyLength < initialBodyLength * 0.8 && initialBodyLength > 100) {
      console.warn("Significant content loss detected during processing. Using content preservation fallback.");
      // Implement fallback strategy
      const bodyWithStructure = document.createElement('div');
      bodyWithStructure.innerHTML = initialBodyContent;
      
      // Add structure with minimal DOM manipulation
      const bodyElements = Array.from(bodyWithStructure.childNodes);
      document.body.innerHTML = '';
      bodyElements.forEach(node => {
        document.body.appendChild(node.cloneNode(true));
      });
      
      // Re-apply minimal styling
      document.querySelectorAll('p, h1, h2, h3, h4, h5, h6').forEach(el => {
        const text = el.textContent.trim();
        // Apply minimal context-specific styling
        if (el.tagName.toLowerCase().startsWith('h')) {
          el.classList.add(`docx-heading${el.tagName.slice(1)}`);
        }
      });
    }
    
    // Serialize and return HTML
    return dom.serialize();
  } catch (error) {
    console.error(
      "Error in applyStylesAndProcessHtml:",
      error.message,
      error.stack
    );
    
    // Create fallback HTML with error message
    return `<!DOCTYPE html>
      <html>
        <head>
          <meta charset="utf-8">
          <title>Error</title>
          <link rel="stylesheet" href="./${cssFilename}">
        </head>
        <body>
          <p>Error during HTML processing: ${error.message}</p>
          <pre>${html.substring(0, 1000)}...</pre>
        </body>
      </html>`;
  }
}

/**
 * Insert header content before the TOC in the HTML document
 * 
 * @param {Document} document - HTML document
 * @param {Object} headerInfo - Header information from extraction
 * @param {Object} styleInfo - Style information
 * @param {Object} zip - JSZip instance for image extraction (optional)
 * @param {string} outputDir - Output directory for images (optional)
 */
async function insertHeaderBeforeTOC(document, headerInfo, styleInfo, zip = null, outputDir = null) {
  try {
    console.log(`Inserting ${headerInfo.headerParagraphs.length} header paragraphs before TOC`);
    
    // Create header container
    const headerContainer = document.createElement('header');
    headerContainer.className = 'docx-document-header';
    headerContainer.setAttribute('role', 'banner');
    
    // Process header paragraphs into HTML elements
    const headerElements = await processHeaderForHtml(headerInfo, document, styleInfo, zip, outputDir);
    
    // Add header elements to container
    headerElements.forEach(headerElement => {
      headerContainer.appendChild(headerElement);
    });
    
    // Insert the header container at the very beginning of the body
    // This ensures it appears before any other content including TOC
    if (document.body.firstElementChild) {
      document.body.insertBefore(headerContainer, document.body.firstElementChild);
    } else {
      document.body.appendChild(headerContainer);
    }
    
    console.log('Header content inserted successfully');
    
  } catch (error) {
    console.error('Error inserting header before TOC:', error);
  }
}

module.exports = {
  applyStylesAndProcessHtml,
  insertHeaderBeforeTOC,
};
