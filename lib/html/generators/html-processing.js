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
    
    // Remove page numbers from TOC entries after TOC processing
    const { removeTOCPageNumbers } = require("../processors/toc-processor");
    removeTOCPageNumbers(document);
    
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

    // Post-process bullet lists that weren't converted by mammoth.js
    postProcessBulletLists(document, styleInfo);
    
    // Process existing li elements that mammoth.js converted but need bullet styling
    processExistingBulletElements(document, styleInfo);
    
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
    
    // Preserve italic formatting validation
    preserveItalicFormatting(document);
    
    // FINAL STEP: Add TOC linking functionality after all HTML processing is complete
    const { linkTOCEntries } = require("../processors/toc-linker");
    linkTOCEntries(document, styleInfo);
    
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

/**
 * Post-process paragraphs that should be bullet lists
 */
function postProcessBulletLists(document, styleInfo) {
  try {
    // Find paragraphs that look like bullet points but weren't converted
    const bulletParagraphs = document.querySelectorAll('p[data-format="bullet"], p.docx-bullet-item, p[class*="docx-bullet-"]');
    
    if (bulletParagraphs.length === 0) return;
    
    console.log(`Post-processing ${bulletParagraphs.length} bullet paragraphs`);
    
    // Group consecutive bullet paragraphs into lists
    let currentList = null;
    let currentLevel = -1;
    
    bulletParagraphs.forEach(p => {
      const level = parseInt(p.getAttribute('data-num-level') || '0');
      
      if (level === 0 || !currentList || !p.previousElementSibling || 
          !p.previousElementSibling.classList.contains('docx-bullet-item')) {
        // Create new list
        currentList = document.createElement('ul');
        currentList.className = 'docx-bullet-list';
        p.parentNode.insertBefore(currentList, p);
        currentLevel = level;
      }
      
      // Convert paragraph to list item
      const li = document.createElement('li');
      
      // Clean up class names and ensure proper bullet styling
      let cleanClassName = p.className.replace(/docx-p-/g, 'docx-bullet-').trim();
      if (!cleanClassName.includes('docx-bullet-') && !cleanClassName.includes('docx-list-item')) {
        cleanClassName = cleanClassName ? `${cleanClassName} docx-list-item` : 'docx-list-item';
      }
      li.className = cleanClassName;
      li.innerHTML = p.innerHTML;
      
      // Copy data attributes
      Array.from(p.attributes).forEach(attr => {
        if (attr.name.startsWith('data-')) {
          li.setAttribute(attr.name, attr.value);
        }
      });
      
      currentList.appendChild(li);
      p.remove();
    });
    
    console.log('Bullet list post-processing completed');
    
  } catch (error) {
    console.error('Error in postProcessBulletLists:', error.message);
  }
}

/**
 * Process existing li elements that mammoth.js converted but need bullet styling
 */
function processExistingBulletElements(document, styleInfo) {
  try {
    // Find all li elements with bullet data attributes that don't have proper classes
    const bulletElements = document.querySelectorAll('li[data-format="bullet"]');
    
    if (bulletElements.length === 0) return;
    
    console.log(`Processing ${bulletElements.length} existing bullet elements`);
    
    bulletElements.forEach(li => {
      // Add bullet styling class if not already present
      if (!li.classList.contains('docx-list-item') && !li.className.includes('docx-bullet-')) {
        li.classList.add('docx-list-item');
      }
      
      // Ensure parent ul has proper bullet list class
      const parentUl = li.closest('ul');
      if (parentUl && !parentUl.classList.contains('docx-bullet-list')) {
        parentUl.classList.add('docx-bullet-list');
      }
      
      // Add proper styling attributes for consistency
      if (!li.style.position) {
        li.style.position = 'relative';
      }
    });
    
    console.log('Existing bullet elements processing completed');
    
  } catch (error) {
    console.error('Error in processExistingBulletElements:', error.message);
  }
}

/**
 * Preserve italic formatting during HTML processing
 * Ensures em tags and italic styles are not stripped
 * 
 * @param {Document} document - HTML document
 */
function preserveItalicFormatting(document) {
  try {
    // Find all em tags and ensure they have proper styling
    const emTags = document.querySelectorAll('em');
    console.log(`Found ${emTags.length} italic elements in HTML`);
    
    // Validate italic elements have proper styling
    emTags.forEach(em => {
      if (!em.style.fontStyle && !em.className.includes('italic')) {
        em.style.fontStyle = 'italic';
      }
    });
    
    // Find elements with italic classes and ensure they're preserved
    const italicElements = document.querySelectorAll('.italic, [class*="italic"]');
    italicElements.forEach(el => {
      if (!el.style.fontStyle) {
        el.style.fontStyle = 'italic';
      }
    });
    
    // Find elements with inline italic styles and preserve them
    const inlineItalicElements = document.querySelectorAll('[style*="font-style: italic"]');
    console.log(`Found ${inlineItalicElements.length} elements with inline italic styles`);
    
  } catch (error) {
    console.error('Error preserving italic formatting:', error);
  }
}

/**
 * Flatten nested paragraph elements to sibling elements
 * Preserves all existing attributes and content exactly as-is
 * @param {Document} document - HTML document
 */
function flattenNestedParagraphs(document) {
  const nestedParagraphs = document.querySelectorAll('p p');
  
  if (nestedParagraphs.length === 0) {
    return; // No nested paragraphs found
  }
  
  console.log(`Flattening ${nestedParagraphs.length} nested paragraph elements`);
  
  // Process all root paragraphs (paragraphs not inside other paragraphs)
  const rootParagraphs = Array.from(document.querySelectorAll('p')).filter(p => !p.closest('p'));
  
  rootParagraphs.forEach(rootP => {
    extractNestedParagraphs(rootP);
  });
}

/**
 * Extract nested paragraphs from a root paragraph and insert as siblings
 * @param {Element} paragraph - Root paragraph element
 */
function extractNestedParagraphs(paragraph) {
  const nestedParagraphs = Array.from(paragraph.querySelectorAll(':scope > p'));
  
  if (nestedParagraphs.length === 0) {
    return; // No nested paragraphs
  }
  
  const parentElement = paragraph.parentElement;
  let insertionPoint = paragraph.nextSibling;
  
  nestedParagraphs.forEach(nestedP => {
    // Remove from current location
    nestedP.remove();
    
    // Recursively process any deeper nesting
    extractNestedParagraphs(nestedP);
    
    // Insert as sibling after current paragraph
    if (insertionPoint) {
      parentElement.insertBefore(nestedP, insertionPoint);
    } else {
      parentElement.appendChild(nestedP);
    }
    
    // Update insertion point for next paragraph
    insertionPoint = nestedP.nextSibling;
  });
}

module.exports = {
  applyStylesAndProcessHtml,
  insertHeaderBeforeTOC,
  postProcessBulletLists,
  processExistingBulletElements,
  preserveItalicFormatting,
  flattenNestedParagraphs,
  extractNestedParagraphs,
};
