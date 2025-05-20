// style-extractor.js - Extract styles directly from DOCX files
const mammoth = require("mammoth");
const { JSDOM } = require("jsdom");
const path = require("path");
const {
  parseDocxStyles,
  generateCssFromStyleInfo,
} = require("./docx-style-parser");

/**
 * Extract and apply styles from a DOCX document to HTML
 * @param {string} docxPath - Path to the DOCX file
 * @param {string} cssFilename - Filename for the CSS file (without path)
 * @returns {Promise<{html: string, styles: string}>} - HTML with embedded styles
 */
async function extractAndApplyStyles(docxPath, cssFilename = null) {
  try {
    // Extract document information including styles
    const documentInfo = await parseDocxStyles(docxPath);

    // Generate CSS from the extracted document information
    const css = generateCssFromStyleInfo(documentInfo);

    // Convert DOCX to HTML
    const htmlResult = await convertToStyledHtml(docxPath, documentInfo);

    // Get the CSS filename based on the input file if not provided
    const cssFile = cssFilename || path.basename(docxPath, path.extname(docxPath)) + ".css";

    // Apply enhanced styling to HTML and link the external CSS
    const styledHtml = applyStylesToHtml(htmlResult.value, documentInfo, cssFile);

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
 * Convert DOCX to HTML while preserving structure and classes
 * @param {string} docxPath - Path to the DOCX file
 * @param {Object} documentInfo - Document information
 * @returns {Promise<{value: string, messages: Array}>} - HTML with appropriate class attributes
 */
async function convertToStyledHtml(docxPath, documentInfo) {
  // Create custom style map based on extracted styles
  const styleMap = createStyleMap(documentInfo);

  // Configure image handling
  const imageOptions = {
    convertImage: mammoth.images.imgElement(function (image) {
      return image.read("base64").then(function (imageBuffer) {
        const extension = image.contentType.split("/")[1];
        const filename = `image-${image.altText || Date.now()}.${extension}`;

        return {
          src: `./images/${filename}`,
          alt: image.altText || "Image",
          className: "doc-image",
        };
      });
    }),
  };

  // Convert using mammoth with custom style mappings
  const result = await mammoth.convertToHtml({
    path: docxPath,
    styleMap: styleMap,
    includeDefaultStyleMap: true,
    ...imageOptions,
  });

  return result;
}

/**
 * Create a custom style map for mammoth conversion
 * @param {Object} documentInfo - Document information
 * @returns {Array<string>} - Style map entries
 */
function createStyleMap(documentInfo) {
  const styleMap = [];

  // Add paragraph styles
  if (documentInfo.styles && documentInfo.styles.paragraph) {
    Object.entries(documentInfo.styles.paragraph).forEach(([id, style]) => {
      if (style.name) {
        styleMap.push(
          `p[style-name='${style.name}'] => p.doc-p-${id.replace(/[^a-zA-Z0-9]/g, '-').toLowerCase()}`
        );
      }
    });
  }

  // Add character styles
  if (documentInfo.styles && documentInfo.styles.character) {
    Object.entries(documentInfo.styles.character).forEach(([id, style]) => {
      if (style.name) {
        styleMap.push(
          `r[style-name='${style.name}'] => span.doc-c-${id.replace(/[^a-zA-Z0-9]/g, '-').toLowerCase()}`
        );
      }
    });
  }

  // Add table styles
  if (documentInfo.styles && documentInfo.styles.table) {
    Object.entries(documentInfo.styles.table).forEach(([id, style]) => {
      if (style.name) {
        styleMap.push(
          `table[style-name='${style.name}'] => table.doc-t-${id.replace(/[^a-zA-Z0-9]/g, '-').toLowerCase()}`
        );
      }
    });
  }

  // Map headings
  styleMap.push("p[style-name='heading 1'] => h1.doc-heading");
  styleMap.push("p[style-name='heading 2'] => h2.doc-heading");
  styleMap.push("p[style-name='heading 3'] => h3.doc-heading");
  styleMap.push("p[style-name='heading 4'] => h4.doc-heading");
  styleMap.push("p[style-name='heading 5'] => h5.doc-heading");
  styleMap.push("p[style-name='heading 6'] => h6.doc-heading");

  // Map standard formatting elements
  styleMap.push("p:fresh => p");
  styleMap.push("r[bold] => strong");
  styleMap.push("r[italic] => em");
  styleMap.push("r[underline] => span.doc-underline");
  styleMap.push("r[strikethrough] => span.doc-strike");
  styleMap.push("r[subscript] => sub");
  styleMap.push("r[superscript] => sup");
  
  // Map list items with appropriate attributes
  styleMap.push("p[numbering] => li.doc-list-item");
  styleMap.push("p[numbering=1] => li.doc-list-item[data-level='1']");
  styleMap.push("p[numbering=2] => li.doc-list-item[data-level='2']");
  styleMap.push("p[numbering=3] => li.doc-list-item[data-level='3']");

  return styleMap;
}

/**
 * Apply document styles to HTML and link the external CSS
 * @param {string} html - Raw HTML content
 * @param {Object} documentInfo - Document information
 * @param {string} cssFilename - Filename for external CSS
 * @returns {string} - Enhanced HTML with link to external CSS
 */
function applyStylesToHtml(html, documentInfo, cssFilename) {
  try {
    // Create a DOM to manipulate the HTML
    const dom = new JSDOM(html);
    const document = dom.window.document;

    // Ensure proper HTML structure (html, head, body)
    ensureHtmlStructure(document);

    // Add external CSS link
    const linkElement = document.createElement("link");
    linkElement.rel = "stylesheet";
    linkElement.href = `./${cssFilename}`;
    document.head.appendChild(linkElement);

    // Set document title based on the first heading or filename
    setDocumentTitle(document, cssFilename);

    // Add document language attribute if available
    if (documentInfo.hasRTL) {
      document.documentElement.setAttribute("dir", "rtl");
    }

    // Add standard document class
    document.body.classList.add("doc-content");

    // Process document structure
    enhanceDocumentStructure(document, documentInfo);

    // Return the serialized HTML
    return dom.serialize();
  } catch (error) {
    console.error("Error applying styles to HTML:", error.message);
    return html; // Return original HTML if there was an error
  }
}

/**
 * Ensure HTML has proper structure (html, head, body)
 * @param {Document} document - DOM document
 */
function ensureHtmlStructure(document) {
  // Create html element if not exists
  if (!document.documentElement || document.documentElement.nodeName.toLowerCase() !== "html") {
    const html = document.createElement("html");
    while (document.childNodes.length > 0) {
      html.appendChild(document.childNodes[0]);
    }
    document.appendChild(html);
  }

  // Create head if not exists
  if (!document.head) {
    const head = document.createElement("head");
    document.documentElement.insertBefore(head, document.documentElement.firstChild);
  }

  // Add meta charset
  if (!document.querySelector("meta[charset]")) {
    const meta = document.createElement("meta");
    meta.setAttribute("charset", "utf-8");
    document.head.appendChild(meta);
  }

  // Add viewport meta
  if (!document.querySelector("meta[name='viewport']")) {
    const viewport = document.createElement("meta");
    viewport.setAttribute("name", "viewport");
    viewport.setAttribute("content", "width=device-width, initial-scale=1");
    document.head.appendChild(viewport);
  }

  // Create body if not exists
  if (!document.body) {
    const body = document.createElement("body");
    Array.from(document.documentElement.childNodes).forEach((node) => {
      if (node !== document.head && node.nodeName.toLowerCase() !== "body") {
        body.appendChild(node);
      }
    });
    document.documentElement.appendChild(body);
  }
}

/**
 * Set document title based on content
 * @param {Document} document - DOM document
 * @param {string} cssFilename - CSS filename to extract title
 */
function setDocumentTitle(document, cssFilename) {
  if (!document.title) {
    let title = "";
    
    // Try to get title from first heading
    const firstHeading = document.querySelector("h1, h2, h3");
    if (firstHeading) {
      title = firstHeading.textContent.trim();
    } else {
      // Use filename as fallback
      title = cssFilename.replace(/\.(css|CSS)$/, '');
    }
    
    const titleElement = document.createElement("title");
    titleElement.textContent = title;
    document.head.appendChild(titleElement);
  }
}

/**
 * Enhance document structure with semantic improvements
 * @param {Document} document - DOM document
 * @param {Object} documentInfo - Document information
 */
function enhanceDocumentStructure(document, documentInfo) {
  // Process tables
  enhanceTables(document);
  
  // Process lists
  enhanceLists(document, documentInfo);
  
  // Process images
  enhanceImages(document);
  
  // Process language attributes
  processLanguageAttributes(document);
}

/**
 * Enhance tables with responsive container and consistent styling
 * @param {Document} document - DOM document
 */
function enhanceTables(document) {
  const tables = document.querySelectorAll("table");
  tables.forEach((table) => {
    // Add standard class if not already present
    if (!table.className.includes("doc-t-")) {
      table.classList.add("doc-table");
    }
    
    // Ensure tbody exists
    if (!table.querySelector("tbody")) {
      const tbody = document.createElement("tbody");
      Array.from(table.querySelectorAll("tr")).forEach((row) => {
        tbody.appendChild(row);
      });
      table.appendChild(tbody);
    }
    
    // Add responsive container
    const wrapper = document.createElement("div");
    wrapper.className = "doc-table-container";
    table.parentNode.insertBefore(wrapper, table);
    wrapper.appendChild(table);
  });
}

/**
 * Enhance lists with proper structure and class attributes
 * @param {Document} document - DOM document
 * @param {Object} documentInfo - Document information
 */
function enhanceLists(document, documentInfo) {
  // Find all list items
  const listItems = document.querySelectorAll("li");
  
  // Group list items into appropriate lists
  const listGroups = {};
  
  listItems.forEach((item) => {
    // Skip items already in a list
    if (item.parentNode.tagName === "UL" || item.parentNode.tagName === "OL") {
      return;
    }
    
    const level = item.getAttribute("data-level") || "1";
    const numId = item.getAttribute("data-numid") || "default";
    const key = `${numId}-${level}`;
    
    if (!listGroups[key]) {
      listGroups[key] = [];
    }
    
    listGroups[key].push(item);
  });
  
  // Create proper lists from groups
  Object.entries(listGroups).forEach(([key, items]) => {
    if (items.length === 0) return;
    
    const [numId, level] = key.split("-");
    const firstItem = items[0];
    const parent = firstItem.parentNode;
    
    // Create list element
    const list = document.createElement("ol");
    list.className = `doc-list doc-list-${numId}`;
    
    // Add level-specific class
    list.classList.add(`doc-list-${numId}-${level}`);
    
    // Insert list before first item
    parent.insertBefore(list, firstItem);
    
    // Move items to list
    items.forEach((item) => {
      list.appendChild(item);
    });
  });
  
  // Process list items for consistent formatting
  document.querySelectorAll(".doc-list-item").forEach((item, index) => {
    // Add index number as data attribute if not present
    if (!item.hasAttribute("data-index")) {
      item.setAttribute("data-index", (index + 1).toString());
    }
    
    // Check for list label in content
    const text = item.textContent.trim();
    const prefixMatch = text.match(/^([0-9]+|[a-z]|[A-Z]|[ivxIVX]+)[\.\)]\s+(.+)$/);
    
    if (prefixMatch) {
      // Add prefix as attribute and remove from content
      item.setAttribute("data-prefix", prefixMatch[1]);
      
      // Remove prefix from list item text if it's a direct text node
      // This should be done carefully to avoid removing from nested elements
      Array.from(item.childNodes).forEach(node => {
        if (node.nodeType === 3) { // Text node
          if (node.textContent.match(/^([0-9]+|[a-z]|[A-Z]|[ivxIVX]+)[\.\)]\s+/)) {
            node.textContent = node.textContent.replace(/^([0-9]+|[a-z]|[A-Z]|[ivxIVX]+)[\.\)]\s+/, '');
          }
        }
      });
    }
  });
}

/**
 * Enhance images with figure elements and proper sizing
 * @param {Document} document - DOM document
 */
function enhanceImages(document) {
  const images = document.querySelectorAll("img");
  
  images.forEach((img) => {
    // Add standard class if not present
    if (!img.classList.contains("doc-image")) {
      img.classList.add("doc-image");
    }
    
    // Ensure all images have alt text
    if (!img.hasAttribute("alt")) {
      img.setAttribute("alt", "Document image");
    }
    
    // Add figure wrapper if not already wrapped
    if (img.parentNode.tagName !== "FIGURE") {
      const figure = document.createElement("figure");
      figure.className = "doc-figure";
      img.parentNode.insertBefore(figure, img);
      figure.appendChild(img);
    }
  });
}

/**
 * Process language attributes in the document
 * @param {Document} document - DOM document
 */
function processLanguageAttributes(document) {
  // Find all elements with language attributes
  const langElements = document.querySelectorAll("[lang]");
  
  langElements.forEach((el) => {
    const lang = el.getAttribute("lang");
    
    // Add language-specific class
    el.classList.add(`doc-lang-${lang}`);
    
    // Add RTL direction for RTL languages
    if (["ar", "he", "fa", "ur"].includes(lang)) {
      el.setAttribute("dir", "rtl");
      el.classList.add("doc-rtl");
    }
  });
}

module.exports = {
  extractAndApplyStyles
};
