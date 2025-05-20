// style-extractor.js - Improved version
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
      styles: css + generateAdditionalCss(),
      messages: htmlResult.messages,
    };
  } catch (error) {
    console.error("Error extracting styles:", error);
    throw error;
  }
}

/**
 * Convert DOCX to HTML while preserving style information
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
 * Create a custom style map based on extracted styles
 * @param {Object} styleInfo - Extracted style information
 * @returns {Array<string>} - Style map entries
 */
function createStyleMap(styleInfo) {
  const styleMap = [];

  // Map paragraph styles
  Object.entries(styleInfo.styles.paragraph || {}).forEach(([id, style]) => {
    styleMap.push(
      `p[style-name='${style.name}'] => p.docx-p-${id.toLowerCase()}`
    );
  });

  // Map character styles
  Object.entries(styleInfo.styles.character || {}).forEach(([id, style]) => {
    styleMap.push(
      `r[style-name='${style.name}'] => span.docx-c-${id.toLowerCase()}`
    );
  });

  // Map table styles
  Object.entries(styleInfo.styles.table || {}).forEach(([id, style]) => {
    styleMap.push(
      `table[style-name='${style.name}'] => table.docx-t-${id.toLowerCase()}`
    );
  });

  // Improved heading mappings
  styleMap.push("p[style-name='heading 1'] => h1.docx-heading1");
  styleMap.push("p[style-name='heading 2'] => h2.docx-heading2");
  styleMap.push("p[style-name='heading 3'] => h3.docx-heading3");
  styleMap.push("p[style-name='heading 4'] => h4.docx-heading4");
  styleMap.push("p[style-name='heading 5'] => h5.docx-heading5");
  styleMap.push("p[style-name='heading 6'] => h6.docx-heading6");

  // Additional custom mappings for specific elements
  styleMap.push("p:fresh => p");
  styleMap.push("r[bold] => strong");
  styleMap.push("r[italic] => em");
  styleMap.push("r[underline] => span.docx-underline");
  styleMap.push("r[strikethrough] => span.docx-strike");
  styleMap.push("r[subscript] => sub");
  styleMap.push("r[superscript] => sup");

  // Handle specific paragraph types
  styleMap.push("p[style-name='Normal Web'] => p.docx-normalweb"); 
  styleMap.push("p[style-name='Body Text'] => p.docx-bodytext");
  styleMap.push("p[style-name='Rationale'] => p.docx-rationale");
  
  return styleMap;
}

/**
 * Create document transformer function for style preservation
 * @param {Object} styleInfo - Style information
 * @returns {Function} - Document transformer
 */
function createDocumentTransformer(styleInfo) {
  return function (document) {
    // In a complete implementation, this would walk the document tree
    // and enhance elements with style attributes based on the styleInfo

    // Process paragraphs, runs, tables and other elements
    // Add class names and style attributes as needed

    return document;
  };
}

/**
 * Apply generated CSS to HTML
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
      processHeadings(document);
    } catch (error) {
      console.error("Error processing headings:", error.message);
    }

    // Process the Table of Contents specifically
    try {
      processTOC(document);
    } catch (error) {
      console.error("Error processing TOC:", error.message);
    }

    // Process numbered paragraphs with proper nesting for hierarchical lists
    try {
      processNestedNumberedParagraphs(document);
    } catch (error) {
      console.error("Error processing numbered paragraphs:", error.message);
    }

    // Improve rationale paragraphs
    try {
      processRationales(document);
    } catch (error) {
      console.error("Error processing rationales:", error.message);
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
 * Process heading elements to match TOC structure
 * @param {Document} document - DOM document
 */
function processHeadings(document) {
  // Get all headings
  const headings = document.querySelectorAll('h1, h2, h3, h4, h5, h6');
  
  // Current section counters
  const counters = {
    h1: 0,
    h2: 0,
    h3: 0,
    h4: 0,
    h5: 0,
    h6: 0
  };
  
  // Track current level to reset lower levels
  let currentLevel = 0;
  
  for (let i = 0; i < headings.length; i++) {
    const heading = headings[i];
    const level = parseInt(heading.tagName.substring(1), 10);
    
    // Reset all counters below the current level
    if (level <= currentLevel) {
      for (let j = level + 1; j <= 6; j++) {
        counters[`h${j}`] = 0;
      }
    }
    
    // Increment counter for this level
    counters[`h${level}`]++;
    currentLevel = level;
    
    // Create numbered prefix
    let prefix = '';
    if (level === 1) {
      prefix = `${counters.h1}.`;
    } else if (level === 2) {
      if (heading.textContent.trim().toLowerCase().startsWith('rationale')) {
        // Skip numbering for rationale headings
        continue;
      }
      prefix = `${counters.h1}.${String.fromCharCode(96 + counters.h2)}.`;
    } else if (level === 3) {
      prefix = `${counters.h1}.${counters.h2}.${counters.h3}.`;
    }
    
    // Add number before heading text if not already there
    if (!heading.textContent.startsWith(prefix)) {
      heading.setAttribute('data-number', prefix);
      
      // Create a wrapper for the number
      const numberSpan = document.createElement('span');
      numberSpan.className = 'heading-number';
      numberSpan.textContent = prefix + ' ';
      
      // Insert the number span at the beginning of the heading
      if (heading.firstChild) {
        heading.insertBefore(numberSpan, heading.firstChild);
      } else {
        heading.appendChild(numberSpan);
      }
    }
  }
}

/**
 * Process Table of Contents for proper structure
 * @param {Document} document - DOM document
 */
function processTOC(document) {
  // Find the TOC (usually the first ordered list with specific structure)
  const tocLists = document.querySelectorAll('ol.docx-numbered-list, ol.docx-alpha-list');
  
  if (tocLists.length === 0) return;
  
  // Assume the first such list is the TOC
  const tocMainList = tocLists[0];
  
  // Add specific TOC classes
  tocMainList.classList.add('toc-container');
  
  // Process all list items in the TOC
  const tocItems = tocMainList.querySelectorAll('li');
  tocItems.forEach((item) => {
    // Find or create the components of the TOC entry
    const textSpan = item.querySelector('.docx-toc-text');
    const dotsSpan = item.querySelector('.docx-toc-dots');
    const pageSpan = item.querySelector('.docx-toc-pagenum');
    
    if (textSpan && pageSpan) {
      // Create a cleaner structure for the TOC entry
      item.classList.add('toc-entry');
      
      // Ensure the dots connect properly
      if (dotsSpan) {
        dotsSpan.classList.add('toc-dots');
      }
    }
  });
  
  // Add a heading for the TOC if not already present
  const prevElement = tocMainList.previousElementSibling;
  if (!prevElement || !prevElement.matches('h1, h2, h3, h4, h5, h6')) {
    const tocHeading = document.createElement('h2');
    tocHeading.textContent = "Table of Contents";
    tocHeading.classList.add('toc-heading');
    tocMainList.parentNode.insertBefore(tocHeading, tocMainList);
  }
}

/**
 * Process rationale paragraphs for better styling
 * @param {Document} document - DOM document
 */
function processRationales(document) {
  // Find paragraphs that appear to be rationales
  const rationaleParagraphs = document.querySelectorAll('p');
  
  rationaleParagraphs.forEach(para => {
    const text = para.textContent.trim();
    if (text.startsWith('Rationale for Resolution') || 
        text.startsWith('Rationale for Resolutions')) {
      para.classList.add('docx-rationale');
      
      // Italicize the content if not already
      if (!para.querySelector('em')) {
        const em = document.createElement('em');
        // Move all child nodes to the em element
        while (para.firstChild) {
          em.appendChild(para.firstChild);
        }
        para.appendChild(em);
      }
    }
  });
}

/**
 * Ensure HTML has proper structure (html, head, body)
 * @param {Document} document - DOM document
 */
function ensureHtmlStructure(document) {
  // Create html element if not exists
  if (
    !document.documentElement ||
    document.documentElement.nodeName.toLowerCase() !== "html"
  ) {
    const html = document.createElement("html");

    // Move existing content
    while (document.childNodes.length > 0) {
      html.appendChild(document.childNodes[0]);
    }

    document.appendChild(html);
  }

  // Create head if not exists
  if (!document.head) {
    const head = document.createElement("head");
    document.documentElement.insertBefore(
      head,
      document.documentElement.firstChild
    );
  }

  // Add meta charset
  if (!document.querySelector("meta[charset]")) {
    const meta = document.createElement("meta");
    meta.setAttribute("charset", "utf-8");
    document.head.appendChild(meta);
  }

  // Create body if not exists
  if (!document.body) {
    const body = document.createElement("body");

    // Move content to body
    Array.from(document.documentElement.childNodes).forEach((node) => {
      if (node !== document.head && node.nodeName.toLowerCase() !== "body") {
        body.appendChild(node);
      }
    });

    document.documentElement.appendChild(body);
  }
}

/**
 * Add document metadata based on style information
 * @param {Document} document - DOM document
 * @param {Object} styleInfo - Style information
 */
function addDocumentMetadata(document, styleInfo) {
  // Add title
  const title = document.createElement("title");
  title.textContent = "ICANN Board Resolutions";
  document.head.appendChild(title);

  // Add viewport meta
  const viewport = document.createElement("meta");
  viewport.setAttribute("name", "viewport");
  viewport.setAttribute("content", "width=device-width, initial-scale=1");
  document.head.appendChild(viewport);

  // Add class to body for specific document styling
  document.body.classList.add('icann-document');
}

/**
 * Process tables for better styling
 * @param {Document} document - DOM document
 */
function processTables(document) {
  const tables = document.querySelectorAll("table");
  tables.forEach((table) => {
    // Add default class if no class is present
    if (!table.classList.length) {
      table.classList.add("docx-table-default");
    }

    // Make sure tables have tbody
    if (!table.querySelector("tbody")) {
      const tbody = document.createElement("tbody");

      // Move rows to tbody
      Array.from(table.querySelectorAll("tr")).forEach((row) => {
        tbody.appendChild(row);
      });

      table.appendChild(tbody);
    }
    
    // Add responsive table wrapper
    const wrapper = document.createElement('div');
    wrapper.className = 'table-responsive';
    table.parentNode.insertBefore(wrapper, table);
    wrapper.appendChild(table);
  });
}

/**
 * Process images for better styling
 * @param {Document} document - DOM document
 */
function processImages(document) {
  const images = document.querySelectorAll("img");
  images.forEach((img) => {
    // Add default class if no class is present
    if (!img.classList.contains("docx-image")) {
      img.classList.add("docx-image");
    }

    // Ensure max-width for responsive images
    img.style.maxWidth = "100%";

    // Make sure images have alt text
    if (!img.hasAttribute("alt")) {
      img.setAttribute("alt", "Document image");
    }
    
    // Add figure wrapper for better semantics
    const figure = document.createElement('figure');
    img.parentNode.insertBefore(figure, img);
    figure.appendChild(img);
    
    // If there's text immediately after, it might be a caption
    if (img.nextSibling && img.nextSibling.nodeType === 3) {
      const caption = document.createElement('figcaption');
      caption.textContent = img.nextSibling.textContent.trim();
      figure.appendChild(caption);
      img.nextSibling.remove();
    }
  });
}

/**
 * Process language-specific elements
 * @param {Document} document - DOM document
 */
function processLanguageElements(document) {
  // Find elements with dir="rtl" and add class
  const rtlElements = document.querySelectorAll('[dir="rtl"]');
  rtlElements.forEach((el) => {
    el.classList.add("docx-rtl");
  });
  
  // Look for non-Latin scripts and add appropriate language attributes
  const allTextNodes = [];
  
  function collectTextNodes(node) {
    if (node.nodeType === 3) { // TEXT_NODE
      allTextNodes.push(node);
    } else {
      for (let i = 0; i < node.childNodes.length; i++) {
        collectTextNodes(node.childNodes[i]);
      }
    }
  }
  
  collectTextNodes(document.body);
  
  // Check for scripts and add language attributes
  const cyrillicRegex = /[\u0400-\u04FF]/;
  const chineseRegex = /[\u4E00-\u9FFF]/;
  const arabicRegex = /[\u0600-\u06FF]/;
  
  allTextNodes.forEach(textNode => {
    const text = textNode.nodeValue;
    const parent = textNode.parentNode;
    
    if (cyrillicRegex.test(text) && !parent.hasAttribute('lang')) {
      parent.setAttribute('lang', 'ru');
    } else if (chineseRegex.test(text) && !parent.hasAttribute('lang')) {
      parent.setAttribute('lang', 'zh');
    } else if (arabicRegex.test(text) && !parent.hasAttribute('lang')) {
      parent.setAttribute('lang', 'ar');
      if (!parent.hasAttribute('dir')) {
        parent.setAttribute('dir', 'rtl');
      }
    }
  });
}

/**
 * Process numbered paragraphs with proper nesting for hierarchical lists
 * @param {Document} document - DOM document
 */
function processNestedNumberedParagraphs(document) {
  // First, look for any existing lists (ordered lists) that may be TOC or other structured content
  const existingLists = document.querySelectorAll('ol.docx-numbered-list, ol.docx-alpha-list');
  
  // Process paragraphs that look like they should be part of a numbered list
  const paragraphs = document.querySelectorAll('p');
  
  // Regex patterns for different types of numbering
  const mainNumberPattern = /^\s*(\d+)\.\s+(.+)$/;
  const alphaPattern = /^\s*([a-z])\.\s+(.+)$/;
  const subNumberPattern = /^\s*(\d+\.\d+)\.\s+(.+)$/;
  
  // Current list structures
  let currentMainList = null;
  let currentMainItem = null;
  let currentSubList = null;
  let lastMainNumber = 0;
  
  // Process paragraphs
  for (let i = 0; i < paragraphs.length; i++) {
    const p = paragraphs[i];
    const text = p.textContent.trim();
    
    // Skip if this paragraph is already part of a list
    if (p.closest('ol, ul')) continue;
    
    // Skip any rationale paragraphs
    if (text.startsWith('Rationale for')) continue;

    // Check for main numbered items (1., 2., etc.)
    let mainMatch = text.match(mainNumberPattern);
    if (mainMatch) {
      const number = parseInt(mainMatch[1], 10);
      const content = mainMatch[2];
      
      // If this is a new list or a restart
      if (!currentMainList || number === 1) {
        currentMainList = document.createElement('ol');
        currentMainList.className = 'docx-numbered-list';
        p.parentNode.insertBefore(currentMainList, p);
      }
      
      // Create the list item
      currentMainItem = document.createElement('li');
      currentMainItem.textContent = content;
      currentMainItem.setAttribute('data-prefix', number);
      
      // Add the item to the list
      currentMainList.appendChild(currentMainItem);
      
      // Remember this number
      lastMainNumber = number;
      
      // Reset sublist tracking
      currentSubList = null;
      
      // Remove the original paragraph
      p.parentNode.removeChild(p);
      i--; // Adjust counter
      
      continue;
    }
    
    // Check for alpha subitems (a., b., etc.)
    let alphaMatch = text.match(alphaPattern);
    if (alphaMatch && currentMainItem) {
      const letter = alphaMatch[1];
      const content = alphaMatch[2];
      
      // If we don't have a sublist for this main item yet
      if (!currentSubList) {
        currentSubList = document.createElement('ol');
        currentSubList.className = 'docx-alpha-list';
        currentMainItem.appendChild(currentSubList);
      }
      
      // Create the sub-item
      const subItem = document.createElement('li');
      subItem.textContent = content;
      subItem.setAttribute('data-prefix', letter);
      
      // Add the sub-item to the sublist
      currentSubList.appendChild(subItem);
      
      // Remove the original paragraph
      p.parentNode.removeChild(p);
      i--; // Adjust counter
      
      continue;
    }
  }
}

/**
 * Generate additional CSS for TOC and list styling
 * @returns {string} - Additional CSS
 */
function generateAdditionalCss() {
  return `
/* Document defaults */
body {
  font-family: "Cambria", sans-serif;
  font-size: 12pt;
  line-height: 1.15;
  margin: 20px;
  padding: 0;
  color: #333;
  background-color: #fff;
}

/* For better mobile readability */
@media (max-width: 768px) {
  body {
    margin: 10px;
    font-size: 14pt;
  }
}

/* ICANN Document Specific Styling */
.icann-document h1, 
.icann-document h2, 
.icann-document h3, 
.icann-document h4, 
.icann-document h5, 
.icann-document h6 {
  color: #345A8A;
  font-family: "Cambria", sans-serif;
  margin-top: 1.5em;
  margin-bottom: 0.5em;
}

.icann-document h1 {
  font-size: 16pt;
  border-bottom: 1px solid #ddd;
  padding-bottom: 0.2em;
}

.icann-document h2 {
  font-size: 14pt;
}

.icann-document h3 {
  font-size: 13pt;
  font-style: italic;
}

/* Heading numbers */
.heading-number {
  font-weight: bold;
  color: #345A8A;
}

/* Improved TOC styling */
.toc-container {
  margin: 1.5em 0;
  border: 1px solid #eee;
  padding: 1em;
  background-color: #f9f9f9;
}

.toc-heading {
  margin-top: 1em !important;
  margin-bottom: 0.5em !important;
}

.toc-entry {
  display: flex !important;
  align-items: baseline;
  margin-bottom: 0.25em;
  line-height: 1.5;
}

.toc-entry .docx-toc-text {
  flex-grow: 0;
  flex-shrink: 1;
  margin-right: 0.5em;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  max-width: 80%;
}

.toc-entry .docx-toc-dots {
  flex-grow: 1;
  margin: 0 0.5em;
  height: 1px;
  border-bottom: 1px dotted #999;
}

.toc-entry .docx-toc-pagenum {
  flex-shrink: 0;
  text-align: right;
  min-width: 1.5em;
  font-weight: normal;
}

/* Improved list styling */
ol.docx-numbered-list {
  counter-reset: item;
  list-style-type: none;
  padding-left: 0;
  margin: 1em 0;
}

ol.docx-numbered-list > li {
  position: relative;
  padding-left: 2em;
  margin-bottom: 0.5em;
  line-height: 1.4;
}

ol.docx-numbered-list > li::before {
  position: absolute;
  left: 0;
  content: attr(data-prefix) ".";
  font-weight: bold;
  color: #345A8A;
}

ol.docx-alpha-list {
  counter-reset: alpha;
  list-style-type: none;
  padding-left: 0;
  margin: 0.5em 0 0.5em 1.5em;
}

ol.docx-alpha-list > li {
  position: relative;
  padding-left: 2em;
  margin-bottom: 0.3em;
  line-height: 1.4;
}

ol.docx-alpha-list > li::before {
  position: absolute;
  left: 0;
  content: attr(data-prefix) ".";
  font-weight: bold;
  color: #345A8A;
}

/* Special formatting for Rationale sections */
.docx-rationale {
  font-style: italic;
  margin: 0.5em 0 1em 2em;
  color: #666;
  border-left: 3px solid #ddd;
  padding-left: 1em;
}

/* Table improvements */
.table-responsive {
  overflow-x: auto;
  margin: 1em 0;
}

table.docx-table-default {
  width: 100%;
  border-collapse: collapse;
  margin: 0;
}

table.docx-table-default th,
table.docx-table-default td {
  border: 1px solid #ddd;
  padding: 8px;
  vertical-align: top;
}

table.docx-table-default th {
  background-color: #f2f2f2;
  font-weight: bold;
  text-align: left;
}

/* Image handling */
figure {
  margin: 1.5em 0;
  text-align: center;
}

.docx-image {
  max-width: 100%;
  height: auto;
  display: block;
  margin: 0 auto;
}

figcaption {
  font-style: italic;
  color: #666;
  margin-top: 0.5em;
  font-size: 0.9em;
}

/* Right-to-left text */
.docx-rtl {
  direction: rtl;
  text-align: right;
}

/* Language-specific styling */
[lang="zh"] {
  font-family: "SimSun", "NSimSun", "Microsoft YaHei", sans-serif;
}

[lang="ru"] {
  font-family: "Arial", sans-serif;
}

[lang="ar"] {
  font-family: "Tahoma", "Arial", sans-serif;
}

/* Print styling */
@media print {
  body {
    font-size: 10pt;
    margin: 0;
  }
  
  .toc-container {
    border: none;
    background: none;
  }
  
  h1, h2, h3, h4, h5, h6 {
    page-break-after: avoid;
    page-break-inside: avoid;
  }
  
  img, table, figure {
    page-break-inside: avoid;
  }
  
  .table-responsive {
    overflow-x: visible;
  }
}

/* Utility styles */
.docx-underline { text-decoration: underline; }
.docx-strike { text-decoration: line-through; }
.docx-tab { display: inline-block; width: 36pt; }
.docx-placeholder { 
  font-weight: bold; 
  text-align: center; 
  padding: 15px; 
  margin: 20px 0; 
  background-color: #f0f0f0; 
  border: 1px dashed #999; 
  border-radius: 5px; 
}
`;
}

module.exports = {
  extractAndApplyStyles,
  generateAdditionalCss,
};
