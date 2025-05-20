// style-extractor.js - Enhanced version
const mammoth = require("mammoth");
const { JSDOM } = require("jsdom");
const path = require("path");
const {
  parseDocxStyles,
  generateCssFromStyleInfo,
  selectNodes,
  selectSingleNode,
  convertTwipToPt
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
      styles: css + generateAdditionalCss(styleInfo), // Use document-specific styles
      messages: htmlResult.messages,
    };
  } catch (error) {
    console.error("Error extracting styles:", error);
    throw error;
  }
}

/**
 * Generate additional CSS for TOC and list styling based on document analysis
 * @param {Object} styleInfo - Extracted style information
 * @returns {string} - Additional CSS
 */
function generateAdditionalCss(styleInfo) {
  // Use the document's TOC style information to create appropriate CSS
  const tocStyles = styleInfo.tocStyles || {};
  const leaderStyle = tocStyles.leaderStyle || { character: '.', spacesBetween: 3 };
  
  return `
/* Enhanced TOC and list styling based on document analysis */
ol.docx-numbered-list {
  counter-reset: item;
  list-style-type: none;
  padding-left: 0;
  margin: 0;
  font-weight: normal;
}

ol.docx-numbered-list li {
  counter-increment: item;
  position: relative;
  padding-left: 2em;
  margin-bottom: 0.25em;
  line-height: 1.5;
}

ol.docx-numbered-list li::before {
  position: absolute;
  left: 0;
  content: attr(data-prefix);
  font-weight: bold;
}

/* Alpha list styles */
ol.docx-alpha-list {
  list-style-type: none;
  padding-left: 1em;
  margin: 0;
  font-weight: normal;
}

ol.docx-alpha-list li {
  position: relative;
  padding-left: 2em;
  margin-bottom: 0.25em;
  line-height: 1.5;
}

ol.docx-alpha-list li::before {
  position: absolute;
  left: 0;
  content: attr(data-prefix) ".";
  font-weight: bold;
}

/* TOC specific styles optimized for the document */
.docx-toc-list {
  margin: 0;
  padding: 0;
  line-height: 1.5;
}

.docx-toc-item {
  display: flex;
  align-items: baseline;
  margin-bottom: 0.25em;
  width: 100%;
  white-space: nowrap;
}

.docx-toc-text {
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}

.docx-toc-dots {
  flex-grow: 1;
  margin: 0 0.5em;
  height: 1px;
  border-bottom: 1px dotted #000;
  min-width: 2em;
}

.docx-toc-pagenum {
  flex-shrink: 0;
  text-align: right;
  min-width: 1.5em;
}

/* Special formatting for TOC headings */
.docx-toc-heading {
  font-weight: bold;
  margin-bottom: 1em;
}

/* Make the TOC look more like the Word document */
ol.docx-numbered-list.docx-toc-list {
  margin-bottom: 1em;
}

ol.docx-numbered-list.docx-toc-list li {
  padding-left: 2.5em;
  margin-bottom: 0;
  line-height: 1.8;
}

ol.docx-alpha-list.docx-toc-list li {
  padding-left: 2.5em;
  margin-bottom: 0;
  line-height: 1.8;
}

/* Improve the appearance of the dotted lines - use actual document settings */
.docx-toc-dots {
  border-bottom: 1px ${leaderStyle.character === '.' ? 'dotted' : leaderStyle.character === '_' ? 'solid' : 'dashed'} #666;
  margin: 0 0.5em;
  position: relative;
  top: -0.3em;
}

/* Ensure page numbers are properly aligned */
.docx-toc-pagenum {
  padding-left: 0.5em;
  font-weight: normal;
}

/* Fix spacing between main items and sub-items */
ol.docx-numbered-list li ol.docx-alpha-list {
  margin-top: 0.25em;
}

/* Special styling for "Rationale for Resolution" sections */
.docx-rationale {
  font-style: italic;
  margin-top: 0.25em;
  margin-bottom: 0.5em;
  margin-left: 2em;
  font-size: 0.95em;
  color: #333;
}
`;
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

  // Map TOC styles if they exist
  if (styleInfo.tocStyles && styleInfo.tocStyles.tocEntryStyles) {
    styleInfo.tocStyles.tocEntryStyles.forEach((style) => {
      if (style.id) {
        styleMap.push(
          `p[style-name='${style.name || style.id}'] => p.docx-toc-entry.docx-toc-level-${style.level || 1}`
        );
      }
    });
  }

  // Additional custom mappings for specific elements
  styleMap.push("p:fresh => p");
  styleMap.push("r[bold] => strong");
  styleMap.push("r[italic] => em");
  styleMap.push("r[underline] => span.docx-underline");
  styleMap.push("r[strikethrough] => span.docx-strike");
  styleMap.push("r[subscript] => sub");
  styleMap.push("r[superscript] => sup");

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
      processTables(document, styleInfo);
    } catch (error) {
      console.error("Error processing tables:", error.message);
    }

    // Process images to maintain aspect ratio and positioning
    try {
      processImages(document, styleInfo);
    } catch (error) {
      console.error("Error processing images:", error.message);
    }

    // Handle language-specific elements
    try {
      processLanguageElements(document, styleInfo);
    } catch (error) {
      console.error("Error processing language elements:", error.message);
    }

    // Process numbered paragraphs with proper nesting
    try {
      processNestedNumberedParagraphs(document, styleInfo);
    } catch (error) {
      console.error("Error processing numbered paragraphs:", error.message);
    }

    // Style and enhance TOC and index elements
    try {
      detectAndStyleTocAndIndex(document, styleInfo);
    } catch (error) {
      console.error("Error processing TOC:", error.message);
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
  title.textContent = "DOCX Document";
  document.head.appendChild(title);

  // Add viewport meta
  const viewport = document.createElement("meta");
  viewport.setAttribute("name", "viewport");
  viewport.setAttribute("content", "width=device-width, initial-scale=1");
  document.head.appendChild(viewport);
}

/**
 * Process tables for better styling based on document analysis
 * @param {Document} document - DOM document
 * @param {Object} styleInfo - Style information
 */
function processTables(document, styleInfo) {
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
    
    // Apply any document-specific table styles
    if (styleInfo.styles && styleInfo.styles.table) {
      // Find matching style and apply additional properties
      Object.entries(styleInfo.styles.table).forEach(([id, style]) => {
        if (table.classList.contains(`docx-t-${id.toLowerCase()}`)) {
          // Apply any specific styles from the document
          if (style.borders) {
            const borderWidth = style.borders.top?.size ? 
                              convertTwipToPt(style.borders.top.size) + 'pt' : 
                              '1px';
            const borderColor = style.borders.top?.color ? 
                              `#${style.borders.top.color}` : 
                              '#000';
            
            table.style.border = `${borderWidth} solid ${borderColor}`;
          }
        }
      });
    }
  });
}

/**
 * Process images for better styling based on document analysis
 * @param {Document} document - DOM document
 * @param {Object} styleInfo - Style information
 */
function processImages(document, styleInfo) {
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
  });
}

/**
 * Process language-specific elements based on document analysis
 * @param {Document} document - DOM document
 * @param {Object} styleInfo - Style information
 */
function processLanguageElements(document, styleInfo) {
  // Find elements with dir="rtl" and add class
  const rtlElements = document.querySelectorAll('[dir="rtl"]');
  rtlElements.forEach((el) => {
    el.classList.add("docx-rtl");
  });
  
  // Apply document-specific language settings if available
  if (styleInfo.settings?.rtlGutter) {
    document.body.dir = "rtl";
    document.body.classList.add("docx-rtl");
  }
}

/**
 * Process numbered paragraphs with proper nesting based on document analysis
 * @param {Document} document - DOM document
 * @param {Object} styleInfo - Style information
 */
function processNestedNumberedParagraphs(document, styleInfo) {
  const paragraphs = document.querySelectorAll("p");

  // Patterns for identifying different types of paragraph numbering
  const numberPattern = /^\s*(\d+)\.(.+)$/;
  const alphaPattern = /^\s*([a-z])\.(.+)$/;
  const romanPattern = /^\s*([ivx]+)\.(.+)$/;

  // Use document structure analysis to find special sections
  const specialSections = styleInfo.documentStructure?.specialSections || [];
  const rationaleSections = specialSections.filter(section => 
    section.type === 'rationale');

  // Track processed items to avoid duplicates
  const processedTexts = new Set();

  // Track lists and their hierarchy
  let mainList = null;
  let currentMainItem = null;
  let currentSubList = null;
  let lastListType = null;
  let lastMainNumber = 0;

  // Process paragraphs sequentially
  for (let i = 0; i < paragraphs.length; i++) {
    const p = paragraphs[i];
    
    // Skip if p is null or doesn't have text content
    if (!p || !p.textContent) continue;
    
    const text = p.textContent;
    
    // Skip if we've already processed this exact text
    if (processedTexts.has(text)) {
      // Remove the duplicate paragraph
      if (p.parentNode) {
        p.parentNode.removeChild(p);
        i--; // Adjust counter since we removed an element
      }
      continue;
    }
    
    // Check for "Rationale for Resolution" special case
    if (text.includes('Rationale for Resolution')) {
      p.classList.add('docx-rationale');
      processedTexts.add(text);
      continue;
    }
    
    // Check for different numbering patterns
    let match = null;
    let listType = null;
    
    if (numberPattern.test(text)) {
      match = text.match(numberPattern);
      listType = "numbered";
    } else if (alphaPattern.test(text)) {
      match = text.match(alphaPattern);
      listType = "alpha";
    } else if (romanPattern.test(text)) {
      match = text.match(romanPattern);
      listType = "roman";
    }
    
    if (match) {
      // Extract the number/letter and content
      const prefix = match[1];
      const content = match[2].trim();
      
      // Add to processed texts set
      processedTexts.add(text);
      
      // Check if this is a TOC item
      const isTocItem = text.includes("\t") || text.includes("   ") || /\d+$/.test(text);
      
      // Handle main numbered items (1., 2., etc.)
      if (listType === "numbered") {
        // If we don't have a main list yet, or if this is a new main list
        if (!mainList || (lastListType === "numbered" && parseInt(prefix) === 1)) {
          // Create a new main list
          mainList = document.createElement("ol");
          mainList.className = "docx-numbered-list";
          
          if (isTocItem) {
            mainList.classList.add("docx-toc-list");
          }
          
          // Insert the list before the paragraph
          if (p.parentNode) {
            p.parentNode.insertBefore(mainList, p);
          } else {
            console.log("Warning: Cannot insert list, paragraph has no parent");
            continue;
          }
        }
        
        // Create the list item
        currentMainItem = document.createElement("li");
        currentMainItem.textContent = content;
        currentMainItem.setAttribute("data-prefix", prefix);
        
        if (isTocItem) {
          currentMainItem.classList.add("docx-toc-item");
          
          // Format TOC item with page number
          const pageNumMatch = content.match(/(\d+)$/);
          if (pageNumMatch) {
            const textPart = content
              .substring(0, content.lastIndexOf(pageNumMatch[1]))
              .trim();
            const pageNum = pageNumMatch[1];
            
            // Clear the content and recreate with spans
            currentMainItem.textContent = "";
            
            const textSpan = document.createElement("span");
            textSpan.classList.add("docx-toc-text");
            textSpan.textContent = textPart;
            
            const dotsSpan = document.createElement("span");
            dotsSpan.classList.add("docx-toc-dots");
            
            const pageSpan = document.createElement("span");
            pageSpan.classList.add("docx-toc-pagenum");
            pageSpan.textContent = pageNum;
            
            currentMainItem.appendChild(textSpan);
            currentMainItem.appendChild(dotsSpan);
            currentMainItem.appendChild(pageSpan);
          }
        }
        
        // Add the item to the main list
        mainList.appendChild(currentMainItem);
        
        // Reset sub-list tracking
        currentSubList = null;
        lastMainNumber = parseInt(prefix);
        
        // Remove the original paragraph
        if (p.parentNode) {
          p.parentNode.removeChild(p);
          i--; // Adjust counter since we removed an element
        }
      }
      // Handle alpha items (a., b., etc.) as sub-items
      else if (listType === "alpha") {
        // If we have a main item to attach to
        if (currentMainItem) {
          // If we don't have a sub-list for this main item yet
          if (!currentSubList) {
            // Create a new sub-list
            currentSubList = document.createElement("ol");
            currentSubList.className = "docx-alpha-list";
            
            if (isTocItem) {
              currentSubList.classList.add("docx-toc-list");
            }
            
            // Append the sub-list to the current main item
            currentMainItem.appendChild(currentSubList);
          }
          
          // Create the sub-item
          const subItem = document.createElement("li");
          subItem.textContent = content;
          subItem.setAttribute("data-prefix", prefix);
          
          if (isTocItem) {
            subItem.classList.add("docx-toc-item");
            
            // Format TOC item with page number
            const pageNumMatch = content.match(/(\d+)$/);
            if (pageNumMatch) {
              const textPart = content
                .substring(0, content.lastIndexOf(pageNumMatch[1]))
                .trim();
              const pageNum = pageNumMatch[1];
              
              // Clear the content and recreate with spans
              subItem.textContent = "";
              
              const textSpan = document.createElement("span");
              textSpan.classList.add("docx-toc-text");
              textSpan.textContent = textPart;
              
              const dotsSpan = document.createElement("span");
              dotsSpan.classList.add("docx-toc-dots");
              
              const pageSpan = document.createElement("span");
              pageSpan.classList.add("docx-toc-pagenum");
              pageSpan.textContent = pageNum;
              
              subItem.appendChild(textSpan);
              subItem.appendChild(dotsSpan);
              subItem.appendChild(pageSpan);
            }
          }
          
          // Add the sub-item to the sub-list
          currentSubList.appendChild(subItem);
          
          // Remove the original paragraph
          if (p.parentNode) {
            p.parentNode.removeChild(p);
            i--; // Adjust counter since we removed an element
          }
        }
        // If we don't have a main item to attach to, create a standalone alpha list
        else {
          // This is a special case where we have alpha items without a preceding numbered item
          // Create a new list if needed
          if (!currentSubList || lastListType !== "alpha") {
            currentSubList = document.createElement("ol");
            currentSubList.className = "docx-alpha-list";
            
            if (isTocItem) {
              currentSubList.classList.add("docx-toc-list");
            }
            
            // Insert the list before the paragraph
            if (p.parentNode) {
              p.parentNode.insertBefore(currentSubList, p);
            } else {
              console.log("Warning: Cannot insert list, paragraph has no parent");
              continue;
            }
          }
          
          // Create the list item
          const item = document.createElement("li");
          item.textContent = content;
          item.setAttribute("data-prefix", prefix);
          
          if (isTocItem) {
            item.classList.add("docx-toc-item");
            
            // Format TOC item with page number
            const pageNumMatch = content.match(/(\d+)$/);
            if (pageNumMatch) {
              const textPart = content
                .substring(0, content.lastIndexOf(pageNumMatch[1]))
                .trim();
              const pageNum = pageNumMatch[1];
              
              // Clear the content and recreate with spans
              item.textContent = "";
              
              const textSpan = document.createElement("span");
              textSpan.classList.add("docx-toc-text");
              textSpan.textContent = textPart;
              
              const dotsSpan = document.createElement("span");
              dotsSpan.classList.add("docx-toc-dots");
              
              const pageSpan = document.createElement("span");
              pageSpan.classList.add("docx-toc-pagenum");
              pageSpan.textContent = pageNum;
              
              item.appendChild(textSpan);
              item.appendChild(dotsSpan);
              item.appendChild(pageSpan);
            }
          }
          
          // Add the item to the list
          currentSubList.appendChild(item);
          
          // Remove the original paragraph
          if (p.parentNode) {
            p.parentNode.removeChild(p);
            i--; // Adjust counter since we removed an element
          }
        }
      }
      // Handle roman numerals similarly to alpha items
      else if (listType === "roman") {
        // Similar handling as alpha items...
        // (Implementation would be similar to alpha items)
      }
      
      // Update last list type
      lastListType = listType;
    }
    // Handle non-list paragraphs
    else {
      // Add to processed texts set for non-list paragraphs too
      processedTexts.add(text);
      
      // If this is a paragraph that separates list sections but doesn't break the overall list
      // (like a "Rationale for Resolution" paragraph between list items)
      if (text.trim().startsWith("Rationale for") && mainList) {
        // Don't reset list tracking, apply special styling
        p.classList.add("docx-rationale");
      }
      // Otherwise, reset list tracking for non-list paragraphs
      else {
        // Reset all list tracking
        if (lastListType !== null) {
          lastListType = null;
          
          // Only reset current items if this isn't a special case paragraph
          if (!text.trim().startsWith("Rationale for")) {
            currentMainItem = null;
            currentSubList = null;
          }
        }
      }
    }
  }
}

/**
 * Detect and style TOC and index elements based on document analysis
 * @param {Document} document - DOM document
 * @param {Object} styleInfo - Style information
 */
function detectAndStyleTocAndIndex(document, styleInfo) {
  try {
    // Use TOC information from document analysis
    const tocInfo = styleInfo.tocStyles || {};
    const hasToc = tocInfo.hasTableOfContents || styleInfo.documentStructure?.hasToc;
    
    if (hasToc) {
      // Look for TOC headings first
      const headings = Array.from(document.querySelectorAll("h1, h2, h3, p")).filter(
        el => el.textContent.includes("Table of Contents") || 
             el.textContent.includes("Contents") ||
             (el.className && el.className.includes("TOC"))
      );
      
      if (headings.length > 0) {
        // Style the first matching heading as TOC heading
        const tocHeading = headings[0];
        tocHeading.classList.add("docx-toc-heading");
        
        // Apply any specific TOC heading styles from the document
        if (tocInfo.tocHeadingStyle) {
          if (tocInfo.tocHeadingStyle.fontSize) {
            tocHeading.style.fontSize = tocInfo.tocHeadingStyle.fontSize;
          }
          if (tocInfo.tocHeadingStyle.fontFamily) {
            tocHeading.style.fontFamily = tocInfo.tocHeadingStyle.fontFamily;
          }
        }
      }
      
      // Now look for paragraphs with tab characters and page numbers that follow a TOC heading
      if (headings.length > 0) {
        const tocHeading = headings[0];
        let nextElement = tocHeading.nextElementSibling;
        
        // Skip until we find a paragraph or list
        while (nextElement && 
              nextElement.nodeName !== 'P' && 
              nextElement.nodeName !== 'OL' &&
              nextElement.nodeName !== 'UL') {
          nextElement = nextElement.nextElementSibling;
        }
        
        // If we found a list, we've already processed it as a TOC list
        if (nextElement && (nextElement.nodeName === 'OL' || nextElement.nodeName === 'UL')) {
          // If it doesn't already have the TOC list class, add it
          if (!nextElement.classList.contains('docx-toc-list')) {
            nextElement.classList.add('docx-toc-list');
          }
        }
        // If we found a paragraph, check if it's a potential TOC entry
        else if (nextElement && nextElement.nodeName === 'P') {
          const text = nextElement.textContent;
          
          // Check if it might be a TOC entry (has tabs, spaces, or ends with a number)
          if (text.includes("\t") || text.includes("   ") || /\d+\s*$/.test(text)) {
            // Create a container for TOC entries
            const tocContainer = document.createElement('div');
            tocContainer.classList.add('docx-toc');
            
            // Insert the container before the first TOC entry
            nextElement.parentNode.insertBefore(tocContainer, nextElement);
            
            // Process potential TOC entries
            let currentElement = nextElement;
            let entryCount = 0;
            
            while (currentElement && entryCount < 50) { // Limit to avoid infinite loop
              const text = currentElement.textContent;
              
              // Check if this still looks like a TOC entry
              if (text.includes("\t") || text.includes("   ") || /\d+\s*$/.test(text)) {
                // Create a TOC entry element
                const entry = document.createElement('div');
                entry.classList.add('docx-toc-entry');
                
                // Determine the level based on indentation or numbering
                let level = 1;
                if (text.match(/^\s*[a-z]\./)) {
                  level = 2; // Alpha entries are level 2
                } else if (text.trim().match(/^\d+\.\d+/)) {
                  level = 2; // Entries with format 1.1 are level 2
                }
                
                entry.classList.add(`docx-toc-level-${level}`);
                
                // Extract text and page number
                const pageMatch = text.match(/(.*?)(\d+)\s*$/);
                if (pageMatch) {
                  const content = pageMatch[1].trim();
                  const pageNum = pageMatch[2].trim();
                  
                  // Create formatted entry
                  const textSpan = document.createElement('span');
                  textSpan.classList.add('docx-toc-text');
                  textSpan.textContent = content;
                  
                  const dotsSpan = document.createElement('span');
                  dotsSpan.classList.add('docx-toc-dots');
                  
                  const pageSpan = document.createElement('span');
                  pageSpan.classList.add('docx-toc-pagenum');
                  pageSpan.textContent = pageNum;
                  
                  entry.appendChild(textSpan);
                  entry.appendChild(dotsSpan);
                  entry.appendChild(pageSpan);
                } else {
                  // Just use the text if we can't extract a page number
                  entry.textContent = text;
                }
                
                // Add the entry to the TOC container
                tocContainer.appendChild(entry);
                
                // Save the next element before removing the current one
                const nextToProcess = currentElement.nextElementSibling;
                
                // Remove the original paragraph
                currentElement.parentNode.removeChild(currentElement);
                
                // Move to the next element
                currentElement = nextToProcess;
                entryCount++;
              } else {
                // Break if we've found a non-TOC element
                break;
              }
            }
          }
        }
      }
    }
    
    // Process Index elements
    // Look for Index headings
    const indexHeadings = Array.from(document.querySelectorAll("h1, h2, h3, p")).filter(
      el => el.textContent.includes("Index")
    );
    
    if (indexHeadings.length > 0) {
      // Style the first matching heading as Index heading
      const indexHeading = indexHeadings[0];
      indexHeading.classList.add("docx-index-heading");
      
      // Create a placeholder for the index if needed
      if (indexHeading.nextElementSibling && 
          indexHeading.nextElementSibling.nodeName === 'P' &&
          indexHeading.nextElementSibling.textContent.trim() === '') {
        const placeholder = document.createElement('p');
        placeholder.classList.add('docx-placeholder');
        placeholder.textContent = '** INDEX HERE **';
        
        indexHeading.parentNode.insertBefore(placeholder, indexHeading.nextElementSibling);
      }
    }
    
  } catch (error) {
    console.error("Error in detectAndStyleTocAndIndex:", error.message);
  }
}

module.exports = {
  extractAndApplyStyles,
  generateAdditionalCss,
};
