// lib/html/element-processors.js - Element processors (tables, images, language)

/**
 * Process tables for better styling
 * Enhances tables with proper structure and responsive wrappers
 * 
 * @param {Document} document - DOM document
 */
function processTables(document) {
  const tables = document.querySelectorAll("table");
  tables.forEach((table) => {
    // Add default class if no class is present
    if (!table.classList.length) {
      table.classList.add("docx-table-default");
    }

    // Process table structure and headers
    processTableStructure(table, document);
    
    // Add responsive table wrapper
    if (!table.parentNode.classList.contains('table-responsive')) {
      const wrapper = document.createElement('div');
      wrapper.className = 'table-responsive';
      table.parentNode.insertBefore(wrapper, table);
      wrapper.appendChild(table);
    }
  });
}

/**
 * Process table structure to add proper thead/tbody and convert first row to headers
 * @param {Element} table - Table element
 * @param {Document} document - DOM document
 */
function processTableStructure(table, document) {
  const rows = Array.from(table.querySelectorAll("tr"));
  
  if (rows.length === 0) return;
  
  // Check if table already has proper structure
  if (table.querySelector("thead") && table.querySelector("tbody")) {
    return;
  }
  
  // Create thead and tbody if they don't exist
  let thead = table.querySelector("thead");
  let tbody = table.querySelector("tbody");
  
  if (!thead) {
    thead = document.createElement("thead");
  }
  
  if (!tbody) {
    tbody = document.createElement("tbody");
  }
  
  // Clear existing structure
  table.innerHTML = '';
  
  // Process first row as header if it looks like a header
  const firstRow = rows[0];
  if (firstRow && isLikelyHeaderRow(firstRow)) {
    // Convert td elements to th elements in first row
    const cells = Array.from(firstRow.querySelectorAll("td"));
    cells.forEach(cell => {
      const th = document.createElement("th");
      th.innerHTML = cell.innerHTML;
      
      // Copy attributes
      Array.from(cell.attributes).forEach(attr => {
        th.setAttribute(attr.name, attr.value);
      });
      
      // Add scope attribute for accessibility
      th.setAttribute("scope", "col");
      
      cell.parentNode.replaceChild(th, cell);
    });
    
    thead.appendChild(firstRow);
    
    // Add remaining rows to tbody
    for (let i = 1; i < rows.length; i++) {
      tbody.appendChild(rows[i]);
    }
  } else {
    // All rows go to tbody
    rows.forEach(row => tbody.appendChild(row));
  }
  
  // Add thead and tbody to table
  if (thead.children.length > 0) {
    table.appendChild(thead);
  }
  if (tbody.children.length > 0) {
    table.appendChild(tbody);
  }
}

/**
 * Determine if a row is likely to be a header row
 * @param {Element} row - Table row element
 * @returns {boolean} - True if likely a header row
 */
function isLikelyHeaderRow(row) {
  const cells = Array.from(row.querySelectorAll("td"));
  
  if (cells.length === 0) return false;
  
  // Check if cells contain header-like content
  let headerIndicators = 0;
  
  cells.forEach(cell => {
    const text = cell.textContent.trim();
    
    // Check for header indicators
    if (text.length > 0) {
      // Short text (likely column names)
      if (text.length < 50) headerIndicators++;
      
      // Contains common header words
      if (/^(name|type|id|date|status|description|value|string|number|title|category)$/i.test(text)) {
        headerIndicators += 2;
      }
      
      // All caps (common in headers)
      if (text === text.toUpperCase() && text.length > 1) {
        headerIndicators++;
      }
      
      // Bold formatting (check for strong tags or bold styles)
      if (cell.querySelector("strong") || cell.style.fontWeight === "bold") {
        headerIndicators++;
      }
    }
  });
  
  // Consider it a header if we have enough indicators
  return headerIndicators >= cells.length * 0.5;
}

/**
 * Process images for better styling
 * Enhances images with proper attributes and figure wrappers
 * 
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
 * Handles RTL text and detects language attributes
 * 
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

module.exports = {
  processTables,
  processImages,
  processLanguageElements
};
