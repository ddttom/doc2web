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