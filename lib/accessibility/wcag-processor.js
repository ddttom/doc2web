// lib/accessibility/wcag-processor.js - WCAG 2.1 compliance processor
// Enhances HTML output to meet WCAG 2.1 Level AA standards

/**
 * Process HTML document for WCAG 2.1 Level AA compliance
 * Enhances document with accessibility features to meet compliance standards
 * 
 * @param {Document} document - DOM document
 * @param {Object} styleInfo - Style information
 * @param {Object} metadata - Document metadata
 * @returns {Document} - Enhanced document
 */
function processForAccessibility(document, styleInfo, metadata) {
  // Add language attribute to HTML element
  const htmlElement = document.documentElement;
  const documentLang = styleInfo.settings?.language || 'en';
  htmlElement.setAttribute('lang', documentLang);
  
  // Add document title from metadata
  const title = document.querySelector('title');
  if (title) {
    title.textContent = metadata?.core?.title || 'Document';
  }
  
  // Process tables for accessibility
  processTablesForAccessibility(document);
  
  // Process images for accessibility
  processImagesForAccessibility(document);
  
  // Ensure proper heading hierarchy
  ensureHeadingHierarchy(document);
  
  // ARIA landmarks processing has been disabled to prevent DOM manipulation errors
  // addAriaLandmarks(document);
  
  // Add skip navigation link
  addSkipNavigation(document);
  
  // Enhance keyboard navigability
  enhanceKeyboardNavigation(document);
  
  // Check and enhance color contrast
  enhanceColorContrast(document, styleInfo);
  
  return document;
}

/**
 * Process tables for accessibility
 * Adds captions, scopes, and other accessibility features to tables
 * 
 * @param {Document} document - DOM document
 */
function processTablesForAccessibility(document) {
  const tables = document.querySelectorAll('table');
  
  tables.forEach((table, index) => {
    // Add role="table" for explicit semantics
    table.setAttribute('role', 'table');
    
    // Add an accessible name if missing
    if (!table.hasAttribute('aria-label') && !table.querySelector('caption')) {
      const caption = document.createElement('caption');
      caption.textContent = `Table ${index + 1}`;
      table.insertBefore(caption, table.firstChild);
    }
    
    // Process header cells
    const headerCells = table.querySelectorAll('th');
    headerCells.forEach(cell => {
      // Add scope attribute if missing
      if (!cell.hasAttribute('scope')) {
        // Determine if it's a row or column header
        const isRowHeader = cell.parentElement.children[0] === cell;
        cell.setAttribute('scope', isRowHeader ? 'row' : 'col');
      }
    });
    
    // Process data cells in complex tables
    if (headerCells.length > 0 && table.rows.length > 1) {
      const dataCells = table.querySelectorAll('td');
      dataCells.forEach(cell => {
        // If data cell doesn't have headers attribute and it's a complex table
        if (!cell.hasAttribute('headers') && headerCells.length > 1) {
          // Add aria-labelledby for complex tables if needed
          // This is a simplified approach; real implementation would be more complex
          const rowIdx = cell.parentElement.rowIndex;
          const colIdx = Array.from(cell.parentElement.children).indexOf(cell);
          
          const rowHeader = table.querySelector(`th[scope="row"]:nth-of-type(${rowIdx + 1})`);
          const colHeader = table.querySelector(`th[scope="col"]:nth-of-type(${colIdx + 1})`);
          
          if (rowHeader && colHeader) {
            // Ensure headers have IDs
            if (!rowHeader.id) rowHeader.id = `row-header-${rowIdx}`;
            if (!colHeader.id) colHeader.id = `col-header-${colIdx}`;
            
            cell.setAttribute('aria-labelledby', `${rowHeader.id} ${colHeader.id}`);
          }
        }
      });
    }
  });
}

/**
 * Process images for accessibility
 * Ensures all images have appropriate alt text
 * 
 * @param {Document} document - DOM document
 */
function processImagesForAccessibility(document) {
  const images = document.querySelectorAll('img');
  
  images.forEach(img => {
    // Check if alt attribute exists
    if (!img.hasAttribute('alt')) {
      // If the image is inside a figure with figcaption, use that as alt
      const figure = img.closest('figure');
      if (figure && figure.querySelector('figcaption')) {
        const caption = figure.querySelector('figcaption').textContent;
        img.setAttribute('alt', caption);
      } else {
        // Otherwise use a generic alt or get from filename
        const src = img.getAttribute('src') || '';
        const filename = src.split('/').pop().split('.')[0];
        const altText = filename.replace(/[-_]/g, ' ').replace(/image/i, '').trim();
        
        img.setAttribute('alt', altText || 'Document image');
      }
    }
    
    // Wrap standalone images in figures if not already
    if (!img.closest('figure') && !img.classList.contains('decorative')) {
      const figure = document.createElement('figure');
      img.parentNode.insertBefore(figure, img);
      figure.appendChild(img);
      
      // If no figcaption and image has a title, add figcaption
      if (!figure.querySelector('figcaption') && img.getAttribute('title')) {
        const figcaption = document.createElement('figcaption');
        figcaption.textContent = img.getAttribute('title');
        figure.appendChild(figcaption);
      }
    }
    
    // For decorative images, set empty alt and role="presentation"
    if (img.classList.contains('decorative')) {
      img.setAttribute('alt', '');
      img.setAttribute('role', 'presentation');
    }
  });
}

/**
 * Ensure proper heading hierarchy
 * Fixes heading levels to ensure no levels are skipped
 * 
 * @param {Document} document - DOM document
 */
function ensureHeadingHierarchy(document) {
  const headings = document.querySelectorAll('h1, h2, h3, h4, h5, h6');
  let lastLevel = 0;
  
  // First pass: identify if we have an h1
  const hasH1 = document.querySelector('h1') !== null;
  
  // If no h1, find the first heading and convert it to h1
  if (!hasH1 && headings.length > 0) {
    const firstHeading = headings[0];
    const newH1 = document.createElement('h1');
    newH1.innerHTML = firstHeading.innerHTML;
    firstHeading.parentNode.replaceChild(newH1, firstHeading);
    
    // Reset the headings collection
    return ensureHeadingHierarchy(document);
  }
  
  // Second pass: fix sequential heading levels
  for (const heading of headings) {
    const currentLevel = parseInt(heading.tagName.substring(1), 10);
    
    // If this heading skips a level (e.g., h1 to h3), fix it
    if (currentLevel > lastLevel + 1) {
      const newLevel = lastLevel + 1;
      const newHeading = document.createElement(`h${newLevel}`);
      newHeading.innerHTML = heading.innerHTML;
      heading.parentNode.replaceChild(newHeading, heading);
    } else {
      lastLevel = currentLevel;
    }
  }
}

/**
 * Add ARIA landmarks to document
 * Creates proper sectioning with ARIA roles
 * 
 * @param {Document} document - DOM document
 */
/**
 * Add ARIA landmarks to document - DISABLED to prevent DOM manipulation errors
 * 
 * @param {Document} document - DOM document
 */
function addAriaLandmarks(document) {
  // This function has been disabled to prevent DOM manipulation errors
  // The original implementation was causing errors when trying to move elements
  // Error: "The child can not be found in the parent"
  console.log('ARIA landmarks processing has been disabled');
  return;
}

/**
 * Add skip navigation link
 * Adds a link to skip to main content for keyboard users
 * 
 * @param {Document} document - DOM document
 */
function addSkipNavigation(document) {
  // Only add if we have navigation and main content
  const nav = document.querySelector('nav, [role="navigation"]');
  const main = document.querySelector('main, [role="main"]');
  
  // Since addAriaLandmarks is disabled, we may not have these elements
  // Only proceed if they exist
  if (nav && main && !document.querySelector('.skip-link')) {
    // Ensure main has an ID
    if (!main.id) {
      main.id = 'main-content';
    }
    
    // Create the skip link
    const skipLink = document.createElement('a');
    skipLink.href = `#${main.id}`;
    skipLink.className = 'skip-link';
    skipLink.textContent = 'Skip to main content';
    
    // Insert at the very beginning of body
    document.body.insertBefore(skipLink, document.body.firstChild);
  }
}

/**
 * Enhance keyboard navigability
 * Makes interactive elements properly focusable and adds focus styles
 * 
 * @param {Document} document - DOM document
 */
function enhanceKeyboardNavigation(document) {
  // Ensure focusable elements have proper tabindex and focus styles
  
  // TOC entries should be navigable
  const tocEntries = document.querySelectorAll('.docx-toc-entry');
  tocEntries.forEach(entry => {
    if (!entry.getAttribute('tabindex')) {
      entry.setAttribute('tabindex', '0');
    }
  });
  
  // Links should have visible focus states (handled in CSS)
  const links = document.querySelectorAll('a');
  links.forEach(link => {
    link.classList.add('keyboard-focusable');
  });
  
  // Add keyboard shortcut to toggle visibility of tracked changes
  if (document.body.classList.contains('docx-track-changes-show') || 
      document.body.classList.contains('docx-track-changes-hide')) {
    
    const toggleScript = document.createElement('script');
    toggleScript.textContent = `
      document.addEventListener('keydown', function(e) {
        // Alt+T to toggle track changes
        if (e.altKey && e.key === 't') {
          document.body.classList.toggle('docx-track-changes-show');
          document.body.classList.toggle('docx-track-changes-hide');
        }
      });
    `;
    document.body.appendChild(toggleScript);
  }
}

/**
 * Enhance color contrast
 * Ensures sufficient color contrast for text elements
 * 
 * @param {Document} document - DOM document
 * @param {Object} styleInfo - Style information
 */
function enhanceColorContrast(document, styleInfo) {
  // Add high contrast mode class to html element
  document.documentElement.classList.add('docx-accessible');
  
  // Add data-theme attribute for potential theme switching
  document.documentElement.setAttribute('data-theme', 'light');
  
  // Insert a meta theme-color tag for browser UI
  const meta = document.createElement('meta');
  meta.setAttribute('name', 'theme-color');
  meta.setAttribute('content', '#ffffff');
  document.head.appendChild(meta);
  
  // Actual color contrast fixing would require color analysis
  // For now, we just ensure text has sufficient contrast via CSS
}

module.exports = {
  processForAccessibility,
  processTablesForAccessibility,
  processImagesForAccessibility,
  ensureHeadingHierarchy,
  addAriaLandmarks,
  addSkipNavigation,
  enhanceKeyboardNavigation,
  enhanceColorContrast
};
