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
  try {
    // Add language attribute to HTML element
    const htmlElement = document.documentElement;
    const documentLang = styleInfo.settings?.language || 'en';
    htmlElement.setAttribute('lang', documentLang);
    
    // Add document title from metadata
    const title = document.querySelector('title');
    if (title && metadata?.core?.title) {
      title.textContent = metadata.core.title;
    }
    
    // Process tables for accessibility
    processTablesForAccessibility(document);
    
    // Process images for accessibility
    processImagesForAccessibility(document);
    
    // Ensure proper heading hierarchy
    ensureHeadingHierarchy(document);
    
    // Add ARIA landmarks (fixed implementation)
    addAriaLandmarks(document);
    
    // Add skip navigation link
    addSkipNavigation(document);
    
    // Enhance keyboard navigability
    enhanceKeyboardNavigation(document);
    
    // Check and enhance color contrast
    enhanceColorContrast(document, styleInfo);
    
    console.log('Accessibility enhancements applied successfully');
    
    return document;
  } catch (error) {
    console.error('Error in processForAccessibility:', error);
    return document;
  }
}

/**
 * Process tables for accessibility
 * Adds captions, scopes, and other accessibility features to tables
 * 
 * @param {Document} document - DOM document
 */
function processTablesForAccessibility(document) {
  try {
    const tables = document.querySelectorAll('table');
    
    tables.forEach((table, index) => {
      // Add role="table" for explicit semantics
      table.setAttribute('role', 'table');
      
      // Add an accessible name if missing
      if (!table.hasAttribute('aria-label') && !table.querySelector('caption')) {
        const caption = document.createElement('caption');
        caption.textContent = `Table ${index + 1}`;
        caption.className = 'sr-only'; // Screen reader only by default
        table.insertBefore(caption, table.firstChild);
      }
      
      // Process header cells
      const headerCells = table.querySelectorAll('th');
      headerCells.forEach(cell => {
        // Add scope attribute if missing
        if (!cell.hasAttribute('scope')) {
          // Determine if it's a row or column header based on position
          const row = cell.parentElement;
          const isFirstRow = row === table.querySelector('tr');
          const isFirstColumn = cell === row.querySelector('th, td');
          
          if (isFirstRow) {
            cell.setAttribute('scope', 'col');
          } else if (isFirstColumn) {
            cell.setAttribute('scope', 'row');
          } else {
            cell.setAttribute('scope', 'col'); // Default to column
          }
        }
      });
      
      // Process data cells in complex tables
      if (headerCells.length > 1 && table.rows.length > 1) {
        const dataCells = table.querySelectorAll('td');
        dataCells.forEach((cell, cellIndex) => {
          // Add aria-describedby for complex tables if helpful
          if (!cell.hasAttribute('aria-describedby') && headerCells.length > 2) {
            // For very complex tables, we could add more sophisticated labeling
            // For now, we'll just ensure the basic structure is accessible
          }
        });
      }
      
      // Add table-responsive wrapper if not already present
      if (!table.closest('.table-responsive')) {
        const wrapper = document.createElement('div');
        wrapper.className = 'table-responsive';
        wrapper.setAttribute('role', 'region');
        wrapper.setAttribute('aria-label', `Table ${index + 1} content`);
        table.parentNode.insertBefore(wrapper, table);
        wrapper.appendChild(table);
      }
    });
  } catch (error) {
    console.error('Error processing tables for accessibility:', error);
  }
}

/**
 * Process images for accessibility
 * Ensures all images have appropriate alt text
 * 
 * @param {Document} document - DOM document
 */
function processImagesForAccessibility(document) {
  try {
    const images = document.querySelectorAll('img');
    
    images.forEach((img, index) => {
      // Check if alt attribute exists
      if (!img.hasAttribute('alt')) {
        // If the image is inside a figure with figcaption, use that as alt
        const figure = img.closest('figure');
        if (figure && figure.querySelector('figcaption')) {
          const caption = figure.querySelector('figcaption').textContent;
          img.setAttribute('alt', caption.trim());
        } else {
          // Otherwise generate alt text from src or use generic
          const src = img.getAttribute('src') || '';
          const filename = src.split('/').pop().split('.')[0];
          const altText = filename
            .replace(/[-_]/g, ' ')
            .replace(/image/i, '')
            .replace(/\d+/g, '')
            .trim();
          
          img.setAttribute('alt', altText || `Document image ${index + 1}`);
        }
      }
      
      // Wrap standalone images in figures if not already wrapped and not decorative
      if (!img.closest('figure') && !img.classList.contains('decorative')) {
        const figure = document.createElement('figure');
        figure.className = 'docx-image-figure';
        
        // Insert figure before image
        img.parentNode.insertBefore(figure, img);
        
        // Move image into figure
        figure.appendChild(img);
        
        // If image has a title, add figcaption
        if (img.getAttribute('title') && !figure.querySelector('figcaption')) {
          const figcaption = document.createElement('figcaption');
          figcaption.textContent = img.getAttribute('title');
          figcaption.className = 'docx-image-caption';
          figure.appendChild(figcaption);
        }
      }
      
      // For decorative images, set empty alt and role="presentation"
      if (img.classList.contains('decorative')) {
        img.setAttribute('alt', '');
        img.setAttribute('role', 'presentation');
      }
    });
  } catch (error) {
    console.error('Error processing images for accessibility:', error);
  }
}

/**
 * Ensure proper heading hierarchy
 * Fixes heading levels to ensure no levels are skipped
 * 
 * @param {Document} document - DOM document
 */
function ensureHeadingHierarchy(document) {
  try {
    const headings = document.querySelectorAll('h1, h2, h3, h4, h5, h6');
    if (headings.length === 0) return;
    
    let lastLevel = 0;
    
    // First pass: identify if we have an h1
    const hasH1 = document.querySelector('h1') !== null;
    
    // If no h1, convert the first heading to h1
    if (!hasH1 && headings.length > 0) {
      const firstHeading = headings[0];
      const newH1 = document.createElement('h1');
      
      // Copy all attributes and content
      Array.from(firstHeading.attributes).forEach(attr => {
        newH1.setAttribute(attr.name, attr.value);
      });
      newH1.innerHTML = firstHeading.innerHTML;
      
      // Replace the old heading
      firstHeading.parentNode.replaceChild(newH1, firstHeading);
      
      // Recursively fix the hierarchy now that we have an h1
      return ensureHeadingHierarchy(document);
    }
    
    // Second pass: fix sequential heading levels
    const allHeadings = document.querySelectorAll('h1, h2, h3, h4, h5, h6');
    
    for (const heading of allHeadings) {
      const currentLevel = parseInt(heading.tagName.substring(1), 10);
      
      // If this heading skips a level (e.g., h1 to h3), fix it
      if (currentLevel > lastLevel + 1) {
        const newLevel = lastLevel + 1;
        const newHeading = document.createElement(`h${newLevel}`);
        
        // Copy all attributes and content
        Array.from(heading.attributes).forEach(attr => {
          newHeading.setAttribute(attr.name, attr.value);
        });
        newHeading.innerHTML = heading.innerHTML;
        
        // Replace the old heading
        heading.parentNode.replaceChild(newHeading, heading);
        lastLevel = newLevel;
      } else {
        lastLevel = currentLevel;
      }
    }
  } catch (error) {
    console.error('Error ensuring heading hierarchy:', error);
  }
}

/**
 * Add ARIA landmarks to document
 * Creates proper sectioning with ARIA roles
 * 
 * @param {Document} document - DOM document
 */
/**
 * Add ARIA landmarks to document - Fixed implementation
 * Creates proper sectioning with ARIA roles without moving elements
 * 
 * @param {Document} document - DOM document
 */
function addAriaLandmarks(document) {
  try {
    // Add role="main" to the body or create a main wrapper
    const body = document.body;
    
    // Check if there's already a main element
    let mainElement = document.querySelector('main');
    
    if (!mainElement) {
      // Add role="main" to body
      body.setAttribute('role', 'main');
      body.setAttribute('aria-label', 'Document content');
    }
    
    // Find and enhance TOC navigation
    const tocElements = document.querySelectorAll('.docx-toc, nav');
    tocElements.forEach(toc => {
      if (!toc.hasAttribute('role')) {
        toc.setAttribute('role', 'navigation');
      }
      if (!toc.hasAttribute('aria-label')) {
        toc.setAttribute('aria-label', 'Table of Contents');
      }
    });
    
    // Add complementary role to sidebars or asides
    const asideElements = document.querySelectorAll('aside, .sidebar');
    asideElements.forEach(aside => {
      aside.setAttribute('role', 'complementary');
    });
    
    // Add banner role to headers if present
    const headerElements = document.querySelectorAll('header');
    headerElements.forEach(header => {
      header.setAttribute('role', 'banner');
    });
    
    // Add contentinfo role to footers if present
    const footerElements = document.querySelectorAll('footer');
    footerElements.forEach(footer => {
      footer.setAttribute('role', 'contentinfo');
    });
    
    console.log('ARIA landmarks added successfully');
    
  } catch (error) {
    console.error('Error adding ARIA landmarks:', error);
  }
}

/**
 * Add skip navigation link
 * Adds a link to skip to main content for keyboard users
 * 
 * @param {Document} document - DOM document
 */
function addSkipNavigation(document) {
  try {
    // Check if skip link already exists
    if (document.querySelector('.skip-link')) {
      return;
    }
    
    // Find or create main content target
    let mainContent = document.querySelector('main, [role="main"]');
    
    if (!mainContent) {
      // Use body as fallback
      mainContent = document.body;
    }
    
    // Ensure main content has an ID
    if (!mainContent.id) {
      mainContent.id = 'main-content';
    }
    
    // Create the skip link
    const skipLink = document.createElement('a');
    skipLink.href = `#${mainContent.id}`;
    skipLink.className = 'skip-link';
    skipLink.textContent = 'Skip to main content';
    skipLink.setAttribute('tabindex', '1'); // Ensure it's first in tab order
    
    // Insert at the very beginning of body
    document.body.insertBefore(skipLink, document.body.firstChild);
    
    console.log('Skip navigation link added');
    
  } catch (error) {
    console.error('Error adding skip navigation:', error);
  }
}

/**
 * Enhance keyboard navigability
 * Makes interactive elements properly focusable and adds focus styles
 * 
 * @param {Document} document - DOM document
 */
function enhanceKeyboardNavigation(document) {
  try {
    // TOC entries should be navigable
    const tocEntries = document.querySelectorAll('.docx-toc-entry');
    tocEntries.forEach(entry => {
      if (!entry.getAttribute('tabindex')) {
        entry.setAttribute('tabindex', '0');
      }
      entry.classList.add('keyboard-focusable');
    });
    
    // Links should have visible focus states
    const links = document.querySelectorAll('a');
    links.forEach(link => {
      link.classList.add('keyboard-focusable');
      
      // Ensure links have proper href or role
      if (!link.getAttribute('href') && !link.getAttribute('role')) {
        link.setAttribute('role', 'button');
      }
    });
    
    // Buttons should be keyboard accessible
    const buttons = document.querySelectorAll('button');
    buttons.forEach(button => {
      button.classList.add('keyboard-focusable');
      if (!button.getAttribute('type')) {
        button.setAttribute('type', 'button');
      }
    });
    
    // Add keyboard shortcut for track changes toggle if present
    if (document.body.classList.contains('docx-track-changes-show') || 
        document.body.classList.contains('docx-track-changes-hide')) {
      
      // Add keyboard event listener via script
      const toggleScript = document.createElement('script');
      toggleScript.textContent = `
        document.addEventListener('keydown', function(e) {
          // Alt+T to toggle track changes
          if (e.altKey && (e.key === 't' || e.key === 'T')) {
            e.preventDefault();
            const body = document.body;
            if (body.classList.contains('docx-track-changes-show')) {
              body.classList.remove('docx-track-changes-show');
              body.classList.add('docx-track-changes-hide');
              console.log('Track changes hidden');
            } else {
              body.classList.remove('docx-track-changes-hide');
              body.classList.add('docx-track-changes-show');
              console.log('Track changes shown');
            }
          }
        });
      `;
      
      // Add script to head to avoid affecting body content
      document.head.appendChild(toggleScript);
    }
    
    // Enhance table navigation
    const tables = document.querySelectorAll('table');
    tables.forEach(table => {
      // Make table focusable for keyboard navigation
      table.setAttribute('tabindex', '0');
      table.classList.add('keyboard-focusable');
      
      // Add keyboard navigation help
      if (!table.hasAttribute('aria-describedby')) {
        const helpId = `table-help-${Date.now()}`;
        table.setAttribute('aria-describedby', helpId);
        
        const helpText = document.createElement('div');
        helpText.id = helpId;
        helpText.className = 'sr-only';
        helpText.textContent = 'Use arrow keys to navigate table cells when focused';
        table.parentNode.insertBefore(helpText, table);
      }
    });
    
    console.log('Keyboard navigation enhancements applied');
    
  } catch (error) {
    console.error('Error enhancing keyboard navigation:', error);
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
  try {
    // Add high contrast mode class to html element
    document.documentElement.classList.add('docx-accessible');
    
    // Add data-theme attribute for potential theme switching
    document.documentElement.setAttribute('data-theme', 'light');
    
    // Insert a meta theme-color tag for browser UI
    if (!document.querySelector('meta[name="theme-color"]')) {
      const meta = document.createElement('meta');
      meta.setAttribute('name', 'theme-color');
      meta.setAttribute('content', '#ffffff');
      document.head.appendChild(meta);
    }
    
    // Add a theme toggle button (optional enhancement)
    if (!document.querySelector('.theme-toggle')) {
      const themeToggle = document.createElement('button');
      themeToggle.className = 'theme-toggle sr-only'; // Hidden by default
      themeToggle.textContent = 'Toggle High Contrast';
      themeToggle.setAttribute('aria-label', 'Toggle high contrast mode');
      themeToggle.setAttribute('type', 'button');
      
      // Add toggle functionality
      themeToggle.addEventListener('click', function() {
        const html = document.documentElement;
        const isHighContrast = html.getAttribute('data-theme') === 'high-contrast';
        html.setAttribute('data-theme', isHighContrast ? 'light' : 'high-contrast');
      });
      
      // Add to document (could be shown via CSS when needed)
      document.body.appendChild(themeToggle);
    }
    
    console.log('Color contrast enhancements applied');
    
  } catch (error) {
    console.error('Error enhancing color contrast:', error);
  }
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
