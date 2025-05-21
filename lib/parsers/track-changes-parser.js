// lib/parsers/track-changes-parser.js - Track changes parsing functions
const { selectNodes, selectSingleNode } = require('../xml/xpath-utils');

/**
 * Parse track changes information from document.xml
 * Extracts insertions, deletions, moves, and formatting changes
 * 
 * @param {Document} documentDoc - Document XML document
 * @returns {Object} - Track changes information
 */
function parseTrackChanges(documentDoc) {
  const changes = {
    hasTrackedChanges: false,
    insertions: [],
    deletions: [],
    moves: [],
    formattingChanges: []
  };
  
  try {
    // Check if track changes are present
    const trackChangesNodes = selectNodes("//w:ins|//w:del|//w:moveFrom|//w:moveTo|//w:rPrChange", documentDoc);
    changes.hasTrackedChanges = trackChangesNodes.length > 0;
    
    // If no tracked changes, return early
    if (!changes.hasTrackedChanges) {
      return changes;
    }
    
    // Extract insertions
    const insertionNodes = selectNodes("//w:ins", documentDoc);
    insertionNodes.forEach((node, index) => {
      const author = node.getAttribute('w:author') || 'Unknown';
      const date = node.getAttribute('w:date') || '';
      const id = node.getAttribute('w:id') || `ins-${index}`;
      
      // Get text content of the insertion
      const runNodes = selectNodes(".//w:r", node);
      let text = '';
      runNodes.forEach(run => {
        const textNodes = selectNodes(".//w:t", run);
        textNodes.forEach(t => {
          text += t.textContent || '';
        });
      });
      
      changes.insertions.push({
        id,
        author,
        date: formatDate(date),
        text
      });
    });
    
    // Extract deletions
    const deletionNodes = selectNodes("//w:del", documentDoc);
    deletionNodes.forEach((node, index) => {
      const author = node.getAttribute('w:author') || 'Unknown';
      const date = node.getAttribute('w:date') || '';
      const id = node.getAttribute('w:id') || `del-${index}`;
      
      // Get text content of the deletion
      const runNodes = selectNodes(".//w:r", node);
      let text = '';
      runNodes.forEach(run => {
        const delTextNodes = selectNodes(".//w:delText", run);
        delTextNodes.forEach(t => {
          text += t.textContent || '';
        });
      });
      
      changes.deletions.push({
        id,
        author,
        date: formatDate(date),
        text
      });
    });
    
    // Extract moves (move from location)
    const moveFromNodes = selectNodes("//w:moveFrom", documentDoc);
    moveFromNodes.forEach((node, index) => {
      const author = node.getAttribute('w:author') || 'Unknown';
      const date = node.getAttribute('w:date') || '';
      const id = node.getAttribute('w:id') || `moveFrom-${index}`;
      
      // Get text content of the move
      const runNodes = selectNodes(".//w:r", node);
      let text = '';
      runNodes.forEach(run => {
        const textNodes = selectNodes(".//w:t", run);
        textNodes.forEach(t => {
          text += t.textContent || '';
        });
      });
      
      changes.moves.push({
        id,
        author,
        date: formatDate(date),
        text,
        type: 'moveFrom'
      });
    });
    
    // Extract moves (move to location)
    const moveToNodes = selectNodes("//w:moveTo", documentDoc);
    moveToNodes.forEach((node, index) => {
      const author = node.getAttribute('w:author') || 'Unknown';
      const date = node.getAttribute('w:date') || '';
      const id = node.getAttribute('w:id') || `moveTo-${index}`;
      
      // Get text content of the move
      const runNodes = selectNodes(".//w:r", node);
      let text = '';
      runNodes.forEach(run => {
        const textNodes = selectNodes(".//w:t", run);
        textNodes.forEach(t => {
          text += t.textContent || '';
        });
      });
      
      changes.moves.push({
        id,
        author,
        date: formatDate(date),
        text,
        type: 'moveTo'
      });
    });
    
    // Extract formatting changes
    const formattingChangeNodes = selectNodes("//w:rPrChange", documentDoc);
    formattingChangeNodes.forEach((node, index) => {
      const author = node.getAttribute('w:author') || 'Unknown';
      const date = node.getAttribute('w:date') || '';
      const id = `format-${index}`;
      
      // Get the parent run to determine the affected text
      const parentRun = getParentNode(node, 'w:r');
      let text = '';
      
      if (parentRun) {
        const textNodes = selectNodes(".//w:t", parentRun);
        textNodes.forEach(t => {
          text += t.textContent || '';
        });
      }
      
      // Extract the formatting properties that changed
      const formattingProps = [];
      
      // Check for bold change
      if (selectSingleNode(".//w:b", node)) {
        formattingProps.push('bold');
      }
      
      // Check for italic change
      if (selectSingleNode(".//w:i", node)) {
        formattingProps.push('italic');
      }
      
      // Check for underline change
      if (selectSingleNode(".//w:u", node)) {
        formattingProps.push('underline');
      }
      
      // Check for font size change
      const szNode = selectSingleNode(".//w:sz", node);
      if (szNode) {
        const size = szNode.getAttribute('w:val');
        if (size) {
          formattingProps.push(`size: ${parseInt(size, 10) / 2}pt`);
        }
      }
      
      // Check for font family change
      const fontNode = selectSingleNode(".//w:rFonts", node);
      if (fontNode) {
        const font = fontNode.getAttribute('w:ascii') || fontNode.getAttribute('w:hAnsi');
        if (font) {
          formattingProps.push(`font: ${font}`);
        }
      }
      
      // Check for color change
      const colorNode = selectSingleNode(".//w:color", node);
      if (colorNode) {
        const color = colorNode.getAttribute('w:val');
        if (color) {
          formattingProps.push(`color: #${color}`);
        }
      }
      
      changes.formattingChanges.push({
        id,
        author,
        date: formatDate(date),
        text,
        properties: formattingProps
      });
    });
    
  } catch (error) {
    console.error('Error parsing track changes:', error);
  }
  
  return changes;
}

/**
 * Process track changes in HTML document
 * Applies visual indicators for track changes in the HTML
 * 
 * @param {Document} document - DOM document
 * @param {Object} changes - Track changes information
 * @param {Object} options - Processing options
 * @returns {Document} - Enhanced document
 */
function processTrackChanges(document, changes, options = {}) {
  // Default options
  const defaultOptions = {
    mode: 'show', // 'show', 'hide', 'accept', or 'reject'
    showAuthor: true,
    showDate: true,
    highlightColor: '#E6F4FF', // Light blue for insertions
    deletionColor: '#FFEBEE'   // Light red for deletions
  };
  
  const opts = { ...defaultOptions, ...options };
  
  try {
    // If no tracked changes, return the document unchanged
    if (!changes.hasTrackedChanges) {
      return document;
    }
    
    // Add track changes mode class to body
    document.body.classList.add(`docx-track-changes-${opts.mode}`);
    
    // Process track changes based on the mode
    if (opts.mode === 'hide') {
      // In hide mode, we don't show any tracked changes
      return document;
    }
    
    // Add a track changes legend if there are changes to show
    if (opts.mode === 'show' && 
       (changes.insertions.length > 0 || 
        changes.deletions.length > 0 || 
        changes.moves.length > 0 || 
        changes.formattingChanges.length > 0)) {
      
      addTrackChangesLegend(document);
    }
    
    // For 'show' or 'accept' mode, process insertions
    if ((opts.mode === 'show' || opts.mode === 'accept') && changes.insertions.length > 0) {
      processInsertions(document, changes.insertions, opts);
    }
    
    // For 'show' or 'reject' mode, process deletions
    if ((opts.mode === 'show' || opts.mode === 'reject') && changes.deletions.length > 0) {
      processDeletions(document, changes.deletions, opts);
    }
    
    // For 'show' mode, process moves and formatting changes
    if (opts.mode === 'show') {
      if (changes.moves.length > 0) {
        processMoves(document, changes.moves, opts);
      }
      
      if (changes.formattingChanges.length > 0) {
        processFormattingChanges(document, changes.formattingChanges, opts);
      }
    }
    
  } catch (error) {
    console.error('Error processing track changes:', error);
  }
  
  return document;
}

/**
 * Process insertions
 * Applies visual styling to inserted content
 * 
 * @param {Document} document - DOM document
 * @param {Array} insertions - Insertion changes
 * @param {Object} options - Processing options
 */
function processInsertions(document, insertions, options) {
  // Find all elements with data-change-id attribute for insertions
  insertions.forEach(insertion => {
    const elements = document.querySelectorAll(`[data-change-id="${insertion.id}"]`);
    
    if (elements.length === 0) {
      // If no elements found, try to find text content that matches
      // This is a fallback approach for when direct IDs are not available
      const textNodes = findTextNodes(document.body, insertion.text);
      
      textNodes.forEach(node => {
        // Wrap the found text in a span for styling
        const span = document.createElement('span');
        span.classList.add('docx-insertion');
        span.setAttribute('data-change-id', insertion.id);
        
        if (options.showAuthor && insertion.author) {
          span.setAttribute('data-author', insertion.author);
        }
        
        if (options.showDate && insertion.date) {
          span.setAttribute('data-date', insertion.date);
        }
        
        // Set background color
        span.style.backgroundColor = options.highlightColor;
        
        // Replace the text node with the span
        node.parentNode.insertBefore(span, node);
        span.appendChild(node);
      });
    } else {
      // Style the found elements
      elements.forEach(element => {
        if (options.mode === 'show') {
          element.classList.add('docx-insertion');
          
          if (options.showAuthor && insertion.author) {
            element.setAttribute('data-author', insertion.author);
          }
          
          if (options.showDate && insertion.date) {
            element.setAttribute('data-date', insertion.date);
          }
          
          // Set background color
          element.style.backgroundColor = options.highlightColor;
        } else if (options.mode === 'accept') {
          // For 'accept' mode, remove the track changes marking but keep the content
          element.classList.remove('docx-insertion');
          element.removeAttribute('data-author');
          element.removeAttribute('data-date');
          element.removeAttribute('data-change-id');
          element.style.backgroundColor = '';
        }
      });
    }
  });
}

/**
 * Process deletions
 * Applies visual styling to deleted content
 * 
 * @param {Document} document - DOM document
 * @param {Array} deletions - Deletion changes
 * @param {Object} options - Processing options
 */
function processDeletions(document, deletions, options) {
  // Find all elements with data-change-id attribute for deletions
  deletions.forEach(deletion => {
    const elements = document.querySelectorAll(`[data-change-id="${deletion.id}"]`);
    
    if (elements.length === 0 && options.mode === 'show') {
      // In 'show' mode, if no elements found, we need to create elements for deleted content
      // Find a suitable place to insert the deletion (e.g., before the main content)
      const mainContent = document.querySelector('main') || document.body;
      
      // Create a deleted content container if it doesn't exist
      let deletedContent = document.querySelector('.docx-deleted-content');
      if (!deletedContent) {
        deletedContent = document.createElement('div');
        deletedContent.className = 'docx-deleted-content';
        deletedContent.setAttribute('aria-label', 'Deleted content');
        deletedContent.setAttribute('role', 'complementary');
        
        // Insert at the start of main content
        mainContent.insertBefore(deletedContent, mainContent.firstChild);
      }
      
      // Create element for this deletion
      const deletionElement = document.createElement('del');
      deletionElement.className = 'docx-deletion';
      deletionElement.setAttribute('data-change-id', deletion.id);
      
      if (options.showAuthor && deletion.author) {
        deletionElement.setAttribute('data-author', deletion.author);
      }
      
      if (options.showDate && deletion.date) {
        deletionElement.setAttribute('data-date', deletion.date);
      }
      
      // Style the deletion
      deletionElement.style.backgroundColor = options.deletionColor;
      deletionElement.style.textDecoration = 'line-through';
      
      // Set the deleted text
      deletionElement.textContent = deletion.text;
      
      // Add to the deleted content container
      deletedContent.appendChild(deletionElement);
    } else {
      // Style the found elements
      elements.forEach(element => {
        if (options.mode === 'show') {
          element.classList.add('docx-deletion');
          
          if (options.showAuthor && deletion.author) {
            element.setAttribute('data-author', deletion.author);
          }
          
          if (options.showDate && deletion.date) {
            element.setAttribute('data-date', deletion.date);
          }
          
          // Style the deletion
          element.style.backgroundColor = options.deletionColor;
          element.style.textDecoration = 'line-through';
        } else if (options.mode === 'reject') {
          // For 'reject' mode, remove the track changes marking but keep the content
          element.classList.remove('docx-deletion');
          element.removeAttribute('data-author');
          element.removeAttribute('data-date');
          element.removeAttribute('data-change-id');
          element.style.backgroundColor = '';
          element.style.textDecoration = '';
        }
      });
    }
  });
}

/**
 * Process moves
 * Applies visual styling to moved content
 * 
 * @param {Document} document - DOM document
 * @param {Array} moves - Move changes
 * @param {Object} options - Processing options
 */
function processMoves(document, moves, options) {
  // Find all elements with data-change-id attribute for moves
  moves.forEach(move => {
    const elements = document.querySelectorAll(`[data-change-id="${move.id}"]`);
    
    elements.forEach(element => {
      // Style based on move type
      if (move.type === 'moveFrom') {
        element.classList.add('docx-move-from');
        element.style.borderBottom = '2px dashed #9575CD'; // Purple for move from
      } else {
        element.classList.add('docx-move-to');
        element.style.borderBottom = '2px solid #9575CD'; // Purple for move to
      }
      
      if (options.showAuthor && move.author) {
        element.setAttribute('data-author', move.author);
      }
      
      if (options.showDate && move.date) {
        element.setAttribute('data-date', move.date);
      }
    });
  });
}

/**
 * Process formatting changes
 * Applies visual styling to formatting changes
 * 
 * @param {Document} document - DOM document
 * @param {Array} formattingChanges - Formatting changes
 * @param {Object} options - Processing options
 */
function processFormattingChanges(document, formattingChanges, options) {
  // Find all elements with data-change-id attribute for formatting changes
  formattingChanges.forEach(change => {
    const elements = document.querySelectorAll(`[data-change-id="${change.id}"]`);
    
    if (elements.length === 0) {
      // If no elements found, try to find text content that matches
      const textNodes = findTextNodes(document.body, change.text);
      
      textNodes.forEach(node => {
        // Wrap the found text in a span for styling
        const span = document.createElement('span');
        span.classList.add('docx-formatting-change');
        span.setAttribute('data-change-id', change.id);
        
        if (options.showAuthor && change.author) {
          span.setAttribute('data-author', change.author);
        }
        
        if (options.showDate && change.date) {
          span.setAttribute('data-date', change.date);
        }
        
        if (change.properties && change.properties.length > 0) {
          span.setAttribute('data-formatting', change.properties.join(', '));
        }
        
        // Add subtle styling
        span.style.borderBottom = '1px dotted #FFA000'; // Amber for formatting changes
        
        // Replace the text node with the span
        node.parentNode.insertBefore(span, node);
        span.appendChild(node);
      });
    } else {
      // Style the found elements
      elements.forEach(element => {
        element.classList.add('docx-formatting-change');
        
        if (options.showAuthor && change.author) {
          element.setAttribute('data-author', change.author);
        }
        
        if (options.showDate && change.date) {
          element.setAttribute('data-date', change.date);
        }
        
        if (change.properties && change.properties.length > 0) {
          element.setAttribute('data-formatting', change.properties.join(', '));
        }
        
        // Add subtle styling
        element.style.borderBottom = '1px dotted #FFA000'; // Amber for formatting changes
      });
    }
  });
}

/**
 * Add track changes legend
 * Adds a legend explaining the track changes markings
 * 
 * @param {Document} document - DOM document
 */
function addTrackChangesLegend(document) {
  // Create legend element
  const legend = document.createElement('div');
  legend.className = 'docx-track-changes-legend';
  legend.setAttribute('role', 'complementary');
  legend.setAttribute('aria-label', 'Track changes legend');
  
  // Style the legend
  legend.style.margin = '1em 0';
  legend.style.padding = '0.5em';
  legend.style.border = '1px solid #BDBDBD';
  legend.style.borderRadius = '4px';
  legend.style.backgroundColor = '#F5F5F5';
  
  // Create toggle button
  const toggleButton = document.createElement('button');
  toggleButton.textContent = 'Toggle Track Changes';
  toggleButton.className = 'docx-track-changes-toggle';
  toggleButton.setAttribute('aria-controls', 'docx-track-changes');
  toggleButton.setAttribute('type', 'button');
  
  // Style the button
  toggleButton.style.padding = '0.25em 0.5em';
  toggleButton.style.marginRight = '1em';
  toggleButton.style.border = '1px solid #BDBDBD';
  toggleButton.style.borderRadius = '3px';
  toggleButton.style.cursor = 'pointer';
  
  // Add click event to toggle
  toggleButton.addEventListener('click', function() {
    const body = document.body;
    if (body.classList.contains('docx-track-changes-show')) {
      body.classList.remove('docx-track-changes-show');
      body.classList.add('docx-track-changes-hide');
    } else {
      body.classList.add('docx-track-changes-show');
      body.classList.remove('docx-track-changes-hide');
    }
  });
  
  // Create legend title
  const title = document.createElement('h3');
  title.textContent = 'Track Changes';
  title.style.margin = '0 0 0.5em 0';
  title.style.display = 'inline-block';
  
  // Create legend items container
  const itemsContainer = document.createElement('div');
  itemsContainer.id = 'docx-track-changes';
  itemsContainer.style.display = 'flex';
  itemsContainer.style.flexWrap = 'wrap';
  itemsContainer.style.gap = '1em';
  
  // Create legend items for different change types
  const insertionItem = createLegendItem('Insertion', '#E6F4FF');
  const deletionItem = createLegendItem('Deletion', '#FFEBEE', 'line-through');
  const moveItem = createLegendItem('Move', 'transparent', 'none', '2px dashed #9575CD');
  const formatItem = createLegendItem('Format Change', 'transparent', 'none', '1px dotted #FFA000');
  
  // Add items to container
  itemsContainer.appendChild(insertionItem);
  itemsContainer.appendChild(deletionItem);
  itemsContainer.appendChild(moveItem);
  itemsContainer.appendChild(formatItem);
  
  // Assemble the legend
  legend.appendChild(toggleButton);
  legend.appendChild(title);
  legend.appendChild(itemsContainer);
  
  // Add keyboard shortcut info
  const shortcutInfo = document.createElement('p');
  shortcutInfo.textContent = 'Keyboard shortcut: Alt+T to toggle track changes';
  shortcutInfo.style.margin = '0.5em 0 0 0';
  shortcutInfo.style.fontSize = '0.9em';
  shortcutInfo.style.color = '#616161';
  legend.appendChild(shortcutInfo);
  
  // Insert at the beginning of the document
  const main = document.querySelector('main') || document.body;
  main.insertBefore(legend, main.firstChild);
}

/**
 * Create legend item
 * Helper function to create a track changes legend item
 * 
 * @param {string} text - Item text
 * @param {string} bgColor - Background color
 * @param {string} textDecoration - Text decoration
 * @param {string} borderBottom - Border style
 * @returns {HTMLElement} - Legend item element
 */
function createLegendItem(text, bgColor, textDecoration = 'none', borderBottom = 'none') {
  const item = document.createElement('div');
  item.style.display = 'flex';
  item.style.alignItems = 'center';
  
  const sample = document.createElement('span');
  sample.textContent = 'Sample';
  sample.style.backgroundColor = bgColor;
  sample.style.textDecoration = textDecoration;
  sample.style.borderBottom = borderBottom;
  sample.style.padding = '0.1em 0.3em';
  sample.style.marginRight = '0.5em';
  
  const label = document.createElement('span');
  label.textContent = text;
  
  item.appendChild(sample);
  item.appendChild(label);
  
  return item;
}

/**
 * Find text nodes matching content
 * Helper function to find text nodes with specific content
 * 
 * @param {Node} root - Root node to search in
 * @param {string} text - Text to find
 * @returns {Array} - Array of matching text nodes
 */
function findTextNodes(root, text) {
  if (!text) return [];
  
  const matches = [];
  const walker = document.createTreeWalker(
    root,
    NodeFilter.SHOW_TEXT,
    {
      acceptNode: function(node) {
        // Ignore empty text nodes and those that don't contain our text
        if (node.nodeValue.trim() === '' || !node.nodeValue.includes(text)) {
          return NodeFilter.FILTER_REJECT;
        }
        return NodeFilter.FILTER_ACCEPT;
      }
    }
  );
  
  let node;
  while (node = walker.nextNode()) {
    matches.push(node);
  }
  
  return matches;
}

/**
 * Get parent node by name
 * Helper function to get the parent node of a specific type
 * 
 * @param {Node} node - Child node
 * @param {string} parentName - Parent node name
 * @returns {Node|null} - Parent node or null if not found
 */
function getParentNode(node, parentName) {
  let parent = node.parentNode;
  
  while (parent) {
    if (parent.nodeName === parentName) {
      return parent;
    }
    parent = parent.parentNode;
  }
  
  return null;
}

/**
 * Format date string
 * Formats DOCX date string to human-readable format
 * 
 * @param {string} dateString - DOCX date string
 * @returns {string} - Formatted date string
 */
function formatDate(dateString) {
  if (!dateString) return '';
  
  try {
    const date = new Date(dateString);
    return date.toLocaleString();
  } catch (error) {
    return dateString;
  }
}

module.exports = {
  parseTrackChanges,
  processTrackChanges,
  processInsertions,
  processDeletions,
  processMoves,
  processFormattingChanges,
  addTrackChangesLegend,
  findTextNodes,
  getParentNode,
  formatDate
};
