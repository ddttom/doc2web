// lib/html/content-processors.js - Content element processors (headings, TOC, lists)
const { selectNodes, selectSingleNode } = require('../xml/xpath-utils');

/**
 * Enhanced heading processing with better numbering
 * Adds numbered prefixes to headings and structures them hierarchically
 * 
 * @param {Document} document - DOM document
 * @param {Object} styleInfo - Style information
 */
function processHeadings(document, styleInfo) {
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
    
    // Skip TOC headings
    if (heading.classList.contains('docx-toc-heading')) {
      continue;
    }
    
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
      // Use number.letter format for h2
      const letter = String.fromCharCode(96 + counters.h2); // 'a', 'b', 'c', etc.
      prefix = `${counters.h1}.${letter}.`;
    } else if (level === 3) {
      prefix = `${counters.h1}.${counters.h2}.${counters.h3}.`;
    }
    
    // Text already starts with the prefix? Skip adding it again
    if (heading.textContent.startsWith(prefix)) {
      continue;
    }
    
    // Store the prefix for potential use in CSS
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

/**
 * Enhanced TOC processing for better structure and styling
 * Creates a properly structured Table of Contents with leader lines and page numbers
 * 
 * @param {Document} document - DOM document
 * @param {Object} styleInfo - Style information
 */
function processTOC(document, styleInfo) {
  // Find TOC entries which might be paragraphs with TOC classes
  let tocEntries = document.querySelectorAll('.docx-toc-entry, p[class*="toc"]');
  
  if (tocEntries.length === 0) {
    // Try to find any paragraphs that might be TOC entries based on content pattern
    const paragraphs = document.querySelectorAll('p');
    let tocCandidates = [];
    
    // Look for paragraphs with page numbers at the end
    for (let i = 0; i < paragraphs.length; i++) {
      const p = paragraphs[i];
      const text = p.textContent.trim();
      
      // Check for the pattern: text followed by dots/spaces, then digits at the end
      const pageNumMatch = text.match(/^(.*?)[\.\s]+(\d+)$/);
      
      if (pageNumMatch) {
        p.classList.add('docx-toc-entry');
        tocCandidates.push(p);
      }
    }
    
    // If we found potential TOC entries, process them
    if (tocCandidates.length > 0) {
      tocEntries = tocCandidates;
    }
  }
  
  // Create TOC container if we have entries
  if (tocEntries.length > 0) {
    // Check if there's already a TOC container
    let tocContainer = document.querySelector('.docx-toc');
    
    if (!tocContainer) {
      // Create a container for all TOC entries
      tocContainer = document.createElement('div');
      tocContainer.className = 'docx-toc';
      
      // Check if there's an existing TOC heading
      const tocHeading = document.querySelector('.docx-toc-heading');
      
      if (!tocHeading) {
        // Create a TOC heading if needed
        const heading = document.createElement('h2');
        heading.className = 'docx-toc-heading';
        heading.textContent = 'Table of Contents';
        tocContainer.appendChild(heading);
      } else {
        // Move existing TOC heading
        tocContainer.appendChild(tocHeading.cloneNode(true));
        tocHeading.parentNode.removeChild(tocHeading);
      }
      
      // Insert the TOC container before the first TOC entry
      const firstEntry = tocEntries[0];
      firstEntry.parentNode.insertBefore(tocContainer, firstEntry);
    }
    
    // Process each TOC entry
    Array.from(tocEntries).forEach((entry, index) => {
      // Skip if already processed
      if (entry.classList.contains('processed-toc-entry')) {
        return;
      }
      
      // Mark as processed
      entry.classList.add('processed-toc-entry');
      
      // Determine TOC level from class name or calculate from indentation
      let level = 1;
      
      // Try to get level from class
      const classArray = Array.from(entry.classList);
      for (const cls of classArray) {
        if (cls.includes('toc') && cls.match(/\d+$/)) {
          level = parseInt(cls.match(/\d+$/)[0], 10);
          break;
        }
      }
      
      // Add level class if not present
      const levelClass = `docx-toc-level-${level}`;
      if (!entry.classList.contains(levelClass)) {
        entry.classList.add(levelClass);
      }
      
      // Get the text content
      const text = entry.textContent.trim();
      
      // Check for the pattern: text followed by dots/spaces, then digits at the end
      const pageNumMatch = text.match(/^(.*?)[\.\s]*(\d+)$/);
      
      if (pageNumMatch) {
        // Clear existing content
        entry.textContent = '';
        
        // Create text span
        const textSpan = document.createElement('span');
        textSpan.className = 'docx-toc-text';
        textSpan.textContent = pageNumMatch[1].trim();
        entry.appendChild(textSpan);
        
        // Create dots span
        const dotsSpan = document.createElement('span');
        dotsSpan.className = 'docx-toc-dots';
        entry.appendChild(dotsSpan);
        
        // Create page number span
        const pageSpan = document.createElement('span');
        pageSpan.className = 'docx-toc-pagenum';
        pageSpan.textContent = pageNumMatch[2];
        entry.appendChild(pageSpan);
      }
      
      // Move this entry to the TOC container if it's not already there
      if (entry.parentNode !== tocContainer) {
        const clonedEntry = entry.cloneNode(true);
        tocContainer.appendChild(clonedEntry);
        entry.parentNode.removeChild(entry);
      }
    });
  }
}

/**
 * Identify list patterns in the document
 * Analyzes paragraphs to identify common list patterns
 * 
 * @param {Array} paragraphs - Array of paragraph elements
 * @returns {Object} - Identified patterns
 */
function identifyListPatterns(paragraphs) {
  const patterns = {
    numberedItems: [],
    alphaItems: [],
    romanItems: [],
    specialParagraphs: []
  };
  
  // Counters for pattern detection
  const prefixCounts = {};
  const patternCounts = {};
  
  // First pass - count patterns
  paragraphs.forEach(p => {
    const text = p.textContent.trim();
    
    // Skip empty paragraphs
    if (!text) return;
    
    // Look for common list item patterns
    const numberMatch = text.match(/^\s*(\d+)[\.\)\s]\s+/);
    const alphaMatch = text.match(/^\s*([a-z])[\.\)]\s+/i);
    const romanMatch = text.match(/^\s*([ivxlcdm]+)[\.\)]\s+/i);
    
    // Count prefixes
    if (numberMatch) {
      const prefix = 'number';
      prefixCounts[prefix] = (prefixCounts[prefix] || 0) + 1;
    } else if (alphaMatch) {
      const prefix = 'alpha';
      prefixCounts[prefix] = (prefixCounts[prefix] || 0) + 1;
    } else if (romanMatch) {
      const prefix = 'roman';
      prefixCounts[prefix] = (prefixCounts[prefix] || 0) + 1;
    }
    
    // Look for special paragraph patterns (generic approach, not hardcoded)
    const forMatch = text.match(/^(\w+)\s+for\s+/i);
    const commaMatch = text.match(/^(\w+),\s/);
    const parenMatch = text.match(/^(\w+)\s+\(([^)]+)\)/);
    
    if (forMatch) {
      const pattern = 'word_for_word';
      patternCounts[pattern] = patternCounts[pattern] || { count: 0, examples: [] };
      patternCounts[pattern].count++;
      if (patternCounts[pattern].examples.length < 3) {
        patternCounts[pattern].examples.push(text);
      }
    } else if (commaMatch) {
      const pattern = 'word_comma';
      patternCounts[pattern] = patternCounts[pattern] || { count: 0, examples: [] };
      patternCounts[pattern].count++;
      if (patternCounts[pattern].examples.length < 3) {
        patternCounts[pattern].examples.push(text);
      }
    } else if (parenMatch) {
      const pattern = 'word_parenthesis';
      patternCounts[pattern] = patternCounts[pattern] || { count: 0, examples: [] };
      patternCounts[pattern].count++;
      if (patternCounts[pattern].examples.length < 3) {
        patternCounts[pattern].examples.push(text);
      }
    }
  });
  
  // Identify special paragraph patterns that appear multiple times (not hardcoded specific terms)
  Object.entries(patternCounts).forEach(([pattern, info]) => {
    if (info.count >= 2) {
      patterns.specialParagraphs.push({
        type: pattern,
        count: info.count,
        examples: info.examples
      });
    }
  });
  
  return patterns;
}

/**
 * Check if a paragraph is a special type based on patterns
 * Identifies special paragraph types based on content patterns
 * 
 * @param {string} text - Paragraph text
 * @param {Array} specialPatterns - Special paragraph patterns
 * @returns {Object|null} - Special paragraph info or null
 */
function isSpecialParagraph(text, specialPatterns) {
  // Check each pattern type (not hardcoded)
  const forMatch = text.match(/^(\w+)\s+for\s+/i);
  if (forMatch) {
    return { type: 'word_for_word', prefix: forMatch[1] };
  }
  
  const commaMatch = text.match(/^(\w+),\s/);
  if (commaMatch) {
    return { type: 'word_comma', prefix: commaMatch[1] };
  }
  
  const parenMatch = text.match(/^(\w+)\s+\(([^)]+)\)/);
  if (parenMatch) {
    return { type: 'word_parenthesis', prefix: parenMatch[1] };
  }
  
  return null;
}

/**
 * Process special paragraph types based on document analysis
 * Applies special styling to paragraphs based on content patterns
 * 
 * @param {Document} document - DOM document
 * @param {Object} styleInfo - Style information
 */
function processSpecialParagraphs(document, styleInfo) {
  // Get special paragraph patterns from document analysis
  const specialPatterns = styleInfo.documentStructure?.specialParagraphPatterns || [];
  
  // Process paragraphs for special formatting
  const paragraphs = document.querySelectorAll('p');
  
  paragraphs.forEach(para => {
    // Skip paragraphs that are already processed as list items or TOC entries
    if (para.classList.contains('processed-list-item') || 
        para.classList.contains('processed-special-para') ||
        para.classList.contains('docx-toc-entry')) {
      return;
    }
    
    const text = para.textContent.trim();
    
    // Check for generic patterns (not hardcoded to specific terms)
    
    // Pattern 1: Word followed by "for" followed by more text
    if (text.match(/^(\w+)\s+for\s+/i)) {
      para.classList.add('docx-word_for_word');
      para.classList.add('docx-special-pattern');
    }
    
    // Pattern 2: Word followed by comma
    if (text.match(/^(\w+),\s/)) {
      para.classList.add('docx-word_comma');
      para.classList.add('docx-special-pattern');
    }
    
    // Pattern 3: Word followed by parenthesized text
    if (text.match(/^(\w+)\s+\(([^)]+)\)/)) {
      para.classList.add('docx-word_parenthesis');
      para.classList.add('docx-special-pattern');
    }
    
    // Add classes based on document structure analysis
    specialPatterns.forEach(pattern => {
      if (pattern.type) {
        // Check if this paragraph matches the pattern's examples
        // We use a more generic approach here to avoid hardcoding specific terms
        if (pattern.examples.some(example => {
          const exampleFirstWord = example.split(' ')[0];
          return text.startsWith(exampleFirstWord);
        })) {
          para.classList.add(`docx-${pattern.type}`);
          para.classList.add('docx-special-pattern');
        }
      }
    });
  });
}

/**
 * Enhanced numbered paragraph processing for better hierarchical structure
 * Converts paragraphs with numbering into properly structured lists
 * 
 * @param {Document} document - DOM document
 * @param {Object} styleInfo - Style information
 */
function processNestedNumberedParagraphs(document, styleInfo) {
  // Process paragraphs that look like they should be part of a numbered list
  const paragraphs = Array.from(document.querySelectorAll('p'));
  
  // Skip paragraphs that are already in lists
  const filteredParagraphs = paragraphs.filter(p => !p.closest('ol, ul'));
  
  // Track list structure
  let currentLists = {}; // Use map to track different list hierarchies
  let lastMainNumbers = {}; // Keep track of last number in each list
  
  // First pass: identify list types and patterns
  const listPatterns = identifyListPatterns(filteredParagraphs);
  
  // Regex patterns for different types of numbering - more comprehensive
  const mainNumberPatterns = [
    /^\s*(\d+)\.\s+(.+)$/,                  // 1. Text
    /^\s*(\d+)\)\s+(.+)$/,                  // 1) Text
    /^\s*(\d+)\s+(.+)$/                     // 1 Text (no punctuation)
  ];
  
  const alphaPatterns = [
    /^\s*([a-z])\.\s+(.+)$/,                // a. Text
    /^\s*([a-z])\)\s+(.+)$/,                // a) Text
    /^\s*([A-Z])\.\s+(.+)$/,                // A. Text
    /^\s*([A-Z])\)\s+(.+)$/                 // A) Text
  ];
  
  const romanPatterns = [
    /^\s*([ivxlcdm]+)\.\s+(.+)$/i,          // i. or I. Text
    /^\s*([ivxlcdm]+)\)\s+(.+)$/i           // i) or I) Text
  ];
  
  const subNumberPatterns = [
    /^\s*(\d+\.\d+)\.\s+(.+)$/,             // 1.1. Text
    /^\s*(\d+\.\d+)\)\s+(.+)$/,             // 1.1) Text
    /^\s*(\d+\.\d+)\s+(.+)$/                // 1.1 Text
  ];
  
  // Generic pattern for items with a word prefix followed by parenthesized text
  const parenthesesPattern = /^\s*([\w]+)\s+\(([^)]+)\)\s*(.+)$/;
  
  // Process each paragraph
  for (let i = 0; i < filteredParagraphs.length; i++) {
    const p = filteredParagraphs[i];
    const text = p.textContent.trim();
    
    // Skip already processed paragraphs
    if (p.classList.contains('processed-list-item')) {
      continue;
    }
    
    // Skip TOC entries
    if (p.classList.contains('docx-toc-entry')) {
      continue;
    }
    
    // Try to match with list patterns
    let isListItem = false;
    let listType = null;
    let prefix = null;
    let content = null;
    let level = 0;
    let match = null;
    
    // Check for main numbered items
    for (const pattern of mainNumberPatterns) {
      match = text.match(pattern);
      if (match) {
        isListItem = true;
        listType = 'main';
        prefix = match[1];
        content = match[2];
        level = 1;
        break;
      }
    }
    
    // Check for alpha sub-items
    if (!isListItem) {
      for (const pattern of alphaPatterns) {
        match = text.match(pattern);
        if (match) {
          isListItem = true;
          listType = 'alpha';
          prefix = match[1];
          content = match[2];
          level = 2;
          break;
        }
      }
    }
    
    // Check for Roman numeral items
    if (!isListItem) {
      for (const pattern of romanPatterns) {
        match = text.match(pattern);
        if (match) {
          isListItem = true;
          listType = 'roman';
          prefix = match[1];
          content = match[2];
          level = match[1].toUpperCase() === match[1] ? 1 : 2; // Uppercase = main, lowercase = sub
          break;
        }
      }
    }
    
    // Check for sub-numbered items (1.1, 1.2, etc.)
    if (!isListItem) {
      for (const pattern of subNumberPatterns) {
        match = text.match(pattern);
        if (match) {
          isListItem = true;
          listType = 'sub';
          prefix = match[1];
          content = match[2];
          level = 2;
          break;
        }
      }
    }
    
    // Check for word + parentheses pattern (generic approach)
    if (!isListItem) {
      match = text.match(parenthesesPattern);
      if (match) {
        isListItem = true;
        listType = 'wordparen';
        prefix = match[2];
        content = match[1] + ' (' + match[2] + ') ' + match[3];
        level = 1;
      }
    }
    
    // Process the list item if found
    if (isListItem) {
      // Mark as processed
      p.classList.add('processed-list-item');
      
      // Determine the list hierarchy ID
      // For main lists, use a unique ID; for sublist, use the parent's ID
      let listHierarchyId;
      
      if (listType === 'main' || listType === 'wordparen') {
        // Start a new list or continue an existing one
        listHierarchyId = 'list-' + listType;
        
        // Check if we need to create a new list or use an existing one
        if (!currentLists[listHierarchyId] || 
            !document.body.contains(currentLists[listHierarchyId]) ||
            // Check if this is a new sequence (for numbered lists)
            (listType === 'main' && parseInt(prefix, 10) === 1 && lastMainNumbers[listHierarchyId] > 1)) {
          
          // Create a new list
          const list = document.createElement('ol');
          list.className = 'docx-numbered-list';
          if (listType === 'wordparen') {
            list.classList.add('docx-wordparen-list');
          }
          p.parentNode.insertBefore(list, p);
          currentLists[listHierarchyId] = list;
          lastMainNumbers[listHierarchyId] = 0;
        }
        
        // Update last number
        if (listType === 'main') {
          lastMainNumbers[listHierarchyId] = parseInt(prefix, 10);
        }
      } else {
        // This is a sublist item, it should be under the most recent main list
        listHierarchyId = 'list-main';
      }
      
      // Get the current list
      const currentList = currentLists[listHierarchyId];
      
      if (!currentList) {
        // If we don't have a main list yet but have a sublist item,
        // create a main list first
        const list = document.createElement('ol');
        list.className = 'docx-numbered-list';
        p.parentNode.insertBefore(list, p);
        currentLists['list-main'] = list;
      }
      
      // Now process the list item
      if (listType === 'main' || listType === 'wordparen') {
        // This is a main list item
        const listItem = document.createElement('li');
        listItem.textContent = content;
        listItem.setAttribute('data-prefix', prefix);
        currentLists[listHierarchyId].appendChild(listItem);
        
        // Store this list item as the current main item for potential sublists
        currentLists['current-main-item'] = listItem;
        
        // Remove the original paragraph
        p.parentNode.removeChild(p);
        
      } else {
        // This is a sublist item
        
        // Get the current main item
        const currentMainItem = currentLists['current-main-item'];
        
        if (currentMainItem) {
          // Look for an existing sublist
          let sublist = currentMainItem.querySelector('ol');
          
          if (!sublist) {
            // Create a new sublist
            sublist = document.createElement('ol');
            sublist.className = listType === 'alpha' ? 'docx-alpha-list' : 
                             listType === 'roman' ? 'docx-roman-list' : 
                             'docx-numbered-list';
            currentMainItem.appendChild(sublist);
          }
          
          // Create the sublist item
          const subItem = document.createElement('li');
          subItem.textContent = content;
          subItem.setAttribute('data-prefix', prefix);
          sublist.appendChild(subItem);
          
          // Remove the original paragraph
          p.parentNode.removeChild(p);
          
        } else {
          // No main item found, treat as normal paragraph
          // This might happen if the document structure is unusual
          console.warn('Found sublist item without main item:', text);
        }
      }
      
    } else {
      // Not a list item - check if it's a special paragraph that belongs in the current list
      
      // Try to identify special paragraphs between list items (without hardcoding specific terms)
      const isSpecial = isSpecialParagraph(text, listPatterns.specialParagraphs);
      
      if (isSpecial && currentLists['list-main']) {
        // Create a special paragraph in the list
        const specialPara = document.createElement('p');
        specialPara.className = 'docx-list-special-paragraph';
        if (isSpecial.type) {
          specialPara.classList.add(`docx-${isSpecial.type}`);
        }
        specialPara.textContent = text;
        
        // Add it after the current main item
        const currentMainItem = currentLists['current-main-item'];
        if (currentMainItem) {
          // Insert after the main item
          if (currentMainItem.nextSibling) {
            currentLists['list-main'].insertBefore(specialPara, currentMainItem.nextSibling);
          } else {
            currentLists['list-main'].appendChild(specialPara);
          }
        } else {
          // No current main item, add to the end of the list
          currentLists['list-main'].appendChild(specialPara);
        }
        
        // Remove the original paragraph
        p.parentNode.removeChild(p);
        
        // Mark as processed
        p.classList.add('processed-special-para');
      }
    }
  }
}

module.exports = {
  processHeadings,
  processTOC,
  processNestedNumberedParagraphs,
  identifyListPatterns,
  isSpecialParagraph,
  processSpecialParagraphs
};