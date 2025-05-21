// lib/parsers/metadata-parser.js - Document metadata parsing functions
const { selectSingleNode, selectNodes } = require('../xml/xpath-utils');

/**
 * Parse document metadata from core.xml and app.xml
 * Extracts title, author, creation date, and other metadata
 * 
 * @param {Document} corePropsDoc - Core properties XML document
 * @param {Document} appPropsDoc - Application properties XML document
 * @returns {Object} - Document metadata
 */
function parseDocumentMetadata(corePropsDoc, appPropsDoc) {
  const metadata = {
    core: {},
    app: {}
  };
  
  try {
    // Extract core properties if available
    if (corePropsDoc) {
      // Define namespaces for core properties
      const coreNS = {
        'dc': 'http://purl.org/dc/elements/1.1/',
        'dcterms': 'http://purl.org/dc/terms/',
        'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties'
      };
      
      // Dublin Core elements
      metadata.core.title = getNodeText(corePropsDoc, "//dc:title", coreNS) || '';
      metadata.core.subject = getNodeText(corePropsDoc, "//dc:subject", coreNS) || '';
      metadata.core.creator = getNodeText(corePropsDoc, "//dc:creator", coreNS) || '';
      metadata.core.description = getNodeText(corePropsDoc, "//dc:description", coreNS) || '';
      
      // Package core properties
      metadata.core.keywords = getNodeText(corePropsDoc, "//cp:keywords", coreNS) || '';
      metadata.core.category = getNodeText(corePropsDoc, "//cp:category", coreNS) || '';
      metadata.core.lastModifiedBy = getNodeText(corePropsDoc, "//cp:lastModifiedBy", coreNS) || '';
      metadata.core.contentStatus = getNodeText(corePropsDoc, "//cp:contentStatus", coreNS) || '';
      metadata.core.revision = getNodeText(corePropsDoc, "//cp:revision", coreNS) || '';
      
      // Dublin Core terms
      metadata.core.created = getNodeText(corePropsDoc, "//dcterms:created", coreNS) || '';
      metadata.core.modified = getNodeText(corePropsDoc, "//dcterms:modified", coreNS) || '';
      
      // Format dates properly if they exist
      if (metadata.core.created) {
        metadata.core.created = formatW3CDate(metadata.core.created);
      }
      
      if (metadata.core.modified) {
        metadata.core.modified = formatW3CDate(metadata.core.modified);
      }
    }
    
    // Extract application properties if available
    if (appPropsDoc) {
      // App properties are in the default namespace, no specific NS required
      metadata.app.application = getNodeText(appPropsDoc, "//Application") || '';
      metadata.app.appVersion = getNodeText(appPropsDoc, "//AppVersion") || '';
      metadata.app.company = getNodeText(appPropsDoc, "//Company") || '';
      metadata.app.manager = getNodeText(appPropsDoc, "//Manager") || '';
      
      // Document statistics
      metadata.app.totalTime = getNodeText(appPropsDoc, "//TotalTime") || '';
      metadata.app.pages = getNodeText(appPropsDoc, "//Pages") || '';
      metadata.app.words = getNodeText(appPropsDoc, "//Words") || '';
      metadata.app.characters = getNodeText(appPropsDoc, "//Characters") || '';
      metadata.app.characterWithSpaces = getNodeText(appPropsDoc, "//CharactersWithSpaces") || '';
      metadata.app.paragraphs = getNodeText(appPropsDoc, "//Paragraphs") || '';
      metadata.app.lines = getNodeText(appPropsDoc, "//Lines") || '';
      
      // Template information
      metadata.app.template = getNodeText(appPropsDoc, "//Template") || '';
    }
  } catch (error) {
    console.error('Error parsing document metadata:', error);
  }
  
  return metadata;
}

/**
 * Apply metadata to HTML document
 * Adds metadata as meta tags in the document head
 * 
 * @param {Document} document - DOM document
 * @param {Object} metadata - Document metadata
 * @returns {Document} - Enhanced document
 */
function applyMetadataToHtml(document, metadata) {
  try {
    const head = document.head;
    
    // Add title if available
    if (metadata.core.title) {
      let titleElement = head.querySelector('title');
      if (!titleElement) {
        titleElement = document.createElement('title');
        head.appendChild(titleElement);
      }
      titleElement.textContent = metadata.core.title;
    }
    
    // Add standard meta tags
    addMetaTag(head, 'description', metadata.core.description);
    addMetaTag(head, 'author', metadata.core.creator);
    addMetaTag(head, 'keywords', metadata.core.keywords);
    addMetaTag(head, 'category', metadata.core.category);
    
    // Add created and modified dates
    if (metadata.core.created) {
      addMetaTag(head, 'date', metadata.core.created);
    }
    
    if (metadata.core.modified) {
      addMetaTag(head, 'last-modified', metadata.core.modified);
    }
    
    // Add Dublin Core metadata
    addDublinCoreMetadata(head, metadata.core);
    
    // Add Open Graph metadata
    addOpenGraphMetadata(head, metadata.core);
    
    // Add Twitter Card metadata
    addTwitterCardMetadata(head, metadata.core);
    
    // Add JSON-LD structured data
    addJsonLdStructuredData(head, metadata);
    
    // Add generator tag
    addMetaTag(head, 'generator', `${metadata.app.application || 'Microsoft Word'} - Converted by doc2web`);
    
    // Add document statistics to comments at the end of head
    const statsComment = document.createComment(`
      Document Statistics:
      - Pages: ${metadata.app.pages || 'Unknown'}
      - Words: ${metadata.app.words || 'Unknown'}
      - Characters: ${metadata.app.characters || 'Unknown'}
      - Paragraphs: ${metadata.app.paragraphs || 'Unknown'}
      - Lines: ${metadata.app.lines || 'Unknown'}
    `);
    head.appendChild(statsComment);
    
  } catch (error) {
    console.error('Error applying metadata to HTML:', error);
  }
  
  return document;
}

/**
 * Add Dublin Core metadata
 * Adds Dublin Core metadata elements to the document head
 * 
 * @param {HTMLElement} head - Head element
 * @param {Object} coreMetadata - Core metadata
 */
function addDublinCoreMetadata(head, coreMetadata) {
  // Add Dublin Core metadata
  if (coreMetadata.title) {
    addMetaTag(head, 'DC.title', coreMetadata.title);
  }
  
  if (coreMetadata.creator) {
    addMetaTag(head, 'DC.creator', coreMetadata.creator);
  }
  
  if (coreMetadata.subject) {
    addMetaTag(head, 'DC.subject', coreMetadata.subject);
  }
  
  if (coreMetadata.description) {
    addMetaTag(head, 'DC.description', coreMetadata.description);
  }
  
  if (coreMetadata.created) {
    addMetaTag(head, 'DC.date', coreMetadata.created);
  }
  
  // Add DC.rights if copyright info is available
  // This is a simplified approach; real implementation would check for copyright info
  addMetaTag(head, 'DC.rights', 'All rights reserved');
  
  // Add DC.type for document type
  addMetaTag(head, 'DC.type', 'Text');
  
  // Add DC.format for format
  addMetaTag(head, 'DC.format', 'text/html');
}

/**
 * Add Open Graph metadata
 * Adds Open Graph protocol metadata for social sharing
 * 
 * @param {HTMLElement} head - Head element
 * @param {Object} coreMetadata - Core metadata
 */
function addOpenGraphMetadata(head, coreMetadata) {
  // Basic Open Graph meta tags
  addMetaTag(head, 'og:type', 'article');
  
  // Use title if available
  if (coreMetadata.title) {
    addMetaTag(head, 'og:title', coreMetadata.title);
  }
  
  // Use description if available
  if (coreMetadata.description) {
    addMetaTag(head, 'og:description', coreMetadata.description);
  }
  
  // Add article specific tags
  if (coreMetadata.created) {
    addMetaTag(head, 'article:published_time', coreMetadata.created);
  }
  
  if (coreMetadata.modified) {
    addMetaTag(head, 'article:modified_time', coreMetadata.modified);
  }
  
  if (coreMetadata.creator) {
    addMetaTag(head, 'article:author', coreMetadata.creator);
  }
  
  // Add tags from keywords if available
  if (coreMetadata.keywords) {
    const keywords = coreMetadata.keywords.split(',');
    keywords.forEach(keyword => {
      if (keyword.trim()) {
        addMetaTag(head, 'article:tag', keyword.trim());
      }
    });
  }
}

/**
 * Add Twitter Card metadata
 * Adds Twitter Card metadata for Twitter sharing
 * 
 * @param {HTMLElement} head - Head element
 * @param {Object} coreMetadata - Core metadata
 */
function addTwitterCardMetadata(head, coreMetadata) {
  // Basic Twitter Card meta tags
  addMetaTag(head, 'twitter:card', 'summary');
  
  // Use title if available
  if (coreMetadata.title) {
    addMetaTag(head, 'twitter:title', coreMetadata.title);
  }
  
  // Use description if available
  if (coreMetadata.description) {
    addMetaTag(head, 'twitter:description', coreMetadata.description);
  }
  
  // Add creator if available
  if (coreMetadata.creator) {
    addMetaTag(head, 'twitter:creator', coreMetadata.creator);
  }
}

/**
 * Add JSON-LD structured data
 * Adds schema.org structured data as JSON-LD
 * 
 * @param {HTMLElement} head - Head element
 * @param {Object} metadata - Document metadata
 */
function addJsonLdStructuredData(head, metadata) {
  // Create basic JSON-LD for Article type
  const jsonLd = {
    '@context': 'https://schema.org',
    '@type': 'Article',
    'headline': metadata.core.title || 'Document',
    'description': metadata.core.description || '',
    'keywords': metadata.core.keywords || '',
    'datePublished': metadata.core.created || '',
    'dateModified': metadata.core.modified || '',
    'wordCount': metadata.app.words || '',
    'author': {
      '@type': 'Person',
      'name': metadata.core.creator || ''
    }
  };
  
  // Add publisher if company is available
  if (metadata.app.company) {
    jsonLd.publisher = {
      '@type': 'Organization',
      'name': metadata.app.company
    };
  }
  
  // Create script element for JSON-LD
  const script = document.createElement('script');
  script.type = 'application/ld+json';
  script.textContent = JSON.stringify(jsonLd, null, 2);
  head.appendChild(script);
}

/**
 * Add meta tag to head
 * Helper function to create and add meta tags
 * 
 * @param {HTMLElement} head - Head element
 * @param {string} name - Meta name or property
 * @param {string} content - Meta content
 */
function addMetaTag(head, name, content) {
  if (!content) return;
  
  const meta = document.createElement('meta');
  
  // Determine if this is a property (og:, twitter:) or name
  if (name.startsWith('og:') || name.startsWith('article:') || name.startsWith('twitter:')) {
    meta.setAttribute('property', name);
  } else {
    meta.setAttribute('name', name);
  }
  
  meta.setAttribute('content', content);
  head.appendChild(meta);
}

/**
 * Get text content of an XML node
 * Helper function to extract text from XML nodes
 * 
 * @param {Document} doc - XML document
 * @param {string} xpath - XPath expression
 * @param {Object} namespaces - Optional namespaces object
 * @returns {string|null} - Text content or null if not found
 */
function getNodeText(doc, xpath, namespaces = null) {
  try {
    let node;
    if (namespaces) {
      // Create a custom selectSingleNode function with the specified namespaces
      const xpath = require('xpath');
      const select = xpath.useNamespaces(namespaces);
      node = select(xpath, doc, true);
    } else {
      // Use the standard selectSingleNode
      node = selectSingleNode(xpath, doc);
    }
    
    return node ? node.textContent : null;
  } catch (error) {
    console.error(`Error getting node text for ${xpath}:`, error);
    return null;
  }
}

/**
 * Format W3C date string
 * Converts W3C date format to ISO date format
 * 
 * @param {string} dateString - W3C date string
 * @returns {string} - Formatted ISO date string
 */
function formatW3CDate(dateString) {
  try {
    // Remove any timezone offset information for consistency
    const dateWithoutTZ = dateString.replace(/Z|[+-]\d{2}:\d{2}$/, '');
    
    // Create a date object and format as ISO string
    const date = new Date(dateWithoutTZ);
    return date.toISOString();
  } catch (error) {
    console.error(`Error formatting date ${dateString}:`, error);
    return dateString;
  }
}

module.exports = {
  parseDocumentMetadata,
  applyMetadataToHtml,
  addDublinCoreMetadata,
  addOpenGraphMetadata,
  addTwitterCardMetadata,
  addJsonLdStructuredData,
  addMetaTag,
  getNodeText,
  formatW3CDate
};
