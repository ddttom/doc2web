// lib/parsers/metadata-parser.js - Document metadata parsing functions
const { selectSingleNode, selectNodes } = require('../xml/xpath-utils');
const xpath = require('xpath');

/**
 * Parse document metadata from core.xml and app.xml
 * Extracts title, author, creation date, and other metadata
 * 
 * @param {Document} corePropsDoc - Core properties XML document
 * @param {Document} appPropsDoc - Application properties XML document
 * @param {Document} documentDoc - Optional document XML for calculating statistics
 * @returns {Object} - Document metadata
 */
function parseDocumentMetadata(corePropsDoc, appPropsDoc, documentDoc) {
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
      
      // Create namespace-aware selector
      const select = xpath.useNamespaces(coreNS);
      
      // Dublin Core elements
      metadata.core.title = getNodeTextWithNS(corePropsDoc, "//dc:title", select) || '';
      metadata.core.subject = getNodeTextWithNS(corePropsDoc, "//dc:subject", select) || '';
      metadata.core.creator = getNodeTextWithNS(corePropsDoc, "//dc:creator", select) || '';
      metadata.core.description = getNodeTextWithNS(corePropsDoc, "//dc:description", select) || '';
      
      // Package core properties
      metadata.core.keywords = getNodeTextWithNS(corePropsDoc, "//cp:keywords", select) || '';
      metadata.core.category = getNodeTextWithNS(corePropsDoc, "//cp:category", select) || '';
      metadata.core.lastModifiedBy = getNodeTextWithNS(corePropsDoc, "//cp:lastModifiedBy", select) || '';
      metadata.core.contentStatus = getNodeTextWithNS(corePropsDoc, "//cp:contentStatus", select) || '';
      metadata.core.revision = getNodeTextWithNS(corePropsDoc, "//cp:revision", select) || '';
      
      // Dublin Core terms
      metadata.core.created = getNodeTextWithNS(corePropsDoc, "//dcterms:created", select) || '';
      metadata.core.modified = getNodeTextWithNS(corePropsDoc, "//dcterms:modified", select) || '';
      
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
      metadata.app.totalTime = getNodeText(appPropsDoc, "//TotalTime") || '0';
      metadata.app.pages = getNodeText(appPropsDoc, "//Pages") || '0';
      metadata.app.words = getNodeText(appPropsDoc, "//Words") || '0';
      metadata.app.characters = getNodeText(appPropsDoc, "//Characters") || '0';
      metadata.app.characterWithSpaces = getNodeText(appPropsDoc, "//CharactersWithSpaces") || '0';
      metadata.app.paragraphs = getNodeText(appPropsDoc, "//Paragraphs") || '0';
      metadata.app.lines = getNodeText(appPropsDoc, "//Lines") || '0';
      
      // If document statistics are not available and documentDoc is provided, calculate them
      if (documentDoc && 
          (metadata.app.words === '0' || 
           metadata.app.characters === '0' || 
           metadata.app.paragraphs === '0')) {
        try {
          // Calculate document statistics from content
          const paragraphNodes = selectNodes("//w:p", documentDoc);
          const textNodes = selectNodes("//w:t", documentDoc);
          
          // Count paragraphs
          if (metadata.app.paragraphs === '0' && paragraphNodes.length > 0) {
            metadata.app.paragraphs = paragraphNodes.length.toString();
          }
          
          // Extract all text content
          let textContent = '';
          textNodes.forEach(t => {
            textContent += t.textContent || '';
          });
          
          // Count words
          if (metadata.app.words === '0' && textContent) {
            const words = textContent.trim().split(/\s+/).filter(word => word.length > 0);
            metadata.app.words = words.length.toString();
          }
          
          // Count characters
          if (metadata.app.characters === '0' && textContent) {
            metadata.app.characters = textContent.replace(/\s/g, '').length.toString();
          }
          
          // Count characters with spaces
          if (metadata.app.characterWithSpaces === '0' && textContent) {
            metadata.app.characterWithSpaces = textContent.length.toString();
          }
          
          // Estimate lines (rough approximation)
          if (metadata.app.lines === '0' && textContent) {
            // Assuming average line length of 80 characters
            metadata.app.lines = Math.ceil(textContent.length / 80).toString();
          }
          
          // Estimate pages (rough approximation)
          if (metadata.app.pages === '0' && textContent) {
            // Assuming average page has 3000 characters
            metadata.app.pages = Math.ceil(textContent.length / 3000).toString();
          }
          
          console.log('Calculated document statistics from content');
        } catch (error) {
          console.error('Error calculating document statistics:', error);
        }
      }
      
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
 * @param {Object} options - Configuration options
 * @returns {Document} - Enhanced document
 */
function applyMetadataToHtml(document, metadata, options = {}) {
  try {
    if (!document || !document.head) {
      console.error('Invalid document or missing head element');
      return document;
    }
    
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
    addMetaTag(head, document, 'description', metadata.core.description);
    
    // Configure author metadata based on options
    const authorConfig = options.metadata || {};
    let authorName = 'doc2web'; // Default fallback
    
    if (authorConfig.preserveOriginalAuthor && metadata.core.creator) {
      authorName = metadata.core.creator;
    } else if (authorConfig.overrideAuthor) {
      authorName = authorConfig.overrideAuthor;
    } else if (authorConfig.useDoc2webAuthor !== false) {
      authorName = 'doc2web';
    }
    
    addMetaTag(head, document, 'author', authorName);
    addMetaTag(head, document, 'keywords', metadata.core.keywords);
    addMetaTag(head, document, 'category', metadata.core.category);
    
    // Add created and modified dates
    if (metadata.core.created) {
      addMetaTag(head, document, 'date', metadata.core.created);
    }
    
    if (metadata.core.modified) {
      addMetaTag(head, document, 'last-modified', metadata.core.modified);
    }
    
    // Add Dublin Core metadata
    addDublinCoreMetadata(head, document, metadata.core, options);
    
    // Add Open Graph metadata
    addOpenGraphMetadata(head, document, metadata.core, options);
    
    // Add Twitter Card metadata
    addTwitterCardMetadata(head, document, metadata.core, options);
    
    // Add JSON-LD structured data
    addJsonLdStructuredData(head, document, metadata, options);
    
    // Add generator tag
    addMetaTag(head, document, 'generator', `${metadata.app.application || 'Microsoft Word'} - Converted by doc2web`);
    
    // Add document statistics to comments at the end of head
    const statsComment = document.createComment(`
      Document Statistics:
      - Pages: ${metadata.app.pages}
      - Words: ${metadata.app.words}
      - Characters: ${metadata.app.characters}
      - Paragraphs: ${metadata.app.paragraphs}
      - Lines: ${metadata.app.lines}
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
 * @param {Document} doc - Document object
 * @param {Object} coreMetadata - Core metadata
 * @param {Object} options - Configuration options
 */
function addDublinCoreMetadata(head, doc, coreMetadata, options = {}) {
  // Add Dublin Core metadata
  if (coreMetadata.title) {
    addMetaTag(head, doc, 'DC.title', coreMetadata.title);
  }
  
  // Configure creator metadata based on options
  const authorConfig = options.metadata || {};
  let creatorName = 'doc2web'; // Default fallback
  
  if (authorConfig.preserveOriginalAuthor && coreMetadata.creator) {
    creatorName = coreMetadata.creator;
  } else if (authorConfig.overrideAuthor) {
    creatorName = authorConfig.overrideAuthor;
  } else if (authorConfig.useDoc2webAuthor !== false) {
    creatorName = 'doc2web';
  }
  
  addMetaTag(head, doc, 'DC.creator', creatorName);
  
  if (coreMetadata.subject) {
    addMetaTag(head, doc, 'DC.subject', coreMetadata.subject);
  }
  
  if (coreMetadata.description) {
    addMetaTag(head, doc, 'DC.description', coreMetadata.description);
  }
  
  if (coreMetadata.created) {
    addMetaTag(head, doc, 'DC.date', coreMetadata.created);
  }
  
  // Add DC.rights if copyright info is available
  // This is a simplified approach; real implementation would check for copyright info
  addMetaTag(head, doc, 'DC.rights', 'All rights reserved');
  
  // Add DC.type for document type
  addMetaTag(head, doc, 'DC.type', 'Text');
  
  // Add DC.format for format
  addMetaTag(head, doc, 'DC.format', 'text/html');
}

/**
 * Add Open Graph metadata
 * Adds Open Graph protocol metadata for social sharing
 * 
 * @param {HTMLElement} head - Head element
 * @param {Document} doc - Document object
 * @param {Object} coreMetadata - Core metadata
 * @param {Object} options - Configuration options
 */
function addOpenGraphMetadata(head, doc, coreMetadata, options = {}) {
  // Basic Open Graph meta tags
  addMetaTag(head, doc, 'og:type', 'article');
  
  // Use title if available
  if (coreMetadata.title) {
    addMetaTag(head, doc, 'og:title', coreMetadata.title);
  }
  
  // Use description if available
  if (coreMetadata.description) {
    addMetaTag(head, doc, 'og:description', coreMetadata.description);
  }
  
  // Add article specific tags
  if (coreMetadata.created) {
    addMetaTag(head, doc, 'article:published_time', coreMetadata.created);
  }
  
  if (coreMetadata.modified) {
    addMetaTag(head, doc, 'article:modified_time', coreMetadata.modified);
  }
  
  // Configure article author metadata based on options
  const authorConfig = options.metadata || {};
  let articleAuthor = 'doc2web'; // Default fallback
  
  if (authorConfig.preserveOriginalAuthor && coreMetadata.creator) {
    articleAuthor = coreMetadata.creator;
  } else if (authorConfig.overrideAuthor) {
    articleAuthor = authorConfig.overrideAuthor;
  } else if (authorConfig.useDoc2webAuthor !== false) {
    articleAuthor = 'doc2web';
  }
  
  addMetaTag(head, doc, 'article:author', articleAuthor);
  
  // Add tags from keywords if available
  if (coreMetadata.keywords) {
    const keywords = coreMetadata.keywords.split(',');
    keywords.forEach(keyword => {
      if (keyword.trim()) {
        addMetaTag(head, doc, 'article:tag', keyword.trim());
      }
    });
  }
}

/**
 * Add Twitter Card metadata
 * Adds Twitter Card metadata for Twitter sharing
 * 
 * @param {HTMLElement} head - Head element
 * @param {Document} doc - Document object
 * @param {Object} coreMetadata - Core metadata
 * @param {Object} options - Configuration options
 */
function addTwitterCardMetadata(head, doc, coreMetadata, options = {}) {
  // Basic Twitter Card meta tags
  addMetaTag(head, doc, 'twitter:card', 'summary');
  
  // Use title if available
  if (coreMetadata.title) {
    addMetaTag(head, doc, 'twitter:title', coreMetadata.title);
  }
  
  // Use description if available
  if (coreMetadata.description) {
    addMetaTag(head, doc, 'twitter:description', coreMetadata.description);
  }
  
  // Configure twitter creator metadata based on options
  const authorConfig = options.metadata || {};
  let twitterCreator = 'doc2web'; // Default fallback
  
  if (authorConfig.preserveOriginalAuthor && coreMetadata.creator) {
    twitterCreator = coreMetadata.creator;
  } else if (authorConfig.overrideAuthor) {
    twitterCreator = authorConfig.overrideAuthor;
  } else if (authorConfig.useDoc2webAuthor !== false) {
    twitterCreator = 'doc2web';
  }
  
  addMetaTag(head, doc, 'twitter:creator', twitterCreator);
}

/**
 * Add JSON-LD structured data
 * Adds schema.org structured data as JSON-LD
 * 
 * @param {HTMLElement} head - Head element
 * @param {Document} doc - Document object
 * @param {Object} metadata - Document metadata
 * @param {Object} options - Configuration options
 */
function addJsonLdStructuredData(head, doc, metadata, options = {}) {
  // Create basic JSON-LD for Article type
  const jsonLd = {
    '@context': 'https://schema.org',
    '@type': 'Article',
    'headline': metadata.core.title || 'Document',
    'description': metadata.core.description || '',
    'keywords': metadata.core.keywords || '',
    'datePublished': metadata.core.created || '',
    'dateModified': metadata.core.modified || '',
    'wordCount': metadata.app.words,
    'author': {
      '@type': 'Person',
      'name': (() => {
        const authorConfig = options.metadata || {};
        if (authorConfig.preserveOriginalAuthor && metadata.core.creator) {
          return metadata.core.creator;
        } else if (authorConfig.overrideAuthor) {
          return authorConfig.overrideAuthor;
        } else {
          return 'doc2web';
        }
      })()
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
  const script = doc.createElement('script');
  script.type = 'application/ld+json';
  script.textContent = JSON.stringify(jsonLd, null, 2);
  head.appendChild(script);
}

/**
 * Add meta tag to head
 * Helper function to create and add meta tags
 * 
 * @param {HTMLElement} head - Head element
 * @param {Document} doc - Document object
 * @param {string} name - Meta name or property
 * @param {string} content - Meta content
 */
function addMetaTag(head, doc, name, content) {
  if (!content) return;
  
  const meta = doc.createElement('meta');
  
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
 * Get text content of an XML node using namespace-aware XPath
 * Helper function to extract text from XML nodes with namespaces
 * 
 * @param {Document} doc - XML document
 * @param {string} xpathExpr - XPath expression
 * @param {Function} select - Namespace-aware selector
 * @returns {string|null} - Text content or null if not found
 */
function getNodeTextWithNS(doc, xpathExpr, select) {
  try {
    const node = select(xpathExpr, doc, true);
    return node ? node.textContent : null;
  } catch (error) {
    console.error(`Error getting node text for ${xpathExpr}:`, error);
    return null;
  }
}

/**
 * Get text content of an XML node
 * Helper function to extract text from XML nodes
 * 
 * @param {Document} doc - XML document
 * @param {string} xpath - XPath expression
 * @returns {string|null} - Text content or null if not found
 */
function getNodeText(doc, xpath) {
  try {
    const node = selectSingleNode(xpath, doc);
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
