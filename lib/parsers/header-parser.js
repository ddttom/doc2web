// lib/parsers/header-parser.js - Document header extraction and processing
const { selectSingleNode, selectNodes } = require('../xml/xpath-utils');
const { DOMParser } = require('xmldom');

/**
 * Extract document header from DOCX content
 * Extracts header content from header XML files and document structure
 * 
 * @param {Object} zip - JSZip instance with DOCX content
 * @param {Document} documentDoc - Document XML
 * @param {Document} styleDoc - Style XML document
 * @param {Object} styleInfo - Parsed style information
 * @returns {Object} - Header information with content and formatting
 */
async function extractDocumentHeader(zip, documentDoc, styleDoc, styleInfo) {
  const headerInfo = {
    hasHeaderContent: false,
    headerParagraphs: [],
    headerStyles: {},
    headerImages: [],
    headerType: null // 'first', 'even', 'default'
  };

  try {
    // Strategy 1: Extract from header XML files
    const headerFiles = ['word/header1.xml', 'word/header2.xml', 'word/header3.xml'];
    
    for (const headerFile of headerFiles) {
      const headerXml = await zip.file(headerFile)?.async("string");
      if (headerXml) {
        console.log(`Found header file: ${headerFile}`);
        const headerDoc = new DOMParser().parseFromString(headerXml);
        const headerContent = await extractHeaderFromXml(headerDoc, styleInfo, headerFile);
        
        if (headerContent.paragraphs.length > 0) {
          headerInfo.hasHeaderContent = true;
          
          // Adjust paragraph indices for images to account for combining multiple header files
          const currentParagraphCount = headerInfo.headerParagraphs.length;
          const adjustedImages = headerContent.images.map(img => ({
            ...img,
            paragraphIndex: img.paragraphIndex + currentParagraphCount
          }));
          
          headerInfo.headerParagraphs.push(...headerContent.paragraphs);
          headerInfo.headerImages.push(...adjustedImages);
          
          // Determine header type from file name
          if (headerFile.includes('header1')) {
            headerInfo.headerType = 'default';
          } else if (headerFile.includes('header2')) {
            headerInfo.headerType = 'first';
          } else if (headerFile.includes('header3')) {
            headerInfo.headerType = 'even';
          }
        }
      }
    }
    
    // Strategy 2: If no header files, look for header-like content at the top of the document
    if (!headerInfo.hasHeaderContent) {
      const documentHeaderContent = extractHeaderFromDocument(documentDoc, styleInfo);
      if (documentHeaderContent.paragraphs.length > 0) {
        headerInfo.hasHeaderContent = true;
        headerInfo.headerParagraphs.push(...documentHeaderContent.paragraphs);
        headerInfo.headerImages.push(...documentHeaderContent.images);
        headerInfo.headerType = 'document';
      }
    }
    
    console.log(`Found ${headerInfo.headerParagraphs.length} header paragraphs`);
    
  } catch (error) {
    console.error('Error extracting document header:', error);
  }
  
  return headerInfo;
}

/**
 * Extract header content from header XML document
 * 
 * @param {Document} headerDoc - Header XML document
 * @param {Object} styleInfo - Parsed style information
 * @param {string} headerFile - Header file name for context
 * @returns {Object} - Header content with paragraphs and images
 */
async function extractHeaderFromXml(headerDoc, styleInfo, headerFile) {
  const headerContent = {
    paragraphs: [],
    images: []
  };
  
  try {
    // Get all paragraphs in the header
    const paragraphNodes = selectNodes("//w:p", headerDoc);
    
    paragraphNodes.forEach((paragraph, index) => {
      const headerParagraph = analyzeHeaderParagraph(paragraph, styleInfo, index);
      
      if (headerParagraph.content.trim() || headerParagraph.hasImages) {
        headerContent.paragraphs.push(headerParagraph);
        
        // Extract any images in this paragraph
        const images = extractImagesFromParagraph(paragraph, index);
        // Add paragraph formatting context to images
        images.forEach(img => {
          img.paragraphFormatting = headerParagraph.formatting;
        });
        headerContent.images.push(...images);
      }
    });
    
  } catch (error) {
    console.error(`Error extracting header from ${headerFile}:`, error);
  }
  
  return headerContent;
}

/**
 * Extract header-like content from the main document
 * Looks for content at the top that might be header material
 * 
 * @param {Document} documentDoc - Document XML
 * @param {Object} styleInfo - Parsed style information
 * @returns {Object} - Header content with paragraphs and images
 */
function extractHeaderFromDocument(documentDoc, styleInfo) {
  const headerContent = {
    paragraphs: [],
    images: []
  };
  
  try {
    // Look at the first few paragraphs of the document
    const documentParagraphs = selectNodes("//w:body/w:p", documentDoc);
    const candidateParagraphs = documentParagraphs.slice(0, Math.min(5, documentParagraphs.length));
    
    for (let i = 0; i < candidateParagraphs.length; i++) {
      const paragraph = candidateParagraphs[i];
      const headerCandidate = analyzeHeaderParagraph(paragraph, styleInfo, i);
      
      // Check if this looks like header content
      if (isLikelyHeaderContent(headerCandidate, i)) {
        headerContent.paragraphs.push(headerCandidate);
        
        // Extract any images in this paragraph
        const images = extractImagesFromParagraph(paragraph, index);
        headerContent.images.push(...images);
      } else if (headerContent.paragraphs.length > 0) {
        // Stop looking once we find non-header content
        break;
      }
    }
    
  } catch (error) {
    console.error('Error extracting header from document:', error);
  }
  
  return headerContent;
}

/**
 * Analyze a paragraph to determine if it's header content
 * 
 * @param {Element} paragraph - Paragraph XML element
 * @param {Object} styleInfo - Parsed style information
 * @param {number} position - Position of paragraph (0-based)
 * @returns {Object} - Analysis result
 */
function analyzeHeaderParagraph(paragraph, styleInfo, position) {
  const result = {
    isHeader: false,
    isEmpty: false,
    styleId: null,
    styleInfo: null,
    content: '',
    formatting: {},
    hasImages: false,
    position: position
  };
  
  try {
    // Get paragraph style
    const pStyleNode = selectSingleNode("w:pPr/w:pStyle", paragraph);
    const styleId = pStyleNode ? pStyleNode.getAttribute('w:val') : null;
    
    // Get text content
    const textNodes = selectNodes(".//w:t", paragraph);
    let textContent = '';
    textNodes.forEach(t => {
      textContent += t.textContent || '';
    });
    
    result.content = textContent.trim();
    result.styleId = styleId;
    
    // Filter out page numbering text (e.g., "Page 1 of 5", "Page xx of xx", "1 of 5", etc.)
    if (result.content) {
      const trimmedContent = result.content.trim();
      const pageNumberPatterns = [
        /^page\s+\d+\s+of\s+\d+$/i,           // "Page 1 of 5"
        /^page\s+\d+$/i,                      // "Page 1"
        /^\d+\s+of\s+\d+$/i,                  // "1 of 5"
        /^page\s+\w+\s+of\s+\w+$/i,          // "Page xx of xx"
        /^\w+\s+of\s+\w+$/i                   // "xx of xx"
      ];
      
      if (pageNumberPatterns.some(pattern => pattern.test(trimmedContent))) {
        result.content = '';
        result.isEmpty = true;
        return result; // Skip page numbering content
      }
    }
    
    result.isEmpty = result.content.length === 0;
    
    // Check for images
    const imageNodes = selectNodes(".//w:drawing | .//w:pict", paragraph);
    result.hasImages = imageNodes.length > 0;
    
    // Extract formatting information
    result.formatting = extractParagraphFormatting(paragraph, styleInfo, styleId);
    
    if (styleId && styleInfo.styles && styleInfo.styles.paragraph && styleInfo.styles.paragraph[styleId]) {
      result.styleInfo = styleInfo.styles.paragraph[styleId];
    }
    
    // Determine if this is likely header content
    result.isHeader = isLikelyHeaderContent(result, position);
    
  } catch (error) {
    console.error('Error analyzing header paragraph:', error);
  }
  
  return result;
}

/**
 * Determine if content is likely to be header material
 * 
 * @param {Object} paragraphInfo - Paragraph analysis result
 * @param {number} position - Position in document
 * @returns {boolean} - True if likely header content
 */
function isLikelyHeaderContent(paragraphInfo, position) {
  // Empty paragraphs with images are likely header content
  if (paragraphInfo.hasImages) {
    return true;
  }
  
  // Skip empty paragraphs without images
  if (paragraphInfo.isEmpty) {
    return false;
  }
  
  // Check for header-like style names
  if (paragraphInfo.styleId) {
    const lowerStyleId = paragraphInfo.styleId.toLowerCase();
    if (lowerStyleId.includes('header') || 
        lowerStyleId.includes('title') ||
        lowerStyleId.includes('caption')) {
      return true;
    }
  }
  
  // Check for header-like formatting
  const formatting = paragraphInfo.formatting;
  let headerScore = 0;
  
  // Position-based scoring (earlier = more likely to be header)
  if (position === 0) headerScore += 20;
  else if (position === 1) headerScore += 15;
  else if (position === 2) headerScore += 10;
  else if (position > 3) headerScore -= 10;
  
  // Formatting-based scoring
  if (formatting.isCentered) headerScore += 15;
  if (formatting.isBold) headerScore += 10;
  if (formatting.isLarge) headerScore += 10;
  
  // Content-based scoring
  if (paragraphInfo.content.length > 0) {
    // Shorter text more likely to be header
    if (paragraphInfo.content.length < 100) headerScore += 10;
    else if (paragraphInfo.content.length > 200) headerScore -= 10;
    
    // All caps might indicate header
    if (paragraphInfo.content === paragraphInfo.content.toUpperCase() && 
        paragraphInfo.content.length > 3) {
      headerScore += 10;
    }
  }
  
  return headerScore >= 25; // Threshold for considering it header content
}

/**
 * Extract formatting information from a paragraph
 * 
 * @param {Element} paragraph - Paragraph XML element
 * @param {Object} styleInfo - Parsed style information
 * @param {string} styleId - Style ID of the paragraph
 * @returns {Object} - Formatting information
 */
function extractParagraphFormatting(paragraph, styleInfo, styleId) {
  const formatting = {
    isBold: false,
    isItalic: false,
    isLarge: false,
    isCentered: false,
    fontSize: null,
    fontFamily: null,
    color: null,
    alignment: null
  };
  
  try {
    // Check paragraph properties
    const pPr = selectSingleNode("w:pPr", paragraph);
    if (pPr) {
      // Check alignment
      const jcNode = selectSingleNode("w:jc", pPr);
      if (jcNode) {
        const alignment = jcNode.getAttribute('w:val');
        formatting.alignment = alignment;
        formatting.isCentered = alignment === 'center';
      }
    }
    
    // Check run properties for formatting
    const runs = selectNodes(".//w:r", paragraph);
    runs.forEach(run => {
      const rPr = selectSingleNode("w:rPr", run);
      if (rPr) {
        // Check for bold
        if (selectSingleNode("w:b", rPr)) {
          formatting.isBold = true;
        }
        
        // Check for italic
        if (selectSingleNode("w:i", rPr)) {
          formatting.isItalic = true;
        }
        
        // Check font size
        const szNode = selectSingleNode("w:sz", rPr);
        if (szNode) {
          const size = parseInt(szNode.getAttribute('w:val'), 10);
          if (size) {
            formatting.fontSize = size / 2; // Convert half-points to points
            formatting.isLarge = formatting.fontSize > 14;
          }
        }
        
        // Check font family
        const rFontsNode = selectSingleNode("w:rFonts", rPr);
        if (rFontsNode) {
          formatting.fontFamily = rFontsNode.getAttribute('w:ascii') || 
                                  rFontsNode.getAttribute('w:hAnsi') ||
                                  rFontsNode.getAttribute('w:cs');
        }
        
        // Check color
        const colorNode = selectSingleNode("w:color", rPr);
        if (colorNode) {
          formatting.color = colorNode.getAttribute('w:val');
        }
      }
    });
    
    // Also check style-based formatting
    if (styleId && styleInfo.styles && styleInfo.styles.paragraph && styleInfo.styles.paragraph[styleId]) {
      const style = styleInfo.styles.paragraph[styleId];
      
      if (style.fontSize && !formatting.fontSize) {
        formatting.fontSize = parseFloat(style.fontSize);
        formatting.isLarge = formatting.fontSize > 14;
      }
      
      if (style.fontFamily && !formatting.fontFamily) {
        formatting.fontFamily = style.fontFamily;
      }
      
      if (style.textAlign && !formatting.alignment) {
        formatting.alignment = style.textAlign;
        formatting.isCentered = style.textAlign === 'center';
      }
      
      if (style.fontWeight === 'bold') {
        formatting.isBold = true;
      }
      
      if (style.fontStyle === 'italic') {
        formatting.isItalic = true;
      }
    }
    
  } catch (error) {
    console.error('Error extracting paragraph formatting:', error);
  }
  
  return formatting;
}

/**
 * Extract images from a paragraph
 * 
 * @param {Element} paragraph - Paragraph XML element
 * @param {number} paragraphIndex - Index of the paragraph (optional)
 * @returns {Array} - Array of image information
 */
function extractImagesFromParagraph(paragraph, paragraphIndex = null) {
  const images = [];
  
  try {
    // Look for drawing elements (newer format)
    const drawingNodes = selectNodes(".//w:drawing", paragraph);
    drawingNodes.forEach(drawing => {
      const imageInfo = extractImageInfo(drawing, 'drawing');
      if (imageInfo) {
        imageInfo.paragraphIndex = paragraphIndex;
        images.push(imageInfo);
      }
    });
    
    // Look for picture elements (older format)
    const pictNodes = selectNodes(".//w:pict", paragraph);
    pictNodes.forEach(pict => {
      const imageInfo = extractImageInfo(pict, 'pict');
      if (imageInfo) {
        imageInfo.paragraphIndex = paragraphIndex;
        images.push(imageInfo);
      }
    });
    
  } catch (error) {
    console.error('Error extracting images from paragraph:', error);
  }
  
  return images;
}

/**
 * Extract information about an image element
 * 
 * @param {Element} imageElement - Image XML element
 * @param {string} type - Type of image element ('drawing' or 'pict')
 * @returns {Object|null} - Image information or null
 */
function extractImageInfo(imageElement, type) {
  try {
    
    const imageInfo = {
      type: type,
      altText: '',
      title: '',
      width: null,
      height: null,
      relationshipId: null,
      positioning: {
        horizontal: null,
        vertical: null,
        xOffset: null,
        yOffset: null,
        relativeFrom: null,
        wrapType: null,
        distT: null,
        distB: null,
        distL: null,
        distR: null
      }
    };
    
    if (type === 'drawing') {
      // Extract from drawing element
      const docPr = selectSingleNode(".//wp:docPr", imageElement);
      if (docPr) {
        imageInfo.altText = docPr.getAttribute('descr') || '';
        imageInfo.title = docPr.getAttribute('title') || '';
      }
      
      const extent = selectSingleNode(".//wp:extent", imageElement);
      if (extent) {
        imageInfo.width = extent.getAttribute('cx');
        imageInfo.height = extent.getAttribute('cy');
      }
      
      const blip = selectSingleNode(".//a:blip", imageElement);
      if (blip) {
        imageInfo.relationshipId = blip.getAttribute('r:embed');
      }
      
      // Extract positioning information
      const anchor = selectSingleNode(".//wp:anchor", imageElement);
      if (anchor) {
        // Floating/anchored positioning
        const positionH = selectSingleNode(".//wp:positionH", anchor);
        if (positionH) {
          imageInfo.positioning.horizontal = positionH.getAttribute('relativeFrom');
          const posOffset = selectSingleNode(".//wp:posOffset", positionH);
          if (posOffset) {
            imageInfo.positioning.xOffset = posOffset.textContent;
          }
          const align = selectSingleNode(".//wp:align", positionH);
          if (align) {
            imageInfo.positioning.horizontal = align.textContent;
          }
        }
        
        const positionV = selectSingleNode(".//wp:positionV", anchor);
        if (positionV) {
          imageInfo.positioning.vertical = positionV.getAttribute('relativeFrom');
          const posOffset = selectSingleNode(".//wp:posOffset", positionV);
          if (posOffset) {
            imageInfo.positioning.yOffset = posOffset.textContent;
          }
          const align = selectSingleNode(".//wp:align", positionV);
          if (align) {
            imageInfo.positioning.vertical = align.textContent;
          }
        }
        
        const wrapType = selectSingleNode(".//wp:wrapNone | .//wp:wrapSquare | .//wp:wrapTight | .//wp:wrapThrough | .//wp:wrapTopAndBottom", anchor);
        if (wrapType) {
          imageInfo.positioning.wrapType = wrapType.tagName.replace('wp:', '');
        }
      }
      
      const inline = selectSingleNode(".//wp:inline", imageElement);
      if (inline) {
        // Inline positioning
        imageInfo.positioning.wrapType = 'inline';
        
        // Extract distance attributes for inline images
        imageInfo.positioning.distT = inline.getAttribute('distT');
        imageInfo.positioning.distB = inline.getAttribute('distB');
        imageInfo.positioning.distL = inline.getAttribute('distL');
        imageInfo.positioning.distR = inline.getAttribute('distR');
      }
      
      // Extract transform information (applies to both inline and anchored)
      const xfrm = selectSingleNode(".//a:xfrm", imageElement);
      if (xfrm) {
        const off = selectSingleNode(".//a:off", xfrm);
        if (off) {
          imageInfo.positioning.xOffset = off.getAttribute('x');
          imageInfo.positioning.yOffset = off.getAttribute('y');
        }
      }
      
    } else if (type === 'pict') {
      // Extract from pict element
      const shape = selectSingleNode(".//v:shape", imageElement);
      if (shape) {
        imageInfo.altText = shape.getAttribute('alt') || '';
        imageInfo.title = shape.getAttribute('title') || '';
      }
      
      const imagedata = selectSingleNode(".//v:imagedata", imageElement);
      if (imagedata) {
        imageInfo.relationshipId = imagedata.getAttribute('r:id');
      }
    }
    
    return imageInfo;
    
  } catch (error) {
    console.error('Error extracting image info:', error);
    return null;
  }
}

/**
 * Process header content for HTML output
 * Converts header paragraphs to HTML with preserved formatting
 * 
 * @param {Object} headerInfo - Header information from extraction
 * @param {Document} document - HTML document
 * @param {Object} styleInfo - Style information for CSS generation
 * @param {Object} zip - JSZip instance for image extraction (optional)
 * @param {string} outputDir - Output directory for images (optional)
 * @returns {Promise<Array>} - Array of HTML elements representing the header
 */
async function processHeaderForHtml(headerInfo, document, styleInfo, zip = null, outputDir = null) {
  const headerElements = [];
  
  if (!headerInfo.hasHeaderContent) {
    return headerElements;
  }
  
  try {
    // Process header images if we have access to the zip file
    let headerImageMap = {};
    if (zip && outputDir && headerInfo.headerImages.length > 0) {
      try {
        headerImageMap = await processHeaderImages(headerInfo.headerImages, zip, outputDir);
      } catch (error) {
        console.error('Error processing header images:', error);
      }
    }
    
    // Sort header paragraphs to put image paragraphs first
    const sortedParagraphs = [...headerInfo.headerParagraphs].sort((a, b) => {
      // Images first, then text paragraphs
      if (a.hasImages && !b.hasImages) return -1;
      if (!a.hasImages && b.hasImages) return 1;
      return 0; // Maintain original order within each group
    });
    
    sortedParagraphs.forEach((headerParagraph, index) => {
      const headerElement = document.createElement('div');
      headerElement.className = 'docx-header-paragraph';
      
      // Add style class if available
      if (headerParagraph.styleId) {
        const safeClassName = headerParagraph.styleId.toLowerCase().replace(/[^a-z0-9-_]/g, '-');
        headerElement.classList.add(`docx-p-${safeClassName}`);
      }
      
      // Add content
      if (headerParagraph.content) {
        const textSpan = document.createElement('span');
        textSpan.className = 'docx-header-text';
        textSpan.textContent = headerParagraph.content;
        headerElement.appendChild(textSpan);
      }
      
      // Add images if this paragraph has them
      if (headerParagraph.hasImages) {
        // Find the original index of this paragraph in the unsorted array
        const originalIndex = headerInfo.headerParagraphs.findIndex(p => p === headerParagraph);
        const paragraphImages = headerInfo.headerImages.filter(img => img.paragraphIndex === originalIndex);
        paragraphImages.forEach(imageInfo => {
          if (headerImageMap[imageInfo.relationshipId]) {
            
            // Create image container to handle paragraph-level alignment
            const imgContainer = document.createElement('div');
            imgContainer.className = 'docx-image-container';
            
            const imgElement = document.createElement('img');
            imgElement.src = headerImageMap[imageInfo.relationshipId].src;
            imgElement.alt = imageInfo.altText || 'Header image';
            imgElement.className = 'docx-header-image';
            
            // Apply paragraph-level alignment to container
            if (imageInfo.paragraphFormatting) {
              if (imageInfo.paragraphFormatting.isCentered) {
                imgContainer.style.textAlign = 'center';
              } else if (imageInfo.paragraphFormatting.isRightAligned) {
                imgContainer.style.textAlign = 'right';
              } else if (imageInfo.paragraphFormatting.isLeftAligned) {
                imgContainer.style.textAlign = 'left';
              }
}
            
            // Apply positioning styles
            if (imageInfo.positioning) {
              const pos = imageInfo.positioning;
              
              // Handle inline positioning with distances
              if (pos.wrapType === 'inline') {
                // Convert EMU to pixels for distances (1 EMU = 1/914400 inch, 1 inch = 96 pixels)
                if (pos.distT && parseInt(pos.distT) > 0) {
                  const topPx = Math.round(parseInt(pos.distT) / 914400 * 96);
                  imgElement.style.marginTop = `${topPx}px`;
                }
                if (pos.distB && parseInt(pos.distB) > 0) {
                  const bottomPx = Math.round(parseInt(pos.distB) / 914400 * 96);
                  imgElement.style.marginBottom = `${bottomPx}px`;
                }
                if (pos.distL && parseInt(pos.distL) > 0) {
                  const leftPx = Math.round(parseInt(pos.distL) / 914400 * 96);
                  imgElement.style.marginLeft = `${leftPx}px`;
                }
                if (pos.distR && parseInt(pos.distR) > 0) {
                  const rightPx = Math.round(parseInt(pos.distR) / 914400 * 96);
                  imgElement.style.marginRight = `${rightPx}px`;
                }
                
                // Apply transform offsets for inline images
                if (pos.xOffset && parseInt(pos.xOffset) !== 0) {
                  const xPx = Math.round(parseInt(pos.xOffset) / 914400 * 96);
                  imgElement.style.transform = `translateX(${xPx}px)`;
                }
                if (pos.yOffset && parseInt(pos.yOffset) !== 0) {
                  const yPx = Math.round(parseInt(pos.yOffset) / 914400 * 96);
                  const existingTransform = imgElement.style.transform || '';
                  imgElement.style.transform = existingTransform ? 
                    `${existingTransform} translateY(${yPx}px)` : `translateY(${yPx}px)`;
                }
              } else if (pos.wrapType && pos.wrapType !== 'inline') {
                // Handle floating/anchored positioning
                imgElement.style.position = 'absolute';
                
                // Convert EMU to pixels for offsets (1 EMU = 1/914400 inch, 1 inch = 96 pixels)
                if (pos.xOffset) {
                  const xPx = Math.round(parseInt(pos.xOffset) / 914400 * 96);
                  imgElement.style.left = `${xPx}px`;
                }
                if (pos.yOffset) {
                  const yPx = Math.round(parseInt(pos.yOffset) / 914400 * 96);
                  imgElement.style.top = `${yPx}px`;
                }
                
                // Handle alignment-based positioning
                if (pos.horizontal === 'center') {
                  imgElement.style.left = '50%';
                  imgElement.style.transform = 'translateX(-50%)';
                } else if (pos.horizontal === 'right') {
                  imgElement.style.right = '0';
                  imgElement.style.left = 'auto';
                }
                
                if (pos.vertical === 'center') {
                  imgElement.style.top = '50%';
                  imgElement.style.transform = imgElement.style.transform ? 
                    imgElement.style.transform + ' translateY(-50%)' : 'translateY(-50%)';
                } else if (pos.vertical === 'bottom') {
                  imgElement.style.bottom = '0';
                  imgElement.style.top = 'auto';
                }
              }
            }
            
            if (imageInfo.title) {
              imgElement.title = imageInfo.title;
            }
            
            // Add width and height if available
            if (imageInfo.width && imageInfo.height) {
              // Convert EMU (English Metric Units) to pixels (1 EMU = 1/914400 inch, 1 inch = 96 pixels)
              const widthPx = Math.round(parseInt(imageInfo.width) / 914400 * 96);
              const heightPx = Math.round(parseInt(imageInfo.height) / 914400 * 96);
              imgElement.style.width = `${widthPx}px`;
              imgElement.style.height = `${heightPx}px`;
            }
            // Add image to container and container to header
            imgContainer.appendChild(imgElement);
            
            headerElement.appendChild(imgContainer);
          }
        });
      }
      
      // Add inline styles for formatting preservation
      applyInlineFormatting(headerElement, headerParagraph.formatting);
      
      // Add data attributes for CSS targeting
      headerElement.setAttribute('data-header-paragraph', index.toString());
      if (headerParagraph.styleId) {
        headerElement.setAttribute('data-style-id', headerParagraph.styleId);
      }
      if (headerInfo.headerType) {
        headerElement.setAttribute('data-header-type', headerInfo.headerType);
      }
      
      headerElements.push(headerElement);
    });
    
  } catch (error) {
    console.error('Error processing header for HTML:', error);
  }
  
  return headerElements;
}

/**
 * Process header images and extract them to the output directory
 * 
 * @param {Array} headerImages - Array of header image information
 * @param {Object} zip - JSZip instance with DOCX content
 * @param {string} outputDir - Output directory for images
 * @returns {Promise<Object>} - Map of relationship IDs to image information
 */
async function processHeaderImages(headerImages, zip, outputDir) {
  const imageMap = {};
  const fs = require('fs');
  const path = require('path');
  
  try {
    console.log(`Processing ${headerImages.length} header images...`);
    console.log(`Output directory: ${outputDir}`);
    
    // Ensure images directory exists
    const imagesDir = path.join(outputDir, 'images');
    console.log(`Images directory path: ${imagesDir}`);
    if (!fs.existsSync(imagesDir)) {
      fs.mkdirSync(imagesDir, { recursive: true });
    }
    
    // Load header relationships to resolve image paths
    const headerRelsFiles = ['word/_rels/header1.xml.rels', 'word/_rels/header2.xml.rels', 'word/_rels/header3.xml.rels'];
    let headerRels = {};
    
    for (const relsFile of headerRelsFiles) {
      const relsFileObj = zip.file(relsFile);
      console.log(`Checking for relationship file: ${relsFile} - ${relsFileObj ? 'Found' : 'Not found'}`);
      if (relsFileObj) {
        try {
          // Use JSZip 3.0 async API
          const relsXml = await relsFileObj.async("string");
          console.log(`Successfully read ${relsFile}, content length: ${relsXml ? relsXml.length : 0}`);
          if (relsXml) {
            const relsDoc = new (require('xmldom').DOMParser)().parseFromString(relsXml);
            console.log(`Relationship XML content: ${relsXml.substring(0, 500)}...`);
            
            // Try different XPath queries to handle namespace issues
            let relationships = selectNodes("//r:Relationship", relsDoc);
            if (relationships.length === 0) {
              relationships = selectNodes("//Relationship", relsDoc);
            }
            if (relationships.length === 0) {
              // Try without namespace prefix
              const nodeList = relsDoc.getElementsByTagName('Relationship');
              relationships = Array.from(nodeList);
            }
            console.log(`Found ${relationships.length} relationships in ${relsFile}`);
            
            relationships.forEach(rel => {
              const id = rel.getAttribute('Id');
              const target = rel.getAttribute('Target');
              const type = rel.getAttribute('Type');
              console.log(`Relationship: ${id} -> ${target} (type: ${type})`);
              
              if (type && type.includes('image')) {
                headerRels[id] = target;
                console.log(`Found header image relationship: ${id} -> ${target}`);
              }
            });
          }
        } catch (relError) {
          console.warn(`Could not process ${relsFile}:`, relError.message);
        }
      }
    }
    
    // Process each header image
    console.log(`Processing ${headerImages.length} header images with relationships:`, Object.keys(headerRels));
    for (const imageInfo of headerImages) {
      console.log(`Processing image with relationshipId: ${imageInfo.relationshipId}, altText: ${imageInfo.altText}`);
      if (imageInfo.relationshipId && headerRels[imageInfo.relationshipId]) {
        const imagePath = headerRels[imageInfo.relationshipId];
        const fullImagePath = imagePath.startsWith('media/') ? `word/${imagePath}` : `word/media/${imagePath}`;
        console.log(`Looking for image file: ${fullImagePath}`);
        
        const imageFile = zip.file(fullImagePath);
        console.log(`Image file ${fullImagePath}: ${imageFile ? 'Found' : 'Not found'}`);
        if (imageFile) {
          try {
            // Use JSZip 3.0 async API
            const imageBuffer = await imageFile.async("nodebuffer");
            console.log(`Successfully read image buffer, size: ${imageBuffer.length} bytes`);
            const extension = path.extname(imagePath) || '.png';
            const sanitizedAltText = (imageInfo.altText || `header-image-${Date.now()}`)
              .replace(/[^\w-]/g, "_")
              .substring(0, 50);
            const filename = `${sanitizedAltText}${extension}`;
            const outputPath = path.join(imagesDir, filename);
            console.log(`Attempting to save image to: ${outputPath}`);
            console.log(`Images directory exists: ${fs.existsSync(imagesDir)}`);
            
            fs.writeFileSync(outputPath, imageBuffer);
            console.log(`Image file saved successfully: ${fs.existsSync(outputPath)}`);
            
            imageMap[imageInfo.relationshipId] = {
              src: `./images/${filename}`,
              alt: imageInfo.altText || 'Header image',
              title: imageInfo.title || '',
              width: imageInfo.width,
              height: imageInfo.height
            };
            
            console.log(`Extracted header image: ${filename}`);
          } catch (imgError) {
            console.warn(`Could not extract image ${fullImagePath}:`, imgError.message);
          }
        } else {
          console.warn(`Image file not found in zip: ${fullImagePath}`);
        }
      } else {
        console.warn(`No relationship found for image ${imageInfo.relationshipId} or missing relationshipId`);
      }
    }
    
  } catch (error) {
    console.error('Error processing header images:', error);
  }
  
  return imageMap;
}

/**
 * Apply inline formatting to a header element
 * 
 * @param {Element} element - HTML element
 * @param {Object} formatting - Formatting information
 */
function applyInlineFormatting(element, formatting) {
  try {
    if (formatting.fontSize) {
      element.style.fontSize = `${formatting.fontSize}pt`;
    }
    
    if (formatting.fontFamily) {
      element.style.fontFamily = formatting.fontFamily;
    }
    
    if (formatting.isBold) {
      element.style.fontWeight = 'bold';
    }
    
    if (formatting.isItalic) {
      element.style.fontStyle = 'italic';
    }
    
    if (formatting.alignment) {
      element.style.textAlign = formatting.alignment;
    }
    
    if (formatting.color && formatting.color !== 'auto') {
      element.style.color = `#${formatting.color}`;
    }
    
  } catch (error) {
    console.error('Error applying inline formatting:', error);
  }
}

module.exports = {
  extractDocumentHeader,
  processHeaderForHtml,
  processHeaderImages,
  extractHeaderFromXml,
  extractHeaderFromDocument,
  analyzeHeaderParagraph,
  extractParagraphFormatting,
  extractImagesFromParagraph,
  isLikelyHeaderContent
};
