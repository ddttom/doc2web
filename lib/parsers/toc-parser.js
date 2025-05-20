// lib/parsers/toc-parser.js - TOC (Table of Contents) parsing functions
const { selectNodes, selectSingleNode } = require('../xml/xpath-utils');
const { convertTwipToPt } = require('../utils/unit-converter');
const { getLeaderChar } = require('../utils/common-utils');

/**
 * Enhanced TOC parsing to better capture leader lines and formatting
 * Extracts TOC styles, heading styles, and leader line information
 * 
 * @param {Document} documentDoc - Document XML
 * @param {Document} styleDoc - Style XML
 * @returns {Object} - TOC specific style information
 */
function parseTocStyles(documentDoc, styleDoc) {
  const tocStyles = {
    hasTableOfContents: false,
    tocHeadingStyle: {},
    tocEntryStyles: [],
    leaderStyle: {
      character: '.',
      spacesBetween: 3,
      position: 0
    }
  };
  
  try {
    // Check for TOC field
    const tocFieldNodes = selectNodes("//w:fldChar[@w:fldCharType='begin']/following-sibling::w:instrText[contains(., 'TOC')]", documentDoc);
    tocStyles.hasTableOfContents = tocFieldNodes.length > 0;
    
    if (tocStyles.hasTableOfContents) {
      // Extract TOC field properties
      for (const tocNode of tocFieldNodes) {
        const instrText = tocNode.textContent || '';
        
        // Extract leader character if specified
        const leaderMatch = instrText.match(/\\[a-z]\s+([\.\_\-])/);
        if (leaderMatch) {
          tocStyles.leaderStyle.character = leaderMatch[1];
        }
        
        // Extract heading levels if specified
        const levelMatch = instrText.match(/\\o\s+"([1-9])"-"([1-9])"/);
        if (levelMatch) {
          tocStyles.startLevel = parseInt(levelMatch[1], 10);
          tocStyles.endLevel = parseInt(levelMatch[2], 10);
        }
      }
    }
    
    // Look for TOC styles in the style definitions
    const tocStyleNodes = selectNodes("//w:style[contains(@w:styleId, 'TOC') or contains(w:name/@w:val, 'TOC') or contains(w:name/@w:val, 'Contents') or contains(@w:styleId, 'toc')]", styleDoc);
    
    tocStyleNodes.forEach((node, index) => {
      const styleId = node.getAttribute('w:styleId');
      const nameNode = selectSingleNode("w:name", node);
      const name = nameNode ? nameNode.getAttribute('w:val') : styleId;
      
      // Parse specific style properties
      const pPrNode = selectSingleNode("w:pPr", node);
      const rPrNode = selectSingleNode("w:rPr", node);
      
      const style = {
        id: styleId,
        name,
        indentation: {},
        fontSize: null,
        fontFamily: null,
        tabs: []
      };
      
      // Extract paragraph properties
      if (pPrNode) {
        // Get indentation
        const indNode = selectSingleNode("w:ind", pPrNode);
        if (indNode) {
          const left = indNode.getAttribute('w:left');
          const hanging = indNode.getAttribute('w:hanging');
          const firstLine = indNode.getAttribute('w:firstLine');
          
          if (left) style.indentation.left = convertTwipToPt(left);
          if (hanging) style.indentation.hanging = convertTwipToPt(hanging);
          if (firstLine) style.indentation.firstLine = convertTwipToPt(firstLine);
        }
        
        // Get tab stops
        const tabsNode = selectSingleNode("w:tabs", pPrNode);
        if (tabsNode) {
          const tabNodes = selectNodes("w:tab", tabsNode);
          
          tabNodes.forEach(tabNode => {
            const pos = tabNode.getAttribute('w:pos');
            const val = tabNode.getAttribute('w:val');
            const leader = tabNode.getAttribute('w:leader');
            
            if (pos && val) {
              const tabPosition = convertTwipToPt(pos);
              style.tabs.push({
                position: tabPosition,
                type: val,
                leader: leader || 'none',
                leaderChar: getLeaderChar(leader)
              });
              
              // If we find a right-aligned tab with leader, use it for TOC dots
              if (val === 'right' && leader) {
                tocStyles.leaderStyle.character = getLeaderChar(leader);
                tocStyles.leaderStyle.position = tabPosition;
              }
            }
          });
        }
      }
      
      // Extract text properties
      if (rPrNode) {
        // Font size
        const szNode = selectSingleNode("w:sz", rPrNode);
        if (szNode) {
          const size = parseInt(szNode.getAttribute('w:val'), 10) || 22; // Default 11pt
          style.fontSize = (size / 2) + 'pt';
        }
        
        // Font family
        const fontNode = selectSingleNode("w:rFonts", rPrNode);
        if (fontNode) {
          style.fontFamily = fontNode.getAttribute('w:ascii') || 
                           fontNode.getAttribute('w:hAnsi') || 
                           'Calibri';
        }
      }
      
      // Determine if this is a TOC heading or entry style
      if (name.includes('Heading') || styleId === 'TOCHeading') {
        tocStyles.tocHeadingStyle = style;
      } else {
        // Extract level info from the style ID if possible
        const levelMatch = styleId.match(/TOC(\d+)/i);
        if (levelMatch) {
          style.level = parseInt(levelMatch[1], 10);
        } else {
          // Default to sequential level based on order
          style.level = index + 1;
        }
        
        tocStyles.tocEntryStyles.push(style);
      }
    });
    
    // Sort TOC entry styles by level
    tocStyles.tocEntryStyles.sort((a, b) => a.level - b.level);
    
    // Scan document for actual TOC entries to get better tab/dot information
    const paragraphNodes = selectNodes("//w:p", documentDoc);
    let inTocSection = false;
    
    for (let i = 0; i < paragraphNodes.length; i++) {
      const p = paragraphNodes[i];
      
      // Check if this is a TOC heading paragraph
      const pStyle = selectSingleNode("w:pPr/w:pStyle", p);
      const styleId = pStyle ? pStyle.getAttribute('w:val') : '';
      
      // Start of TOC detection - look for TOC heading styles or content that indicates TOC
      if (styleId?.toLowerCase().includes('toc') || 
          (selectSingleNode(".//w:t[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'table of contents') or contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'contents')]", p) !== null)) {
        inTocSection = true;
        continue;
      }
      
      // If we're in TOC section, analyze the paragraph for tab stops and dots
      if (inTocSection) {
        // Get text content to check if it looks like TOC entry
        const textNodes = selectNodes(".//w:t", p);
        let text = '';
        textNodes.forEach(t => text += t.textContent || '');
        
        // Check if this paragraph has page number pattern (text....number)
        // Look for text followed by whitespace and/or dots, then numbers at the end
        const pageNumMatch = text.match(/^(.*?)[\.\s]*(\d+)$/);
        
        if (pageNumMatch) {
          // This looks like a TOC entry with page number
          const tabsNode = selectSingleNode("w:pPr/w:tabs", p);
          
          if (tabsNode) {
            const rightTab = selectSingleNode("w:tab[@w:val='right']", tabsNode);
            if (rightTab) {
              const leader = rightTab.getAttribute('w:leader');
              const pos = rightTab.getAttribute('w:pos');
              
              if (leader) {
                tocStyles.leaderStyle.character = getLeaderChar(leader);
              }
              
              if (pos) {
                tocStyles.leaderStyle.position = convertTwipToPt(pos);
              }
            }
          }
        } else if (text.trim() === '' || styleId === 'Normal') {
          // Empty line or normal paragraph might indicate end of TOC
          // But only if we're past at least a few TOC entries
          if (i > 5) {
            inTocSection = false;
          }
        }
      }
    }
    
    // If leader style wasn't found but we have TOC entries, set a default
    if (tocStyles.leaderStyle.position === 0 && tocStyles.tocEntryStyles.length > 0) {
      tocStyles.leaderStyle.position = 6 * 72; // Default 6 inches
      tocStyles.leaderStyle.character = '.';
    }
    
  } catch (error) {
    console.error('Error parsing TOC styles:', error);
  }
  
  return tocStyles;
}

module.exports = {
  parseTocStyles
};