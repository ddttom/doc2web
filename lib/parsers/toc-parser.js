// lib/parsers/toc-parser.js - TOC (Table of Contents) parsing functions
const { selectNodes, selectSingleNode } = require('../xml/xpath-utils');
const { convertTwipToPt } = require('../utils/unit-converter');
const { getLeaderChar } = require('../utils/common-utils');

/**
 * Parse run properties specifically for TOC entries
 * Extracts character-level formatting that affects TOC appearance
 * 
 * @param {Node} rPrNode - Running properties node
 * @returns {Object} - Running properties with italic, bold, font, etc.
 */
function parseRunPropertiesForTOC(rPrNode) {
  const props = {};
  
  try {
    // Font information
    const fontNode = selectSingleNode("w:rFonts", rPrNode);
    if (fontNode) {
      props.font = {
        ascii: fontNode.getAttribute('w:ascii'),
        hAnsi: fontNode.getAttribute('w:hAnsi')
      };
    }
    
    // Size (convert from half-points to points)
    const szNode = selectSingleNode("w:sz", rPrNode);
    if (szNode) {
      const sizeHalfPoints = parseInt(szNode.getAttribute('w:val'), 10) || 22;
      props.fontSize = (sizeHalfPoints / 2) + 'pt';
    }
    
    // Bold
    const bNode = selectSingleNode("w:b", rPrNode);
    props.bold = bNode !== null;
    
    // Italic - CRITICAL for TOC fidelity
    const iNode = selectSingleNode("w:i", rPrNode);
    props.italic = iNode !== null;
    
    // Color
    const colorNode = selectSingleNode("w:color", rPrNode);
    if (colorNode) {
      const colorVal = colorNode.getAttribute('w:val');
      if (colorVal && colorVal !== 'auto') {
        props.color = '#' + colorVal;
      }
    }
    
    // Underline
    const uNode = selectSingleNode("w:u", rPrNode);
    if (uNode) {
      props.underline = uNode.getAttribute('w:val') || 'single';
    }
    
  } catch (error) {
    console.error('Error parsing TOC run properties:', error);
  }
  
  return props;
}

/**
 * Enhanced TOC parsing to better capture leader lines and formatting</search>
</search_and_replace>
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
      character: null, // Remove leader character for web-based TOC
      spacesBetween: 0, // Remove spacing for web-based TOC
      position: 0, // Remove position for web-based TOC
    },
  };

  try {
    // Check for TOC field
    const tocFieldNodes = selectNodes(
      "//w:fldChar[@w:fldCharType='begin']/following-sibling::w:instrText[contains(., 'TOC')]",
      documentDoc
    );
    tocStyles.hasTableOfContents = tocFieldNodes.length > 0;

    // Improved: Detect paragraph styles that might be part of TOC even if not explicitly marked
    const paragraphNodes = selectNodes("//w:p", documentDoc);
    let inTocSection = false;
    let potentialTocEntries = [];

    // First pass: identify the TOC section
    for (let i = 0; i < paragraphNodes.length; i++) {
      const p = paragraphNodes[i];
      const pStyle = selectSingleNode("w:pPr/w:pStyle", p);
      const styleId = pStyle ? pStyle.getAttribute("w:val") : "";

      // Check if this is a TOC heading or start of TOC
      if (
        styleId &&
        (styleId.toLowerCase().includes("toc") ||
          styleId.toLowerCase() === "tableofcontents" ||
          styleId.toLowerCase() === "contentsheading")
      ) {
        inTocSection = true;
        continue;
      }

      if (inTocSection) {
        // Look for tab stops with right alignment and leaders - key indicator of TOC entries
        const tabsNode = selectSingleNode("w:pPr/w:tabs", p);
        if (tabsNode) {
          const rightTabWithLeader = selectSingleNode(
            "w:tab[@w:val='right' and @w:leader]",
            tabsNode
          );
          if (rightTabWithLeader) {
            // Found a paragraph with right-aligned tab with leader - likely a TOC entry
            // Note: We skip leader and position extraction to avoid page number generation
            // const leader = rightTabWithLeader.getAttribute("w:leader");
            // const pos = rightTabWithLeader.getAttribute("w:pos");

            // Skip leader and position processing for web-based TOC without page numbers
            // if (leader) {
            //   tocStyles.leaderStyle.character = getLeaderChar(leader);
            // }

            // if (pos) {
            //   tocStyles.leaderStyle.position = convertTwipToPt(pos);
            // }

            // Get indentation info to help identify TOC level
            const indNode = selectSingleNode("w:pPr/w:ind", p);
            const leftIndent = indNode
              ? parseInt(indNode.getAttribute("w:left") || "0", 10)
              : 0;
            const hangingIndent = indNode
              ? parseInt(indNode.getAttribute("w:hanging") || "0", 10)
              : 0;
            const firstLineIndent = indNode
              ? parseInt(indNode.getAttribute("w:firstLine") || "0", 10)
              : 0;

            // NEW: Extract run properties for character formatting
            const runNodes = selectNodes("w:r", p);
            const runProperties = [];
            runNodes.forEach(runNode => {
              const rPrNode = selectSingleNode("w:rPr", runNode);
              if (rPrNode) {
                const runProps = parseRunPropertiesForTOC(rPrNode);
                const textNode = selectSingleNode("w:t", runNode);
                if (textNode) {
                  runProperties.push({
                    text: textNode.textContent,
                    formatting: runProps
                  });
                }
              }
            });

            // Store enhanced style info for this TOC entry
            potentialTocEntries.push({
              styleId: styleId,
              leftIndent: leftIndent,
              hangingIndent: hangingIndent,
              firstLineIndent: firstLineIndent,
              tabPos: 0, // Remove tab position to avoid page number generation
              leader: null, // Remove leader to avoid page number generation
              paragraphIndex: i,
              runProperties: runProperties, // NEW: Character-level formatting
              hasItalic: runProperties.some(rp => rp.formatting.italic), // NEW: Italic detection
              hasBold: runProperties.some(rp => rp.formatting.bold), // NEW: Bold detection
            });
          }
        }

        // Check for non-TOC styles that might indicate end of TOC section
        if (
          styleId &&
          !styleId.toLowerCase().includes("toc") &&
          styleId.toLowerCase().includes("heading") &&
          potentialTocEntries.length > 0
        ) {
          // Found a heading after TOC entries - likely end of TOC
          inTocSection = false;
          break;
        }
      }
    }

    // Process potential TOC entries to identify TOC levels
    if (potentialTocEntries.length > 0) {
      // Get unique indentation values, sorted
      const indentLevels = [
        ...new Set(potentialTocEntries.map((e) => e.leftIndent)),
      ].sort((a, b) => a - b);

      // Create TOC entry styles based on indentation levels
      indentLevels.forEach((indent, index) => {
        const entriesWithThisIndent = potentialTocEntries.filter(
          (e) => e.leftIndent === indent
        );
        if (entriesWithThisIndent.length > 0) {
          // Use style info from the first entry at this indentation level
          const entry = entriesWithThisIndent[0];

          // Get style details if available
          let styleDetails = {};
          if (entry.styleId) {
            const styleNode = selectSingleNode(
              `//w:style[@w:styleId='${entry.styleId}']`,
              styleDoc
            );
            if (styleNode) {
              const nameNode = selectSingleNode("w:name", styleNode);
              styleDetails.name = nameNode
                ? nameNode.getAttribute("w:val")
                : entry.styleId;

              // Extract font and other properties
              const rPrNode = selectSingleNode("w:rPr", styleNode);
              if (rPrNode) {
                const fontNode = selectSingleNode("w:rFonts", rPrNode);
                if (fontNode) {
                  styleDetails.fontFamily =
                    fontNode.getAttribute("w:ascii") ||
                    fontNode.getAttribute("w:hAnsi");
                }

                const szNode = selectSingleNode("w:sz", rPrNode);
                if (szNode) {
                  const size = parseInt(szNode.getAttribute("w:val"), 10) || 22; // Default 11pt
                  styleDetails.fontSize = size / 2 + "pt";
                }

                const colorNode = selectSingleNode("w:color", rPrNode);
                if (colorNode) {
                  styleDetails.color = "#" + colorNode.getAttribute("w:val");
                }
              }
            }
          }

          tocStyles.tocEntryStyles.push({
            level: index + 1,
            indentation: { left: convertTwipToPt(indent) },
            tabPosition: 0, // Remove tab position for web-based TOC
            leader: null, // Remove leader for web-based TOC
            leaderChar: null, // Remove leader character for web-based TOC
            styleId: entry.styleId,
            fontFamily: styleDetails.fontFamily,
            fontSize: styleDetails.fontSize,
            color: styleDetails.color,
          });
        }
      });
    }

    // If no TOC entry styles were found but we have TOC field, create default styles
    if (tocStyles.tocEntryStyles.length === 0 && tocStyles.hasTableOfContents) {
      for (let i = 1; i <= 3; i++) {
        // Default to 3 TOC levels
        tocStyles.tocEntryStyles.push({
          level: i,
          indentation: { left: (i - 1) * 20 }, // Default 20pt indentation per level
          tabPosition: 0, // Remove tab position for web-based TOC
          leader: null, // Remove leader for web-based TOC
          leaderChar: null, // Remove leader character for web-based TOC
        });
      }
    }

    // Look for TOC heading style
    const tocHeadingNodes = selectNodes(
      "//w:style[w:name/@w:val='TOC Heading' or contains(@w:styleId, 'TOCHeading')]",
      styleDoc
    );
    if (tocHeadingNodes.length > 0) {
      const headingNode = tocHeadingNodes[0];
      const nameNode = selectSingleNode("w:name", headingNode);
      tocStyles.tocHeadingStyle.name = nameNode
        ? nameNode.getAttribute("w:val")
        : "TOC Heading";

      // Extract font and other properties for TOC heading
      const rPrNode = selectSingleNode("w:rPr", headingNode);
      if (rPrNode) {
        const fontNode = selectSingleNode("w:rFonts", rPrNode);
        if (fontNode) {
          tocStyles.tocHeadingStyle.fontFamily =
            fontNode.getAttribute("w:ascii") ||
            fontNode.getAttribute("w:hAnsi");
        }

        const szNode = selectSingleNode("w:sz", rPrNode);
        if (szNode) {
          const size = parseInt(szNode.getAttribute("w:val"), 10) || 28; // Default 14pt
          tocStyles.tocHeadingStyle.fontSize = size / 2 + "pt";
        }

        const bNode = selectSingleNode("w:b", rPrNode);
        tocStyles.tocHeadingStyle.bold = bNode !== null;

        const colorNode = selectSingleNode("w:color", rPrNode);
        if (colorNode) {
          tocStyles.tocHeadingStyle.color =
            "#" + colorNode.getAttribute("w:val");
        }
      }
    }
  } catch (error) {
    console.error("Error parsing TOC styles:", error);
  }

// Initialize potentialTocEntries if not defined (fallback for error cases)
  if (typeof potentialTocEntries === 'undefined') {
    potentialTocEntries = [];
  }
  // Add enhanced TOC entries data for improved style fidelity
  tocStyles.enhancedTOCEntries = potentialTocEntries;
  
  return tocStyles;
}
module.exports = {
  parseTocStyles
};
