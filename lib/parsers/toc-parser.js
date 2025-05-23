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
      character: ".",
      spacesBetween: 3,
      position: 0,
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
            const leader = rightTabWithLeader.getAttribute("w:leader");
            const pos = rightTabWithLeader.getAttribute("w:pos");

            if (leader) {
              tocStyles.leaderStyle.character = getLeaderChar(leader);
            }

            if (pos) {
              tocStyles.leaderStyle.position = convertTwipToPt(pos);
            }

            // Get indentation info to help identify TOC level
            const indNode = selectSingleNode("w:pPr/w:ind", p);
            const leftIndent = indNode
              ? parseInt(indNode.getAttribute("w:left") || "0", 10)
              : 0;

            // Store style info for this potential TOC entry
            potentialTocEntries.push({
              styleId: styleId,
              leftIndent: leftIndent,
              tabPos: pos ? parseInt(pos, 10) : 0,
              leader: leader,
              paragraphIndex: i,
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
            tabPosition: convertTwipToPt(entry.tabPos),
            leader: entry.leader,
            leaderChar: getLeaderChar(entry.leader),
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
          tabPosition: 468, // Default 6.5 inches
          leader: "1", // Default dot leader
          leaderChar: ".",
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

  return tocStyles;
}
module.exports = {
  parseTocStyles
};