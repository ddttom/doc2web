// File: ddttom/doc2web/doc2web-804421e293a28695ca8c4527a7c79f342ff0c562/lib/parsers/numbering-parser.js

const { selectNodes, selectSingleNode } = require("../xml/xpath-utils");
const { convertTwipToPt } = require("../utils/unit-converter");

/**
 * Enhanced numbering definition parsing to capture complete numbering context
 * Extracts numbering definitions, levels, and formatting information with full context
 * * @param {Document} numberingDoc - Numbering XML document
 * @returns {Object} - Complete numbering definitions with resolution context
 */
function parseNumberingDefinitions(numberingDoc) {
  const numberingDefs = {
    abstractNums: {},
    nums: {},
    numIdMap: {},
    sequenceTrackers: {},
  };

  if (!numberingDoc) {
    return numberingDefs;
  }

  try {
    const abstractNumNodes = selectNodes("//w:abstractNum", numberingDoc);

    abstractNumNodes.forEach((node) => {
      const id = node.getAttribute("w:abstractNumId");
      if (!id) return;

      const abstractNum = {
        id,
        levels: {},
        multiLevelType: null,
        restartNumberingAfterBreak: false,
      };

      const multiLevelTypeNode = selectSingleNode("w:multiLevelType", node);
      if (multiLevelTypeNode) {
        abstractNum.multiLevelType = multiLevelTypeNode.getAttribute("w:val");
      }

      const levelNodes = selectNodes("w:lvl", node);
      levelNodes.forEach((lvlNode) => {
        const ilvl = lvlNode.getAttribute("w:ilvl");
        if (ilvl === null) return;

        const level = {
          level: parseInt(ilvl, 10), // 0-indexed
          format: "decimal",
          text: `%${parseInt(ilvl, 10) + 1}.`, // Default text like %1. %2.
          alignment: "left",
          indentation: {},
          isLegal: false,
          textFormat: "", // Will be w:lvlText@w:val
          restart: null,
          start: 1,
          suffix: "tab",
          tentative: false,
          runProps: {}, // For font/color of the number itself
          paragraphProps: {}, // For paragraph properties tied to this list level
        };

        const numFmtNode = selectSingleNode("w:numFmt", lvlNode);
        if (numFmtNode) {
          level.format = numFmtNode.getAttribute("w:val") || "decimal";
        }

        const lvlTextNode = selectSingleNode("w:lvlText", lvlNode);
        if (lvlTextNode) {
          level.text =
            lvlTextNode.getAttribute("w:val") || `%${level.level + 1}.`;
          level.textFormat = level.text; // Store the original w:lvlText value
          level.parsedFormat = parseNumberingTextFormat(level.text);
        } else {
          // If no w:lvlText, generate a basic one based on format for current level
          level.textFormat = `%${level.level + 1}${
            level.format === "bullet" ? "" : "."
          }`;
          level.parsedFormat = parseNumberingTextFormat(level.textFormat);
        }

        const suffixNode = selectSingleNode("w:suff", lvlNode);
        if (suffixNode) {
          level.suffix = suffixNode.getAttribute("w:val") || "tab";
        }

        const isLegalNode = selectSingleNode("w:isLgl", lvlNode);
        if (isLegalNode) {
          level.isLegal =
            isLegalNode.getAttribute("w:val") === "1" ||
            isLegalNode.getAttribute("w:val") === "true";
        }

        const tentativeAttr = lvlNode.getAttribute("w:tentative");
        if (tentativeAttr) {
          level.tentative = tentativeAttr === "1" || tentativeAttr === "true";
        }

        const alignmentNode = selectSingleNode("w:lvlJc", lvlNode);
        if (alignmentNode) {
          level.alignment = alignmentNode.getAttribute("w:val") || "left";
        }

        const indentNode = selectSingleNode("w:pPr/w:ind", lvlNode);
        if (indentNode) {
          const left = indentNode.getAttribute("w:left");
          const hanging = indentNode.getAttribute("w:hanging");
          const firstLine = indentNode.getAttribute("w:firstLine");
          const startIndent = indentNode.getAttribute("w:start"); // w:start for overall indent
          const endIndent = indentNode.getAttribute("w:end");

          if (left) level.indentation.left = convertTwipToPt(left);
          if (hanging) level.indentation.hanging = convertTwipToPt(hanging);
          if (firstLine)
            level.indentation.firstLine = convertTwipToPt(firstLine);
          if (startIndent)
            level.indentation.start = convertTwipToPt(startIndent);
          if (endIndent) level.indentation.end = convertTwipToPt(endIndent);
        }

        const lvlRestartNode = selectSingleNode("w:lvlRestart", lvlNode);
        if (lvlRestartNode) {
          level.restart = parseInt(lvlRestartNode.getAttribute("w:val"), 10);
        }

        const startNode = selectSingleNode("w:start", lvlNode);
        if (startNode) {
          level.start = parseInt(startNode.getAttribute("w:val"), 10) || 1;
        }

        const pPrNode = selectSingleNode("w:pPr", lvlNode);
        if (pPrNode) {
          level.paragraphProps = parseParagraphPropertiesForNumbering(pPrNode);
        }

        const rPrNode = selectSingleNode("w:rPr", lvlNode);
        if (rPrNode) {
          level.runProps = parseRunPropertiesForNumbering(rPrNode);
        }

        abstractNum.levels[ilvl] = level;
      });

      numberingDefs.abstractNums[id] = abstractNum;
    });

    const numNodes = selectNodes("//w:num", numberingDoc);

    numNodes.forEach((node) => {
      const id = node.getAttribute("w:numId");
      if (!id) return;

      const abstractNumIdNode = selectSingleNode("w:abstractNumId", node);
      if (!abstractNumIdNode) return;

      const abstractNumId = abstractNumIdNode.getAttribute("w:val");
      if (!abstractNumId) return;

      const levelOverrideNodes = selectNodes("w:lvlOverride", node);
      const overrides = {};

      levelOverrideNodes.forEach((overrideNode) => {
        const ilvl = overrideNode.getAttribute("w:ilvl");
        if (ilvl === null) return;

        const override = {};

        const startOverrideNode = selectSingleNode(
          "w:startOverride",
          overrideNode
        );
        if (startOverrideNode) {
          override.startOverride = parseInt(
            startOverrideNode.getAttribute("w:val"),
            10
          );
        }

        const lvlNode = selectSingleNode("w:lvl", overrideNode);
        if (lvlNode) {
          override.completeLevel = parseLevelDefinition(lvlNode); // This parses a w:lvl element
        }

        overrides[ilvl] = override;
      });

      numberingDefs.nums[id] = {
        id,
        abstractNumId,
        overrides: Object.keys(overrides).length > 0 ? overrides : null,
      };

      numberingDefs.numIdMap[id] = abstractNumId;
      numberingDefs.sequenceTrackers[id] = {
        levelCounters: {},
        lastSequence: {},
        restartPoints: {},
      };
    });
  } catch (error) {
    console.error("Error parsing numbering definitions:", error);
  }

  return numberingDefs;
}

function parseNumberingTextFormat(textFormat) {
  const parsed = {
    pattern: textFormat,
    segments: [], // Will store literal strings and level placeholders
  };
  if (!textFormat) return parsed;

  let currentSegment = "";
  for (let i = 0; i < textFormat.length; i++) {
    if (
      textFormat[i] === "%" &&
      i + 1 < textFormat.length &&
      /\d/.test(textFormat[i + 1])
    ) {
      if (currentSegment) {
        parsed.segments.push({ type: "literal", value: currentSegment });
        currentSegment = "";
      }
      const levelNum = parseInt(textFormat[i + 1], 10);
      parsed.segments.push({ type: "level", value: levelNum }); // levelNum is 1-indexed here
      i++; // Skip the digit
    } else {
      currentSegment += textFormat[i];
    }
  }
  if (currentSegment) {
    parsed.segments.push({ type: "literal", value: currentSegment });
  }
  return parsed;
}

function parseParagraphPropertiesForNumbering(pPrNode) {
  const props = {};
  try {
    const spacingNode = selectSingleNode("w:spacing", pPrNode);
    if (spacingNode) {
      props.spacing = {
        before: spacingNode.getAttribute("w:before"),
        after: spacingNode.getAttribute("w:after"),
        line: spacingNode.getAttribute("w:line"),
        lineRule: spacingNode.getAttribute("w:lineRule"),
      };
    }
    const tabsNode = selectSingleNode("w:tabs", pPrNode);
    if (tabsNode) {
      const tabNodes = selectNodes("w:tab", tabsNode);
      props.tabs = [];
      tabNodes.forEach((tabNode) => {
        const pos = tabNode.getAttribute("w:pos");
        const val = tabNode.getAttribute("w:val");
        const leader = tabNode.getAttribute("w:leader");
        if (pos && val) {
          props.tabs.push({
            position: pos,
            positionPt: convertTwipToPt(pos),
            type: val,
            leader: leader || "none",
          });
        }
      });
    }
  } catch (error) {
    console.error("Error parsing paragraph properties for numbering:", error);
  }
  return props;
}

function parseRunPropertiesForNumbering(rPrNode) {
  const props = {};
  try {
    const fontNode = selectSingleNode("w:rFonts", rPrNode);
    if (fontNode) {
      props.font = {
        ascii: fontNode.getAttribute("w:ascii"),
        hAnsi: fontNode.getAttribute("w:hAnsi"),
        eastAsia: fontNode.getAttribute("w:eastAsia"),
        cs: fontNode.getAttribute("w:cs"),
      };
    }
    const szNode = selectSingleNode("w:sz", rPrNode);
    if (szNode) {
      props.fontSize =
        (parseInt(szNode.getAttribute("w:val"), 10) || 22) / 2 + "pt";
    }
    props.bold = selectSingleNode("w:b", rPrNode) !== null;
    props.italic = selectSingleNode("w:i", rPrNode) !== null;
    const colorNode = selectSingleNode("w:color", rPrNode);
    if (colorNode) {
      const colorVal = colorNode.getAttribute("w:val");
      if (colorVal) {
        props.color = "#" + colorVal;
      }
    }
  } catch (error) {
    console.error("Error parsing run properties for numbering:", error);
  }
  return props;
}

function parseLevelDefinition(lvlNode) {
  // Used for lvlOverride/w:lvl
  const level = { indentation: {}, runProps: {}, paragraphProps: {} };
  const numFmtNode = selectSingleNode("w:numFmt", lvlNode);
  if (numFmtNode) {
    level.format = numFmtNode.getAttribute("w:val");
  }

  const lvlTextNode = selectSingleNode("w:lvlText", lvlNode);
  if (lvlTextNode) {
    level.text = lvlTextNode.getAttribute("w:val");
    level.textFormat = level.text;
    level.parsedFormat = parseNumberingTextFormat(level.text);
  }

  const startNode = selectSingleNode("w:start", lvlNode);
  if (startNode) {
    level.start = parseInt(startNode.getAttribute("w:val"), 10);
  }

  const pPrNode = selectSingleNode("w:pPr", lvlNode);
  if (pPrNode) {
    level.paragraphProps = parseParagraphPropertiesForNumbering(pPrNode);
    const indentNode = selectSingleNode("w:ind", pPrNode); // also parse indent for the override
    if (indentNode) {
      const left = indentNode.getAttribute("w:left");
      const hanging = indentNode.getAttribute("w:hanging");
      const firstLine = indentNode.getAttribute("w:firstLine");
      if (left) level.indentation.left = convertTwipToPt(left);
      if (hanging) level.indentation.hanging = convertTwipToPt(hanging);
      if (firstLine) level.indentation.firstLine = convertTwipToPt(firstLine);
    }
  }
  const rPrNode = selectSingleNode("w:rPr", lvlNode);
  if (rPrNode) {
    level.runProps = parseRunPropertiesForNumbering(rPrNode);
  }

  return level;
}

function getCSSCounterFormat(format) {
  const formatMap = {
    decimal: "decimal",
    lowerLetter: "lower-alpha",
    upperLetter: "upper-alpha",
    lowerRoman: "lower-roman",
    upperRoman: "upper-roman",
    bullet: "disc", // CSS 'disc' is common for bullet
    decimalZero: "decimal-leading-zero",
    // Add more mappings as needed based on DOCX common values and CSS equivalents
    none: "none", // For no numbering
  };
  return formatMap[format] || "decimal"; // Fallback to decimal
}

/**
 * Get CSS counter content for numbering based on parsed format.
 * This version constructs the CSS 'content' string by interpreting levelDef.parsedFormat.
 * @param {Object} levelDef - The definition for the current level.
 * @param {string} abstractNumId - The ID of the abstract numbering definition.
 * @param {Object} numberingDefs - The complete numbering definitions object.
 * @returns {string} - CSS content value, e.g., "'Chapter ' counter(c1-0, decimal) '. ' counter(c1-1, decimal)"
 */
function getCSSCounterContent(levelDef, abstractNumId, numberingDefs) {
  if (
    !levelDef ||
    !levelDef.parsedFormat ||
    !levelDef.parsedFormat.segments ||
    levelDef.parsedFormat.segments.length === 0
  ) {
    // Fallback for the current level if parsedFormat is not helpful
    const currentLevelIndex = levelDef.level; // Assuming levelDef.level is 0-indexed
    const counterName = `docx-counter-${abstractNumId}-${currentLevelIndex}`;
    const format = getCSSCounterFormat(levelDef.format);
    let suffix = "'.'"; // Default suffix for decimal numbers
    if (levelDef.format === "bullet") suffix = "''"; // No suffix for bullets
    else if (levelDef.suffix === "tab")
      suffix = "'\\00a0\\00a0'"; // Represent tab as spaces
    else if (levelDef.suffix === "space") suffix = "' '";
    else if (levelDef.suffix === "nothing") suffix = "''";

    if (format === "none") return "''"; // No content if format is 'none'
    return `counter(${counterName}, ${format}) ${suffix}`;
  }

  let contentValue = "";
  levelDef.parsedFormat.segments.forEach((segment, index) => {
    if (index > 0) {
      contentValue += " + "; // CSS content concatenation
    }
    if (segment.type === "literal") {
      contentValue += `"${segment.value.replace(/"/g, '\\"')}"`; // Quote literal parts
    } else if (segment.type === "level") {
      const targetLevelIndex = segment.value - 1; // Placeholder %N is 1-indexed, our levels are 0-indexed
      const targetLevelDef =
        numberingDefs.abstractNums[abstractNumId]?.levels[targetLevelIndex];
      if (targetLevelDef) {
        const counterNameForPlaceholder = `docx-counter-${abstractNumId}-${targetLevelIndex}`;
        const formatForPlaceholder = getCSSCounterFormat(targetLevelDef.format);
        if (formatForPlaceholder === "none") {
          // If a referenced level is 'none', it shouldn't produce output.
          // We might need to adjust logic if " + '' + " is problematic.
          // For now, let's output an empty string if it's not the only segment.
          if (levelDef.parsedFormat.segments.length > 1) contentValue += "''";
          else {
            /* if it's the only segment and none, the whole thing is blank */
          }
        } else {
          contentValue += `counter(${counterNameForPlaceholder}, ${formatForPlaceholder})`;
        }
      } else {
        if (levelDef.parsedFormat.segments.length > 1) contentValue += "''"; // Placeholder for non-existent level
      }
    }
  });

  // If the whole thing ended up empty (e.g. format was 'none' and no literals)
  if (
    contentValue.replace(/\s*\+\s*''/g, "").trim() === "''" ||
    contentValue.trim() === ""
  ) {
    if (levelDef.format === "bullet") return "'•'"; // Default bullet char if format is bullet but text is empty
    return "''"; // No visual number
  }

  return contentValue;
}

module.exports = {
  parseNumberingDefinitions,
  parseNumberingTextFormat,
  parseParagraphPropertiesForNumbering,
  parseRunPropertiesForNumbering,
  parseLevelDefinition,
  getCSSCounterFormat,
  getCSSCounterContent,
};
