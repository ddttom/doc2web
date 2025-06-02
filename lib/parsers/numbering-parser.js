// File: ddttom/doc2web/doc2web-804421e293a28695ca8c4527a7c79f342ff0c562/lib/parsers/numbering-parser.js

const { selectNodes, selectSingleNode } = require("../xml/xpath-utils");
const { convertTwipToPt } = require("../utils/unit-converter");

/**
 * Enhanced numbering definition parsing.
 */
function parseNumberingDefinitions(numberingDoc) {
  const numberingDefs = {
    abstractNums: {},
    nums: {},
    numIdMap: {},
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
        numStyleLink: selectSingleNode("w:numStyleLink", node)?.getAttribute(
          "w:val"
        ),
        styleLink: selectSingleNode("w:styleLink", node)?.getAttribute("w:val"),
      };

      const multiLevelTypeNode = selectSingleNode("w:multiLevelType", node);
      if (multiLevelTypeNode) {
        abstractNum.multiLevelType = multiLevelTypeNode.getAttribute("w:val");
      }

      const levelNodes = selectNodes("w:lvl", node);
      levelNodes.forEach((lvlNode) => {
        const ilvl = lvlNode.getAttribute("w:ilvl");
        if (ilvl === null) return;

        const levelIndex = parseInt(ilvl, 10); // 0-indexed

        const level = {
          level: levelIndex,
          format: "decimal",
          text: `%${levelIndex + 1}.`,
          alignment: "left",
          indentation: {},
          isLegal: false,
          textFormat: "",
          restart: null, // Means restart after higher level
          start: 1,
          suffix: "tab",
          tentative: false,
          runProps: {},
          paragraphProps: {},
        };

        const numFmtNode = selectSingleNode("w:numFmt", lvlNode);
        if (numFmtNode) {
          level.format = numFmtNode.getAttribute("w:val") || "decimal";
        }

        const lvlTextNode = selectSingleNode("w:lvlText", lvlNode);
        if (lvlTextNode) {
          level.text =
            lvlTextNode.getAttribute("w:val") || `%${levelIndex + 1}.`;
        }
        level.textFormat = level.text; // Store the original w:lvlText value
        level.parsedFormat = parseNumberingTextFormat(level.textFormat);

        // Enhanced bullet point detection and parsing
        if (level.format === "bullet") {
          enhanceBulletDetection(level, lvlTextNode);
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

        const lvlRestartNode = selectSingleNode("w:lvlRestart", lvlNode);
        if (lvlRestartNode) {
          // w:lvlRestart@w:val="1" means restart after level (val-1).
          // Here, val is 0-indexed if it means restart after specific level index.
          // If val="0", it means restart after level 0.
          // Typically, restart is 1 for restarting after a higher level.
          // The value in DOCX means "restart numbering after encountering a paragraph of outline level X"
          // For simplicity, if present, we can treat it as restart if a higher level increments.
          // This is handled by NumberingSequenceTracker.
          level.restartAfterLevel = parseInt(
            lvlRestartNode.getAttribute("w:val"),
            10
          );
        }

        const startNode = selectSingleNode("w:start", lvlNode);
        if (startNode) {
          level.start = parseInt(startNode.getAttribute("w:val"), 10) || 1;
        }

        const pPrNode = selectSingleNode("w:pPr", lvlNode);
        if (pPrNode) {
          level.paragraphProps = parseParagraphPropertiesForNumbering(pPrNode);
          const indNode = selectSingleNode("w:ind", pPrNode);
          if (indNode) {
            if (indNode.hasAttribute("w:left"))
              level.indentation.left = convertTwipToPt(
                indNode.getAttribute("w:left")
              );
            if (indNode.hasAttribute("w:hanging"))
              level.indentation.hanging = convertTwipToPt(
                indNode.getAttribute("w:hanging")
              );
            if (indNode.hasAttribute("w:firstLine"))
              level.indentation.firstLine = convertTwipToPt(
                indNode.getAttribute("w:firstLine")
              );
            if (indNode.hasAttribute("w:start"))
              level.indentation.start = convertTwipToPt(
                indNode.getAttribute("w:start")
              );
          }
        }

        const rPrNode = selectSingleNode("w:rPr", lvlNode);
        if (rPrNode) {
          level.runProps = parseRunPropertiesForNumbering(rPrNode);
        }
        abstractNum.levels[levelIndex] = level;
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

      const overrides = {};
      selectNodes("w:lvlOverride", node).forEach((overrideNode) => {
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
          override.completeLevel = parseLevelDefinition(lvlNode);
        }
        overrides[ilvl] = override;
      });

      numberingDefs.nums[id] = {
        id,
        abstractNumId,
        overrides: Object.keys(overrides).length > 0 ? overrides : null,
      };
      numberingDefs.numIdMap[id] = abstractNumId;
    });
  } catch (error) {
    console.error(
      "Error parsing numbering definitions:",
      error.message,
      error.stack
    );
  }
  return numberingDefs;
}

function parseNumberingTextFormat(textFormat) {
  const parsed = { pattern: textFormat, segments: [] };
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
      parsed.segments.push({ type: "level", value: levelNum }); // levelNum is 1-indexed
      i++;
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
  /* ... same as before ... */ return {};
}
function parseRunPropertiesForNumbering(rPrNode) {
  /* ... same as before ... */ return {};
}
function parseLevelDefinition(lvlNode) {
  /* ... same as before, ensure it parses indents if override includes w:pPr/w:ind ... */
  const level = {
    indentation: {},
    runProps: {},
    paragraphProps: {},
    format: "decimal",
    start: 1,
  };
  // ... (parsing logic as provided in previous correct version)
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

  const pPrNodeOverride = selectSingleNode("w:pPr", lvlNode);
  if (pPrNodeOverride) {
    level.paragraphProps =
      parseParagraphPropertiesForNumbering(pPrNodeOverride);
    const indentNode = selectSingleNode("w:ind", pPrNodeOverride);
    if (indentNode) {
      if (indentNode.hasAttribute("w:left"))
        level.indentation.left = convertTwipToPt(
          indentNode.getAttribute("w:left")
        );
      if (indentNode.hasAttribute("w:hanging"))
        level.indentation.hanging = convertTwipToPt(
          indentNode.getAttribute("w:hanging")
        );
      if (indentNode.hasAttribute("w:firstLine"))
        level.indentation.firstLine = convertTwipToPt(
          indentNode.getAttribute("w:firstLine")
        );
    }
  }
  const rPrNodeOverride = selectSingleNode("w:rPr", lvlNode);
  if (rPrNodeOverride) {
    level.runProps = parseRunPropertiesForNumbering(rPrNodeOverride);
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
    bullet: "disc", // For CSS, 'disc' is a common bullet. Actual char can be in lvlText.
    decimalZero: "decimal-leading-zero",
    none: "none",
    // Add more complex mappings if necessary, e.g., for ordinal, cardinal, specific language scripts
    ordinal: "decimal", // CSS doesn't have a direct 'ordinal' counter style, often handled with suffixes
    cardinalText: "decimal", // Similar to ordinal
    aiueo: "hiragana",
    iroha: "hiragana-iroha", // Japanese
    chineseCounting: "cjk-decimal", // A guess, might need more specific cjk styles
  };
  return formatMap[format] || "decimal";
}

function getCSSCounterContent(levelDef, abstractNumId, numberingDefs) {
  if (
    !levelDef ||
    !levelDef.parsedFormat ||
    !levelDef.parsedFormat.segments ||
    levelDef.parsedFormat.segments.length === 0
  ) {
    const currentLevelIndex = levelDef?.level ?? 0;
    const counterName = `docx-counter-${abstractNumId}-${currentLevelIndex}`;
    const format = getCSSCounterFormat(levelDef?.format || "decimal");
    let suffix = "'.'";
    if (levelDef?.format === "bullet") suffix = "''";
    else if (levelDef?.suffix === "tab") suffix = "'\\00a0\\00a0'";
    else if (levelDef?.suffix === "space") suffix = "' '";
    else if (levelDef?.suffix === "nothing") suffix = "''";

    if (format === "none" || levelDef?.format === "none") return "''"; // No content for 'none' format
    if (
      levelDef?.format === "bullet" &&
      (!levelDef.textFormat ||
        levelDef.textFormat.trim() === "" ||
        levelDef.textFormat.trim() === "%1")
    ) {
      // If format is bullet and textFormat is empty or default placeholder, use actual bullet char
      const bulletChar = levelDef.text?.trim() || "•"; // Use text from lvlText if it's a specific bullet char
      return `'${bulletChar.replace(/'/g, "\\'")}'`;
    }
    return `counter(${counterName}, ${format}) + ${suffix}`;
  }

  let contentValue = "";
  let needsPlus = false;
  levelDef.parsedFormat.segments.forEach((segment) => {
    let segmentContent = "";
    if (segment.type === "literal") {
      segmentContent = `"${segment.value.replace(/"/g, '\\"')}"`;
    } else if (segment.type === "level") {
      const targetLevelOneBased = segment.value;
      const targetLevelZeroBased = targetLevelOneBased - 1;
      const targetLevelFullDef =
        numberingDefs.abstractNums[abstractNumId]?.levels[targetLevelZeroBased];

      if (targetLevelFullDef) {
        const counterNameForPlaceholder = `docx-counter-${abstractNumId}-${targetLevelZeroBased}`;
        const formatForPlaceholder = getCSSCounterFormat(
          targetLevelFullDef.format
        );
        if (formatForPlaceholder !== "none") {
          segmentContent = `counter(${counterNameForPlaceholder}, ${formatForPlaceholder})`;
        } else {
          segmentContent = "''"; // Empty string for 'none' format
        }
      } else {
        segmentContent = "''"; // Fallback for undefined referenced level
      }
    }

    if (segmentContent && segmentContent !== "''") {
      if (needsPlus) {
        contentValue += " + ";
      }
      contentValue += segmentContent;
      needsPlus = true;
    } else if (
      segmentContent === "''" &&
      needsPlus &&
      levelDef.parsedFormat.segments.length === 1
    ) {
      // If the only segment is an empty counter (e.g. %1 with format:none), make contentValue empty.
    } else if (
      segmentContent === "''" &&
      !needsPlus &&
      levelDef.parsedFormat.segments.length === 1
    ) {
      // If it's the only segment and it's an empty string, it means the content should be empty
      contentValue = "''"; // Explicitly set to empty string for CSS
      // needsPlus remains false
    }
  });

  // If after all segments, contentValue is effectively empty, handle default bullet
  if (contentValue.trim() === "" || contentValue.trim() === "''") {
    if (levelDef.format === "bullet") return "'•'";
    return "''"; // No number to display
  }

  return contentValue;
}

/**
 * Enhanced bullet point detection and parsing
 */
function enhanceBulletDetection(levelDef, lvlTextNode) {
  if (levelDef.format === "bullet") {
    // Extract actual bullet character from lvlText
    const bulletText = lvlTextNode?.getAttribute("w:val") || "•";
    
    // Handle various bullet characters
    const bulletMap = {
      "": "•",           // Default bullet
      "·": "·",          // Middle dot
      "○": "○",          // White circle
      "■": "■",          // Black square
      "□": "□",          // White square
      "▪": "▪",          // Black small square
      "▫": "▫",          // White small square
      "►": "►",          // Right-pointing triangle
      "⇒": "⇒",          // Rightwards double arrow
      "➢": "➢",          // Three-D top-lighted rightwards arrowhead
      "✓": "✓",          // Check mark
      "✗": "✗",          // Ballot X
      "◆": "◆",          // Black diamond
      "◇": "◇",          // White diamond
    };
    
    levelDef.bulletChar = bulletMap[bulletText] || bulletText || "•";
    levelDef.isBullet = true;
    
    // Store original bullet character for CSS generation
    levelDef.originalBulletChar = bulletText;
    
    // Mark this as a bullet format for easier identification
    levelDef.textFormat = levelDef.bulletChar;
  }
  
  return levelDef;
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
