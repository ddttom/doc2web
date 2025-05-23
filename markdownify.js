// markdownify.js
const { JSDOM } = require("jsdom");

/**
 * Convert HTML to Markdown
 * @param {string} html - HTML content from mammoth
 * @returns {string} - Converted markdown
 */
function markdownify(html) {
  const dom = new JSDOM(html, {
    contentType: "text/html; charset=utf-8",
  });
  const document = dom.window.document;
  let markdown = "";

  // Process the body element recursively
  markdown = processNode(document.body, 0);

  // Clean up extra whitespace and line breaks and fix linting issues
  markdown = markdown
    .replace(/\n{3,}/g, "\n\n") // Reduce multiple line breaks
    .replace(/\t/g, "  ") // Replace tabs with spaces
    .replace(/[ \t]+$/gm, "") // Remove trailing spaces (MD009)
    .replace(/\n(#+)\s+(.*?)[:.]\s*$/gm, "\n$1 $2") // Remove trailing punctuation in headings (MD026)
    .replace(/\[\]\(\)/g, "[placeholder](#)") // Fix empty links (MD042)
    .replace(/\[placeholder\]\(\)/g, "[placeholder](#)") // Fix any remaining empty links
    .trim();

  // Fix spaces inside emphasis markers (MD037) - more thorough approach
  markdown = fixEmphasisSpacing(markdown);

  // Fix heading increment issues (MD001)
  markdown = fixHeadingIncrements(markdown);

  // Fix list marker spacing (MD030) and ordered list prefixes (MD029)
  markdown = fixListFormatting(markdown);

  // Fix blanks around headings (MD022)
  markdown = fixBlanksAroundHeadings(markdown);

  // Ensure first line is a top-level heading (MD041)
  markdown = ensureFirstLineHeading(markdown);

  return markdown;
}

/**
 * Fix spaces inside emphasis markers (MD037)
 * @param {string} markdown - Markdown content
 * @returns {string} - Fixed markdown content
 */
function fixEmphasisSpacing(markdown) {
  // Fix spaces inside single asterisk emphasis
  // *text* is correct, * text* or *text * or * text * are incorrect
  let result = markdown.replace(/\*(\s+)([^*]+)(\s*)\*/g, "*$2*");
  result = result.replace(/\*([^*]+)(\s+)\*/g, "*$1*");

  // Fix spaces inside double asterisk (strong emphasis)
  // **text** is correct, ** text** or **text ** or ** text ** are incorrect
  result = result.replace(/\*\*(\s+)([^*]+)(\s*)\*\*/g, "**$2**");
  result = result.replace(/\*\*([^*]+)(\s+)\*\*/g, "**$1**");

  // Fix spaces inside underscore emphasis
  // _text_ is correct, _ text_ or _text _ or _ text _ are incorrect
  result = result.replace(/_(\s+)([^_]+)(\s*)_/g, "_$2_");
  result = result.replace(/_([^_]+)(\s+)_/g, "_$1_");

  // Fix spaces inside double underscore (strong emphasis)
  // __text__ is correct, __ text__ or __text __ or __ text __ are incorrect
  result = result.replace(/__(\s+)([^_]+)(\s*)__/g, "__$2__");
  result = result.replace(/__([^_]+)(\s+)__/g, "__$1__");

  return result;
}

/**
 * Ensure the first line in the file is a top-level heading (MD041)
 * @param {string} markdown - Markdown content
 * @returns {string} - Fixed markdown content
 */
function ensureFirstLineHeading(markdown) {
  const lines = markdown.split("\n");

  // Find the first non-blank line
  let firstNonBlankIndex = 0;
  while (
    firstNonBlankIndex < lines.length &&
    lines[firstNonBlankIndex].trim() === ""
  ) {
    firstNonBlankIndex++;
  }

  // If there are no non-blank lines, return the original markdown
  if (firstNonBlankIndex >= lines.length) {
    return markdown;
  }

  const firstNonBlankLine = lines[firstNonBlankIndex];

  // Check if the first non-blank line is already a top-level heading
  if (/^# .+/.test(firstNonBlankLine)) {
    // Ensure there's only one top-level heading (MD025)
    return ensureSingleTopLevelHeading(markdown);
  }

  // If the first non-blank line is a heading but not level 1, convert it to level 1
  if (/^#{2,6} .+/.test(firstNonBlankLine)) {
    const headingContent = firstNonBlankLine.replace(/^#{2,6} /, "");
    lines[firstNonBlankIndex] = `# ${headingContent}`;
    return ensureSingleTopLevelHeading(lines.join("\n"));
  }

  // If the first non-blank line is not a heading, add a document title heading
  const title = "Document";
  lines.splice(firstNonBlankIndex, 0, `# ${title}`, "");
  return ensureSingleTopLevelHeading(lines.join("\n"));
}

/**
 * Ensure there's only one top-level heading in the document (MD025)
 * @param {string} markdown - Markdown content
 * @returns {string} - Fixed markdown content
 */
function ensureSingleTopLevelHeading(markdown) {
  const lines = markdown.split("\n");
  let foundFirstH1 = false;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    // Check if this line is a top-level heading
    if (/^# .+/.test(line)) {
      if (!foundFirstH1) {
        // This is the first top-level heading, keep it
        foundFirstH1 = true;
      } else {
        // This is a subsequent top-level heading, demote it to H2
        lines[i] = line.replace(/^# /, "## ");
      }
    }
  }

  return lines.join("\n");
}

/**
 * Fix blank lines around headings (MD022)
 * @param {string} markdown - Markdown content
 * @returns {string} - Fixed markdown content
 */
function fixBlanksAroundHeadings(markdown) {
  const lines = markdown.split("\n");
  const result = [];

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const isHeading = /^#+\s+.+$/.test(line);

    if (isHeading) {
      // If the previous line is not blank and this is not the first line, add a blank line
      if (i > 0 && result[result.length - 1] !== "") {
        result.push("");
      }

      // Add the heading
      result.push(line);

      // If the next line is not blank and this is not the last line, add a blank line
      if (i < lines.length - 1 && lines[i + 1] !== "") {
        result.push("");
      }
    } else {
      result.push(line);
    }
  }

  return result.join("\n");
}

/**
 * Fix list formatting issues (MD030 and MD029)
 * @param {string} markdown - Markdown content
 * @returns {string} - Fixed markdown content
 */
function fixListFormatting(markdown) {
  const lines = markdown.split("\n");
  let inOrderedList = false;
  let orderedListCounter = 1;

  for (let i = 0; i < lines.length; i++) {
    // Fix unordered list marker spacing (MD030)
    // Match: indentation + list marker (* or - or +) + more than one space + content
    let match = lines[i].match(/^(\s*)([\*\-\+])(\s{2,})(.+)$/);
    if (match) {
      // Replace with: indentation + list marker + exactly one space + content
      lines[i] = `${match[1]}${match[2]} ${match[4]}`;
    }

    // Fix ordered list marker spacing and prefixes (MD030 and MD029)
    // Match: indentation + number + period + more than one space + content
    match = lines[i].match(/^(\s*)(\d+)\.(\s{2,})(.+)$/);
    if (match) {
      // If this is the start of a new ordered list or continuing an existing one
      if (!inOrderedList) {
        inOrderedList = true;
        orderedListCounter = 1;
      }

      // Replace with: indentation + sequential number + period + exactly one space + content
      lines[i] = `${match[1]}${orderedListCounter}. ${match[4]}`;
      orderedListCounter++;
    } else {
      // If this line is not an ordered list item, reset the counter
      if (!/^\s*\d+\.\s+/.test(lines[i])) {
        inOrderedList = false;
      }
    }
  }

  return lines.join("\n");
}

/**
 * Fix heading increment issues to ensure they only increment by one level
 * @param {string} markdown - Markdown content
 * @returns {string} - Fixed markdown content
 */
function fixHeadingIncrements(markdown) {
  const lines = markdown.split("\n");
  let currentLevel = 0;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const headingMatch = line.match(/^(#+)\s+(.*)$/);

    if (headingMatch) {
      const level = headingMatch[1].length;
      const content = headingMatch[2];

      // If this is the first heading, or if the level is only one more than the current level, keep it
      if (currentLevel === 0 || level <= currentLevel + 1) {
        currentLevel = level;
      } else {
        // Otherwise, adjust the heading level to be only one more than the current level
        const newLevel = currentLevel + 1;
        lines[i] = "#".repeat(newLevel) + " " + content;
        currentLevel = newLevel;
      }
    }
  }

  return lines.join("\n");
}

/**
 * Process an HTML node and its children to markdown
 * @param {Node} node - DOM node
 * @param {number} depth - Current nesting depth for lists
 * @returns {string} - Markdown output
 */
function processNode(node, depth = 0) {
  if (!node) return "";
  let result = "";

  // Handle text nodes
  if (node.nodeType === 3) {
    // Text node
    return node.textContent;
  }

  // Handle element nodes
  if (node.nodeType === 1) {
    // Element node
    const nodeName = node.nodeName.toLowerCase();
    const children = Array.from(node.childNodes);

    // Check for directionality
    const dir = node.getAttribute("dir");
    const isRTL = dir === "rtl";

    // Check for placeholder class
    const isPlaceholder =
      node.classList && node.classList.contains("docx-placeholder");
    if (isPlaceholder) {
      return `\n\n**${node.textContent.trim()}**\n\n`;
    }

    // IMPROVED: Check for TOC-specific classes
    const isToc =
      node.classList &&
      (node.classList.contains("docx-toc") ||
        node.classList.contains("docx-toc-entry") ||
        node.classList.contains("docx-toc-item"));

    if (isToc) {
      // Handle TOC entry
      const level = 1;
      // Check if there's a level class
      if (node.classList) {
        for (const className of node.classList) {
          if (className.startsWith("docx-toc-level-")) {
            const levelMatch = className.match(/docx-toc-level-(\d+)/);
            if (levelMatch) {
              level = parseInt(levelMatch[1], 10);
            }
          }
        }
      }

      // Get the text and page number if available
      let text = "";
      let pageNum = "";

      const textSpan = node.querySelector(".docx-toc-text");
      const pageSpan = node.querySelector(".docx-toc-pagenum");

      if (textSpan && pageSpan) {
        text = textSpan.textContent.trim();
        pageNum = pageSpan.textContent.trim();
        result += `${text.padEnd(60, ".")} ${pageNum}\n`;
      } else {
        // Just use the node's text content if structured spans aren't found
        result += `${node.textContent.trim()}\n`;
      }

      return result;
    }

    // Process different element types
    switch (nodeName) {
      case "h1":
        result += `# ${processChildren(node)}\n\n`;
        break;
      case "h2":
        result += `## ${processChildren(node)}\n\n`;
        break;
      case "h3":
        result += `### ${processChildren(node)}\n\n`;
        break;
      case "h4":
        result += `#### ${processChildren(node)}\n\n`;
        break;
      case "h5":
        result += `##### ${processChildren(node)}\n\n`;
        break;
      case "h6":
        result += `###### ${processChildren(node)}\n\n`;
        break;
      case "p":
        // Handle right-to-left text direction
        if (isRTL) {
          result += `<div dir="rtl">\n\n${processChildren(node)}\n\n</div>\n\n`;
        } else {
          result += `${processChildren(node)}\n\n`;
        }
        break;
      case "strong":
      case "b":
        result += `**${processChildren(node)}**`;
        break;
      case "em":
      case "i":
        result += `*${processChildren(node)}*`;
        break;
      case "u":
        result += `<u>${processChildren(node)}</u>`;
        break;
      case "strike":
      case "s":
      case "del":
        result += `~~${processChildren(node)}~~`;
        break;
      case "code":
        result += `\`${processChildren(node)}\``;
        break;
      case "pre":
        result += `\`\`\`\n${processChildren(node)}\n\`\`\`\n\n`;
        break;
      case "a":
        const href = node.getAttribute("href") || "";
        // Escape special characters in URLs that would break markdown syntax
        const escapedHref = href
          .replace(/\(/g, "%28")
          .replace(/\)/g, "%29")
          .replace(/\s/g, "%20")
          .replace(/\[/g, "%5B")
          .replace(/\]/g, "%5D")
          .replace(/&/g, "%26")
          .replace(/;/g, "%3B")
          .replace(/,/g, "%2C")
          .replace(/'/g, "%27")
          .replace(/"/g, "%22")
          .replace(/</g, "%3C")
          .replace(/>/g, "%3E");

        // If the link text is empty, use the URL as the text
        const linkText = processChildren(node) || escapedHref;
        result += `[${linkText}](${escapedHref})`;
        break;
      case "img":
        const src = node.getAttribute("src") || "";
        const alt = node.getAttribute("alt") || "";
        result += `![${alt}](${src})`;
        break;
      case "blockquote":
        // Split into lines and add > to each line
        const blockquoteContent = processChildren(node)
          .split("\n")
          .map((line) => `> ${line}`)
          .join("\n");
        result += `${blockquoteContent}\n\n`;
        break;
      case "ul":
        result += processListItems(node, "*", depth);
        break;
      case "ol":
        // IMPROVED: Check if this is a TOC list and process accordingly
        if (node.classList && node.classList.contains("docx-toc-list")) {
          result += processTocList(node, depth);
        } else {
          result += processListItems(node, "1.", depth);
        }
        break;
      case "li":
        // This is handled by processListItems
        result += processChildren(node);
        break;
      case "table":
        result += processTable(node);
        break;
      case "hr":
        result += `\n---\n\n`;
        break;
      case "br":
        result += `\n`;
        break;
      case "div":
        // IMPROVED: Special handling for TOC divs
        if (node.classList && node.classList.contains("docx-toc")) {
          result += processTocDiv(node);
        } else {
          // For other divs, just process children
          result += processChildren(node) + "\n\n";
        }
        break;
      default:
        // For other elements, just process children
        if (children.length > 0) {
          result += processChildren(node);
        }
    }
  }

  return result;
}

/**
 * NEW: Process TOC div into markdown
 * @param {Node} tocNode - TOC container node
 * @returns {string} - Markdown TOC
 */
function processTocDiv(tocNode) {
  let result = "\n\n";
  const entries = tocNode.querySelectorAll(".docx-toc-entry");

  entries.forEach((entry) => {
    // Get level from class
    let level = 1;
    for (const className of entry.classList) {
      if (className.startsWith("docx-toc-level-")) {
        const levelMatch = className.match(/docx-toc-level-(\d+)/);
        if (levelMatch) {
          level = parseInt(levelMatch[1], 10);
        }
      }
    }

    // Determine appropriate indentation
    const indent = "  ".repeat(level - 1);

    // Get text content and page number
    const textSpan = entry.querySelector(".docx-toc-text");
    const pageSpan = entry.querySelector(".docx-toc-pagenum");

    if (textSpan && pageSpan) {
      const text = textSpan.textContent.trim();
      const pageNum = pageSpan.textContent.trim();

      // Format with dots for padding (60 chars total width for text + dots)
      result += `${indent}${text}${".".repeat(
        Math.max(1, 60 - indent.length - text.length)
      )} ${pageNum}\n`;
    } else {
      // Fallback to just using the entry's text content
      result += `${indent}${entry.textContent.trim()}\n`;
    }
  });

  return result + "\n";
}

/**
 * NEW: Process TOC list into markdown
 * @param {Node} listNode - TOC list node
 * @param {number} depth - Nesting depth
 * @returns {string} - Markdown TOC
 */
function processTocList(listNode, depth) {
  let result = "\n\n";
  const items = Array.from(listNode.querySelectorAll(":scope > li"));

  items.forEach((item) => {
    // Check if this is a main or sub item based on class or prefix
    let isSubItem =
      item.classList && item.classList.contains("docx-sublist-item");
    let prefix = item.getAttribute("data-prefix") || "";

    // Determine if it's a main or sub item based on the prefix if not determined by class
    if (!isSubItem && prefix) {
      isSubItem = /^[a-z]$/.test(prefix); // If prefix is a single lowercase letter, it's a sub-item
    }

    // Calculate indentation
    const indent = isSubItem ? "  ".repeat(depth + 1) : "  ".repeat(depth);

    // Process the item's content
    const itemContent = processChildren(item);

    // Check if this is a TOC item with page number
    const tocSpans = item.querySelectorAll(
      "span.docx-toc-text, span.docx-toc-pagenum"
    );
    if (tocSpans.length > 0) {
      const textSpan = item.querySelector(".docx-toc-text");
      const pageSpan = item.querySelector(".docx-toc-pagenum");

      if (textSpan && pageSpan) {
        const text = textSpan.textContent.trim();
        const pageNum = pageSpan.textContent.trim();

        // Format with dots for padding (60 chars total width for text + dots)
        result += `${indent}${prefix}. ${text}${".".repeat(
          Math.max(1, 50 - text.length)
        )} ${pageNum}\n`;
      } else {
        // Fallback
        result += `${indent}${prefix}. ${itemContent}\n`;
      }
    } else {
      // Regular list item
      result += `${indent}${prefix}. ${itemContent}\n`;
    }
  });

  return result + "\n";
}

/**
 * Process child nodes
 * @param {Node} node - Parent node
 * @returns {string} - Processed content
 */
function processChildren(node) {
  let result = "";
  const children = Array.from(node.childNodes);

  children.forEach((child) => {
    result += processNode(child);
  });

  return result;
}

/**
 * Process list items with proper indentation
 * @param {Node} listNode - UL or OL node
 * @param {string} marker - List item marker (* or 1.)
 * @param {number} depth - Nesting depth
 * @returns {string} - Processed list
 */
function processListItems(listNode, marker, depth) {
  // Add a blank line before the list (MD032)
  let result = "\n\n";
  const items = Array.from(listNode.querySelectorAll(":scope > li"));

  // IMPROVED: Check if this is specifically a TOC list using classes
  const isTocList =
    listNode.classList &&
    (listNode.classList.contains("docx-toc-list") ||
      (listNode.classList.contains("docx-alpha-list") &&
        items.some(
          (item) =>
            item.textContent.includes("\t") || /\d+\s*$/.test(item.textContent)
        )));

  items.forEach((item, index) => {
    const indent = "  ".repeat(depth);

    // Get the prefix from data attribute if available (for preserving original numbering)
    const prefix = item.getAttribute("data-prefix");

    // IMPROVED: Detect if this is a nested sublist within a TOC
    const isSublistItem =
      item.classList && item.classList.contains("docx-sublist-item");

    // For TOC lists, try to format with dots and page numbers
    if (isTocList) {
      const itemContent = processNode(item, depth + 1);

      // Try to extract page number if it exists
      const pageMatch = itemContent.match(/(\d+)\s*$/);
      if (pageMatch) {
        const content = itemContent
          .substring(0, itemContent.lastIndexOf(pageMatch[1]))
          .trim();
        const pageNum = pageMatch[1];

        // Use dots to connect content to page number (like a TOC)
        const dotsCount = Math.max(1, 60 - indent.length - content.length);
        result += `${indent}${
          prefix || (marker === "1." ? index + 1 + "." : marker)
        } ${content}${".".repeat(dotsCount)} ${pageNum}\n`;
      } else {
        // Just use normal list item formatting
        result += `${indent}${
          prefix || (marker === "1." ? index + 1 + "." : marker)
        } ${itemContent}\n`;
      }
    } else {
      // Regular list item
      const actualMarker = marker === "1." ? `${index + 1}.` : marker;
      const itemContent = processNode(item, depth + 1);

      // Use exactly one space after the marker (MD030)
      result += `${indent}${
        prefix ? prefix + "." : actualMarker
      } ${itemContent}\n`;
    }
  });

  // Add a blank line after the list (MD032)
  return result + "\n";
}

/**
 * Process table to markdown
 * @param {Node} tableNode - Table node
 * @returns {string} - Markdown table
 */
function processTable(tableNode) {
  let result = "\n";
  const rows = Array.from(tableNode.querySelectorAll("tr"));

  if (rows.length === 0) return "";

  // Process header row
  const headerCells = Array.from(rows[0].querySelectorAll("th, td"));
  if (headerCells.length === 0) return "";

  // Create header row
  const headerTexts = headerCells.map((cell) => processChildren(cell).trim());
  result += `| ${headerTexts.join(" | ")} |\n`;

  // Create separator row
  result += `| ${headerTexts.map(() => "---").join(" | ")} |\n`;

  // Process data rows
  for (let i = 1; i < rows.length; i++) {
    const cells = Array.from(rows[i].querySelectorAll("td"));
    if (cells.length === 0) continue;

    const cellTexts = cells.map((cell) => processChildren(cell).trim());
    result += `| ${cellTexts.join(" | ")} |\n`;
  }

  return result + "\n";
}

module.exports = { markdownify };
