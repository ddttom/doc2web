// File: lib/html/generators/html-formatting.js
// HTML formatting and structure utilities

/**
 * Format HTML with proper indentation and line breaks for better readability
 * @param {string} html - The HTML string to format
 * @returns {string} - Formatted HTML string
 */
function formatHtml(html) {
  try {
    // Normalize self-closing tags for consistency
    html = html.replace(/<([^>]+?)([^\s/])\/>/g, "<$1$2 />");

    // Split into lines and prepare for formatting
    let lines = html.split("\n");
    let formattedLines = [];
    let indentLevel = 0;
    const indentSize = 2;
    
    // Define block elements that affect indentation
    const blockElements = new Set([
      "div", "p", "h1", "h2", "h3", "h4", "h5", "h6", 
      "ul", "ol", "li", "table", "thead", "tbody", "tfoot", "tr", "td", "th",
      "nav", "section", "article", "header", "footer", "aside", "main",
      "figure", "figcaption", "blockquote", "pre", "form", "fieldset",
      "dl", "dt", "dd"
    ]);
    
    // Define self-closing elements
    const selfClosingElements = new Set([
      "area", "base", "br", "col", "embed", "hr", "img", "input",
      "link", "meta", "param", "source", "track", "wbr"
    ]);

    // Process each line
    for (let i = 0; i < lines.length; i++) {
      let line = lines[i].trim();
      if (!line) continue;

      // Check for closing tags
      let closingTagMatch = line.match(/^<\/\s*([\w-]+)\s*>/);
      let openingTagMatch = line.match(/^<\s*([\w-]+)([^>]*)>/);

      // Decrease indent for closing tags
      if (closingTagMatch) {
        const tagName = closingTagMatch[1].toLowerCase();
        if (blockElements.has(tagName) && indentLevel > 0) {
          indentLevel--;
        }
      }

      // Add the line with proper indentation
      formattedLines.push(" ".repeat(indentLevel * indentSize) + line);

      // Increase indent for opening tags
      if (openingTagMatch) {
        const tagName = openingTagMatch[1].toLowerCase();
        const attributes = openingTagMatch[2];
        const isSelfClosingInTag = attributes.endsWith("/") || selfClosingElements.has(tagName);

        // Only increase indent for block elements with content
        if (
          blockElements.has(tagName) &&
          !isSelfClosingInTag &&
          !line.includes(`</${tagName}>`) // Don't indent if closing tag on same line
        ) {
          indentLevel++;
        }
      }
    }

    // Join formatted lines and add final newline
    return formattedLines.join("\n") + "\n";
  } catch (error) {
    console.error("Error formatting HTML:", error.message, error.stack);
    return html; // Return original on error
  }
}

module.exports = {
  formatHtml,
};