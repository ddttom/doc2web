// File: lib/html/generators/html-formatting.js
// HTML formatting and structure utilities

/**
 * Format HTML with proper indentation and line breaks for better readability
 * @param {string} html - The HTML string to format
 * @returns {string} - Formatted HTML string
 */
function formatHtml(html) {
  try {
    // First, normalize and break up the HTML into proper lines
    html = preprocessHtml(html);
    
    // Split into lines and prepare for formatting
    let lines = html.split("\n");
    let formattedLines = [];
    
    // Process each line - remove indentation but keep line structure
    for (let i = 0; i < lines.length; i++) {
      let line = lines[i].trim();
      if (!line) continue;

      // Add the line without any indentation
      formattedLines.push(line);
    }

    // Join formatted lines and add final newline
    return formattedLines.join("\n") + "\n";
  } catch (error) {
    console.error("Error formatting HTML:", error.message, error.stack);
    return html; // Return original on error
  }
}

/**
 * Preprocess HTML to break it into proper lines for formatting
 * @param {string} html - The HTML string to preprocess
 * @returns {string} - HTML with proper line breaks
 */
function preprocessHtml(html) {
  // Normalize self-closing tags
  html = html.replace(/<([^>]+?)([^\s/])\/>/g, "<$1$2 />");
  
  // Add line breaks after closing tags
  html = html.replace(/<\/([^>]+)>/g, "</$1>\n");
  
  // Add line breaks before opening tags (but not if preceded by another tag)
  html = html.replace(/(?<!>)\s*<([^/!][^>]*?)>/g, "\n<$1>");
  
  // Add line breaks after opening tags for block elements
  const blockElements = [
    "html", "head", "body", "div", "p", "h1", "h2", "h3", "h4", "h5", "h6",
    "ul", "ol", "li", "table", "thead", "tbody", "tfoot", "tr", "td", "th",
    "nav", "section", "article", "header", "footer", "aside", "main",
    "figure", "figcaption", "blockquote", "pre", "form", "fieldset",
    "dl", "dt", "dd", "script", "style", "title", "caption"
  ];
  
  blockElements.forEach(tag => {
    // Add line break after opening tag
    html = html.replace(new RegExp(`<${tag}([^>]*?)>(?!\n)`, "gi"), `<${tag}$1>\n`);
  });
  
  // Special handling for meta tags, links, and other self-closing elements
  html = html.replace(/<(meta|link|br|hr|img|input|area|base|col|embed|param|source|track|wbr)([^>]*?)>/gi, "<$1$2>\n");
  
  // Clean up multiple consecutive newlines
  html = html.replace(/\n\s*\n\s*\n/g, "\n\n");
  
  // Clean up leading/trailing whitespace on lines
  html = html.split("\n").map(line => line.trim()).join("\n");
  
  // Remove empty lines at the beginning
  html = html.replace(/^\n+/, "");
  
  return html;
}

module.exports = {
  formatHtml,
  preprocessHtml,
};
