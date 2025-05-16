// markdownify.js
const { JSDOM } = require('jsdom');

/**
 * Convert HTML to Markdown
 * @param {string} html - HTML content from mammoth
 * @returns {string} - Converted markdown
 */
function markdownify(html) {
  const dom = new JSDOM(html, { 
    contentType: 'text/html; charset=utf-8' 
  });
  const document = dom.window.document;
  let markdown = '';

  // Process the body element recursively
  markdown = processNode(document.body, 0);
  
  // Clean up extra whitespace and line breaks
  markdown = markdown
    .replace(/\n{3,}/g, '\n\n')  // Reduce multiple line breaks
    .trim();

  return markdown;
}

/**
 * Process an HTML node and its children to markdown
 * @param {Node} node - DOM node
 * @param {number} depth - Current nesting depth for lists
 * @returns {string} - Markdown output
 */
function processNode(node, depth = 0) {
  if (!node) return '';
  let result = '';

  // Handle text nodes
  if (node.nodeType === 3) { // Text node
    return node.textContent;
  }

  // Handle element nodes
  if (node.nodeType === 1) { // Element node
    const nodeName = node.nodeName.toLowerCase();
    const children = Array.from(node.childNodes);
    
    // Check for directionality
    const dir = node.getAttribute('dir');
    const isRTL = dir === 'rtl';
    
    // Process different element types
    switch (nodeName) {
      case 'h1':
        result += `# ${processChildren(node)}\n\n`;
        break;
      case 'h2':
        result += `## ${processChildren(node)}\n\n`;
        break;
      case 'h3':
        result += `### ${processChildren(node)}\n\n`;
        break;
      case 'h4':
        result += `#### ${processChildren(node)}\n\n`;
        break;
      case 'h5':
        result += `##### ${processChildren(node)}\n\n`;
        break;
      case 'h6':
        result += `###### ${processChildren(node)}\n\n`;
        break;
      case 'p':
        // Handle right-to-left text direction
        if (isRTL) {
          result += `<div dir="rtl">\n\n${processChildren(node)}\n\n</div>\n\n`;
        } else {
          result += `${processChildren(node)}\n\n`;
        }
        break;
      case 'strong':
      case 'b':
        result += `**${processChildren(node)}**`;
        break;
      case 'em':
      case 'i':
        result += `*${processChildren(node)}*`;
        break;
      case 'u':
        result += `<u>${processChildren(node)}</u>`;
        break;
      case 'strike':
      case 's':
      case 'del':
        result += `~~${processChildren(node)}~~`;
        break;
      case 'code':
        result += `\`${processChildren(node)}\``;
        break;
      case 'pre':
        result += `\`\`\`\n${processChildren(node)}\n\`\`\`\n\n`;
        break;
      case 'a':
        const href = node.getAttribute('href') || '';
        result += `[${processChildren(node)}](${href})`;
        break;
      case 'img':
        const src = node.getAttribute('src') || '';
        const alt = node.getAttribute('alt') || '';
        result += `![${alt}](${src})`;
        break;
      case 'blockquote':
        // Split into lines and add > to each line
        const blockquoteContent = processChildren(node).split('\n')
          .map(line => `> ${line}`)
          .join('\n');
        result += `${blockquoteContent}\n\n`;
        break;
      case 'ul':
        result += processListItems(node, '*', depth);
        break;
      case 'ol':
        result += processListItems(node, '1.', depth);
        break;
      case 'li':
        // This is handled by processListItems
        result += processChildren(node);
        break;
      case 'table':
        result += processTable(node);
        break;
      case 'hr':
        result += `\n---\n\n`;
        break;
      case 'br':
        result += `\n`;
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
 * Process child nodes
 * @param {Node} node - Parent node
 * @returns {string} - Processed content
 */
function processChildren(node) {
  let result = '';
  const children = Array.from(node.childNodes);
  
  children.forEach(child => {
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
  let result = '\n';
  const items = Array.from(listNode.querySelectorAll(':scope > li'));
  
  items.forEach(item => {
    const indent = '  '.repeat(depth);
    const itemContent = processNode(item, depth + 1);
    result += `${indent}${marker} ${itemContent}\n`;
  });
  
  return result + '\n';
}

/**
 * Process table to markdown
 * @param {Node} tableNode - Table node
 * @returns {string} - Markdown table
 */
function processTable(tableNode) {
  let result = '\n';
  const rows = Array.from(tableNode.querySelectorAll('tr'));
  
  if (rows.length === 0) return '';
  
  // Process header row
  const headerCells = Array.from(rows[0].querySelectorAll('th, td'));
  if (headerCells.length === 0) return '';
  
  // Create header row
  const headerTexts = headerCells.map(cell => processChildren(cell).trim());
  result += `| ${headerTexts.join(' | ')} |\n`;
  
  // Create separator row
  result += `| ${headerTexts.map(() => '---').join(' | ')} |\n`;
  
  // Process data rows
  for (let i = 1; i < rows.length; i++) {
    const cells = Array.from(rows[i].querySelectorAll('td'));
    if (cells.length === 0) continue;
    
    const cellTexts = cells.map(cell => processChildren(cell).trim());
    result += `| ${cellTexts.join(' | ')} |\n`;
  }
  
  return result + '\n';
}

module.exports = { markdownify };
