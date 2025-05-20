// lib/html/structure-processor.js - HTML structure handling functions

/**
 * Ensure HTML has proper structure (html, head, body)
 * This function ensures that the HTML document has the proper structure
 * by creating missing elements if necessary
 * 
 * @param {Document} document - DOM document
 */
function ensureHtmlStructure(document) {
  // Create html element if not exists
  if (
    !document.documentElement ||
    document.documentElement.nodeName.toLowerCase() !== "html"
  ) {
    const html = document.createElement("html");

    // Move existing content
    while (document.childNodes.length > 0) {
      html.appendChild(document.childNodes[0]);
    }

    document.appendChild(html);
  }

  // Create head if not exists
  if (!document.head) {
    const head = document.createElement("head");
    document.documentElement.insertBefore(
      head,
      document.documentElement.firstChild
    );
  }

  // Add meta charset
  if (!document.querySelector("meta[charset]")) {
    const meta = document.createElement("meta");
    meta.setAttribute("charset", "utf-8");
    document.head.appendChild(meta);
  }

  // Create body if not exists
  if (!document.body) {
    const body = document.createElement("body");

    // Move content to body
    Array.from(document.documentElement.childNodes).forEach((node) => {
      if (node !== document.head && node.nodeName.toLowerCase() !== "body") {
        body.appendChild(node);
      }
    });

    document.documentElement.appendChild(body);
  }
}

/**
 * Add document metadata based on style information
 * Adds title, viewport meta tag, and body class
 * 
 * @param {Document} document - DOM document
 * @param {Object} styleInfo - Style information
 */
function addDocumentMetadata(document, styleInfo) {
  // Add title
  const title = document.createElement("title");
  title.textContent = "Document";
  document.head.appendChild(title);

  // Add viewport meta
  const viewport = document.createElement("meta");
  viewport.setAttribute("name", "viewport");
  viewport.setAttribute("content", "width=device-width, initial-scale=1");
  document.head.appendChild(viewport);

  // Add class to body 
  document.body.classList.add('doc2web-document');
}

module.exports = {
  ensureHtmlStructure,
  addDocumentMetadata
};