// File: lib/css/generators/specialized-styles.js
// Specialized styles for accessibility, track changes, and headers

/**
 * Generate accessibility styles
 */
function generateAccessibilityStyles(styleInfo) {
  return `
/* Accessibility Styles */
.skip-link { position: absolute; top: -100px; left: 0; padding: 10px; background-color: #f0f0f0; color: #333; z-index: 9999; transition: top 0.3s ease-in-out; text-decoration: none; border: 1px solid #ccc; border-top: none; box-shadow: 0 2px 5px rgba(0,0,0,0.2); }
.skip-link:focus { top: 0; outline: 2px solid #005a9c; outline-offset: -2px; }
.keyboard-focusable:focus, *:focus-visible { outline: 2px dashed #005a9c !important; outline-offset: 2px !important; box-shadow: 0 0 0 2px rgba(0,90,156,0.3) !important; }
@media (prefers-reduced-motion: reduce) { * { transition-duration: 0.001ms !important; animation-duration: 0.001ms !important; animation-iteration-count: 1 !important; scroll-behavior: auto !important; } }
@media (prefers-contrast: more) {
  body { color: #000000 !important; background-color: #ffffff !important; }
  a { color: #0000EE !important; text-decoration: underline !important; } a:visited { color: #551A8B !important; }
  th, td { border: 1px solid #000000 !important; }
  .docx-toc-dots::after { background-image: linear-gradient(to right, #000 1px, transparent 0) !important; }
  .docx-insertion { background-color: #cce5ff !important; border: 1px solid #007bff !important; color: #004085 !important; }
  .docx-deletion { background-color: #f8d7da !important; border: 1px solid #dc3545 !important; color: #721c24 !important; text-decoration: line-through !important; }
}
.sr-only { position: absolute; width: 1px; height: 1px; padding: 0; margin: -1px; overflow: hidden; clip: rect(0, 0, 0, 0); white-space: nowrap; border-width: 0; }
table caption.sr-only { position: static; width: auto; height: auto; margin: 0; overflow: visible; clip: auto; white-space: normal; }
figure { margin: 1em 0; }
figcaption { font-style: italic; margin-top: 0.5em; }
`;
}

/**
 * Generate track changes styles
 */
function generateTrackChangesStyles(styleInfo) {
  return `
/* Track Changes Styles */
.docx-track-changes-legend { margin: 1em 0; padding: 0.5em; border: 1px solid #BDBDBD; border-radius: 4px; background-color: #F5F5F5; font-size: 0.9em; }
.docx-track-changes-toggle { padding: 0.3em 0.6em; margin-right: 1em; border: 1px solid #BDBDBD; border-radius: 3px; cursor: pointer; background-color: #e9e9e9; }
.docx-track-changes-toggle:hover { background-color: #dcdcdc; }
.docx-track-changes-show .docx-insertion { background-color: #e6ffed; border-bottom: 1px solid #a2d5ab; position: relative; }
.docx-track-changes-show .docx-insertion:hover::after, .docx-track-changes-show .docx-deletion:hover::after { content: attr(data-author) " (" attr(data-date) ")"; display: block; position: absolute; bottom: 100%; left: 0; margin-bottom: 2px; background-color: #333; color: #fff; border-radius: 3px; padding: 0.25em 0.5em; font-size: 0.8em; z-index: 100; white-space: nowrap; box-shadow: 0 1px 3px rgba(0,0,0,0.2); }
.docx-track-changes-show .docx-deletion { background-color: #ffe6e6; border-bottom: 1px solid #f5b7b1; text-decoration: line-through; color: #d9534f; position: relative; }
.docx-track-changes-hide .docx-insertion, .docx-track-changes-hide .docx-deletion { background-color: transparent; text-decoration: none; border-bottom: none; color: inherit; }
.docx-track-changes-hide .docx-deleted-content { display: none; }
.docx-deleted-content { margin: 1em 0; padding: 0.5em; border: 1px dashed #FFCDD2; border-radius: 4px; background-color: #FFEBEE; color: #b71c1c; }
.docx-deleted-content::before { content: "Deleted Content Section"; display: block; font-weight: bold; margin-bottom: 0.5em; }
`;
}

/**
 * Generate header styles
 */
function generateHeaderStyles(styleInfo) {
  return `
/* Document Header Styles */
.docx-document-header {
  margin-bottom: 2em;
  padding-bottom: 1em;
  border-bottom: 1px solid #e0e0e0;
  text-align: center;
  page-break-after: avoid;
}

.docx-header-paragraph {
  margin: 0.5em 0;
  line-height: 1.3;
  word-wrap: break-word;
  overflow-wrap: break-word;
}

.docx-header-text {
  display: block;
  word-wrap: break-word;
  overflow-wrap: break-word;
}

/* Default header styling for common header styles */
.docx-header-paragraph[data-style-id*="header"],
.docx-header-paragraph[data-style-id*="Header"] {
  font-size: 16pt;
  font-weight: bold;
  margin: 1em 0;
}

.docx-header-paragraph[data-style-id*="title"],
.docx-header-paragraph[data-style-id*="Title"] {
  font-size: 18pt;
  font-weight: bold;
  margin: 1em 0;
}

.docx-header-paragraph[data-style-id*="subtitle"],
.docx-header-paragraph[data-style-id*="Subtitle"] {
  font-size: 14pt;
  font-style: italic;
  margin: 0.5em 0;
}

/* Header type specific styles */
.docx-header-paragraph[data-header-type="first"] {
  /* Styles for first page header */
  font-weight: bold;
}

.docx-header-paragraph[data-header-type="even"] {
  /* Styles for even page header */
  text-align: left;
}

.docx-header-paragraph[data-header-type="default"] {
  /* Styles for default header */
  text-align: center;
}

/* Ensure header images are properly sized */
.docx-document-header img {
  max-width: 100%;
  height: auto;
  display: block;
  margin: 0.5em auto;
}

/* Print-specific header styles */
@media print {
  .docx-document-header {
    page-break-after: avoid;
    break-after: avoid;
  }
  
  .docx-header-paragraph {
    page-break-inside: avoid;
    break-inside: avoid;
  }
}

/* Responsive header styles */
@media (max-width: 768px) {
  .docx-document-header {
    margin-bottom: 1.5em;
    padding-bottom: 0.5em;
  }
  
  .docx-header-paragraph[data-style-id*="title"],
  .docx-header-paragraph[data-style-id*="Title"] {
    font-size: 16pt;
  }
  
  .docx-header-paragraph[data-style-id*="subtitle"],
  .docx-header-paragraph[data-style-id*="Subtitle"] {
    font-size: 12pt;
  }
  
  .docx-header-paragraph[data-style-id*="header"],
  .docx-header-paragraph[data-style-id*="Header"] {
    font-size: 14pt;
  }
}
`;
}

module.exports = {
  generateAccessibilityStyles,
  generateTrackChangesStyles,
  generateHeaderStyles,
};