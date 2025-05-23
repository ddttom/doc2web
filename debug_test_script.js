// Debug test script to verify numbering fixes
const fs = require('fs');
const path = require('path');
const { JSDOM } = require('jsdom');

// Load the HTML file
const htmlPath = path.join(__dirname, 'output/input/TEST-TWO.html');
const html = fs.readFileSync(htmlPath, 'utf8');

// Create a DOM from the HTML
const dom = new JSDOM(html);
const document = dom.window.document;

// Function to simulate the numbering fix for headings
function fixHeadingNumbering(doc) {
  const headings = doc.querySelectorAll("h1, h2, h3, h4, h5, h6");
  
  headings.forEach(heading => {
    // Skip TOC headings
    if (heading.classList.contains("docx-toc-heading")) return;
    
    // Get the numbering data
    const numId = heading.getAttribute("data-num-id");
    const abstractNum = heading.getAttribute("data-abstract-num");
    const numLevel = heading.getAttribute("data-num-level");
    const format = heading.getAttribute("data-format");
    
    if (numId && abstractNum && numLevel !== null) {
      // Find the heading-number span
      const numberSpan = heading.querySelector(".heading-number");
      
      if (numberSpan) {
        // Generate a number based on the data attributes
        let number = '';
        
        // For decimal format (1, 2, 3...)
        if (format === "decimal") {
          number = `${parseInt(numLevel) + 1}.`;
        } 
        // For lowercase letter format (a, b, c...)
        else if (format === "lowerLetter") {
          number = `${String.fromCharCode(97 + parseInt(numLevel))}.`;
        }
        // For uppercase letter format (A, B, C...)
        else if (format === "upperLetter") {
          number = `${String.fromCharCode(65 + parseInt(numLevel))}.`;
        }
        // For lowercase Roman format (i, ii, iii...)
        else if (format === "lowerRoman") {
          const romans = ["i", "ii", "iii", "iv", "v", "vi", "vii", "viii", "ix", "x"];
          number = `${romans[parseInt(numLevel)]}.`;
        }
        // For uppercase Roman format (I, II, III...)
        else if (format === "upperRoman") {
          const romans = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X"];
          number = `${romans[parseInt(numLevel)]}.`;
        }
        
        // Set the number text
        numberSpan.textContent = number;
      }
    }
  });
}

// Function to simulate the numbering fix for TOC entries
function fixTOCNumbering(doc) {
  const tocEntries = doc.querySelectorAll(".docx-toc-entry");
  
  tocEntries.forEach(entry => {
    // Get the numbering data
    const numId = entry.getAttribute("data-num-id");
    const abstractNum = entry.getAttribute("data-abstract-num");
    const numLevel = entry.getAttribute("data-num-level");
    const format = entry.getAttribute("data-format");
    
    if (numId && abstractNum && numLevel !== null) {
      // Find the text span
      const textSpan = entry.querySelector(".docx-toc-text");
      
      if (textSpan) {
        // Get the anchor element
        const anchor = textSpan.querySelector("a");
        
        if (anchor) {
          // Generate a number based on the data attributes
          let number = '';
          
          // For decimal format (1, 2, 3...)
          if (format === "decimal") {
            number = `${parseInt(numLevel) + 1}.`;
          } 
          // For lowercase letter format (a, b, c...)
          else if (format === "lowerLetter") {
            number = `${String.fromCharCode(97 + parseInt(numLevel))}.`;
          }
          // For uppercase letter format (A, B, C...)
          else if (format === "upperLetter") {
            number = `${String.fromCharCode(65 + parseInt(numLevel))}.`;
          }
          // For lowercase Roman format (i, ii, iii...)
          else if (format === "lowerRoman") {
            const romans = ["i", "ii", "iii", "iv", "v", "vi", "vii", "viii", "ix", "x"];
            number = `${romans[parseInt(numLevel)]}.`;
          }
          // For uppercase Roman format (I, II, III...)
          else if (format === "upperRoman") {
            const romans = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X"];
            number = `${romans[parseInt(numLevel)]}.`;
          }
          
          // Prepend the number to the anchor text
          anchor.textContent = `${number} ${anchor.textContent}`;
        }
      }
    }
  });
}

// Apply the fixes
console.log("Applying numbering fixes...");
fixHeadingNumbering(document);
fixTOCNumbering(document);

// Save the modified HTML
const outputPath = path.join(__dirname, 'output/input/TEST-TWO-fixed.html');
fs.writeFileSync(outputPath, dom.serialize());

console.log(`Fixed HTML saved to ${outputPath}`);