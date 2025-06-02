#!/usr/bin/env node

// Test script for italic formatting fix
// This script will help us test the italic formatting conversion

const fs = require('fs');
const path = require('path');
const { extractAndApplyStyles } = require('./lib/html/html-generator');

/**
 * Test italic formatting conversion
 */
async function testItalicFormatting() {
  console.log('=== Italic Formatting Test ===\n');
  
  // Note: For now, we'll test with any existing DOCX file
  // In a real scenario, we would create a test DOCX with various italic scenarios
  
  const testFiles = [
    // Add any test DOCX files here
  ];
  
  // Look for any DOCX files in current directory for testing
  const currentDir = process.cwd();
  const docxFiles = fs.readdirSync(currentDir)
    .filter(file => file.toLowerCase().endsWith('.docx'))
    .slice(0, 1); // Test with first DOCX file found
  
  if (docxFiles.length === 0) {
    console.log('No DOCX files found for testing.');
    console.log('Please place a DOCX file with italic text in the current directory.');
    return;
  }
  
  for (const docxFile of docxFiles) {
    console.log(`Testing with: ${docxFile}`);
    
    try {
      const result = await extractAndApplyStyles(docxFile);
      
      // Check for italic elements in HTML
      const italicMatches = result.html.match(/<em[^>]*>|<i[^>]*>|font-style:\s*italic/gi) || [];
      console.log(`Found ${italicMatches.length} italic elements/styles in HTML output`);
      
      if (italicMatches.length > 0) {
        console.log('Italic elements found:');
        italicMatches.forEach((match, index) => {
          console.log(`  ${index + 1}: ${match}`);
        });
      }
      
      // Check CSS for italic styles
      const cssItalicMatches = result.styles.match(/font-style:\s*italic/gi) || [];
      console.log(`Found ${cssItalicMatches.length} italic styles in CSS`);
      
      // Save test output
      const outputDir = 'test-output';
      if (!fs.existsSync(outputDir)) {
        fs.mkdirSync(outputDir);
      }
      
      const baseName = path.basename(docxFile, '.docx');
      fs.writeFileSync(path.join(outputDir, `${baseName}-test.html`), result.html);
      fs.writeFileSync(path.join(outputDir, `${baseName}-test.css`), result.styles);
      
      console.log(`Test output saved to ${outputDir}/`);
      
    } catch (error) {
      console.error(`Error testing ${docxFile}:`, error.message);
    }
    
    console.log('---\n');
  }
}

// Run test if called directly
if (require.main === module) {
  testItalicFormatting().catch(console.error);
}

module.exports = { testItalicFormatting };
