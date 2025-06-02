#!/usr/bin/env node

// Script to create a test DOCX file with italic formatting
// This will help us test the italic formatting fix

const fs = require('fs');
const path = require('path');

/**
 * Create a simple test DOCX file with italic text
 * Note: This creates a minimal DOCX structure for testing
 */
function createTestDocx() {
  console.log('Creating test DOCX file with italic formatting...');
  
  // Create a simple XML structure for a DOCX with italic text
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>This is normal text. </w:t>
      </w:r>
      <w:r>
        <w:rPr>
          <w:i/>
        </w:rPr>
        <w:t>This text should be italic.</w:t>
      </w:r>
      <w:r>
        <w:t> Back to normal text.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:rPr>
          <w:b/>
          <w:i/>
        </w:rPr>
        <w:t>This text should be bold and italic.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Another paragraph with </w:t>
      </w:r>
      <w:r>
        <w:rPr>
          <w:i/>
        </w:rPr>
        <w:t>italic words</w:t>
      </w:r>
      <w:r>
        <w:t> in the middle.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

  // Create basic content types
  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

  // Create basic relationships
  const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

  const wordRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>`;

  // Note: Creating a proper DOCX requires ZIP manipulation
  // For now, we'll create the XML files and suggest using an existing DOCX
  
  const testDir = 'test-docx-files';
  if (!fs.existsSync(testDir)) {
    fs.mkdirSync(testDir);
  }
  
  // Save the XML files for reference
  fs.writeFileSync(path.join(testDir, 'document.xml'), documentXml);
  fs.writeFileSync(path.join(testDir, '[Content_Types].xml'), contentTypes);
  fs.writeFileSync(path.join(testDir, '_rels.xml'), rels);
  fs.writeFileSync(path.join(testDir, 'word_rels.xml'), wordRels);
  
  console.log(`Test XML files created in ${testDir}/`);
  console.log('Note: To create a proper DOCX file, you would need to:');
  console.log('1. Create a ZIP archive');
  console.log('2. Add these XML files in the correct structure');
  console.log('3. Rename the ZIP to .docx');
  console.log('');
  console.log('For testing, please use an existing DOCX file with italic text.');
  console.log('You can create one in Microsoft Word with:');
  console.log('- Normal text');
  console.log('- Italic text (Ctrl+I)');
  console.log('- Bold and italic text');
  console.log('- Mixed formatting in paragraphs');
}

// Run if called directly
if (require.main === module) {
  createTestDocx();
}

module.exports = { createTestDocx };
