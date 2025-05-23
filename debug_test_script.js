// debug-test.js - Debug script to test doc2web components with enhanced section ID debugging
const fs = require('fs');
const path = require('path');
const { extractAndApplyStyles } = require('./lib');

/**
 * Test the DOCX processing pipeline with detailed logging
 */
async function testDocxProcessing(docxPath) {
  console.log('=== DEBUG TEST: DOCX Processing Pipeline ===');
  console.log(`Testing file: ${docxPath}`);
  
  try {
    // Check if file exists
    if (!fs.existsSync(docxPath)) {
      throw new Error(`File not found: ${docxPath}`);
    }
    
    // Check file stats
    const stats = fs.statSync(docxPath);
    console.log(`File size: ${stats.size} bytes`);
    console.log(`File modified: ${stats.mtime}`);
    
    // Test the main processing function
    console.log('\n--- Testing extractAndApplyStyles ---');
    const result = await extractAndApplyStyles(docxPath, 'test.css', {
      enhanceAccessibility: true,
      preserveMetadata: true,
      trackChangesMode: 'show'
    });
    
    // Analyze results
    console.log('\n--- Results Analysis ---');
    console.log(`HTML length: ${result.html?.length || 0} characters`);
    console.log(`CSS length: ${result.styles?.length || 0} characters`);
    console.log(`Messages count: ${result.messages?.length || 0}`);
    console.log(`Metadata present: ${!!result.metadata}`);
    console.log(`Track changes present: ${result.trackChanges?.hasTrackedChanges || false}`);
    console.log(`Numbering context entries: ${result.numberingContext?.length || 0}`);
    
    // Check for common issues
    console.log('\n--- Issue Detection ---');
    
    if (!result.html || result.html.length < 100) {
      console.error('❌ HTML output is too short or empty');
      console.log('HTML preview:', result.html?.substring(0, 200));
    } else {
      console.log('✅ HTML output looks valid');
    }
    
    if (!result.styles || result.styles.length < 100) {
      console.warn('⚠️  CSS output is very short');
    } else {
      console.log('✅ CSS output looks valid');
    }
    
    // Check HTML structure
    if (result.html) {
      const hasDoctype = result.html.includes('<!DOCTYPE');
      const hasHtml = result.html.includes('<html');
      const hasBody = result.html.includes('<body');
      const hasContent = result.html.includes('<p') || result.html.includes('<h1') || result.html.includes('<div');
      
      console.log(`HTML structure check:`);
      console.log(`  DOCTYPE: ${hasDoctype ? '✅' : '❌'}`);
      console.log(`  <html>: ${hasHtml ? '✅' : '❌'}`);
      console.log(`  <body>: ${hasBody ? '✅' : '❌'}`);
      console.log(`  Content: ${hasContent ? '✅' : '❌'}`);
      
      // Enhanced section ID validation
      console.log('\n--- Section ID Analysis ---');
      const sectionIds = result.html.match(/id="section-[^"]+"/g) || [];
      console.log(`Section IDs found: ${sectionIds.length}`);
      
      if (sectionIds.length > 0) {
        console.log('Sample section IDs:');
        sectionIds.slice(0, 5).forEach((id, index) => {
          console.log(`  ${index + 1}. ${id}`);
        });
        
        // Check for malformed section IDs (ending with dash or other issues)
        const malformedIds = result.html.match(/id="section-[^"]*-"/g) || [];
        if (malformedIds.length > 0) {
          console.warn(`⚠️ Found ${malformedIds.length} malformed section IDs ending with dash:`);
          malformedIds.slice(0, 3).forEach((id, index) => {
            console.warn(`  ${index + 1}. ${id}`);
          });
        } else {
          console.log('✅ No malformed section IDs ending with dash');
        }
        
        // Check for proper hierarchical structure
        const hierarchicalIds = result.html.match(/id="section-\d+(-[a-z0-9]+)*"/g) || [];
        console.log(`Properly formed hierarchical IDs: ${hierarchicalIds.length}`);
        
        // Check for empty section IDs
        const emptyIds = result.html.match(/id="section-\s*"/g) || [];
        if (emptyIds.length > 0) {
          console.warn(`⚠️ Found ${emptyIds.length} empty section IDs`);
        }
        
        // Analyze section ID patterns
        const patterns = {
          levelOne: (result.html.match(/id="section-\d+"/g) || []).length,
          levelTwo: (result.html.match(/id="section-\d+-[a-z0-9]+"/g) || []).length,
          levelThree: (result.html.match(/id="section-\d+-[a-z0-9]+-[a-z0-9]+"/g) || []).length,
        };
        
        console.log('Section ID hierarchy analysis:');
        console.log(`  Level 1 (section-X): ${patterns.levelOne}`);
        console.log(`  Level 2 (section-X-Y): ${patterns.levelTwo}`);
        console.log(`  Level 3 (section-X-Y-Z): ${patterns.levelThree}`);
        
      } else {
        console.warn('⚠️ No section IDs found in HTML output');
      }
      
      // Check for data attributes related to numbering
      const dataNumElements = (result.html.match(/data-num-id="/g) || []).length;
      const dataAbstractNumElements = (result.html.match(/data-abstract-num="/g) || []).length;
      const dataNumLevelElements = (result.html.match(/data-num-level="/g) || []).length;
      
      console.log('\n--- Numbering Data Attributes ---');
      console.log(`Elements with data-num-id: ${dataNumElements}`);
      console.log(`Elements with data-abstract-num: ${dataAbstractNumElements}`);
      console.log(`Elements with data-num-level: ${dataNumLevelElements}`);
    }
    
    // Show conversion messages if any
    if (result.messages && result.messages.length > 0) {
      console.log('\n--- Conversion Messages ---');
      result.messages.forEach((msg, index) => {
        console.log(`${index + 1}. [${msg.type}] ${msg.message}`);
      });
    }
    
    // Enhanced numbering context analysis with section IDs
    if (result.numberingContext && result.numberingContext.length > 0) {
      console.log('\n--- Numbering Context with Section IDs ---');
      const numberedItems = result.numberingContext.filter(ctx => ctx.resolvedNumbering);
      console.log(`Total contexts: ${result.numberingContext.length}`);
      console.log(`With resolved numbering: ${numberedItems.length}`);
      
      if (numberedItems.length > 0) {
        console.log('Sample numbered items with section analysis:');
        numberedItems.slice(0, 5).forEach((ctx, index) => {
          console.log(`  ${index + 1}. Level ${ctx.numberingLevel}: "${ctx.textContent?.substring(0, 50)}..."`);
          
          if (ctx.resolvedNumbering) {
            console.log(`     Raw Number: ${ctx.resolvedNumbering.rawNumber}`);
            console.log(`     Format: ${ctx.resolvedNumbering.format}`);
            console.log(`     Full Display: ${ctx.resolvedNumbering.fullNumbering || 'N/A'}`);
            
            // Show hierarchical numbering if available
            if (ctx.resolvedNumbering.hierarchicalNumbering) {
              console.log(`     Hierarchical: ${ctx.resolvedNumbering.hierarchicalNumbering}`);
            }
            
            // Show section ID information
            if (ctx.resolvedNumbering.sectionId) {
              console.log(`     Section ID: ${ctx.resolvedNumbering.sectionId}`);
              
              // Validate section ID format
              if (ctx.resolvedNumbering.sectionId.endsWith('-')) {
                console.warn(`     ⚠️ Section ID ends with dash (malformed)`);
              } else if (ctx.resolvedNumbering.sectionId === 'section-') {
                console.warn(`     ⚠️ Section ID is empty (malformed)`);
              } else {
                console.log(`     ✅ Section ID looks valid`);
              }
            } else {
              console.warn(`     ⚠️ No section ID generated`);
            }
          }
          console.log(''); // Empty line for readability
        });
        
        // Analyze numbering patterns
        const levelCounts = {};
        const formatCounts = {};
        numberedItems.forEach(ctx => {
          const level = ctx.numberingLevel;
          const format = ctx.resolvedNumbering?.format;
          
          levelCounts[level] = (levelCounts[level] || 0) + 1;
          if (format) {
            formatCounts[format] = (formatCounts[format] || 0) + 1;
          }
        });
        
        console.log('Numbering level distribution:');
        Object.keys(levelCounts).sort((a, b) => parseInt(a) - parseInt(b)).forEach(level => {
          console.log(`  Level ${level}: ${levelCounts[level]} items`);
        });
        
        console.log('Numbering format distribution:');
        Object.entries(formatCounts).forEach(([format, count]) => {
          console.log(`  ${format}: ${count} items`);
        });
      }
    } else {
      console.warn('⚠️ No numbering context found');
    }
    
    return result;
    
  } catch (error) {
    console.error('\n❌ ERROR in processing pipeline:');
    console.error(`Message: ${error.message}`);
    console.error(`Stack: ${error.stack}`);
    return null;
  }
}

/**
 * Test individual components
 */
async function testComponents(docxPath) {
  console.log('\n=== DEBUG TEST: Individual Components ===');
  
  try {
    // Test style parsing
    console.log('\n--- Testing Style Parser ---');
    const { parseDocxStyles } = require('./lib/parsers/style-parser');
    const styleInfo = await parseDocxStyles(docxPath);
    
    console.log(`Paragraph styles: ${Object.keys(styleInfo.styles.paragraph || {}).length}`);
    console.log(`Character styles: ${Object.keys(styleInfo.styles.character || {}).length}`);
    console.log(`Table styles: ${Object.keys(styleInfo.styles.table || {}).length}`);
    console.log(`Theme fonts: ${styleInfo.theme?.fonts?.major || 'N/A'}, ${styleInfo.theme?.fonts?.minor || 'N/A'}`);
    console.log(`TOC detected: ${styleInfo.tocStyles?.hasTableOfContents || false}`);
    console.log(`Numbering definitions: ${Object.keys(styleInfo.numberingDefs?.abstractNums || {}).length}`);
    
    // Enhanced numbering analysis
    if (styleInfo.numberingDefs && Object.keys(styleInfo.numberingDefs.abstractNums || {}).length > 0) {
      console.log('\n--- Numbering Definitions Analysis ---');
      Object.entries(styleInfo.numberingDefs.abstractNums).forEach(([id, abstractNum]) => {
        console.log(`Abstract Numbering ${id}:`);
        console.log(`  Levels: ${Object.keys(abstractNum.levels || {}).length}`);
        console.log(`  Multi-level type: ${abstractNum.multiLevelType || 'N/A'}`);
        
        if (abstractNum.levels) {
          Object.entries(abstractNum.levels).forEach(([levelIndex, levelDef]) => {
            console.log(`    Level ${levelIndex}: format=${levelDef.format}, text="${levelDef.textFormat}"`);
          });
        }
      });
    }
    
    // Enhanced numbering context analysis
    if (styleInfo.numberingContext && styleInfo.numberingContext.length > 0) {
      console.log('\n--- Numbering Context Analysis ---');
      console.log(`Total numbering contexts: ${styleInfo.numberingContext.length}`);
      
      const withResolvedNumbering = styleInfo.numberingContext.filter(ctx => ctx.resolvedNumbering);
      console.log(`With resolved numbering: ${withResolvedNumbering.length}`);
      
      if (withResolvedNumbering.length > 0) {
        console.log('Sample resolved numbering:');
        withResolvedNumbering.slice(0, 3).forEach((ctx, index) => {
          console.log(`  ${index + 1}. "${ctx.textContent?.substring(0, 30)}..."`);
          console.log(`     Abstract Num: ${ctx.abstractNumId}`);
          console.log(`     Level: ${ctx.numberingLevel}`);
          if (ctx.resolvedNumbering.sectionId) {
            console.log(`     Section ID: ${ctx.resolvedNumbering.sectionId}`);
          }
        });
      }
    }
    
    // Test CSS generation
    console.log('\n--- Testing CSS Generator ---');
    const { generateCssFromStyleInfo } = require('./lib/css/css-generator');
    const css = generateCssFromStyleInfo(styleInfo);
    
    console.log(`Generated CSS length: ${css.length} characters`);
    
    // Check for section ID related CSS
    const sectionCssRules = (css.match(/\[id\^="section-"\]/g) || []).length;
    const counterRules = (css.match(/counter-increment:/g) || []).length;
    const beforeRules = (css.match(/::before/g) || []).length;
    
    console.log(`Section ID CSS rules: ${sectionCssRules}`);
    console.log(`Counter increment rules: ${counterRules}`);
    console.log(`::before pseudo-element rules: ${beforeRules}`);
    
    // Test mammoth conversion
    console.log('\n--- Testing Mammoth Conversion ---');
    const mammoth = require('mammoth');
    const basicResult = await mammoth.convertToHtml({ path: docxPath });
    
    console.log(`Basic mammoth result length: ${basicResult.value.length} characters`);
    console.log(`Basic mammoth messages: ${basicResult.messages.length}`);
    
    if (basicResult.messages.length > 0) {
      console.log('Mammoth messages:');
      basicResult.messages.slice(0, 5).forEach(msg => {
        console.log(`  [${msg.type}] ${msg.message}`);
      });
    }
    
    return {
      styleInfo,
      css,
      mammothResult: basicResult
    };
    
  } catch (error) {
    console.error('\n❌ ERROR in component testing:');
    console.error(`Message: ${error.message}`);
    console.error(`Stack: ${error.stack}`);
    return null;
  }
}

/**
 * Main debug function
 */
async function debugMain() {
  const args = process.argv.slice(2);
  
  if (args.length === 0) {
    console.log('Usage: node debug-test.js <path-to-docx-file>');
    console.log('Example: node debug-test.js ./test-documents/sample.docx');
    process.exit(1);
  }
  
  const docxPath = args[0];
  
  console.log('doc2web Debug Test');
  console.log('=================');
  console.log(`Node.js version: ${process.version}`);
  console.log(`Current working directory: ${process.cwd()}`);
  console.log(`Testing file: ${docxPath}`);
  console.log(`File exists: ${fs.existsSync(docxPath)}`);
  
  try {
    // Test components first
    const componentResults = await testComponents(docxPath);
    
    // Test full pipeline
    const pipelineResults = await testDocxProcessing(docxPath);
    
    // Summary
    console.log('\n=== SUMMARY ===');
    if (componentResults && pipelineResults) {
      console.log('✅ All tests completed successfully');
      
      // Save test output for inspection
      const outputDir = './debug-output';
      if (!fs.existsSync(outputDir)) {
        fs.mkdirSync(outputDir);
      }
      
      if (pipelineResults.html) {
        fs.writeFileSync(path.join(outputDir, 'debug-output.html'), pipelineResults.html);
        console.log(`✅ Test HTML saved to: ${path.join(outputDir, 'debug-output.html')}`);
      }
      
      if (pipelineResults.styles) {
        fs.writeFileSync(path.join(outputDir, 'debug-output.css'), pipelineResults.styles);
        console.log(`✅ Test CSS saved to: ${path.join(outputDir, 'debug-output.css')}`);
      }
      
      // Save debug info with enhanced section ID metrics
      const debugInfo = {
        timestamp: new Date().toISOString(),
        file: docxPath,
        results: {
          htmlLength: pipelineResults.html?.length || 0,
          cssLength: pipelineResults.styles?.length || 0,
          messagesCount: pipelineResults.messages?.length || 0,
          hasMetadata: !!pipelineResults.metadata,
          hasTrackChanges: pipelineResults.trackChanges?.hasTrackedChanges || false,
          numberingContextLength: pipelineResults.numberingContext?.length || 0,
          sectionIdsCount: (pipelineResults.html?.match(/id="section-[^"]+"/g) || []).length,
          malformedSectionIds: (pipelineResults.html?.match(/id="section-[^"]*-"/g) || []).length,
          hierarchicalSectionIds: (pipelineResults.html?.match(/id="section-\d+(-[a-z0-9]+)*"/g) || []).length
        },
        components: {
          paragraphStyles: Object.keys(componentResults.styleInfo?.styles?.paragraph || {}).length,
          characterStyles: Object.keys(componentResults.styleInfo?.styles?.character || {}).length,
          tableStyles: Object.keys(componentResults.styleInfo?.styles?.table || {}).length,
          cssLength: componentResults.css?.length || 0,
          mammothLength: componentResults.mammothResult?.value?.length || 0,
          numberingDefinitions: Object.keys(componentResults.styleInfo?.numberingDefs?.abstractNums || {}).length,
          resolvedNumberingContexts: (componentResults.styleInfo?.numberingContext?.filter(ctx => ctx.resolvedNumbering) || []).length
        },
        sectionIdAnalysis: {
          total: (pipelineResults.html?.match(/id="section-[^"]+"/g) || []).length,
          malformed: (pipelineResults.html?.match(/id="section-[^"]*-"/g) || []).length,
          hierarchical: (pipelineResults.html?.match(/id="section-\d+(-[a-z0-9]+)*"/g) || []).length,
          levelOne: (pipelineResults.html?.match(/id="section-\d+"/g) || []).length,
          levelTwo: (pipelineResults.html?.match(/id="section-\d+-[a-z0-9]+"/g) || []).length,
          levelThree: (pipelineResults.html?.match(/id="section-\d+-[a-z0-9]+-[a-z0-9]+"/g) || []).length
        }
      };
      
      fs.writeFileSync(path.join(outputDir, 'debug-info.json'), JSON.stringify(debugInfo, null, 2));
      console.log(`✅ Debug info saved to: ${path.join(outputDir, 'debug-info.json')}`);
      
      // Final summary of section ID health
      console.log('\n=== Section ID Health Summary ===');
      const sectionIdStats = debugInfo.sectionIdAnalysis;
      console.log(`Total section IDs: ${sectionIdStats.total}`);
      if (sectionIdStats.malformed > 0) {
        console.warn(`⚠️ Malformed section IDs: ${sectionIdStats.malformed}`);
      } else if (sectionIdStats.total > 0) {
        console.log(`✅ All section IDs are well-formed`);
      }
      console.log(`Hierarchical structure: L1=${sectionIdStats.levelOne}, L2=${sectionIdStats.levelTwo}, L3=${sectionIdStats.levelThree}`);
      
    } else {
      console.log('❌ Tests failed - check error messages above');
    }
    
  } catch (error) {
    console.error('\n❌ FATAL ERROR in debug test:');
    console.error(`Message: ${error.message}`);
    console.error(`Stack: ${error.stack}`);
    process.exit(1);
  }
}

// Run if called directly
if (require.main === module) {
  debugMain().catch(error => {
    console.error('Unhandled error in debug test:', error);
    process.exit(1);
  });
}

module.exports = {
  testDocxProcessing,
  testComponents,
  debugMain
};