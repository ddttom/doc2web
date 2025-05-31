// doc2web.js - Enhanced version with improved error handling and logging
const fs = require('fs');
const path = require('path');
const { promisify } = require('util');
const mammoth = require('mammoth');
const { markdownify } = require('./markdownify');
// Import extractAndApplyStyles and extractImagesFromDocx from the lib
const { extractAndApplyStyles, extractImagesFromDocx } = require('./lib');

// Promisify fs functions for async/await usage
const readFile = promisify(fs.readFile);
const writeFile = promisify(fs.writeFile);
const stat = promisify(fs.stat);
const readdir = promisify(fs.readdir);
const mkdir = promisify(fs.mkdir);

// Base output directory
const OUTPUT_BASE_DIR = './output';

/**
 * Process command-line arguments
 */
function processArgs() {
  const args = process.argv.slice(2);
  
  if (args.length === 0) {
    console.log('Usage:');
    console.log('  node doc2web.js <file.docx|directory|list-file.txt> [--html-only]');
    console.log('\nOptions:');
    console.log('  --html-only    Generate only HTML output, skip markdown');
    console.log('  --list         Treat the input file as a list of files to process');
    console.log('\nExamples:');
    console.log('  node doc2web.js document.docx');
    console.log('  node doc2web.js ./documents/');
    console.log('  node doc2web.js file-list.txt --list');
    process.exit(1);
  }
  
  const inputPath = args[0];
  const options = {
    htmlOnly: args.includes('--html-only'),
    isList: args.includes('--list')
  };
  
  return { inputPath, options };
}

/**
 * Ensure output directory exists
 * @param {string} outputPath - Directory path to create
 */
async function ensureDirectory(outputPath) {
  try {
    await mkdir(outputPath, { recursive: true });
    console.log(`Created directory: ${outputPath}`);
  } catch (error) {
    if (error.code !== 'EEXIST') {
      console.error(`Failed to create directory ${outputPath}:`, error.message);
      throw error;
    }
  }
}

/**
 * Generate output path for a file
 * @param {string} inputFilePath - Input file path
 * @returns {string} - Output file path under OUTPUT_BASE_DIR
 */
function getOutputPath(inputFilePath) {
  // Normalize path to handle different path separators
  const normalizedPath = path.normalize(inputFilePath);
  
  // Extract directory and file name
  const dirName = path.dirname(normalizedPath);
  const fileName = path.basename(normalizedPath, path.extname(normalizedPath));
  
  // Create output directory path that mirrors input structure
  const outputDir = path.join(OUTPUT_BASE_DIR, dirName);
  
  return {
    directory: outputDir,
    markdownFile: path.join(outputDir, `${fileName}.md`),
    htmlFile: path.join(outputDir, `${fileName}.html`),
    cssFile: path.join(outputDir, `${fileName}.css`)
  };
}

/**
 * Validate DOCX file
 * @param {string} filePath - Path to the file
 * @returns {boolean} - True if valid DOCX file
 */
function validateDocxFile(filePath) {
  try {
    // Check if file exists
    if (!fs.existsSync(filePath)) {
      console.error(`Error: File "${filePath}" not found.`);
      return false;
    }
    
    // Check file extension
    const ext = path.extname(filePath).toLowerCase();
    if (ext !== '.docx') {
      console.error(`Error: File "${filePath}" is not a .docx document. Extension: ${ext}`);
      return false;
    }
    
    // Check if file is readable
    try {
      fs.accessSync(filePath, fs.constants.R_OK);
    } catch (error) {
      console.error(`Error: Cannot read file "${filePath}". Check permissions.`);
      return false;
    }
    
    // Check file size (should be > 0)
    const stats = fs.statSync(filePath);
    if (stats.size === 0) {
      console.error(`Error: File "${filePath}" is empty.`);
      return false;
    }
    
    return true;
  } catch (error) {
    console.error(`Error validating file "${filePath}":`, error.message);
    return false;
  }
}

/**
 * Process a single DOCX file
 * @param {string} filePath - Path to the DOCX file
 * @param {Object} options - Processing options
 */
async function processDocxFile(filePath, options) {
  console.log(`\n=== Processing: ${filePath} ===`);
  
  try {
    // Validate the DOCX file
    if (!validateDocxFile(filePath)) {
      return;
    }
    
    // Get output paths
    const outputPaths = getOutputPath(filePath);
    console.log(`Output directory: ${outputPaths.directory}`);
    
    // Create output directory
    await ensureDirectory(outputPaths.directory);
    
    // Create images directory if it doesn't exist
    const imagesDir = path.join(outputPaths.directory, 'images');
    await ensureDirectory(imagesDir);
    
    // Extract images
    console.log('Extracting images...');
    try {
      await extractImagesFromDocx(filePath, imagesDir);
      console.log('Image extraction completed.');
    } catch (error) {
      console.warn(`Warning: Image extraction failed: ${error.message}`);
      // Continue processing even if image extraction fails
    }
    
    // Extract styles and convert to styled HTML with improved style extraction
    console.log(`Extracting styled content from "${filePath}"...`);
    const cssFilename = path.basename(outputPaths.cssFile);
    
    // Use our enhanced style extractor for better TOC and list handling
    const result = await extractAndApplyStyles(filePath, cssFilename, {
      enhanceAccessibility: true,
      preserveMetadata: true,
      trackChangesMode: 'show'
    }, outputPaths.directory);
    
    console.log(`Conversion completed. HTML length: ${result.html.length}, CSS length: ${result.styles.length}`);
    
    // Validate the results
    if (!result.html || result.html.length < 100) {
      console.error('ERROR: HTML output is too short or empty');
      console.log('HTML content preview:', result.html.substring(0, 500));
      return;
    }
    
    if (!result.styles || result.styles.length < 100) {
      console.warn('WARNING: CSS output is very short');
    }
    
    // Save the HTML with link to external CSS
    console.log(`Saving HTML to "${outputPaths.htmlFile}"...`);
    await writeFile(outputPaths.htmlFile, result.html, 'utf8');
    
    // Save separate CSS file
    console.log(`Saving CSS to "${outputPaths.cssFile}"...`);
    await writeFile(outputPaths.cssFile, result.styles, 'utf8');
    
    console.log(`✓ Styled HTML saved to "${outputPaths.htmlFile}"`);
    console.log(`✓ CSS styles saved to "${outputPaths.cssFile}"`);
    
    // If HTML only mode is enabled, skip markdown conversion
    if (options.htmlOnly) {
      console.log('HTML-only mode enabled. Skipping markdown conversion.');
      return;
    }
    
    // Convert HTML to markdown
    console.log('Converting HTML to markdown...');
    try {
      const markdown = markdownify(result.html);
      
      if (!markdown || markdown.trim().length === 0) {
        console.warn('WARNING: Markdown conversion resulted in empty content');
        return;
      }
      
      // Write markdown to file with explicit UTF-8 encoding
      await writeFile(outputPaths.markdownFile, markdown, 'utf8');
      console.log(`✓ Markdown saved to "${outputPaths.markdownFile}"`);
    } catch (markdownError) {
      console.error(`Error converting to markdown: ${markdownError.message}`);
      // Don't fail the entire process if markdown conversion fails
    }
    
    console.log(`✓ Processing completed successfully for "${filePath}"`);
    
  } catch (error) {
    console.error(`✗ Error processing file "${filePath}":`, error.message);
    console.error('Stack trace:', error.stack);
    
    // Try to provide helpful debugging information
    if (error.message.includes('Invalid DOCX file')) {
      console.error('The file may be corrupted or not a valid DOCX document.');
    } else if (error.message.includes('ENOENT')) {
      console.error('File or directory not found. Check the path.');
    } else if (error.message.includes('EACCES')) {
      console.error('Permission denied. Check file/directory permissions.');
    }
  }
}

/**
 * Process a directory of DOCX files
 * @param {string} dirPath - Path to the directory
 * @param {Object} options - Processing options
 */
async function processDirectory(dirPath, options) {
  try {
    console.log(`\n=== Processing directory: ${dirPath} ===`);
    
    // Validate directory exists
    if (!fs.existsSync(dirPath)) {
      console.error(`Error: Directory "${dirPath}" not found.`);
      return;
    }
    
    // Check if it's actually a directory
    const stats = await stat(dirPath);
    if (!stats.isDirectory()) {
      console.error(`Error: "${dirPath}" is not a directory.`);
      return;
    }
    
    // Read directory contents
    const files = await readdir(dirPath);
    let docxFound = false;
    let processedCount = 0;
    let errorCount = 0;
    
    console.log(`Found ${files.length} items in directory`);
    
    // Process each item
    for (const file of files) {
      const filePath = path.join(dirPath, file);
      
      try {
        const itemStats = await stat(filePath);
        
        if (itemStats.isDirectory()) {
          // Recursively process subdirectories
          console.log(`\nEntering subdirectory: ${file}`);
          await processDirectory(filePath, options);
        } else if (path.extname(filePath).toLowerCase() === '.docx') {
          docxFound = true;
          await processDocxFile(filePath, options);
          processedCount++;
        }
      } catch (error) {
        console.error(`Error processing item "${filePath}":`, error.message);
        errorCount++;
      }
    }
    
    if (!docxFound) {
      console.log(`No DOCX files found in directory: ${dirPath}`);
    } else {
      console.log(`\n=== Directory processing summary ===`);
      console.log(`Processed: ${processedCount} files`);
      console.log(`Errors: ${errorCount} files`);
    }
    
  } catch (error) {
    console.error(`Error processing directory "${dirPath}":`, error.message);
  }
}

/**
 * Process a file containing a list of file paths
 * @param {string} listFilePath - Path to the list file
 * @param {Object} options - Processing options
 */
async function processFileList(listFilePath, options) {
  try {
    console.log(`\n=== Processing file list from: ${listFilePath} ===`);
    
    // Validate list file exists
    if (!fs.existsSync(listFilePath)) {
      console.error(`Error: List file "${listFilePath}" not found.`);
      return;
    }
    
    // Read the list file
    const fileContent = await readFile(listFilePath, 'utf8');
    
    // Split by newlines and filter out empty lines
    const filePaths = fileContent.split('\n')
      .map(line => line.trim())
      .filter(line => line.length > 0 && !line.startsWith('#')); // Ignore comments
    
    if (filePaths.length === 0) {
      console.log('No file paths found in the list file.');
      return;
    }
    
    console.log(`Found ${filePaths.length} file paths in the list.`);
    
    // Process each DOCX file in the list
    let docxCount = 0;
    let processedCount = 0;
    let errorCount = 0;
    
    for (const filePath of filePaths) {
      // Only process DOCX files
      if (path.extname(filePath).toLowerCase() === '.docx') {
        docxCount++;
        try {
          await processDocxFile(filePath, options);
          processedCount++;
        } catch (error) {
          console.error(`Error processing file from list "${filePath}":`, error.message);
          errorCount++;
        }
      } else {
        console.log(`Skipping non-DOCX file: ${filePath}`);
      }
    }
    
    console.log(`\n=== File list processing summary ===`);
    console.log(`DOCX files found: ${docxCount}`);
    console.log(`Successfully processed: ${processedCount}`);
    console.log(`Errors: ${errorCount}`);
    
    if (docxCount === 0) {
      console.log('No DOCX files found in the list.');
    }
    
  } catch (error) {
    console.error(`Error processing file list "${listFilePath}":`, error.message);
  }
}

/**
 * Main function to start processing
 */
async function main() {
  const startTime = Date.now();
  
  try {
    console.log('doc2web - DOCX to HTML/Markdown Converter');
    console.log('==========================================');
    
    const { inputPath, options } = processArgs();
    
    console.log(`Input: ${inputPath}`);
    console.log(`Options: ${JSON.stringify(options)}`);
    
    // Determine input type and process accordingly
    if (options.isList || inputPath.endsWith('.txt')) {
      // Process as list file
      await processFileList(inputPath, options);
    } else {
      // Check if input is a file or directory
      try {
        const stats = await stat(inputPath);
        
        if (stats.isDirectory()) {
          // Process directory
          await processDirectory(inputPath, options);
        } else if (path.extname(inputPath).toLowerCase() === '.docx') {
          // Process single DOCX file
          await processDocxFile(inputPath, options);
        } else {
          console.error(`Error: Unsupported file type "${path.extname(inputPath)}".`);
          console.log('Supported inputs: .docx files, directories, or .txt list files.');
          process.exit(1);
        }
      } catch (error) {
        if (error.code === 'ENOENT') {
          console.error(`Error: Input path "${inputPath}" not found.`);
        } else {
          console.error(`Error accessing input path "${inputPath}":`, error.message);
        }
        process.exit(1);
      }
    }
    
    const endTime = Date.now();
    const duration = (endTime - startTime) / 1000;
    
    console.log('\n==========================================');
    console.log(`Processing completed in ${duration.toFixed(2)} seconds`);
    console.log('==========================================');
    
  } catch (error) {
    console.error('\n==========================================');
    console.error('FATAL ERROR:', error.message);
    console.error('Stack trace:', error.stack);
    console.error('==========================================');
    process.exit(1);
  }
}

// Run the main function
if (require.main === module) {
  main().catch(error => {
    console.error('Unhandled error in main:', error);
    process.exit(1);
  });
}

module.exports = {
  processDocxFile,
  processDirectory,
  processFileList,
  getOutputPath,
  ensureDirectory,
  validateDocxFile
};
