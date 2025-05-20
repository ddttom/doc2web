// doc2web.js - Enhanced version with improved style extraction
const fs = require('fs');
const path = require('path');
const { promisify } = require('util');
const mammoth = require('mammoth');
const { markdownify } = require('./markdownify');
const { extractAndApplyStyles } = require('./style-extractor');

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
  // Create directory recursively
  try {
    await mkdir(outputPath, { recursive: true });
  } catch (error) {
    if (error.code !== 'EEXIST') {
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
 * Process a single DOCX file
 * @param {string} filePath - Path to the DOCX file
 * @param {Object} options - Processing options
 */
async function processDocxFile(filePath, options) {
  try {
    console.log(`Processing: ${filePath}`);
    
    // Check if file exists and is a DOCX file
    if (!fs.existsSync(filePath)) {
      console.error(`Error: File "${filePath}" not found.`);
      return;
    }
    
    if (path.extname(filePath).toLowerCase() !== '.docx') {
      console.error(`Error: File "${filePath}" is not a .docx document.`);
      return;
    }
    
    // Get output paths
    const outputPaths = getOutputPath(filePath);
    
    // Create output directory
    await ensureDirectory(outputPaths.directory);
    
    // Create images directory if it doesn't exist
    const imagesDir = path.join(outputPaths.directory, 'images');
    await ensureDirectory(imagesDir);
    
    // Extract images
    await processImages(filePath, imagesDir);
    
    // Extract styles and convert to styled HTML with improved style extraction
    console.log(`Extracting styled content from "${filePath}"...`);
    const cssFilename = path.basename(outputPaths.cssFile);
    
    // Use our enhanced style extractor for better TOC and list handling
    const { html, styles } = await extractAndApplyStyles(filePath, cssFilename);
    
    // Save the HTML with link to external CSS
    await writeFile(outputPaths.htmlFile, html, 'utf8');
    
    // Save separate CSS file
    await writeFile(outputPaths.cssFile, styles, 'utf8');
    
    console.log(`Styled HTML saved to "${outputPaths.htmlFile}"`);
    console.log(`CSS styles saved to "${outputPaths.cssFile}"`);
    
    // If HTML only mode is enabled, skip markdown conversion
    if (options.htmlOnly) {
      console.log('HTML-only mode enabled. Skipping markdown conversion.');
      return;
    }
    
    // Convert HTML to markdown
    console.log('Converting HTML to markdown...');
    const markdown = markdownify(html);
    
    // Write markdown to file with explicit UTF-8 encoding
    await writeFile(outputPaths.markdownFile, markdown, 'utf8');
    
    console.log(`Markdown conversion complete. Output saved to "${outputPaths.markdownFile}"`);
    
  } catch (error) {
    console.error(`Error processing file "${filePath}":`, error.message);
  }
}

/**
 * Process images from a DOCX file
 * @param {string} docxPath - Path to the DOCX file
 * @param {string} outputDir - Directory to save images
 */
async function processImages(docxPath, outputDir) {
  try {
    const options = {
      path: docxPath,
      encoding: 'utf8',
      convertImage: mammoth.images.imgElement(function(image) {
        return image.read("base64").then(function(imageBuffer) {
          const extension = image.contentType.split('/')[1];
          const filename = `image-${image.altText || Date.now()}.${extension}`;
          
          // Save the image to output directory
          const imagePath = path.join(outputDir, filename);
          fs.writeFileSync(imagePath, Buffer.from(imageBuffer, 'base64'));
          
          return {
            src: `./images/${filename}`,
            alt: image.altText || 'Image'
          };
        });
      })
    };
    
    // Convert to HTML with images
    await mammoth.convertToHtml(options);
    console.log('Images extracted successfully.');
  } catch (error) {
    console.error('Error extracting images:', error.message);
  }
}

/**
 * Process a directory of DOCX files
 * @param {string} dirPath - Path to the directory
 * @param {Object} options - Processing options
 */
async function processDirectory(dirPath, options) {
  try {
    console.log(`Processing directory: ${dirPath}`);
    
    // Read directory contents
    const files = await readdir(dirPath);
    let docxFound = false;
    
    // Process each DOCX file
    for (const file of files) {
      const filePath = path.join(dirPath, file);
      const stats = await stat(filePath);
      
      if (stats.isDirectory()) {
        // Recursively process subdirectories
        await processDirectory(filePath, options);
      } else if (path.extname(filePath).toLowerCase() === '.docx') {
        docxFound = true;
        await processDocxFile(filePath, options);
      }
    }
    
    if (!docxFound) {
      console.log(`No DOCX files found in directory: ${dirPath}`);
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
    console.log(`Processing file list from: ${listFilePath}`);
    
    // Read the list file
    const fileContent = await readFile(listFilePath, 'utf8');
    
    // Split by newlines and filter out empty lines
    const filePaths = fileContent.split('\n')
      .map(line => line.trim())
      .filter(line => line.length > 0);
    
    if (filePaths.length === 0) {
      console.log('No file paths found in the list file.');
      return;
    }
    
    console.log(`Found ${filePaths.length} file paths in the list.`);
    
    // Process each DOCX file in the list
    let docxCount = 0;
    for (const filePath of filePaths) {
      // Only process DOCX files
      if (path.extname(filePath).toLowerCase() === '.docx') {
        docxCount++;
        await processDocxFile(filePath, options);
      }
    }
    
    if (docxCount === 0) {
      console.log('No DOCX files found in the list.');
    } else {
      console.log(`Processed ${docxCount} DOCX files from the list.`);
    }
    
  } catch (error) {
    console.error(`Error processing file list "${listFilePath}":`, error.message);
  }
}

/**
 * Main function to start processing
 */
async function main() {
  try {
    const { inputPath, options } = processArgs();
    
    // Determine input type and process accordingly
    if (options.isList || inputPath.endsWith('.txt')) {
      // Process as list file
      await processFileList(inputPath, options);
    } else {
      // Check if input is a file or directory
      const stats = await stat(inputPath);
      
      if (stats.isDirectory()) {
        // Process directory
        await processDirectory(inputPath, options);
      } else if (path.extname(inputPath).toLowerCase() === '.docx') {
        // Process single DOCX file
        await processDocxFile(inputPath, options);
      } else {
        console.error(`Error: Unsupported file type "${path.extname(inputPath)}".`);
        console.log('Supported file types: .docx or directory or .txt list file.');
      }
    }
    
    console.log('Processing complete!');
    
  } catch (error) {
    console.error('Error:', error.message);
    process.exit(1);
  }
}

// Run the main function
if (require.main === module) {
  main();
}

module.exports = {
  processDocxFile,
  processDirectory,
  processFileList,
  getOutputPath,
  ensureDirectory
};
