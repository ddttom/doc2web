#!/usr/bin/env node
// doc2web-run.js - Command-line tool to run the doc2web converter

const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');
const readline = require('readline');

// Create a readline interface for user input
const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

/**
 * Display menu and get user choice
 */
function showMenu() {
  console.log('\n=== doc2web Converter ===\n');
  console.log('1. Process a single DOCX file');
  console.log('2. Process a directory of DOCX files');
  console.log('3. Process a list of files from a text file');
  console.log('4. Search and process DOCX files (using find command)');
  console.log('5. Exit');
  
  rl.question('\nSelect an option (1-5): ', (choice) => {
    processChoice(choice);
  });
}

/**
 * Process user's menu choice
 * @param {string} choice - User's menu selection
 */
function processChoice(choice) {
  switch (choice) {
    case '1':
      processSingleFile();
      break;
    case '2':
      processDirectory();
      break;
    case '3':
      processListFile();
      break;
    case '4':
      searchAndProcess();
      break;
    case '5':
      console.log('Exiting...');
      rl.close();
      break;
    default:
      console.log('Invalid choice. Please try again.');
      showMenu();
      break;
  }
}

/**
 * Process a single DOCX file
 */
function processSingleFile() {
  rl.question('Enter the path to the DOCX file: ', (filePath) => {
    if (!filePath || !fs.existsSync(filePath)) {
      console.log('Error: File not found.');
      return showMenu();
    }
    
    rl.question('Generate HTML only? (y/n): ', (htmlOnly) => {
      const options = htmlOnly.toLowerCase() === 'y' ? '--html-only' : '';
      
      try {
        console.log(`Processing file: ${filePath}`);
        execSync(`node doc2web.js "${filePath}" ${options}`, { stdio: 'inherit' });
        console.log('Processing complete!');
      } catch (error) {
        console.error('Error processing file:', error.message);
      }
      
      showMenu();
    });
  });
}

/**
 * Process a directory of DOCX files
 */
function processDirectory() {
  rl.question('Enter the directory path: ', (dirPath) => {
    if (!dirPath || !fs.existsSync(dirPath)) {
      console.log('Error: Directory not found.');
      return showMenu();
    }
    
    rl.question('Generate HTML only? (y/n): ', (htmlOnly) => {
      const options = htmlOnly.toLowerCase() === 'y' ? '--html-only' : '';
      
      try {
        console.log(`Processing directory: ${dirPath}`);
        execSync(`node doc2web.js "${dirPath}" ${options}`, { stdio: 'inherit' });
        console.log('Processing complete!');
      } catch (error) {
        console.error('Error processing directory:', error.message);
      }
      
      showMenu();
    });
  });
}

/**
 * Process a list of files from a text file
 */
function processListFile() {
  rl.question('Enter the path to the list file: ', (filePath) => {
    if (!filePath || !fs.existsSync(filePath)) {
      console.log('Error: File not found.');
      return showMenu();
    }
    
    rl.question('Generate HTML only? (y/n): ', (htmlOnly) => {
      const options = htmlOnly.toLowerCase() === 'y' ? '--html-only' : '';
      
      try {
        console.log(`Processing file list from: ${filePath}`);
        execSync(`node doc2web.js "${filePath}" --list ${options}`, { stdio: 'inherit' });
        console.log('Processing complete!');
      } catch (error) {
        console.error('Error processing file list:', error.message);
      }
      
      showMenu();
    });
  });
}

/**
 * Search for DOCX files using find command and process them
 */
function searchAndProcess() {
  rl.question('Enter the directory to search in: ', (searchDir) => {
    if (!searchDir || !fs.existsSync(searchDir)) {
      console.log('Error: Directory not found.');
      return showMenu();
    }
    
    rl.question('Enter search pattern (e.g., "*.docx" or leave empty for all .docx files): ', (pattern) => {
      const searchPattern = pattern || '*.docx';
      
      rl.question('Generate HTML only? (y/n): ', (htmlOnly) => {
        const options = htmlOnly.toLowerCase() === 'y' ? '--html-only' : '';
        
        try {
          // Create temp file for find results
          const tempFile = path.join(process.cwd(), `docx-find-results-${Date.now()}.txt`);
          
          console.log(`Searching for ${searchPattern} in ${searchDir}...`);
          
          // Use platform-specific find command
          if (process.platform === 'win32') {
            // Windows
            execSync(`dir "${searchDir}\\${searchPattern}" /s /b > "${tempFile}"`, { stdio: 'inherit' });
          } else {
            // Unix/Linux/MacOS
            execSync(`find "${searchDir}" -name "${searchPattern}" > "${tempFile}"`, { stdio: 'inherit' });
          }
          
          console.log(`Found files listed in: ${tempFile}`);
          console.log('Processing files...');
          
          execSync(`node doc2web.js "${tempFile}" --list ${options}`, { stdio: 'inherit' });
          
          // Clean up temp file
          fs.unlinkSync(tempFile);
          console.log('Processing complete!');
        } catch (error) {
          console.error('Error searching and processing files:', error.message);
        }
        
        showMenu();
      });
    });
  });
}

// Start the application
showMenu();
