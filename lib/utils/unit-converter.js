// lib/utils/unit-converter.js - Unit conversion utilities for DOCX processing

/**
 * Convert twip value to points
 * DOCX uses twips (twentieth of a point) for measurements
 * This function converts twips to points for CSS usage
 * 
 * @param {string} twip - Value in twentieths of a point
 * @returns {number} - Value in points
 */
function convertTwipToPt(twip) {
  const twipNum = parseInt(twip, 10) || 0;
  // 20 twips = 1 point
  return twipNum / 20;
}

/**
 * Convert border size to point value
 * DOCX uses 1/8th points for border sizes
 * This function converts these values to points for CSS usage
 * 
 * @param {string} size - Border size value
 * @returns {number} - Size in points
 */
function convertBorderSizeToPt(size) {
  const sizeNum = parseInt(size, 10) || 1;
  // Border size is in 1/8th points
  return sizeNum / 8;
}

/**
 * Get CSS border style from Word border value
 * Maps Word's border style names to CSS border style values
 * 
 * @param {string} value - Word border value
 * @returns {string} - CSS border style
 */
function getBorderTypeValue(value) {
  const borderTypes = {
    'single': 'solid',
    'double': 'double',
    'dotted': 'dotted',
    'dashed': 'dashed',
    'wavy': 'wavy',
    'dashSmallGap': 'dashed',
    'dotDash': 'dashed',
    'dotDotDash': 'dashed',
    'triple': 'double',
    'thinThickSmallGap': 'double',
    'thickThinSmallGap': 'double'
  };
  
  return borderTypes[value] || 'solid';
}

module.exports = {
  convertTwipToPt,
  convertBorderSizeToPt,
  getBorderTypeValue
};