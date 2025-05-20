// lib/utils/common-utils.js - Common utility functions used across modules

/**
 * Get leader character based on leader type
 * Maps Word's leader type codes to actual characters
 * 
 * @param {string} leader - Leader type
 * @returns {string} - Leader character
 */
function getLeaderChar(leader) {
  if (!leader) return '.';
  
  switch (leader) {
    case 'dot':
    case '1':
      return '.';
    case 'hyphen':
    case '2':
      return '-';
    case 'underscore':
    case '3':
      return '_';
    case 'heavy':
    case '4':
      return '=';
    case 'middleDot':
    case '5':
      return '·';
    default:
      return '.';
  }
}

module.exports = {
  getLeaderChar
};