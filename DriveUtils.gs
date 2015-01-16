/*
 * DriveUtils
 * ==========
 *
 * A collection of utilities for use with Google Drive and the DriveApp object.
 *
 * To utilize this library, select Resources > Libraries... and enter the following project key:
 * MUDdULBfiLdgEZ13bA9paOlVaKzeOjMwH
 */

/**
 * Get a folder or create it if it does not exist
 *
 * @method getFolder
 * @param {string} name
 * @return {Folder} object
 */
function getFolder(name) {
  var folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}
