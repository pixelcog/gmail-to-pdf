/**
 * A collection of utilities for use with Google Drive and the DriveApp object
 *
 * @module DriveUtils
 */

var DriveUtils = new function() {

  /**
   * Get a folder or create it if it does not exist
   *
   * @method getFolder
   * @param {String} name
   * @return {Object} Folder object
   */
  this.getFolder = function(name) {
    var folders = DriveApp.getFoldersByName(name);
    return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
  };
};

