/**
 * Displays the message with a "new version" notice.
 */
function newVersionNotification(versionRelease, versionDate, versionUrl) {
    var d = versionDate;
    var r = versionRelease;

    var msg  = 'A new version is available: ' + r + ' (released on ' + d + ')';
    msg += '\nWould you like to download it now?';

    if (confirm(msg)) {
        window.location = versionUrl;
    } 
}

/**
 * Opens a new windows and points the location to winURL.
 * The window frame is called winName.
 *
 * @todo remove in favor of some AJAX/Js effect.
 */
function openWin(winURL, winName, winFeatures) {
    window.open(winURL, winName, winFeatures);
}
