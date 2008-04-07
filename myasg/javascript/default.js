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
