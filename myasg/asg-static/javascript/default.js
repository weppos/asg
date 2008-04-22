/**
 * = ASP Stats Generator - Powerful and reliable ASP website counter
 * 
 * Copyright (c) 2003-2008 Simone Carletti <weppos@weppos.net>
 * 
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 * 
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 * 
 * 
 * @category        ASP Stats Generator
 * @package         ASP Stats Generator
 * @author          Simone Carletti <weppos@weppos.net>
 * @copyright       2003-2008 Simone Carletti
 * @license         http://www.opensource.org/licenses/mit-license.php
 * @version         SVN: $Id$
 */

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
