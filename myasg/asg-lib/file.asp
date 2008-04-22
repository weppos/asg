<%

' 
' = ASP Stats Generator - Powerful and reliable ASP website counter
' 
' Copyright (c) 2003-2008 Simone Carletti <weppos@weppos.net>
' 
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
' 
' The above copyright notice and this permission notice shall be included in
' all copies or substantial portions of the Software.
' 
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
' THE SOFTWARE.
' 
' 
' @category        ASP Stats Generator
' @package         ASP Stats Generator
' @author          Simone Carletti <weppos@weppos.net>
' @copyright       2003-2008 Simone Carletti
' @license         http://www.opensource.org/licenses/mit-license.php
' @version         SVN: $Id$
' 


'
' Copies a file from strSourcePath to strTargetPath.
'
' @param  strSourcePath
' @param  strTargetPath
' @param  blnOverwrite
'
function asgFileCopy(strSourcePath, strTargetPath, blnOverwrite)
  Dim objFso

  Set objFso = Server.CreateObject("Scripting.FileSystemObject")
  objFso.copyFile strSourcePath, strTargetPath, blnOverwrite
  Set objFso = Nothing
end function


'
' Deletes the file at strSourcePath.
'
' @param  strSourcePath
' @param  strTargetPath
'
function asgFileDelete(strSourcePath)
  Dim objFso

  Set objFso = Server.CreateObject("Scripting.FileSystemObject")
  objFso.deleteFile strSourcePath
  Set objFso = Nothing
end function


'
' Deletes the file at strSourcePath
' only if the file exists.
'
' The main difference between this method
' and asgFileDelete is that the latter
' doesn't check whether strSourcePath exists
' and crashes on failure.
'
' @param  strSourcePath
' @param  strTargetPath
'
function asgFileDeleteIfExists(strSourcePath)
  Dim objFso

  Set objFso = Server.CreateObject("Scripting.FileSystemObject")
  if objFso.fileExists(strSourcePath) then
    objFso.deleteFile strSourcePath
  end if
  Set objFso = Nothing
end function


'
' Checks and returns whether strSourcePath file exists.
'
' @param  string  strSourcePath
' @return bool
'
function asgFileExists(strSourcePath)
  Dim objFso, blnResponse

  Set objFso = Server.CreateObject("Scripting.FileSystemObject")
  blnResponse = objFso.fileExists(strSourcePath)
  Set objFso = Nothing
  
  asgFileExists = blnResponse
end function


'
' Moves a file from strSourcePath to strTargetPath.
'
' @param  strSourcePath
' @param  strTargetPath
'
function asgFileMove(strSourcePath, strTargetPath)
  Dim objFso

  Set objFso = Server.CreateObject("Scripting.FileSystemObject")
  objFso.moveFile strSourcePath, strTargetPath
  Set objFso = Nothing
end function


'
' Replace file strTargetPath with strSourcePath.
' If strTargetPath doesn't exist, strSourcePath file is simply renamed.
'
' @param  strSourcePath
' @param  strTargetPath
'
function asgFileReplace(strSourcePath, strTargetPath)
  Dim objFso

  Set objFso = Server.CreateObject("Scripting.FileSystemObject")
  objFso.copyFile strSourcePath, strTargetPath, true
  objFso.deleteFile strSourcePath
  Set objFso = Nothing
end function


%>
