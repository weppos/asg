<%

'-------------------------------------------------------------------------------'
'	ASP Stats Generator															'
'	Copyright © 2003-2005 Simone Carletti [ aka weppos ]						'
'-------------------------------------------------------------------------------'


' ***** APPLICATION VARIABLES *****
' Use application variables to store configuration settings.
' Enabling this feature will improve loading performances.
Dim blnApplicationConfig
blnApplicationConfig = false

' ***** COOKIE PREFIX *****
' Prefix that all ASG cookies will have.
' This is useful if you run multiple copies of the program on the same site so that 
' cookies don't interfer with each other.
Const ASG_COOKIE_PREFIX = "asg_"

' ***** APPLICATION PREFIX *****
' Prefix that all ASG application variables will have.
' This is useful if you run multiple copies of the program on the same site so that 
' application variables don't interfer with each other.
Const ASG_APPLICATION_PREFIX = "asg_"

' ***** TABLE PREFIX *****
' Prefix that all ASG database tables will have.
' This is useful if you want to run multiple versions or copies on the same database 
' or if you are sharing the database with other applications.
Const ASG_TABLE_PREFIX = "asg_"


' ***** UNKNOWN INFORMATION *****
' This string will replace empty detail variables.
Const ASG_UNKNOWN = "(unknown)"

' ***** OWNSERVER *****
' This string will replace the real referer URL 
' if the 'Disable internal referer' option is enabled.
Const ASG_OWNSERVER = "(ownserver)"


' ***** ELABORATION TIME *****
' Show elaboration time at the bottom of the pages.
Const ASG_ELABORATION_TIME = true

' ***** BUILD INFORMATION *****
' Show 'build' info of the version at the bottom of the pages
' near the application version.
Const ASG_BUILDINFO = true

' ***** HIDE TOOLBAR *****
' Hide the menu in the following pages: default, login.
' This is useful to preserve the program from curious.
Const ASG_MENUBAR_HIDELOGIN = true

' ***** DISABLE TOOLBAR *****
' Disable the admin toolbar if the user is not logged in.
' This is useful to improve security settings.
Const ASG_ADMINBAR_NOLOGIN = true

' ***** SETUPLOCK CHECKING *****
' The program will not check if the locking file exists and
' there will be no control on the setup file.
Const ASG_SETUPLOCK = true						' NOW DISABLED

' ***** SETUPLOCK FILENAME *****
' Setup lock file name.
Const ASG_SETUPLOCK_FILE = "setuplock.asg"		' NOW DISABLED


' *****  *****
' 
Const ASG_CACHE = true


' ***** DEBUG *****
' Enter the program in debug mode
Const ASG_DEBUG_MODE = false

' ***** DEBUG SEARCH ENGINES *****
' Debug search engine from referers
' Test referers with a regular expression pattern to find
' unknown new search engines.
Const ASG_DEBUG_SEARCHENGINES = false


' ***** DOCTYPE *****
' You can define page DOCTYPE.

' Const STR_ASG_PAGE_DOCTYPE = "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN"">"

Const STR_ASG_PAGE_DOCTYPE = "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">"

' ***** GRAPH COLUMNS MAX WIDTH *****
' Max width of graph area in pixel.
Const ASG_COL_MAXWIDTH = 600

' ***** GRAPH COLUMNS MAX HEIGHT *****
' Max height of graph area in pixel.
Const ASG_COL_MAXHEIGHT = 200

' ***** ADVANCED VISITOR DETAILS *****
' Show advanced details in the visitors report such as user agent.
Const ASG_VISITOR_ADVANCED = true


' You can use a normal variable instad of a constant to overwrite
' the variable if you need to change something on the fly.
' Dim STR_ASG_PAGE_DOCTYPE
' STR_ASG_PAGE_DOCTYPE = "<!-- IE in quirk mode -->" & vbCrLf & "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" " & vbCrLf & " ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">"

' ***** FOLDER WITH WRITE PERMISSION *****
' This folder needs write permission.
' You can enable write permission by selecting 'write' checkbox from IIS Console.
' If you have no access to IIS ask your provider to enable permissions for you
' or use the default one provided by your hoster, usually called /public .
Const STR_ASG_PATH_FOLDER_WR = "write_permission/"


%>