<%

' Tracking file
'Response.Write(vbCrLf & "var file='/myasg/count.asp';")
Response.Write(vbCrLf & "var file='http://www.weppos.net/myasg/count.asp';")

' Referer
' Check if the page is frame based
Response.Write(vbCrLf & "if(document.referrer){")
Response.Write(vbCrLf & "var r=document.referrer;")
Response.Write(vbCrLf & "}else{")
Response.Write(vbCrLf & "var r=top.document.referrer;")
Response.Write(vbCrLf & "}")
' Filter
Response.Write(vbCrLf & "r=escape(r);")
' Prevent Referer locking
Response.Write(vbCrLf & "if((r=='null') || (r=='unknown') || (r=='undefined')) r='';")

' Navigation page
Response.Write(vbCrLf & "var u='' + escape(document.URL); ")

' Screen resolution
Response.Write(vbCrLf & "var w=screen.width; ")
Response.Write(vbCrLf & "var h=screen.height; ")

' Color depth
Response.Write(vbCrLf & "var v=navigator.appName; ")
Response.Write(vbCrLf & "if (v != 'Netscape') {c=screen.colorDepth;}")
Response.Write(vbCrLf & "else {c=screen.pixelDepth;}")

' Anti-Aliasing Fonts
Response.Write(vbCrLf & "var fs = window.screen.fontSmoothingEnabled;")
' Page title
Response.Write(vbCrLf & "var t = escape(document.title);")
' Ricava il supporto per Java abilitato
Response.Write(vbCrLf & "var j = navigator.javaEnabled();")

' Passa la stringa con i valori
' Standard - 2.1
' Response.Write "info='w=' + w + '&h=' + h + '&c=' + c + '&r=' + r + '&u='+ u;"
' Avanzata - 3.x
Response.Write(vbCrLf & "info='w=' + w + '&h=' + h + '&c=' + c + '&r=' + r + '&u='+ u + '&fs=' + fs + '&t=' + t + '&j=' + j;")

'// Richiama l'img e passa i valori
Response.Write(vbCrLf & "document.open();")
Response.Write(vbCrLf & "document.write('<img src=' + file + '?' + info + ' border=0>');")
Response.Write(vbCrLf & "document.close();")

%>