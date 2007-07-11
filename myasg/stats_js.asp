<%

'// Definisce dove si trova il file per il conteggio
Response.Write "var file='/myasg/count.asp';"

'// Ricava il Referer = Pagina di Provenienza
Response.Write "f='' + escape(document.referrer);"

'// Ricava la pagina attuale nel sito
Response.Write "u='' + escape(document.URL); "

'// Ricava la risoluzione video
Response.Write "var w=screen.width; "
Response.Write "var h=screen.height; "

'// Ricava il nome del browser per valutare la profondità di colore
Response.Write "v=navigator.appName; "
Response.Write "if (v != 'Netscape') {c=screen.colorDepth;}"
Response.Write "else {c=screen.pixelDepth;}"

'// Ricava Anti-Aliasing Fonts
Response.Write "var fs = window.screen.fontSmoothingEnabled;"

'// Ricava il supporto per Java abilitato
Response.Write "j=navigator.javaEnabled();"

'// Passa la stringa con i valori
'Response.Write "info='w=' + w + '&h=' + h + '&c=' + c + '&r=' + f + '&u='+ u + '&fs=' + fs + '&b=' + b + '&x=' + x;"
Response.Write "info='w=' + w + '&h=' + h + '&c=' + c + '&r=' + f + '&u='+ u + '&fs=' + fs + '&j=' + j;"

'// Richiama l'img e passa i valori
Response.Write "document.open();"
Response.Write "document.write('<img src=' + file + '?'+info+ ' border=0>');"
Response.Write "document.close();"

%>