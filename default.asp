<% @LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--#include file="w2k3_config.asp" -->
<%

'-------------------------------------------------------------------------------'
'	ASP Stats Generator															'
'	Copyright © 2003-2005 Simone Carletti [ aka weppos ]						'
'-------------------------------------------------------------------------------'


' Reset objects
Set objAsgRs = Nothing
objAsgConn.Close
Set objAsgConn = Nothing

' Hide menubar
if ASG_MENUBAR_HIDELOGIN then blnAsgShowToolbar = false


%>
<%= STR_ASG_PAGE_DOCTYPE %>
<html>
<head>
<title><%= appAsgSiteName %> | powered by ASP Stats Generator v<%= ASG_VERSION %></title>
<%= STR_ASG_PAGE_CHARSET %>
<meta name="copyright" content="Copyright (C) 2003-2005 Carletti Simone" />
<!--#include file="includes/meta.inc.asp" -->

<!-- ASP Stats Generator v. <%= ASG_VERSION %> is created and developed by Simone Carletti.
To download your Free copy visit the official site http://www.weppos.com/asg/ -->

</head>

<body>
<!--#include file="includes/header.asp" -->

<div align="center">
<%

Response.Write(vbCrLf & "<div>" &_
	"<h1>" & TXT_error & "</h1>" &_
	"<p>" & TXT_error_defaultpage & "</p>" &_
	"<p><a href=""" & Request.ServerVariables("REMOTE_HOST") & """>" & Request.ServerVariables("REMOTE_HOST") & "</a></p>" &_
	"</div>")

%>

</div>

<br /></div>
<!-- / body -->
<%

' Footer
Response.Write(vbCrLf & "<div id=""footer"">")
' ***** START WARNING - REMOVAL or MODIFICATION IN PART or ALL OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT	******
' ***** INIZIO AVVERTENZA - RIMOZIONE o MODIFICA PARZIALE/TOTALE DEL CODICE COMPORTA VIOLAZIONE DELLA LICENZA  	******
Response.Write(vbCrLf & "<br /><div style=""text-align: center;"">Powered by <a href=""http://www.weppos.com/asg/"" title=""ASP Stats Generator"">ASP Stats Generator</a> v" & ASG_VERSION & " ") 
if ASG_BUILDINFO then Response.Write("build " & ASG_VERSION_BUILD)
Response.Write(vbCrLf & "<br />Copyright &copy; 2003-2005 <a href=""http://www.weppos.com/"">weppos</a></div>")
if ASG_ELABORATION_TIME then Response.Write("<div class=""elabtime"">" & Replace(TXT_elabtime, "$time$", FormatNumber(Timer() - startAsgElab, 4)) & "</div>")
' ***** END WARNING - REMOVAL or MODIFICATION IN PART or ALL OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT	******
' ***** FINE AVVERTENZA - RIMOZIONE o MODIFICA PARZIALE/TOTALE DEL CODICE COMPORTA VIOLAZIONE DELLA LICENZA  ******
Response.Write(vbCrLf & "</div>")

%>
<!--#include file="includes/footer.asp" -->
</body></html>