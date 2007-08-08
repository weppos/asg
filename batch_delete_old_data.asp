<% @LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--#include file="w2k3_config.asp" -->
<!--#include file="includes/inc_array_table.asp" -->
<%

'-------------------------------------------------------------------------------'
'	ASP Stats Generator															'
'	Copyright © 2003-2005 Simone Carletti [ aka weppos ]						'
'-------------------------------------------------------------------------------'


'// WARNING! Program protection.
'	Changing default values may allow users to access the page.
Call checkPermission("False", "False", "False", appAsgSecurity)


Dim aryAsgTableWarning(2, 10)		' Holds the array containing warning informations

' Set warning record level
aryAsgTableWarning(1, 1) = 5000
aryAsgTableWarning(1, 2) = 2000
aryAsgTableWarning(1, 3) = 750
aryAsgTableWarning(1, 4) = 2000
aryAsgTableWarning(1, 5) = 500
aryAsgTableWarning(1, 6) = 3500
aryAsgTableWarning(1, 7) = 3500
aryAsgTableWarning(1, 8) = 2000
aryAsgTableWarning(1, 9) = 750
aryAsgTableWarning(1, 10) = 3500

Dim i

' Execute a loop to count the records of each table
Dim intAsgTableLoop
for intAsgTableLoop = 1 to Ubound(aryAsgTable)

	aryAsgTableWarning(0,intAsgTableLoop) = false
	
	' Initialise SQL string to count records
	strAsgSQL = "SELECT COUNT(*) FROM "& ASG_TABLE_PREFIX & aryAsgTable(intAsgTableLoop, 1)
	' Open Rs
	objAsgRs.Open strAsgSQL, objAsgConn
	' Set the number of total hits
	if not objAsgRs.EOF then 

		' Check warning limit
		if Cint(objAsgRs(0)) > Cint(aryAsgTableWarning(1, intAsgTableLoop)) then
			aryAsgTableWarning(0, intAsgTableLoop) = true
		end If
		aryAsgTableWarning(2, intAsgTableLoop) = Cint(objAsgRs(0))

	end If
	' Close Rs
	objAsgRs.Close

next

' Reset objects
Set objAsgRs = Nothing
objAsgConn.Close
Set objAsgConn = Nothing

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
	<div id="layout">
		<div id="layout_menutitle"><%= "<span class=""menusubtitle"">" & MENUGROUP_Administration & " &raquo; " & MENUSECTION_Maintenance & " &raquo;</span> " & MENUSECTION_BatchDeleteOldData %></div>
		<div id="layout_content">

<%

' :: Open tlayout :: MENUSECTION_BatchDeleteOldData
Response.Write(builTableTlayout("", "open", MENUSECTION_BatchDeleteOldData))
	
	' 
	if Request.QueryString("msg") = "error" then Response.Write("<p class=""errortext"" style=""text-align: center;"">" & TXT_Error_CheckTableMatching & "</p>")
	
	' Loop database tables
	for intAsgTableLoop = 0 to Ubound(aryAsgTable)
	
		' Change password
		strAsgTmpLayer = "<form action=""batch_delete_old_data_execute.asp"" method=""get"" name=""frmDel" & intAsgTableLoop & """>"
		strAsgTmpLayer = strAsgTmpLayer & "<p style=""text-align: center;""><input type=""hidden"" name=""table"" value=""" & intAsgTableLoop & """ />" 

			' Table title and description
			' strAsgTmpLayer = strAsgTmpLayer & "<span class=""notetext"">" & aryAsgTable(intAsgTableLoop,1) & "</span>&nbsp;(" & aryAsgTableWarning(2,intAsgTableLoop) & "&nbsp;" & TXT_Records & ")&nbsp;-&nbsp;" & aryAsgTable(intAsgTableLoop,1) & "</p>"
			strAsgTmpLayer = strAsgTmpLayer & aryAsgTable(intAsgTableLoop, 2) & "</p>"

		strAsgTmpLayer = strAsgTmpLayer & "<table align=""center"" border=""0"" cellspacing=""1"" cellpadding=""1""><tr>" 
		strAsgTmpLayer = strAsgTmpLayer & "<td align=""right"" width=""60%"">" & TXT_Deldata & "&nbsp;"
		strAsgTmpLayer = strAsgTmpLayer & "<select name=""timerange""><option value=""full"">" & TXT_Deldata_all & "</option>"
			
			' In these cases show a different select
			' depending on database structure
			if aryAsgTable(intAsgTableLoop, 1) <> "IP" then
				strAsgTmpLayer = strAsgTmpLayer & "<option value=""0"">" & TXT_Deldata_OlderThan_monthC & "</option>"
				strAsgTmpLayer = strAsgTmpLayer & "<option value=""1"">" & TXT_Deldata_OlderThan_month1 & "</option>"
				for i = 2 to 12
				strAsgTmpLayer = strAsgTmpLayer & "<option value=""" & i & """>" & Replace(TXT_Deldata_OlderThan_weekN, "$var1$", i) & "</option>"
				next 
		
			elseif aryAsgTable(intAsgTableLoop, 1) = "Detail" then
				strAsgTmpLayer = strAsgTmpLayer & "<option value=""0"">" & TXT_Deldata_OlderThan_weekC & "</option>"
				strAsgTmpLayer = strAsgTmpLayer & "<option value=""1"">" & TXT_Deldata_OlderThan_week1 & "</option>"
				for i = 2 to 12
				strAsgTmpLayer = strAsgTmpLayer & "<option value=""" & i & """>" & Replace(TXT_Deldata_OlderThan_weekN, "$var1$", i) & "</option>"
				next 
			
			end if
			
		strAsgTmpLayer = strAsgTmpLayer & "</select></td>"
		strAsgTmpLayer = strAsgTmpLayer & "<td align=""left""><input type=""image"" src=""" & STR_ASG_SKIN_PATH_IMAGE & "icons/remove.png"" name=""delete"" value=""deletenormal"" onclick=""return confirm('" & TXT_Deldata_conf & "');"" /></td>"

		' For Detail table show a week select
		if aryAsgTable(intAsgTableLoop, 1) = "Detail" then
		
			strAsgTmpLayer = strAsgTmpLayer & "</tr><tr>"
			strAsgTmpLayer = strAsgTmpLayer & "<td align=""right"" width=""60%"">" & TXT_Deldata & "&nbsp;"
			strAsgTmpLayer = strAsgTmpLayer & "<select name=""weekrange""><option value=""""></option>"
			strAsgTmpLayer = strAsgTmpLayer & "<option value=""0"">" & TXT_Deldata_OlderThan_weekC & "</option>"
			strAsgTmpLayer = strAsgTmpLayer & "<option value=""1"">" & TXT_Deldata_OlderThan_week1 & "</option>"
			for i = 2 to 12
			strAsgTmpLayer = strAsgTmpLayer & "<option value=""" & i & """>" & Replace(TXT_Deldata_OlderThan_weekN, "$var1$", i) & "</option>"
			next 
			strAsgTmpLayer = strAsgTmpLayer & "</select></td>"
			strAsgTmpLayer = strAsgTmpLayer & "<td align=""left""><input type=""image"" src=""" & STR_ASG_SKIN_PATH_IMAGE & "icons/remove.png"" name=""delete"" value=""deleteweek"" onclick=""return confirm('" & TXT_Deldata_conf & "');"" /></td>"
		
		end if

		strAsgTmpLayer = strAsgTmpLayer & "</tr></table>"
		strAsgTmpLayer = strAsgTmpLayer & "</form>"

	
		' Show the icon if the table need a reset
		if aryAsgTableWarning(0,intAsgTableLoop) then
			' :: Create the layer ::
			Response.Write(buildLayer("layerDelete" & intAsgTableLoop, "<img src=""" & STR_ASG_SKIN_PATH_IMAGE & "/icons/warning.png"" border=""0"" align=""middle"" alt=""" & TXT_Warning & """  onmouseover=""stm(Warning[1],Style[3])"" onmouseout=""htm()"" />&nbsp;" & aryAsgTable(intAsgTableLoop,1), aryAsgTableWarning(2,intAsgTableLoop) & "&nbsp;" & TXT_Records, strAsgTmpLayer))
		else
			' :: Create the layer ::
			Response.Write(buildLayer("layerDelete" & intAsgTableLoop, aryAsgTable(intAsgTableLoop,1), aryAsgTableWarning(2,intAsgTableLoop) & "&nbsp;" & TXT_Records, strAsgTmpLayer))
		end If

	next

' :: Open tlayout :: MENUSECTION_BatchDeleteOldData
Response.Write(builTableTlayout("", "close", ""))

%>

		</div>
	</div>
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