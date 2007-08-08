<% @LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--#include file="w2k3_config.asp" -->
<%

'-------------------------------------------------------------------------------'
'	ASP Stats Generator															'
'	Copyright © 2003-2005 Simone Carletti [ aka weppos ]						'
'-------------------------------------------------------------------------------'


'// WARNING! Program protection.
'	Changing default values may allow users to access the page.
Call checkPermission("False", "False", "False", appAsgSecurity)


'/**
' * 
' * 
' * @param 		()  	- 
' * @return 	() 
' *
' * @since 		3.0
' *
' * @author		Simone Carletti <carletti@weppos.net>
' */ 
public function exportXls(ByRef rs)

	Dim return
	Dim ii
	
	' Header
	return = vbCrLf & "<tr class=""exceldatatitle"">"
	for ii = 0 to rs.Fields.Count - 1
		return = return & "<td>" &_
			objAsgRs.Fields(x).Name &_
		"</td>"
	next
	return = return & "</tr>"
	
	' Content
	return = return & vbCrLf & "<tr class=""exceldatavalue""><td>" &_
		objAsgRs.GetString(,, "</td><td>", "</td></tr>" & vbCrLf & "<tr class=""exceldatavalue""><td>", "-") &_
		"</td></tr>"
		
	exportXls = return
	
end function



Dim strAsgType
Dim strAsgTable
Dim strAsgContent

strAsgType = Request.QueryString("type")
strAsgTable = Request.QueryString("table")


select case strAsgType

	case "xls"
		
		strAsgContent = exportXls(objAsgRs)
				
		objAsgRs.Close
		
		Set objAsgRs = Nothing
		objAsgConn.Close
		Set objAsgConn = Nothing
		
		%>
		<style type="text/css">
		/* Excel */
		.exceldatatitle {
			font-family: Tahoma, Arial, Helvetica, sans-serif;
			font-size: 12px;
			color: #0000FF;
		}
		.exceldatavalue {
			font-family: Tahoma, Arial, Helvetica, sans-serif;
			font-size: 11px;
			color: #000000;
		}
		</style>
		<%

		Response.ContentType = "application/vnd.ms-excel"
		Response.AddHeader "content-disposition", "inline; filename=" & strAsgExcelFile & ".xls"
		Response.Write ("<table>" & strAsgContent &  "</table>")
	
	case else
		'
	
end select

%>