<% Option Explicit
response.buffer = false %>
<!--#Include file="wbstat3_class.asp"-->
<%

response.write "<p><div style=""text-align:center;font-family:Tahoma;""><span style=""font-size:12px;font-weight:bold"">WBstat 3.0beta - Copyleft 2003-2004 Simone Cingano - <a href=""http://www.imente.it/wbstat"">http://www.imente.it/wbstat</a></span></p>"
dim tempo:tempo=timer
Dim oBrowser
Dim objConn, strConn, strSQL, objRs, Count
Set oBrowser = new wbstatclass
Set objConn = Server.CreateObject("ADODB.Connection")
Set objRS = Server.CreateObject("ADODB.Recordset")

If Request.QueryString("elabora") = "Elabora" AND Len(Request.QueryString("tabella")) > 0 AND Len(Request.QueryString("database")) > 0 Then
	
	Dim tabella, database
	If Request.QueryString("default") = "true" then
		tabella = "tblSt_Detail"
		database = Request.QueryString("database") & "dbStats.mdb"
	Else
		tabella = Request.QueryString("tabella")
		database = Request.QueryString("database")
	End If
	strConn = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath(database)

	oBrowser.SetPath "wbstat3_spec/"
	oBrowser.Options.IncludeUserAgent = false
	oBrowser.Options.Cache = false
	Count = -1
	
	objConn.Open strConn
	strSQL = "SELECT DISTINCT(User_Agent) FROM " & tabella & ""
	objRs.Open strSQL, objConn
	Do While Not objRs.EOF
		oBrowser.SetUserAgent(objRs("User_Agent"))
		oBrowser.Eval()
		oBrowser.Debug "Key",Count
		Count = Count + 1
	objRs.MoveNext
	Loop
	objRs.Close
	objConn.Close

Else

Response.Write("<p><div style=""text-align:center;font-family:Tahoma;""><span style=""font-size:12px"">")
Response.Write("<form name=""frmMultidatabase"" method=""get"">")
Response.Write("Tabella <br /> <input name=""tabella"" value=""" & Request.QueryString("tabella") & """ type=""text"" lenght=""25"" /><br />")
Response.Write("Database <br /> <input name=""database"" value=""" & Request.QueryString("database") & """ type=""text"" lenght=""25"" /><br /> ")
Response.Write("Default <input name=""default"" value=""true"" type=""checkbox"" /><br /> ")
Response.Write("<br /> <input name=""elabora"" value=""Elabora"" type=""submit"" lenght=""25"" /><br /> ")
Response.Write("</form>")
Response.Write("</div></p>")

End If

Set objRs = Nothing
Set objConn = Nothing
Set oBrowser = Nothing

response.write "<p><div style=""text-align:center;font-family:Tahoma;""><span style=""font-size:12px;font-weight:bold"">WBstat 3.0beta - Copyleft 2003-2004 Simone Cingano - <a href=""http://www.imente.it/wbstat"">http://www.imente.it/wbstat</a></span><span style=""font-size:10px;""><br />totale tempo di elaborazione " & formatnumber(timer - tempo,4) & "s" & " - tool by <a href=""http://www.weppos.com/"" target=""_blank"">weppos</span></div></p>"
	
%>