<% Option Explicit
response.buffer = false%>
<!--#Include file="../wbstat3_class.asp"-->
<%
dim tempo:tempo=timer
DIm oBrowser,Fso,FileName,File,Line,Count,realfilename,engineversion
Set oBrowser = new wbstatclass
oBrowser.SetPath "../wbstat3_spec/"
oBrowser.Options.IncludeUserAgent = false
oBrowser.Options.Cache = false
engineversion = oBrowser.Version
filename = request.QueryString("file")
if filename = "" then
	realfilename = "wbstat3_multi_list.txt"
	FileName = Server.MapPath("wbstat3_multi_list.txt")
else
	realfilename = filename
	filename = Server.MapPath(request.QueryString("file"))
end if


response.write "<p><div style=""text-align:center;font-family:Tahoma;"">"
response.write "<span style=""font-size:12px;font-weight:bold"">WBstat 3.x - Copyleft 2003-2004 Simone Cingano - <a href=""http://www.imente.it/wbstat"">http://www.imente.it/wbstat</a></span>"
response.write "<span style=""font-size:10px;"">"
response.write "<br /><br />"
response.write "WBStat &raquo; <b>Version:</b> " & engineversion
response.write "<br />"
response.write "WBStat &raquo; <b>Processed File:</b> " & realfilename
response.write "</span>"
response.write "</div></p>"

Set Fso = Server.CreateObject("Scripting.FileSystemObject")
Count = 1
Set File = Fso.OpenTextFile(FileName,1)
	While Not File.AtEndOfStream
		Line = File.ReadLine()
		if not(left(line,1) = "#") then
			oBrowser.SetUserAgent(Line)
			oBrowser.Eval()
			oBrowser.Debug "Key",Count
			Count = Count + 1
		end if
	Wend
Set oBrowser = Nothing
response.write "<p><div style=""text-align:center;font-family:Tahoma;"">"
response.write "<span style=""font-size:12px;font-weight:bold"">WBstat 3.x - Copyleft 2003-2004 Simone Cingano - <a href=""http://www.imente.it/wbstat"">http://www.imente.it/wbstat</a></span>"
response.write "<span style=""font-size:10px;"">"
response.write "<br /><br />"
response.write "WBStat &raquo; <b>Version:</b> " & engineversion
response.write "<br />"
response.write "WBStat &raquo; <b>Processed File:</b> " & realfilename
response.write "<br />"
response.write "WBstat &raquo; <b>Processing Time:</b> " & formatnumber(timer - tempo,4) & "s"
response.write "</span>"
response.write "</div></p>"
%>