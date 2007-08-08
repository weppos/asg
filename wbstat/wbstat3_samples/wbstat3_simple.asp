<% Option Explicit
response.buffer = false%>
<!--#Include file="../wbstat3_class.asp"-->
<%
response.write "<p><div style=""text-align:center;font-family:Tahoma;""><span style=""font-size:12px;font-weight:bold"">WBstat 3.1 - Copyleft 2003-2004 Simone Cingano - <a href=""http://www.imente.it/wbstat"">http://www.imente.it/wbstat</a></span></p>"
dim tempo:tempo=timer
dim oBrowser
'Set oBrowser = CreateWBstat("../wbstat3_spec/",false,"Sconosciuto",1,0,True,False,False,False,True,True,True,True,True,True,True,True,True,True,True)
Set oBrowser = CreateWBstatSimple("../wbstat3_spec/",False,"Sconosciuto",True)
oBrowser.Debug "Key",false
response.write "<p><div style=""text-align:center;font-family:Tahoma;""><span style=""font-size:12px;font-weight:bold"">WBstat 3.1 - Copyleft 2003-2004 Simone Cingano - <a href=""http://www.imente.it/wbstat"">http://www.imente.it/wbstat</a></span><span style=""font-size:10px;""><br />totale tempo di elaborazione " & formatnumber(timer - tempo,4) & "s" & "</span></div></p>"
%>