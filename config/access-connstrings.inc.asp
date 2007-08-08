<%

'-------------------------------------------------------------------------------'
'	ASP Stats Generator															'
'	Copyright © 2003-2005 Simone Carletti [ aka weppos ]						'
'-------------------------------------------------------------------------------'


'--------------------------------------------------------------------------------
' Microsoft Access 97
'--------------------------------------------------------------------------------

' Microsoft Access Driver
'strAsgConn = "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & strAsgMapPath

' OLEDB 3.51
'strAsgConn = "Provider=Microsoft.Jet.OLEDB.3.51; Data Source=" & strAsgMapPath

'--------------------------------------------------------------------------------
' Microsoft Access 2000, 2002, 2003	
'--------------------------------------------------------------------------------

' OLEDB 4.0
strAsgConn = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & strAsgMapPath

%>
