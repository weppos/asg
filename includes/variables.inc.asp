<%

'-------------------------------------------------------------------------------'
'	ASP Stats Generator															'
'	Copyright © 2003-2005 Simone Carletti [ aka weppos ]						'
'-------------------------------------------------------------------------------'

' Dimension variables
Dim strLayerAdvDataSorting
Dim strAsgAppend		' Holds a temp querystring
Dim strAsgMode			' Holds the sorting mode : all | month
Dim strAsgPeriod		' Holds the report period : format mm-yyyy
Dim intAsgPeriodY		' Holds the year : format yyyy
Dim intAsgPeriodM		' Holds the month : format mm

' Graph
Dim intAsgBarPart					' 
Dim intAsgMaxRecValue			' Holds the max value for the recordset item
Dim intAsgTotMonthHits			' Holds the total monthly value of hits
Dim intAsgTotMonthVisits		' Holds the total monthly value of visits

' Read report settings from querystring
strAsgSortBy = Request.QueryString("sortby")
strAsgSortOrder = formatSetting("sortorder", "DESC")
strAsgMode = formatSetting("mode", "month")
intAsgPeriodM = formatSetting("periodm", Month(dtmAsgNow))
intAsgPeriodY = formatSetting("periody", Year(dtmAsgNow))
strAsgPeriod = formatSetting("period", "")

%>
