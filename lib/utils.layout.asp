<%

'-------------------------------------------------------------------------------'
'	ASP Stats Generator															'
'	Copyright © 2003-2005 Simone Carletti [ aka weppos ]						'
'-------------------------------------------------------------------------------'


' Common layout variables
Dim strAsgTmpLayer


'-----------------------------------------------------------------------------------------
' Build the .EOF line for the report recordset.
'-----------------------------------------------------------------------------------------
public function buildTableContNoRecord(colspanValue, message)
	
	Dim layout
	
	' Check for an alternative message.
	' If there's no alternative use the default one.
	if message = "standard" then 
		message = TXT_Nodata_db
	elseif message = "search" then
		message = TXT_Nodata_search
	end If 
			
	layout = layout & "<tr class=""treport_row"">"
	layout = layout & "<td class=""treport_col"" align=""center"" colspan=""" & colspanValue & """>" & message & "</td>"
	layout = layout & "</tr>"
	
	' Return the function
	buildTableContNoRecord = layout

end function
			

'-----------------------------------------------------------------------------------------
' Costruisci Riga Tabella Contenuti - Debug automatico icone non riconosciute
'-----------------------------------------------------------------------------------------
' Function:	
' Date: 	14.03.2004 | 
' Comment:	
'-----------------------------------------------------------------------------------------
Function buildTableContCheckIcon(ByVal colspanValue, ByVal iconType, ByVal pageNum)
	
	Dim strAsgTableContent
	strAsgTableContent = ""
	
	strAsgTableContent = strAsgTableContent & vbCrLf & "<!-- Informazioni icone non riconosciute -->"
	strAsgTableContent = strAsgTableContent & vbCrLf & "<tr class=""normaltext"" align=""center"" valign=""top"">"
	strAsgTableContent = strAsgTableContent & vbCrLf & "<td width=""100%"" colspan=""" & colspanValue & """><br /><img src=""" & STR_ASG_SKIN_PATH_IMAGE & iconType & ".asp?icon=checkicon&page=" & pageNum & """ alt="""" /><br /></td>"
	strAsgTableContent = strAsgTableContent & vbCrLf & "</tr>"
	strAsgTableContent = strAsgTableContent & vbCrLf & "<!-- Informazioni icone non riconosciute -->"
			  
			
	If iconType = "browser" AND Session("blnAsgIconBrowser" & pageNum) <> "notified" AND appAsgDebugIcon Then
	
		Response.Write(strAsgTableContent)
	
	ElseIf iconType = "os" AND Session("blnAsgIconOs" & pageNum) <> "notified" AND appAsgDebugIcon Then
	
		Response.Write(strAsgTableContent)
	
	ElseIf iconType = "engine" AND Session("blnAsgIconEngine" & pageNum) <> "notified" AND appAsgDebugIcon Then
	
		Response.Write(strAsgTableContent)
	
	End If
			
End Function
			

'-----------------------------------------------------------------------------------------
' Costruisci Riga Tabella Contenuti - Spaziatore finale
'-----------------------------------------------------------------------------------------
' Function:	
' Date: 	14.03.2004 | 
' Comment:	
'-----------------------------------------------------------------------------------------
Function buildTableContEndSpacer(ByVal colspanValue)

	Response.Write(vbCrLf & "<tr bgcolor=""" & STR_ASG_SKIN_TABLE_TITLE_BGCOLOUR & """>")
	Response.Write(vbCrLf & "<td colspan=""" & colspanValue & """ background=""" & STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_TITLE_BGIMAGE & """ height=""2""></td>")
	Response.Write(vbCrLf & "</tr>")

End Function
			

'-----------------------------------------------------------------------------------------
' Build table Bar 
'-----------------------------------------------------------------------------------------
' Function:	
' Date: 	28.10.2004 | 06.11.2004
' Comment:	
'-----------------------------------------------------------------------------------------
Function buildTableBar(barTitle, menuTitle)

	Response.Write(vbCrLf & "<div class=""table_bar""><span class=""menubartitle"">" & Ucase(menuTitle) & "</span> &raquo; " & barTitle & "</div>")

End Function
			

'-----------------------------------------------------------------------------------------
' Table title
'-----------------------------------------------------------------------------------------
Function buildTableTitle(rowStyle)

	' check style class value
	if Len(rowStyle) > 0 then Response.Write("class=""" & rowStyle & """ ")

End Function
			

'-----------------------------------------------------------------------------------------
' Table content row with rollover effect
'-----------------------------------------------------------------------------------------
Function buildTableContRollover(rowStyle)

	' check style class value
	if Len(rowStyle) > 0 then Response.Write("class=""" & rowStyle & """ ")
	' check background image value
	if not Len(STR_ASG_SKIN_TABLE_CONT_BGIMAGE) > 0 then
		Response.Write("onmouseover=""this.style.background='" & STR_ASG_SKIN_TABLE_CONT_BGOVER & "';"" onmouseout=""this.style.background='" & STR_ASG_SKIN_TABLE_CONT_BGOUT & "';"" ")
	end if

End Function

'/**
' * Build the layer to search reports.
' * 
' * @return 	
' *
' * @since 		3.0
' * @see		buildLayer()
' *
' * @author		Simone Carletti <carletti@weppos.net>
' */ 
public function buildLayerAdvDataSorting()
	
	Dim pageMax, pageMid, pageLast, pageIndex, frompageIndex, topageIndex
	Dim txt_pagePosition, url, querystring
	Dim layout
	
	txt_pagePosition = TXT_navigation_schema
	txt_pagePosition = Replace(txt_pagePosition, "$pageCurrent$", page)
	txt_pagePosition = Replace(txt_pagePosition, "$pageCount$", objAsgRs.PageCount)
	url = Request.ServerVariables("URL")

	layout = "<p>"
	
	Dim objItem 

	' Run and build the new querystring checking all
	' items stored at the moment in the querystring
	for Each objItem in Request.QueryString
		If Not objItem = "page" Then 
			querystring = querystring & "&" & objItem & "=" & Request.QueryString(objItem)
		End If
	next
	
	pageMax = 10 
	pageMid = (pageMax \ 2) + 1
	pageLast = objAsgRs.pageCount
	
	if pageMax > pageLast then
		for pageIndex = 1 to pageLast
			if CInt(pageIndex) = CInt(page) then
				layout = layout & " - <strong>" & pageIndex & "</strong> "
			else
				layout = layout & " - <a href=""" & url & "?page=" & pageIndex & querystring & """ title=""" & Replace(TXT_gotoPage_number_schema, "$page$", pageIndex) & """>" & pageIndex & "</a> "
			end if
		next 
		layout = layout & " - "
	else
		if CInt(pageMid) < CInt(page) then
			frompageIndex = page - pageMid + 1
			topageIndex = page + pageMid - 1
			if topageIndex > pageLast then 
				topageIndex = pageLast
				frompageIndex = topageIndex - pageMax + 1
			end if
		else
			frompageIndex = 1
			topageIndex = pageMax
		end if

		if frompageIndex <> 1 then
			layout = layout & " <a href=""" & url & "?page=1" & querystring & """ title=""" & Replace(TXT_gotoPage_number_schema, "$page$", "1") & """>&lt;&lt;</a> "
		end if
		for pageIndex = frompageIndex to topageIndex
			if CInt(pageIndex) = CInt(page) then
				layout = layout & " - <strong>" & pageIndex & "</strong> "
			else
				layout = layout & " - <a href=""" & url & "?page=" & pageIndex & querystring & """ title=""" & Replace(TXT_gotoPage_number_schema, "$page$", pageIndex) & """>" & pageIndex & "</a> "
			end if
		next 
		if frompageIndex < pageLast - pageMax then
			layout = layout & " - <a href=""" & url & "?page=" & pageLast & querystring & """ title=""" & Replace(TXT_gotoPage_number_schema, "$page$", pageLast) & """>&gt;&gt;</a> "
		else
			layout = layout & " - "
		end if
	end if

	layout = layout & "</p>"
	
	' Call the build layer function to fill the layer with the group condition content
	layout = buildLayer("layerAdvDataSorting", LABEL_Navigation, txt_pagePosition, layout)
	
	' Return function
	buildLayerAdvDataSorting = layout

end function

'/**
' * Build the layer to search reports.
' * 
' * @return 	
' *
' * @since 		3.0
' * @see		buildLayer()
' *
' * @author		Simone Carletti <carletti@weppos.net>
' */ 
public function buildLayerAdvDetDataSorting()

	Dim pageMax, pageMid, pageLast, pageIndex, frompageIndex, topageIndex
	Dim txt_pagePosition, url, querystring
	Dim layout
	
	txt_pagePosition = TXT_navigation_schema
	txt_pagePosition = Replace(txt_pagePosition, "$pageCurrent$", page)
	txt_pagePosition = Replace(txt_pagePosition, "$pageCount$", objAsgRs2.PageCount)
	url = Request.ServerVariables("URL")

	layout = "<p>"
	
	Dim objItem 

	' Run and build the new querystring checking all
	' items stored at the moment in the querystring
	for Each objItem in Request.QueryString
		If Not objItem = "detpage" Then 
			strQuerystring = strQuerystring & "&" & objItem & "=" & Request.QueryString(objItem)
		End If
	next
	
	pageMax = 10 
	pageMid = (pageMax \ 2) + 1
	pageLast = objAsgRs2.pageCount
	
	if pageMax > pageLast then
		for pageIndex = 1 to pageLast
			if CInt(pageIndex) = CInt(detpage) then
				layout = layout & " - <strong>" & pageIndex & "</strong> "
			else
				layout = layout & " - <a href=""" & url & "?detpage=" & pageIndex & querystring & """ title=""" & Replace(TXT_gotoPage_number_schema, "$page$", pageIndex) & """>" & pageIndex & "</a> "
			end if
		next 
		layout = layout & " - "
	else
		if CInt(pageMid) < CInt(detpage) then
			frompageIndex = detpage - pageMid + 1
			topageIndex = detpage + pageMid - 1
			if topageIndex > pageLast then 
				topageIndex = pageLast
				frompageIndex = topageIndex - pageMax + 1
			end if
		else
			frompageIndex = 1
			topageIndex = pageMax
		end if

		if frompageIndex <> 1 then
			layout = layout & " <a href=""" & url & "?detpage=1" & querystring & """ title=""" & Replace(TXT_gotoPage_number_schema, "$page$", "1") & """>&lt;&lt;</a> "
		end if
		for pageIndex = frompageIndex to topageIndex
			if CInt(pageIndex) = CInt(detpage) then
				layout = layout & " - <strong>" & pageIndex & "</strong> "
			else
				layout = layout & " - <a href=""" & url & "?detpage=" & pageIndex & querystring & """ title=""" & Replace(TXT_gotoPage_number_schema, "$page$", pageIndex) & """>" & pageIndex & "</a> "
			end if
		next 
		if frompageIndex < pageLast - pageMax then
			layout = layout & " - <a href=""" & url & "?detpage=" & pageLast & querystring & """ title=""" & Replace(TXT_gotoPage_number_schema, "$page$", pageLast) & """>&gt;&gt;</a> "
		else
			layout = layout & " - "
		end if
	end if

	layout = layout & "</p>"
	
	' Call the build layer function to fill the layer with the group condition content
	layout = buildLayer("layerAdvDataSorting", LABEL_Navigation, strPagePosition, layout)
	
	' Return function
	buildLayerAdvDetDataSorting = layout

end function
			

'-----------------------------------------------------------------------------------------
' Table content row legend
'-----------------------------------------------------------------------------------------
' Function:	
' Date: 	12.11.2004 | 
' Comment:	
'-----------------------------------------------------------------------------------------
Function buildTableContLegend(colspanValue)

	Response.Write(vbCrLf & "<tr bgcolor=""" & STR_ASG_SKIN_TABLE_CONT_BGCOLOUR & """ align=""center"">")
	Response.Write(vbCrLf & "<td colspan=""" & colspanValue & """ background=""" & STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_CONT_BGIMAGE & """><p>")
	Response.Write(vbCrLf & "<img src=""" & STR_ASG_SKIN_PATH_IMAGE & "bar_graph_image_hits_h.gif"" width="""" height=""9"" alt=""" & TXT_pageviews & """ align=""middle"" />&nbsp;" & TXT_pageviews)
	Response.Write(vbCrLf & "<img src=""" & STR_ASG_SKIN_PATH_IMAGE & "bar_graph_image_visits_h.gif"" width="""" height=""9"" alt=""" & TXT_visits & """ align=""middle"" />&nbsp;" & TXT_visits)
	Response.Write(vbCrLf & "</p></td>")
	Response.Write(vbCrLf & "</tr>")

End Function
			

'-----------------------------------------------------------------------------------------
' Create the report legend layer 
'-----------------------------------------------------------------------------------------
public function buildLayerReportLegend()

	Dim layout
	
	layout = layout & vbcrLf & "<div class=""treport_legend"">"
	layout = layout & "&nbsp;<img src=""" & STR_ASG_SKIN_PATH_IMAGE & "bar_graph_image_hits_h.gif"" width="""" height=""9"" alt=""" & TXT_pageviews & """ align=""middle"" />&nbsp;" & TXT_pageviews
	layout = layout & "&nbsp;<img src=""" & STR_ASG_SKIN_PATH_IMAGE & "bar_graph_image_visits_h.gif"" width="""" height=""9"" alt=""" & TXT_visits & """ align=""middle"" />&nbsp;" & TXT_visits
	layout = layout & "</div>"

	' Return the function
	buildLayerReportLegend = layout

end function
			

'-----------------------------------------------------------------------------------------
' Create the line to swap layer display status. 
'-----------------------------------------------------------------------------------------
public function buildSwapDisplay(rowID, rowtitle)

	Dim layout
	
	layout = layout & vbcrLf & "<span class=""swapdisplay"" onclick=""swapDisplayTr('" & rowID & "');"">" & rowtitle & "</span>"

	' Return the function
	buildSwapDisplay = layout

end function


'-----------------------------------------------------------------------------------------
' Build the layer to choose the report period and mode.
'-----------------------------------------------------------------------------------------
public function buildLayerForm(action)

	Dim layout
	
	if action = "open" then
		layout = vbCrLf & "<form name=""frmNavy"" action=""?"" method=""get"">"
	elseif action = "close" then
		layout = vbCrLf & "</form>"
	end if

	' Return the function
	buildLayerForm = layout

end function


'-----------------------------------------------------------------------------------------
' 
'-----------------------------------------------------------------------------------------
private function searchQuerystringItem(argNoappend, argItem)

	Dim i
	Dim aryItem
	aryItem = Split(argNoappend, "|")

	for i = 0 to Ubound(aryItem)
		if argItem = aryItem(i) then exit for
	next
	
	' Return the function
	if i > Ubound(aryItem) then
		searchQuerystringItem = false
	else
		searchQuerystringItem = true
	end if

end function


Const STR_NOAPPEND_LAYER = "periodm|periody|showsubmit|mode|type|group"
Const STR_NOAPPEND_LAYER_SEARCH = "searchfor|searchin|showsearch|page"

'/**
' * Build the layer to search reports.
' * 
' * @param		
' * @param		
' * @return 	string § layout of the search layer.
' *
' * @since 		3.0
' *
' * @author		Simone Carletti <carletti@weppos.net>
' */ 
public function buildLayerSearch(actionQuerystring, databaseFields)

	Dim aryField		' Holds the array of the database fields to search in
	Dim objItem			' Holds the querystring collection item
	Dim i
	Dim layout
	Dim lvTmp
	Dim lvRemoveSearch

	lvTmp = ""
	if Len(asgSearchfor) > 0 AND Len(asgSearchin) > 0 then
		' Fill the legend info with the search terms
		lvTmp = TXT_Search_fieldquery
		lvTmp = Replace(lvTmp, "$var1$", asgSearchfor)
		lvTmp = Replace(lvTmp, "$var2$", asgSearchin)
	
		' Allow user to come back to the page without the search mode enable
		lvRemoveSearch = appendToQuerystring("searchfor||searchin||showsearch")
	
	else
	
		lvRemoveSearch = Request.ServerVariables("QUERY_STRING")
	
	end if
			
	' Split the variable to an array containing all database fields to search in
	aryField = Split(databaseFields, "|")		

'	if Len(Request.QueryString("searchfor")) > 0 then
		layout = layout & vbCrLf & "<div id=""layerSearch"" style=""display: block;"">"
'	else
'		layout = layout & vbCrLf & "<div id=""labelSearch"" style=""display: none;"">"
'	end if
	
	' layout = layout & vbCrLf & "<table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"">"
	' layout = layout & vbCrLf & "<tr><td>"
	layout = layout & vbCrLf & "<form name=""frmSearch"" action=""?" & actionQuerystring & """ method=""get"">"
	layout = layout & vbCrLf & "<fieldset class=""fldlayer""><legend class=""fldlegendtext""><span class=""fldlegendtitle"">" & LABEL_Searchform & " :: </span>" & lvTmp & "</legend>"
	layout = layout & vbCrLf & "<p><input type=""text"" name=""searchfor"" value=""" & Request.QueryString("searchfor") & """ size=""10"" />"
	'layout = layout & "&nbsp;<img src=""" & STR_ASG_SKIN_PATH_IMAGE & "icons/remove_button.png"" alt=""" & TXT_Search_cleanfield & """ border=""0"" onclick=""document.frmSearch.searchfor.value='';"" onmouseover=""stm(Info[],Style[1])"" onmouseout=""htm()"" />"
	layout = layout & "&nbsp;<img src=""" & STR_ASG_SKIN_PATH_IMAGE & "icons/remove_button.png"" alt=""" & TXT_Search_cleanfield & """ border=""0"" onclick=""document.frmSearch.searchfor.value='';"" />"
	layout = layout & vbCrLf & "<select name=""searchin"">"
	' Show fields in a select
	for i = 0 to Ubound(aryField)
		layout = layout & "<option value=""" & aryField(i) & """"
		If Request.QueryString("searchin") = aryField(i) then layout = layout & " selected"
		layout = layout & " >" & aryField(i) & "</option>"
	next
	layout = layout & "</select>"
	' layout = layout & vbCrLf & "<input type=""image"" src=""" & STR_ASG_SKIN_PATH_IMAGE & "icons/search_button.png"" name=""showsearch"" value=""" & TXT_Search & """ align=""middle"" />"
	layout = layout & vbCrLf & "<img src=""" & STR_ASG_SKIN_PATH_IMAGE & "icons/search_button.png"" name=""showsearch"" alt=""" & TXT_Search & """ onclick=""document.frmSearch.submit();"" />&nbsp;"
	
	' Allow user to come back to the page without the search mode enable
	layout = layout & "<a href=""?" & lvRemoveSearch & """ title=""" & TXT_Search_delquery & """><img src=""" & STR_ASG_SKIN_PATH_IMAGE & "icons/search_delquery_button.png"" alt=""" & TXT_Search_delquery & """ border=""0"" /></a>"
	
	' Loop all Querystring items
	for each objItem in Request.QueryString
		' if not objItem = "searchfor" AND not objItem = "searchin" AND not objItem = "showsearch" then
		if not searchQuerystringItem(STR_NOAPPEND_LAYER_SEARCH, objItem) then
		layout = layout & vbCrLf & "<input type=""hidden"" name=""" & objItem & """ value=""" & Request.QueryString(objItem) & """ />"
		end If
	next
	
	layout = layout & vbCrLf & "</p></fieldset></form>"
	'layout = layout & "</td></tr></table>"
	layout = layout & "</div>"
	
	' Return the function
	buildLayerSearch = layout

end function

'/**
' * Build the layer to choose the report period and view mode.
' * 
' * @param		
' * @param		
' * @return 	string § layout of the  report period and view mode.
' *
' * @since 		3.0
' *
' * @author		Simone Carletti <carletti@weppos.net>
' */ 
%><!--#include file="month.array.inc.asp" --><%
public function buildLayerPeriod()

	Dim objItem
	Dim ii
	Dim layout

	layout = "<p>"

	' Show all months in a select
	layout = layout & vbCrLf & "<select name=""periodm"">"
	for ii = 1 to Ubound(aryAsgMonth, 2)
		layout = layout & "<option value=""" & ii & """" 
		if intAsgPeriodM = ii then 
			' Check the DTD
			if InStr(STR_ASG_PAGE_DOCTYPE, "XHTML") > 0 then
				layout = layout & " selected=""selected"""
			else
				layout = layout & " selected"
			end if
		end if
		layout = layout & " >" & aryAsgMonth(1,ii) & "</option>"
	next
	layout = layout & "</select>"

	' Show all years in a select
	layout = layout & vbCrLf & "<select name=""periody"">"
	for ii = Year(appAsgProgramSetup) to dtmAsgYear 
		layout = layout & "<option value=""" & ii & """" 
		if intAsgPeriodY = ii then 
			' Check the DTD
			if InStr(STR_ASG_PAGE_DOCTYPE, "XHTML") > 0 then
				layout = layout & " selected=""selected"""
			else
				layout = layout & " selected"
			end if
		end if
		layout = layout & " >" & ii & "</option>"
	next
	layout = layout & "</select>"

	layout = layout & vbCrLf & "<input type=""submit"" name=""showsubmit"" value=""" & TXT_button_show & """ />"
	
	' Loop all Querystring items
	for each objItem in Request.QueryString
		' if not objItem = "periodm" AND not objItem = "periody" AND not objItem = "showsubmit" AND not objItem = "mode" then
		if not searchQuerystringItem(STR_NOAPPEND_LAYER, objItem) then
		layout = layout & vbCrLf & "<input type=""hidden"" name=""" & objItem & """ value=""" & Request.QueryString(objItem) & """ />"
		end if
	next
	layout = layout & "</p>"
	
	' Call the build layer function to fill the layer with the group condition content
	layout = buildLayer("layerPeriod", LABEL_ViewPeriod, aryAsgMonth(1, intAsgPeriodM) & "&nbsp;" & intAsgPeriodY, layout)
	
	' Return the function
	buildLayerPeriod = layout

end function


'-----------------------------------------------------------------------------------------
' Build the layer to choose the report year period.
'-----------------------------------------------------------------------------------------
public function buildLayerPeriodY()

	Dim objItem
	Dim i
	Dim layout

	layout = "<p>"

	' Show all years in a select
	layout = layout & vbCrLf & "<select name=""periody"">"
	for i = Year(appAsgProgramSetup) to dtmAsgYear 
		layout = layout & "<option value=""" & i & """" 
		if intAsgPeriodY = i then 
			' Check the DTD
			if InStr(STR_ASG_PAGE_DOCTYPE, "XHTML") > 0 then
				layout = layout & " selected=""selected"""
			else
				layout = layout & " selected"
			end if
		end if
		layout = layout & " >" & i & "</option>"
	next
	layout = layout & "</select>"

	layout = layout & vbCrLf & "<input type=""submit"" name=""showsubmit"" value=""" & TXT_button_show & """ />"
	
	' Loop all Querystring items
	for each objItem in Request.QueryString
		' if not objItem = "periodm" AND not objItem = "periody" AND not objItem = "showsubmit" AND not objItem = "mode" then
		if not searchQuerystringItem(STR_NOAPPEND_LAYER, objItem) then
		layout = layout & vbCrLf & "<input type=""hidden"" name=""" & objItem & """ value=""" & Request.QueryString(objItem) & """ />"
		end if
	next
	layout = layout & "</p>"
	
	' Call the build layer function to fill the layer with the group condition content
	layout = buildLayer("layerPeriodY", LABEL_ViewYear, intAsgPeriodY, layout)
	
	' Return the function
	buildLayerPeriodY = layout

end function


'-----------------------------------------------------------------------------------------
' Build the layer to choose the report sorting mode.
'-----------------------------------------------------------------------------------------
public function buildLayerMode()

	Dim i
	Dim layout
	Dim aryMode(2)
	
	aryMode(1) = "month"
	aryMode(2) = "all"

	layout = "<p>"
	
	' Show all months in a select
	layout = layout & vbCrLf & "<select name=""mode"">"
	for i = 1 to Ubound(aryMode)
		layout = layout & "<option value=""" & aryMode(i) & """" 
		if strAsgMode = aryMode(i) then 
			' Check the DTD
			if InStr(STR_ASG_PAGE_DOCTYPE, "XHTML") > 0 then
				layout = layout & " selected=""selected"""
			else
				layout = layout & " selected"
			end if
		end if
		layout = layout & " >" & aryMode(i) & "</option>"
	next
	layout = layout & "</select>"

	layout = layout & vbCrLf & "<input type=""submit"" name=""showsubmit"" value=""" & TXT_button_show & """ />"
	layout = layout & "</p>"
	
	' Call the build layer function to fill the layer with the group condition content
	layout = buildLayer("layerMode", LABEL_ViewMode, strAsgMode, layout)
	
	' Return the function
	buildLayerMode = layout

end function


'-----------------------------------------------------------------------------------------
' Build the layer to choose the report sorting mode.
'-----------------------------------------------------------------------------------------
public function buildLayerGroup(argItem, argItemValue)

	Dim layout
	Dim i
	Dim aryItem
	Dim aryItemValue
	aryItem = Split(argItem, "|")
	aryItemValue = Split(argItemValue, "|")

	layout = "<p>"
	
	' Grouping mode layer
	layout = layout & "<select name=""group"">"
	
	for i = 0 to Ubound(aryItem)
		
		layout = layout & "<option value=""" & aryItemValue(i) & """" 
			if strAsgGroup = aryItemValue(i) then 
				' Check the DTD
				if InStr(STR_ASG_PAGE_DOCTYPE, "XHTML") > 0 then
					layout = layout & " selected=""selected"""
				else
					layout = layout & " selected"
				end if
			end if
		layout = layout & " >" & aryItem(i) & "</option>"

	next
	
	layout = layout & "</select>"
	layout = layout & "</p>"
	
	' Call the build layer function to fill the layer with the group condition content
	layout = buildLayer("layerGroup", LABEL_Group, "", layout)
		
	' Return the function
	buildLayerGroup = layout

end function


'-----------------------------------------------------------------------------------------
' Build a common layer with default style that may be used as a template.
' The layer is filled with the content gived as a function argument and called
' depending on the layer name gived as a function argument.
' If the layer name start with 'x-' the content will be hidden, at the opposite
' by default the content is displayed.
'
' @since 3.0
'-----------------------------------------------------------------------------------------
public function buildLayer(layerID, legendTitle, legendText, databaseField)

	Dim layout

	' layerID argument holds the unique layer id to swap display status.
	' If the layerID start with x- the default display status is hidden
	if Len(layerID) > 0 AND Instr(layerID, "x-") > 0 then
		layout = layout & vbCrLf & "<div id=""" & Replace(layerID, "x-", "") & """ style=""display: none;"">"
	else
		layout = layout & vbCrLf & "<div id=""" & layerID & """ style=""display: block;"">"
	end if

	'layout = layout & vbCrLf & "<div id=""" & layerID & """ style=""display: block;"">"
	layout = layout & vbCrLf & "<fieldset class=""fldlayer""><legend class=""fldlegendtext""><span class=""fldlegendtitle"">" & legendTitle & " :: </span>" & legendText & "</legend>"
	' layout = layout & vbCrLf & "<p>"

	layout = layout & vbCrLf & databaseField
	
	' layout = layout & vbCrLf & "</p>"
	layout = layout & vbCrLf & "</fieldset>"
	layout = layout & "</div>"
	
	' Return the function
	buildLayer = layout

end function


'-----------------------------------------------------------------------------------------
' Build the main layout table and open or close it depending on settings.
' The table is filled with a content layer gived as function argument.
' If the row name start with 'x-' the content will be hidden, at the opposite
' by default the content is displayed.
'
' @since 3.0
'-----------------------------------------------------------------------------------------
public function builTableTlayout(rowID, action, title)

	Dim layout
	
	if action = "open" then
		layout = layout & vbCrLf & "<!-- content table " & title & " -->"
		layout = layout & vbCrLf & "<table class=""tlayout_border"" cellpadding=""5"" cellspacing=""1"" border=""0"" width=""100%"" align=""center"">"
		layout = layout & "<tr>"
		layout = layout & vbCrLf & "<td class=""tlayout_cat"">" & title & "</td>"
		layout = layout & vbCrLf & "</tr>"
		
		' Exception:
		' The system is creating a search for layout and the user has searched a query.
		' - Show an open layer -
		if Len(asgSearchfor) > 0 AND Len(asgSearchin) > 0 and Instr(rowID, "x-rowSearch") > 0 then
			rowID = "rowSearch"
		end if
		
		' rowID argument holds the unique row id to swap display status.
		' If the rowID start with x- the default display status is hidden
		if Len(rowID) > 0 AND Instr(rowID, "x-") > 0 then
			layout = layout & vbCrLf & "<tr id=""" & Replace(rowID, "x-", "") & """ style=""display: none;"">"
			layout = layout & vbCrLf & "<td style=""padding:0px"">"
		elseif Len(rowID) > 0 AND not Instr(rowID, "x-") > 0 then
			layout = layout & vbCrLf & "<tr id=""" & rowID & """ style=""display: table-row;"">"
			layout = layout & vbCrLf & "<td style=""padding:0px"">"
		else
			layout = layout & vbCrLf & "<tr>"
			layout = layout & vbCrLf & "<td style=""padding:0px"">"
		end if
	
	elseif action = "close" then
		layout = layout & vbCrLf & "<div class=""top""><a href=""#top"" title=""" & TXT_pagetop & """>" & TXT_pagetop & " ^</a></div>"
		layout = layout & "</td>"
		layout = layout & "</tr>"
		layout = layout & "</table>"
		layout = layout & vbCrLf & "<!-- / content table " & title & " -->"
	end if
		
	' Return the function
	builTableTlayout = layout

end function
			

%>