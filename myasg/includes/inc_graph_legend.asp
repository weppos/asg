<%

Response.Write(vbCrLf & "			  <tr class=""smalltext"" align=""center"">")
Response.Write(vbCrLf & "				<td width=""100%"" colspan=""" & intAsgNumCol & """><br />")
Response.Write("<img src=""images/bar_graph_image_visits.gif"" width=""10"" height=""8"" alt=""" & strAsgTxtVisits & """ align=""absmiddle"" />&nbsp;&nbsp;" & strAsgTxtVisits & "&nbsp;&nbsp;")
Response.Write("<img src=""images/bar_graph_image_hits.gif"" width=""10"" height=""8"" alt=""" & strAsgTxtHits & """ align=""absmiddle"" />&nbsp;&nbsp;" & strAsgTxtHits & "")
Response.Write("</td>")
Response.Write(vbCrLf & "			  </tr>")

%>