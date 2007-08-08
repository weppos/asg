<%

' General transparent .gif
Const ASG_ICON_EMPTY = "47494638396110001000910000000000FFFFFFFFFFFF00000021F90401000002002C000000001000100000020E948FA9CBED0FA39CB4DA8BB33E05003B"

'/** 
' * Convert the encoded icon to the gif content type and print it.
' * 
' * @param		(string) icon		- encoded icon
' * 
' * @author		Simone Carletti <carletti@weppos.net>
' */ 
public function printIcon(icon)

	Dim Index
	
	Response.ContentType = "Image/gif"
	for Index = 1 to Len(icon) step 2
		Response.BinaryWrite(ChrB("&h" & Mid(icon,Index,2)))
	next
	
end function

%>