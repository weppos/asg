<%

'-------------------------------------------------------------------------------'
'	ASP Stats Generator															'
'	Copyright © 2003-2005 Simone Carletti [ aka weppos ]						'
'-------------------------------------------------------------------------------'


'/**
' * Get the day count for the selected month.
' * 
' * @param		(int) mm 	- month value
' *				1 --> January ... 12 --> December
' * @return 	(int) the number of days in the selected month,
' *				0 if the number index is not valid.
' *
' * @since 		2.x
' *
' * @author		Simone Carletti <carletti@weppos.net>
' */ 
public function getDayCount(mm)
	
	Dim return
	
	if IsNumeric(mm) then
		select case Cint(mm)
			case 1, 3, 5 , 7, 8, 10, 12
				return = 31

			case 4, 6, 9, 11 
				return = 30

			case 2 
				if IsDate("29/02/" & Year(Date())) then
					return = 29
				else
					return = 28
				end if 		

			' Invalid index
			case else
				return = 0
		end select
	else
		return = 0
	end if
	
	getDayCount = return
	
end function

'/**
' * Get the current system date time and update it following
' * time zone settings defined by the user.
' * 
' * @param		(int) datetimevalue	- current date and time
' * @param		(string) timezone	- timezone value. 
' * 			A valid timezone is +H or -H where H is the offset.
' * @return 	(date) the new date and time value. 
' *
' * @since 		2.x
' *
' * @author		Simone Carletti <carletti@weppos.net>
' */ 
public function formatTimeZone(datetimevalue, timezone)
	
	' Check the offset
	if Left(timeZone, 1) = "+" then
		datetimevalue = DateAdd("h", timezone, datetimevalue)
	elseif Left(timeZone, 1) = "-" then
		datetimevalue = DateAdd("h", timezone, datetimevalue)
	end if
	
	' Split
	' dtmAsgNow = datetimevalue
	dtmAsgDay = Day(datetimevalue)
	dtmAsgMonth = Month(datetimevalue)
	dtmAsgYear = Year(datetimevalue)
	if ASG_USE_MYSQL then
	dtmAsgDate = Year(datetimevalue) & "-" & Month(datetimevalue) & "-" & Day(datetimevalue)
	dtmAsgNow = Year(datetimevalue) & "-" & Month(datetimevalue) & "-" & Day(datetimevalue) & " " & Hour(datetimevalue) & ":" & Minute(datetimevalue) & ":" & Second(datetimevalue)
	else
	dtmAsgDate = Year(datetimevalue) & "/" & Month(datetimevalue) & "/" & Day(datetimevalue)
	dtmAsgNow = Year(datetimevalue) & "/" & Month(datetimevalue) & "/" & Day(datetimevalue) & " " & Hour(datetimevalue) & "." & Minute(datetimevalue) & "." & Second(datetimevalue)
	end if
	
	' Add leading 0
	if len(dtmAsgDay) < 2 then dtmAsgDay = "0" & dtmAsgDay
	if len(dtmAsgMonth) < 2 then dtmAsgMonth = "0" & dtmAsgMonth
	
end function

'/**
' * Calculate time from seconds.
' * 
' * @param		
' * @param		
' * @return 	
' *
' * @since 		3.0
' *
' * @author		Simone Carletti <carletti@weppos.net>
' */ 
public function convertSecondToTime(seconds)
	
	Dim aryTime(2)
	Dim intSeconds
	Dim intMinutes
	Dim intHours
	Dim intBuffer
	
	aryTime(0) = 0
	aryTime(1) = 0
	aryTime(2) = 0
	
	if IsNumeric(seconds) then
		
		intHours = Cint(seconds / 3600)
		intBuffer = seconds Mod 3600
		intMinutes = Cint(intBuffer / 60)
		intBuffer = intBuffer Mod 60
		intSeconds = intBuffer
		
		aryTime(0) = intSeconds
		aryTime(1) = intMinutes
		aryTime(2) = intHours
		
	end if
	
	convertSecondToTime = aryTime

end function

'/**
' * Conver date and time format depending on user settings.
' * 
' * @param		
' * @param		
' * @return 	string § converted date/time.
' *
' * @since 		2.0
' * @version	1.01 , 2005-04-12
' *
' * @author		Simone Carletti <carletti@weppos.net>
' */ 
public function formatDateTimeValue(datetimevalue, datetimetype)
	
	Dim return
	
	if IsNull(datetimevalue) OR not Len(datetimevalue) > 0 then datetimevalue = dtmAsgNow
	
	' 
	select case datetimetype
		
		Case "Year"
			return = CInt(Year(datetimevalue))
		Case "Month"
			return = CInt(Month(datetimevalue))
		Case "Day"
			return = CInt(Day(datetimevalue))
		Case "Hour"
			return = CInt(Hour(datetimevalue))
		Case "Minute"
			return = CInt(Minute(datetimevalue))
		Case "Second"
			return = CInt(Second(datetimevalue))
		Case "Time"
			return = CDate(TimeSerial(Hour(datetimevalue), Minute(datetimevalue), Second(datetimevalue)))
		Case "Date"
			return = CDate(DateSerial(Year(datetimevalue), Month(datetimevalue), Day(datetimevalue)))
		
		' Month value for stats report
		Case "MonthReport"
			return = Right("0" & Month(datetimevalue) & "-" & Year(datetimevalue), 7)

	end select
	
	if not datetimetype = "Time" AND not datetimetype = "Date" AND not datetimetype = "MonthReport" then
		if datetimevalue < 10 then datetimevalue = "0" & datetimevalue 
	end if
	
	' Return the formatted date time
	formatDateTimeValue = return
		
end function

%>