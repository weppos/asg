<%

'-------------------------------------------------------------------------------'
'	ASP Stats Generator															'
'	Copyright © 2003-2005 Simone Carletti [ aka weppos ]						'
'-------------------------------------------------------------------------------'


'/**
' * TODO
' * 
' * @param		
' * @return 	(array) the list of all email component
' * 			supported by the program.
' *
' * @since 	3.0
' */ 
public function mail_components()
	
	Dim components
	components = Array("CDOSYS", "CDONTS")
	mail_components = components

end function


'/**
' * Send an email using the mail component
' * specified in the mailComponent argument.
' * 
' * @param		TODO
' * @return 	(boolean) true if the mail has been sent,
' *				false if an error occured.
' *
' * @since 	3.0
' */ 
public function mail(Message, fromEmail, toEmail, fromName, toName, Subject, smtpServer, smtpPort, mailComponent)

	on error resume next
	Dim sent
	sent = false
	
	select case mailComponent
		case "CDONTS"
			sent = mail_CDONTS(Message, fromEmail, toEmail, fromName, toName, Subject, smtpServer, smtpPort)
		case "CDOSYS"
			sent = mail_CDOSYS(Message, fromEmail, toEmail, fromName, toName, Subject, smtpServer, smtpPort)
	end select

	on error goto 0
	mail = sent

end function

'/**
' * TODO
' * 
' * @param		TODO
' * @return 	true if the mail has been sent,
' *				false if an error occured.
' *
' * @since 	3.0
' */ 
public function mail_CDONTS(Message, fromEmail, toEmail, fromName, toName, Subject, smtpServer, smtpPort)

	on error resume next
	Dim objMail
	
	Set objMail = Server.CreateObject("CDONTS.NewMail") 
	if err.number <> 0 then
		mail_CDONTS = false : Set objMail = Nothing : err.clear() : err = 0
		exit function
	end if
	
	objMail.To = toEmail
	objMail.From = fromEmail
	objMail.Subject = Subject
	objMail.Body = Message
	objMail.MailFormat = 1
	objMail.BodyFormat = 1
	objMail.Send
	
	if err.number <> 0 then
		mail_CDONTS = false : Set objMail = Nothing : err.clear() : err = 0
		exit function
	end if
	Set objMail = Nothing
	mail_CDONTS = true

end function

'/**
' * TODO
' * 
' * @param		
' * @return 	(boolean) true if the mail has been sent,
' *				false if an error occured.
' *
' * @since 	3.0
' */ 
function mail_CDOSYS(Message, fromEmail, toEmail, fromName, toName, Subject, smtpServer, smtpPort)

	on error resume next
	Dim objMail
	
	Set objMail = server.createobject("CDO.Message") 
	if err.number <> 0 then
		mail_CDOSYS = false : Set objMail = Nothing : err.clear() : err = 0
		exit function
	end if
	
	objMail.From = fromName & " <" & fromEmail & ">"
	objMail.To = toName & " <" & toEmail & ">"
	objMail.TextBody = Message
	objMail.Subject = Subject
	with objMail.Configuration
		.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = smtpServer
		.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = smtpPort
		.Fields.Update
	end with
	objMail.Send
	
	if err.number <> 0 then
		mail_CDOSYS = false : Set objMail = Nothing : err.clear() : err = 0
		exit function
	end if
	set objMail = Nothing
	mail_CDOSYS = true

end function

%>