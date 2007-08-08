<%

'-------------------------------------------------------------------------------'
'	ASP Stats Generator															'
'	Copyright © 2003-2005 Simone Carletti [ aka weppos ]						'
'-------------------------------------------------------------------------------'


' ***** DATABASE TYPE *****

' Set True to use MySQL
Const ASG_USE_MYSQL = true			

' Set True to use Microsoft Access
Const ASG_USE_ACCESS = false		


' ***** ACCESS DATABASE SETTINGS *****
'

' Database folder
Const ASG_ACCESS_PATH = "database/" 
' Database name
Const ASG_ACCESS_DATABASE = "dbstats" 


' ***** MYSQL DATABASE SERVER NAME *****
'

' Server IP or name
Const ASG_MYSQL_SERVER 		= "localhost"
' Database name
Const ASG_MYSQL_DATABASE 	= "weppos_asg3"
' Database user
Const ASG_MYSQL_USER 		= "wep-weppos"
' Database password for the selected user
Const ASG_MYSQL_PASSWORD 	= "SC1985sql"


' ***** IP2COUNTRY DATABASE SETTINGS *****
'

' Database folder
Const ASG_IP2C_PATH = "database/" 
' Database name
Const ASG_IP2C_DATABASE = "ip-to-country" 
' Database table
Const ASG_IP2C_TABLE = "ip2country" 

' MySQL users can choose to store ip2country table
' in the same database where data is stored
' or to create a new dedicated database.
' Set true to use the same database.
Const ASG_IP2C_SAMEDATABASE = true

' If ASG_IP2C_SAMEDATABASE = false you must
' fill the following variables with MySQL configuration.
Const ASG_IP2C_SERVER 		= ""
Const ASG_IP2C_USER 		= ""
Const ASG_IP2C_PASSWORD 	= ""

%>
