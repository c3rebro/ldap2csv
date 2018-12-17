# ldap2csv
### export an ldap query result to a csv file on windows operating systems

LDAP2CSV usage:
  
ldap2csv.exe LDAP://10.0.0.4:389/ou=people,dc=example,dc=de -u username -p password 
-f (&(ou=people)) -a sn,givenName,ou -o \"D:\\Path\\To\\Destination.csv"  
  
### parameters  
  
optional parmeter -e: encoding  
supported attributes are: UTF8, ASCII(=ANSI) and Unicode.  
**Hint:** If omitted, encoding remain system default  
usage: -e UTF8  
  
optional parmeter -s: csv seperator char  
supported attributes: maybe all ASCII characters  
**Hint:** If omitted, comma will be used as seperator  
usage: -s ;  

optional parmeter -p: ldap protocol version  
supported attributes: 2, 3  
**Hint:** If omitted, version 3 is assumed  
usage: -p 2 ; 

optional parmeter -ssl: use ssl encryption  
supported attributes: none  
**Hint:** If omitted, basic authentication will be used  
usage: -ssl ; 

optional parmeter -t: timeout in minutes before a connection\nor a ldap query will raise a timeout exception
supported attributes: 1 - n  
**Hint:** If omitted, timeout will be set to 10 minutes  
usage: -t 5; 

optional parameter -c: encryption of credentials  
this will create an encrypted \"cred.lcv\" in the root directory of ldap2csv.exe  
where the credentials will be stored safely. May need elevated privileges "Run As Administrator"  
supported attributes: "username|password"  
**Hint:** Be aware to use the | character (pipe) to seperate user and password.  
**Hint:** Be aware to set the attribute in quotes.  
**Hint:** To use the created credential file see "-C" parameter.  
usage: ldap2csv.exe -c "username|password"  
  
optional paremeter -C: use encrypted credential file  
this will tell ldap2csv.exe to search for the "cred.lcv" file   
created with the -c parameter above.  
supported attributes: none  
usage: ldap2csv.exe LDAP://10.0.0.4:389/ou=people,dc=example,dc=de -C -f (&(ou=people))  
-a sn,givenName,ou -o \"D:\\Path\\To\\Destination.csv\  
  
### Important: 
for security reasons cred.lcv file is bound to the environment where it was created as well as ldap2csv.exe build and could not be transferred between different machines or releases
            
### Additional Hints:
+ errorlogs are created automatically to C:\User\%CurrentUser%\appdata\local\ldap2csv\log\err.log
