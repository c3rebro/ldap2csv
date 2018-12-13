# ldap2csv
### export ldap query result to csv file

LDAP2CSV usage:

ldap2csv.exe LDAP://10.0.0.4:389/ou=people,dc=example,dc=de -u username -p password 
-f (&(ou=people)) -a sn,givenName,ou -o \"D:\\Path\\To\\Destination.csv

optional parmeter: encoding, supported encodings are: UTF8, ASCII(=ANSI) and Unicode

if omitted, encoding remain system default
usage: -e UTF8

optional parmeter: seperator char

if omitted, comma will be used as seperator
usage: -s ;

### Hint:
+ errorlogs are created automatically to C:\User\appdata\local\ldap2csv\log\err.log
