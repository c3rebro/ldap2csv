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

optional encryption of credentials:
will create an encrypted \"cred.lcv\" in the root directory of ldap2csv.exe
where the credentials will be stored safely. Needs elevated privileges "Run As"
usage: ldap2csv.exe -c "username|password"

to use the stored credentials replace -u and -p parameters by -C parameter
without any additional attributes

example:
ldap2csv.exe LDAP://10.0.0.4:389/ou=people,dc=example,dc=de -C -f (&(ou=people)) -a sn,givenName,ou 
-o \"D:\\Path\\To\\Destination.csv\

### Important: 
for security reasons cred.lcv file is bound to the environment where it was created as well as ldap2csv.exe build and could not be transferred between different machines or releases
            
### Hint:
+ errorlogs are created automatically to C:\User\appdata\local\ldap2csv\log\err.log
