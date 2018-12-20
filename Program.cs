/*
 * Created by SharpDevelop.
 * User: c3rebro
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using System.Net;
using System.Reflection;
using System.DirectoryServices;
using System.DirectoryServices.Protocols;

namespace LDAP2CSV
{
	class Program
	{
		public static int Main(string[] args){
			try	{
				Version v = Assembly.GetExecutingAssembly().GetName().Version;
				//fill list of arguments
				List<string> arguments = new List<string>(args);
				string attrString = "";
				Encoding textEncoding;
				
				// some fancy window title
				Console.Title = string.Format("LDAP2CSV v{0}.{1}.{2}",v.Major, v.Minor, v.Build);
				
				
				if(args.Contains("-c"))	{
					Crypto.Protect(args[arguments.IndexOf("-c") + 1]);
					return 0;
				}
				
				// check existance of desirend amount of parameters at least: server:port/rootdn and attributes -a
				if(args.Length < 2)	{
					Console.WriteLine(
						"LDAP2CSV usage:\n" +
						"\n" +
						"ldap2csv.exe LDAP://10.0.0.4:389 -dn ou=people,dc=example,dc=de \n -u username -p password -f (&(ou=people)) -a sn,givenName,ou " +
						"\n -o \"D:\\Path\\To\\Destination.csv -e UTF8 -s ;\"\n" +
						"\n" +
						"optional parmeter: encoding, supported encodings are:\n" +
						"UTF8, ASCII(=ANSI) and Unicode. If omitted, encoding remain system default\n" +
						"usage: -e UTF8\n" +
						"\n" +
						"optional parmeter: seperator character\n" +
						"if omitted, comma will be used as seperator\n" +
						"usage: -s ;\n" +
						"\n" +
						"optional encryption of credentials:\n" +
						"will create an encrypted \"cred.lcv\" in the root directory of ldap2csv.exe\n" +
						"where the credentials will be stored safely. Needs elevated privileges \"RunAs\"\n" +
						"usage: ldap2csv.exe -c \"username|password\"\n" +
						"\n" +
						"to use the stored credentials replace -u and -p parameters by -C parameter\n" +
						"without any additional attributes\n" +
						"example:\n" +
						"ldap2csv.exe LDAP://10.0.0.4:389/ou=people,dc=example,dc=de -C -f (&(ou=people)) -a sn,givenName,ou " +
						"-o \"D:\\Path\\To\\Destination.csv\"\n" +
						"\nImportant: cred.lcv file is bound to the environment where it was created and could not be transferred between different machines"
					);
					
					Console.Write(
						"\n" +
						"Press any key to exit . . . "
					);
					
					Console.ReadKey(true);
					
					return 0;
				}
				
				else {
					try	{
						//extract server and dn from ldap path attribute
						string dn = args[0].Replace("LDAP://",string.Empty).Remove(0,args[0].Replace("LDAP://",string.Empty).IndexOf('/') +1);
						string server = args[0].Replace("LDAP://",string.Empty).Remove(args[0].Replace("LDAP://",string.Empty).IndexOf('/'), dn.Length + 1);
						
						LdapDirectoryIdentifier ldapDir = new LdapDirectoryIdentifier(server);
						
						LdapConnection ldapConnection = new LdapConnection(ldapDir);
						//protocol version specified? use version 3 if not
						ldapConnection.SessionOptions.ProtocolVersion = args.Contains("-v") ? int.Parse(args[arguments.IndexOf("-v") +1 ]) : 3;
						//use ssl?
						ldapConnection.SessionOptions.SecureSocketLayer = args.Contains("-ssl") ? true : false;
						//You may need to try different types of Authentication depending on your setup
						ldapConnection.AuthType = (args.Contains("-u")||args.Contains("-C")) ? AuthType.Basic : AuthType.Anonymous;
						
						//use credentials from encrypted credential file?
						if(args.Contains("-C"))	{
							try{
								//split back to user|pwd
								string[] credentials = Crypto.Unprotect().Split('|');

								NetworkCredential cred =
									new NetworkCredential(credentials[0], credentials[1]);
								
								//ldap connect
								ldapConnection.Bind(cred);
							}
							catch(Exception e) {
								LogWriter.CreateLogEntry(string.Format("ERROR:{0} {1}", e.Message, (e.InnerException != null) ? e.InnerException.Message : string.Empty));
								ldapConnection.Dispose();
								return 1;
							}
						}
						
						else if (args.Contains("-u")){
							NetworkCredential cred =
								new NetworkCredential(args[arguments.IndexOf("-u") +1 ],args[arguments.IndexOf("-p") +1 ]);
							
							ldapConnection.Bind(cred);
						}
						
						else {
							NetworkCredential cred =
								new NetworkCredential();
							cred.Password = null;
							cred.UserName = null;
							
							ldapConnection.Bind(cred);
						}
						
						SearchRequest searchRequest = new SearchRequest();
						//set the query timeout
						searchRequest.TimeLimit = new TimeSpan(0, args.Contains("-t") ? int.Parse(args[arguments.IndexOf("-t") +1 ]) : 10, 0);
						ldapConnection.Timeout = new TimeSpan(0, args.Contains("-t") ? int.Parse(args[arguments.IndexOf("-t") +1 ]) : 10, 0);
						//set searchbase
						searchRequest.DistinguishedName = string.IsNullOrWhiteSpace(dn) ? "*" : dn;
						//set filter
						searchRequest.Filter = args.Contains("-f") ? args[arguments.IndexOf("-f") +1 ] : "(objectClass=*)" ;
						//set scope
						searchRequest.Scope = System.DirectoryServices.Protocols.SearchScope.Subtree;
						//run search
						SearchResponse response = (SearchResponse)ldapConnection.SendRequest(searchRequest);
						//and fill collection with results
						SearchResultEntryCollection entries =  response.Entries;

						// create an array of properties that we would like and
						// add them to the search object
						string[] requiredAttributes = args[arguments.IndexOf("-a") +1 ].Split(',');
						
						//prepare output directory of -o parameter was supplied
						if(args.Contains("-o"))	{
							// try to remove previous csv file if it exists
							if(File.Exists(args[arguments.IndexOf("-o") +1 ])) {
								try	{
									File.Delete(args[arguments.IndexOf("-o") +1 ]);
								}
								catch(Exception e) {
									LogWriter.CreateLogEntry(string.Format("ERROR:{0} {1}", e.Message, (e.InnerException != null) ? e.InnerException.Message : string.Empty));
									ldapConnection.Dispose();
									return 1;
								}
							}
							
							// the argument right after the "-o" argument should be an existing directory as this will be used for the output file. create if not exists
							if(!Directory.GetParent(args[arguments.IndexOf("-o") +1 ]).Exists)
								Directory.CreateDirectory(Directory.GetParent(args[arguments.IndexOf("-o") +1 ]).ToString());
						}
						
						// encoding parameter was specified, try to apply
						if(args.Contains("-e"))	{
							switch(args[arguments.IndexOf("-e") + 1].ToUpper())
							{
								case "UTF8":
									textEncoding = Encoding.UTF8;
									break;
									
								case "ANSI":
								case "ASCII":
									textEncoding = Encoding.ASCII;
									break;
									
								case "UNICODE":
									textEncoding = Encoding.Unicode;
									break;
									
								default:
									ldapConnection.Dispose();
									throw new ArgumentException("Unkown Encoding: Expected one of the following Encoding types: UTF8, ANSI or Unicode");
							}
						}
						
						// if not specified use system default
						else {
							textEncoding = Encoding.Default;
						}
						
						foreach(SearchResultEntry searchResult in entries) {
							
							IDictionaryEnumerator attribEnum = searchResult.Attributes.GetEnumerator();
							while (attribEnum.MoveNext())//Iterate through the result attributes
							{
								//Attribute Name below
								if(requiredAttributes.Contains(attribEnum.Key.ToString()) || (requiredAttributes.Count() == 1 && requiredAttributes.Contains("*")))
								{
									//Attributes have one or more values so we iterate through all the values of each attribute
									DirectoryAttribute subAttrib = (DirectoryAttribute)attribEnum.Value;
									
									for (int i = 0; i < subAttrib.Count; i++)
									{
										//Add Attributes to string and seperate with a space character if more that one value is present
										attrString += subAttrib[i].ToString() + ((subAttrib.Count > 1) ? " " : string.Empty);
									}
									
									//separate each attribute with comma or supplied separation character
									attrString += args.Contains("-s") ? args[arguments.IndexOf("-s") +1 ] : ",";
								}
							}
							
							try {
								//remove last seperator
								if(!string.IsNullOrEmpty(attrString)){
									attrString = attrString.Remove(attrString.Length - 1, 1);
									
									//output to console instead of csv?
									if(args.Contains("-o")){
										

										using (FileStream fs = new FileStream(args[arguments.IndexOf("-o") + 1], FileMode.Append, FileAccess.Write)) {
											using (StreamWriter sw = new StreamWriter(fs, textEncoding)) {
												
												sw.WriteLine(attrString);
												
												attrString = string.Empty;
												
												sw.Close();
											}
										}
									}
									
									//prefer console output?
									else {
										Console.WriteLine(attrString + "\n");
									}
								}
								
								
							} catch (Exception e) {
								LogWriter.CreateLogEntry(string.Format("ERROR:{0} {1}", e.Message, (e.InnerException != null) ? e.InnerException.Message : string.Empty));
								//nothing found?, dont worry -> go on
								continue;
							}
						}
						ldapConnection.Dispose();
					}
					catch (Exception e)	{
						LogWriter.CreateLogEntry(string.Format("ERROR:{0} {1}", e.Message, (e.InnerException != null) ? e.InnerException.Message : string.Empty));
						return 1;
					}
				}
				return 0;
			}

			catch (Exception e)	{
				LogWriter.CreateLogEntry(string.Format("ERROR:{0} {1}", e.Message, (e.InnerException != null) ? e.InnerException.Message : string.Empty));
				return 1;
			}
		}
	}
}