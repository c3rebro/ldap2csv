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

namespace LDAP2CSV
{
	class Program
	{
		public static int Main(string[] args)
		{
			try
			{
				Version v = Assembly.GetExecutingAssembly().GetName().Version;
				//fill list of arguments
				List<string> arguments = new List<string>(args);
				string attrString = "";
				Encoding textEncoding;
				
				// some fancy window title
				Console.Title = string.Format("LDAP2CSV v{0}.{1}.{2}",v.Major, v.Minor, v.Build);
				
				
				if(args.Contains("-c"))
				{
					Crypto.Protect(args[arguments.IndexOf("-c") + 1]);
					return 0;
				}
				
				// check existance of desirend amount of parameters at least: server:port/rootdn and attributes -a
				if(args.Length < 2)
				{
					Console.WriteLine(
						"LDAP2CSV usage:\n" +
						"\n" +
						"ldap2csv.exe LDAP://10.0.0.4:389/ou=people,dc=example,dc=de \n -u username -p password -f (&(ou=people)) -a sn,givenName,ou " +
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
				
				else
				{
					try
					{
						DirectoryEntry ldapConnection;
						
						//use credentials from encrypted credential file?
						if(args.Contains("-C"))
						{
							try{
								//split back to user|pwd
								string[] credentials = Crypto.Unprotect().Split('|');
								ldapConnection = new DirectoryEntry(args[0],credentials[0],credentials[1]);
							}
							catch(Exception e)
							{
								LogWriter.CreateLogEntry(string.Format("ERROR:{0} {1}", e.Message, (e.InnerException != null) ? e.InnerException.Message : string.Empty));
								return 1;
							}
						}
						
						else {
							//first parameter should be ldap path, usr and pwd are optional
							ldapConnection = new DirectoryEntry(args[0],args[arguments.IndexOf("-u") +1 ],args[arguments.IndexOf("-p") +1 ]);
						}
						
						//basic authentication required? no encryption supportet yet
						ldapConnection.AuthenticationType = args.Contains("-u") ? AuthenticationTypes.ServerBind : AuthenticationTypes.Anonymous;
						
						//connect to ldap
						ldapConnection.RefreshCache();

						DirectorySearcher ldapSearch = new DirectorySearcher(ldapConnection);
						
						//select rootDSE from ldap path
						ldapSearch.SearchRoot = ldapConnection;
						//currently only search throuout all subtrees is supportet
						ldapSearch.SearchScope = SearchScope.Subtree;
						//specify a filter if needed
						ldapSearch.Filter = args.Contains("-f") ? args[arguments.IndexOf("-f") +1 ] : "(objectClass=*)" ;
						//use a seperate thread for the search
						ldapSearch.Asynchronous = true;
						
						// create an array of properties that we would like and
						// add them to the search object
						string[] requiredAttributes = args[arguments.IndexOf("-a") +1 ].Split(',');
						

						foreach(SearchResult searchResult in ldapSearch.FindAll())
						{
							foreach(string attr in requiredAttributes)
							{
								try {
									//only try to read an attrubute if it actually exists in the search scope
									if(searchResult.Properties.Contains(attr))
										//and add it to the csv line followed by a seperator
										attrString += searchResult.Properties[attr][0].ToString() + (args.Contains("-s") ? args[arguments.IndexOf("-s") +1] : ",");
								} catch {
									//nothing found?, dont worry -> go on
									continue;
								}
								
							}
							
							try {
								//remove last seperator
								attrString = attrString.Remove(attrString.Length - 1, 1);
							} catch {
								//nothing found?, dont worry -> go on
								continue;
							}
							
							//and last but not least add a crlf to every line
							attrString += "\n";
						}

						
					}
					catch (Exception e)
					{
						LogWriter.CreateLogEntry(string.Format("ERROR:{0} {1}", e.Message, (e.InnerException != null) ? e.InnerException.Message : string.Empty));
						return 1;
					}
				}
				
				// encoding parameter was specified, try to apply
				if(arguments.IndexOf("-e") > 0)
				{
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
							throw new ArgumentException("Unkown Encoding: Expected one of the following Encoding types: UTF8, ANSI or Unicode");
					}
				}
				
				// if not specified use system default
				else {
					textEncoding = Encoding.Default;
				}
				
				//output to console instead of csv?
				if(args.Contains("-o"))
				{
					// try to remove previous csv file if it exists
					if(File.Exists(args[arguments.IndexOf("-o") +1 ]))
					{
						try
						{
							File.Delete(args[arguments.IndexOf("-o") +1 ]);
						}
						catch(Exception e)
						{
							LogWriter.CreateLogEntry(string.Format("ERROR:{0} {1}", e.Message, (e.InnerException != null) ? e.InnerException.Message : string.Empty));
							return 1;
						}
					}
					
					// the argument right after the "-o" argument should be an existing directory as this will be used for the output file. create if not exists
					if(!Directory.GetParent(args[arguments.IndexOf("-o") +1 ]).Exists)
						Directory.CreateDirectory(Directory.GetParent(args[arguments.IndexOf("-o") +1 ]).ToString());
					
					File.AppendAllLines(args[arguments.IndexOf("-o") +1 ], attrString.Split('\n'),textEncoding);
				}
				
				//prefer console output?
				else
				{
					foreach (string s in attrString.Split('\n'))
						Console.WriteLine(s);
					
					Console.Write(
						"\n" +
						"Press any key to exit . . . ");
					Console.ReadKey(true);
				}
				
				return 0;
			}

			catch (Exception e)
			{
				LogWriter.CreateLogEntry(string.Format("ERROR:{0} {1}", e.Message, (e.InnerException != null) ? e.InnerException.Message : string.Empty));
				return 1;
			}
		}
		
	}
}