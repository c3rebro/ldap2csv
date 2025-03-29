/*
 * Created by SharpDevelop.
 * User: c3rebro
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 * 
 * credits and many thanks goes to Ryan Dunn: http://dunnry.com/blog/AsynchronousLDAPSearchingWithSystemDirectoryServicesProtocols.aspx
 * 
 */
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using System.Threading;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.DirectoryServices.Protocols;
using Microsoft.Win32.SafeHandles;

namespace LDAP2CSV
{
	class Program : IDisposable
	{
		static ManualResetEvent _resetEvent = new ManualResetEvent(false);
		
		static Guid _guid;
		
		static List<string> _arguments;
		
		static int _resultsCounter;
		
		static string[] attr;
		
		public static int Main(string[] args){
			try	{
				Version v = Assembly.GetExecutingAssembly().GetName().Version;
				//fill list of arguments
				_arguments = new List<string>();
				
				foreach(string arg in args)
				{
					_arguments.Add(arg.Replace("\"",string.Empty));
				}
				
				// some fancy window title
				Console.Title = string.Format("LDAP2CSV v{0}.{1}.{2}",v.Major, v.Minor, v.Build);
				
				//want to create a lcv file? exit after job is done
				if(_arguments.Contains("-c")){
					Crypto.Protect(_arguments[_arguments.IndexOf("-c") + 1]);
					return 0;
				}
				
				//messing around with parameters? dont try
				if(_arguments.Contains("-u") && ! _arguments.Contains("-p")){
					throw new ArgumentException("-u and -p parameter must be specified in conjuction");
				}
				
				// check existance of desirend amount of parameters at least: server:port/rootdn
				if(_arguments.Count < 1)	{
					Console.WriteLine(
						"LDAP2CSV usage:\n" +
						"\n" +
						"ldap2csv.exe LDAP://10.0.0.4:389/ou=people,dc=example,dc=de \n -u username -p password -f (&(ou=people)) -a sn,givenName,ou " +
						"\n -o \"D:\\Path\\To\\Destination.csv -e UTF8 -s ;\"\n" +
						"\n" +
						"optional parmeter -a: ldap attributes\n" +
						"supported attributes are: comma seperated list of attributes\n" +
						"Hint: If omitted, attribute will be set to wildcard \"*\"\n" +
						"usage: -a sn,givenName\n" +
						"\n" +
						"optional parmeter -f: ldap search filter\n" +
						"supported attributes are: standard ldap query\n" +
						"Hint: If omitted, filter is set to \"objectClass=*\"\n" +
						"usage: -f ou=people\n" +
						"\n" +
						"optional parmeter -u: username\n" +
						"supported attributes are: ldap username\n" +
						"Hint: refer -c and -C parameter for encrypted credentials\n" +
						"Hint: Anonymous Login is used if -u, -p and -C are omitted\n" +
						"usage: -u cn=admin,dc=example,dc=de\n" +
						"\n" +
						"optional parmeter -p: plain password\n" +
						"supported attributes are: ldap userpassword\n" +
						"Hint: -u parameter is mandatory if -p should be used \n" +
						"Hint: refer -c and -C parameter for encrypted credentials\n" +
						"usage: -p password\n" +
						"\n" +
						"optional parmeter -e: encoding\n" +
						"supported attributes are: UTF8, ASCII(=ANSI) and Unicode\n" +
						"Hint: If omitted, encoding remain system default\n" +
						"usage: -e UTF8\n" +
						"\n" +
						"optional parmeter -s: csv seperator char\n" +
						"supported attributes: maybe all ASCII characters" +
						"Hint: If omitted, comma will be used as seperator\n" +
						"usage: -s ;\n" +
						"\n" +
						"optional parmeter -v: ldap protocol version\n" +
						"supported attributes: 2, 3\n" +
						"Hint: If omitted, version 3 is assumed\n" +
						"usage: -v 2\n" +
						"\n" +
						"optional parmeter -ssl: use ssl encryption\n" +
						"supported attributes: none\n" +
						"Hint: If omitted, basic authentication will be used\n" +
						"usage: -ssl\n" +
						"\n" +
						"optional parmeter -tls: use tls encryption\n" +
						"supported attributes: none\n" +
						"Hint: If omitted, basic authentication will be used\n" +
						"usage: -tls\n" +
						"\n" +
						"optional parmeter -k: use kerberos encryption\n" +
						"supported attributes: none\n" +
						"Hint: If omitted, basic authentication will be used\n" +
						"usage: -k\n" +
						"\n" +
						"optional parmeter -t: timeout in minutes before a connection\n" +
						"or a ldap query will raise a timeout exception\n" +
						"supported attributes: 1 - n\n" +
						"Hint: If omitted, timeout will be set to 10 minutes\n" +
						"usage: -t 5\n" +
						"\n" +
						"optional parameter -c: encryption of credentials\n" +
						"this will create an encrypted \"cred.lcv\" in\n" +
						"the root directory of ldap2csv.exe where the credentials\n" +
						"will be stored safely. May need elevated privileges \"Run As Administrator\"\n" +
						"supported attributes: \"username|password\"\n" +
						"Hint: Be aware to use the | character (pipe) to seperate user and password.\n" +
						"Hint: Be aware to set the attribute in quotes.\n" +
						"Hint: To use the created credential file see \"-C\" parameter.\n" +
						"usage: ldap2csv.exe -c \"username|password\"\n" +
						"\n" +
						"optional paremeter -C: use encrypted credential file\n" +
						"this will tell ldap2csv.exe to search for the \"cred.lcv\" file\n" +
						"created with the -c parameter above.\n" +
						"supported attributes: none\n" +
						"usage: ldap2csv.exe LDAP://10.0.0.4:389/ou=people,dc=example,dc=de -C -f (&(ou=people))\n" +
						"-a sn,givenName,ou -o D:\\Path\\To\\Destination.csv\\\n" +
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
						string dn = _arguments[0].ToUpper().Replace("LDAP://",string.Empty)
							.Replace("LDAPS://",string.Empty)
							.Remove(0,_arguments[0].ToUpper()
							        .Replace("LDAP://",string.Empty).Replace("LDAPS://",string.Empty)
							        .IndexOf('/') +1)
							.ToLower();
						
						string server = _arguments[0].ToUpper().Replace("LDAP://",string.Empty)
							.Replace("LDAPS://",string.Empty)
							.Remove(_arguments[0].ToUpper()
							        .Replace("LDAP://",string.Empty).Replace("LDAPS://",string.Empty)
							        .IndexOf('/'), dn.Length + 1)
							.ToLower();
						
						LdapDirectoryIdentifier ldapDir = new LdapDirectoryIdentifier(server);
						
						//prepare output directory if -o parameter was supplied
						if(_arguments.Contains("-o"))	{
							
							// the argument right after the "-o" argument should be an existing directory as this will be used for the output file. create if not exists
							if(!Directory.GetParent(_arguments[_arguments.IndexOf("-o") +1 ]).Exists)
								Directory.CreateDirectory(Directory.GetParent(_arguments[_arguments.IndexOf("-o") +1 ]).ToString());
							
							// try to remove previous csv file if it exists
							if(File.Exists(_arguments[_arguments.IndexOf("-o") +1 ])) {
								try	{
									File.Delete(_arguments[_arguments.IndexOf("-o") +1 ]);
								}
								catch(Exception ex) {
									LogWriter.CreateLogEntry(string.Format("ERROR:{0} {1}", ex.Message, (ex.InnerException != null) ? ex.InnerException.Message : string.Empty));
								}
							}
						}
						
						using (LdapConnection ldapConnection = CreateConnection(ldapDir))
						{
							//use ssl?
							ldapConnection.SessionOptions.SecureSocketLayer = _arguments.Contains("-ssl") ? true : false;
							
							//use tls?
							if(_arguments.Contains("-tls") && !_arguments.Contains("-ssl"))
								ldapConnection.SessionOptions.StartTransportLayerSecurity(null);
							
							//protocol version specified? use version 3 if not
							ldapConnection.SessionOptions.ProtocolVersion = _arguments.Contains("-v") ? int.Parse(_arguments[_arguments.IndexOf("-v") +1 ]) : 3;
							
							//do not chase referral
							ldapConnection.SessionOptions.ReferralChasing = ReferralChasingOptions.None;

							//You may need to try different types of Authentication depending on the setup
							ldapConnection.AuthType = (_arguments.Contains("-u")||_arguments.Contains("-C")) ? AuthType.Basic : AuthType.Anonymous;
							
							//use kerberos?
							if(_arguments.Contains("-k"))
							{
								ldapConnection.AuthType = AuthType.Kerberos;
								ldapConnection.SessionOptions.Sealing = true;
								ldapConnection.SessionOptions.Signing = true;
							}
							
							//set or get timeout setting
							ldapConnection.Timeout = new TimeSpan(0, _arguments.Contains("-t") ? int.Parse(_arguments[_arguments.IndexOf("-t") +1 ]) : 10, 0);
							
							//use credentials from encrypted credential file?
							if(_arguments.Contains("-C"))	{
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
							
							//use plain authentication?
							else if (_arguments.Contains("-u") && _arguments.Contains("-p")){
								NetworkCredential cred =
									new NetworkCredential(_arguments[_arguments.IndexOf("-u") +1 ],_arguments[_arguments.IndexOf("-p") +1 ]);
								
								ldapConnection.Bind(cred);
							}
							
							//or try anonymously?
							else {
								NetworkCredential cred =
									new NetworkCredential();
								cred.Password = null;
								cred.UserName = null;
								
								ldapConnection.Bind(cred);
							}
							
							AsyncSearcher searcher = CreateSearcher(ldapConnection);

							//this call is asynch, so we need to keep this main
							//thread alive in order to see anything
							//we can use the same searcher for multiple requests - we just have to track which one
							//is which, so we can interpret the results later in our events.
							attr = _arguments.Contains("-a") ? _arguments[_arguments.IndexOf("-a") +1 ].Split(new char[]{',',';'}) : new string[]{"*"};
							
							_guid = searcher.BeginPagedSearch(
								//set searchbase
								string.IsNullOrWhiteSpace(dn) ? "*" : dn,
								//set filter
								_arguments.Contains("-f") ? _arguments[_arguments.IndexOf("-f") +1 ] : "(objectClass=*)",
								//set attributes
								attr,
								150);

							//we will use a reset event to signal when we are done (using Sleep() on
							//current thread would work too...)
							_resetEvent.WaitOne(); //wait for signal;
						}
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
		
		static AsyncSearcher CreateSearcher(LdapConnection connection)
		{
			AsyncSearcher searcher = new AsyncSearcher(connection);

			//assign some handlers for our events
			searcher.PageCompleted += new EventHandler<AsyncEventArgs>(searcher_PageCompleted);
			searcher.SearchCompleted += new EventHandler<AsyncEventArgs>(searcher_SearchCompleted);

			return searcher;
		}

		static LdapConnection CreateConnection(LdapDirectoryIdentifier server)
		{
			LdapConnection connect = new LdapConnection(server);

			return connect;
		}

		static void searcher_SearchCompleted(object sender, AsyncEventArgs e)
		{
			//this is volatile, so we need check it first or another thread
			//could change this from under us
			bool lastSearch = (((AsyncSearcher)sender).PendingSearches == 0);

			Console.WriteLine(
				"{0} Search Complete on thread {1}",
				e.RequestID.Equals(_guid) ? "First" : "Second",
				Thread.CurrentThread.ManagedThreadId
			);

			if (lastSearch)
				_resetEvent.Set();
		}

		static void searcher_PageCompleted(object sender, AsyncEventArgs e)
		{
			try {
				Encoding textEncoding;
				string attrString = "";
				_resultsCounter += e.Results.Count;
				
				//or do something with e.Results here...
				Console.WriteLine(
					"Found {0} results on search thread {1}",
					_resultsCounter,
					Thread.CurrentThread.ManagedThreadId
				);

				// encoding parameter was specified, try to apply
				if(_arguments.Contains("-e"))	{
					switch(_arguments[_arguments.IndexOf("-e") + 1].ToUpper())
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

                foreach (SearchResultEntry searchResult in e.Results)
                {
                    List<string> rowValues = new List<string>();

                    foreach (string a in attr)
                    {
                        string value = "";

                        if (searchResult.Attributes.Contains(a))
                        {
                            DirectoryAttribute subAttrib = searchResult.Attributes[a];
                            List<string> values = new List<string>();

                            foreach (var v in subAttrib)
                            {
                                if (v is byte[] bytes)
                                {
                                    // Convert byte[] to base64 string, or hex, or UTF8 string
                                    // Here we’ll try UTF8, fallback to base64 if it fails
                                    try
                                    {
                                        values.Add(Encoding.UTF8.GetString(bytes));
                                    }
                                    catch
                                    {
                                        values.Add(Convert.ToBase64String(bytes));
                                    }
                                }
                                else
                                {
                                    values.Add(v.ToString());
                                }
                            }

                            value = string.Join(" ", values); // use space between multiple values
                        }

                        rowValues.Add(value);
                    }

                    string separator = _arguments.Contains("-s") ? _arguments[_arguments.IndexOf("-s") + 1] : ",";

                    string csvLine = string.Join(separator, rowValues);

                    if (_arguments.Contains("-o"))
                    {
                        using (FileStream fs = new FileStream(_arguments[_arguments.IndexOf("-o") + 1], FileMode.Append, FileAccess.Write))
                        using (StreamWriter sw = new StreamWriter(fs, textEncoding))
                        {
                            sw.WriteLine(csvLine);
                        }
                    }
                    else
                    {
                        Console.WriteLine(csvLine + "\n");
                    }
                }
            }
			catch (Exception ex) {
				LogWriter.CreateLogEntry(string.Format("ERROR:{0} {1}", ex.Message, (ex.InnerException != null) ? ex.InnerException.Message : string.Empty));
			}
		}
		/// <summary>
		/// Listing 5.10 Modified...
		/// </summary>
		public class AsyncSearcher
		{
			LdapConnection _connect;
			Hashtable _results = new Hashtable();

			public event EventHandler<AsyncEventArgs> SearchCompleted;
			public event EventHandler<AsyncEventArgs> PageCompleted;

			public AsyncSearcher(LdapConnection connection)
			{
				this._connect = connection;
			}

			/// <summary>
			/// Volatile count of outstanding searches in process
			/// </summary>
			public int PendingSearches
			{
				get
				{
					return _results.Count;
				}
			}

			private void InternalCallback(IAsyncResult result)
			{
				SearchResponse response = this._connect.EndSendRequest(result) as SearchResponse;

				ProcessResponse(response, ((SearchRequest)result.AsyncState).RequestId);

				//find the returned page response control
				foreach (DirectoryControl control in response.Controls)
				{
					if (control is PageResultResponseControl)
					{
						//call paged search again
						NextPage((SearchRequest)result.AsyncState, ((PageResultResponseControl)control).Cookie);
						break;
					}
				}
			}

			private void ProcessResponse(SearchResponse response, string guid)
			{
				//only 1 thread at a time gets here...
				List<SearchResultEntry> entries = new List<SearchResultEntry>();

				foreach (SearchResultEntry entry in response.Entries)
				{
					entries.Add(entry);
				}

				//signal our caller that we have a page
				EventHandler<AsyncEventArgs> OnPage = PageCompleted;
				if (OnPage != null)
				{
					OnPage(
						this,
						new AsyncEventArgs(entries, guid)
					);
				}

				//add to the main collection
				((List<SearchResultEntry>)_results[guid]).AddRange(entries);
			}

			public Guid BeginPagedSearch(
				string baseDN,
				string filter,
				string[] attribs,
				int pageSize
			)
			{
				Guid guid = Guid.NewGuid();

				SearchRequest request = new SearchRequest(
					baseDN,
					filter,
					System.DirectoryServices.Protocols.SearchScope.Subtree,
					attribs
				);

				PageResultRequestControl prc = new PageResultRequestControl(pageSize);

				//add the paging control
				request.Controls.Add(prc);

				//we will use this to distinguish multiple searches.
				request.RequestId = guid.ToString();

				//create a temporary placeholder for the results
				_results.Add(request.RequestId, new List<SearchResultEntry>());

				//kick off async
				IAsyncResult result = this._connect.BeginSendRequest(
					request,
					PartialResultProcessing.NoPartialResultSupport,
					new AsyncCallback(InternalCallback),
					request
				);

				return guid;
			}

			private void NextPage(SearchRequest request, byte[] cookie)
			{
				//our last page is when the cookie is empty
				if (cookie != null && cookie.Length != 0)
				{
					//update the cookie and preserve page size
					foreach (DirectoryControl control in request.Controls)
					{
						if (control is PageResultRequestControl)
						{
							((PageResultRequestControl)control).Cookie = cookie;
							break;
						}
					}

					//call it again to get next page
					IAsyncResult result = this._connect.BeginSendRequest(
						request,
						PartialResultProcessing.NoPartialResultSupport,
						new AsyncCallback(InternalCallback),
						request
					);
				}
				else
				{
					List<SearchResultEntry> results = (List<SearchResultEntry>)_results[request.RequestId];

					//decrement our collection when we are done
					_results.Remove(request.RequestId);

					//we have finished, signal the caller
					EventHandler<AsyncEventArgs> OnComplete = SearchCompleted;
					if (OnComplete != null)
					{
						OnComplete(
							this,
							new AsyncEventArgs(results, request.RequestId)
						);
					}
				}
			}
		}

		/// <summary>
		/// Just a simple class to hold some results
		/// </summary>
		public class AsyncEventArgs : EventArgs
		{
			Guid _id;
			List<SearchResultEntry> _entries;

			public AsyncEventArgs(List<SearchResultEntry> entries, string requestID)
			{
				_entries = entries;
				_id = new Guid(requestID);
			}

			public List<SearchResultEntry> Results
			{
				get
				{
					return _entries;
				}
			}

			public Guid RequestID
			{
				get
				{
					return _id;
				}
			}
		}
		
		// Flag: Has Dispose already been called?
		bool disposed = false;
		// Instantiate a SafeHandle instance.
		SafeHandle handle = new SafeFileHandle(IntPtr.Zero, true);
		
		// Public implementation of Dispose pattern callable by consumers.
		public void Dispose()
		{
			Dispose(true);
			GC.SuppressFinalize(this);
		}
		
		// Protected implementation of Dispose pattern.
		protected virtual void Dispose(bool disposing)
		{
			if (disposed)
				return;
			
			if (disposing) {
				handle.Dispose();
				// Free any other managed objects here.
			}
			
			disposed = true;
		}
	}
}