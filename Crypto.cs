/*
 * Created by SharpDevelop.
 * User: C3rebro
 * Date: 12.12.2018
 * Time: 22:43
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.IO;
using System.Text;
using System.Reflection;
using System.Security.Cryptography;

namespace LDAP2CSV
{
	/// <summary>
	/// Description of Crypto.
	/// </summary>
	public static class Crypto
	{
		// Create byte array for additional entropy when using Protect method.
		static byte [] s_aditionalEntropy = Encoding.Default.GetBytes(
			string.Format("{0},{1},{2}",
			              Assembly.GetExecutingAssembly().GetName().Version.Major,
			              Assembly.GetExecutingAssembly().GetName().Version.Minor,
			              Assembly.GetExecutingAssembly().GetName().Version.Build));
		
		static string pathToCredFile = Path.Combine(Environment.CurrentDirectory, "cred.lcv");
		
		public static int Protect( string cred )
		{
			try
			{
				//pipe character is used as seperator username|pwd
				if(cred.Contains("|"))
				{
					if(File.Exists(pathToCredFile))
						File.Delete(pathToCredFile);
					
					// Encrypt the data using DataProtectionScope.LocalMachine. The result can be decrypted
					// only on the same machine.
					File.AppendAllText(pathToCredFile,
					                   Encoding.Default.GetString(
					                   	ProtectedData.Protect(Encoding.Default.GetBytes(cred), s_aditionalEntropy, DataProtectionScope.LocalMachine )
					                   ), Encoding.Default);

					return 0;
				} else
					return 1;

			}
			
			catch(Exception e)
			{
				LogWriter.CreateLogEntry(string.Format("ERROR:{0} {1}", e.Message, (e.InnerException != null) ? e.InnerException.Message : string.Empty));
				return 1;
			}
		}

		public static string Unprotect()
		{
			try
			{
				//Decrypt the data
				return Encoding.Default.GetString(
					ProtectedData.Unprotect(File.ReadAllBytes(pathToCredFile), s_aditionalEntropy, DataProtectionScope.LocalMachine )
				);
			}
			catch (Exception e)
			{
				LogWriter.CreateLogEntry(string.Format("ERROR:{0} {1}", e.Message, (e.InnerException != null) ? e.InnerException.Message : string.Empty));
				return null;
			}
		}
		
		private static byte[] CreateRandomEntropy()
		{
			// Create a byte array to hold the random value.
			byte[] entropy = new byte[16];

			// Create a new instance of the RNGCryptoServiceProvider.
			// Fill the array with a random value.
			new RNGCryptoServiceProvider().GetBytes(entropy);

			// Return the array.
			return entropy;
		}
	}
}


