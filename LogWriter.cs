﻿/*
 * Created by SharpDevelop.
 * Date: 31.08.2017
 * Time: 20:32
 *
 */

using System;
using System.IO;
using System.Reflection;

namespace LDAP2CSV
{
    /// <summary>
    /// Description of LogWriter.
    /// </summary>
    public static class LogWriter
    {
        private static StreamWriter textStream;

        private static readonly string _logFileName = "err.log";

        /// <summary>
        ///
        /// </summary>
        /// <param name="entry"></param>
        public static void CreateLogEntry(string entry)
        {
        	string _logFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), Assembly.GetExecutingAssembly().GetName().Name, "log");

            if (!Directory.Exists(_logFilePath))
                Directory.CreateDirectory(_logFilePath);
            try
            {
                if (!File.Exists(Path.Combine(_logFilePath, _logFileName)))
                    textStream = File.CreateText(Path.Combine(_logFilePath, _logFileName));
                else
                    textStream = File.AppendText(Path.Combine(_logFilePath, _logFileName));
                textStream.WriteAsync(string.Format("{0}" + Environment.NewLine, entry.Replace("\r\n", "; ").Insert(0,DateTime.Now.ToString() + " ")));
                textStream.Close();
                textStream.Dispose();
            }
            catch
            {
            }
        }
    }
}