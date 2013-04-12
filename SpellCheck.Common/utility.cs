using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Text;
using System.Text.RegularExpressions;
using System.Management;
using System.Linq;
using System.Globalization;
using System.Web;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

using System.Diagnostics;

using HtmlAgilityPack;


namespace SpellCheck.Common.Utility
{
    public class Locale
    {
        public const string JA = "ja";
        public const string EN = "en";
        public const string DE = "de";
        public const string FR = "fr";
        public const string ES = "es";

        public const string ENUS = "en-US";
    }

    /// <summary>
    /// Add text file listener to Trace subsystem
    /// </summary>
    public class TraceLog
    {
        static TextWriterTraceListener myWriter;
        static string c_path;
        static bool c_deleteOldLog;
        
    
        
        public TraceLog()
        {
        	c_path = Assembly.GetCallingAssembly().GetName().Name + ".log";
			c_deleteOldLog = true;
			
			
			CreateTraceLog();
        }
        
        public TraceLog(string exePath, bool deleteOldLog)
        {
        	c_path = exePath + ".log";;

			c_deleteOldLog = deleteOldLog;
			
			if(c_deleteOldLog)
        	{
	
	            if (File.Exists(c_path))
	            {
	                File.Delete(c_path);
	            }
        	}

			 CreateTraceLog();
        }
        
        private void CreateTraceLog()
        {
        	myWriter = new TextWriterTraceListener(c_path);
            Trace.Listeners.Add(myWriter);
            Lib.Log("================================================start new log==============================================================");
            Lib.Log("Trace.Listeners count={0} file={1}", Trace.Listeners.Count, c_path);
        }
    }
    

		

    public class Lib
    {
        const string SPACE= " ";
        
        public const int EXCELTEXTMAXLENGTH = 32767;
        
        static DirectoryInfo c_ExtractionDirectory;
        
        static bool c_logToConsole;
        
        static string c_reportname;

        private static Regex _splitCamelCaseRegex;
        private static Regex _matchCamelCaseRegex;

        static Lib()
        {
            _splitCamelCaseRegex = new Regex("([A-Z a-z][a-z]+)", RegexOptions.Compiled);
            _matchCamelCaseRegex = new Regex(@"(\b[a-z]|\B[A-Z])", RegexOptions.Compiled);
        }
        
		public static bool LogToConsole {
			get { return c_logToConsole; }
			set { c_logToConsole = value; }
		}
        
        
    
        /// <summary>
        /// Log
        /// </summary>
        /// <param name="text"></param>
        /// <param name="args"></param>
        public static void Log(string text, params object[] args)
        {
            Log(string.Format(text, args));
        }

        public static void Log(object obj)
        {
            Log(obj.ToString());
        }

        public static void Log(string text)
        {
            Trace.WriteLine(string.Format("[{0}] {1}", DateTime.Now, text));
            Trace.Flush();
            
            if(LogToConsole)
            {
                if (_consoleOutAction == null)
                {
                    Console.WriteLine(text);
                }
                else
                {
                    _consoleOutAction(text);
                }
            }
        }
        
        public static void DebugLog(string text, params object[] args)
        {
        	text = string.Format("**Debug**> {0} {1} <**Debug**",GetMethodName(2) ,text);
        	Log(string.Format(text,args));
        }
        
        public static void ConsoleLog(string text, params object[] args)
        {
        	Log(string.Format(text,args));
        	// Console.WriteLine(string.Format(text,args));
        }

        static Action<string> _consoleOutAction;

        public static Action<string> ConsoleOutAction
        {
            get 
            {
                return _consoleOutAction;
            }
            set
            {
                _consoleOutAction = value;
            }
        }

        /// <summary>
        /// returns name of calling method
        /// </summary>
        /// <returns></returns> 
        public static string GetMethodName()
        {
        	return GetMethodName(2);
        }
        
                /// <summary>
        /// returns name of calling method
        /// use 1 if calling from the method you want to log
        /// use 2 if logging from here to get the calling method
        /// </summary>
        /// <returns></returns> 
        private static string GetMethodName(int frame)
        {
            StackTrace stackTrace = new StackTrace();
            StackFrame stackFrame = stackTrace.GetFrame(frame);

            MethodBase methodBase = stackFrame.GetMethod();
            return string.Format("{0}.{1}()", methodBase.DeclaringType, methodBase.Name);
        }

        /// <summary>
        /// Urlencode a string
        /// </summary>
        /// <param name="inStr"></param>
        /// <returns></returns>
        public static string UrlEncode(string inStr)
        {
            string outStr = null;

            if (inStr != null)
            {
               return HttpUtility.UrlEncode(inStr);
            }

            return outStr;
        }

        /// <summary>
        /// Remove all <TAGS> from HTML and XML files, decode HTML encoding and strips white space
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string StripTags(string str)
        {
            System.Text.RegularExpressions.Regex objRegEx1 = new System.Text.RegularExpressions.Regex("<!--[*]-->");
            System.Text.RegularExpressions.MatchCollection col = objRegEx1.Matches(str);
            string text = objRegEx1.Replace(str, "");
            
            //need to remove script first.
            System.Text.RegularExpressions.Regex objRegEx2 = new System.Text.RegularExpressions.Regex("<script>[^<]*</script>");
            //System.Text.RegularExpressions.MatchCollection col = objRegEx2.Matches(str);
            text = objRegEx2.Replace(text, "");
            //<script>*</script>

            // Removes tags from passed HTML            
            System.Text.RegularExpressions.Regex objRegEx = new System.Text.RegularExpressions.Regex("<[^>]*>");

            text = objRegEx.Replace(str, "");

            text = HttpUtility.HtmlDecode(text);

            text = text.Replace("&apos;", "'");

            text = text.Replace("\n\r", string.Empty);
            text = text.Replace(";", SPACE);
            text = text.Replace(":", SPACE);

            text = text.Trim();

            return text;
        }
        
        //public static string TruncatePath( string path, int length )
        //{
        //    StringBuilder sb = new StringBuilder();
        //    PathCompactPathEx( sb, path, length, 0 );
        //    return sb.ToString();
        //}


        /// <summary>
        /// 
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static string SplitCamelCase(string input)
        {
            //string output = System.Text.RegularExpressions.Regex.Replace(
            //input,
            //"([A-Z]+)",
            //" $1",
            //System.Text.RegularExpressions.RegexOptions.Compiled).Trim();

            //([A-Z])([A-Z][a-z]+)
            
            var output = _splitCamelCaseRegex.Replace(input, " $1 ").Trim();

            return output;
        }

    
        
        /// <summary>
        /// Is the string CamelCase or camelCase 
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static bool IsCamelCase(string str)
        {
            // if this is all lowercase skip
            if (!str.Equals(str.ToLower(), StringComparison.Ordinal) && !str.Equals(str.ToUpper(), StringComparison.Ordinal))
            {
                //same
                //MatchCollection m = Regex.Matches(str, @"(\b[a-z]|\B[A-Z])", RegexOptions.Compiled);
                var m = _matchCamelCaseRegex.Matches(str);
                if (m.Count > 0)
                {
                    //is cammelCase
                    return true;
                }
            }
            return false;
        }
        
        /// <summary>
        /// If the word has non alpha charaters ignore it strips out $Name, \\path\path\ or //file/file etc
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static bool WordContainsOnlyAlpha(string str)
        {
            if (Regex.Match(str, @"([^a-zA-Z])", RegexOptions.Compiled).Success)
            {
                if (str.Equals("."))
                {
                    return false;
                }

                if (!str.EndsWith(".") && str.IndexOf("'") == -1)
                {
                    return false;
                }
            }
            return true;
        }

        /// <summary>
        /// Split a large string and return String collection 
        /// </summary>
        /// <param name="?"></param>
        /// <returns></returns>
        public static StringCollection TokenizeString(string text)
        {
            StringCollection strings = new StringCollection();
 
            string[] arr = text.Split(new char[] { '\r', '\n' ,' '}, StringSplitOptions.RemoveEmptyEntries);
            strings.AddRange(arr);

            return strings;
        }

        /// <summary>
        /// removes script, applet and object nodes from html files
        /// </summary>
        /// <param name="html"></param>
        /// <returns></returns>
        public static string ScrubHTML(string html)
        {
            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(html);

            //Remove potentially harmful elements
            HtmlNodeCollection nc = doc.DocumentNode.SelectNodes("//script|//applet|//object");
            if (nc != null)
            {
                foreach (HtmlNode node in nc)
                {
                    node.ParentNode.RemoveChild(node, false);

                }
            }


            return doc.DocumentNode.WriteTo();
        } 

        /// <summary>
        /// Get the CultureInfo from the directory name 
        /// </summary>
        /// <param name="fileInfo"></param>
        /// <returns>null if JA</returns>
        public static CultureInfo GetLocaleFromFilePath(System.IO.FileInfo fileInfo)
        {
            string dirName = fileInfo.Directory.Name;

            CultureInfo locale = GetLocaleFromTWOLetterString(dirName);

            if (locale == null)
            {
                locale = GetLocaleFromFileName(fileInfo);
            }   

            return locale;
        }

        public static CultureInfo GetLocaleFromTWOLetterString(string twoLetterStr)
        {
            CultureInfo locale = null;
            switch (twoLetterStr)
            {
                case Locale.DE:
                case Locale.FR:
                case Locale.ES:
                case Locale.JA:
                    locale = new CultureInfo(twoLetterStr);
                    break;
                case Locale.EN:
                    locale = new CultureInfo(Locale.ENUS);
                    break;
            }
            return locale;
        }

        public static CultureInfo GetLocaleFromFileName(FileInfo fileInfo)
        {
            string[] arr = fileInfo.Name.Split(new char[] { '.'});
            if (arr.Length == 3)
            {
                string locStr = arr[1];
                try
                {
                    CultureInfo locale = GetLocaleFromTWOLetterString(locStr);
                    return locale;
                }
                catch (Exception ex)
                {
                    Lib.Log("Can't get locale from {0}: {1}", fileInfo.FullName,ex.Message);
                }
            }

            return null;
        }
        

        /// <summary>
        /// Get a unique report name using current date and time
        /// </summary>
        /// <returns></returns>
        public static string ReportName
        {
			get
			{
				if(c_reportname == null)
				{
		        	c_reportname = string.Format("{0}_{1}","spellingreport", DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture));
				}
				return c_reportname;
			}
        }

 

        /// <summary>
        /// Run the gived method by refelection
        /// </summary>
        /// <param name="assemblyName"></param>
        /// <param name="typeName"></param>
        /// <param name="GetMethodName"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public static object RunMethodByReflection(string assemblyName, string typeName, string GetMethodName, object[] parameters)
        {
            object obj = AppDomain.CurrentDomain.CreateInstanceAndUnwrap(assemblyName, typeName);

            MethodInfo minfo = obj.GetType().GetMethod(GetMethodName);

            object retObj = minfo.Invoke(obj, parameters);

            return retObj;
        }
        
        /// <summary>
        /// Creates an object in a sandboxed AppDomain which can then be disposed of
        /// </summary>
        /// <param name="sandbox"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public static T CreateObjectInstance<T>(out AppDomain sandbox)
        {
    		Type type = typeof(T);
    		AppDomainSetup setup = new AppDomainSetup();
            setup.ApplicationBase = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            setup.PrivateBinPath = AppDomain.CurrentDomain.SetupInformation.PrivateBinPath;
            setup.ApplicationName = null;
            setup.ShadowCopyFiles = "false";
            
            string sandBoxName = string.Format("SandBox{0}",Guid.NewGuid().ToString());
            
            sandbox = AppDomain.CreateDomain(sandBoxName, null, setup);
            
            string adstr = AppDomain.CurrentDomain.FriendlyName;
            
            T obj = (T) sandbox.CreateInstanceAndUnwrap(type.Assembly.GetName().Name, type.FullName); 
    		
            adstr = AppDomain.CurrentDomain.FriendlyName;
            
            return obj;
        }

        
        static DirectoryInfo c_WorkingDirectory;
        
        /// <summary>
        /// Returns the directory for the executing assembly
        /// </summary>
        public static DirectoryInfo WorkingDirectory
        {
            get
            {
                if (c_WorkingDirectory == null)
                {
                    FileInfo fi = new FileInfo(Assembly.GetExecutingAssembly().Location);
                    c_WorkingDirectory = new DirectoryInfo(fi.DirectoryName);
                }

                return c_WorkingDirectory;
            }
        }
        
        /// <summary>
        /// Return the path for the result xml/xls file
        /// </summary>
        /// <param name="htmlString"></param>
        /// <returns></returns>
        public static DirectoryInfo ReportPath
        {
        	get
        	{
        		string REPORTDIR = "reports";
        	    DirectoryInfo resultPath = new DirectoryInfo(Path.Combine(Lib.WorkingDirectory.FullName, REPORTDIR));

                if (!resultPath.Exists)
                {
                    resultPath.Create();
                }
                
                return resultPath;
        	}
        }
        

        /// <summary>
        /// Try to get the locale from the HTML lang value
        /// </summary>
        /// <param name="htmlString"></param>
        /// <returns></returns>
        public static CultureInfo GetHtmlLocale(string htmlString)
        {
            //htmlString = @"<html lang=""en-us"" xml:lang=""en-us""/>";

            CultureInfo ci = CultureInfo.InvariantCulture;

            HtmlAgilityPack.HtmlDocument doc = new HtmlDocument();
            try
            {
                doc.LoadHtml(htmlString);
                HtmlAttribute attribute= doc.DocumentNode.SelectSingleNode("//html").Attributes["lang"];
                if(attribute != null)
                {
	                ci = new CultureInfo(attribute.Value);
	                return ci;
                }
            }
            catch (ArgumentException ex)
            {
                Log(ex);
            }


            // if loading html fails try getting the locale from the text
            string groupStr = "extract";

            string pattern = string.Format(@"<html lang=[^""]*""(?'{0}'.*?)""",groupStr);

            Match match = Regex.Match(htmlString, pattern);

            if (match.Success)
            {
                string lang = match.Groups[groupStr].Value;

                ci = new CultureInfo(lang);

                return ci;
            }

            return ci;
        }

        /// <summary>
        /// The current working directory
        /// </summary>
        public static DirectoryInfo ExtractionDirectory
        {
            get
            {
                if (c_ExtractionDirectory == null)
                {
                    //FileInfo fi = new FileInfo(Assembly.GetExecutingAssembly().Location);
                    string rootDrive = new DirectoryInfo(WorkingDirectory.FullName).Root.FullName;

                    c_ExtractionDirectory = new DirectoryInfo(Path.Combine(rootDrive,"extract"));
                    c_ExtractionDirectory.Create();
                }

                return c_ExtractionDirectory;
            }
            
            set
            {
            	c_ExtractionDirectory = value;
            }
        }
        
        
        /// <summary>
        /// Get the TwoLetter lang name from the dictionary file name. null if not language specfic.
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
//        private static string GetDictionarytLangFromFileName(FileInfo file)
//        {
//            string groupStr = "extract";
//            string pattern = string.Format("[^_]*_(?'{1}'.*?){2}", c_CitrixCustomDicName, groupStr, c_CitrixCustomDicExtension);
//            Match match = Regex.Match(file.Name, pattern);
//
//            if (match.Success)
//            {
//                string lang = match.Groups[groupStr].Value;
//                return lang;            
//            } 
//
//            return null;
//        }
//        
                /// <summary>
        /// Get the processor type for this machine and check if it is 64 bit
        /// </summary>
        /// <returns></returns>
        public static bool IsRunningOnWin64()
        {
            ManagementClass class1 = new ManagementClass("Win32_Processor");

            foreach (ManagementObject ob in class1.GetInstances())
            {
                PropertyData data = ob.Properties["AddressWidth"];
                UInt16 AddressWidth = (UInt16)data.Value;

                return AddressWidth == 64;
            } 

            throw new ApplicationException("Failed to get processor type");
        }
        
        /// <summary>
        /// Check to see if we can write to the disk we are running on.
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        public static bool IsValidDrive(FileInfo file)
        {
            DriveInfo dinfo = new DriveInfo(file.Directory.Root.Name);

           // Lib.Log("Drive: {0} is {1}", file.FullName.Substring(0, 1),dinfo);

            switch (dinfo.DriveType)
            {
                case DriveType.Unknown:
                case DriveType.Network:
                case DriveType.CDRom:
        		{
            		Lib.Log("Drive: {0} is {1} and is readonly", file.FullName.Substring(0, 1),dinfo);
                    return false;
        		}
                    
            }

            return true;
        }


    }
}
