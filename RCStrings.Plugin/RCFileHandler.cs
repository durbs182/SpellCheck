using System;
using System.Collections.Generic;
using SpellCheck.Common.XmlData;
using SpellCheck.Common;

using SpellCheck.Common.Utility;
using System.IO;
using System.Diagnostics;
using System.Globalization;

namespace SpellCheck.Plugin.RCStrings
{
  

    [Serializable]
    public class RCFileHandler : StringFileHandler
    {
	    /// <summary>
	    /// The main entry point for the application.
	    /// </summary>
	    [STAThread]
	    static void Main()
	    {
	        RCFileHandler rfh = new RCFileHandler();
	        rfh.LoadStrings(new FileInfo(@"C:\msi\test1.rc"), null, null);
	    }           
	    
    
    	private List<FileExtension> m_supportedExtensions;

        #region IStringResourceFileHandler Members

        public override FileData LoadStrings(FileInfo fileInfo, IList<CheckStringDelegate> checkStringDelegates,  CultureInfo locale)
        {
            List<StringValue> strings = ParseRCFile(fileInfo);

            FileData fileData = new FileData(fileInfo.FullName);

            foreach (StringValue stringVal in strings)
            {
               fileData = HandleSpellingDelegates( checkStringDelegates, fileData,stringVal);
            }

            return fileData;
        }

        private List<StringValue> ParseRCFile(FileInfo fileInfo)
        {
            List<StringValue> strings = new List<StringValue>();
            
            using(StreamReader reader = new StreamReader(fileInfo.OpenRead()))
            {
                    string STRINGTABLE = "STRINGTABLE";
                    string END = "END";
                    string BEGIN = "BEGIN";
                    string DIALOGEX = "DIALOGEX";
                    string DIALOG = "DIALOG";
                    string CAPTION = "CAPTION";

                bool foundStringTable = false;
                bool foundDialogex = false;
                bool foundDialogexBegin = false;

                string line;
                string dialogID = string.Empty;
                int dialogIDStrCount = 0;

                while (( line = reader.ReadLine()) != null)
                {
                    line = line.Trim(); // drop leading whitespace
                    

                    if (line.Equals(STRINGTABLE))
                    {
                        foundStringTable = true;
                        continue;
                    }

                    if (line.IndexOf(DIALOGEX) != -1 || line.IndexOf(DIALOG) != -1)
                    { 
                        foundDialogex = true;
                        int i = line.IndexOf(' ');
                        dialogID = line.Substring(0,i);
                        continue;
                    }

                    try
                    {
                        if (foundDialogex)
                        {
                            if (line.Equals(END))
                            {
                                foundDialogex = false;
                                foundDialogexBegin = false;
                                dialogIDStrCount = 0;
                                continue;
                            }
                            if (line.Equals(BEGIN))
                            {
                                foundDialogexBegin = true;
                                continue;
                            }

                            if (line.StartsWith(CAPTION))
                            { 
                                
                                int i = line.IndexOf('"');
                                string Key = string.Format("{0}_{1}", dialogID, line.Substring(0, i).Trim());

                                string value = line.Substring(i, (line.Length - i));
                                value = value.Replace("\"", "");

                                StringValue sv = new StringValue(Key, value, CultureInfo.InvariantCulture);
                            }

                            if (foundDialogexBegin)
                            {
                                if (line != string.Empty)
                                {                                        
                                    int i = line.IndexOf('"');
                                    if (i > 0)
                                    {
                                        
                                        string key = string.Format("{0}{1}_{2}",dialogID,dialogIDStrCount, line.Substring(0, i).Trim());

                                        dialogIDStrCount++;
                                        string str2 = line.Substring(i, (line.Length - i));
                                        str2 = str2.Substring(1);

                                        i = str2.IndexOf('"');

                                        str2 = str2.Substring(0, i);
                                        str2 = str2.Trim();

                                        if (str2 != string.Empty)
                                        {

                                            str2 = str2.Replace("&", string.Empty);
                                            str2 = str2.Replace(">", string.Empty);
                                            str2 = str2.Replace("<", string.Empty);
                                            str2 = str2.Replace("\\n", " ");
                                            str2 = str2.Replace("\\t", " ");
                                            str2 = str2.Replace("\\", " ");

                                            StringValue sv = new StringValue(key, str2, CultureInfo.InvariantCulture);

                                            strings.Add(sv);
                                        }
                                    }
                                }                                    
                            }

                            continue;
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new Exception(string.Format("Error parsing FILE[{0} @ DIALOGEX: {1}]", fileInfo.FullName, dialogID), ex);
                    }


                    try
                    {

                        if (foundStringTable)
                        {
                            if (line.Equals(END))
                            {
                                foundStringTable = false;
                                continue;
                            }
                            if (line.Equals(BEGIN))
                            {
                                continue;
                            }


                            	
                            if (line != string.Empty)
                            {
                                string key = null;
                                string str2 = null;

                                int i = line.IndexOf('"');
                                if (i == -1)
                                {
                                    key = line;
                                    line = reader.ReadLine();

                                    str2 = line.Trim();
                                }
                                else
                                {
                                    key = line.Substring(0, i).Trim();
                                    str2 = line.Substring(i, (line.Length - i));

                                    while (true)
                                    {
                                        if (str2.EndsWith("\""))
                                        {
                                            str2 = str2.Replace("\"", "").Trim();

                                            break;
                                        }
                                        else
                                        {
                                            line = reader.ReadLine();

                                            if (line == null || line.Trim().Equals(BEGIN) || line.Trim().Equals(END))
                                            {
                                                throw new Exception("Failed to find enf of STRINGTABLE block");
                                            }

                                            str2 += line.Trim();
                                        }
                                    }
                                }

                                // remove control characters that mess up spell checking
                                str2 = str2.Replace("&", string.Empty);
                                str2 = str2.Replace("\\n", " ");
                                str2 = str2.Replace("\\t", " ");
                                str2 = str2.Replace("\\", " ");


                                StringValue sv = new StringValue(key, str2, CultureInfo.InvariantCulture);
                                strings.Add(sv);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new Exception(string.Format("Error parsing FILE[{0} @ LINE: {1}]", fileInfo.FullName,line), ex);
                    }



                }
                reader.Close();
            }
            return strings;
        }



        #endregion

        #region IFileHandler Members

        public override List<FileExtension> SupportedExtensions
        {
            get
            {
                if (m_supportedExtensions == null)
                {
                    m_supportedExtensions = new List<FileExtension>();
                    m_supportedExtensions.Add(new FileExtension("rc"));
                }
                return m_supportedExtensions;
            }            
        }

        public override string Description
        {
            get { return "Parses .rc files and find strings in the stringtable"; }
        }

        #endregion
    }

}