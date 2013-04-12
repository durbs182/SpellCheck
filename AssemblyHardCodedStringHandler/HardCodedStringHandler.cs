using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Globalization;

using System.Reflection;
using SpellCheck.Common;
using SpellCheck.Common.XmlData;


using SpellCheck.Common.Utility;

namespace SpellCheck.Plugin.AssemblyHardCodedHandler
{
    [Serializable]
    public class HardCodedStringFileHandler : StringFileHandler
	{

        public HardCodedStringFileHandler()
        {
            m_stringsDic = new Dictionary<string, List<StringValue>>();
        }
        
        private void Log(string text)
        {
            Lib.ConsoleLog(text);
        }

         #region IFileHandler Members

        List<FileExtension> m_supportedExtensions; 

        public override List<FileExtension> SupportedExtensions
        {
            get
            {
                if (m_supportedExtensions == null)
                {
                    m_supportedExtensions = new List<FileExtension>();
                    m_supportedExtensions.Add(new FileExtension("dll"));
                    m_supportedExtensions.Add(new FileExtension("exe"));
                }
                return m_supportedExtensions;
            }
        }

        public override string Description
        {
            get { return "Spell checks all hard coded strings in an Assembly"; }
        }

        #endregion

        #region IStringResourceFileHandler Members

        Dictionary<string, List<StringValue>> m_stringsDic;

        public override FileData LoadStrings(System.IO.FileInfo fileInfo, IList<CheckStringDelegate> checkStringDelegates,  CultureInfo locale)
        {
	            FileData fileData = null;
                DateTime start = DateTime.Now;

                
                Dictionary<string, List<StringValue>>  dic = AssemblyParser.RunInSandbox(fileInfo);

                if (dic.Count == 0)
                {
                    return null;
                }
                else
                {
                    fileData = new FileData(fileInfo.FullName);
 
                    foreach (string key in dic.Keys)
                    {
                        if (!m_stringsDic.ContainsKey(key))
                        {   
                            m_stringsDic.Add(key, null);

                            List<StringValue> strings = dic[key];
                            if (strings != null)
                            {
                                foreach (StringValue stringVal in strings)
                                {
                                    fileData = HandleSpellingDelegates(checkStringDelegates, fileData, stringVal);
                                    if (stringVal.Value == "True")
                                    { }
                                }
                            }
                        }
                    }

                    return fileData;
                }
        }

        #endregion

    }
}
