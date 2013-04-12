using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Text;
using System.Text.RegularExpressions;
using System.IO;
using System.Windows.Forms;
using System.Globalization;

using SpellCheck.Common.XmlData;
using SpellCheck.Common;

using SpellCheck.Common.Utility;

namespace SpellCheck.Plugin.TextFileHandler
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            TextFileHandler tfh = new TextFileHandler();
            tfh.LoadStrings(new FileInfo(@"I:\test\de\Localized.xml"), null, null);
        }
    }



    [Serializable]
    public class TextFileHandler : StringFileHandler
    {
        const string RTF = "rtf";
        const string HTML = "html";
        const string HTM = "htm";
        const string XML = "xml";
        const string TXT = "txt";
        const string XSLT = "xslt";
        const string LOG = "log";
        const string PSM1 = "psm1";
        const string PS1 = "ps1";
        const string PSD1 = "psd1";



        #region helper methods
        private string ReadFile(FileInfo fileInfo)
        {
            StreamReader reader = null;
            try
            {
                reader = fileInfo.OpenText();
                return reader.ReadToEnd();
            }
            finally
            {
                reader.Close();
            }
        }



        #endregion

        #region IStringResourceFileHandler Members

        public override FileData LoadStrings(FileInfo fileInfo, IList<CheckStringDelegate> checkStringDelegates,  CultureInfo fileLocale)
        {
            string text = ReadFile(fileInfo);

            switch (fileInfo.Extension.Substring(1))
            { 
                case RTF:
                    using (RichTextBox rtb = new RichTextBox())
                    {
                        rtb.Rtf = text;
                        text = rtb.Text;
                    }
                    break;
                case HTM:
                case HTML:
                case XML:
                case XSLT:
                    text = Lib.ScrubHTML(text);
                    text = Lib.StripTags(text);
                    break;
                case TXT:
                case LOG:
                case PS1:
                case PSD1:
                case PSM1:
                    break;
                    
            }

            FileData fileData = new FileData(fileInfo.FullName);
            
            // guess the locale from the directory name

            CultureInfo textFileLocale = Lib.GetLocaleFromFilePath(fileInfo);
            
			StringValue sv = new StringValue("FullFileText", text.Trim(), textFileLocale);
			
			fileData = HandleSpellingDelegates(checkStringDelegates, fileData, sv);

            return fileData;        
        }
  
        #endregion

        #region IFileHandler Members

        private List<FileExtension> m_supportedExtensions;


        public override List<FileExtension> SupportedExtensions
        {
            get
            {
                if (m_supportedExtensions == null)
                {
                    m_supportedExtensions = new List<FileExtension>();
                    m_supportedExtensions.Add(new FileExtension(TXT));
                    m_supportedExtensions.Add(new FileExtension(LOG));
                    m_supportedExtensions.Add(new FileExtension(RTF));
                    m_supportedExtensions.Add(new FileExtension(HTML));
                    m_supportedExtensions.Add(new FileExtension(HTM));
                    m_supportedExtensions.Add(new FileExtension(XML));
                    m_supportedExtensions.Add(new FileExtension(XSLT));
                    m_supportedExtensions.Add(new FileExtension(PS1));
                    m_supportedExtensions.Add(new FileExtension(PSD1));
                    m_supportedExtensions.Add(new FileExtension(PSM1));
                }
                return m_supportedExtensions;
            }
        }

        public override string Description
        {
            get { return "Parses LOG, HTML, HTM, TXT, RTF, XML, XLST files"; }
        }

        #endregion
    }
}
