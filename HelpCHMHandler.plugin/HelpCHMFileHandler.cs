using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Text;
using System.IO;
using System.Web;
using System.Globalization;

using SpellCheck.Common.XmlData;
using SpellCheck.Common;

using SpellCheck.Common.Utility;

using HtmlHelp.Storage;

namespace SpellCheck.Plugin.HelpCHMHandler
{
    [Serializable]
    public class HelpCHMFileHandler : StringFileHandler
    {


        private List<FileExtension> m_supportedExtensions;

        #region IStringResourceFileHandler Members

        public override FileData LoadStrings(System.IO.FileInfo fileInfo, IList<CheckStringDelegate> checkStringDelegates, CultureInfo fileLocale)
        {
            FileData fileData = new FileData(fileInfo.FullName);

            // guess the locale from the directory name
           
            CultureInfo chmFileLocale = Lib.GetLocaleFromFilePath(fileInfo);


            // Create Instance of ITStorageWrapper.
            // During initialization constructor will process CHM file 
            // and create collection of file objects stored inside CHM file.
            ITStorageWrapper iw = new ITStorageWrapper(fileInfo.FullName, true);


            // Loop through collection of objects stored inside IStorage
            foreach (FileObject fileObject in iw.foCollection)
            {
                // Check to make sure we can READ stream of an individual file object
                if (fileObject.CanRead)
                {
                    string fileString1 = fileObject.ReadFromFile();
                    // We only want to extract HTM files in this example
                    // fileObject is our representation of internal file stored in IStorage
                    if (fileObject.FileName.EndsWith(".htm") || fileObject.FileName.EndsWith(".html"))
                    {
              
                        Lib.Log("File: " + fileObject.FileName);

                        string fileString = fileObject.ReadFromFile();

                        // Read first and then save later example
                        //StreamWriter sw = File.CreateText(@"c:\test1\" + fileObject.FileName);
                        //sw.WriteLine(fileString);
                        //sw.Close();

                        

                        chmFileLocale = Lib.GetHtmlLocale(fileString);

                        fileString = Lib.StripTags(fileString);

                        // Need to tokenize a long string as word spell checking doesn't seem to like a long string
                        StringValue sv = new StringValue(fileInfo.Name, fileString.Trim(), chmFileLocale);
						fileData = HandleSpellingDelegates(checkStringDelegates ,fileData, sv);



                    }
                }
            }

            return fileData;
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
                    m_supportedExtensions.Add(new FileExtension("chm"));
                }
                return m_supportedExtensions;
            }
        }

        public override string Description
        {
            get { return "Parses compiled HTML help file *.chm"; }
        }

        #endregion
    }
}
