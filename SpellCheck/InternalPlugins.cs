using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Globalization;

using System.Collections.Specialized;
using System.Diagnostics;
using System.Windows.Forms;


using SpellCheck.Common;

using SpellCheck.Common.Utility;
using SpellCheck.Common.XmlData;


using Microsoft.Tools.WindowsInstallerXml.Cab;




namespace SpellCheck.InternalPlugins
{


    [Serializable]
    class AssemblyPlugin : InternalStringFileHandlerPlugin
    {
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
            get { return "Parses embedded resource in .Net assemblies"; }
        }

        #endregion

        #region IPlugin Members

        public override FileData LoadStrings(System.IO.FileInfo file, IList<CheckStringDelegate> checkStringDelegates, CultureInfo locale)
        {
            if (file.Exists)
            {
                return AssemblyLoader.LoadFileData(file.FullName, this, checkStringDelegates, locale);
            }
            else
            {
                throw new System.IO.FileNotFoundException();
            }
        }



        #endregion
    }


    class ResXPlugin : InternalStringFileHandlerPlugin
    {
       #region IFileHandler Members

        List<FileExtension> m_supportedExtensions; 

        public override List<FileExtension> SupportedExtensions
        {
            get
            {
                if (m_supportedExtensions == null)
                {
                    m_supportedExtensions = new List<FileExtension>();
                    m_supportedExtensions.Add(new FileExtension("resx"));
                }
                return m_supportedExtensions;
            }
        }

        public override string Description
        {
            get { return "Parses .Net ResX files"; }
        }

        #endregion


        #region IPlugin Members

        public override FileData LoadStrings(System.IO.FileInfo file, IList<CheckStringDelegate> checkStringDelegates, CultureInfo locale)
        {
            FileData fileData = null;
            if (file.Exists)
            {
                Lib.Log(file.FullName);

                CultureInfo culture = Lib.GetLocaleFromFilePath(file);

                // if a locale is set only check the string if the file locale matches
                if (locale != null && !locale.Equals(culture))
                {
                    return fileData;
                }

                System.Resources.ResXResourceReader resxReader = new System.Resources.ResXResourceReader(file.FullName);
                System.Collections.IDictionaryEnumerator idicEnum = resxReader.GetEnumerator();
                fileData = new FileData(file.FullName);
           
                while (idicEnum.MoveNext())
                {
                    if (idicEnum.Value is string && !ExcludedKey(idicEnum.Key.ToString()))
                    {
                        string str = ParseString(idicEnum.Value as string);

                        if (str.Length != 0)
                        {
                            StringValue stringVal = new StringValue(idicEnum.Key.ToString(), str, culture);

                            fileData = HandleSpellingDelegates(checkStringDelegates, fileData, stringVal);
                        }
                  
                    }
                }
                
                resxReader.Close();
            }
            else
            {
                throw new System.IO.FileNotFoundException();
            }

            return fileData;
        }

        #endregion

    }







    class MSIExtractPlugin : ArchiveFileHandler
    {

        public override DirectoryInfo Extract(FileInfo msiFile)
        {
            DirectoryInfo extractDirectory = base.Extract(msiFile);
            ExtractMSI.Wixtracts.ExtractFiles(msiFile, extractDirectory);

            return extractDirectory;
        }

        #region IFileHandler Members

        
        private List<FileExtension> m_SupportedExtensions;

        public override List<FileExtension> SupportedExtensions
        {
            get
            {
                if (m_SupportedExtensions == null)
                {
                    m_SupportedExtensions = new List<FileExtension>();
                    m_SupportedExtensions.Add(new FileExtension("msi"));
                }
                return m_SupportedExtensions;
            }
        }

        public override string Description { get { return "This handler extracts all files from an MSI installer"; } }


        #endregion
    
    }

    class CabExtractPlugin : ArchiveFileHandler
    {

        public override DirectoryInfo Extract(FileInfo cabFile)
        {
            DirectoryInfo extractDirectory = base.Extract(cabFile);

            WixExtractCab extCab = new WixExtractCab();

            extractDirectory.Create();
            // track the created folder so we can return it in the list.
            extCab.Extract(cabFile.FullName, extractDirectory.FullName);
            extCab.Dispose();

            return extractDirectory;
        }

        #region IFileHandler Members

        private List<FileExtension> m_SupportedExtensions;

        public override List<FileExtension> SupportedExtensions
        {
            get
            {
                if (m_SupportedExtensions == null)
                {
                    m_SupportedExtensions = new List<FileExtension>();
                    m_SupportedExtensions.Add(new FileExtension("cab"));
                }
                return m_SupportedExtensions;
            }
        }

        public override string Description { get { return "This handler extracts all files from a cab file"; } }

        #endregion

    }

}
