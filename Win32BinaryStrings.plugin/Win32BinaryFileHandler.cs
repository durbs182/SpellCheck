using System;
using System.Collections.Generic;
using System.Text;

using SpellCheck.Common.XmlData;
using SpellCheck.Common;

using SpellCheck.Common.Utility;



namespace SpellCheck.Plugin.Win32BinaryStrings
{
    [Serializable]
    public class Win32BinaryFileHandler : StringFileHandler
    {
        private List<FileExtension> m_supportedExtensions;

        #region IStringResourceFileHandler Members

        public override FileData LoadStrings(System.IO.FileInfo fileInfo, IList<CheckStringDelegate> checkStringDelegates, System.Globalization.CultureInfo locale)
        {
            return Library.LoadFile(this, fileInfo,checkStringDelegates);
        }

     

        #endregion

        #region IFileHandler Members
        /// <summary>
        /// handles Win32 DLLs and EXEs poss .RES as well
        /// </summary>
        public override List<FileExtension> SupportedExtensions
        {
            get
            {
                if (m_supportedExtensions == null)
                {
                    m_supportedExtensions = new List<FileExtension>();
                    m_supportedExtensions.Add(new FileExtension("exe"));
                    m_supportedExtensions.Add(new FileExtension("dll"));
                }
                return m_supportedExtensions;
            }
        }

        public override string Description
        {
            get { return "Parses a Win32 binary and extracts StringTable and Dialog strings"; }
        }

        #endregion
    }


}
