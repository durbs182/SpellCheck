using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Xml;
using System.Xml.Serialization;
using System.Windows.Forms;
using System.Reflection;
using System.Linq;

using System.Runtime.InteropServices;

// using Microsoft.Tools.WindowsInstallerXml.Msi;
using ExtractMSI;

using SpellCheck.Common.XmlData;
using SpellCheck.Common;

using SpellCheck.Common.Utility;
using Microsoft.Deployment.WindowsInstaller;


namespace SpellCheck.Plugin.MSIStrings
{


    [Serializable]
    public class MSIFileHandler : StringFileHandler
    {
        static void Main(string[] args)
        {
            MSIFileHandler msiStrings = new MSIFileHandler();

            List<string> cols = msiStrings.GetColumnNames();

            Columns columns = new Columns();
            columns.Names = new string[cols.Count];

            int count = 0;
            foreach (string col in cols)
            {
                columns.Names[count] = col;
                count++;                  
            }

            StreamWriter writer = new StreamWriter(@"C:\columns.xml");

            try
            {
                XmlSerializer s = new XmlSerializer(typeof(Columns));
                s.Serialize(writer, columns);
            }
            finally
            {
                writer.Close();
            }
        }

        /// <summary>
        /// Default column names if columns.xml is not found
        /// </summary>
        /// <returns></returns>
        private List<string> GetColumnNames()
        {
            List<string> columnNames = new List<string>();
            columnNames.Add("Description");
            columnNames.Add("Value");
            columnNames.Add("DisplayName");
            columnNames.Add("Text");
            columnNames.Add("Message");
            columnNames.Add("Title");
            columnNames.Add("DiskPrompt");
            return columnNames;
        }

        private List<CultureInfo> GetTableCultures(Database msidb)
        {
            List<CultureInfo> tableCultures = new List<CultureInfo>();
            //tableCultures.Add("fr");
            //tableCultures.Add("de");
            //tableCultures.Add("ja");
            //tableCultures.Add("ko");
            //tableCultures.Add("ru");
            //tableCultures.Add("es");
            //tableCultures.Add("zh_CN");
            //tableCultures.Add("zh_TW");

            string query = string.Format("SELECT Locale FROM {0}",m_CTXUILocale);

            var list = msidb.ExecuteQuery(query, new Record(0));

            foreach(var locale in list)
            {
                string productLang = locale.ToString();

                CultureInfo culture = new CultureInfo(productLang);

                tableCultures.Add(culture);
            }

            //using (var view = msidb.OpenView(query))
            //{
            //    var record = view.Fetch();
            //    var fieldCount = record.;

            //    foreach (object[] values in record.FieldCount GetString()
            //    {
            //        string productLang = values[0].ToString();

            //        CultureInfo culture = new CultureInfo(productLang);

            //        tableCultures.Add(culture);
            //    }

            //}

            return tableCultures;
            
        }

        /// <summary>
        /// Reads the columns names that are used to find columns in any table that contains strings
        /// </summary>
        /// <returns></returns>
        private List<string> LoadColumns()
        {
            FileInfo fi = new FileInfo(Assembly.GetExecutingAssembly().Location);
            string columnFile = fi.DirectoryName + @"\columns.xml";

            if(!File.Exists(columnFile))
            {
                return GetColumnNames();
            }

            StreamReader reader = new StreamReader(columnFile);
            try
            {
                XmlSerializer serializer = new XmlSerializer(typeof(Columns));
                Columns columns = (Columns)serializer.Deserialize(reader);
                List<string> columnList = new List<string>(columns.Names.Length);
                columnList.AddRange(columns.Names);
                return columnList;
            }
            finally
            {
                reader.Close();
            }
        }

        /// <summary>
        /// Use ProductLanguage Property to get the language the msi is localised in
        /// </summary>
        /// <param name="msidb"></param>
        /// <returns></returns>
        private CultureInfo GetMSICulture(Database msidb)
        {
            string query = "SELECT Value FROM Property WHERE Property='ProductLanguage'";

            var list = msidb.ExecuteQuery(query, new Record(0));

            foreach (var value in list)
            {
                string productLang = value.ToString();

                int lcid = int.Parse(productLang);

                if (lcid > 0)
                {
                    CultureInfo culture = new CultureInfo(lcid);

                    return culture;
                }
                else
                {
                    Lib.Log("ProductLanguage in MSI Property table is an invalid Culture code: ProductLanguage={0}", lcid);
                }
            }

            return CultureInfo.InvariantCulture;
        }

        /// <summary>
        /// Gets the names of all the tables in the MSI DB
        /// </summary>
        /// <param name="msidb"></param>
        /// <returns></returns>
        private List<string> GetMSITableNames(Database msidb)
        {
            string tableName = "_Tables";
            string query = string.Concat("SELECT `Name` FROM `", tableName, "`");

            var list = msidb.ExecuteQuery(query, new Record(0));

            return list.Cast<string>().ToList();
        }

        #region IStringResourceFileHandler Members

        const string MSI = "msi";
        const string MST = "mst";

        public override FileData LoadStrings(FileInfo fileInfo, IList<CheckStringDelegate> checkStringDelegates,  CultureInfo locale)
        {
            string extension = fileInfo.Extension.Substring(1);

            switch (extension)
            { 
                case MST:
                    return ParseMst(fileInfo, checkStringDelegates);
                case MSI:
                    return ParseMsi(fileInfo, checkStringDelegates);
            }

            return null;
            
        }

        /// <summary>
        /// Parse MSI transform file. Load the correct MSI for the transform them apply the transform
        /// </summary>
        /// <param name="fileInfo"></param>
        /// <param name="checkStringDelegate"></param>
        /// <param name="wordApp"></param>
        /// <returns></returns>
        private FileData ParseMst(FileInfo fileInfo, IList<CheckStringDelegate> checkStringDelegates)
        {
            FileInfo msiFileInfo = FindMsi(fileInfo);

            if (msiFileInfo != null)
            {
                WindowsInstaller.Installer installer = CreateWindowsInstaller();

                WindowsInstaller.Database database = installer.OpenDatabase(msiFileInfo.FullName, WindowsInstaller.MsiOpenDatabaseMode.msiOpenDatabaseModeReadOnly);

                // guess the local from the directory name
                CultureInfo locale = Lib.GetLocaleFromFilePath(fileInfo);

                database.ApplyTransform(fileInfo.FullName, WindowsInstaller.MsiTransformError.msiTransformErrorViewTransform);

                List<StringValue> strings = ListTransform(database, locale);

                FileData fileData = new FileData(fileInfo.FullName);

                foreach (StringValue sv in strings)
                {
                    fileData = HandleSpellingDelegates( checkStringDelegates, fileData, sv);
                }

                return fileData;
            }

            Lib.Log("Can not find matching MSI for " + fileInfo.FullName);

            return null;
        }

        const uint icdLong       = 0;
        const uint icdShort      = 0x400;
        const uint icdObject     = 0x800;
        const uint icdString     = 0xC00;
        const uint icdNullable   = 0x1000;
        const uint icdPrimaryKey = 0x2000;
        const uint icdNoNulls    = 0x0000;
        const uint icdPersistent = 0x0100;
        const uint icdTemporary  = 0x0000;

        private List<CultureInfo> m_tableCultures;
        private string m_CTXUILocale = "CTXUILocale";

        private static string DecodeColDef(int colDef)
        {
            string def = string.Empty;
	        switch(colDef & (icdShort | icdObject))
            {
	            case icdLong:
		            def = "LONG";
                    break;
	            case icdShort:
		            def = "SHORT";
                    break;
	            case icdObject:
		            def  ="OBJECT";
                    break;
	            case icdString:
                    def = "CHAR(" + (colDef & 255) + ")";
                    break;
            }
            if ((colDef & icdNullable)   == 0 )
            {
                def = def + " NOT NULL";
            }
            if ((colDef & icdPrimaryKey) != 0 )
            {
                def = def + " PRIMARY KEY";
            }

            return def;
        }

        /// <summary>
        /// Walk the _TransformView table and list all the changed strings
        /// </summary>
        /// <param name="database"></param>
        /// <param name="locale"></param>
        /// <returns></returns>
        private static List<StringValue> ListTransform(WindowsInstaller.Database database,CultureInfo locale)
        {
            List<StringValue> strings = new List<StringValue>();

            var view = database.OpenView("SELECT * FROM `_TransformView` ORDER BY `Table`, `Row`");

            view.Execute(null);
            var record = view.Fetch();

            while (record != null)
            {
                string row = string.Empty;
                string change = string.Empty;
                string column = string.Empty;

                //not sure what the first part does here
                if (record.get_IsNull(3))
                {
                    row = "<DDL>";
                    if (!record.get_IsNull(4))
                    {
                        change = "[" + record.get_StringData(5) + "]: " + DecodeColDef(int.Parse(record.get_StringData(4)));
                    }
                }
                else
                {
                    row = record.get_StringData(3);

                    if (record.get_StringData(2) != "INSERT" && record.get_StringData(2) != "DELETE")
                    {
                        change = record.get_StringData(4);
                    }
                }

                column = record.get_StringData(1) + "." + record.get_StringData(2);

                StringValue sv = new StringValue(string.Format("{0}.{1}", column, row), change, locale);
                strings.Add(sv);

                record = view.Fetch();
            }

            return strings;
        }

        /// <summary>
        ///  create an instance of the Windows Installer. Can't use Wix for this as the Wix Database object
        /// doesn't have ApplyTransform() method
        /// </summary>
        /// <returns></returns>
        private static WindowsInstaller.Installer CreateWindowsInstaller()
        {
            WindowsInstaller.Installer i = null;
            Type oType = Type.GetTypeFromProgID("WindowsInstaller.Installer");
            if (oType != null)
            {
                i = (WindowsInstaller.Installer)Activator.CreateInstance(oType);
            }

            return i;
        }

        /// <summary>
        /// Walk back down the directory tree and look for the matching MSI
        /// only works if the MSI has the same name
        /// </summary>
        /// <param name="fileInfo"></param>
        /// <returns></returns>
        private FileInfo FindMsi(FileInfo fileInfo)
        {
            string directory = fileInfo.DirectoryName;
            string fileName = fileInfo.Name;

            fileName = fileName.Substring(0, fileName.IndexOf(fileInfo.Extension));

            fileName = string.Format("{0}.{1}",fileName,MSI); 

            while (true)
            {
                DirectoryInfo dirInfo = new DirectoryInfo(directory);
                foreach (FileInfo filePathStr in dirInfo.GetFiles(fileName,SearchOption.AllDirectories))
                {
                    return filePathStr;
                }

                directory = directory.Substring(0, directory.LastIndexOf(@"\"));

                if (directory.Length <= 3)
                {
                    break;
                }
            }

            return null;
        }

        private CultureInfo IsMuiTable(string tableName)
        {           

            foreach (CultureInfo tableCulture in m_tableCultures)
            {
                string suffix = string.Format("_{0}", tableCulture.Name);
                suffix = suffix.Replace("-", "_");
                if (tableName.EndsWith(suffix, true, CultureInfo.InvariantCulture))
                {
                    return tableCulture;
                }
            }

            return null;
        }

        /// <summary>
        /// Parse an MSI file for strings
        /// </summary>
        /// <param name="fileInfo"></param>
        /// <param name="checkStringDelegate"></param>
        /// <param name="wordApp"></param>
        /// <returns></returns>
        private FileData ParseMsi(FileInfo fileInfo, IList<CheckStringDelegate> checkStringDelegates)
        {
            FileData fileData = new FileData(fileInfo.FullName);

            using(Database msidb = new Database(fileInfo.FullName, DatabaseOpenMode.ReadOnly))
            {

                List<string> tableNames = GetMSITableNames(msidb);

                bool isMiuMSI = false;

                // check to see if the MSI is multi langauge
                if (tableNames.Contains(m_CTXUILocale))
                {
                    // extract table cultures from MSI
                    m_tableCultures = GetTableCultures(msidb);
                    isMiuMSI = true;
                }

                List<string> columnNames = LoadColumns();

                CultureInfo msiCulture = GetMSICulture(msidb);
                
                foreach (string tableName in tableNames)
                {
                    // these tables cause memory exceptions if you try to access them
                    // they just contain binary data so we can skip them
                    if (!tableName.Equals("icon", StringComparison.InvariantCultureIgnoreCase) && !tableName.Equals("binary", StringComparison.InvariantCultureIgnoreCase))
                    {
                        if (tableName == "CTXUIText")
                        {
                            Lib.Log("");
                        }
                        CultureInfo culture = null;

                        if (isMiuMSI)
                        {
                            if ((culture = IsMuiTable(tableName)) == null)
                            {
                                culture = msiCulture;
                            }
                        }
                        else
                        {
                            culture = msiCulture;
                        }

                        List<TableRow> rows = TableRow.GetRowsFromTable(msidb, tableName);

                        var pkRec = msidb.Tables[tableName].PrimaryKeys;
                        string primaryKeyStr = pkRec[0];
                        int num = 0;

                        try
                        {
                            foreach (TableRow row in rows)
                            {
                                num++;

                                foreach (string columnName in columnNames)
                                {
                                    string text = row.GetString(columnName);

                                    if (text == string.Empty) continue;
                                    string rowName = string.Empty;

                                    try
                                    {
                                        rowName = row.GetString(primaryKeyStr);
                                    }
                                    catch
                                    {
                                        System.Diagnostics.Debug.WriteLine("caught");
                                        rowName = Guid.NewGuid().ToString();
                                    }

                                    try
                                    {
                                        if (text.Length > 1000)
                                        { }
                                        if (text.StartsWith(@"{\rtf1\"))
                                        {
                                            using (RichTextBox rtb = new RichTextBox())
                                            {
                                                rtb.Rtf = text;
                                                text = rtb.Text;
                                            }
                                        }
                                    }
                                    catch (ArgumentException argEx)
                                    {
                                	    Lib.Log(argEx.Message);                                    	
                                    }


                                    //strip out & for accelerator keys
                                    if (text.IndexOf("&") != -1)
                                    {
                                        text = text.Replace("&", string.Empty);
                                    }

                                    if (text == string.Empty) continue;

                                    // skip [name] or {value} as they are not user facing strings
                                    string groupStr = "extract";
                                    string pattern = "[*{\\[](?'" + groupStr + "'.*?)[}\\]]";

                                    MatchCollection matches = Regex.Matches(text, pattern);

                                    foreach (Match match in matches)
                                    {
                                        if (match.Success)
                                        {
                                            string val = match.Groups[groupStr].Value;
                                            if (!val.Equals(string.Empty))
                                            {
                                                text = text.Replace(val, "**");
                                            }
                                        }
                                    }
                                
                                    text = text.Replace("_", " ").Trim();

                                    if (text == string.Empty) continue;

                                    string id = null;
                                    if (rowName == string.Empty)
                                    {
                                        id = string.Format("{0}.{1}.[{2}]", tableName, columnName, num);
                                    }
                                    else
                                    {
                                        id = string.Format("{0}.{1}.[{2}]", tableName, columnName, rowName);
                                    }

                                    StringValue stringVal = new StringValue(id, text, culture);
                                    fileData = HandleSpellingDelegates(checkStringDelegates, fileData, stringVal);

                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Lib.Log(ex);
                        }
                    }
                    else
                    {
                        Lib.Log("Skip table: " + tableName);
                    }
                }

                msidb.Close();
            
            }

             return fileData;
       }


  

        public override string Description { get { return "Spell checks strings in an MSI installer and MST transforms. Use Columns.xml to specify which columns should be spell checked."; } }

        #endregion

        #region IFileHandler Members

        private List<FileExtension> m_SupportedExtensions;

        public override List<FileExtension> SupportedExtensions
        {
            get
            {
                if (m_SupportedExtensions == null)
                {
                    m_SupportedExtensions = new List<FileExtension>();
                    m_SupportedExtensions.Add(new FileExtension(MSI));
                    m_SupportedExtensions.Add(new FileExtension(MST));
                }
                return m_SupportedExtensions;
            }
        }

        #endregion
    }
}
