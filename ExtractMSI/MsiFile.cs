using System.Collections;
using System.Collections.Generic;

using System.Diagnostics;
using Microsoft.Tools.WindowsInstallerXml.Msi;

using SpellCheck.Common.Utility;

namespace ExtractMSI
{
	/// <summary>
	/// Represents a file in the msi file table/view.
	/// </summary>
	public class MsiFile
	{
		public string File;// a unique id for the file
		public string LongFileName;
		public string ShortFileName;
		public int FileSize;
		public string Version;
		public string Component;
		private MsiDirectory _directory;

		/// <summary>
		/// Returns the directory that this file belongs in.
		/// </summary>
		public MsiDirectory Directory
		{
			get { return _directory; }
		}

		private MsiFile()
		{
		}

		/// <summary>
		/// Creates a list of <see cref="MsiFile"/> objects from the specified database.
		/// </summary>
        public static List<MsiFile> CreateMsiFilesFromMSI(Database msidb)
		{
			List<TableRow> rows = TableRow.GetRowsFromTable(msidb, "File");

			// do some prep work to cache values from MSI for finding directories later...
			MsiDirectory[] rootDirectories;
			MsiDirectory[] allDirectories;
			MsiDirectory.GetMsiDirectories(msidb, out rootDirectories, out allDirectories);

			//find the target directory for each by reviewing the Component Table
			List<TableRow> components = TableRow.GetRowsFromTable(msidb, "Component");
			//build a table of components keyed by it's "Component" column value
			Hashtable componentsByComponentTable = new Hashtable();
			foreach (TableRow component in components)
            {
                string str = component.GetString("Component").Trim();
				componentsByComponentTable[component.GetString("Component")] = component;
			}

            List<MsiFile> files = new List<MsiFile>(rows.Count);
			foreach (TableRow row in rows)
			{
				MsiFile file = new MsiFile();
				
				string fileName = row.GetString("FileName");
				string[] split = fileName.Split('|');
				file.ShortFileName = split[0];
				if (split.Length > 1)
					file.LongFileName = split[1];
				else
					file.LongFileName = split[0];

				file.File = row.GetString("File");
				file.FileSize = row.GetInt32("FileSize");
				file.Version = row.GetString("Version");
				file.Component = row.GetString("Component_");

            	file._directory = GetDirectoryForFile(file, allDirectories, componentsByComponentTable);

                if (file._directory != null)
                {
                    files.Add(file);
                }
                else
                {
                    Lib.ConsoleLog("***Skipping {0} as it has an invalid MSI entry", fileName);
                }
			}
			return files;
		}

		private static MsiDirectory GetDirectoryForFile(MsiFile file, MsiDirectory[] allDirectories, IDictionary componentsByComponentTable)
		{
			// get the component for the file
			TableRow componentRow = componentsByComponentTable[file.Component] as TableRow;
			if (componentRow == null)
			{
				Lib.Log("File '{0}' has no component entry.", file.LongFileName);
				return null;
			}
			// found component, get the directory:
			string componentDirectory = componentRow.GetString("Directory_");
			MsiDirectory directory = FindDirectoryByDirectoryKey(allDirectories, componentDirectory);
			if (directory != null)
			{
				Lib.Log("Directory for '{0}' is '{1}'.", file.LongFileName, directory.GetPath());
			}
			else
			{
				Lib.Log("directory not found for file '{0}'.", file.LongFileName);
			}
			return directory;
			
		}

		/// <summary>
		/// Returns the directory with the specified value for <see cref="MsiDirectory.Directory"/> or null if it cannot be found.
		/// </summary>
		/// <param name="directory_Value">The value for the sought directory's <see cref="MsiDirectory.Directory"/> column.</param>
		private static MsiDirectory FindDirectoryByDirectoryKey(MsiDirectory[] allDirectories, string directory_Value)
		{
			foreach (MsiDirectory dir in allDirectories)
			{
				if (0 == string.CompareOrdinal(dir.Directory, directory_Value))
				{
					return dir;
				}
			}
			return null;
		}
	}
}
