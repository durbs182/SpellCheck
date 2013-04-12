using System.Collections;
using System.IO;
using System.Collections.Generic;

using Microsoft.Tools.WindowsInstallerXml.Msi;

namespace ExtractMSI
{
	/// <summary>
	/// Represents an entry in the Directory table of an MSI file.
	/// </summary>
	public class MsiDirectory
	{
		private string _name = "";
		private string _shortName = "";
		//the "DefaultDir value from the MSI table.
		private string _defaultDir="";
		/// The "Directory" entry from the MSI
		private string _directory="";
		/// The "Directory_Parent" entry
		private string _directoryParent;
		/// Stores the child directories
		private ArrayList _children = new ArrayList();
		private MsiDirectory _parent;

		private MsiDirectory()
		{
		}

		/// <summary>
		/// Returns the name of this directory.
		/// </summary>
		public string Name
		{
			get { return _name; }
		}

		/// <summary>
		/// Returns the alternative short name (8.3 format) of this directory.
		/// </summary>
		public string ShortName
		{
			get { return _shortName; }
		}

		/// The "Directory" entry from the MSI
		public string Directory
		{
			get { return _directory; }
		}

		/// The "Directory_Parent" entry
		public string DirectoryParent
		{
			get { return _directoryParent; }
		}

		/// <summary>
		/// The direct child directories of this directory.
		/// </summary>
		public ICollection Children
		{
			get
			{
				return _children;
			}
		}

		/// <summary>
		/// Returns this directory's parent or null if it is a root directory.
		/// </summary>
		public MsiDirectory Parent
		{
			get { return _parent; }
		}

		/// <summary>
		/// Returns the full path considering it's parent directories.
		/// </summary>
		/// <returns></returns>
		public string GetPath()
		{
            string path = this.Name;

  
			MsiDirectory parent = this.Parent;
			while (parent != null)
			{
				path = Path.Combine(parent.Name, path);
				parent = parent.Parent;
			}

           ArrayList invalidPathChars = new ArrayList();
            invalidPathChars.AddRange(Path.GetInvalidPathChars());
            invalidPathChars.Add('^');
            invalidPathChars.Add('<');
            invalidPathChars.Add('>');
            invalidPathChars.Add(':');
            invalidPathChars.Add('|');
            invalidPathChars.Add('?');
            invalidPathChars.Add('*');
            //invalidPathChars.Add('.');


            foreach (char invalidChar in invalidPathChars)
                path = path.Replace(invalidChar, '_'); 

			return path;
		}

		/// <summary>
		/// Creates a list of <see cref="MsiFile"/> objects from the specified database.
		/// </summary>
		/// <param name="allDirectories">All directories in the table.</param>
		/// <param name="msidb">The databse to get directories from.</param>
		/// <param name="rootDirectories">
		/// Only the root directories (those with no parent). Use <see cref="MsiDirectory.Children"/> to traverse the rest of the directories.
		/// </param>
		public static void GetMsiDirectories(Database msidb, out MsiDirectory[] rootDirectories, out MsiDirectory[] allDirectories)
		{
			List<TableRow> rows = TableRow.GetRowsFromTable(msidb, "Directory");
			Hashtable directoriesByDirID = new Hashtable();

			foreach (TableRow row in rows)
			{
				MsiDirectory directory = new MsiDirectory();
				directory._defaultDir = row.GetString("DefaultDir");
				if (directory._defaultDir != null && directory._defaultDir.Length > 0)
				{
					string[] split = directory._defaultDir.Split('|');
					if (split.Length >= 1)
					{
						directory._shortName = split[0];
						if (split.Length > 1)
							directory._name = split[1];
						else
							directory._name = split[0];
					}
				}
				directory._directory = row.GetString("Directory");
				directory._directoryParent = row.GetString("Directory_Parent");
				directoriesByDirID.Add(directory.Directory, directory);
			}
			//Now we have all directories inthe table, create a structure for them based on their parents.
			ArrayList rootDirectoriesList = new ArrayList();
			foreach (MsiDirectory dir in directoriesByDirID.Values)
			{
				if (dir.DirectoryParent == null || dir.DirectoryParent.Length == 0)
				{
					rootDirectoriesList.Add(dir);
					continue;
				}

				MsiDirectory parent = directoriesByDirID[dir.DirectoryParent] as MsiDirectory;
				dir._parent = parent;
				parent._children.Add(dir);
			}
			// return the values:
			rootDirectories = (MsiDirectory[])rootDirectoriesList.ToArray(typeof(MsiDirectory));
			
			MsiDirectory[] allDirectoriesLocal = new MsiDirectory[directoriesByDirID.Values.Count];
			directoriesByDirID.Values.CopyTo(allDirectoriesLocal,0);
			allDirectories = allDirectoriesLocal;
		}
	}
}