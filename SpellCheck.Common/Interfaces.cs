using System;
using System.Collections.Generic;
using System.Text;
using System.Collections.Specialized;
using System.Globalization;
using System.IO;
using SpellCheck.Common.XmlData;
using SpellCheck.Common.Utility;

namespace SpellCheck.Common
{

	public delegate StringValue CheckStringDelegate(StringValue stringVal);

	#region interfaces

	public interface IFileHandler
	{
		List<FileExtension> SupportedExtensions { get; }
		string Description { get; }
	}

	public interface IArchiveFileHandler : IFileHandler
	{
		System.IO.DirectoryInfo Extract(System.IO.FileInfo archiveFile);
	}

	public interface IStringResourceFileHandler : IFileHandler
	{
		FileData LoadStrings(System.IO.FileInfo fileInfo, IList<CheckStringDelegate> checkStringDelegates,CultureInfo locale);
		string StrongName { get; }
	}

    public interface ISpellingEngine 
	{
		StringValue CheckString(StringValue stringVal);
	}
    
	#endregion

	#region abstract classes
	[Serializable]
	public abstract class SpellingEngineBase : MarshalByRefObject,  ISpellingEngine
	{

		public abstract StringValue CheckString(StringValue stringVal);
		
		//public abstract ISpellingEngine CreateSpellingEngineInstance(System.Collections.Specialized.StringCollection alertStrings, System.Globalization.CultureInfo locale, bool splitCamelCase) ;

		public abstract void CheckForAlertStrings(System.Collections.Specialized.StringCollection alertStrings, string stringToCheck);
		
		/// <summary>
		/// Excel can't handle text longer than 32767 in a cell. IF longer
		/// store strings in a file and referernce the file path in the stringvalue
		/// </summary>
		/// <param name="stringValue"></param>
		/// <returns></returns>
		protected StringValue ExcelCheckString(StringValue stringValue)
		{
            if(stringValue.Value.Length > Lib.EXCELTEXTMAXLENGTH)
            {
            	string guid = Guid.NewGuid().ToString();
            	
            	string longbadstring_fileName = string.Format( "longbadstring_{0}.txt",guid);
            	string longgoodstring_fileName = string.Format( "longgoodstring_{0}.txt",guid);
            	

            	// create a directory path to store the long string files 
            	string longStringFilePath = Path.Combine(Lib.ReportPath.FullName, Lib.ReportName);
            	
            	// check whether the directory exists create it if it doesn't
            	if(!Directory.Exists(longStringFilePath))
            	{
            		Directory.CreateDirectory(longStringFilePath);
            	}
				
            	// create a file to store the long good and bad string
            	string longbadstring_fileNamePath = Path.Combine(longStringFilePath,longbadstring_fileName);
            	string longgoodstring_fileNamePath = Path.Combine(longStringFilePath,longgoodstring_fileName);
            	File.WriteAllText(longbadstring_fileNamePath,stringValue.Value);
            	File.WriteAllText(longgoodstring_fileNamePath,stringValue.SuggestedCorrection);
            	
            	// replace the error with the file path
            	stringValue.Value = longbadstring_fileNamePath;
            	// replace the suggested correctioon with the file path
            	stringValue.SuggestedCorrection = longgoodstring_fileNamePath;
            }
            
            return stringValue;
		}
	}
	

	[Serializable]
	public abstract class StringFileHandler : IStringResourceFileHandler
	{
		#region IStringResourceFileHandler Members

		public string StrongName {
			get { return this.GetType().Assembly.FullName; }
		}

		public abstract FileData LoadStrings(System.IO.FileInfo fileInfo, IList<CheckStringDelegate> checkStringDelegates, CultureInfo locale);

		public abstract List<FileExtension> SupportedExtensions { get; }
		public abstract string Description { get; }
		
		public FileData HandleSpellingDelegates(IList<CheckStringDelegate> checkStringDelegates, FileData fileData, StringValue stringValue)
		{
			IList<StringValue> stringValues = new List<StringValue>();
			
			foreach(CheckStringDelegate csDelegate in checkStringDelegates)
			{
				StringValue strVal = csDelegate( stringValue);
				
				// double check that the engine plugin has validated this already
				if(stringValue.Value.Length > Lib.EXCELTEXTMAXLENGTH || (stringValue.SuggestedCorrection != null && stringValue.SuggestedCorrection.Length > Lib.EXCELTEXTMAXLENGTH))
				{
					throw new ArgumentException(string.Format("StringValue contains values that will exceed the cell text limted in Excel: {0} {1}" ,fileData.FilePath, stringValue.Id));
				}
				
				if(strVal != null)
				{
					 stringValues.Add(strVal);
				}
			}
                
			foreach(StringValue stringValueOut in stringValues)
			{
				try
				{
		            if(stringValueOut.Errors.Count != 0)
		            {
		            	fileData.AddString(stringValueOut);
		            }
				}
	            catch(Exception ex)
	            {
	            	Utility.Lib.Log(ex);
	            }
			}
        
        	return fileData;
		}

		#endregion
	}


	public abstract class ArchiveFileHandler : IArchiveFileHandler
	{
		#region IArchiveFileHandler Members

		public virtual DirectoryInfo Extract(FileInfo archiveFile)
		{
			//trim dir name down or the path ends up too long
			
			string pathTempStr = archiveFile.Name.Replace(".", "_"); //+ "_ext";
			
			string[] bits = pathTempStr.Split(new string[]{"_"},StringSplitOptions.None);
			
			StringBuilder truncatedStr = new StringBuilder();
			
			foreach(string bit in bits)
			{
				if(bit.Length <3)
				{
					truncatedStr.Append(bit);
				}
				else
				{
					truncatedStr.Append(bit.Substring(0,3));
				}
				
				truncatedStr.Append("_");
			}
			
			pathTempStr = truncatedStr.ToString() + "ext";
			
			DirectoryInfo archiveFileDirInfo = new DirectoryInfo(archiveFile.DirectoryName);

			if (archiveFileDirInfo.Attributes == (FileAttributes.ReadOnly | FileAttributes.Directory)) 
			{	
				pathTempStr = EnsureUniqueDirName(Path.Combine(Utility.Lib.ExtractionDirectory.FullName, pathTempStr));
				DirectoryInfo dinfo = new DirectoryInfo(pathTempStr);
				Utility.Lib.ConsoleLog(string.Format("{0} is ReadOnly. {1} will be extracted to {2}", archiveFileDirInfo, archiveFile.Name, dinfo));
				return dinfo;
			} 
			else
			{
				pathTempStr = EnsureUniqueDirName(Path.Combine(archiveFileDirInfo.FullName, pathTempStr));
				return new DirectoryInfo(pathTempStr);
			}
		}
		
		private string EnsureUniqueDirName(string dirName)
		{
			int id = 1;
			while(Directory.Exists(dirName))
			{
				dirName = dirName + id;
				id++;						
			}
			return dirName;
		}

		#endregion

		#region IFileHandler Members

		public abstract List<FileExtension> SupportedExtensions { get; }
		public abstract string Description { get; }


		#endregion
	}

	#endregion



	
	[Serializable]
	public sealed class FileExtension
	{
		public override string ToString()
		{
			return Extension;
		}

		string m_fileExtension;

		public string Extension {
			get { return m_fileExtension; }
		}

		public string DotExtension {
			get { return "." + m_fileExtension; }
		}


		public string StarDotExtension {
			get { return "*." + m_fileExtension; }
		}

		public FileExtension(string extension)
		{
			m_fileExtension = extension;
		}

	}


		
		
}
