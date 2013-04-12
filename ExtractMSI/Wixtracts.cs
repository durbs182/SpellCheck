
using System;
using System.Collections;
using System.Collections.Generic;

using System.Diagnostics;
using System.IO;
using System.Threading;
//using Microsoft.Tools.WindowsInstallerXml.Cab;
//using Microsoft.Tools.WindowsInstallerXml.Msi;

using SpellCheck.Common.Utility;
using Microsoft.Deployment.WindowsInstaller.Package;
using Microsoft.Deployment.WindowsInstaller;

namespace ExtractMSI
{
	public class Wixtracts
	{
        private const string TEMP_DIR_NAME = "WITEMP";

		/// <summary>
		/// Extracts the compressed files from the specified MSI file to the specified output directory.
		/// If specified, the list of <paramref name="filesToExtract"/> objects are the only files extracted.
		/// </summary>
		/// <param name="filesToExtract">The files to extract or null or empty to extract all files.</param>
		/// <param name="progressCallback">Will be called during during the operation with progress information, and upon completion. The argument will be of type <see cref="ExtractionProgress"/>.</param>
        public static void ExtractFiles(FileInfo msiFilePath, DirectoryInfo outputDir)
		{
            if (msiFilePath == null)
				throw new ArgumentNullException("msi");
			if (outputDir == null)
				throw new ArgumentNullException("outputDir");

			//int filesExtractedSoFar = 0;

            using (var msiPackage = new InstallPackage(msiFilePath.FullName, DatabaseOpenMode.Transact))
            {
                msiPackage.WorkingDirectory = outputDir.FullName;

                var dirMapping = msiPackage.Directories;

                msiPackage.UpdateDirectories();
                msiPackage.ExtractFiles();

                //Close **without** calling Commit() to ensure changes are not persisted                  
                msiPackage.Close();
            }


            // cleanup temp dir

            var tempDir = Path.Combine(outputDir.FullName, TEMP_DIR_NAME);

            if (Directory.Exists(tempDir))
            {
                Lib.Log("Delete WIX temp directory: {0}", tempDir);
                Directory.Delete(tempDir, true);
            }

            //Database msidb = new Database(msi.FullName, OpenDatabase.ReadOnly);
            //try(
            //{
            //    List<MsiFile> filesToExtract = MsiFile.CreateMsiFilesFromMSI(msidb);

            //    if (!msi.Exists)
            //    {
            //        Lib.Log("File \'" + msi.FullName + "\' not found.");
            //        return;
            //    }

            //    outputDir.Create();

            //    //map short file names to the msi file entry
            //    Dictionary<string, MsiFile> fileEntryMap = new Dictionary<string, MsiFile>(StringComparer.InvariantCultureIgnoreCase);
 
            //    foreach (MsiFile fileEntry in filesToExtract)
            //        fileEntryMap[fileEntry.File] = fileEntry;

            //    //extract ALL the files to the folder:
            //    List<DirectoryInfo>mediaCabs = ExplodeAllMediaCabs(msidb, outputDir);

            //    // now rename or remove any files not desired.
            //    foreach (DirectoryInfo cabDir in mediaCabs)
            //    {
            //        foreach (FileInfo sourceFile in cabDir.GetFiles())
            //        {
            //            if (fileEntryMap.ContainsKey(sourceFile.Name))
            //            {
            //                MsiFile entry = fileEntryMap[sourceFile.Name];

            //                if (entry != null)
            //                {
            //                    DirectoryInfo targetDirectoryForFile = GetTargetDirectory(outputDir, entry.Directory);

            //                    //TODO: need to add some logic to truncate the path if it's too long, not sure if this will actually occur??
            //                    if (targetDirectoryForFile.FullName.Length > 248)
            //                    {
            //                        Lib.Log("targetDirectoryForFile length is greater than 248: {0}",targetDirectoryForFile.FullName);
            //                        throw new System.IO.PathTooLongException();                                  
            //                    }
								 
            //                    string destName = Path.Combine(targetDirectoryForFile.FullName, entry.LongFileName);

            //                    if (File.Exists(destName))
            //                    {
            //                        //make unique
            //                        Lib.Log(string.Concat("Duplicate file found \'", destName, "\'"));
            //                        int duplicateCount = 0;
            //                        string uniqueName = destName;
            //                        do
            //                        {
            //                            uniqueName = string.Concat(destName, ".", "duplicate", ++duplicateCount);
            //                        } while (File.Exists(uniqueName));
            //                        destName = uniqueName;
            //                    }
            //                    //rename
            //                    Lib.Log(string.Concat("Renaming File \'", sourceFile.Name, "\' to \'", entry.LongFileName, "\'"));
            //                    try
            //                    {   
            //                        sourceFile.MoveTo(destName);

            //                        filesExtractedSoFar++;
            //                    }
            //                    catch (System.IO.PathTooLongException ptlex)
            //                    {
            //                        Lib.Log("Exception caught moving {0} to {1}:  {2}", sourceFile, destName, ptlex.Message);
            //                    }
            //                }
            //            }
            //            else
            //            { 
            //                Lib.Log("{0} not found in fileEntryMap",sourceFile.Name);
            //            }
            //        }
            //        cabDir.Delete(true);
            //    }
            //}
            //finally
            //{
            //    if (msidb != null)
            //        msidb.Close();
            //}
		}

		/// <summary>
		/// Create the directory for the MSI file about to be extracted
		/// </summary>
		/// <param name="rootDirectory"></param>
		/// <param name="relativePath"></param>
		/// <returns></returns>
        //private static DirectoryInfo GetTargetDirectory(DirectoryInfo rootDirectory, MsiDirectory relativePath)
        //{
        //    string fullPath = Path.Combine(rootDirectory.FullName, relativePath.GetPath());
            
        //    try
        //    {
        //        if (!Directory.Exists(fullPath))
        //        {
        //            Directory.CreateDirectory(fullPath);
        //        }
        //        return new DirectoryInfo(fullPath);
        //    }
        //    catch(Exception ex)
        //    {
        //        Lib.Log(ex.Message);
        //    }
        //    return null;
        //}

		/// <summary>
		/// Dumps the entire contents of each cab into it's own subfolder in the specified baseOutputPath.
		/// </summary>
		/// <remarks>
		/// A list of Directories containing the files that were the contents of the cab files.
		/// </remarks>
        //private static List<DirectoryInfo> ExplodeAllMediaCabs(Database msidb, DirectoryInfo baseOutputPath)
        //{
        //    List<DirectoryInfo> cabFolders = new List<DirectoryInfo>();

        //    const string tableName = "Media";
        //    if (!msidb.TableExists(tableName))
        //    {
        //        //return (DirectoryInfo[]) cabFolders.ToArray(typeof (DirectoryInfo));
        //        return cabFolders;
        //    }

        //    string query = String.Concat("SELECT * FROM `", tableName, "`");
        //    using (View view = msidb.OpenExecuteView(query))
        //    {
        //        Record record;
        //        while (view.Fetch(out record))
        //        {
        //            const int MsiInterop_Media_Cabinet = 4;
        //            string cabSourceName = record[MsiInterop_Media_Cabinet];

        //            DirectoryInfo cabFolder;
        //            // ensure it's a unique folder
        //            int uniqueCounter = 0;
        //            do
        //            {
        //                cabFolder = new DirectoryInfo(Path.Combine(baseOutputPath.FullName, string.Concat("_cab_", cabSourceName, ++uniqueCounter)));
        //            } while (cabFolder.Exists);

        //            Lib.Log(string.Concat("Exploding media cab \'", cabSourceName, "\' to folder \'", cabFolder.FullName, "\'."));
        //            if (0 < cabSourceName.Length)
        //            {
        //                if (cabSourceName.StartsWith("#"))
        //                {
        //                    cabSourceName = cabSourceName.Substring(1);

        //                    // extract cabinet, then explode all of the files to a temp directory
        //                    string cabFileSpec = Path.Combine(baseOutputPath.FullName, cabSourceName);

        //                    ExtractCabFromPackage(cabFileSpec, cabSourceName, msidb);
        //                    WixExtractCab extCab = new WixExtractCab();
        //                    if (File.Exists(cabFileSpec))
        //                    {
        //                        cabFolder.Create();
        //                        // track the created folder so we can return it in the list.
        //                        cabFolders.Add(cabFolder);

        //                        extCab.Extract(cabFileSpec, cabFolder.FullName);
        //                    }
        //                    extCab.Close();
        //                    File.Delete(cabFileSpec);
        //                }
        //            }
        //        }
        //    }

        //    return cabFolders;
        //}


		/// <summary>
		/// Write the Cab to disk.
		/// </summary>
		/// <param name="filePath">Specifies the path to the file to contain the stream.</param>
		/// <param name="cabName">Specifies the name of the file in the stream.</param>
        //public static void ExtractCabFromPackage(string filePath, string cabName, Database inputDatabase)
        //{
        //    using (View view = inputDatabase.OpenExecuteView(string.Format("SELECT * FROM `_Streams` WHERE `Name` = '{0}'", cabName)))
        //    {
        //        Record record;
        //        if (view.Fetch(out record))
        //        {
        //            FileStream cabFilestream = null;
        //            BinaryWriter writer = null;
        //            try
        //            {
        //                cabFilestream = new FileStream(filePath, FileMode.Create);

        //                // Create the writer for data.
        //                writer = new BinaryWriter(cabFilestream);
        //                int count = 512;
        //                byte[] buf = new byte[count];
        //                while (count == buf.Length)
        //                {
        //                    const int MsiInterop_Storages_Data = 2; //From wiX:Index to column name Data into Record for row in Msi Table Storages
        //                    count = record.GetStream(MsiInterop_Storages_Data, buf, count);
        //                    if (buf.Length > 0)
        //                    {
        //                        // Write data to Test.data.
        //                        writer.Write(buf);
        //                    }
        //                }
        //            }
        //            finally
        //            {
        //                if (writer != null)
        //                {
        //                    writer.Close();
        //                }

        //                if (cabFilestream != null)
        //                {
        //                    cabFilestream.Close();
        //                }
        //            }
        //        }
        //    }
        //}
	}
}