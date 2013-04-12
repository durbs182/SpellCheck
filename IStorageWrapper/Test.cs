using System;
using System.IO;
using System.Runtime.InteropServices;

using System.Security;

using HtmlHelp.Storage;



namespace TestStorage
{
    
	class Test
	{
		[STAThread]
		static void Main(string[] args)
		{
			// Create Instance of ITStorageWrapper.
			// During initialization constructor will process CHM file 
			// and create collection of file objects stored inside CHM file.
            ITStorageWrapper iw = new ITStorageWrapper(@"C:\msi\ASC_XenDesktop_msi_extracted\SourceDir\Citrix\Access Management Console - Help\asc_help_dse.chm",true);

			// Loop through collection of objects stored inside IStorage
            foreach (FileObject fileObject in iw.foCollection)
			{
				// Check to make sure we can READ stream of an individual file object
				if (fileObject.CanRead)
				{
					// We only want to extract HTM files in this example
					// fileObject is our representation of internal file stored in IStorage
					if (fileObject.FileName.EndsWith(".htm"))
					{
						Console.WriteLine("Path: " + fileObject.FilePath);
						Console.WriteLine("File: " + fileObject.FileName);

						// FileUrl - is a external reference to the internal object
						// allows you to display content of single file in Internet Explorer
						// without extracting content from the archive
						Console.WriteLine("Url: " + fileObject.FileUrl);
			
						string fileString = fileObject.ReadFromFile();

                        fileString = StripTags(fileString);

                        fileString = fileString.Replace("&nbsp;", string.Empty).Trim();

						Console.WriteLine("Text: " + fileString);

						// Direct Extraction sample
						//fileObject.Save(@"c:\test1\" + fileObject.FileName);

						// Read first and then save later example
						StreamWriter sw = File.CreateText(@"c:\test1\" + fileObject.FileName);
						sw.WriteLine(fileString);
						sw.Close();

						Console.ReadLine();
					}
				}
			}
			Console.ReadLine();
		}

        private static string StripTags(string HTML)
        {
            // Removes tags from passed HTML            
            System.Text.RegularExpressions.Regex objRegEx = new System.Text.RegularExpressions.Regex("<[^>]*>");

            return objRegEx.Replace(HTML, "");
        }
	}
}
