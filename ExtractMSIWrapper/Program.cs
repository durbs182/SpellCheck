using System;
using System.Collections.Generic;
using System.Text;

namespace ExtractMSI
{
    class Program
    {
        static void Main(string[] args)
        {
            System.IO.FileInfo fileInfo = new System.IO.FileInfo(args[0]);
            System.IO.DirectoryInfo dirInfo = new System.IO.DirectoryInfo(args[1]);
            ExtractMSI.Wixtracts.ExtractFiles(fileInfo, dirInfo);
        }
    }
}
