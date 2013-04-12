using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;
using System.Globalization;
using System.Diagnostics;

using SpellCheck.Common.XmlData;
using SpellCheck.Common;

using SpellCheck.Common.Utility;


namespace PluginWrapper
{
    class Program
    {
        static int Main(string[] args)
        {
            StringFileHandler handler = LoadPlugin(args[0]);

            CheckStringDelegate checkStringDelegate = new CheckStringDelegate(SpellCheck);
            
            IList<CheckStringDelegate> delegates = new List<CheckStringDelegate>();
            delegates.Add(checkStringDelegate);
            System.IO.DirectoryInfo dInfo = new System.IO.DirectoryInfo(args[1]);
            foreach (System.IO.FileInfo fileInfo in dInfo.GetFiles())
            {
                Console.WriteLine(fileInfo.FullName);
                handler.LoadStrings(fileInfo, delegates, null);
            }
       
            return 0;
        }

        private static StringValue SpellCheck(StringValue strVal)
        {
            Lib.Log("{0}-{1}", strVal.Id, strVal.Value);
            
            return null;
        }

        private static StringFileHandler LoadPlugin(string fileName)
        {
            Assembly assembly = Assembly.LoadFrom(fileName);

            AssemblyName[] names =  assembly.GetReferencedAssemblies();

            foreach (Module mod in assembly.GetModules())
            {
                foreach (Type type in mod.GetTypes())
                {
                    if (type.IsSubclassOf(typeof(StringFileHandler)))
                    {
                        byte[] token = assembly.GetName().GetPublicKeyToken();
                        if (token == null || token.Length == 0)
                        {
                            Lib.Log("SpellCheck Plugins must have a strongname: {0}", assembly.FullName);
                            return null; ;
                        }
                        else
                        {
                            StringFileHandler handler = Activator.CreateInstance(type) as StringFileHandler;
                            return handler;
                        }
                    }
                }
            }
            return null;
        }
    }
}
