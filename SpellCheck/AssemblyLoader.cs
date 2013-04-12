using System;
using System.Collections.Generic;
using System.Reflection;
using System.IO;
using System.Windows.Forms;
using System.Globalization;


using SpellCheck.Common.XmlData;
using SpellCheck.Common;
using SpellCheck.Common.Utility;



namespace SpellCheck
{
    public class AssemblyLoader: MarshalByRefObject
    {


        static Dictionary<string, FileData> assemblyLog;

        static AssemblyLoader()
        {
            assemblyLog = new Dictionary<string, FileData>();
        }

        /// <summary>
        /// Load assemblies into a seperate AppDomain
        /// </summary>
        /// <param name="path"></param>
        /// <param name="plugin"></param>
        /// <param name="checkString"></param>
        /// <param name="wordApp"></param>
        /// <returns></returns>
        internal static FileData LoadFileData(string path, InternalStringFileHandlerPlugin plugin, IList<CheckStringDelegate> checkStringDelegates,System.Globalization.CultureInfo locale)
        {
            //AppDomain sandbox = null;

            try
            {
            	string adstr = AppDomain.CurrentDomain.FriendlyName;

                //AssemblyLoader assemblyLoader = Lib.CreateObjectInstance<AssemblyLoader>(out sandbox);
                AssemblyLoader assemblyLoader = Citrix.Automation.ObjectCreator.LocalObjectFactory<AssemblyLoader>.Instance.GetObject(true);
                
                Dictionary<string, FileData> fileDataTable = assemblyLoader.LoadAssembly(path, plugin, checkStringDelegates, locale, assemblyLog);
                if (fileDataTable.Count == 0)
                {
                    return null;
                }
                else
                {
                    FileData fileData = null;
                    foreach(string key in fileDataTable.Keys)
                    {
                        fileData = fileDataTable[key];
                        assemblyLog.Add(key, fileData);
                    }

                    return fileData;
                }
            }
            catch (System.Exception systemException)
            {
                Lib.Log(systemException.Message);
                //return null;
                throw;
            }
            finally
            {
                Citrix.Automation.ObjectCreator.IsolatedObjectManager.Instance.DisposeAll();
                //AppDomain.Unload(sandbox);
                //GC.Collect();
            }            
        }

        
        /// <summary>
        /// Load each assembly, then enumerate the embedded resources and find all the strings
        /// </summary>
        /// <param name="assemblyPath"></param>
        /// <param name="plugin"></param>
        /// <param name="checkString"></param>
        /// <param name="wordApp"></param>
        /// <returns></returns>
        private Dictionary<string,FileData> LoadAssembly(string assemblyPath, 
            InternalStringFileHandlerPlugin plugin, 
            IList<CheckStringDelegate> checkStringDelegates, 
            CultureInfo locale,
            Dictionary<string, FileData> assemblyLog)
        {
            FileData fileData = null;
            Dictionary<string, FileData> result = new Dictionary<string, FileData>();
            try
            {
            	
            	Lib.Log("Executing {0} in AppDomain: {1}", Lib.GetMethodName(), AppDomain.CurrentDomain.FriendlyName);

                Assembly assembly = Assembly.LoadFile(assemblyPath);

                if (assemblyLog.ContainsKey(assembly.FullName))
                {
                    Lib.ConsoleLog("***Skipping [{0}] aleady checked", assembly.FullName);
                    return result;
                }
                
				System.Globalization.CultureInfo assCulture = null;
                try
                {
                 	assCulture = assembly.GetName().CultureInfo;
                }
                catch(System.ArgumentException argEx)// if getting the local from the assembly fails set to InvariantCulture
                {
                	Lib.Log(argEx.Message);
                	Lib.Log("Set assembly culture to {0}", System.Globalization.CultureInfo.InvariantCulture.DisplayName);
                	assCulture = System.Globalization.CultureInfo.InvariantCulture;
                }

                // if a locale is set only check the string if the file CultureInfo matches
                if (locale != null)
                {
                    if (!locale.Equals(assCulture) && !assCulture.Equals(System.Globalization.CultureInfo.InvariantCulture))
                        return result;
                }

                if (assembly.GetManifestResourceNames().Length == 0)
                {
                    return result;
                }

                fileData = new FileData(assemblyPath);

                result.Add(assembly.FullName, fileData);

                foreach (string resourceName in assembly.GetManifestResourceNames())
                {
                    Lib.Log("{0} {1}", resourceName, Lib.GetMethodName());

                    try
                    {
                        System.Resources.ResourceReader resReader = new System.Resources.ResourceReader(assembly.GetManifestResourceStream(resourceName));
                        System.Collections.IDictionaryEnumerator idicEnum = resReader.GetEnumerator();

                        while (idicEnum.MoveNext())
                        {
                            if (idicEnum.Value is string && !plugin.ExcludedKey(idicEnum.Key.ToString()))
                            {
                                string str = plugin.ParseString(idicEnum.Value as string);

                                if (str.Length != 0)
                                {
                                    StringValue stringValue = new StringValue(string.Format("{0}[{1}]", resourceName, idicEnum.Key), str, assCulture);

                                    
                                    fileData = plugin.HandleSpellingDelegates(checkStringDelegates,fileData,stringValue);
                                }
                  
                            }
                        }
                        resReader.Close();

                    }
                    //catch any other embedded resource
                    catch (System.ArgumentException argex)
                    {
                        Lib.Log("{0} Exception loading resource from: {1} [{2}]", Lib.GetMethodName(), assemblyPath, argex.Message);
                    }

                }
            }
            catch (BadImageFormatException badImgEx)
            {
                Lib.Log("File not assembly: {0} [{1}]", assemblyPath, badImgEx.Message);
            }

            catch (FileLoadException flex)
            {
                Lib.Log("File not assembly: {0} [{1}]", assemblyPath, flex.Message);
            }
            catch (FileNotFoundException fnfex)
            {
                Lib.Log("Error loading assembly: {0} [{1}]", assemblyPath, fnfex.Message);
            }


            return result;
        }

        
    }
}
