using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;


using Microsoft.Office.Interop.Word;


using Reflector.CodeModel;
using Reflector.CodeModel.Memory;

using SpellCheck.Interfaces;
using SpellCheck.Utility;

namespace Reflector.AssemblyHardCodedStringHandler
{
	/// <summary>
	/// Summary description for FileDisassemblerHelper.
	/// </summary>
    /// [
    /// 
    public class Helper: MarshalByRefObject
	{
		private IAssemblyManager m_assemblyManager;
        private ILanguageManager m_languageManager;
        private ITranslatorManager m_translatorManager;
        private IServiceProvider m_serviceProvider;



		public Helper()
		{
            m_serviceProvider = new ApplicationManager(null);
            m_assemblyManager = (IAssemblyManager)m_serviceProvider.GetService(typeof(IAssemblyManager));
            m_languageManager = (ILanguageManager)m_serviceProvider.GetService(typeof(ILanguageManager));
            m_translatorManager = (ITranslatorManager)m_serviceProvider.GetService(typeof(ITranslatorManager));
        }

        public ILanguageManager LanguageManager
		{
			get { return m_languageManager; }
		}

        public ITranslatorManager TranslatorManager
		{
			get { return m_translatorManager; }
		}

		public IAssemblyManager AssemblyManager
		{
			get	{ return m_assemblyManager; }
		}

        /// <summary>
        /// Use Reflector's ability to translate back into C# to get access tio hard coded strings
        /// </summary>
        /// <param name="fileInfo"></param>
        /// <returns></returns>
        public Dictionary<string, string> GetStrings(FileInfo fileInfo)
		{
            IAssembly assembly = m_assemblyManager.LoadFile(fileInfo.FullName);

			ILanguageWriterConfiguration configuration = new LanguageWriterConfiguration();
            IAssemblyResolver resolver = AssemblyManager.Resolver;
            AssemblyManager.Resolver = new AssemblyResolver(resolver);

            LoadReferences(assembly);

			try
			{
				// write assembly
				if (assembly != null)
				{
                    Dictionary<string, string> strings = new Dictionary<string, string>();

					// write modules
					string location = string.Empty;
					foreach (IModule module in assembly.Modules)
					{
						if (module.Location != location)
							location = module.Location;

						foreach (ITypeDeclaration typeDeclaration in module.Types)
						{
                            if ((typeDeclaration.Namespace.Length != 0) || (
								(typeDeclaration.Name != "<Module>") &&
								(typeDeclaration.Name != "<PrivateImplementationDetails>")))
							{
                                WriteTypeDeclaration(typeDeclaration, configuration, strings);
							}
						}
					}

                    return strings;
				}
		    }
			catch (Exception ex)
			{
                Lib.Log("Error: " + ex.Message);
			}
			finally
			{
				AssemblyManager.Resolver = resolver;
			}

            return null;
		}

        /// <summary>
        /// Load a known references, stops reflector popping up a dialog.
        /// </summary>
        /// <param name="assembly"></param>
        private void LoadReferences(IAssembly assembly)
        {
            List<string> searchPaths = new List<string>();
            searchPaths.Add(new FileInfo( assembly.Location).DirectoryName + @"\");
            searchPaths.Add(System.Runtime.InteropServices.RuntimeEnvironment.GetRuntimeDirectory());
            searchPaths.Add(System.Runtime.InteropServices.RuntimeEnvironment.GetRuntimeDirectory()+ @"..\");

            List<string> extensions = new List<string>();
            extensions.Add("dll");
            extensions.Add("exe");

            foreach (IModule mod in assembly.Modules)
            {
                foreach (IAssemblyReference refer in mod.AssemblyReferences)
                {
                    string s = refer.Name;
                    bool loaded = false;
                    foreach (IAssembly loadedAss in AssemblyManager.Assemblies)
                    {
                        if (loadedAss.Name.Equals(s, StringComparison.InvariantCultureIgnoreCase))
                        {
                            loaded = true;
                            break;
                        }
                    }

                    if (!loaded)
                    {
                        bool loadFailed = false;

                        foreach(string path in searchPaths)
                        {           
                            foreach(string extension in extensions)
                            {
                                string fullpath = string.Format(@"{0}{1}.{2}", path, refer.Name,extension);

                                if (File.Exists(fullpath))
                                {
                                    IAssembly justLoadedAss = AssemblyManager.LoadFile(fullpath);

                                    if (justLoadedAss.Type != AssemblyType.None)
                                    {
                                        loaded = true;
                                        break;
                                    }
                                    else
                                    {
                                        Lib.Log("{0} ERROR: {1} exists but failed to load.",Lib.MethodName(), fullpath);
                                        AssemblyManager.Unload(justLoadedAss);
                                        loadFailed = true;
                                        break;
                                    }
                                }
                            }

                            if (loaded || loadFailed) break;
                        }
                    }
                }
            }
        }


        private Dictionary<string, string> WriteTypeDeclaration(ITypeDeclaration typeDeclaration, ILanguageWriterConfiguration configuration, Dictionary<string, string> strings)
		{
			ILanguage language = LanguageManager.ActiveLanguage;
			ITranslator translator = TranslatorManager.CreateDisassembler(null, null);

            INamespace namespaceItem = new Namespace();
            namespaceItem.Name = typeDeclaration.Namespace;
 
            try
            {
                if (language.Translate)
                {
                    typeDeclaration = translator.TranslateTypeDeclaration(typeDeclaration, true, true);
                }
                namespaceItem.Types.Add(typeDeclaration);
            }
            catch (Exception ex)
            {
                Lib.Log(ex.Message);
            }

            TextFormatter formatter = new TextFormatter(strings, typeDeclaration.Name);
            ILanguageWriter writer = language.GetWriter(formatter, configuration);

            try
            {
                writer.WriteNamespace(namespaceItem);
            }
            catch (Exception exception)
            {
                Lib.Log(exception.Message);
            }

            translator = null;
            language = null; 

            return strings;

            //string output = formatter.ToString().Replace("\r\n", "\n").Replace("\n", "\r\n");


            //if (output.IndexOf("\"") != -1)
            //{
                
            //    string groupStr = "extract";
            //    string pattern = string.Format("[^\\\"]*\\\"(?'{0}'.*?)\\\"", groupStr);
            //    System.Text.RegularExpressions.MatchCollection matches = System.Text.RegularExpressions.Regex.Matches(output, pattern);


            //    int stringCounter = 0;
            //    foreach (System.Text.RegularExpressions.Match match in matches)
            //    {
            //        if (match.Success)
            //        {
            //            stringCounter++;
            //            string text = match.Groups[groupStr].Value;
            //            string id = string.Format("{0}.{1}.[string{2}]", typeDeclaration.Namespace, typeDeclaration.Name, stringCounter);

            //            strings.Add(id, text);
            //        }
            //    }
            //}



            //return strings;
        }

        

		private class LanguageWriterConfiguration : ILanguageWriterConfiguration
		{
			private IVisibilityConfiguration visibility = new VisibilityConfiguration();

			public IVisibilityConfiguration Visibility
			{
				get
				{
					return this.visibility;
				}
			}

			public string this[string name]
			{
				get
				{
					switch (name)
					{
						case "ShowDocumentation":
						case "ShowCustomAttributes":
						case "ShowNamespaceImports":
						case "ShowNamespaceBody":
						case "ShowTypeDeclarationBody":
						case "ShowMethodDeclarationBody":
							return "true";
					}

					return "false";
				}
			}
		}

		private class VisibilityConfiguration : IVisibilityConfiguration
		{
			public bool Public { get { return true; } }
			public bool Private { get { return true; } }
			public bool Family { get { return true; } }
			public bool Assembly { get { return true; } }
			public bool FamilyAndAssembly { get { return true; } }
			public bool FamilyOrAssembly { get { return true; } }
		}

		private class AssemblyResolver : IAssemblyResolver
		{
			private IDictionary _assemblyTable;
			private IAssemblyResolver _assemblyResolver;

			public AssemblyResolver(IAssemblyResolver assemblyResolver)
			{
				_assemblyTable = new Hashtable();
				_assemblyResolver = assemblyResolver;
			}

			public IAssembly Resolve(IAssemblyReference assemblyName, string localPath)
			{
				if (_assemblyTable.Contains(assemblyName))
				{
					return (IAssembly) _assemblyTable[assemblyName];
				}

                IAssembly assembly = null; // _assemblyResolver.Resolve(assemblyName, localPath);

				_assemblyTable.Add(assemblyName, assembly);

				return assembly;
			}
		}

    }
}
