using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Text;
using System.Reflection;
using System.IO;
using System.Management;
using System.Text.RegularExpressions;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.Serialization;
using System.Xml;


using SpellCheck.Common.XmlData;
using SpellCheck.Common;
using SpellCheck.Common.Utility;

using SpellCheck.InternalPlugins;



namespace SpellCheck
{
    public class SpellCheck
    {        
        IDictionary<string, IList<IFileHandler>> _extensionDictionary;

        string _FileHandlerpluginName;
        string _FileHandlerpluginPattern = "SpellCheck.Plugin.*.dll";
        string _SpellCheckpluginPattern = "SpellCheck.Plugin.SpellingEngine.*.dll";
        string _inputFileOrDir;
        CultureInfo _locale;
        bool _doNotLoadExternalPlugins;
        bool _doNotLoadInternalFileHandlers;
        bool _notResXFileHandler;
        bool _notResourceFileHandler;
        bool _ShowResults;
        bool _SplitCamelCase = true;
        bool _SplitUnderScore = true;
        bool _suggestions = false;

        public bool DoNotLoadExternalPlugins
        {
            get { return _doNotLoadExternalPlugins; }
            set { _doNotLoadExternalPlugins = value; }
        }

        public bool DoNotLoadInternalFileHandlers
        {
            get { return _doNotLoadInternalFileHandlers; }
            set { _doNotLoadInternalFileHandlers = value; }
        }

        public bool NotResXFileHandler
        {
            get { return _notResXFileHandler; }
            set { _notResXFileHandler = value; }
        }

        public bool NotResourceFileHandler
        {
            get { return _notResourceFileHandler; }
            set { _notResourceFileHandler = value; }
        }

        public bool ShowResults
        {
            get { return _ShowResults; }
            set { _ShowResults = value; }
        }

        public bool SplitCamelCase
        {
            get { return _SplitCamelCase; }
            set { _SplitCamelCase = value; }
        }

        public bool SplitUnderScore
        {
            get { return _SplitUnderScore; }
            set { _SplitUnderScore = value; }
        }

        public bool Suggestions
        {
            get { return _suggestions; }
            set { _suggestions = value; }
        }

        Settings _settings;

        SpellingErrors _errors;

        Dictionary<string, FileHandler> _results;
        
        IList<ISpellingEngine> _spellingEngines;

        const string REPORTXMLEXT = "xml";

        /// <summary>
        /// Directory that holds the custom word dictionaries for each language
        /// </summary>
        private DirectoryInfo DictionariesDirectory
        {
            get
            {
            	DirectoryInfo dicDir = new DirectoryInfo(Path.Combine(Lib.WorkingDirectory.FullName, "dictionaries"));
            	if(dicDir.Exists)
            	{
	                return dicDir;
            	}
            	else
            	{
            		throw new DirectoryNotFoundException(string.Format("Directory {0} not found: can not load dictionaries",dicDir));
            	}
            }
        }

        /// <summary>
        /// Collection of file data 
        /// </summary>
        private Dictionary<string,FileHandler> Results
        {
            get
            {
                if (_results == null)
                {
                    _results = new Dictionary<string, FileHandler>();
                }
                return _results;
            }
        }

        /// <summary>
        /// Spelling errors object
        /// </summary>
        private SpellingErrors SpellingErrors
        {
            get
            {
                if (_errors == null)
                {
                    _errors = new SpellingErrors();
                }
                return _errors;
            }
        }

        StringCollection _alertStrings;

        public StringCollection AlertStrings
        {
            get 
            {
                if (_alertStrings == null)
                {
                    return _settings.AlertStrings;
                }

                return _alertStrings;
            }

            set
            {
                _alertStrings = value;
            }
        }

       

        public SpellCheck()
        {         
            //clean up extract dir before we start a new run
            if (Lib.ExtractionDirectory.Exists)
            {
                Lib.ExtractionDirectory.Delete(true);
                Lib.ExtractionDirectory = null;
            }

            _settings = new Settings();
            
            Lib.LogToConsole = _settings.LogToConsole;
            TraceLog _traceLOg = new TraceLog();
            _ShowResults = true;
        }

        public SpellingErrors Run(string[] args)
        {            
            string reportName = Lib.ReportName;

            try
            {
                // set local to EN-US as other language support doesn't work
            	_locale = new CultureInfo(Locale.ENUS);
                
                ProcessArgs(args);
                
                LoadSpellingEngines();

                LoadFileHandlerPlugins();

                HandleInputFileOrDirectory(_inputFileOrDir);

                WriteResult();
            }
            catch (Exception ex)
            {
                Lib.Log(ex.ToString());
                Lib.ConsoleLog(ex.Message);
            }

            return SpellingErrors;
        }
        
        /// <summary>
        /// At the moment this will load the first Engine found
        /// </summary>
        private void LoadSpellingEngines()
        {
			IDictionary<string, string> pluginDictionary = PluginManager.FindPlugins<ISpellingEngine>(Lib.WorkingDirectory.FullName, _SpellCheckpluginPattern);
			
			_spellingEngines = new List<ISpellingEngine>();	
			
			if(pluginDictionary != null)
			{
//				foreach(string assemblyPath in pluginDictionary.Keys)
//				{
					object[] parameters = new object[]{_settings.AlertStrings, _locale,_SplitCamelCase,_SplitUnderScore, _suggestions};
					_spellingEngines = PluginManager.LoadPlugins<ISpellingEngine>( pluginDictionary,parameters );
//				}
			}
        }

		/// <summary>
		/// Enumerate the cmdline args
		/// </summary>
		/// <param name="args"></param>
        private void ProcessArgs(string[] args)
        {
            for (int currentArg = 0; currentArg < args.Length; currentArg++)
            {
                string arg = args[currentArg];

                if (arg[0] == '-' || arg[0] == '/')
                {
                    switch (arg.Substring(1).ToLower(CultureInfo.InvariantCulture))
                    {
                        case "plugin":
                             currentArg++;
                            if (!File.Exists(Path.Combine(Lib.WorkingDirectory.FullName, args[currentArg])))
                            {
                                throw new FileNotFoundException("Can not find plugin", args[currentArg]);
                            }
                            _FileHandlerpluginName = args[currentArg];
                            break;

                        case "lang":
                            currentArg++;
                            string lang = args[currentArg].ToLower();

                            if (lang == Locale.EN)
                            {
                                _locale = new CultureInfo(Locale.ENUS);
                            }
                            else
                            {
                                _locale = new CultureInfo(lang);
                            }
                            break;

                        case "noplugins":
                            _doNotLoadExternalPlugins = true;
                            break;
                        case "notresx":
                            _notResXFileHandler = true;
                            break;
                        case "notresource":
                            _notResourceFileHandler = true;
                            break;
                        case "notbuiltinhandlers":
                            _doNotLoadInternalFileHandlers = true;
                            break;
                        case "notshowresults":
                            _ShowResults = false;
                            break;
                        case "notsplitbyunderscore":
                            _SplitUnderScore = false;
                            break;
                        case "notsplitcamelcase":
                            _SplitCamelCase = false;
                            break;
                        case "suggestions":
                            _suggestions = true;
                            break;
                        case "h":
                        case "help":
                        case "?":
                            ShowUsage();                            
                            break;
                        default:
                            Lib.ConsoleLog("Unknown switch '" + args[currentArg] + "'");
                            throw new ArgumentException("Unknown switch '" + args[currentArg] + "'");
                    }
                }
                else
                {

                    if (!Directory.Exists(arg) && !File.Exists(arg))
                    {
                        throw new IOException(string.Format("{0} is not a valid file or directory", arg));
                    }
                    
                    _inputFileOrDir = arg;

                }
            }

            if (_inputFileOrDir == null)
            {
                ShowUsage();
            }
        }

 

        private void ShowUsage()
        {
            StringBuilder builder = new StringBuilder();

            builder.Append("\n\r"); //suggestions
            builder.Append(string.Format("{0} <file or directory> [options]\n\r",Assembly.GetEntryAssembly().GetName().Name.ToUpper()));
            builder.Append("\n\r");
            builder.Append("-plugin <plugin name>       | name of only plugin to load.\n\r");
//            builder.Append("-lang                       | en|fr|es|de Restricts checking to langauage.\n\r");
            builder.Append("-noplugins                  | plugins will not load.\n\r");
            builder.Append("-notresx                    | resx file handler will not run.\n\r");
            builder.Append("-notresource                | embedded resource file handler will not run.\n\r");
            builder.Append("-notbuiltinhandlers         | resx and resource file handlers will not run.\n\r");
            builder.Append("-notshowresults             | stops results showing in default browser.\n\r");
            builder.Append("-notsplitbyunderscore       | do not split strings with _");
            builder.Append("-notsplitcamelcase          | do not split camelcase strings");
            builder.Append("-suggestions          		| get suggested corrections for spelling errors. This will slow down execution");
            builder.Append("-help                       | this help message.\n\r");
            Console.WriteLine(builder.ToString());

            Environment.Exit(0);
        }

  
        /// <summary>
        /// Take the cmd line argument and handle whether its a file or directory
        /// </summary>
        /// <param name="inputFileOrDir"></param>
        private void HandleInputFileOrDirectory(string inputFileOrDir)
        {
            if (Directory.Exists(inputFileOrDir))
            {
                HandleDirectory(new DirectoryInfo(inputFileOrDir));
            }
            else if (File.Exists(inputFileOrDir))
            {
                HandleFile(new FileInfo(inputFileOrDir));
            }
            else
            {
                throw new ApplicationException("Shouldn't reach here. Something odd has happened!");
            }
        }

        /// <summary>
        /// Get the files from the directory and all sub directories and handle each file
        /// </summary>
        /// <param name="directory"></param>
        private void HandleDirectory(DirectoryInfo directory)
        {
            foreach (string extension in _extensionDictionary.Keys)
            {
            	try
            	{
	                foreach (FileInfo file in directory.GetFiles(extension, SearchOption.AllDirectories))
	                {
	                    HandleFile(file);
	                }
            	}
            	catch(UnauthorizedAccessException uaex)
            	{
            		Lib.ConsoleLog("Access to directory failed: {0}", uaex.Message);
            		return;
            	}
            }
        }

        /// <summary>
        /// Get the file handler for the file extension 
        /// </summary>
        /// <param name="file"></param>
        /// <param name="extension"></param>
        private void HandleFile(FileInfo fileInfo)
        {
            string extension = string.Format("*{0}", fileInfo.Extension);

            if (!_extensionDictionary.Keys.Contains(extension))
            {
                Lib.ConsoleLog("No handler for file type [{0}]:  {1}", extension, fileInfo.FullName);
                return;
            }

            if(_settings.ExcludedFiles.Contains(fileInfo.Name))
            {
                Lib.ConsoleLog("File [{0}] has been excluded.",fileInfo.FullName);
                return;
            }
            
            if (!Lib.IsValidDrive(fileInfo))
            {
            	Lib.Log("Can't handle {0}. Drive is readonly. Copying to {1}", fileInfo, Lib.ExtractionDirectory.FullName);
                FileInfo copiedFile = fileInfo.CopyTo(Path.Combine(Lib.ExtractionDirectory.FullName, fileInfo.Name), true);
                copiedFile.IsReadOnly = false;
                HandleFile(copiedFile);
                // if copied don't process tyhe original file any further
                return; 
            }

            foreach (IFileHandler iFileHandler in _extensionDictionary[extension])
            {
                if (iFileHandler is IStringResourceFileHandler)
                {
                	IList<CheckStringDelegate> checkStringDelegates = new List<CheckStringDelegate>();
                	
                	foreach(ISpellingEngine engine in _spellingEngines)
                	{
                		checkStringDelegates.Add( new CheckStringDelegate(engine.CheckString) );
                	}
                	
                    IStringResourceFileHandler stringResourceFileHandler = iFileHandler as IStringResourceFileHandler;
                    Lib.ConsoleLog("{0}: {1}", stringResourceFileHandler,fileInfo);
                    ParseSpellingErrors(stringResourceFileHandler.LoadStrings(fileInfo, checkStringDelegates, _locale), stringResourceFileHandler);
                }
                else if (iFileHandler is IArchiveFileHandler)
                {
//                    if (Lib.IsValidDrive(fileInfo))
//                    {
                        Lib.Log("Expanding {0} using {1}",fileInfo.FullName, iFileHandler);
                        IArchiveFileHandler archiveFileHandler = iFileHandler as IArchiveFileHandler;
                        DirectoryInfo extractDirectory = archiveFileHandler.Extract(fileInfo);

                        HandleDirectory(extractDirectory);
//                    }
//                    else
//                    {
//                        Lib.Log("Can not extract {0}. Drive is not valid. Copying to {1}", fileInfo, Lib.ExtractionDirectory.FullName);
//                        FileInfo copiedFile = fileInfo.CopyTo(Path.Combine(Lib.ExtractionDirectory.FullName, fileInfo.Name), true);
//                        HandleFile(copiedFile);
//                    }
                }

                else
                {
                    // should not happen
                    throw new Exception("Unknown file handler found");
                }
            }
        }




        /// <summary>
        /// Load internal and external plugins
        /// </summary>
        private void LoadFileHandlerPlugins()
        {
            List<IFileHandler> spellCheckFileHandlers = new List<IFileHandler>();

            if (!_doNotLoadExternalPlugins)
            {
                if (_FileHandlerpluginName == null)
                {
                    _FileHandlerpluginName = _FileHandlerpluginPattern;
                }

                IDictionary<string, string> pluginDictionary = PluginManager.FindPlugins<IFileHandler>(Lib.WorkingDirectory.FullName, _FileHandlerpluginName);
                
                if(pluginDictionary != null)
                {
	                spellCheckFileHandlers.AddRange(PluginManager.LoadPlugins<IFileHandler>(pluginDictionary));
                }
                
                
            }

            LoadInternalPlugins(spellCheckFileHandlers);

            _extensionDictionary = new Dictionary<string, IList<IFileHandler>>();

            foreach (IFileHandler handler in spellCheckFileHandlers)
            {
                foreach (FileExtension extension in handler.SupportedExtensions)
                {
                    if (_extensionDictionary.ContainsKey(extension.StarDotExtension))
                    {
                        _extensionDictionary[extension.StarDotExtension].Add(handler);
                    }
                    else
                    {
                        IList<IFileHandler> handlers = new List<IFileHandler>();
                        handlers.Add(handler);
                        _extensionDictionary.Add(extension.StarDotExtension, handlers);
                    }
                }
            }

        }

        /// <summary>
        /// Load the internal Plugins for Embedded .net resources, resx
        /// Cab extracters and Msi
        /// </summary>
        /// <param name="spellCheckFileHandlers"></param>
        private void LoadInternalPlugins(List<IFileHandler> spellCheckFileHandlers)
        {
            if (!_doNotLoadInternalFileHandlers)
            {
                if (!_notResourceFileHandler)
                {
                    AssemblyPlugin assFileHandler = new AssemblyPlugin();
                    spellCheckFileHandlers.Add(assFileHandler);
                }
                if (!_notResXFileHandler)
                {
                    ResXPlugin resxFileHandler = new ResXPlugin();
                    spellCheckFileHandlers.Add(resxFileHandler);
                }

                CabExtractPlugin cabFileHandler = new CabExtractPlugin();
                MSIExtractPlugin msiFileHandler = new MSIExtractPlugin();

                spellCheckFileHandlers.Add(cabFileHandler);
                spellCheckFileHandlers.Add(msiFileHandler);
            }

        }

 
        /// <summary>
        /// Iterate through each spell check result, and store it by handler
        /// </summary>
        /// <param name="data"></param>
        private void ParseSpellingErrors(FileData data,IStringResourceFileHandler iStringResourceFileHandler)
        {
            if (data != null)
            {
                if (data.GetStringValueList.Count > 0)
                {
                    if (!SpellingErrors.FileHandlerTable.ContainsKey(iStringResourceFileHandler.StrongName))
                    {
                        FileHandler handler = new FileHandler();
                        handler.Description = iStringResourceFileHandler.Description;
                        handler.Files.Add(data);
                        handler.ErrorTotal += data.ErrorTotal;
                        SpellingErrors.FileHandlerTable.Add(iStringResourceFileHandler.StrongName, handler);
                        SpellingErrors.ErrorTotal += data.ErrorTotal;
                    }
                    else
                    {
                        FileHandler handler = SpellingErrors.FileHandlerTable[iStringResourceFileHandler.StrongName];
                        handler.Files.Add(data);
                        handler.ErrorTotal += data.ErrorTotal;
                        SpellingErrors.ErrorTotal += data.ErrorTotal;
                    }
                } 
            }
        }

        /// <summary>
        /// write the spelling errors to xml using serialisation and
        /// excel using OpenOfficeXML
        /// </summary>
        private void WriteResult()
        {
            if (SpellingErrors.ErrorTotal > 0)
            {
                int spellingErrorCount = SpellingErrors.ErrorTotal;

                DirectoryInfo resultPath = Lib.ReportPath;

                string reportname = Lib.ReportName+ ".{0}";                
                               
                string resultFilePath = Path.Combine(resultPath.FullName, reportname);
                
                string xmlResultFilePath = string.Format(resultFilePath,REPORTXMLEXT);

            	using(MemoryStream ms = new MemoryStream())
            	{
            		// write spelling errors to a memory stream
                    DataContractSerializer serializer = new DataContractSerializer(typeof(SpellingErrors));
                    serializer.WriteObject(ms, SpellingErrors);
                    
		            ms.Position = 0;
		            
		            // get the xml string and create an xmldocument
		            string xmlStr = new StreamReader(ms).ReadToEnd();
	                XmlDocument xmlObj = new XmlDocument();
		            xmlObj.LoadXml(xmlStr);
					
		            // add an xmlattribute to point to the schema.
		            // this done to stop Excel from complaing that mo schema is referenced
					XmlElement e = xmlObj.DocumentElement;
					
					// not sure if this is really needed but it makes the xml look cleaner
					// remove xmlns:i="http://www.w3.org/2001/XMLSchema-instance" xmlattribute
					
					string XMLSchemaInstance = null;
									
					if(e.HasAttribute("xmlns:i"))
					{
						XMLSchemaInstance = e.GetAttribute("xmlns:i");
						e.RemoveAttribute("xmlns:i");
					}
					
					// work out the schema name from the namespace
					string nameSpace = e.GetAttribute("xmlns");					
					string[] arr = nameSpace.Split(new char[]{'/'});						
					string xsdName = arr[arr.Length -1] + ".xsd";					   
					
					e.SetAttribute("schemaLocation", XMLSchemaInstance , string.Format("{0} {1}",nameSpace, xsdName));					

					// save the xmldoc to the results xml file
					xmlObj.Save(xmlResultFilePath);
					
					
					
                    string xlsxPath = XmlToExcel.SaveToXlsx(SpellingErrors, resultFilePath);
                    Lib.ConsoleLog("Saved to [{0}]]", xmlResultFilePath);
                    Lib.ConsoleLog("Saved to [{0}]]", xlsxPath);
                    
	                if (_ShowResults)
	                {
                        Lib.ConsoleLog("Loading results in default XML browser");
	                    System.Diagnostics.Process.Start(xmlResultFilePath);
	                }
            	}
            }
            else
            {
                Lib.ConsoleLog("No spelling errors found");                
            }
        }

        /// <summary>
        /// Save the result XML to Excel
        /// **************not used any more*****************
        /// </summary>
        /// <param name="pathToXmlFile"></param>
        private void SaveXmlToExcel(string pathToXmlFile)
        {
            string assemblyName = "XmlToExcel";
            string typeName = "XmlToExcel.XmlToExcelClass";
            string GetMethodName = "SaveToXlsx";
            object[] parameters = new object[] { pathToXmlFile };
            FileInfo pathToXmlFileInfo = new FileInfo(pathToXmlFile);

            try
            {
                object returnObj = Lib.RunMethodByReflection(assemblyName, typeName, GetMethodName, parameters);
                Lib.ConsoleLog("Saved [{0}] to [{1}]]", pathToXmlFileInfo.Name, returnObj);
            }
            catch (Exception ex)
            {
            	if( ex.InnerException == null)
            	{
	                Lib.Log("{0}", ex.Message);
            	}
            	else
            	{
	                Lib.Log("{0} [{1}]", ex.Message, ex.InnerException.Message);
            	}
                Lib.ConsoleLog("***Failed to load Excel");
               
            }

        }

    }

}
