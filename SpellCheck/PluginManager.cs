using System;
using System.Collections.Generic;
using System.Reflection;
using System.IO;

using SpellCheck.Common.XmlData;
using SpellCheck.Common;
using SpellCheck.Common.Utility;


namespace SpellCheck
{
    internal class PluginManager : MarshalByRefObject
    {
    	/// <summary>
    	/// Check that plugins implement the correct interfaces for the type spefcied
    	/// Then unload the assemblies in case there are in correct plugins found
    	/// </summary>
    	/// <param name="path"></param>
    	/// <param name="pattern"></param>
    	/// <returns></returns>
        internal static IDictionary<string, string> FindPlugins<T>(string path, string pattern)
        {
            AppDomain sandbox = null;    
            
            try
            {
				//PluginManager pluginManager = Lib.CreateObjectInstance<PluginManager>(out sandbox); 

                PluginManager pluginManager = Citrix.Automation.ObjectCreator.LocalObjectFactory<PluginManager>.Instance.GetObject(false);
				
				return  pluginManager.SearchForPlugIns( typeof(T), path, pattern);
 				
            }
            catch (System.Exception systemException)
            {
                Lib.Log(systemException.ToString());
                return null;
            }
            finally
            {
                Citrix.Automation.ObjectCreator.IsolatedObjectManager.Instance.DisposeAll();
            }            
        }


        internal static IList<T_PluginType> LoadPlugins<T_PluginType>(IDictionary<string, string> pluginDictionary)
        {
        	IList<T_PluginType> pluginList = LoadPlugins<T_PluginType>(pluginDictionary,null);
        
        	return pluginList;
        }
        
        
        internal static IList<T_PluginType> LoadPlugins<T_PluginType>(IDictionary<string, string> pluginDictionary, object[] parameters)
        {
	        IList<T_PluginType> pluginList = new List<T_PluginType>();
        	
        	foreach(string assemblyPath in pluginDictionary.Keys)
        	{
        		Assembly assembly = Assembly.LoadFrom(assemblyPath);
        		
        		
        		
        		if(parameters != null)
        		{
        			Type type = assembly.GetType(pluginDictionary[assemblyPath]);
        		
        			Type[] paramType = new Type[parameters.Length];
	        		for(int i=0;i<parameters.Length;i++)
	        		{
	        			paramType[i] = parameters[i].GetType();
	        		} 
	        		ConstructorInfo cinfo = type.GetConstructor(paramType);
	        		
	        		foreach(ConstructorInfo ci in type.GetConstructors())
	        		{
	        			CallingConventions cc = ci.CallingConvention;
	        			ParameterInfo[] p = ci.GetParameters();
	        		}
	        		
	        		T_PluginType obj = (T_PluginType) cinfo.Invoke(parameters);
	        		pluginList.Add(obj);        		
        		}
        		else
        		{
        			T_PluginType obj = (T_PluginType) assembly.CreateInstance(pluginDictionary[assemblyPath]);
        			pluginList.Add(obj);
        		}
        		
        		
               
        	}
        
        	return pluginList;
        }    
        

        private IDictionary<string, string> SearchForPlugIns(Type searchType, string path, string pattern)
        {
        	Lib.Log("Executing {0} in AppDomain: {1}", Lib.GetMethodName(), AppDomain.CurrentDomain.FriendlyName);

            IDictionary<string, string> pluginDictionary = new Dictionary<string, string>();

            foreach (string assemblyStr in Directory.GetFiles(path, pattern))
            {
                Assembly assembly = Assembly.LoadFile(assemblyStr);

                foreach (Module mod in assembly.GetModules())
                {
                    foreach (Type type in mod.GetTypes())
                    {
                        if (type.GetInterface(searchType.FullName) != null)
                        {
                            byte[] token = assembly.GetName().GetPublicKeyToken();
                            if (token == null || token.Length == 0)
                            {
                                Lib.Log("{0} SpellCheck Plugins must have a strongname: {1}",Lib.GetMethodName(), assembly.FullName);
                                continue;
                            }     
                            
                            Lib.Log("Found PlugIn: " + type);
                            
                            if(pluginDictionary.ContainsKey(assembly.Location))
                            {
                            	Lib.Log("{0}: SpellCheck Plugins can contain only 1 plug-in type: {1}",Lib.GetMethodName(), assembly.FullName);
                            	throw new NotSupportedException("Assembly can contain only 1 plug-in type");
                            }
                            else
                            {
	                            pluginDictionary.Add(assembly.Location,type.FullName);
                            }

                        }
                    }
                }
            }

            return pluginDictionary;
        }

    }
}
