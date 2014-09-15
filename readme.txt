
AUTOSPELLCHECK <file or directory> [options]

-plugin <plugin name>       | name of only plugin to load.
-lang                       | en   ***** Not working|fr|es|de Restricts checking to langauage.**********
-noplugins                  | plugins will not load.
-notresx                    | resx file handler will not run.
-notresource                | embedded resource file handler will not run.
-notbuiltinhandlers         | resx and resource file handlers will not run.
-notshowresults             | stops results showing in default browser.
-notsplitcamelcase          | do not split camelcase strings
-help                       | this help message.

File Types
==========
Files processed by inbuilt handlers

ArchiveHandlers
msi - extracts all files from an MSI installer
cab - extracts all files from a cab file

StringHandlers
dll or exe (.Net) - extracts strings from embedded resource file
resx - extracts strings from resx files.




Plugins
=======
MSIStrings.plugin - extracts MSI strings. Use the columns.xml file to specify the columns in which to search for strings. There are some already included but you might need to add to the list. Use Orca to help with this.
HelpCHMHandler.plugin - decomposes Chm files to html, removes html tags and spellchecks what is left.
TextFileHandler.plugin - checks htm/html/xml/rtf/xml/xlst files
Win32BinaryStrings.plugin - check win32 dll and exe. Looks for dialogs, stringtables and messagetables

RCStrings.Plugin - parses Win32 .rc resource files


Custom Dictionary
===================
Add to citrix.dic to avoid
false positive errors to include in a customdictionary



3rd party code
==============
This utility makes use of EPPLUS, Microsoft.Cci, wix, ICSharpCode.SharpZipLib, NetSpell.SpellChecker and HtmlAgilityPack


Creating addtional Plugins
==========================
2 types of plugin can be created 

IArchiveFileHandler - used for files that contain other files (i.e. cab / msi / zip etc)

IStringResourceFileHandler - files that contain strings that need spell checking.

Plugin file name must be <name>.plugin.dll. The assembly must have a strong name.

Just reference SpellCheck.Common in your solution and look in the SpellCheck.Common & SpellCheck.Common.Utility namespaces.


