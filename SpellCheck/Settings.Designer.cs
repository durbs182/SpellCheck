﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.17020
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace SpellCheck {
    
    
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "10.0.0.0")]
    internal sealed partial class Settings : global::System.Configuration.ApplicationSettingsBase {
        
        private static Settings defaultInstance = ((Settings)(global::System.Configuration.ApplicationSettingsBase.Synchronized(new Settings())));
        
        public static Settings Default {
            get {
                return defaultInstance;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute(@"<?xml version=""1.0"" encoding=""utf-16""?>
<ArrayOfString xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">
  <string>msvcm80.dll</string>
  <string>mfcm80.dll</string>
  <string>mfcm80u.dll</string>
  <string>msvcp80.dll</string>
  <string>msvcr80.dll</string>
  <string>msvcm90.dll</string>
  <string>msvcp90.dll</string>
  <string>msvcr90.dll</string>
  <string>cmx8.dll</string>
  <string>mfcm90.dll</string>
  <string>mfcm90u.dll</string>
</ArrayOfString>")]
        public global::System.Collections.Specialized.StringCollection ExcludedFiles {
            get {
                return ((global::System.Collections.Specialized.StringCollection)(this["ExcludedFiles"]));
            }
            set {
                this["ExcludedFiles"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("True")]
        public bool LogToConsole {
            get {
                return ((bool)(this["LogToConsole"]));
            }
            set {
                this["LogToConsole"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("<?xml version=\"1.0\" encoding=\"utf-16\"?>\r\n<ArrayOfString xmlns:xsi=\"http://www.w3." +
            "org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\">\r\n  <s" +
            "tring>vdesk</string>\r\n  <string>ringcube</string>\r\n</ArrayOfString>")]
        public global::System.Collections.Specialized.StringCollection AlertStrings {
            get {
                return ((global::System.Collections.Specialized.StringCollection)(this["AlertStrings"]));
            }
            set {
                this["AlertStrings"] = value;
            }
        }
    }
}
