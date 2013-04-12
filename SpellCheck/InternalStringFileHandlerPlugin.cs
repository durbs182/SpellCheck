using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Globalization;

using System.Collections.Specialized;
using System.Diagnostics;
using System.Windows.Forms;


using SpellCheck.Common.XmlData;
using SpellCheck.Common;

using SpellCheck.Common.Utility;

//using Microsoft.Tools.WindowsInstallerXml.Cab;


namespace SpellCheck
{
	[Serializable]
	public abstract class InternalStringFileHandlerPlugin : StringFileHandler
	{
	    string[] exlcusionList = new string[] { ".type", ".parent", ".name", ".zorder", ".Font", ".ImeMode", ".StartPosition", ".BackgroundImage", ".TextAlign" };
	    
	    internal bool ExcludedKey(string key)
	    {
	        foreach (string exlString in exlcusionList)
	        {
	            if (key.EndsWith(exlString, StringComparison.OrdinalIgnoreCase))
	                return true;
	        }
	        return false;
	    }
	
	    internal string ParseString(string str)
	    {
	        
	        // rtf strings?
	        if (str.StartsWith(@"{\rtf1\ansi\ansicpg1252\"))
	        {
	            using (RichTextBox rtb = new RichTextBox())
	            {
	                rtb.Rtf = str;
	                str = rtb.Text;
	            }
	        }
	
	        // xml strings?
	        if (str.StartsWith(@"<?xml"))
	        {
	            str = Lib.StripTags(str);
	        }
	
	        // strip out & 
	        str = str.Replace("&", string.Empty);
	
	        // strip out ;
	        str = str.Replace(";", " ");
	        str = str.Replace(":", " ");
	        str = str.Replace("|", " ");
	
	        // strip white space
	        str = str.Trim();
	        return str;
	    }
	
	
	}
}
