
/*
 * Created by SharpDevelop.
 * User: PaulD
 * Date: 14/07/2010
 * Time: 17:25
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;

using System.IO;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

using NetSpell.SpellChecker;
using SpellCheck.Common;
using SpellCheck.Common.XmlData;

using SpellCheck.Common.Utility;

namespace SpellCheck.Plugin.SpellingEngine.NetSpellEngine
{
	/// <summary>
	/// Description of NetSpellEngine.
	/// </summary>
	
	[Serializable]
	public class Engine: SpellingEngineBase
	{
		
		static Spelling c_Spelling;
		static bool c_splitCamelCase;
        static bool c_splitUnderScore;
		static bool c_suggestions;
		static CultureInfo c_restrictedLocale;	
		
		static StringValue c_StringValue;
		static bool c_WaitEndOfText;
		
		static System.Collections.Specialized.StringCollection c_alertStrings;

		
		private string m_CitrixCustomDicName = "citrix";
		private string m_DicExtension = ".dic";


        public Engine(System.Collections.Specialized.StringCollection alertStrings, System.Globalization.CultureInfo locale, bool splitCamel, bool splitUnderScore, bool suggestions)
        {
         	c_restrictedLocale = locale;
        	c_splitCamelCase = splitCamel;
            c_splitUnderScore = splitUnderScore;
        	c_alertStrings = alertStrings;
        	c_suggestions = suggestions;
        	IntializeEngine();       
        }
        

        
        private void IntializeEngine()
        {
        	c_WaitEndOfText = true;
        	
        	c_Spelling = new Spelling();
        	c_Spelling.Dictionary = new NetSpell.SpellChecker.Dictionary.WordDictionary();
        	
			// stop dialog showing
			c_Spelling.ShowDialog = false;
			c_Spelling.AlertComplete = false; 
			c_Spelling.IgnoreWordsWithDigits = true;
			c_Spelling.IgnoreAllCapsWords = true;
			c_Spelling.MaxSuggestions = 1;
            c_Spelling.SplitCamelCase = c_splitCamelCase;
            c_Spelling.SplitUnderScore = c_splitUnderScore;
			
			// setup event handlers that receive speeling corrections
			c_Spelling.EndOfText += new Spelling.EndOfTextEventHandler(Spelling_EndOfText);
			c_Spelling.MisspelledWord += new Spelling.MisspelledWordEventHandler(Spelling_MisspelledWord);
			c_Spelling.DoubledWord += new Spelling.DoubledWordEventHandler(Spelling_DoubledWord);
			
		    string customCitrixDicPath;
            
            if((customCitrixDicPath = LoadCitrixCustomDictionary()) != null)
            {                
                //c_Spelling.Dictionary.UserFile = customCitrixDicPath;
            }
        }
		
        /// <summary>
        /// Implemneted method from base class. Handles checking the strings for this engine
        /// </summary>
        /// <param name="strVal"></param>
        /// <returns></returns>
		public override StringValue CheckString(StringValue strVal)
		{
			if (strVal.Value.Length > 0)
            {
                if (strVal.Locale == Locale.JA)
                {
                    Lib.Log("Japanese strings can't be checked with this tool [{0}]",strVal.Value);
                    return null;
                }

                if (c_restrictedLocale != null)
                {
                    if (c_restrictedLocale.TwoLetterISOLanguageName != strVal.Locale.Substring(0,2))
                    {
                        Lib.Log("***Skipping [{0}-{1}] string locale isn't {2}",strVal.Id,strVal.Locale,c_restrictedLocale);
                        return null;
                    }
                }

                c_Spelling.Dictionary.DictionaryFile = GetDictionaryFromLocale(strVal.Locale);


                c_Spelling.Text = strVal.Value;      

		        c_StringValue = strVal;
		        
		        Lib.ConsoleLog(" ->CHECKING: {0}", strVal.Value);
        		
		        if(c_alertStrings != null && c_alertStrings.Count > 0)
		        {
		        	CheckForAlertStrings(c_alertStrings,strVal.Value);
		        }
		        
		 
                while(c_WaitEndOfText)
                {
                	if(c_Spelling.SpellCheck(c_Spelling.WordIndex,c_Spelling.WordCount))
                	{
                		// the c_Spelling dialog seems to show up evrey so often so try setting this again
                		c_Spelling.ShowDialog = false;
                		c_Spelling.WordIndex ++;
                	}
                	else
                	{
                		c_WaitEndOfText = false;
                	}
                }
                
                c_WaitEndOfText = true;                
                 
                // check that the string is not too long for an excel cell
                c_StringValue = ExcelCheckString(c_StringValue);
                
                return c_StringValue ;

			}
			
			return null;
		}
		
		/// <summary>
		/// Check to see is speficied alert keywords exsist TODO FIXME etc
		/// </summary>
		/// <param name="alertStrings"></param>
		/// <param name="stringToCheck"></param>
		public override void CheckForAlertStrings(System.Collections.Specialized.StringCollection alertStrings, string stringToCheck)
		{
			foreach(string alertString in alertStrings)
			{
				if(stringToCheck.ToLower().Contains(alertString.ToLower()))
				{
					Error error = new Error(alertString, true, false);
					c_StringValue.Errors.Add(error);
				}
			}
		}
		
		/// <summary>
        /// Load the citrix specific word dictionary
        /// </summary>
        private string LoadCitrixCustomDictionary()
        {
        	FileInfo citrixCustomDir = new FileInfo(Path.Combine(DictionariesDirectory.FullName,string.Format("{0}{1}",m_CitrixCustomDicName,m_DicExtension)));
        	
        	if(citrixCustomDir.Exists)
        	{
        		return citrixCustomDir.FullName;
        	}
        	else
        	{
        		Lib.Log("Custom citrix dictionary not found: {0}", citrixCustomDir.FullName );
        	}
          
            return null;
        }
		
		/// <summary>
        /// Eventhandler for when a duplicate word is found
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Spelling_DoubledWord(object sender, SpellingEventArgs e)
        {
        	Spelling spelling = sender as Spelling;
        	
        	if(e.Word != null)
        	{
	        	string errorText = e.Word;
	        	c_StringValue.Errors.Add(new Error(errorText,false,true));
	        	
	        	Lib.ConsoleLog("Duplicate word found: {0}", errorText);
	        	
	        	if(c_StringValue.SuggestedCorrection == null)
	        	{
	        		c_StringValue.SuggestedCorrection = c_Spelling.Text;
	        	}
	        	
	        	c_StringValue.SuggestedCorrection = spelling.DeleteWord(new StringBuilder(c_StringValue.SuggestedCorrection),true);
        	}
        	else
        	{
        		Lib.Log("Error current word is NULL: Text={0} WordIndex={1} ",spelling.Text,spelling.WordIndex,spelling.CurrentWord);
        	}
        }
        

      
        /// <summary>
        /// Eventhandler for when an incorrect word is found
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Spelling_MisspelledWord(object sender, SpellingEventArgs e)
        {
        	Spelling c_Spelling = sender as Spelling;
        	
        	string errorText = e.Word;
        	c_StringValue.Errors.Add(new Error(errorText));
        	
        	Lib.ConsoleLog("Misspelled word found: {0}", errorText);
        	
        	if(c_suggestions)
        	{
	        	c_Spelling.Suggest();
        	}
        	
        	// replce error with first suggested correction
        	if(c_Spelling.Suggestions != null && c_Spelling.Suggestions.Count > 0)
        	{
	        	string correctedString = c_Spelling.Suggestions[0] as string;
	        	
	        	if(c_StringValue.SuggestedCorrection == null)
	        	{
	        		c_StringValue.SuggestedCorrection = c_Spelling.Text;
	        	}
	        	
	        	c_StringValue.SuggestedCorrection = c_StringValue.SuggestedCorrection.Replace(errorText, correctedString);
        	}
        	
        }

        /// <summary>
        /// Eventhandler to set end of text flag
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Spelling_EndOfText(object sender, EventArgs e)
        {
        	c_WaitEndOfText = false;
        }
        


		
		        /// <summary>
        /// Get the correct dictionary for the locale given
        /// </summary>
        /// <param name="locale"></param>
        /// <returns></returns>
        private string GetDictionaryFromLocale(string locale)
        {
        	DirectoryInfo dicDir = DictionariesDirectory;
        	
        	string dicName = string.Format("{0}{1}",locale,m_DicExtension);
        	
        	foreach(FileInfo finfo in dicDir.GetFiles())
        	{
        		if(finfo.Name.Equals(dicName,StringComparison.OrdinalIgnoreCase))
        		{
        			return finfo.FullName;
        		}
        	}
        	
        	throw new Exception(string.Format("Can not find locale dictionary file: {0}: ", Path.Combine(dicDir.FullName,dicName)));
        }
        
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
	}
}

