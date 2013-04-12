/*
 * Created by SharpDevelop.
 * User: PaulD
 * Date: 19/08/2010
 * Time: 08:56
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Runtime.Serialization;
using System.Collections.Generic;
using System.Xml;
using System.Xml.Serialization;
using System.Xml.Schema;
using System.Reflection;


namespace SpellCheck.Common.XmlData
{
		
	

		
		public class SpellingErrorData
		{
			
			public delegate void DataReceivedHandler(object sender, DataReceivedEventArgs args);
			public static event DataReceivedHandler DataReceivedEvent; 
			
			public static event EventHandler DataCompleteEvent;
			
			public class DataReceivedEventArgs : EventArgs
			{
	    			public SpellingErrorData SpellingErrorData;
			}			
			/// <summary>
	    	/// Column enum for spreadsheet columns
	    	/// </summary>
	    	public enum Columns
	    	{ 
	    		 [StringValue("File Handler Name")]
			     FileHandlerName = 	1,	
	    		 [StringValue("File Handler Description")]
				 FileHandlerDescription, 
	    		 [StringValue("File Path")]
				 FilePath, 
	    		 [StringValue("Error Total")]
				 ErrorTotal, 
	    		 [StringValue("String Id")]
				 StringId, 
	    		 [StringValue("Locale")]
				 Locale, 
	    		 [StringValue("Incorrect String")]
				 IncorrectString,
	    		 [StringValue("Suggested Correction")]
				 SuggestedCorrection, 
	    		 [StringValue("Contains Alert Word")]
				 Alert, 
				 [StringValue("Duplicate Word")]
				 Duplicate, 
	    		 [StringValue("Spelling Error")]
				 Error 
	    	}
	    	
	    	private class StringValueAttribute : System.Attribute
			{		
			    private string _value;
			
			    public StringValueAttribute(string value)
			    {
			        _value = value;
			    }
			
			    public string Value
			    {
			    get { return _value; }
			    }		
			}
	    	
	    	static IDictionary<Enum, StringValueAttribute> _stringValues = new Dictionary<Enum, StringValueAttribute>();
	    	
	    	public static string GetStringValue(Enum value)
			{
			    string output = null;
			    Type type = value.GetType();
			
			    //Check first in our cached results...
			
			    if (_stringValues.ContainsKey(value))
			      output = (_stringValues[value] as StringValueAttribute).Value;
			    else 
			    {
			        //Look for our 'StringValueAttribute' 
			
			        //in the field's custom attributes
			
			        FieldInfo fi = type.GetField(value.ToString());
			        StringValueAttribute[] attrs = 
			           fi.GetCustomAttributes(typeof (StringValueAttribute), 
			                                   false) as StringValueAttribute[];
			        if (attrs.Length > 0)
			        {
			            _stringValues.Add(value, attrs[0]);
			            output = attrs[0].Value;
			        }
			    }
			
			    return output;
			}
	    	
			
			
			public IDictionary<Columns, object> Data
			{
				get
				{
					if(data == null)
					{
						data = new Dictionary<Columns, object>();
					}
					
					return data;
				}
			}
			
			IDictionary<Columns, object> data;
			


			
			/// <summary>
			/// Populate a list data objects for each spelling error
			/// </summary>
			/// <param name="errors"></param>
			/// <returns></returns>
			public static void GetData(SpellingErrors errors)
			{
				IList<SpellingErrorData> dataList = new List<SpellingErrorData>();
					
				foreach(string handlerName in errors.FileHandlerTable.Keys)
    			{	
    				FileHandler handler = errors.FileHandlerTable[handlerName];
    				
    				foreach(FileData fileData in handler.Files)
    				{    					
    					foreach(StringValue strVal in fileData.GetStringValueList)
    					{    						
    						foreach(Error error in strVal.Errors)
    						{    							
								SpellingErrorData errorData = new SpellingErrorData();
								
								errorData.Data.Add(Columns.FileHandlerName, handlerName);								
								errorData.Data.Add(Columns.FileHandlerDescription,  handler.Description);
								errorData.Data.Add(Columns.FilePath, fileData.FilePath);
								errorData.Data.Add(Columns.ErrorTotal, fileData.ErrorTotal);
								errorData.Data.Add(Columns.Locale, strVal.Locale);
								errorData.Data.Add(Columns.StringId, strVal.Id);
								errorData.Data.Add(Columns.IncorrectString, Utility.Lib.UrlEncode( strVal.Value));
								errorData.Data.Add(Columns.SuggestedCorrection, Utility.Lib.UrlEncode(strVal.SuggestedCorrection));
								errorData.Data.Add(Columns.Alert, error.Alert);
								errorData.Data.Add(Columns.Duplicate, error.Duplicate);
								errorData.Data.Add(Columns.Error, error.Value);
								
								dataList.Add(errorData);
								SendData(errorData);
    						}    					
    					}    					
    				}    
    			}
				
				SendDataComplete();
				
			}
			
			private static void SendData(SpellingErrorData errorData)
			{			
				if(DataReceivedEvent != null)
				{
					DataReceivedEventArgs args = new SpellingErrorData.DataReceivedEventArgs();
					args.SpellingErrorData = errorData;
					DataReceivedEvent(null,args);					
				}
			}
			
			private static void SendDataComplete()
			{
				if(DataCompleteEvent != null)
				{
					DataCompleteEvent(null, new EventArgs());
				}
			}
		}

		[DataContract(Name = "SpellingErrors")]
		public partial class SpellingErrors
		{
			public SpellingErrors()
			{
				handlers = new FileHandlersDictionary<string, FileHandler>();
			}
	
			private FileHandlersDictionary<string, FileHandler> handlers;
	
			[DataMember(Name = "FileHandlers")]
			public FileHandlersDictionary<string, FileHandler> FileHandlerTable {
				get { return handlers; }
				set { handlers = value; }
			}
	
			private int errorTotal;
	
			public int ErrorTotal {
				get { return errorTotal; }
				set { errorTotal = value; }
			}
			
			
	
		}
		
		[CollectionDataContract
	    (Name = "FileHandlerName", 
	    ItemName = "FileHandler", 
	    KeyName = "FileHandlerName", 
	    ValueName = "FileHandlerData")]
		public class FileHandlersDictionary<TKey, TValue> : Dictionary<TKey, TValue> 
		{
	
		}
	
		[DataContract(Name = "FileHandler")]
		public partial class FileHandler
		{
	
			private string description;
	
			private List<FileData> files;
	
			public FileHandler()
			{
				files = new List<FileData>();
			}
	
			private int errorTotal;
	
			public int ErrorTotal {
				get { return errorTotal; }
				set { errorTotal = value; }
			}
	
			[DataMember(Order = 2, Name = "Files")]
			public List<FileData> Files {
				get { return files; }
				set { files = value; }
			}
	
			[DataMember(Order = 1, Name = "Description")]
			public string Description {
				get { return this.description; }
				set { this.description = value; }
			}
		}
	
	
		[Serializable]
		[DataContract(Name = "StringData")]
		public partial class StringValue
		{
			internal StringValue()
			{
			}
	
			public StringValue(string key, string value, System.Globalization.CultureInfo culture)
			{
				keyField = key;
				valueField = value;
				if (culture == null || culture.Equals(System.Globalization.CultureInfo.InvariantCulture)) {
					localeField = Utility.Locale.ENUS;
				} else {
					localeField = culture.ToString();
				}
			}
	
			private string valueField;
	
			private string correctedValueField;
	
			private List<Error> errorsField;
	
			private string keyField;
	
			private string localeField;
	
			[DataMember(Order = 1, Name = "StringId")]
			public string Id {
				get { return this.keyField; }
				set { this.keyField = value; }
			}
	
			[DataMember(Order = 2, Name = "Locale")]
			public string Locale {
				get { return localeField; }
				set { localeField = value; }
			}
	
			[DataMember(Order = 3, Name = "IncorrectString")]
			public string Value {
				get { return this.valueField; }
				set { this.valueField = value; }
			}
	
	
			[DataMember(Order = 4, Name = "SuggestedCorrection")]
			public string SuggestedCorrection {
				get { return this.correctedValueField; }
				set { this.correctedValueField = value; }
			}
	
			[DataMember(Order = 5, Name = "Errors")]
			public List<Error> Errors {
				get 
				{ 
					if(this.errorsField == null)
					{
						this.errorsField = new List<Error>();
					}
					return this.errorsField; 
				}
	//			set 
	//			{ 
	//				this.errorsField = value; 
	//			}
			}
	
	
		}
	
		[Serializable]
		[DataContract(Name = "Error")]
		public partial class Error
		{
	
			public Error(string errorText)
			{
				valueField = errorText;
				alertField = false;
			}
			
			public Error(string errorText, bool alert, bool duplicate)
			{
				valueField = errorText;
				alertField = alert;
				duplicateField = duplicate;
			}
			
			private string valueField;
			
			private bool alertField;
			private bool duplicateField;
	
			[DataMember(Name = "Error")]
			public string Value {
				get { return this.valueField; }
				set { this.valueField = value; }
			}
			
			[DataMember(Name = "Duplicate")]
			public bool Duplicate {
				get { return this.duplicateField; }
				set { this.duplicateField = value; }
			}

			
			[DataMember(Name = "Alert")]
			public bool Alert{
				get { return this.alertField; }
				set { this.alertField = value; }
			}
		}
	
		[Serializable]
		[DataContract(Name = "File")]
		public sealed class FileData 
		{
			[DataMember(Name = "Strings")]
			internal List<StringValue> StringValueList;
	
			public List<StringValue> GetStringValueList {
				get { return StringValueList; }
			}
	
			private int m_errorTotal;
	
			private string m_fileName;
	
			public FileData(string fileName)
			{
				StringValueList = new List<StringValue>();
				m_fileName = fileName;
			}
	
			[DataMember(Name = "ErrorTotal")]
			public int ErrorTotal {
				get { return m_errorTotal; }
				set { m_errorTotal = value; }
			}
	
	
			[DataMember(Name = "Name")]
			public string FilePath {
				get { return m_fileName; }
				set { m_fileName = value; }
			}
	
			public void AddString(StringValue sv)
			{
				StringValueList.Add(sv);
				m_errorTotal += sv.Errors.Count;
			}
	
		}

	
}