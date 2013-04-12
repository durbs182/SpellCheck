// ---------------------------------------------------------
// Lutz Roeder's .NET Reflector
// Copyright (c) 2000-2005 Lutz Roeder. All rights reserved.
// http://www.aisto.com/roeder
// ---------------------------------------------------------
namespace Reflector
{
	using System;
	using System.Collections;
	using System.Globalization;
	using System.IO;
    using System.Collections.Generic;
	using Reflector.CodeModel;
	
	public class TextFormatter : IFormatter
	{
		//private StringWriter writer = new StringWriter(CultureInfo.InvariantCulture);
		//private bool allowProperties = false;
		//private bool newLine = false;
		//private int indent = 0;

        //const string CLASS = "class";
        //const string STRUCT  = "struct";

        //bool foundClassKeyword;
        //bool foundClassName;
        string m_className;
        Dictionary<string,string> m_strings;



        public TextFormatter(Dictionary<string, string> strings, string className)
        {
            m_strings = strings;
            m_className = className;
        }

        private int lineNumber = 0;

        public override string ToString()
        {
            return m_className;
        }
	
		public void Write(string text)
		{
			//this.ApplyIndent();
			//this.writer.Write(text);
        }
	
		public void WriteDeclaration(string text)
		{
			//this.WriteBold(text);

            //if (foundClassKeyword && !foundClassName)
            //{
            //    foundClassName = true;
            //    m_className = text;
            //}
		}

        public void WriteDeclaration(string text, object target)
        {
            //this.WriteBold(text);
            //if (foundClassKeyword && !foundClassName)
            //{
            //    foundClassName = true;
            //    m_className = text;
            //}
        }

        public void WriteComment(string text)
		{
			//this.WriteColor(text, (int) 0x808080);
  		}
		
		public void WriteLiteral(string text)
		{
			//this.WriteColor(text, (int) 0x800000);

            
            if (text.IndexOf("\"") != -1)
            {
                //text = text.Replace("\"",string.Empty).Trim();
                //text = text.Replace("."," ");
                //text = text.Replace(";", " ");
                //text = text.Replace(":", " ");
                //text = text.Replace("|", " ");

                if (text != string.Empty)
                {
                    string key = string.Format("{0}.Line[{1}]", m_className, lineNumber);
                    Console.WriteLine("{0} [{1}]",key,text);
       
                    if (m_strings.ContainsKey(key))
                    {
                        string val = m_strings[key];

                        val = string.Format("{0} {1}", val, text);

                        m_strings[key] = val;
                    }
                    else
                    {
                        m_strings.Add(key, text);
                    }
                }
            }
		}
		
		public void WriteKeyword(string text)
		{
            //if ((text == CLASS || text == STRUCT) && !foundClassKeyword)
            //{
            //    foundClassKeyword = true;
            //}
		}
	
		public void WriteIndent()
		{
			//this.indent++;
		}
				
		public void WriteLine()
		{
			//this.writer.WriteLine();
			//this.newLine = true;
            lineNumber++;
		}

		public void WriteOutdent()
		{
			//this.indent--;
		}

		public void WriteReference(string text, string toolTip, Object reference)
		{
			//this.ApplyIndent();
			//this.writer.Write(text);
		}

		public void WriteProperty(string propertyName, string propertyValue)
		{
            //if (this.allowProperties)
            //{
            //    throw new NotSupportedException();
            //}
		}

        //public bool AllowProperties
        //{

			
        //    get 
        //    { 
        //        return false;
        //    }
        //}

        //private void WriteBold(string text)
        //{
        //    this.ApplyIndent();
        //    this.writer.Write(text);
        //}

        //private void WriteColor(string text, int color)
        //{
        //    this.ApplyIndent();
        //    this.writer.Write(text);
        //}

        //private void ApplyIndent()
        //{
        //    if (this.newLine)
        //    {
        //        for (int i = 0; i < this.indent; i++)
        //        {
        //            this.writer.Write("    ");
        //        }

        //        this.newLine = false;
        //    }
        //}
	}
}