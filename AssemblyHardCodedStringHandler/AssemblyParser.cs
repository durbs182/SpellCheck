using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Cci;
using System.IO;

using System.Text.RegularExpressions;

using SpellCheck.Common;
using SpellCheck.Common.XmlData;

using SpellCheck.Common.Utility;

namespace SpellCheck.Plugin.AssemblyHardCodedHandler
{
	
	public class AssemblyParser :MarshalByRefObject
    {

        //static void Main(string[] args)
        //{
        //    new Program();
        //    GC.Collect();
        //}

        static List<StringValue> c_strings;

        static AssemblyParser()
        {
            c_strings = new List<StringValue>();
        }
        
        internal static Dictionary<string, List<StringValue>> RunInSandbox(System.IO.FileInfo fileInfo)
        {
        	AppDomain sandbox = null;
        	
        	try{
        		
        		AssemblyParser parser = Lib.CreateObjectInstance<AssemblyParser>(out sandbox); 
        		
        		Dictionary<string, List<StringValue>> stringsDic = new Dictionary<string, List<StringValue>>();
                
                return parser.Run(fileInfo, stringsDic);
        	}
        	catch (System.Exception systemException)
            {
                Lib.Log(systemException.ToString());
                throw;
            }
            finally
            {
                AppDomain.Unload(sandbox);

                GC.Collect();
            }
        }

        internal Dictionary<string, List<StringValue>> Run(FileInfo fileInfo, Dictionary<string, List<StringValue>> stringsDic)
        {
            try
            {
            	Lib.Log("Executing {0} in AppDomain: {1}", Lib.GetMethodName(), AppDomain.CurrentDomain.FriendlyName);

                using (AssemblyNode assNode = AssemblyNode.GetAssembly(fileInfo.FullName))
                {
                    if (assNode != null)
                    {
                        if (stringsDic.ContainsKey(assNode.StrongName))
                        {
                            Lib.Log("{0} has already been checked", assNode.StrongName);
                            return stringsDic;
                        }

                        stringsDic.Add(assNode.StrongName, c_strings);

                        TypeNodeList typeLIst = assNode.Types;

                        foreach (TypeNode typeNode in typeLIst)
                        {
                            CheckType(typeNode);
                        }
                    }
                    else
                    {
                        Lib.Log("Failed to load: {0}", fileInfo.FullName);
                        return stringsDic;
                    }
                }

            }
            catch(System.TypeInitializationException typeEx)
            {
                Lib.Log("Failed to load: {0}", fileInfo.FullName);
                Lib.Log(typeEx.StackTrace);
            }
            
            return stringsDic;

        }

        /// <summary>
        /// Parse each type and handle Class attributes and each method (properties as well) and field 
        /// </summary>
        /// <param name="typeNode"></param>
        private void CheckType(TypeNode typeNode)
        {
            ParseAttributes(typeNode.Name.Name, typeNode.Attributes);

            foreach (Member member in typeNode.Members)
            {
                if (member is Field)
                {
                    ParseField(member as Field);
                    continue;
                }

                if (member is Method)
                {
                    ParseMethod(member as Method);
                    continue;
                }
            }
        }

        /// <summary>
        /// Pull out the strings in Attributes 
        /// </summary>
        /// <param name="id"></param>
        /// <param name="attributeList"></param>
        private void ParseAttributes(string id, AttributeList attributeList)
        {
            foreach (AttributeNode attNode in attributeList)
            {
                if (attNode.Type.Name.Name == "AttributeUsageAttribute")
                { }
                int expNum = 0;
                foreach (Expression exp in attNode.Expressions)
                {
                    if (exp is NamedArgument && exp.Type.TypeCode == TypeCode.String)
                    {
                        NamedArgument arg = exp as NamedArgument;

                        string str = arg.Value.ToString();

                        if ((str = CheckString(str)) != null)
                        {
                            //Console.WriteLine("{0}.{1} [{2}]", id, attNode.Type.Name, str);
                            string itemID = string.Format("{0}.{1}.{2}.[{3}]", id, attNode.Type.Name, attNode.UniqueKey, expNum);
                            c_strings.Add(new StringValue(itemID, str, System.Globalization.CultureInfo.InvariantCulture));


                            expNum++;
                        }
                        continue;
                    }

                    if (exp.Type.TypeCode == TypeCode.String && exp is Literal)
                    {
                        Literal lit = exp as Literal;
                       
                        if (lit.Value != null)
                        {
                            string str = lit.Value.ToString();

                            if ((str = CheckString(str)) != null)
                            {
                                string itemID = string.Format("{0}.{1}.{2}.[{3}]", id, attNode.Type.Name, attNode.UniqueKey, expNum);

                                c_strings.Add(new StringValue(itemID, str, System.Globalization.CultureInfo.InvariantCulture));

                                expNum++;
                            }
                        }
                 
                        continue;
                    }
                }
            }
        }



        /// <summary>
        /// Look for the deafult value in a field if its a string
        /// </summary>
        /// <param name="field"></param>
        private void ParseField(Field field)
        {
            if (field.DefaultValue != null && field.DefaultValue.Type.TypeCode == TypeCode.String)
            {
                string str = field.DefaultValue.ToString();

                if ((str = CheckString(str)) != null)
                {
                    string id = string.Format("{0}.{1}", field.DeclaringType.FullName, field.Name.Name);

                    c_strings.Add(new StringValue( id, str,System.Globalization.CultureInfo.InvariantCulture));
                }
            }
        }

        /// <summary>
        /// Regex delegate
        /// </summary>
        /// <param name="m"></param>
        /// <returns></returns>
        private string RemoveChar(Match m)
        {
            string s = m.Captures[0].Value;
            return string.Empty;
        }


        /// <summary>
        /// Lots of hard codes strings aren't real words and won't be shown in the UI.
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        private string CheckString(string str)
        {
            string retStr = null;
            bool strOk = true;
            str = str.Trim();

            if (str == string.Empty)
            {
                strOk = false;
            }

            // if the string doesn't have any spaces it my be an internal string so check further
            if (strOk && str.IndexOf(" ") == -1)
            {
                strOk = Lib.WordContainsOnlyAlpha(str);

                if (strOk)
                {
                    strOk = !Lib.IsCamelCase(str);
                }
            }
            else
            {

                if (strOk)
                {
                    //remove ()[]{} from string
                    str = Regex.Replace(str, @"([{}\]\[\(\)]*)", new MatchEvaluator(RemoveChar), RegexOptions.Compiled);

                    string[] words = str.Split(new char[] { ' ' });
                    StringBuilder sb = new StringBuilder();
                    foreach (string word in words)
                    {
                        string subStr = word;
                        while (Regex.Match(subStr, @"([;:,.!?]$)", RegexOptions.Compiled).Success)
                        {
                            subStr = subStr.Substring(0, subStr.Length - 1);
                        }
                        if (Lib.WordContainsOnlyAlpha(subStr) && !Lib.IsCamelCase(subStr))
                        {
                            sb.Append(subStr + " ");
                        }
                    }

                    retStr = sb.ToString().Trim();
                    if (retStr == string.Empty)
                    {
                        strOk = false;
                    }
                }
            }

            if (strOk)
            {
                if (retStr == null)
                {
                    retStr = str;
                }
            }
            else
            {
                Lib.Log("{0} excluding -> [{1}]",Lib.GetMethodName(), str);
                return null;
            }

            return retStr;

        }

        /// <summary>
        /// Is the string CamelCase or camelCase if true its not a real word so ignore
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
//        private bool IsNotCamelCase(string str)
//        {
//            // if this is all lowercase skip
//            if (!str.Equals(str.ToLower(), StringComparison.Ordinal) && !str.Equals(str.ToUpper(), StringComparison.Ordinal))
//            {
//                //same
//                MatchCollection m = Regex.Matches(str, @"(\b[a-z]|\B[A-Z])", RegexOptions.Compiled);
//
//                if (m.Count > 0)
//                {
//                    //is cammelCase
//                    return false;
//                }
//            }
//            return true;
//        }

        /// <summary>
        /// If the word has non alpha charaters ignore it strips out $Name, \\path\path\ or //file/file etc
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
//        private bool WordContainsOnlyAlpha(string str)
//        {
//            if (Regex.Match(str, @"([^a-zA-Z])", RegexOptions.Compiled).Success)
//            {
//                if (str.Equals("."))
//                {
//                    return false;
//                }
//
//                if (!str.EndsWith(".") && str.IndexOf("'") == -1)
//                {
//                    return false;
//                }
//            }
//            return true;
//        }

        /// <summary>
        /// Check the Method attributes for strings and walk the Opcodes and look for OpCode.Ldstr
        /// </summary>
        /// <param name="method"></param>
        private void ParseMethod(Method method)
        {
      
            string id = string.Format("{0}.{1}.{2}", method.DeclaringType.FullName, method.Name.Name,method.UniqueKey);

            if (method.DeclaringMember is Property)
            {
                AttributeList attributeList = method.DeclaringMember.Attributes;
                ParseAttributes(id, attributeList);
            }
            else
            {
                AttributeList attributeList = method.Attributes;
                ParseAttributes(id, attributeList);
            }

            InstructionList instList = method.Instructions;

            for (int index = 0; index < instList.Length; index++)
            {
                Instruction inst = instList[index];
                if (inst.OpCode == OpCode.Ldstr)
                {
                    object obj = inst.Value;
                    string str = obj.ToString();

                    if ((str = CheckString(str)) != null)
                    {
                        string itemID = string.Format("{0}.[{1}]", id, inst.Offset);
                        c_strings.Add(new StringValue( itemID, str,System.Globalization.CultureInfo.InvariantCulture));

                    }
                }
            }
        }
    }
}

