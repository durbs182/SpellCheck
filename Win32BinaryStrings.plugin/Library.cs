using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Globalization;

using SpellCheck.Common.XmlData;
using SpellCheck.Common;

using SpellCheck.Common.Utility;

namespace SpellCheck.Plugin.Win32BinaryStrings
{
    class Library
    {
        #region helper methods

        private static bool IS_INTRESOURCE(IntPtr value)
        {
            if (((uint)value) > ushort.MaxValue)
                return false;
            return true;
        }
        private static uint GET_RESOURCE_ID(IntPtr value)
        {
            if (IS_INTRESOURCE(value) == true)
                return (uint)value;
            throw new System.NotSupportedException(value + " is not an ID!");
        }
        private static string GET_RESOURCE_NAME(IntPtr value)
        {
            if (IS_INTRESOURCE(value) == true)
                return value.ToString();
            return Marshal.PtrToStringUni((IntPtr)value);
        }

        private static String GetStringResource(IntPtr hModuleInstance, uint uiStringID)
        {
            StringBuilder sb = new StringBuilder(1024);
            WindowsAPI.LoadString(hModuleInstance, uiStringID, sb, sb.Capacity + 1);
            return sb.ToString();
        }


        private static int ReadS32(byte[] array, int offset)
        {
            return array[offset] | array[offset + 1] << 8 | array[offset + 2] << 16 | array[offset + 3] << 24;
        }

        private static short ReadS16(byte[] array, int offset)
        {
            return (short)(array[offset] | array[offset + 1] << 8);
        }

        private static int AllignOnDWORD(int wPointer)
        {
            int mod = (wPointer + 3) & ~3;
            return mod;
        }

 
        #endregion 


        static Dictionary<System.Globalization.CultureInfo, List<StringValue>> c_Strings;


        /// <summary>
        /// Method called IStringResourceFileHandler LoadStrings method
        /// </summary>
        /// <param name="fileInfo"></param>
        /// <param name="checkStringDelegate"></param>
        /// <param name="wordApp"></param>
        /// <returns></returns>
        internal static FileData LoadFile(Win32BinaryFileHandler win32BinaryFileHandler, System.IO.FileInfo fileInfo, IList<CheckStringDelegate> checkStringDelegates)
        {
            // instanciate c_strings dictionary
            c_Strings = new Dictionary<System.Globalization.CultureInfo, List<StringValue>>();
            FileData fileData = new FileData(fileInfo.FullName);

            // load the dll or exe using WinAPI LoadLibraryEx call
            IntPtr hModule = WindowsAPI.LoadLibraryEx(fileInfo.FullName, IntPtr.Zero, WindowsAPI.DONT_RESOLVE_DLL_REFERENCES | WindowsAPI.LOAD_LIBRARY_AS_DATAFILE);

            // WinAPI call to enumerate all resources in the file
            if (WindowsAPI.EnumResourceTypes(hModule, new WindowsAPI.EnumTypesDelegate(EnumResTypes), IntPtr.Zero) == false)
            {
                Lib.ConsoleLog("Win32Binary EnumResourceTypes Parsing Error in file: {1} [gle: {0}]", Marshal.GetLastWin32Error(), fileInfo.FullName);
            }

       
            // unload the file
            WindowsAPI.FreeLibrary(hModule);

            // walk through each string and pass it the checkString Delegate   
            foreach (System.Globalization.CultureInfo culture in c_Strings.Keys)
            {
                foreach (StringValue stringVal in c_Strings[culture])
                {
                    fileData = win32BinaryFileHandler.HandleSpellingDelegates(checkStringDelegates, fileData, stringVal);
                }
            }

            return fileData;
            
        }


        /// <summary>
        /// Trim any characters that might interfer with spell checking. Use lock to make adding to the string dictionary
        /// thread safe.
        /// </summary>
        /// <param name="culture"></param>
        /// <param name="key"></param>
        /// <param name="value"></param>
        private static void AddToStrings(System.Globalization.CultureInfo culture,string key,string value)
        {
            value = value.Trim();
            value = value.Replace("&", string.Empty);
            value = value.Replace("%n", string.Empty);
            value = value.Replace("%t", string.Empty);
            value = value.Replace("%s", string.Empty);
            value = value.Replace(">", string.Empty);
            value = value.Replace("<", string.Empty);
            value = value.Replace("\\n", " ");
            value = value.Replace("\\t", " ");
            value = value.Replace("\\", " ");

            if (value.Equals(String.Empty))
                return;

            lock (c_Strings)
            {

                Lib.Log("{0} {1} {2} [{3}]", Lib.GetMethodName(), key, value, System.Threading.Thread.CurrentThread.ManagedThreadId);

                if (c_Strings.ContainsKey(culture))
                {
                    List<StringValue> strings = c_Strings[culture];
                    strings.Add(new StringValue(key, value, culture));
                }
                else
                { 
                    List<StringValue> strings = new List<StringValue>();
                    strings.Add(new StringValue(key, value,culture));
                    c_Strings.Add(culture, strings);
                }
            }
        }

   
        /// <summary>
        /// Parse an RTF embedded resource
        /// </summary>
        /// <param name="hModule"></param>
        /// <param name="lpszType"></param>
        /// <param name="lpszName"></param>
        /// <param name="lParam"></param>
        /// <returns></returns>
        static bool EnumRCRTF(IntPtr hModule,
              IntPtr lpszType,
              IntPtr lpszName,
              IntPtr lParam)
        {

            string type = Marshal.PtrToStringAuto(lpszType);
   
            // Load the resource and get the resource bits.
            IntPtr hResource = WindowsAPI.FindResource(hModule, lpszName, lpszType);
            IntPtr hDialogRes = WindowsAPI.LoadResource(hModule, hResource);
            IntPtr pDibBits = WindowsAPI.LockResource(hDialogRes);

            // We make a local copy of the DIB bits as pDibBits may be 
            // freed when the module is unloaded.
            byte[] arDibBits = new byte[WindowsAPI.SizeofResource(hModule, hResource)];

            Marshal.Copy(pDibBits, arDibBits, 0, arDibBits.Length);

            System.Windows.Forms.RichTextBox rtb = new System.Windows.Forms.RichTextBox();
            
           
            rtb.Rtf = System.Text.Encoding.UTF8.GetString(arDibBits);
            string s = rtb.Text;
            AddToStrings(System.Globalization.CultureInfo.InvariantCulture, "RTF", s);

            return true;
        }

        #region Dialog Parser

        /// <summary>
        /// Parse the reource data and check the dialog type.
        /// Dialog or DialogEx
        /// </summary>
        /// <param name="hModule"></param>
        /// <param name="lpszType"></param>
        /// <param name="lpszName"></param>
        /// <param name="culture"></param>
        /// <returns></returns>
        static bool ParseDialogResource(IntPtr hModule,
            IntPtr lpszType,
            IntPtr lpszName,
            CultureInfo culture)
        {
            Lib.Log("Dialog ID: {0}", GET_RESOURCE_NAME(lpszName));

            // Load the resource and get the resource bits.
            IntPtr hResource = WindowsAPI.FindResource(hModule, lpszName, lpszType);
            IntPtr hDialogRes = WindowsAPI.LoadResource(hModule, hResource);
            IntPtr pDibBits = WindowsAPI.LockResource(hDialogRes);

            // We make a local copy of the DIB bits as pDibBits may be 
            // freed when the module is unloaded.
            byte[] arDibBits = new byte[WindowsAPI.SizeofResource(hModule, hResource)];

            WindowsAPI.DLGTEMPLATEEX dlgtemplateex = (WindowsAPI.DLGTEMPLATEEX)Marshal.PtrToStructure(pDibBits, typeof(WindowsAPI.DLGTEMPLATEEX));
            Marshal.Copy(pDibBits, arDibBits, 0, arDibBits.Length);

            string dialogID = GET_RESOURCE_NAME(lpszName);

            if (dlgtemplateex.signature == 0xFFFF)
            {
                ParseDialogEx(dialogID, arDibBits, dlgtemplateex, culture);
            }
            else
            {
                WindowsAPI.DLGTEMPLATE dlgtemplate = (WindowsAPI.DLGTEMPLATE)Marshal.PtrToStructure(pDibBits, typeof(WindowsAPI.DLGTEMPLATE));
                ParseDialog(dialogID, arDibBits, dlgtemplate, culture);
            }


            return true;

        }

        /// <summary>
        /// Parse a dialog
        /// </summary>
        /// <param name="dialogID"></param>
        /// <param name="arDibBits"></param>
        /// <param name="dlgtemplate"></param>
        /// <param name="culture"></param>
        private static void ParseDialog(string dialogID, byte[] arDibBits, WindowsAPI.DLGTEMPLATE dlgtemplate, CultureInfo culture)
        {
            //dialog
            int controlCount = (int)ReadS16(arDibBits, 8);

            // skip to the end of the template
            int wPointer = Marshal.SizeOf(typeof(WindowsAPI.DLGTEMPLATE));

            //menu
            if (ReadS16(arDibBits, wPointer) == 0x0000)
            {
                // No menu resource
                wPointer += 2;
            }
            else
            {
                if (ReadS16(arDibBits, wPointer) == -1)
                {
                    wPointer += 2;
                }
                else
                {
                    while (ReadS16(arDibBits, wPointer) != 0x0000)
                    {
                        wPointer += 2;
                    }
                }
            }

            // window class
            if (ReadS16(arDibBits, wPointer) == 0x0000)
            {
                wPointer += 2;
            }
            else
            {
                if (ReadS16(arDibBits, wPointer) == -1)
                {
                    wPointer += 2;
                }
                else
                {
                    while (ReadS16(arDibBits, wPointer) != 0x0000)
                    {
                        wPointer += 2;
                    }
                }
            }

            // title string
            if (ReadS16(arDibBits, wPointer) != 0x0000)
            {
                StringBuilder sb = new StringBuilder();
                short sh = 0;
                while ((sh = ReadS16(arDibBits, wPointer)) != 0x0000)
                {
                    char c = (char)sh;
                    sb.Append(c);
                    wPointer += 2;
                }

      
                AddToStrings(culture, string.Format("Dialog[{0}]_{1}_Caption", culture.TwoLetterISOLanguageName, dialogID), sb.ToString());
                
            }

            //TODO: need to work out this flag???
            uint style = (uint)ReadS32(arDibBits, 0);
            uint s = style & 0x40;

            //font
            if (true)
            {
                wPointer += 2;
                short fontSize = ReadS16(arDibBits, wPointer);
                wPointer += 2;

                StringBuilder sb = new StringBuilder();
                short sh = 0;
                while ((sh = ReadS16(arDibBits, wPointer)) != 0x0000)
                {
                    char c = (char)sh;
                    sb.Append(c);
                    wPointer += 2;
                }

                Lib.Log(sb.ToString());
                wPointer += 2;

                // need to allign here first
                
            }

            for (int i = 0; i < controlCount; i++)
            {
                wPointer = AllignOnDWORD(wPointer);
                int controlstyle = ReadS32(arDibBits, wPointer);

                wPointer += Marshal.SizeOf(typeof(WindowsAPI.DLGITEMTEMPLATE));

                if (ReadS16(arDibBits, wPointer) == -1)
                {
                    wPointer += 2;
                }
                else
                {
                    while (ReadS16(arDibBits, wPointer) != 0x0000)
                    {
                        wPointer += 2;
                    }
                }

                wPointer += 2;

                // control text
                if (ReadS16(arDibBits, wPointer) == -1)
                {
                    wPointer += 2;
                }
                else
                {
                    StringBuilder sb = new StringBuilder();
                    short sh = 0;
                    while ((sh = ReadS16(arDibBits, wPointer)) != 0x0000)
                    {
                        char c = (char)sh;
                        sb.Append(c);
                        wPointer += 2;
                    }

                    AddToStrings(culture,string.Format("Dialog[{0}]_{1}_Control{2}",culture.TwoLetterISOLanguageName,dialogID,i),sb.ToString());
                }

                wPointer += 2;

                short creationFlag = 0x00;
                if ((creationFlag = ReadS16(arDibBits, wPointer)) == 0x00)
                {
                    wPointer += 2;
                }

            }
        }

        /// <summary>
        /// Parse a dialogex
        /// </summary>
        /// <param name="dialogID"></param>
        /// <param name="arDibBits"></param>
        /// <param name="dlgtemplateex"></param>
        /// <param name="culture"></param>
        private static void ParseDialogEx(string dialogID, byte[] arDibBits, WindowsAPI.DLGTEMPLATEEX dlgtemplateex, CultureInfo culture)
        {
            //this is a dialogex
            int controlCount = (int)dlgtemplateex.cDlgItems;

            // skip to the end ot the template
            int wPointer = Marshal.SizeOf(typeof(WindowsAPI.DLGTEMPLATEEX));

            //menu
            if (ReadS16(arDibBits, wPointer) == 0x0000)
            {
                // No menu resource
                wPointer += 2;
            }
            else
            {
                if (ReadS16(arDibBits, wPointer) == -1)
                {
                    wPointer += 2;
                }
                else
                {
                    while (ReadS16(arDibBits, wPointer) != 0x0000)
                    {
                        wPointer += 2;
                    }
                }
            }

            // window class
            if (ReadS16(arDibBits, wPointer) == 0x0000)
            {
                wPointer += 2;
            }
            else
            {
                if (ReadS16(arDibBits, wPointer) == -1)
                {
                    wPointer += 2;
                }
                else
                {
                    while (ReadS16(arDibBits, wPointer) != 0x0000)
                    {
                        wPointer += 2;
                    }
                }
            }

            // title string
            if (ReadS16(arDibBits, wPointer) != 0x0000)
            {
                StringBuilder sb = new StringBuilder();
                short sh = 0;
                while ((sh = ReadS16(arDibBits, wPointer)) != 0x0000)
                {
                    char c = (char)sh;
                    sb.Append(c);
                    wPointer += 2;
                }

                AddToStrings(culture, string.Format("Dialog[{0}]_{1}_Caption", culture.TwoLetterISOLanguageName,dialogID), sb.ToString());

            }

            //uint pointsize;
            //uint weight;
            //byte italic;
            //byte charset;

            wPointer += 8;

            //font name
            while (ReadS16(arDibBits, wPointer) != 0x0000)
            {
                wPointer += 2;
            }
            wPointer += 2;


            for (int i = 0; i < controlCount; i++)
            {
                // need to allign to a word boundary here first after font name string
                wPointer = AllignOnDWORD(wPointer);

                //helpID
                wPointer += 4;

                //exStyle
                wPointer += 4;

                //style
                wPointer += 4;

                //x
                short x = ReadS16(arDibBits, wPointer);
                wPointer += 2;
                //y
                short y = ReadS16(arDibBits, wPointer);
                wPointer += 2;

                //cx
                short cx = ReadS16(arDibBits, wPointer);
                wPointer += 2;

                //cy
                short cy = ReadS16(arDibBits, wPointer);
                wPointer += 2;

                Lib.Log(string.Format("x={0}, y={1}, cx={2}, cy={3}", x, y, cx, cy));

                //id
                wPointer += 4;

                //windowclass
                if (ReadS16(arDibBits, wPointer) == -1)
                {
                    wPointer += 4;
                }
                else
                {
                    short sh = 0;
                    while ((sh = ReadS16(arDibBits, wPointer)) != 0x0000)
                    {
                        char c = (char)sh;
                        Lib.Log(c);
                        wPointer += 2;
                    }
                    Lib.Log(string.Empty);
                    wPointer += 2;
                }

                //control text
                if (ReadS16(arDibBits, wPointer) == -1)
                {
                    wPointer += 6;
                }
                else
                {
                    short sh = 0;
                    StringBuilder sb = new StringBuilder();
                    while ((sh = ReadS16(arDibBits, wPointer)) != 0x0000)
                    {
                        char c = (char)sh;
                        sb.Append(c);
                        wPointer += 2;
                    }

                    AddToStrings(culture, string.Format("Dialog[{0}]_{1}_Control{2}",culture.TwoLetterISOLanguageName,dialogID,i),sb.ToString());
                    wPointer += 2;

                }

                short creationFlag = 0x00;
                if((creationFlag = ReadS16(arDibBits,wPointer))== 0x00)
                {
                    wPointer += 2;
                }
            }
        }


        #endregion

        /// <summary>
        /// Callback method from WinAPI call EnumResourceTypes
        /// </summary>
        /// <param name="hModule"></param>
        /// <param name="lpszType"></param>
        /// <param name="lParam"></param>
        /// <returns></returns>
        static bool EnumResTypes(IntPtr hModule,
            IntPtr lpszType,
            IntPtr lParam)
        {

            if (IS_INTRESOURCE(lpszType))
            {
                uint type = GET_RESOURCE_ID(lpszType);
                Lib.Log("ResourceType: " + type);

                switch (type)
                {
                    case WindowsAPI.RT_MESSAGETABLE:
                    case WindowsAPI.RT_DIALOG:
                    case WindowsAPI.RT_STRING:
                        if (WindowsAPI.EnumResourceNamesWithID(hModule, type,
                          new WindowsAPI.EnumResNameDelegate(EnumResType), IntPtr.Zero) == false)
                        {
                            Lib.ConsoleLog("Win32Binary EnumResType Parsing Error in file:  [gle: {0}]", Marshal.GetLastWin32Error());
                        }
                        break;
                }
            }
            else
            { 
                string typeStr = Marshal.PtrToStringAnsi(lpszType);

                switch (typeStr)
                { 
                    case "RTF":
                        if (WindowsAPI.EnumResourceNamesWithName(hModule, typeStr,
                          new WindowsAPI.EnumResNameDelegate(EnumResType), IntPtr.Zero) == false)
                        {
                            Lib.ConsoleLog("Win32Binary RTF Parsing Error in file:  [gle: {0}]", Marshal.GetLastWin32Error());
                        }
                        break;
                        

                }
            }

            return true;
        }


        /// <summary>
        /// Callback method for WinAPI EnumResourceNamesWithID and EnumResourceNamesWithName
        /// </summary>
        /// <param name="hModule"></param>
        /// <param name="lpszType"></param>
        /// <param name="lpszName"></param>
        /// <param name="lParam"></param>
        /// <returns></returns>
        static bool EnumResType(IntPtr hModule,
            IntPtr lpszType,
            IntPtr lpszName,
            IntPtr lParam)
        {
            Lib.Log("String ID: {0}", GET_RESOURCE_NAME(lpszName));

            if (IS_INTRESOURCE(lpszType))
            {

                if (WindowsAPI.EnumResourceLanguagesWithID(hModule, lpszType, lpszName,
                            new WindowsAPI.EnumResLangDelegate(EnumLangRes), IntPtr.Zero) == false)
                {
                    Lib.ConsoleLog("Win32Binary EnumResType Parsing Error in file:  [gle: {0}]", Marshal.GetLastWin32Error());
                }
            }
            else
            {
                string typeStr = Marshal.PtrToStringAuto(lpszType);

                switch (typeStr)
                {
                    case "RTF":
                        if (WindowsAPI.EnumResourceNamesWithName(hModule, typeStr,
                          new WindowsAPI.EnumResNameDelegate(EnumRCRTF), IntPtr.Zero) == false)
                        {
                            Lib.ConsoleLog("Win32Binary RTF Parsing Error in file:  [gle: {0}]", Marshal.GetLastWin32Error());
                        }
                        break;

                }
            }

            return true;
        }


        /// <summary>
        /// Callback method for WInAPI EnumResourceLanguagesWithID
        /// </summary>
        /// <param name="hModule"></param>
        /// <param name="lpszType"></param>
        /// <param name="lpszName"></param>
        /// <param name="wIDLanguage"></param>
        /// <param name="lParam"></param>
        /// <returns></returns>
        static bool EnumLangRes(IntPtr hModule,
        IntPtr lpszType,
        IntPtr lpszName,
        int wIDLanguage,
        IntPtr lParam)
        {
            Lib.Log("String ID: {0}", GET_RESOURCE_NAME(lpszName));

            System.Globalization.CultureInfo culture = null;
            
            System.Globalization.CultureInfo wIDLanguageCI = System.Globalization.CultureInfo.GetCultureInfo(wIDLanguage);
            
         
            switch (wIDLanguageCI.TwoLetterISOLanguageName)
            { 
            	case Locale.DE:
                case Locale.FR:
                case Locale.ES:
                    culture = wIDLanguageCI;
                    break;
                case Locale.JA:
                    Lib.Log("Japanese strings can't be checked with this tool");
                    return true;
                default:
                    culture = System.Globalization.CultureInfo.InvariantCulture;
                    break;
            }

            Lib.Log(culture.EnglishName);

            uint resType = GET_RESOURCE_ID(lpszType);


            switch (resType)
            { 
                case WindowsAPI.RT_STRING:
                    ParseStringTable(hModule, lpszType,lpszName,  culture);
                    break;
                case WindowsAPI.RT_DIALOG:
                    ParseDialogResource(hModule,lpszType, lpszName, culture);
                    break;
                case WindowsAPI.RT_MESSAGETABLE:
                    ParseMessageTable(hModule, lpszType, lpszName, culture);
                    break;
            }


            return true;
        }



        /// <summary>
        /// Uses GetStringResource to lookup each strings resource ID
        /// </summary>
        /// <param name="hModule"></param>
        /// <param name="lpszType"></param>
        /// <param name="lpszName"></param>
        /// <param name="culture"></param>
        /// <returns></returns>
        static bool ParseStringTable(IntPtr hModule, IntPtr lpszType,IntPtr lpszName,  CultureInfo culture)
        {
            Lib.Log("String ID: {0}", GET_RESOURCE_NAME(lpszName));

            //DumpResource(hModule, lpszType, lpszName);

            uint strCount = 16;
            
            for (uint i = 0; i < strCount; i++)
            {
                uint id = (GET_RESOURCE_ID(lpszName) * strCount) - strCount + i;

                string s = GetStringResource(hModule, id);

                AddToStrings(culture, string.Format("StringTable[{0}]_{1}_{2}", culture.TwoLetterISOLanguageName,GET_RESOURCE_NAME(lpszName), i), s);
            }

            return true;
        }

        /// <summary>
        /// Parse MessageTable resource data
        /// </summary>
        /// <param name="hModule"></param>
        /// <param name="lpszType"></param>
        /// <param name="lpszName"></param>
        /// <param name="culture"></param>
        /// <returns></returns>
        static bool ParseMessageTable(IntPtr hModule, IntPtr lpszType,IntPtr lpszName, 
              CultureInfo culture)
        {
            // Load the resource and get the resource bits.
            IntPtr hResource = WindowsAPI.FindResource(hModule, lpszName, lpszType);
            IntPtr hDialogRes = WindowsAPI.LoadResource(hModule, hResource);
            IntPtr pDibBits = WindowsAPI.LockResource(hDialogRes);


            byte[] arDibBits = new byte[WindowsAPI.SizeofResource(hModule, hResource)];

            Marshal.Copy(pDibBits, arDibBits, 0, arDibBits.Length);            

            int wpointer = 0;

            // MESSAGE_RESOURCE_DATA has number of MESSAGE_RESOURCE_BLOCK
            int numBlocks = ReadS32(arDibBits, wpointer);

            uint tableNum = (uint)lpszName; 

            wpointer += 4;

            for (int i = 0; i < numBlocks; i++)
            {
                // MESSAGE_RESOURCE_BLOCK
                //
                // DWORD LowId;
                // DWORD HighId;
                // DWORD OffsetToEntries;
                //

                // Each MESSAGE_RESOURCE_BLOCK represents a sequence of consecutive message table entries 
                // in a message table, starting at the ID indicated by the member LowId and ending with the 
                // ID indicated by the HighId member of the MESSAGE_RESOURCE_BLOCK struct. Adding the value 
                // in the OffsetToEntries member to the address of the MESSAGE_RESOURCE_BLOCK struct itself 
                // then yields the start address of the message table entry with the first ID of the 
                // MESSAGE_RESOURCE_BLOCK which is contained in the LowId member. This address points to a 
                // MESSAGE_RESOURCE_ENTRY data structure, also defined in the winnt.h, as such:

                int low = ReadS32(arDibBits, wpointer);
                wpointer += 4;
                int high = ReadS32(arDibBits, wpointer);
                wpointer += 4;
                int offSet = ReadS32(arDibBits, wpointer);
                int num = high - low + 1;

                for (int k = 0; k < num; k++)
                {
                    string str = null;

                    // MESSAGE_RESOURCE_ENTRY
                    //
                    // WORD Length; - legth of the data but does include padding. so look for \0
                    // WORD Flags; 0x0000 is ANSI & 0x0001 is Unicode
                    // 

                    short length = ReadS16(arDibBits, offSet);
                    offSet += 2;
                    short encoding = ReadS16(arDibBits, offSet);
                    offSet += 2;

                    // for both encoding types read each character until a \0 is 
                    if (encoding == 0x0000)
                    {
                        StringBuilder sb = new StringBuilder();

                        for (int j = 0; j < length; j++)
                        {
                            byte b = arDibBits[offSet + j];

                            if (b == 0x0)
                            {
                                break;
                            }
                            else
                            {
                                sb.Append((char)b);
                            }
                        }
                        str = sb.ToString().Trim();

                    }
                    else
                    {
                        StringBuilder sb = new StringBuilder();

                        for (int j = 0; j < length; j++)
                        {
                            short sh = ReadS16(arDibBits, offSet + j);

                            if (sh == 0x00)
                            {
                                break;
                            }
                            else
                            {
                                sb.Append((char)sh);
                            }
                            j++;
                        }

                        str = sb.ToString().Trim();
                    }

                    AddToStrings(culture, string.Format("MessageTable[{0}]_{1}_{2}", culture.TwoLetterISOLanguageName,tableNum,numBlocks * k), str);

                    offSet += (length -4);
                }

                wpointer += 4;

            }

            return true;
        }

        /// <summary>
        /// Debug method to dump the resource data to a file 
        /// </summary>
        /// <param name="hModule"></param>
        /// <param name="lpszType"></param>
        /// <param name="lpszName"></param>
        private static void DumpResource(IntPtr hModule, IntPtr lpszType, IntPtr lpszName)
        {
            // Load the resource and get the resource bits.
            IntPtr hResource = WindowsAPI.FindResource(hModule, lpszName, lpszType);
            IntPtr hDialogRes = WindowsAPI.LoadResource(hModule, hResource);
            IntPtr pDibBits = WindowsAPI.LockResource(hDialogRes);

            byte[] arDibBits = new byte[WindowsAPI.SizeofResource(hModule, hResource)];

            Marshal.Copy(pDibBits, arDibBits, 0, arDibBits.Length);

            string[] hexArr = new string[arDibBits.Length];

            for (int i = 0; i < arDibBits.Length; i++)
            {
                hexArr[i] = arDibBits[i].ToString("X");
            }

            System.IO.BinaryWriter bw = new System.IO.BinaryWriter(System.IO.File.Create(GET_RESOURCE_NAME(lpszName) + ".bin"));
            bw.Write(arDibBits);
            bw.Close();
        }

    }
}
