using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;


namespace SpellCheck.Plugin.Win32BinaryStrings
{
    class WindowsAPI
    {
        internal const uint RT_CURSOR = 0x00000001;
        internal const uint RT_BITMAP = 0x00000002;
        internal const uint RT_ICON = 0x00000003;
        internal const uint RT_MENU = 0x00000004;
        internal const uint RT_DIALOG = 0x00000005;
        internal const uint RT_STRING = 0x00000006;
        internal const uint RT_FONTDIR = 0x00000007;
        internal const uint RT_FONT = 0x00000008;
        internal const uint RT_ACCELERATOR = 0x00000009;
        internal const uint RT_RCDATA = 0x0000000a;
        internal const uint RT_MESSAGETABLE = 0x0000000b;


        internal const ushort LANG_NEUTRAL = 0x0000;
        internal const ushort LANG_SPANISH = 0x0c0a;
        internal const ushort LANG_GERMAN = 0x0407;
        internal const ushort LANG_FRENCH = 0x040c;
        internal const ushort LANG_ENGLISH_US = 0x0409;


        internal const uint LOAD_LIBRARY_AS_DATAFILE = 0x00000002;
        internal const uint DONT_RESOLVE_DLL_REFERENCES = 0x00000001;

        internal const uint RES_TYPE = 0x00000001;

        [StructLayout(LayoutKind.Sequential, Pack = 2, CharSet = CharSet.Auto)]
        internal struct DLGITEMTEMPLATE
        {
            uint style;
            uint dwExtendedStyle;
            short x;
            short y;
            short cx;
            short cy;
            ushort id;
        }

        [StructLayout(LayoutKind.Sequential, Pack = 2, CharSet = CharSet.Auto)]
        internal struct DLGITEMTEMPLATEEX
        {
            internal uint helpID;
            internal uint exStyle;
            internal uint style;
            internal short x;
            internal short y;
            internal short cx;
            internal short cy;
            internal uint id;
            //sz_Or_Ord windowClass;
            //sz_Or_Ord title;
            //internal short extraCount;
        }

        [StructLayout(LayoutKind.Sequential, Pack = 2, CharSet = CharSet.Auto)]
        internal struct DLGTEMPLATEEX
        {
            internal ushort dlgVer;
            internal ushort signature;
            internal uint helpID;
            internal uint exStyle;
            internal uint style;
            internal ushort cDlgItems;
            internal short x;
            internal short y;
            internal short cx;
            internal short cy;

            // The structure has more fields 
            // but are variable length
            //internal object menu;
            //internal object windowClass;
            //internal StringBuilder title;
            //internal uint pointsize;
            //internal uint weight;
            //internal byte italic;
            //internal byte charset;
            //internal StringBuilder typeface;
        }





        [StructLayout(LayoutKind.Sequential, Pack = 2, CharSet = CharSet.Auto)]
        internal struct DLGTEMPLATE
        {
            internal uint style;
            internal uint extendedStyle;
            internal ushort cdit;
            internal short x;
            internal short y;
            internal short cx;
            internal short cy;
            //internal short menuResource;
            //internal short windowClass;
            //internal short titleArray;
            //internal short fontPointSize;
            //[MarshalAs(UnmanagedType.LPWStr)]
            //internal string fontTypeface;
        }

        [StructLayout(LayoutKind.Sequential)]
        public class MESSAGE_RESOURCE_BLOCK {
        public System.UInt32 LowId;
        public System.UInt32 HighId;
        public System.UInt32 OffsetToEntries;
        }

        [StructLayout(LayoutKind.Sequential)]
        public class MESSAGE_RESOURCE_DATA {
        public System.UInt32 NumberOfBlocks;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 1)]
        public MESSAGE_RESOURCE_BLOCK[] Blocks;
        }


        [StructLayout(LayoutKind.Sequential)]
        public struct MESSAGE_RESOURCE_ENTRY {
        System.UInt16 Length;
        System.UInt16 Flags;
        [MarshalAsAttribute(UnmanagedType.ByValArray,SizeConst=1)]
        System.Byte[] Text;
        }


        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern IntPtr LoadLibraryEx(
          string lpFileName,
          IntPtr hFile,
          uint dwFlags);

        [DllImport("user32", EntryPoint = "DialogBoxIndirectParam")]
        internal static extern IntPtr DialogBoxParam(IntPtr hInstance,
                                                          IntPtr lpTemplateName,
                                                          int hwndParent,
                                                          int lpDialogFunc,
                                                          int dwInitParam);




        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern bool FreeLibrary(IntPtr hModule);

        [DllImport("Kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        internal static extern IntPtr FindResource(IntPtr hModule,
                                                 IntPtr pName,
                                                 IntPtr pType);

        [DllImport("Kernel32.dll", SetLastError = true)]
        internal static extern IntPtr LoadResource(IntPtr hModule,
                                                 IntPtr hResource);

        [DllImport("Kernel32.dll", SetLastError = true)]
        internal static extern uint SizeofResource(IntPtr hModule,
                                                 IntPtr hResource);

        [DllImport("Kernel32.dll")]
        internal static extern IntPtr LockResource(IntPtr hGlobal);

        [DllImport("Kernel32.dll")]
        internal static extern IntPtr UnlockResource(IntPtr hGlobal);

        [DllImport("kernel32.dll", EntryPoint = "EnumResourceNamesW",
          CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern bool EnumResourceNamesWithName(
          IntPtr hModule,
          string lpszType,
          EnumResNameDelegate lpEnumFunc,
          IntPtr lParam);

   


        [DllImport("kernel32.dll", EntryPoint = "EnumResourceNamesW",
          CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern bool EnumResourceNamesWithID(
          IntPtr hModule,
          uint lpszType,
          EnumResNameDelegate lpEnumFunc,
          IntPtr lParam);
              
      

        [DllImport("kernel32.dll", EntryPoint = "EnumResourceLanguagesA",
                       CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern bool EnumResourceLanguagesWithID(
          IntPtr hModule,
          IntPtr lpszType,
          IntPtr lpName,
          EnumResLangDelegate lpEnumFunc,
          IntPtr lParam);

        //[DllImport("kernel32.dll", EntryPoint = "EnumResourceLanguagesA",
        //       CharSet = CharSet.Unicode, SetLastError = true)]
        //internal static extern bool EnumResourceLanguagesWithName(
        //  IntPtr hModule,
        //  string lpszType,
        //  string lpName,
        //  EnumResLangDelegate lpEnumFunc,
        //  IntPtr lParam);

        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern bool EnumResourceTypes(
            IntPtr hModule,
            EnumTypesDelegate lpEnumFunc,
            IntPtr lParam);



        [DllImport("user32.dll")]
        internal static extern int LoadString(IntPtr hInstance, uint uID, StringBuilder lpBuffer, int nBufferMax);


        internal delegate bool EnumResNameDelegate(
          IntPtr hModule,
          IntPtr lpszType,
          IntPtr lpszName,
          IntPtr lParam);

        internal delegate bool EnumResLangDelegate(
          IntPtr hModule,
          IntPtr lpszType,
          IntPtr lpszName,
          int wIDLanguage,
          IntPtr lParam);

        internal delegate bool EnumTypesDelegate(
          IntPtr hModule,
          IntPtr lpszType,
          IntPtr lParam);


    }
}
