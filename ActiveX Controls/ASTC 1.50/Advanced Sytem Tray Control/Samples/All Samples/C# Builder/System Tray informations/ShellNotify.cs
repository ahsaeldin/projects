using System;
using System.Runtime.InteropServices;


namespace Get_System_Tray_Informtions
{
	/// <summary>
	/// Summary description for ShellNotify.
	/// </summary>
	///
	public class ShellNotify
	{
		[StructLayout(LayoutKind.Sequential)]
			public struct NOTIFYICONDATA		{			public Int32 cbSize;
			public Int32 hwnd;
			public Int32 uID;
			public Int32 uFlags;
			public Int32 uCallbackMessage;
			public Int32 hIcon;
			[MarshalAs(UnmanagedType.ByValTStr, SizeConst=64)] public string szTip;		}
		
		[DllImport("shell32.dll", EntryPoint="Shell_NotifyIcon")]
		public static extern Boolean Shell_NotifyIcon(Int32 dwMessage,ref NOTIFYICONDATA lpData);

	}	
}
