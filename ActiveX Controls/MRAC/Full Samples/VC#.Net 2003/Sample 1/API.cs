using System;
using System.Runtime.InteropServices;
namespace WindowsApplication1
{
	/// <summary>
	/// Summary description for API.
	/// </summary>
	public class API
	{
		public API()
		{

		
		}

		[DllImport("user32.dll", EntryPoint="GetAsyncKeyState")]
		public static extern bool GetAsyncKeyState(Int32 vKey);

		[DllImport("shell32.dll", EntryPoint="ShellExecuteA")]
		public static extern Int32 ShellExecute(Int32 hwnd, String lpOperation, String lpFile, String lpParameters, String lpDirectory, Int32 nShowCmd);

	}
}
