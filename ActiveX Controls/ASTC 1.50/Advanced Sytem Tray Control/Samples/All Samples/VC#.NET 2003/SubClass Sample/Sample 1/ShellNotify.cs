using System;
using System.Runtime.InteropServices;
namespace SubClass_Sample
{
	/// <summary>
	/// Summary description for ShellNotify.
	/// </summary>
	public class ShellNotify
	{
		public ShellNotify()
		{
	
		}
        
		public struct POINTAPI
		{
           public Int32 X;
		   public Int32 Y;
		}

        public struct MINMAXINFO
		{
           public POINTAPI ptReserved;  
           public POINTAPI ptMaxSize;  
           public POINTAPI ptMaxPosition;  
           public POINTAPI ptMinTrackSize;  
	       public POINTAPI ptMaxTrackSize;  
		}

         //CopyMemory is the same as CopyMemory2 but we change the parameters
		 //to fit this Sample Case.

		[DllImport("kernel32.dll", EntryPoint="RtlMoveMemory")]
		public static extern Int32 CopyMemory(ref MINMAXINFO pDst,Int32 pSrc,Int32 ByteLen);
        
		[DllImport("kernel32.dll", EntryPoint="RtlMoveMemory")]
		public static extern Int32 CopyMemory2(Int32 pDst,ref MINMAXINFO pSrc,Int32 ByteLen);

		[DllImport("user32.dll", EntryPoint="SendMessage")]
		public static extern Int32 SendMessage(Int32 hwnd,Int32 wMsg,Int32 wParam,Int32 lParam);

	
		[DllImport("user32", EntryPoint="TrackPopupMenu")]
		public static extern Int32 TrackPopupMenu(Int32 hMenu,Int32 wFlags,Int32 x,Int32 y,Int32 nReserved
			,Int32 hwnd,Int32 lprc);
	
	}
}
