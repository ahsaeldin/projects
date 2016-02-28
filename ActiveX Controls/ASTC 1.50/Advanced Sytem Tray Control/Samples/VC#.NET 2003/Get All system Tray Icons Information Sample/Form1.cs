/*
'=========================================================================================
'Publisher      CprinGold Software.
'               http://www.cpringold.com
'               support@cpringold.com
'
'
'Description:  A sample code demonstrate how to use ASTC control to get all system tray icons
'                Information as well as how to hide a specified icon from system tray area.
'==========================================================================================

'First of all don't afraid from all these lines of codes, the main line which return the system
'Tray information is only one line which located in FillFlexGrid method [CoreLine] at line 10
'All the remaining lines is to fill the returned data from GetSysTrayIcons method into FlexGrid
'To display the data in a good manner.
'How i return this information?
'Well
'GetSysTrayIcons method returns an array of TrayIconInfo type
'Wait what does it means?
'We have a Type Called TrayIconInfo; you can use Object Browser to see this type in ASTC
'When you call GetSysTrayIcons method, it returns an array of this type contains the information
'I don't understand
'Well take this example
'If the system tray area contains three icons say [sound control icon,
'MacAfee icon and network connection icon].
'Then the returned data will be organized in an Array and the array will have three elements
'And every element will be a type called TrayIconInfo and TrayIconInfo will contain the information.
'
'So what is this information?
'
'Public Type TrayIconInfo
'    Hwnd As Long                 Handle to the window that will receive notification messages associated with an icon in the taskbar status area.
'    uId As Long                  Application-defined identifier of the taskbar icon.
'    ucallbackMessage As Long     Application-defined message identifier. The system uses this identifier for notification messages that it sends to the window identified in hWnd. These notifications are sent when a mouse event occurs in the bounding rectangle of the icon.
'    Param1(1) As Long            unknown interpretation
'    hIcon As Long                Handle to the icon
'    Param2(2) As Long            unknown interpretation
'    APath As String              The path of the application which adds the icon to system tray area
'    ToolTip As String            The tooltip associated with the icon
'    Bitmap as Long               index of the posted icon
'    Command as Long              command id
'    State as Byte                icon state
'    Style as Byte                icon style
'    Data as Long                 pointer towards the data of the tray
'    str As Long                  pointer to tooltip
'End Type

'I think that the only problem comes from you must know that the variable which will receive the
'Returned data type must be an array of TrayIconInfo
'See this example
'
'Dim TrayList() As TrayIconInfo
'TrayList = TrayIcons1.GetSysTrayIcons
'
'As you see the variable TrayList is an array of TrayIconInfo type

'How to hide icon from system tray area?
'Because every icon has an element in the returned array, all you need to do is to pass
'The element index to HideIcon method like this
'
'TrayList1.HideIcon 1
'Which 1 is the index of the icon element (TrayIconInfo) in the array?

'That's it
'If you have any further questions or need more sample code don't hesitate to contact us
'At support@cpringold.com

*/
using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Runtime.InteropServices;

namespace Get_All_system_Tray_Icons_Information_Sample
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
    public class Form1 : System.Windows.Forms.Form
	{
		internal System.Windows.Forms.Button ButRefresh;
		internal System.Windows.Forms.Button ButRestore;
		internal System.Windows.Forms.Button ButHide;
		private AxMSFlexGridLib.AxMSFlexGrid MSFlexGrid1;
		public ASTC.TrayIconInfo OldTrayInfo = new ASTC.TrayIconInfo();
		private AxASTC.AxTrayList axTrayList1; 
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Form1()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(Form1));
			this.ButRefresh = new System.Windows.Forms.Button();
			this.ButRestore = new System.Windows.Forms.Button();
			this.ButHide = new System.Windows.Forms.Button();
			this.MSFlexGrid1 = new AxMSFlexGridLib.AxMSFlexGrid();
			this.axTrayList1 = new AxASTC.AxTrayList();
			((System.ComponentModel.ISupportInitialize)(this.MSFlexGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.axTrayList1)).BeginInit();
			this.SuspendLayout();
			// 
			// ButRefresh
			// 
			this.ButRefresh.Location = new System.Drawing.Point(184, 232);
			this.ButRefresh.Name = "ButRefresh";
			this.ButRefresh.Size = new System.Drawing.Size(80, 23);
			this.ButRefresh.TabIndex = 5;
			this.ButRefresh.Text = "Refresh";
			this.ButRefresh.Click += new System.EventHandler(this.ButRefresh_Click);
			// 
			// ButRestore
			// 
			this.ButRestore.Location = new System.Drawing.Point(96, 232);
			this.ButRestore.Name = "ButRestore";
			this.ButRestore.Size = new System.Drawing.Size(80, 23);
			this.ButRestore.TabIndex = 4;
			this.ButRestore.Text = "Restore Last";
			this.ButRestore.Click += new System.EventHandler(this.ButRestore_Click);
			// 
			// ButHide
			// 
			this.ButHide.Location = new System.Drawing.Point(8, 232);
			this.ButHide.Name = "ButHide";
			this.ButHide.Size = new System.Drawing.Size(80, 23);
			this.ButHide.TabIndex = 3;
			this.ButHide.Text = "Hide";
			this.ButHide.Click += new System.EventHandler(this.ButHide_Click);
			// 
			// MSFlexGrid1
			// 
			this.MSFlexGrid1.Location = new System.Drawing.Point(8, 8);
			this.MSFlexGrid1.Name = "MSFlexGrid1";
			this.MSFlexGrid1.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("MSFlexGrid1.OcxState")));
			this.MSFlexGrid1.Size = new System.Drawing.Size(688, 216);
			this.MSFlexGrid1.TabIndex = 6;
			// 
			// axTrayList1
			// 
			this.axTrayList1.Enabled = true;
			this.axTrayList1.Location = new System.Drawing.Point(664, 232);
			this.axTrayList1.Name = "axTrayList1";
			this.axTrayList1.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("axTrayList1.OcxState")));
			this.axTrayList1.Size = new System.Drawing.Size(36, 36);
			this.axTrayList1.TabIndex = 7;
			// 
			// Form1
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(704, 266);
			this.Controls.Add(this.axTrayList1);
			this.Controls.Add(this.MSFlexGrid1);
			this.Controls.Add(this.ButRefresh);
			this.Controls.Add(this.ButRestore);
			this.Controls.Add(this.ButHide);
			this.Name = "Form1";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "VC#.NET TrayList Sample Code";
			this.TopMost = true;
			this.Load += new System.EventHandler(this.Form1_Load);
			((System.ComponentModel.ISupportInitialize)(this.MSFlexGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.axTrayList1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new Form1());
		}


        /*
		'==============================================================================
		' Method:       initFlexGrid
		'
		' Description:  setup the flexgrid control to receive the data
		'==============================================================================
		*/
		private void initFlexGrid()
		{ 
         
			MSFlexGrid1.set_ColWidth(0, 300);
			MSFlexGrid1.set_TextMatrix(0, 0, "I");
			MSFlexGrid1.set_ColWidth(1, 3000);
			MSFlexGrid1.set_TextMatrix(0, 1, "Application Path");
			MSFlexGrid1.set_ColWidth(2, 1300);
			MSFlexGrid1.set_TextMatrix(0, 2, "uID");
			MSFlexGrid1.set_ColWidth(3, 1300);
			MSFlexGrid1.set_TextMatrix(0, 3, "hWnd");
			MSFlexGrid1.set_ColWidth(4, 1300);
			MSFlexGrid1.set_TextMatrix(0, 4, "hIcon");
			MSFlexGrid1.set_ColWidth(5, 1700);
			MSFlexGrid1.set_TextMatrix(0, 5, "ToolTip");
			MSFlexGrid1.set_ColWidth(6, 1500);
			MSFlexGrid1.set_TextMatrix(0, 6, "uCallbackMessage");
			MSFlexGrid1.ForeColorFixed = Color.Blue;
           
		}

        /*
		'==============================================================================
		' Method:       FillFlexGrid
		'
		' Description:  The core function here we get the information from GetSysTrayIcons
		'               then fill the FlexGrid control with these info.
		'==============================================================================
		*/
		private void FillFlexGrid()
		{  
		   
			ASTC.TrayIconInfo[] TrayInfo = new ASTC.TrayIconInfo[100]; 

			initFlexGrid();

			//Get the System Tray Data
			TrayInfo = (ASTC.TrayIconInfo[])axTrayList1.GetSysTrayIcons();
            
			//Fill the FlexGrid with Data
			
			for (int i = 0; i != (int)axTrayList1.IconCount; i++ )
			{
				int count = i + 1;
				MSFlexGrid1.set_TextMatrix(i + 1, 0, count.ToString() );
				if (TrayInfo[i].APath == "")
				{
					MSFlexGrid1.set_TextMatrix(i + 1, 1, "N/A");
				}
				else
				{
					MSFlexGrid1.set_TextMatrix(i + 1, 1, TrayInfo[i].APath);
				}
				MSFlexGrid1.set_TextMatrix(i + 1, 2, TrayInfo[i].uId.ToString() );
				MSFlexGrid1.set_TextMatrix(i + 1, 3, TrayInfo[i].hwnd.ToString() );
				MSFlexGrid1.set_TextMatrix(i + 1, 4, TrayInfo[i].hIcon.ToString() );
				MSFlexGrid1.set_TextMatrix(i + 1, 5, TrayInfo[i].ToolTip.ToString() );
				MSFlexGrid1.set_TextMatrix(i + 1, 6, TrayInfo[i].ucallbackMessage.ToString());
			}						   
															   
		}
        
		/*
		 '==============================================================================
         ' Method:       Form1_Load
         '
         ' Description:  Call FillFlexGrid to display the data
         '==============================================================================
		*/
		private void Form1_Load(object sender, System.EventArgs e)
		{
			FillFlexGrid();
		}
        
		/*
		'==============================================================================
		' Method:       ButHide_Click
		'
		' Description:  Cool function used to hide any icon from system tray icon
		'               then refresh the FlexGrid control.
		'==============================================================================
		*/
		private void ButHide_Click(object sender, System.EventArgs e)
		{

			short RowSelected = (short)MSFlexGrid1.RowSel;
		
			ASTC.TrayIconInfo[] TrayInfo = new ASTC.TrayIconInfo[100]; 
            
			TrayInfo = (ASTC.TrayIconInfo[])axTrayList1.GetSysTrayIcons();
            
			//Save the Icon Info before Remove it in order to use
            //when we restore the icon
            OldTrayInfo = TrayInfo[RowSelected - 1];
               
			//Hide the selected Item from the System Tray area
			axTrayList1.HideIcon(ref RowSelected);
			//Clear the FlexGrid
			MSFlexGrid1.Clear();
			//fill the FlexGrid 
			FillFlexGrid();
		}

        /*
		 '==============================================================================
         ' Method:       ButRestore_Click
         '
         ' Description:  Restore the last removed icon To system tray.
         '==============================================================================
        */ 
		private void ButRestore_Click(object sender, System.EventArgs e)
		{
		    RestoreIcon(OldTrayInfo);
		}

        /*
		 '==============================================================================
         ' Method:       ButRefresh_Click
         '
         ' Description:  Refresh FillFlexGrid to display any new data
         '==============================================================================
		*/
		private void ButRefresh_Click(object sender, System.EventArgs e)
		{
	     	FillFlexGrid();
		}
        

		/*
		'==============================================================================
        ' Method:       RestoreIcon
        '
        ' Description:  Restore any icon you removed with Hide Method.
        '==============================================================================
        */
		private void RestoreIcon(ASTC.TrayIconInfo TrayIcon)
		{ 

			const int NIF_TIP = 0x4;
			const int NIF_ICON = 0x2;
			const int NIM_ADD = 0x0;
			const int NIF_MESSAGE = 0x1;

            Boolean Res;
			ShellNotify ShellIcon = new ShellNotify(); 
		    ShellNotify.NOTIFYICONDATA TrayI = new ShellNotify.NOTIFYICONDATA() ;


          if (TrayIcon.hwnd != 0)
			{   
			
			   TrayI.cbSize =  Marshal.SizeOf(TrayI);
               TrayI.hwnd = TrayIcon.hwnd;
			   TrayI.uID = TrayIcon.uId;
			   TrayI.hIcon = TrayIcon.hIcon;
      		   TrayI.szTip = TrayIcon.ToolTip;
			   TrayI.uFlags = NIF_ICON | NIF_TIP | NIF_MESSAGE;
		       TrayI.uCallbackMessage = TrayIcon.ucallbackMessage;
               //Restore the icon
			   Res = ShellNotify.Shell_NotifyIcon(NIM_ADD, ref TrayI);
			   MSFlexGrid1.Clear();
               FillFlexGrid();
			}
		}
	}
}
