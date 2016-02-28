/*
=========================================================================================
Publisher      CprinGold Software.
               http://www.cpringold.com
			   support@cpringold.com


 Description:  A sample code demonstrate how to add your software icon to system tray area
               As well as how to use the advanced features related to system tray area.
==========================================================================================
ASTC provide you with a very simple object, TrayIcon object which is the responsible to add the most of system
Tray functions to your applications.

With a few lines i can demonstrate how TrayIcon object can help you

Use AxTrayIcon1.Show method to add an icon to system tray area
Use AxTrayIcon1.Hide method to hide the icon you had add by TrayIcon.Show method from system tray
Use AxTrayIcon1. ChangeIcon method to change the icon you had added by another icon
Use AxTrayIcon1.Animate method to animate the icon in the system tray
Use AxTrayIcon1.StopAnimateing method to stop animating the icon
Use AxTrayIcon1.ShowBalloon to display a tooltip balloon
Use AxTrayIcon1.HideBalloon to hide the balloon
Use AxTrayIcon1.Popup to display a popup menu
Use AxTrayIcon1.TrackPopupmenu to track the popup menu whenever it lost focus in order to close it.
Use AxTrayIcon1.IsDisplayed to check if the icon is displayed in the system tray area

Note that you have a custom types of events associated with TrayIcon Object like
BalloonShow, BalloonHide, MedMouseUp, LeftMouseUp, RightMouseUp, MedMouseDown, LeftMouseDown
RightMouseDown, MedMouseDBLCLK, LeftMouseDBLCLK, BalloonLeftClick, BalloonRightClick, RightMouseDBLCLK ()

Note that TrayIcon Object doesn't subclass the window in which it contains TrayIcon control
because TrayIcon Object dynamically creates an internal window related to it's instance,
and destroy this window whenever you call Hide method or close your application,
so don 't worry about subclassing
If you have any further questions or need more sample code don't hesitate to contact us
At support@cpringold.com
*/
using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

namespace Project1
{
	/// <summary>
	/// Summary description for WinForm.
	/// </summary>
	public class WinForm : System.Windows.Forms.Form
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button butTrack;
		private System.Windows.Forms.Button butHide;
		private System.Windows.Forms.Button butAni;
		private System.Windows.Forms.Button butShowBalloon;
		private System.Windows.Forms.Button butShow;
		private System.Windows.Forms.Button butSTrack;
		private System.Windows.Forms.Button butCBalloon;
		private System.Windows.Forms.Button butChgIcon;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.Button butStAni;
		private AxMSComctlLib.AxImageList axImageList1;
		private System.Windows.Forms.ContextMenu contextMenu1;
		private System.Windows.Forms.MenuItem menuItem1;
		private System.Windows.Forms.MenuItem menuItem2;
		private AxASTC.AxTrayIcon axTrayIcon1;

		public WinForm()
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
		protected override void Dispose(bool disposing)
		{
			if (disposing)
			{
				if (components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose(disposing);
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(WinForm));
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.axTrayIcon1 = new AxASTC.AxTrayIcon();
			this.axImageList1 = new AxMSComctlLib.AxImageList();
			this.butStAni = new System.Windows.Forms.Button();
			this.butTrack = new System.Windows.Forms.Button();
			this.butHide = new System.Windows.Forms.Button();
			this.butAni = new System.Windows.Forms.Button();
			this.butShowBalloon = new System.Windows.Forms.Button();
			this.butShow = new System.Windows.Forms.Button();
			this.butSTrack = new System.Windows.Forms.Button();
			this.butCBalloon = new System.Windows.Forms.Button();
			this.butChgIcon = new System.Windows.Forms.Button();
			this.button1 = new System.Windows.Forms.Button();
			this.contextMenu1 = new System.Windows.Forms.ContextMenu();
			this.menuItem1 = new System.Windows.Forms.MenuItem();
			this.menuItem2 = new System.Windows.Forms.MenuItem();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.axTrayIcon1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.axImageList1)).BeginInit();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.axTrayIcon1);
			this.groupBox1.Controls.Add(this.axImageList1);
			this.groupBox1.Controls.Add(this.butStAni);
			this.groupBox1.Controls.Add(this.butTrack);
			this.groupBox1.Controls.Add(this.butHide);
			this.groupBox1.Controls.Add(this.butAni);
			this.groupBox1.Controls.Add(this.butShowBalloon);
			this.groupBox1.Controls.Add(this.butShow);
			this.groupBox1.Controls.Add(this.butSTrack);
			this.groupBox1.Controls.Add(this.butCBalloon);
			this.groupBox1.Controls.Add(this.butChgIcon);
			this.groupBox1.Controls.Add(this.button1);
			this.groupBox1.Location = new System.Drawing.Point(8, 8);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(200, 184);
			this.groupBox1.TabIndex = 9;
			this.groupBox1.TabStop = false;
			// 
			// axTrayIcon1
			// 
			this.axTrayIcon1.ContainingControl = this;
			this.axTrayIcon1.Enabled = true;
			this.axTrayIcon1.Location = new System.Drawing.Point(160, 144);
			this.axTrayIcon1.Name = "axTrayIcon1";
			this.axTrayIcon1.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("axTrayIcon1.OcxState")));
			this.axTrayIcon1.Size = new System.Drawing.Size(36, 36);
			this.axTrayIcon1.TabIndex = 21;
			this.axTrayIcon1.RightMouseUp += new System.EventHandler(this.axTrayIcon1_RightMouseUp);
			this.axTrayIcon1.BalloonLeftClick += new System.EventHandler(this.axTrayIcon1_BalloonLeftClick);
			// 
			// axImageList1
			// 
			this.axImageList1.ContainingControl = this;
			this.axImageList1.Enabled = true;
			this.axImageList1.Location = new System.Drawing.Point(8, 136);
			this.axImageList1.Name = "axImageList1";
			this.axImageList1.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("axImageList1.OcxState")));
			this.axImageList1.Size = new System.Drawing.Size(38, 38);
			this.axImageList1.TabIndex = 20;
			// 
			// butStAni
			// 
			this.butStAni.Location = new System.Drawing.Point(56, 144);
			this.butStAni.Name = "butStAni";
			this.butStAni.Size = new System.Drawing.Size(96, 24);
			this.butStAni.TabIndex = 18;
			this.butStAni.Text = "Stop Animation";
			this.butStAni.Click += new System.EventHandler(this.butStAni_Click);
			// 
			// butTrack
			// 
			this.butTrack.Location = new System.Drawing.Point(104, 80);
			this.butTrack.Name = "butTrack";
			this.butTrack.Size = new System.Drawing.Size(88, 24);
			this.butTrack.TabIndex = 17;
			this.butTrack.Text = "TrackPopUp";
			this.butTrack.Click += new System.EventHandler(this.butTrack_Click);
			// 
			// butHide
			// 
			this.butHide.Location = new System.Drawing.Point(8, 80);
			this.butHide.Name = "butHide";
			this.butHide.Size = new System.Drawing.Size(88, 24);
			this.butHide.TabIndex = 16;
			this.butHide.Text = "Hide Icon";
			this.butHide.Click += new System.EventHandler(this.butHide_Click);
			// 
			// butAni
			// 
			this.butAni.Location = new System.Drawing.Point(8, 112);
			this.butAni.Name = "butAni";
			this.butAni.Size = new System.Drawing.Size(88, 24);
			this.butAni.TabIndex = 15;
			this.butAni.Text = "Animate";
			this.butAni.Click += new System.EventHandler(this.butAni_Click);
			// 
			// butShowBalloon
			// 
			this.butShowBalloon.Location = new System.Drawing.Point(104, 16);
			this.butShowBalloon.Name = "butShowBalloon";
			this.butShowBalloon.Size = new System.Drawing.Size(88, 24);
			this.butShowBalloon.TabIndex = 14;
			this.butShowBalloon.Text = "Show Balloon";
			this.butShowBalloon.Click += new System.EventHandler(this.butShowBalloon_Click);
			// 
			// butShow
			// 
			this.butShow.Location = new System.Drawing.Point(8, 16);
			this.butShow.Name = "butShow";
			this.butShow.Size = new System.Drawing.Size(88, 24);
			this.butShow.TabIndex = 13;
			this.butShow.Text = "&Show Icon";
			this.butShow.Click += new System.EventHandler(this.butShow_Click);
			// 
			// butSTrack
			// 
			this.butSTrack.Location = new System.Drawing.Point(104, 112);
			this.butSTrack.Name = "butSTrack";
			this.butSTrack.Size = new System.Drawing.Size(88, 24);
			this.butSTrack.TabIndex = 12;
			this.butSTrack.Text = "Stop Track";
			this.butSTrack.Click += new System.EventHandler(this.butSTrack_Click);
			// 
			// butCBalloon
			// 
			this.butCBalloon.Location = new System.Drawing.Point(104, 48);
			this.butCBalloon.Name = "butCBalloon";
			this.butCBalloon.Size = new System.Drawing.Size(88, 24);
			this.butCBalloon.TabIndex = 11;
			this.butCBalloon.Text = "Close Balloon";
			this.butCBalloon.Click += new System.EventHandler(this.butCBalloon_Click);
			// 
			// butChgIcon
			// 
			this.butChgIcon.Location = new System.Drawing.Point(8, 48);
			this.butChgIcon.Name = "butChgIcon";
			this.butChgIcon.Size = new System.Drawing.Size(88, 24);
			this.butChgIcon.TabIndex = 10;
			this.butChgIcon.Text = "Change Icon";
			this.butChgIcon.Click += new System.EventHandler(this.butChgIcon_Click);
			// 
			// button1
			// 
			this.button1.Location = new System.Drawing.Point(-29, -48);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(80, 42);
			this.button1.TabIndex = 9;
			this.button1.Text = "button1";
			// 
			// contextMenu1
			// 
			this.contextMenu1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
						this.menuItem1,
						this.menuItem2});
			// 
			// menuItem1
			// 
			this.menuItem1.Index = 0;
			this.menuItem1.Text = "Main";
			this.menuItem1.Click += new System.EventHandler(this.menuItem1_Click);
			// 
			// menuItem2
			// 
			this.menuItem2.Index = 1;
			this.menuItem2.Text = "Exit";
			this.menuItem2.Click += new System.EventHandler(this.menuItem2_Click);
			// 
			// WinForm
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(216, 198);
			this.Controls.Add(this.groupBox1);
			this.Name = "WinForm";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "C#Builder TrayIcon";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.WinForm_Closing);
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.axTrayIcon1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.axImageList1)).EndInit();
			this.ResumeLayout(false);
		}
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new WinForm());
		}

		/*
		==============================================================================
		 Method:        butShow_Click
		
		 Description:  Shows the icon in system tray.
		==============================================================================
		*/
		private void butShow_Click(object sender, System.EventArgs e)
		{
		   /*
			Unlike the previous version of ASTC, you don't need to pass a window handle to Show function because
            ASTC dynamically creates an internal window related to ASTC instance, and destroy this window whenever
            you call Hide method or close your application.
		   */
			axTrayIcon1.CtlShow(this.Icon.Handle.ToInt32(),"C#Builder TrayIcon");

		}

		/*
		==============================================================================
		 Method:        butChgIcon_Click
		
		 Description:  Changes the icon in system tray area.
		==============================================================================
		*/
		private void butChgIcon_Click(object sender, System.EventArgs e)
		{
			//change the icon in system tray area
			//by another one in ImageList
			Object i = 2;     
			axTrayIcon1.ChangeIcon(axImageList1.ListImages.get_Item(ref i).ExtractIcon().Handle);
		}

		/*
		==============================================================================
		 Method:       butHide_Click
		
		 Description:  Removes the icon from system tray.
		==============================================================================
		*/
		private void butHide_Click(object sender, System.EventArgs e)
		{
		   axTrayIcon1.CtlHide();
		}

		/*
		'==================================================================================
		' Method:        butAni_Click
		'
		' Description:  Animate the icon in system tray using in icons in imagelist control.
		'==================================================================================
		*/
		private void butAni_Click(object sender, System.EventArgs e)
		{
			//Only if the icon is not Animated
			if (!axTrayIcon1.AnimateState)  
			   axTrayIcon1.Animate(axImageList1,1);
			/*
			    Start animating the icons in system tray
                we use the icons in the imagelist control
                Then
                We pass the ImageList1 control to Animate method
                '''you can replace the icons with your own'''
			*/

		}

		/*
		'==============================================================================
		' Method:       butStAni_Click
		'
		' Description:  Stop animating the icons in system tray.
		'==============================================================================
		*/
		private void butStAni_Click(object sender, System.EventArgs e)
		{
			 if (axTrayIcon1.AnimateState)  
			  axTrayIcon1.StopAnimateing(this.Icon.Handle.ToInt32());
		}

		/*
		'==============================================================================
		' Method:       butShowBalloon_Click
		'
		' Description:  Displays a balloon tooltip points to the icon.
		'==============================================================================
		*/
		private void butShowBalloon_Click(object sender, System.EventArgs e)
		{
			axTrayIcon1.ShowBalloon("C#Builder TrayIcon Balloon Tooltip.","My program",ASTC.BIcoType.Info ,5000);
		}

		/*
		'==============================================================================
		' Method:       butCBalloon_Click
		'
		' Description:  Close the balloon tooltip.
		'==============================================================================
		*/
		private void butCBalloon_Click(object sender, System.EventArgs e)
		{
			axTrayIcon1.CloseBalloon();
		}

		/*
		'==============================================================================
		' Method:       butTrack_Click
		'
		' Description:  Start Tracking the popup menu in system tray area.
		'==============================================================================
		*/
		private void butTrack_Click(object sender, System.EventArgs e)
		{
			axTrayIcon1.TrackPopMenu = true;
			butSTrack.Enabled = true; 
			MessageBox.Show(this,"when you right click the icon a popupmenu will appear,ASTC will Track the popupmenu and close it when you left click again");

		}

		/*
		'==============================================================================
		' Method:       butSTrack_Click
		'
		' Description:  Stop track the popup menu in system tray area.
		'==============================================================================
		*/
		private void butSTrack_Click(object sender, System.EventArgs e)
		{
			axTrayIcon1.TrackPopMenu = false;
			butSTrack.Enabled = false;
			MessageBox.Show(this, "ASTC stop tracking the popupmenu,Now right click the icon again and you will see that popup menu willn't close unless you click an item.");
		}

		/*
		'==============================================================================
		' Method:       WinForm_Closing
		'
		' Description:  Removes the icon from system tray whenever the
		'               terminates the program.
		'==============================================================================
		*/
		private void WinForm_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
		   if (axTrayIcon1.IsDisplayed)
			axTrayIcon1.CtlHide();
		}
		
		private void menuItem1_Click(object sender, System.EventArgs e)
		{
			MessageBox.Show("Main menu Item had been clicked.");
		}

		/*
		'==============================================================================
		' Method:       AxTrayIcon1_BalloonLeftClick
		'
		' Description:  Display a message box whenever you left click the balloon tooltip.
		'==============================================================================
		*/
		private void axTrayIcon1_BalloonLeftClick(object sender, System.EventArgs e)
		{
			 MessageBox.Show("The Balloon Tooltip had been left clicked.");
		}

		/*
		'====================================================================================
		' Method:       AxTrayIcon1_RightMouseUp
		'
		' Description:  Display a popup menu whenever you right click the icon in system tray.
		'====================================================================================
		*/
		private void axTrayIcon1_RightMouseUp(object sender, System.EventArgs e)
		{
			 axTrayIcon1.PopUp(contextMenu1.Handle.ToInt32(),this.Handle.ToInt32());
		}

		/*
		'==============================================================================
		' Method:        MenuItem2_Click
		'
		' Description:  Removes the icon from system tray and ends the program.
		'==============================================================================
		*/
		private void menuItem2_Click(object sender, System.EventArgs e)
		{
		     if (axTrayIcon1.IsDisplayed)  
				axTrayIcon1.CtlHide();  
			 this.Close();
		}
		

	}
}
