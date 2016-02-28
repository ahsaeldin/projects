/*
=========================================================================================
Publisher      CprinGold Software.
               http://www.cpringold.com
               support@cpringold.com


 Description:  A sample code demonstrates how to use Balloon Object.
==========================================================================================

Use the Optional parameter PHwnd to insure that the balloon tooltip will not appear outside a specified
Window or a child window region like a command button or a text box.
Note that:

1.you can't set this parameter to any external window i.e. you can't set to an external Apps window

2.you can set this parameter to the window which contains the current instance of the balloon control

i.e.  For the following code

balloon1.ShowBalloon From1.hwnd

The hwnd is a handle to the window in which contains the Balloon control and that's means that the balloon
Will only appear if the mouse is over the Form itself not any of its child controls [child windows] in which
It contains.

3.you can set this parameter to any control of the window that contains the current instance of the balloon
Control i.e if you have a Form called Form1 which has a command button called Command1 and a text box called
text1, you can set the parameter as the following

balloon1.ShowBalloon command1.hwnd

balloon1.ShowBalloon text1.hwnd

And that's means the balloon will only appear if the mouse is over the control itself

4. If you didn't pass this optional parameter, Balloon1 object will use its parent window automatically
I.e. if you have a form called Form1, Balloon object will Form1.Hwnd for PHwnd parameter silently but in
This case the balloon tooltip may appear at any part of the Form or its Controls

Use DelayTime parameter to set the Delay Time before the balloon appear, the default value for this
Parameter is 2000 milliseconds and the maximum, value for DelayTime Parameter is 65,535 milliseconds,
Is equivalent to just over 1 minute?
The maximum, value for Timeout Parameter is 65,535 milliseconds, is equivalent to just over 1 minute.
*/
using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

namespace Balloon_Tooltip_Sample
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class Form1 : System.Windows.Forms.Form
	{
		internal System.Windows.Forms.GroupBox GroupBox1;
		internal System.Windows.Forms.Button ButBalloon3;
		internal System.Windows.Forms.Button ButBalloon2;
		internal System.Windows.Forms.Button ButBalloon1;
		internal System.Windows.Forms.Label Label1;
		private AxASTC.AxBalloon Balloon1;
		private AxASTC.AxBalloon Balloon2;
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
			this.GroupBox1 = new System.Windows.Forms.GroupBox();
			this.ButBalloon3 = new System.Windows.Forms.Button();
			this.ButBalloon2 = new System.Windows.Forms.Button();
			this.ButBalloon1 = new System.Windows.Forms.Button();
			this.Label1 = new System.Windows.Forms.Label();
			this.Balloon1 = new AxASTC.AxBalloon();
			this.Balloon2 = new AxASTC.AxBalloon();
			this.GroupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.Balloon1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.Balloon2)).BeginInit();
			this.SuspendLayout();
			// 
			// GroupBox1
			// 
			this.GroupBox1.Controls.Add(this.Balloon2);
			this.GroupBox1.Controls.Add(this.Balloon1);
			this.GroupBox1.Controls.Add(this.ButBalloon3);
			this.GroupBox1.Controls.Add(this.ButBalloon2);
			this.GroupBox1.Controls.Add(this.ButBalloon1);
			this.GroupBox1.Controls.Add(this.Label1);
			this.GroupBox1.Location = new System.Drawing.Point(10, 8);
			this.GroupBox1.Name = "GroupBox1";
			this.GroupBox1.Size = new System.Drawing.Size(272, 112);
			this.GroupBox1.TabIndex = 1;
			this.GroupBox1.TabStop = false;
			this.GroupBox1.MouseHover += new System.EventHandler(this.GroupBox1_MouseHover);
			// 
			// ButBalloon3
			// 
			this.ButBalloon3.Location = new System.Drawing.Point(192, 56);
			this.ButBalloon3.Name = "ButBalloon3";
			this.ButBalloon3.TabIndex = 7;
			this.ButBalloon3.Text = "Balloon3";
			this.ButBalloon3.MouseMove += new System.Windows.Forms.MouseEventHandler(this.ButBalloon3_MouseMove);
			// 
			// ButBalloon2
			// 
			this.ButBalloon2.Location = new System.Drawing.Point(96, 80);
			this.ButBalloon2.Name = "ButBalloon2";
			this.ButBalloon2.TabIndex = 6;
			this.ButBalloon2.Text = "Balloon2";
			this.ButBalloon2.MouseMove += new System.Windows.Forms.MouseEventHandler(this.ButBalloon2_MouseMove);
			// 
			// ButBalloon1
			// 
			this.ButBalloon1.Location = new System.Drawing.Point(8, 56);
			this.ButBalloon1.Name = "ButBalloon1";
			this.ButBalloon1.TabIndex = 5;
			this.ButBalloon1.Text = "Balloon1";
			this.ButBalloon1.MouseMove += new System.Windows.Forms.MouseEventHandler(this.ButBalloon1_MouseMove);
			// 
			// Label1
			// 
			this.Label1.ForeColor = System.Drawing.Color.Blue;
			this.Label1.Location = new System.Drawing.Point(8, 8);
			this.Label1.Name = "Label1";
			this.Label1.Size = new System.Drawing.Size(256, 96);
			this.Label1.TabIndex = 0;
			this.Label1.Text = "Move the Cursor to every command button to see the balloon tooltip.";
			this.Label1.MouseMove += new System.Windows.Forms.MouseEventHandler(this.Label1_MouseMove);
			// 
			// Balloon1
			// 
			this.Balloon1.ContainingControl = this;
			this.Balloon1.Enabled = true;
			this.Balloon1.Location = new System.Drawing.Point(56, 80);
			this.Balloon1.Name = "Balloon1";
			this.Balloon1.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("Balloon1.OcxState")));
			this.Balloon1.Size = new System.Drawing.Size(27, 27);
			this.Balloon1.TabIndex = 8;
			this.Balloon1.BalloonLeftClick += new System.EventHandler(this.Balloon1_BalloonLeftClick);
			// 
			// Balloon2
			// 
			this.Balloon2.ContainingControl = this;
			this.Balloon2.Enabled = true;
			this.Balloon2.Location = new System.Drawing.Point(200, 24);
			this.Balloon2.Name = "Balloon2";
			this.Balloon2.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("Balloon2.OcxState")));
			this.Balloon2.Size = new System.Drawing.Size(27, 27);
			this.Balloon2.TabIndex = 9;
			// 
			// Form1
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(292, 126);
			this.Controls.Add(this.GroupBox1);
			this.Name = "Form1";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "VC#.Net Balloon Tooltip";
			this.TopMost = true;
			this.MouseMove += new System.Windows.Forms.MouseEventHandler(this.Form1_MouseMove);
			this.GroupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.Balloon1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.Balloon2)).EndInit();
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
		 ===============================================================================
		 Method:        Form1_MouseMove
		
		 Description:   Destroy the Balloon Tooltip.
	     ================================================================================
		*/
		private void Form1_MouseMove(object sender, System.Windows.Forms.MouseEventArgs e)
		{
		   Balloon1.Destroy(); 
		   Balloon2.Destroy();  
		}
		/*
		 ===============================================================================
		 Method:        GroupBox1_MouseHover
		
		 Description:   Destroy the Balloon Tooltip.
		 ================================================================================
		*/
		private void GroupBox1_MouseHover(object sender, System.EventArgs e)
		{
			Balloon1.Destroy(); 
			Balloon2.Destroy(); 
		}

		/*
		 ===============================================================================
		 Method:        Label1_MouseMove
		
		 Description:   Destroy the Balloon Tooltip.
		 ================================================================================
		*/
		private void Label1_MouseMove(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			Balloon1.Destroy(); 
			Balloon2.Destroy(); 
		}

		/*
		 ==============================================================================
		 Method:       ButBalloon1_MouseMove
		
		 Description:  Display a balloon tooltip for ButBalloon1.
		 ===============================================================================
		*/
		private void ButBalloon1_MouseMove(object sender, System.Windows.Forms.MouseEventArgs e)
		{
		    int ButHandle;
			ButHandle = ButBalloon1.Handle.ToInt32();
			Balloon1.Style = ASTC.BalloonStyle.BalloonType; 
			Balloon1.ShowBalloon(ref ButHandle,"Click Me","Balloon1",ASTC.BIcoType.Info
				,1000,5000);
        }

		/*
		 ==============================================================================
		 Method:       ButBalloon2_MouseMove
		
		 Description:  Display a balloon tooltip for ButBalloon2.
		 ===============================================================================
		*/
		private void ButBalloon2_MouseMove(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			  int ButHandle;
			  ButHandle = ButBalloon2.Handle.ToInt32();
			  Balloon1.Style = ASTC.BalloonStyle.RectangleType;
			  Balloon1.ShowBalloon(ref ButHandle,"Rectangle Type","Balloon2",ASTC.BIcoType.Info
				  ,1000,5000);
		}

		/*
		 ==============================================================================
		 Method:       ButBalloon3_MouseMove
		
		 Description:  Display a balloon tooltip for ButBalloon3.
		 ===============================================================================
		*/
		private void ButBalloon3_MouseMove(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			int ButHandle;
			ButHandle = ButBalloon3.Handle.ToInt32();
			
			Balloon2.CtlForeColor = 9843455;
			Balloon2.CtlBackColor = 16777215;
			Balloon2.ShowBalloon(ref ButHandle,"Customizable balloon tooltip","Balloon3",ASTC.BIcoType.Info
				,1000,5000);
		}

		/*
		 ==============================================================================
		' Method:        Balloon1_BalloonLeftClick
		'
		' Description:   Display a message Box when the left click the balloon.
		 ==============================================================================
		*/
		private void Balloon1_BalloonLeftClick(object sender, System.EventArgs e)
		{
			MessageBox.Show("The Balloon Tooltip had been Clicked");  
		}

	}
}
