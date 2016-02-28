/*
'=========================================================================================
'Publisher      CprinGold Software.
'               http://www.cpringold.com
'               support@cpringold.com
'
'
'Description:  A sample code demonstrates how to use ASTC control to subclass any Form.
'==========================================================================================
'First of all what is the Subclass?
'Subclassing is the processing of intercepting Windows messages that your program normally wouldn't
'Receive, so it extends your ability to process more windows messages and add new features to your program.
''
'With a few lines i can demonstrate how Subclass object can help you
'
'1.use SubClass.BeginSubClass to start subclassing a specified window.
'2.use SubClass.EndSubClass to end subclassing.
'3.use SubClass_MessageReceived event to process the message you want.
'
'About this sample code
'This sample code show you the benefits of subclassing by demonstrate
'How to disable the context menu in a text box?
'How to extend a listbox functions?

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
namespace SubClass_Sample
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class Form1 : System.Windows.Forms.Form
	{   
		int OldItem;
		internal System.Windows.Forms.GroupBox GroupBox1;
		internal System.Windows.Forms.Label Label3;
		internal System.Windows.Forms.ListBox ListBox1;
		internal System.Windows.Forms.Label Label2;
		internal System.Windows.Forms.Label Label1;
		internal System.Windows.Forms.Button ButLUnSubClass;
		internal System.Windows.Forms.Button ButLSubClass;
		private AxASTC.AxSubClass SubClass1;
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
			this.Label3 = new System.Windows.Forms.Label();
			this.ListBox1 = new System.Windows.Forms.ListBox();
			this.Label2 = new System.Windows.Forms.Label();
			this.Label1 = new System.Windows.Forms.Label();
			this.ButLUnSubClass = new System.Windows.Forms.Button();
			this.ButLSubClass = new System.Windows.Forms.Button();
			this.SubClass1 = new AxASTC.AxSubClass();
			this.Balloon1 = new AxASTC.AxBalloon();
			this.Balloon2 = new AxASTC.AxBalloon();
			this.GroupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.SubClass1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.Balloon1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.Balloon2)).BeginInit();
			this.SuspendLayout();
			// 
			// GroupBox1
			// 
			this.GroupBox1.Controls.Add(this.Balloon2);
			this.GroupBox1.Controls.Add(this.Balloon1);
			this.GroupBox1.Controls.Add(this.SubClass1);
			this.GroupBox1.Controls.Add(this.Label3);
			this.GroupBox1.Controls.Add(this.ListBox1);
			this.GroupBox1.Controls.Add(this.Label2);
			this.GroupBox1.Controls.Add(this.Label1);
			this.GroupBox1.Controls.Add(this.ButLUnSubClass);
			this.GroupBox1.Controls.Add(this.ButLSubClass);
			this.GroupBox1.ForeColor = System.Drawing.Color.Blue;
			this.GroupBox1.Location = new System.Drawing.Point(8, 8);
			this.GroupBox1.Name = "GroupBox1";
			this.GroupBox1.Size = new System.Drawing.Size(408, 232);
			this.GroupBox1.TabIndex = 1;
			this.GroupBox1.TabStop = false;
			this.GroupBox1.MouseHover += new System.EventHandler(this.GroupBox1_MouseHover);
			// 
			// Label3
			// 
			this.Label3.ForeColor = System.Drawing.Color.Red;
			this.Label3.Location = new System.Drawing.Point(40, 160);
			this.Label3.Name = "Label3";
			this.Label3.Size = new System.Drawing.Size(256, 32);
			this.Label3.TabIndex = 5;
			// 
			// ListBox1
			// 
			this.ListBox1.Items.AddRange(new object[] {
														  "item0",
														  "item1",
														  "item2",
														  "item3"});
			this.ListBox1.Location = new System.Drawing.Point(40, 96);
			this.ListBox1.Name = "ListBox1";
			this.ListBox1.Size = new System.Drawing.Size(320, 56);
			this.ListBox1.TabIndex = 4;
			// 
			// Label2
			// 
			this.Label2.ForeColor = System.Drawing.Color.Red;
			this.Label2.Location = new System.Drawing.Point(40, 64);
			this.Label2.Name = "Label2";
			this.Label2.Size = new System.Drawing.Size(320, 32);
			this.Label2.TabIndex = 3;
			this.Label2.Text = "Move the Cursor between listBox items before and after you click SubClass command" +
				" Button and watch the difference.";
			// 
			// Label1
			// 
			this.Label1.ForeColor = System.Drawing.Color.Blue;
			this.Label1.Location = new System.Drawing.Point(8, 16);
			this.Label1.Name = "Label1";
			this.Label1.Size = new System.Drawing.Size(392, 40);
			this.Label1.TabIndex = 2;
			this.Label1.Text = "ListBox control doesn\'t provides a mousemove event in which you can use for displ" +
				"aying a balloon tooltip for the list Box, but if you we subclass the listbox we " +
				"can assign a Balloon tooltip for every item in the listbox.";
			// 
			// ButLUnSubClass
			// 
			this.ButLUnSubClass.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(64)));
			this.ButLUnSubClass.Location = new System.Drawing.Point(208, 200);
			this.ButLUnSubClass.Name = "ButLUnSubClass";
			this.ButLUnSubClass.TabIndex = 2;
			this.ButLUnSubClass.Text = "UnSubClass";
			this.ButLUnSubClass.Click += new System.EventHandler(this.ButLUnSubClass_Click);
			// 
			// ButLSubClass
			// 
			this.ButLSubClass.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(64)));
			this.ButLSubClass.Location = new System.Drawing.Point(128, 200);
			this.ButLSubClass.Name = "ButLSubClass";
			this.ButLSubClass.TabIndex = 1;
			this.ButLSubClass.Text = "SubClass";
			this.ButLSubClass.Click += new System.EventHandler(this.ButLSubClass_Click);
			// 
			// SubClass1
			// 
			this.SubClass1.ContainingControl = this;
			this.SubClass1.Enabled = true;
			this.SubClass1.Location = new System.Drawing.Point(8, 192);
			this.SubClass1.Name = "SubClass1";
			this.SubClass1.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("SubClass1.OcxState")));
			this.SubClass1.Size = new System.Drawing.Size(36, 36);
			this.SubClass1.TabIndex = 6;
			this.SubClass1.MessageReceived += new AxASTC.__SubClass_MessageReceivedEventHandler(this.axSubClass1_MessageReceived);
			// 
			// Balloon1
			// 
			this.Balloon1.ContainingControl = this;
			this.Balloon1.Enabled = true;
			this.Balloon1.Location = new System.Drawing.Point(376, 168);
			this.Balloon1.Name = "Balloon1";
			this.Balloon1.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("Balloon1.OcxState")));
			this.Balloon1.Size = new System.Drawing.Size(27, 27);
			this.Balloon1.TabIndex = 7;
			// 
			// Balloon2
			// 
			this.Balloon2.ContainingControl = this;
			this.Balloon2.Enabled = true;
			this.Balloon2.Location = new System.Drawing.Point(376, 200);
			this.Balloon2.Name = "Balloon2";
			this.Balloon2.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("Balloon2.OcxState")));
			this.Balloon2.Size = new System.Drawing.Size(27, 27);
			this.Balloon2.TabIndex = 8;
			this.Balloon2.BalloonLeftClick += new System.EventHandler(this.Balloon2_BalloonLeftClick);
			// 
			// Form1
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(424, 246);
			this.Controls.Add(this.GroupBox1);
			this.Name = "Form1";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "VC#.NET SubClass";
			this.GroupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.SubClass1)).EndInit();
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
		'==============================================================================
		' Method:        ButLSubClass_Click
		'
		' Description:  Start subclassing List1.
		'==============================================================================
        */
		private void ButLSubClass_Click(object sender, System.EventArgs e)
		{   
			SubClass1.BeginSubClass(ListBox1.Handle.ToInt32());
		}
        
		/*
		'==============================================================================
		' Method:        axSubClass1_MessageReceived
		'
		' Description:   Extends the listbox functionality.
		'==============================================================================
        */
		private void axSubClass1_MessageReceived(object sender, AxASTC.__SubClass_MessageReceivedEvent e)
		{
		  /*
		   'e.hwnd()        Handle of the subclassed window or control.
           'e.msg()         The ID of the intercepted message.
           'e.wParam()      The wParam value of the intercepted message.
           'e.lParam()      The lParam value of the intercepted message.
           'e.cancel        Further processing state.
		  */

          Int32 ret;
          
		   
          int ButHandle;

		  const int WM_MOUSEMOVE = 512;
          const int LB_ITEMFROMPOINT = 425;		
	      
		  

			if (e.msg ==  WM_MOUSEMOVE)
			{
			  //Retrieve the zero-based index of the item nearest a specified point in the list box.
              ret = ShellNotify.SendMessage(e.hwnd, LB_ITEMFROMPOINT, 0, e.lParam);
				for (int g = 0; g != (int)ListBox1.Items.Count + 1; g++)
				{
					if (ret == g)
					{   
						if (OldItem != g)
						{
						  Balloon1.Destroy();
                          Balloon2.Destroy();
						}

						ListBox1.SetSelected(g,true);  
				        Label3.Text = "item" + ListBox1.GetItemText(ListBox1.SelectedIndex) + " had been hovered.";
					    
						ButHandle = ListBox1.Handle.ToInt32();
						    
						if (g == 3)
						{
							Balloon1.Destroy();     
							Balloon2.ShowBalloon(ref ButHandle, "Click Me.", "SubClass", ASTC.BIcoType.Info, 500, 5000);
						}
						else
							Balloon1.ShowBalloon(ref ButHandle, "item" + ListBox1.GetItemText(ListBox1.SelectedIndex) + " had been hovered.", "SubClass", ASTC.BIcoType.Info, 500, 5000);
                   
					    OldItem = g;
					}
                     
				}
             }
			 
         }

        /*
		'==============================================================================
		' Method:        ButLUnSubClass_Click
		'
		' Description:  Stop Subclassing ListBOx1.
		'==============================================================================
        */
		private void ButLUnSubClass_Click(object sender, System.EventArgs e)
		{
            SubClass1.EndSubClass();
            Label3.Text = "";
		}

		private void Balloon2_BalloonLeftClick(object sender, System.EventArgs e)
		{
			MessageBox.Show("item" + ListBox1.GetItemText(ListBox1.SelectedIndex) + " had been clicked.");  
		}

		private void GroupBox1_MouseHover(object sender, System.EventArgs e)
		{
			Balloon1.Destroy();
			Balloon2.Destroy();
		}
    }
}
