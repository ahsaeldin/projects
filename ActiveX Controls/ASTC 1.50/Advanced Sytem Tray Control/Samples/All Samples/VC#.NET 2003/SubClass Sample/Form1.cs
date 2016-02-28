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
'How to stop the form from being resized below or above a user-defined amount?

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
		internal System.Windows.Forms.GroupBox GroupBox1;
		internal System.Windows.Forms.Label Label3;
		internal System.Windows.Forms.ListBox ListBox1;
		internal System.Windows.Forms.Label Label2;
		internal System.Windows.Forms.Label Label1;
		internal System.Windows.Forms.Button ButLUnSubClass;
		internal System.Windows.Forms.Button ButLSubClass;
		internal System.Windows.Forms.GroupBox GroupBox2;
		internal System.Windows.Forms.Button ButUnSubCont;
		internal System.Windows.Forms.Button ButSubContext;
		internal System.Windows.Forms.Label Label5;
		internal System.Windows.Forms.TextBox TextBox1;
		internal System.Windows.Forms.Label Label4;
		internal System.Windows.Forms.ContextMenu ContextMenu1;
		internal System.Windows.Forms.MenuItem MenuItem1;
		internal System.Windows.Forms.GroupBox groupBox3;
		internal System.Windows.Forms.Button butUnSubF;
		internal System.Windows.Forms.Button butFSubClass;
		internal System.Windows.Forms.Label label7;
		private AxASTC.AxSubClass axSubClass1;
		private AxASTC.AxSubClass axSubClass2;
		private AxASTC.AxSubClass axSubClass3;
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
			this.GroupBox2 = new System.Windows.Forms.GroupBox();
			this.ButUnSubCont = new System.Windows.Forms.Button();
			this.ButSubContext = new System.Windows.Forms.Button();
			this.Label5 = new System.Windows.Forms.Label();
			this.TextBox1 = new System.Windows.Forms.TextBox();
			this.Label4 = new System.Windows.Forms.Label();
			this.ContextMenu1 = new System.Windows.Forms.ContextMenu();
			this.MenuItem1 = new System.Windows.Forms.MenuItem();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.butUnSubF = new System.Windows.Forms.Button();
			this.butFSubClass = new System.Windows.Forms.Button();
			this.label7 = new System.Windows.Forms.Label();
			this.axSubClass1 = new AxASTC.AxSubClass();
			this.axSubClass2 = new AxASTC.AxSubClass();
			this.axSubClass3 = new AxASTC.AxSubClass();
			this.GroupBox1.SuspendLayout();
			this.GroupBox2.SuspendLayout();
			this.groupBox3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.axSubClass1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.axSubClass2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.axSubClass3)).BeginInit();
			this.SuspendLayout();
			// 
			// GroupBox1
			// 
			this.GroupBox1.Controls.Add(this.axSubClass1);
			this.GroupBox1.Controls.Add(this.Label3);
			this.GroupBox1.Controls.Add(this.ListBox1);
			this.GroupBox1.Controls.Add(this.Label2);
			this.GroupBox1.Controls.Add(this.Label1);
			this.GroupBox1.Controls.Add(this.ButLUnSubClass);
			this.GroupBox1.Controls.Add(this.ButLSubClass);
			this.GroupBox1.ForeColor = System.Drawing.Color.Blue;
			this.GroupBox1.Location = new System.Drawing.Point(10, 9);
			this.GroupBox1.Name = "GroupBox1";
			this.GroupBox1.Size = new System.Drawing.Size(270, 248);
			this.GroupBox1.TabIndex = 1;
			this.GroupBox1.TabStop = false;
			this.GroupBox1.Text = "SubClass Sample 1";
			// 
			// Label3
			// 
			this.Label3.ForeColor = System.Drawing.Color.Red;
			this.Label3.Location = new System.Drawing.Point(8, 184);
			this.Label3.Name = "Label3";
			this.Label3.Size = new System.Drawing.Size(256, 32);
			this.Label3.TabIndex = 5;
			this.Label3.Text = "Then click SubClass command and Right Click any item and you will see the differe" +
				"nce.";
			// 
			// ListBox1
			// 
			this.ListBox1.Items.AddRange(new object[] {
														  "item0",
														  "item1",
														  "item2",
														  "item3"});
			this.ListBox1.Location = new System.Drawing.Point(8, 128);
			this.ListBox1.Name = "ListBox1";
			this.ListBox1.Size = new System.Drawing.Size(256, 56);
			this.ListBox1.TabIndex = 4;
			// 
			// Label2
			// 
			this.Label2.ForeColor = System.Drawing.Color.Red;
			this.Label2.Location = new System.Drawing.Point(8, 96);
			this.Label2.Name = "Label2";
			this.Label2.Size = new System.Drawing.Size(256, 32);
			this.Label2.TabIndex = 3;
			this.Label2.Text = "Right Click any item in listbox before you click the Subclass command button.";
			// 
			// Label1
			// 
			this.Label1.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(64)));
			this.Label1.Location = new System.Drawing.Point(8, 16);
			this.Label1.Name = "Label1";
			this.Label1.Size = new System.Drawing.Size(256, 72);
			this.Label1.TabIndex = 2;
			this.Label1.Text = "In the Normal ListBox you can\'t assign a popup menu to any item so when the user " +
				"right click the item it will appears                                            " +
				"        But if you we use Subclassing we can achieve this.";
			// 
			// ButLUnSubClass
			// 
			this.ButLUnSubClass.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(64)));
			this.ButLUnSubClass.Location = new System.Drawing.Point(144, 216);
			this.ButLUnSubClass.Name = "ButLUnSubClass";
			this.ButLUnSubClass.TabIndex = 2;
			this.ButLUnSubClass.Text = "UnSubClass";
			this.ButLUnSubClass.Click += new System.EventHandler(this.ButLUnSubClass_Click);
			// 
			// ButLSubClass
			// 
			this.ButLSubClass.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(64)));
			this.ButLSubClass.Location = new System.Drawing.Point(56, 216);
			this.ButLSubClass.Name = "ButLSubClass";
			this.ButLSubClass.TabIndex = 1;
			this.ButLSubClass.Text = "SubClass";
			this.ButLSubClass.Click += new System.EventHandler(this.ButLSubClass_Click);
			// 
			// GroupBox2
			// 
			this.GroupBox2.Controls.Add(this.axSubClass2);
			this.GroupBox2.Controls.Add(this.ButUnSubCont);
			this.GroupBox2.Controls.Add(this.ButSubContext);
			this.GroupBox2.Controls.Add(this.Label5);
			this.GroupBox2.Controls.Add(this.TextBox1);
			this.GroupBox2.Controls.Add(this.Label4);
			this.GroupBox2.ForeColor = System.Drawing.Color.Blue;
			this.GroupBox2.Location = new System.Drawing.Point(8, 264);
			this.GroupBox2.Name = "GroupBox2";
			this.GroupBox2.Size = new System.Drawing.Size(272, 216);
			this.GroupBox2.TabIndex = 2;
			this.GroupBox2.TabStop = false;
			this.GroupBox2.Text = "SubClass Sample 2";
			// 
			// ButUnSubCont
			// 
			this.ButUnSubCont.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(64)));
			this.ButUnSubCont.Location = new System.Drawing.Point(144, 176);
			this.ButUnSubCont.Name = "ButUnSubCont";
			this.ButUnSubCont.TabIndex = 11;
			this.ButUnSubCont.Text = "UnSubClass";
			this.ButUnSubCont.Click += new System.EventHandler(this.ButUnSubCont_Click);
			// 
			// ButSubContext
			// 
			this.ButSubContext.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(64)));
			this.ButSubContext.Location = new System.Drawing.Point(56, 176);
			this.ButSubContext.Name = "ButSubContext";
			this.ButSubContext.TabIndex = 10;
			this.ButSubContext.Text = "SubClass";
			this.ButSubContext.Click += new System.EventHandler(this.ButSubContext_Click);
			// 
			// Label5
			// 
			this.Label5.BackColor = System.Drawing.Color.White;
			this.Label5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.Label5.Location = new System.Drawing.Point(56, 136);
			this.Label5.Name = "Label5";
			this.Label5.Size = new System.Drawing.Size(160, 32);
			this.Label5.TabIndex = 2;
			// 
			// TextBox1
			// 
			this.TextBox1.Location = new System.Drawing.Point(56, 64);
			this.TextBox1.Multiline = true;
			this.TextBox1.Name = "TextBox1";
			this.TextBox1.Size = new System.Drawing.Size(160, 64);
			this.TextBox1.TabIndex = 1;
			this.TextBox1.Text = "Right Click me before and after you click Subclass command  button and watch the " +
				"difference.";
			// 
			// Label4
			// 
			this.Label4.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(64)));
			this.Label4.Location = new System.Drawing.Point(8, 16);
			this.Label4.Name = "Label4";
			this.Label4.Size = new System.Drawing.Size(256, 40);
			this.Label4.TabIndex = 0;
			this.Label4.Text = "You can use Subclassing to disable the Textbox context menu in which it appears w" +
				"hen the user right click the Text box.";
			// 
			// ContextMenu1
			// 
			this.ContextMenu1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																						 this.MenuItem1});
			// 
			// MenuItem1
			// 
			this.MenuItem1.Index = 0;
			this.MenuItem1.Text = "Click Me";
			this.MenuItem1.Click += new System.EventHandler(this.MenuItem1_Click);
			// 
			// groupBox3
			// 
			this.groupBox3.Controls.Add(this.axSubClass3);
			this.groupBox3.Controls.Add(this.butUnSubF);
			this.groupBox3.Controls.Add(this.butFSubClass);
			this.groupBox3.Controls.Add(this.label7);
			this.groupBox3.ForeColor = System.Drawing.Color.Blue;
			this.groupBox3.Location = new System.Drawing.Point(8, 488);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(272, 96);
			this.groupBox3.TabIndex = 3;
			this.groupBox3.TabStop = false;
			this.groupBox3.Text = "SubClass Sample 3";
			// 
			// butUnSubF
			// 
			this.butUnSubF.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(64)));
			this.butUnSubF.Location = new System.Drawing.Point(144, 64);
			this.butUnSubF.Name = "butUnSubF";
			this.butUnSubF.TabIndex = 11;
			this.butUnSubF.Text = "UnSubClass";
			this.butUnSubF.Click += new System.EventHandler(this.butUnSubF_Click);
			// 
			// butFSubClass
			// 
			this.butFSubClass.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(64)));
			this.butFSubClass.Location = new System.Drawing.Point(56, 64);
			this.butFSubClass.Name = "butFSubClass";
			this.butFSubClass.TabIndex = 10;
			this.butFSubClass.Text = "SubClass";
			this.butFSubClass.Click += new System.EventHandler(this.butFSubClass_Click);
			// 
			// label7
			// 
			this.label7.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(64)));
			this.label7.Location = new System.Drawing.Point(8, 16);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(256, 40);
			this.label7.TabIndex = 0;
			this.label7.Text = "Try to resize this Form before and after you click the subclass Form command butt" +
				"ton and watch the difference.";
			// 
			// axSubClass1
			// 
			this.axSubClass1.ContainingControl = this;
			this.axSubClass1.Enabled = true;
			this.axSubClass1.Location = new System.Drawing.Point(224, 208);
			this.axSubClass1.Name = "axSubClass1";
			this.axSubClass1.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("axSubClass1.OcxState")));
			this.axSubClass1.Size = new System.Drawing.Size(36, 36);
			this.axSubClass1.TabIndex = 6;
			this.axSubClass1.MessageReceived += new AxASTC.__SubClass_MessageReceivedEventHandler(this.axSubClass1_MessageReceived);
			// 
			// axSubClass2
			// 
			this.axSubClass2.ContainingControl = this;
			this.axSubClass2.Enabled = true;
			this.axSubClass2.Location = new System.Drawing.Point(232, 176);
			this.axSubClass2.Name = "axSubClass2";
			this.axSubClass2.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("axSubClass2.OcxState")));
			this.axSubClass2.Size = new System.Drawing.Size(36, 36);
			this.axSubClass2.TabIndex = 12;
			this.axSubClass2.MessageReceived += new AxASTC.__SubClass_MessageReceivedEventHandler(this.axSubClass2_MessageReceived);
			// 
			// axSubClass3
			// 
			this.axSubClass3.ContainingControl = this;
			this.axSubClass3.Enabled = true;
			this.axSubClass3.Location = new System.Drawing.Point(232, 56);
			this.axSubClass3.Name = "axSubClass3";
			this.axSubClass3.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("axSubClass3.OcxState")));
			this.axSubClass3.Size = new System.Drawing.Size(36, 36);
			this.axSubClass3.TabIndex = 12;
			this.axSubClass3.MessageReceived += new AxASTC.__SubClass_MessageReceivedEventHandler(this.axSubClass3_MessageReceived);
			// 
			// Form1
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(288, 590);
			this.Controls.Add(this.groupBox3);
			this.Controls.Add(this.GroupBox2);
			this.Controls.Add(this.GroupBox1);
			this.Name = "Form1";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "VC#.NET TrayList";
			this.GroupBox1.ResumeLayout(false);
			this.GroupBox2.ResumeLayout(false);
			this.groupBox3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.axSubClass1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.axSubClass2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.axSubClass3)).EndInit();
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
			axSubClass1.BeginSubClass(ListBox1.Handle.ToInt32());
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

		  const int WM_RBUTTONUP = 517;
          const int LB_ITEMFROMPOINT = 425;		
	      
			if (e.msg ==  WM_RBUTTONUP)
			{
			  //Retrieve the zero-based index of the item nearest a specified point in the list box.
              ret = ShellNotify.SendMessage(e.hwnd, LB_ITEMFROMPOINT, 0, e.lParam);
				for (int g = 0; g != (int)ListBox1.Items.Count + 1; g++)
				{
					if (ret == g)
					{
						ListBox1.SetSelected(g,true);  
				        Label3.Text = "item" + ListBox1.GetItemText(ListBox1.SelectedIndex) + " had been clicked";
						int X = Cursor.Position.X;
                        int Y = Cursor.Position.Y;
					    //Extends Listbox functionality by Display a popup menu related to the clicked item.
                        ShellNotify.TrackPopupMenu(ContextMenu1.Handle.ToInt32(), 0, X, Y,0,Handle.ToInt32(), 0);
	                 }
                     
				}


			
			}
			 
         }

        /*
		'==============================================================================
		' Method:        ButLUnSubClass_Click
		'
		' Description:  Stop Subclassing List1.
		'==============================================================================
        */
		private void ButLUnSubClass_Click(object sender, System.EventArgs e)
		{
            axSubClass1.EndSubClass();
            Label3.Text = "";
		}

		private void MenuItem1_Click(object sender, System.EventArgs e)
		{
			MessageBox.Show("item" + ListBox1.GetItemText(ListBox1.SelectedIndex)  + " had been clicked");  
		}
           
        /*
		'==============================================================================
		' Method:        ButSubContext_Click
		'
		' Description:  Starts subclassing TextBox1.
		'==============================================================================
		*/
		private void ButSubContext_Click(object sender, System.EventArgs e)
		{
		   axSubClass2.BeginSubClass(TextBox1.Handle.ToInt32());
		}

        /*
		'==============================================================================
		' Method:        axSubClass2_MessageReceived
		'
		' Description:   Disable a context menu in a text box.
		'==============================================================================
		*/
		private void axSubClass2_MessageReceived(object sender, AxASTC.__SubClass_MessageReceivedEvent e)
		{
           /*
		   'e.hwnd()        Handle of the subclassed window or control.
           'e.msg()         The ID of the intercepted message.
           'e.wParam()      The wParam value of the intercepted message.
           'e.lParam()      The lParam value of the intercepted message.
           'e.cancel        Further processing state.
           */

           const int WM_CONTEXTMENU = 123;
              
			if (e.msg == WM_CONTEXTMENU) 
			{
				Label5.Text = "Context Menu had been disabled";
				//Cancel any further processing.
				e.cancel = true;
			}  
		
		}
    
		/*
		'==============================================================================
        ' Method:        ButUnSubCont_Click
        '
        ' Description:  Stop Subclassing TextBox1.
        '==============================================================================
        */
		private void ButUnSubCont_Click(object sender, System.EventArgs e)
		{
		  axSubClass2.EndSubClass();
          Label5.Text = "";
		}
        
		/*
		'==============================================================================
		' Method:        butFSubClass_Click
		'
		' Description:  Start subclassing the Form.
		'==============================================================================
		*/
		private void butFSubClass_Click(object sender, System.EventArgs e)
		{
		   axSubClass3.BeginSubClass(Handle.ToInt32());
		}

        /* 
		'==============================================================================
		' Method:        axSubClass3_MessageReceived
		'
		' Description:  Stop the form from being resized below or above a user-defined amount.
		'==============================================================================
        */
		private void axSubClass3_MessageReceived(object sender, AxASTC.__SubClass_MessageReceivedEvent e)
		{
			/*
			'e.hwnd()        Handle of the subclassed window or control.
			'e.msg()         The ID of the intercepted message.
			'e.wParam()      The wParam value of the intercepted message.
			'e.lParam()      The lParam value of the intercepted message.
			'e.cancel        Further processing state.
			*/
			
			const int WM_GETMINMAXINFO = 36;

            ShellNotify.MINMAXINFO mmiMsg = new ShellNotify.MINMAXINFO();   

              if (e.msg == WM_GETMINMAXINFO)
			  {
               
			   ShellNotify.CopyMemory (ref mmiMsg,e.lParam, Marshal.SizeOf(mmiMsg));
               if (mmiMsg.ptMinTrackSize.X > 0)
				   mmiMsg.ptMinTrackSize.X = 296;
               if (mmiMsg.ptMinTrackSize.Y > 0)
				   mmiMsg.ptMinTrackSize.Y = 624;
               ShellNotify.CopyMemory2 (e.lParam,ref mmiMsg, Marshal.SizeOf(mmiMsg));
              
			   //Cancel any further processing.
			   e.cancel = true;

			   }
		}
        
		/*
		'==============================================================================
		' Method:        butUnSubF_Click
		'
		' Description:  Stop Subclassing Form1.
		'==============================================================================
		*/
		private void butUnSubF_Click(object sender, System.EventArgs e)
		{
			axSubClass3.EndSubClass();
		}

	
	}
}
