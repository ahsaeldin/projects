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
	/// Summary description for WinForm.
	/// </summary>
	public class WinForm : System.Windows.Forms.Form
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.GroupBox groupBox2;
		private AxASTC.AxSubClass axSubClass2;
		private System.Windows.Forms.Button ButUnSubCont;
		private System.Windows.Forms.Button ButSubContext;
		private System.Windows.Forms.Label Label5;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.TextBox TextBox1;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.ListBox ListBox1;
		private System.Windows.Forms.Label Label3;
		private System.Windows.Forms.Button ButLSubClass;
		private System.Windows.Forms.Button ButLUnSubClass;
		private AxASTC.AxSubClass axSubClass1;
		private System.Windows.Forms.ContextMenu ContextMenu1;
		private System.Windows.Forms.MenuItem menuItem1;
		private System.Windows.Forms.GroupBox groupBox3;
		private AxASTC.AxSubClass axSubClass3;
		private System.Windows.Forms.Button butUnSubF;
		private System.Windows.Forms.Button butFSubClass;
		private System.Windows.Forms.Label label7;

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
			this.axSubClass1 = new AxASTC.AxSubClass();
			this.ButLUnSubClass = new System.Windows.Forms.Button();
			this.ButLSubClass = new System.Windows.Forms.Button();
			this.Label3 = new System.Windows.Forms.Label();
			this.ListBox1 = new System.Windows.Forms.ListBox();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.ContextMenu1 = new System.Windows.Forms.ContextMenu();
			this.menuItem1 = new System.Windows.Forms.MenuItem();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.TextBox1 = new System.Windows.Forms.TextBox();
			this.axSubClass2 = new AxASTC.AxSubClass();
			this.ButUnSubCont = new System.Windows.Forms.Button();
			this.ButSubContext = new System.Windows.Forms.Button();
			this.Label5 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.axSubClass3 = new AxASTC.AxSubClass();
			this.butUnSubF = new System.Windows.Forms.Button();
			this.butFSubClass = new System.Windows.Forms.Button();
			this.label7 = new System.Windows.Forms.Label();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.axSubClass1)).BeginInit();
			this.groupBox2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.axSubClass2)).BeginInit();
			this.groupBox3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.axSubClass3)).BeginInit();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.axSubClass1);
			this.groupBox1.Controls.Add(this.ButLUnSubClass);
			this.groupBox1.Controls.Add(this.ButLSubClass);
			this.groupBox1.Controls.Add(this.Label3);
			this.groupBox1.Controls.Add(this.ListBox1);
			this.groupBox1.Controls.Add(this.label2);
			this.groupBox1.Controls.Add(this.label1);
			this.groupBox1.Location = new System.Drawing.Point(8, 8);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(280, 240);
			this.groupBox1.TabIndex = 0;
			this.groupBox1.TabStop = false;
			// 
			// axSubClass1
			// 
			this.axSubClass1.ContainingControl = this;
			this.axSubClass1.Enabled = true;
			this.axSubClass1.Location = new System.Drawing.Point(240, 192);
			this.axSubClass1.Name = "axSubClass1";
			this.axSubClass1.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("axSubClass1.OcxState")));
			this.axSubClass1.Size = new System.Drawing.Size(36, 36);
			this.axSubClass1.TabIndex = 6;
			this.axSubClass1.MessageReceived += new AxASTC.__SubClass_MessageReceivedEventHandler(this.axSubClass1_MessageReceived);
			// 
			// ButLUnSubClass
			// 
			this.ButLUnSubClass.Location = new System.Drawing.Point(144, 208);
			this.ButLUnSubClass.Name = "ButLUnSubClass";
			this.ButLUnSubClass.Size = new System.Drawing.Size(80, 24);
			this.ButLUnSubClass.TabIndex = 5;
			this.ButLUnSubClass.Text = "UnSubClass";
			this.ButLUnSubClass.Click += new System.EventHandler(this.ButLUnSubClass_Click);
			// 
			// ButLSubClass
			// 
			this.ButLSubClass.Location = new System.Drawing.Point(56, 208);
			this.ButLSubClass.Name = "ButLSubClass";
			this.ButLSubClass.Size = new System.Drawing.Size(80, 24);
			this.ButLSubClass.TabIndex = 4;
			this.ButLSubClass.Text = "SubClass";
			this.ButLSubClass.Click += new System.EventHandler(this.ButLSubClass_Click);
			// 
			// Label3
			// 
			this.Label3.ForeColor = System.Drawing.Color.Red;
			this.Label3.Location = new System.Drawing.Point(8, 176);
			this.Label3.Name = "Label3";
			this.Label3.Size = new System.Drawing.Size(264, 24);
			this.Label3.TabIndex = 3;
			this.Label3.Text = "Then click SubClass command and Right Click any item and you will " +  
				"see the difference.";
			// 
			// ListBox1
			// 
			this.ListBox1.Items.AddRange(new object[] {
						"item0",
						"item1",
						"item2",
						"item3"});
			this.ListBox1.Location = new System.Drawing.Point(8, 104);
			this.ListBox1.Name = "ListBox1";
			this.ListBox1.Size = new System.Drawing.Size(264, 69);
			this.ListBox1.TabIndex = 2;
			// 
			// label2
			// 
			this.label2.ForeColor = System.Drawing.Color.Red;
			this.label2.Location = new System.Drawing.Point(8, 72);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(264, 32);
			this.label2.TabIndex = 1;
			this.label2.Text = "Right Click any item in listbox before you click the Subclass comm" +  
				"and button.";
			// 
			// label1
			// 
			this.label1.ForeColor = System.Drawing.Color.Blue;
			this.label1.Location = new System.Drawing.Point(8, 16);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(264, 56);
			this.label1.TabIndex = 0;
			this.label1.Text = "In the Normal ListBox you can\'t assign a popup menu to any item so" +  
				" when the user right click the item it will appears              " +  
				"                                      But if you we use Subclassi" +  
				"ng we can achieve this.";
			// 
			// ContextMenu1
			// 
			this.ContextMenu1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
						this.menuItem1});
			// 
			// menuItem1
			// 
			this.menuItem1.Index = 0;
			this.menuItem1.Text = "Click Me";
			this.menuItem1.Click += new System.EventHandler(this.menuItem1_Click);
			// 
			// groupBox2
			// 
			this.groupBox2.Controls.Add(this.TextBox1);
			this.groupBox2.Controls.Add(this.axSubClass2);
			this.groupBox2.Controls.Add(this.ButUnSubCont);
			this.groupBox2.Controls.Add(this.ButSubContext);
			this.groupBox2.Controls.Add(this.Label5);
			this.groupBox2.Controls.Add(this.label6);
			this.groupBox2.Location = new System.Drawing.Point(8, 256);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(280, 184);
			this.groupBox2.TabIndex = 1;
			this.groupBox2.TabStop = false;
			// 
			// TextBox1
			// 
			this.TextBox1.Location = new System.Drawing.Point(8, 56);
			this.TextBox1.Multiline = true;
			this.TextBox1.Name = "TextBox1";
			this.TextBox1.Size = new System.Drawing.Size(248, 40);
			this.TextBox1.TabIndex = 7;
			this.TextBox1.Text = "Right Click me before and after you click Subclass command  button" +  
				" and watch the difference.";
			// 
			// axSubClass2
			// 
			this.axSubClass2.ContainingControl = this;
			this.axSubClass2.Enabled = true;
			this.axSubClass2.Location = new System.Drawing.Point(240, 136);
			this.axSubClass2.Name = "axSubClass2";
			this.axSubClass2.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("axSubClass2.OcxState")));
			this.axSubClass2.Size = new System.Drawing.Size(36, 36);
			this.axSubClass2.TabIndex = 6;
			this.axSubClass2.Visible = false;
			this.axSubClass2.MessageReceived += new AxASTC.__SubClass_MessageReceivedEventHandler(this.axSubClass2_MessageReceived);
			// 
			// ButUnSubCont
			// 
			this.ButUnSubCont.Location = new System.Drawing.Point(144, 152);
			this.ButUnSubCont.Name = "ButUnSubCont";
			this.ButUnSubCont.Size = new System.Drawing.Size(80, 24);
			this.ButUnSubCont.TabIndex = 5;
			this.ButUnSubCont.Text = "UnSubClass";
			this.ButUnSubCont.Click += new System.EventHandler(this.ButUnSubCont_Click);
			// 
			// ButSubContext
			// 
			this.ButSubContext.Location = new System.Drawing.Point(56, 152);
			this.ButSubContext.Name = "ButSubContext";
			this.ButSubContext.Size = new System.Drawing.Size(80, 24);
			this.ButSubContext.TabIndex = 4;
			this.ButSubContext.Text = "SubClass";
			this.ButSubContext.Click += new System.EventHandler(this.ButSubContext_Click);
			// 
			// Label5
			// 
			this.Label5.BackColor = System.Drawing.Color.White;
			this.Label5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.Label5.ForeColor = System.Drawing.Color.Red;
			this.Label5.Location = new System.Drawing.Point(8, 104);
			this.Label5.Name = "Label5";
			this.Label5.Size = new System.Drawing.Size(248, 32);
			this.Label5.TabIndex = 3;
			// 
			// label6
			// 
			this.label6.ForeColor = System.Drawing.Color.Blue;
			this.label6.Location = new System.Drawing.Point(8, 16);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(264, 40);
			this.label6.TabIndex = 0;
			this.label6.Text = "You can use Subclassing to disable the Textbox context menu in whi" +  
				"ch it appears when the user right click the Text box.";
			// 
			// groupBox3
			// 
			this.groupBox3.Controls.Add(this.axSubClass3);
			this.groupBox3.Controls.Add(this.butUnSubF);
			this.groupBox3.Controls.Add(this.butFSubClass);
			this.groupBox3.Controls.Add(this.label7);
			this.groupBox3.Location = new System.Drawing.Point(8, 448);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(280, 88);
			this.groupBox3.TabIndex = 2;
			this.groupBox3.TabStop = false;
			// 
			// axSubClass3
			// 
			this.axSubClass3.ContainingControl = this;
			this.axSubClass3.Enabled = true;
			this.axSubClass3.Location = new System.Drawing.Point(240, 40);
			this.axSubClass3.Name = "axSubClass3";
			this.axSubClass3.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("axSubClass3.OcxState")));
			this.axSubClass3.Size = new System.Drawing.Size(36, 36);
			this.axSubClass3.TabIndex = 6;
			this.axSubClass3.Visible = false;
			this.axSubClass3.MessageReceived += new AxASTC.__SubClass_MessageReceivedEventHandler(this.axSubClass3_MessageReceived);
			// 
			// butUnSubF
			// 
			this.butUnSubF.Location = new System.Drawing.Point(144, 56);
			this.butUnSubF.Name = "butUnSubF";
			this.butUnSubF.Size = new System.Drawing.Size(80, 24);
			this.butUnSubF.TabIndex = 5;
			this.butUnSubF.Text = "UnSubClass";
			this.butUnSubF.Click += new System.EventHandler(this.butUnSubF_Click);
			// 
			// butFSubClass
			// 
			this.butFSubClass.Location = new System.Drawing.Point(56, 56);
			this.butFSubClass.Name = "butFSubClass";
			this.butFSubClass.Size = new System.Drawing.Size(80, 24);
			this.butFSubClass.TabIndex = 4;
			this.butFSubClass.Text = "SubClass";
			this.butFSubClass.Click += new System.EventHandler(this.butFSubClass_Click);
			// 
			// label7
			// 
			this.label7.ForeColor = System.Drawing.Color.Blue;
			this.label7.Location = new System.Drawing.Point(8, 16);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(264, 40);
			this.label7.TabIndex = 0;
			this.label7.Text = "Try to resize this Form before and after you click the subclass Fo" +  
				"rm command buttton and watch the difference.";
			// 
			// WinForm
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(296, 542);
			this.Controls.Add(this.groupBox3);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.groupBox1);
			this.Name = "WinForm";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "C#Builder SubClass Sample";
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.axSubClass1)).EndInit();
			this.groupBox2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.axSubClass2)).EndInit();
			this.groupBox3.ResumeLayout(false);
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
			Application.Run(new WinForm());
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
		
		private void menuItem1_Click(object sender, System.EventArgs e)
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
				   mmiMsg.ptMinTrackSize.X = 304;
               if (mmiMsg.ptMinTrackSize.Y > 0)
				   mmiMsg.ptMinTrackSize.Y = 576;
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
