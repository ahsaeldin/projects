using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using MouseKeyboardLibrary;

namespace SampleApplication
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class HookTestForm : System.Windows.Forms.Form
	{
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Label curXYLabel;
		private System.Windows.Forms.ListView listView1;
		private System.Windows.Forms.ListView listView2;
		private System.Windows.Forms.ColumnHeader columnHeader1;
		private System.Windows.Forms.ColumnHeader columnHeader2;
		private System.Windows.Forms.ColumnHeader columnHeader3;
		private System.Windows.Forms.ColumnHeader columnHeader4;
		private System.Windows.Forms.ColumnHeader columnHeader5;
		private System.Windows.Forms.ColumnHeader columnHeader6;
		private System.Windows.Forms.ColumnHeader columnHeader7;
		private System.Windows.Forms.ColumnHeader columnHeader8;
		private System.Windows.Forms.ColumnHeader columnHeader9;
		private System.Windows.Forms.ColumnHeader columnHeader10;
		private System.Windows.Forms.ColumnHeader columnHeader11;

		MouseHook mouseHook = new MouseHook();
		KeyboardHook keyboardHook = new KeyboardHook();

		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public HookTestForm()
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
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.curXYLabel = new System.Windows.Forms.Label();
			this.listView1 = new System.Windows.Forms.ListView();
			this.listView2 = new System.Windows.Forms.ListView();
			this.columnHeader1 = new System.Windows.Forms.ColumnHeader();
			this.columnHeader2 = new System.Windows.Forms.ColumnHeader();
			this.columnHeader3 = new System.Windows.Forms.ColumnHeader();
			this.columnHeader4 = new System.Windows.Forms.ColumnHeader();
			this.columnHeader5 = new System.Windows.Forms.ColumnHeader();
			this.columnHeader6 = new System.Windows.Forms.ColumnHeader();
			this.columnHeader7 = new System.Windows.Forms.ColumnHeader();
			this.columnHeader8 = new System.Windows.Forms.ColumnHeader();
			this.columnHeader9 = new System.Windows.Forms.ColumnHeader();
			this.columnHeader10 = new System.Windows.Forms.ColumnHeader();
			this.columnHeader11 = new System.Windows.Forms.ColumnHeader();
			this.groupBox1.SuspendLayout();
			this.groupBox2.SuspendLayout();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.listView1);
			this.groupBox1.Controls.Add(this.curXYLabel);
			this.groupBox1.Location = new System.Drawing.Point(12, 12);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(391, 150);
			this.groupBox1.TabIndex = 0;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Mouse";
			// 
			// groupBox2
			// 
			this.groupBox2.Controls.Add(this.listView2);
			this.groupBox2.Location = new System.Drawing.Point(12, 168);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(391, 150);
			this.groupBox2.TabIndex = 0;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Keyboard";
			// 
			// curXYLabel
			// 
			this.curXYLabel.Location = new System.Drawing.Point(22, 20);
			this.curXYLabel.Name = "curXYLabel";
			this.curXYLabel.Size = new System.Drawing.Size(330, 16);
			this.curXYLabel.TabIndex = 0;
			// 
			// listView1
			// 
			this.listView1.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
																						this.columnHeader1,
																						this.columnHeader2,
																						this.columnHeader3,
																						this.columnHeader4,
																						this.columnHeader5});
			this.listView1.Location = new System.Drawing.Point(6, 38);
			this.listView1.Name = "listView1";
			this.listView1.Size = new System.Drawing.Size(379, 106);
			this.listView1.TabIndex = 1;
			this.listView1.View = System.Windows.Forms.View.Details;
			// 
			// listView2
			// 
			this.listView2.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
																						this.columnHeader6,
																						this.columnHeader7,
																						this.columnHeader8,
																						this.columnHeader9,
																						this.columnHeader10,
																						this.columnHeader11});
			this.listView2.Location = new System.Drawing.Point(6, 19);
			this.listView2.Name = "listView2";
			this.listView2.Size = new System.Drawing.Size(379, 125);
			this.listView2.TabIndex = 1;
			this.listView2.View = System.Windows.Forms.View.Details;
			// 
			// columnHeader1
			// 
			this.columnHeader1.Text = "EventType";
			this.columnHeader1.Width = 80;
			// 
			// columnHeader2
			// 
			this.columnHeader2.Text = "Button";
			this.columnHeader2.Width = 70;
			// 
			// columnHeader3
			// 
			this.columnHeader3.Text = "X";
			this.columnHeader3.Width = 40;
			// 
			// columnHeader4
			// 
			this.columnHeader4.Text = "Y";
			this.columnHeader4.Width = 40;
			// 
			// columnHeader5
			// 
			this.columnHeader5.Text = "Delta";
			// 
			// columnHeader6
			// 
			this.columnHeader6.Text = "EventType";
			this.columnHeader6.Width = 80;
			// 
			// columnHeader7
			// 
			this.columnHeader7.Text = "KeyCode";
			// 
			// columnHeader8
			// 
			this.columnHeader8.Text = "KeyChar";
			// 
			// columnHeader9
			// 
			this.columnHeader9.Text = "Shift";
			this.columnHeader9.Width = 50;
			// 
			// columnHeader10
			// 
			this.columnHeader10.Text = "Alt";
			this.columnHeader10.Width = 50;
			// 
			// columnHeader11
			// 
			this.columnHeader11.Text = "Control";
			this.columnHeader11.Width = 50;
			// 
			// HookTestForm
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(415, 330);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.groupBox2);
			this.Name = "HookTestForm";
			this.Text = "Form1";
			this.Load += new System.EventHandler(this.HookTestForm_Load);
			this.Closed += new System.EventHandler(this.HookTestForm_Closed);
			this.groupBox1.ResumeLayout(false);
			this.groupBox2.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new HookTestForm());
		}

		private void HookTestForm_Load(object sender, System.EventArgs e)
		{
		
			mouseHook.MouseMove += new MouseEventHandler(mouseHook_MouseMove);
			mouseHook.MouseDown += new MouseEventHandler(mouseHook_MouseDown);
			mouseHook.MouseUp += new MouseEventHandler(mouseHook_MouseUp);
			mouseHook.MouseWheel += new MouseEventHandler(mouseHook_MouseWheel);

			keyboardHook.KeyDown += new KeyEventHandler(keyboardHook_KeyDown);
			keyboardHook.KeyUp += new KeyEventHandler(keyboardHook_KeyUp);
			keyboardHook.KeyPress += new KeyPressEventHandler(keyboardHook_KeyPress);

			mouseHook.Start();
			keyboardHook.Start();

			SetXYLabel(MouseSimulator.X, MouseSimulator.Y);

		}


		void keyboardHook_KeyPress(object sender, KeyPressEventArgs e)
		{

			AddKeyboardEvent(
				"KeyPress",
				"",
				e.KeyChar.ToString(),
				"",
				"",
				""
				);

		}

		void keyboardHook_KeyUp(object sender, KeyEventArgs e)
		{

			AddKeyboardEvent(
				"KeyUp",
				e.KeyCode.ToString(),
				"",
				e.Shift.ToString(),
				e.Alt.ToString(),
				e.Control.ToString()
				);

		}

		void keyboardHook_KeyDown(object sender, KeyEventArgs e)
		{


			AddKeyboardEvent(
				"KeyDown",
				e.KeyCode.ToString(),
				"",
				e.Shift.ToString(),
				e.Alt.ToString(),
				e.Control.ToString()
				);

		}

		void mouseHook_MouseWheel(object sender, MouseEventArgs e)
		{

			AddMouseEvent(
				"MouseWheel",
				"",
				"",
				"",
				e.Delta.ToString()
				);

		}

		void mouseHook_MouseUp(object sender, MouseEventArgs e)
		{


			AddMouseEvent(
				"MouseUp",
				e.Button.ToString(),
				e.X.ToString(),
				e.Y.ToString(),
				""
				);

		}

		void mouseHook_MouseDown(object sender, MouseEventArgs e)
		{


			AddMouseEvent(
				"MouseDown",
				e.Button.ToString(),
				e.X.ToString(),
				e.Y.ToString(),
				""
				);


		}

		void mouseHook_MouseMove(object sender, MouseEventArgs e)
		{

			SetXYLabel(e.X, e.Y);

		}

		void SetXYLabel(int x, int y)
		{

			curXYLabel.Text = String.Format("Current Mouse Point: X={0}, y={1}", x, y);

		}

		void AddMouseEvent(string eventType, string button, string x, string y, string delta)
		{

			listView1.Items.Insert(0,
				new ListViewItem(
				new string[]{
								eventType, 
								button,
								x,
								y,
								delta
							}));

		}

		void AddKeyboardEvent(string eventType, string keyCode, string keyChar, string shift, string alt, string control)
		{

			listView2.Items.Insert(0,
				new ListViewItem(
				new string[]{
								eventType, 
								keyCode,
								keyChar,
								shift,
								alt,
								control
							}));

		}

		private void HookTestForm_Closed(object sender, System.EventArgs e)
		{
			
			mouseHook.Stop();
			keyboardHook.Stop();
		}

	}
}
