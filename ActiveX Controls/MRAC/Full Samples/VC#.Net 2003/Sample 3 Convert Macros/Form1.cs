/*
'=============================================================================
'Publisher      CprinGold Software.
'               http://www.cpringold.com
'               support@cpringold.com
'
'
' Description:  A sample code demonstrates how to use MacroRecorder
'               ActiveX control v1.50 to convert normal macro file to
'               an executable macro file.
'=============================================================================
*/
using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Runtime.InteropServices;

namespace WindowsApplication1
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class Form1 : System.Windows.Forms.Form
	{
		internal System.Windows.Forms.SaveFileDialog SaveFileDialog1;
		internal System.Windows.Forms.OpenFileDialog OpenFileDialog1;
		internal System.Windows.Forms.GroupBox GroupBox1;
		internal System.Windows.Forms.Label Label3;
		internal System.Windows.Forms.Label Label2;
		internal System.Windows.Forms.Label Label1;
		internal System.Windows.Forms.Button ButConvert;
		internal System.Windows.Forms.Button StopRecordBut;
		internal System.Windows.Forms.Button RecordBut;
		internal System.Windows.Forms.Timer Timer1;
		private AxMacroRecorderActX.AxMacroRecorder axMacroRecorder1;
		private System.ComponentModel.IContainer components;

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
			this.components = new System.ComponentModel.Container();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(Form1));
			this.SaveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.OpenFileDialog1 = new System.Windows.Forms.OpenFileDialog();
			this.GroupBox1 = new System.Windows.Forms.GroupBox();
			this.Label3 = new System.Windows.Forms.Label();
			this.Label2 = new System.Windows.Forms.Label();
			this.Label1 = new System.Windows.Forms.Label();
			this.ButConvert = new System.Windows.Forms.Button();
			this.StopRecordBut = new System.Windows.Forms.Button();
			this.RecordBut = new System.Windows.Forms.Button();
			this.Timer1 = new System.Windows.Forms.Timer(this.components);
			this.axMacroRecorder1 = new AxMacroRecorderActX.AxMacroRecorder();
			this.GroupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.axMacroRecorder1)).BeginInit();
			this.SuspendLayout();
			// 
			// GroupBox1
			// 
			this.GroupBox1.Controls.Add(this.axMacroRecorder1);
			this.GroupBox1.Controls.Add(this.Label3);
			this.GroupBox1.Controls.Add(this.Label2);
			this.GroupBox1.Controls.Add(this.Label1);
			this.GroupBox1.Controls.Add(this.ButConvert);
			this.GroupBox1.Controls.Add(this.StopRecordBut);
			this.GroupBox1.Controls.Add(this.RecordBut);
			this.GroupBox1.Location = new System.Drawing.Point(8, 8);
			this.GroupBox1.Name = "GroupBox1";
			this.GroupBox1.Size = new System.Drawing.Size(432, 184);
			this.GroupBox1.TabIndex = 13;
			this.GroupBox1.TabStop = false;
			// 
			// Label3
			// 
			this.Label3.Location = new System.Drawing.Point(8, 128);
			this.Label3.Name = "Label3";
			this.Label3.Size = new System.Drawing.Size(416, 24);
			this.Label3.TabIndex = 5;
			this.Label3.Text = "3. Click Convert button to convert a normal macro file to an executable macro fil" +
				"e.";
			// 
			// Label2
			// 
			this.Label2.Location = new System.Drawing.Point(8, 72);
			this.Label2.Name = "Label2";
			this.Label2.Size = new System.Drawing.Size(392, 24);
			this.Label2.TabIndex = 4;
			this.Label2.Text = "2. Click Stop Record button to stop recording and save the macro to the disk.";
			// 
			// Label1
			// 
			this.Label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(178)));
			this.Label1.Location = new System.Drawing.Point(8, 16);
			this.Label1.Name = "Label1";
			this.Label1.Size = new System.Drawing.Size(288, 24);
			this.Label1.TabIndex = 3;
			this.Label1.Text = "1. Click Record button to start recording.";
			// 
			// ButConvert
			// 
			this.ButConvert.Location = new System.Drawing.Point(8, 152);
			this.ButConvert.Name = "ButConvert";
			this.ButConvert.Size = new System.Drawing.Size(88, 24);
			this.ButConvert.TabIndex = 2;
			this.ButConvert.Text = "&Convert";
			this.ButConvert.Click += new System.EventHandler(this.ButConvert_Click);
			// 
			// StopRecordBut
			// 
			this.StopRecordBut.Location = new System.Drawing.Point(8, 96);
			this.StopRecordBut.Name = "StopRecordBut";
			this.StopRecordBut.Size = new System.Drawing.Size(88, 24);
			this.StopRecordBut.TabIndex = 1;
			this.StopRecordBut.Text = "&Stop Record";
			this.StopRecordBut.Click += new System.EventHandler(this.StopRecordBut_Click);
			// 
			// RecordBut
			// 
			this.RecordBut.Location = new System.Drawing.Point(8, 40);
			this.RecordBut.Name = "RecordBut";
			this.RecordBut.Size = new System.Drawing.Size(88, 24);
			this.RecordBut.TabIndex = 0;
			this.RecordBut.Text = "&Record";
			this.RecordBut.Click += new System.EventHandler(this.RecordBut_Click);
			// 
			// Timer1
			// 
			this.Timer1.Enabled = true;
			this.Timer1.Tick += new System.EventHandler(this.Timer1_Tick);
			// 
			// axMacroRecorder1
			// 
			this.axMacroRecorder1.ContainingControl = this;
			this.axMacroRecorder1.Enabled = true;
			this.axMacroRecorder1.Location = new System.Drawing.Point(392, 16);
			this.axMacroRecorder1.Name = "axMacroRecorder1";
			this.axMacroRecorder1.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("axMacroRecorder1.OcxState")));
			this.axMacroRecorder1.Size = new System.Drawing.Size(35, 35);
			this.axMacroRecorder1.TabIndex = 6;
			this.axMacroRecorder1.RecordStart += new System.EventHandler(this.axMacroRecorder1_RecordStart);
			this.axMacroRecorder1.ConvertComplete += new AxMacroRecorderActX.__MacroRecorder_ConvertCompleteEventHandler(this.axMacroRecorder1_ConvertComplete);
			// 
			// Form1
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(448, 198);
			this.Controls.Add(this.GroupBox1);
			this.Name = "Form1";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "VC# Macro Recorder";
			this.GroupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.axMacroRecorder1)).EndInit();
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
		' Method:        RecordBut_Click
		'
		' Description:  Start recording a new macro.
		'==============================================================================
        */
		private void RecordBut_Click(object sender, System.EventArgs e)
		{   
            if (this.WindowState != FormWindowState.Minimized )
				this.WindowState = FormWindowState.Minimized;
			axMacroRecorder1.Record();  
		}

        /*
		'==============================================================================
		' Method:        StopRecordBut_Click
		'
		' Description:  Stop recording and save the macro to the disk.
		'============================================================================== 
        */
		private void StopRecordBut_Click(object sender, System.EventArgs e)
		{
	       StopRecord();
		}
        
		private void StopRecord()
		{
			//Only if the State is recording.
            if (axMacroRecorder1.IsRecord )
			{   
				axMacroRecorder1.StopRecord();
				this.WindowState = FormWindowState.Normal;
				SaveFileDialog1.FileName = "macro1";
				SaveFileDialog1.ShowDialog(); 
				axMacroRecorder1.Save (SaveFileDialog1.FileName);
			}
		}	
		
		private void ButConvert_Click(object sender, System.EventArgs e)
		{			
			OpenFileDialog1.Title = "Select Macro to Convert to exe file.";
			OpenFileDialog1.ShowDialog();
			axMacroRecorder1.ConvertToExE(OpenFileDialog1.FileName, @"c:\macro1.exe");
		}

		private void Timer1_Tick(object sender, System.EventArgs e)
		{
          //Stop Recording when you Click Alt+F10
          if (API.GetAsyncKeyState(121) && API.GetAsyncKeyState(18))
				StopRecord();
		}

		private void axMacroRecorder1_RecordStart(object sender, System.EventArgs e)
		{
			MessageBox.Show ("Press Alt+F10 to Stop Recording.");
		}

		private void axMacroRecorder1_ConvertComplete(object sender, AxMacroRecorderActX.__MacroRecorder_ConvertCompleteEvent e)
		{
			if (e.macFilePath != "") 
				MessageBox.Show("Macro converted and Saved To " + e.macFilePath);
		}

	


	}
}
