/*
'====================================================================================
'Publisher      CprinGold Software.
'               http://www.cpringold.com
'               support@cpringold.com
'
'
' Description:  A sample code demonstrates how to use MacroRecorder
'               ActiveX control v1.50 to record and replay mouse clicks, keystrokes
'               and bundle them into an executble file [.exe] in order to replay later.
'=====================================================================================
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
		internal System.Windows.Forms.Timer Timer1;
		internal System.Windows.Forms.SaveFileDialog SaveFileDialog1;
		internal System.Windows.Forms.GroupBox GroupBox1;
		internal System.Windows.Forms.Label Label2;
		internal System.Windows.Forms.Label Label1;
		internal System.Windows.Forms.Button butSaveExe;
		internal System.Windows.Forms.Button RecordBut;
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
			this.Timer1 = new System.Windows.Forms.Timer(this.components);
			this.SaveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.GroupBox1 = new System.Windows.Forms.GroupBox();
			this.Label2 = new System.Windows.Forms.Label();
			this.Label1 = new System.Windows.Forms.Label();
			this.butSaveExe = new System.Windows.Forms.Button();
			this.RecordBut = new System.Windows.Forms.Button();
			this.axMacroRecorder1 = new AxMacroRecorderActX.AxMacroRecorder();
			this.GroupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.axMacroRecorder1)).BeginInit();
			this.SuspendLayout();
			// 
			// Timer1
			// 
			this.Timer1.Enabled = true;
			this.Timer1.Tick += new System.EventHandler(this.Timer1_Tick);
			// 
			// GroupBox1
			// 
			this.GroupBox1.Controls.Add(this.Label2);
			this.GroupBox1.Controls.Add(this.Label1);
			this.GroupBox1.Controls.Add(this.butSaveExe);
			this.GroupBox1.Controls.Add(this.RecordBut);
			this.GroupBox1.Controls.Add(this.axMacroRecorder1);
			this.GroupBox1.Location = new System.Drawing.Point(8, 8);
			this.GroupBox1.Name = "GroupBox1";
			this.GroupBox1.Size = new System.Drawing.Size(464, 128);
			this.GroupBox1.TabIndex = 1;
			this.GroupBox1.TabStop = false;
			// 
			// Label2
			// 
			this.Label2.Location = new System.Drawing.Point(8, 72);
			this.Label2.Name = "Label2";
			this.Label2.Size = new System.Drawing.Size(448, 24);
			this.Label2.TabIndex = 15;
			this.Label2.Text = "2.Click Save as Exe button to stop recording and save the macro as an exectutable" +
				" file.";
			// 
			// Label1
			// 
			this.Label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(178)));
			this.Label1.Location = new System.Drawing.Point(8, 16);
			this.Label1.Name = "Label1";
			this.Label1.Size = new System.Drawing.Size(208, 16);
			this.Label1.TabIndex = 14;
			this.Label1.Text = "1.Click Record button to start recording.";
			// 
			// butSaveExe
			// 
			this.butSaveExe.Location = new System.Drawing.Point(8, 96);
			this.butSaveExe.Name = "butSaveExe";
			this.butSaveExe.Size = new System.Drawing.Size(88, 24);
			this.butSaveExe.TabIndex = 13;
			this.butSaveExe.Text = "&Save As Exe";
			this.butSaveExe.Click += new System.EventHandler(this.StopRecordBut_Click);
			// 
			// RecordBut
			// 
			this.RecordBut.Location = new System.Drawing.Point(8, 40);
			this.RecordBut.Name = "RecordBut";
			this.RecordBut.Size = new System.Drawing.Size(88, 24);
			this.RecordBut.TabIndex = 12;
			this.RecordBut.Text = "&Record";
			this.RecordBut.Click += new System.EventHandler(this.RecordBut_Click);
			// 
			// axMacroRecorder1
			// 
			this.axMacroRecorder1.ContainingControl = this;
			this.axMacroRecorder1.Enabled = true;
			this.axMacroRecorder1.Location = new System.Drawing.Point(424, 8);
			this.axMacroRecorder1.Name = "axMacroRecorder1";
			this.axMacroRecorder1.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("axMacroRecorder1.OcxState")));
			this.axMacroRecorder1.Size = new System.Drawing.Size(35, 35);
			this.axMacroRecorder1.TabIndex = 2;
			this.axMacroRecorder1.RecordStart += new System.EventHandler(this.axMacroRecorder1_RecordStart);
			this.axMacroRecorder1.SaveComplete += new AxMacroRecorderActX.__MacroRecorder_SaveCompleteEventHandler(this.axMacroRecorder1_SaveComplete);
			// 
			// Form1
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(480, 142);
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

				/*  
				'Note that
				'Macro Recorder ActiveX Control will add the ".exe" extension
				'to the file path automatically
				'For example suppose that you want to save the macro to c:\
				'you can use one of the following
				'MacroRecorder1.SaveAsExE ("c:\macro.exe")
				'MacroRecorder1.SaveAsExE ("c:\macro")     'Macro Recorder will add the .exe extension to the file
				*/
				axMacroRecorder1.SaveAsExE (SaveFileDialog1.FileName);
				/*
				'You can use command line parameters for the executable macro file to change the replay speed
				'For example if you save an executable macro to macro1.exe
				'macro1 /h    for high replay speed.
				'macro1 /n    for normal replay speed "the default one"
				'macro1 /l    for low replay speed.
				*/

			}
		}
        
		private void Timer1_Tick(object sender, System.EventArgs e)
		{
          //Stop Recording when you Click Alt+F10
          if (API.GetAsyncKeyState(121) && API.GetAsyncKeyState(18))
				StopRecord();
		}

		private void axMacroRecorder1_SaveComplete(object sender, AxMacroRecorderActX.__MacroRecorder_SaveCompleteEvent e)
		{
			if (e.macFilePath != "") 
				MessageBox.Show("Macro Saved To " + e.macFilePath);
		}

		private void axMacroRecorder1_RecordStart(object sender, System.EventArgs e)
		{
			MessageBox.Show ("Press Alt+F10 to Stop Recording.");
		}

	}
}
