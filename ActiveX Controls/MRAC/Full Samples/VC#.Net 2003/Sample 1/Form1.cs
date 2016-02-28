/*
'=========================================================================================
'Publisher      CprinGold Software.
'               http://www.cpringold.com
'               support@cpringold.com
'
'
' Description:  A sample code demonstrates how to use MacroRecorder
'               ActiveX control v1.50 to record and replay mouse clicks, keystrokes
'               and bundle them into a file [macro] in order to replay later.
'==========================================================================================
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
		internal System.Windows.Forms.OpenFileDialog OpenFileDialog1;
		internal System.Windows.Forms.Timer Timer1;
		internal System.Windows.Forms.SaveFileDialog SaveFileDialog1;
		internal System.Windows.Forms.ComboBox ComSpeed;
		internal System.Windows.Forms.Button StopRecordBut;
		internal System.Windows.Forms.Button RecordBut;
		internal System.Windows.Forms.GroupBox GroupBox1;
		internal System.Windows.Forms.Label Label1;
		internal System.Windows.Forms.Label Label2;
		internal System.Windows.Forms.Label Label3;
		internal System.Windows.Forms.Button PlaybackBut;
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
			this.OpenFileDialog1 = new System.Windows.Forms.OpenFileDialog();
			this.Timer1 = new System.Windows.Forms.Timer(this.components);
			this.SaveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.ComSpeed = new System.Windows.Forms.ComboBox();
			this.StopRecordBut = new System.Windows.Forms.Button();
			this.RecordBut = new System.Windows.Forms.Button();
			this.GroupBox1 = new System.Windows.Forms.GroupBox();
			this.PlaybackBut = new System.Windows.Forms.Button();
			this.Label3 = new System.Windows.Forms.Label();
			this.Label2 = new System.Windows.Forms.Label();
			this.Label1 = new System.Windows.Forms.Label();
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
			// ComSpeed
			// 
			this.ComSpeed.ItemHeight = 13;
			this.ComSpeed.Items.AddRange(new object[] {
														  "High Speed",
														  "Normal Speed",
														  "Low Speed"});
			this.ComSpeed.Location = new System.Drawing.Point(104, 128);
			this.ComSpeed.Name = "ComSpeed";
			this.ComSpeed.Size = new System.Drawing.Size(104, 21);
			this.ComSpeed.TabIndex = 4;
			this.ComSpeed.Text = "Replay Speed";
			this.ComSpeed.SelectedIndexChanged += new System.EventHandler(this.ComSpeed_SelectedIndexChanged);
			// 
			// StopRecordBut
			// 
			this.StopRecordBut.Location = new System.Drawing.Point(8, 80);
			this.StopRecordBut.Name = "StopRecordBut";
			this.StopRecordBut.Size = new System.Drawing.Size(88, 24);
			this.StopRecordBut.TabIndex = 1;
			this.StopRecordBut.Text = "&Stop Record";
			this.StopRecordBut.Click += new System.EventHandler(this.StopRecordBut_Click);
			// 
			// RecordBut
			// 
			this.RecordBut.Location = new System.Drawing.Point(8, 32);
			this.RecordBut.Name = "RecordBut";
			this.RecordBut.Size = new System.Drawing.Size(88, 24);
			this.RecordBut.TabIndex = 0;
			this.RecordBut.Text = "&Record";
			this.RecordBut.Click += new System.EventHandler(this.RecordBut_Click);
			// 
			// GroupBox1
			// 
			this.GroupBox1.Controls.Add(this.PlaybackBut);
			this.GroupBox1.Controls.Add(this.Label3);
			this.GroupBox1.Controls.Add(this.Label2);
			this.GroupBox1.Controls.Add(this.Label1);
			this.GroupBox1.Controls.Add(this.ComSpeed);
			this.GroupBox1.Controls.Add(this.StopRecordBut);
			this.GroupBox1.Controls.Add(this.RecordBut);
			this.GroupBox1.Controls.Add(this.axMacroRecorder1);
			this.GroupBox1.Location = new System.Drawing.Point(8, 8);
			this.GroupBox1.Name = "GroupBox1";
			this.GroupBox1.Size = new System.Drawing.Size(408, 160);
			this.GroupBox1.TabIndex = 1;
			this.GroupBox1.TabStop = false;
			// 
			// PlaybackBut
			// 
			this.PlaybackBut.Location = new System.Drawing.Point(8, 128);
			this.PlaybackBut.Name = "PlaybackBut";
			this.PlaybackBut.Size = new System.Drawing.Size(88, 24);
			this.PlaybackBut.TabIndex = 10;
			this.PlaybackBut.Text = "&Replay";
			this.PlaybackBut.Click += new System.EventHandler(this.PlaybackBut_Click);
			// 
			// Label3
			// 
			this.Label3.Location = new System.Drawing.Point(8, 112);
			this.Label3.Name = "Label3";
			this.Label3.Size = new System.Drawing.Size(304, 16);
			this.Label3.TabIndex = 9;
			this.Label3.Text = "Click Replay button to Replay the macro you had saved";
			// 
			// Label2
			// 
			this.Label2.Location = new System.Drawing.Point(8, 64);
			this.Label2.Name = "Label2";
			this.Label2.Size = new System.Drawing.Size(392, 16);
			this.Label2.TabIndex = 8;
			this.Label2.Text = "2. Click Stop Record button to stop recording and save the macro to the disk.";
			// 
			// Label1
			// 
			this.Label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(178)));
			this.Label1.Location = new System.Drawing.Point(8, 16);
			this.Label1.Name = "Label1";
			this.Label1.Size = new System.Drawing.Size(208, 16);
			this.Label1.TabIndex = 7;
			this.Label1.Text = "1. Click Record button to start recording.";
			// 
			// axMacroRecorder1
			// 
			this.axMacroRecorder1.ContainingControl = this;
			this.axMacroRecorder1.Enabled = true;
			this.axMacroRecorder1.Location = new System.Drawing.Point(368, 8);
			this.axMacroRecorder1.Name = "axMacroRecorder1";
			this.axMacroRecorder1.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("axMacroRecorder1.OcxState")));
			this.axMacroRecorder1.Size = new System.Drawing.Size(35, 35);
			this.axMacroRecorder1.TabIndex = 2;
			this.axMacroRecorder1.ReplayFinish += new AxMacroRecorderActX.__MacroRecorder_ReplayFinishEventHandler(this.axMacroRecorder1_ReplayFinish);
			this.axMacroRecorder1.ReplayStart += new AxMacroRecorderActX.__MacroRecorder_ReplayStartEventHandler(this.axMacroRecorder1_ReplayStart);
			this.axMacroRecorder1.RecordStart += new System.EventHandler(this.axMacroRecorder1_RecordStart);
			this.axMacroRecorder1.SaveComplete += new AxMacroRecorderActX.__MacroRecorder_SaveCompleteEventHandler(this.axMacroRecorder1_SaveComplete);
			// 
			// Form1
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(424, 174);
			this.Controls.Add(this.GroupBox1);
			this.Name = "Form1";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "VC# Macro Recorder";
			this.Load += new System.EventHandler(this.Form1_Load);
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

				/*
				You can use any extension for the macro file name, or you can use no extension.
            	hence it will be easy to integrate the MacroRecorder ActiveX Control in your
				Applications by using any file extension you want for the macro file generated by
				MacroRecorder ActiveX Control.   
				*/
				SaveFileDialog1.FileName = "macro1";
				SaveFileDialog1.ShowDialog(); 
				/*  
				You can use any extension for the macro file name, or you can use no extension.
				hence it will be easy to integrate the MacroRecorder ActiveX Control in your
				Applications by using any file extension you want for the macro file generated by
				MacroRecorder ActiveX Control.

				For example suppose that you want to save the macro to c:\
				If you want to assign (*.xxx) as an extension for the generated macro file,
				then all you have to do is
				MacroRecorder1.Save ("c:\mymacro.xxx")
				If you want to use no extension then use the following:
				MacroRecorder1.Save ("c:\mymacro")
                */
				axMacroRecorder1.Save(SaveFileDialog1.FileName);

			}
		}
        
		/*
		'==============================================================================
		' Method:        PlaybackBut_Click
		'
		' Description:  Start playback a macro.
		'==============================================================================
        */ 
		private void PlaybackBut_Click(object sender, System.EventArgs e)
		{
		
			if (! axMacroRecorder1.IsReplay )
			{
				OpenFileDialog1.ShowDialog();
				if (OpenFileDialog1.FileName != "")  
				{   
					/*
					 Pass to MacroPath parameter any file you had saved by Save Method,
					 passing invalid macro file has no effect and no Replaying will happen.
					*/
					axMacroRecorder1.StartReplay (OpenFileDialog1.FileName);
				}
			}
		}

		private void Form1_Load(object sender, System.EventArgs e)
		{
			ComSpeed.SelectedIndex = 1;
		}

	    private void ComSpeed_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			/*
			'To Replay with high speed set Speed parameter to 0
			'To Replay with Normal speed set Speed parameter to 1
			'To Replay with Low speed set Speed parameter to 2
			*/
			switch (ComSpeed.SelectedIndex)
			{
				case (int)MacroRecorderActX.ReplaySpeed.HighSpeed:   
					axMacroRecorder1.SetReplaySpeed(MacroRecorderActX.ReplaySpeed.HighSpeed);
					break;
				case (int)MacroRecorderActX.ReplaySpeed.LowSpeed:   
					axMacroRecorder1.SetReplaySpeed(MacroRecorderActX.ReplaySpeed.LowSpeed);  
					break;
				case (int)MacroRecorderActX.ReplaySpeed.NormalSpeed:   
					axMacroRecorder1.SetReplaySpeed(MacroRecorderActX.ReplaySpeed.NormalSpeed);
					break;

			}	  
	    }

		private void Timer1_Tick(object sender, System.EventArgs e)
		{
          //Stop Recording when you Click Alt+F10
          if (API.GetAsyncKeyState(121) && API.GetAsyncKeyState(18))
				StopRecord();
         
		  //Stop Replaying when you Click Alt+F9
		  if (API.GetAsyncKeyState(120) && API.GetAsyncKeyState(18))
				axMacroRecorder1.StopReplay();
		}

		private void axMacroRecorder1_ReplayFinish(object sender, AxMacroRecorderActX.__MacroRecorder_ReplayFinishEvent e)
		{
			MessageBox.Show ("Repaly " + e.macFilePath + " Macro Finish");
		}

		private void axMacroRecorder1_ReplayStart(object sender, AxMacroRecorderActX.__MacroRecorder_ReplayStartEvent e)
		{
			MessageBox.Show("Press Alt+F9 to terminate Replay.");
		}

		private void axMacroRecorder1_SaveComplete(object sender, AxMacroRecorderActX.__MacroRecorder_SaveCompleteEvent e)
		{
			if (e.macFilePath != "") 
				MessageBox.Show("Macro Saved To " + e.macFilePath);
		}

		private void axMacroRecorder1_RecordStart(object sender, System.EventArgs e)
		{
			MessageBox.Show ("Press Alt+F10 to stop recording.");
		}

		

	}
}
