using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Threading;

using MouseKeyboardLibrary;

namespace GlobalMacroRecorder
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class MacroForm : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button recordStartButton;
		private System.Windows.Forms.Button recordStopButton;
		private System.Windows.Forms.Button playBackMacroButton;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;

		ArrayList events = new ArrayList();
		int lastTimeRecorded = 0;

		MouseHook mouseHook = new MouseHook();
		KeyboardHook keyboardHook = new KeyboardHook();

		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public MacroForm()
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
            this.recordStartButton = new System.Windows.Forms.Button();
            this.recordStopButton = new System.Windows.Forms.Button();
            this.playBackMacroButton = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // recordStartButton
            // 
            this.recordStartButton.Location = new System.Drawing.Point(188, 12);
            this.recordStartButton.Name = "recordStartButton";
            this.recordStartButton.Size = new System.Drawing.Size(75, 23);
            this.recordStartButton.TabIndex = 0;
            this.recordStartButton.Text = "Start";
            this.recordStartButton.Click += new System.EventHandler(this.recordStartButton_Click);
            // 
            // recordStopButton
            // 
            this.recordStopButton.Location = new System.Drawing.Point(268, 12);
            this.recordStopButton.Name = "recordStopButton";
            this.recordStopButton.Size = new System.Drawing.Size(72, 23);
            this.recordStopButton.TabIndex = 1;
            this.recordStopButton.Text = "Stop";
            this.recordStopButton.Click += new System.EventHandler(this.recordStopButton_Click);
            // 
            // playBackMacroButton
            // 
            this.playBackMacroButton.Location = new System.Drawing.Point(188, 90);
            this.playBackMacroButton.Name = "playBackMacroButton";
            this.playBackMacroButton.Size = new System.Drawing.Size(152, 23);
            this.playBackMacroButton.TabIndex = 2;
            this.playBackMacroButton.Text = "Play Back";
            this.playBackMacroButton.Click += new System.EventHandler(this.playBackMacroButton_Click);
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(32, 12);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(100, 23);
            this.label1.TabIndex = 3;
            this.label1.Text = "Record Macro";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(32, 64);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(100, 23);
            this.label2.TabIndex = 4;
            this.label2.Text = "Playback Macro";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // MacroForm
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(352, 125);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.playBackMacroButton);
            this.Controls.Add(this.recordStopButton);
            this.Controls.Add(this.recordStartButton);
            this.Name = "MacroForm";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new MacroForm());
		}

		private void Form1_Load(object sender, System.EventArgs e)
		{
		
			mouseHook.MouseMove += new MouseEventHandler(mouseHook_MouseMove);
			mouseHook.MouseDown += new MouseEventHandler(mouseHook_MouseDown);
			mouseHook.MouseUp += new MouseEventHandler(mouseHook_MouseUp);

			keyboardHook.KeyDown += new KeyEventHandler(keyboardHook_KeyDown);
			keyboardHook.KeyUp += new KeyEventHandler(keyboardHook_KeyUp);

		}

		void mouseHook_MouseMove(object sender, MouseEventArgs e)
		{

			events.Add(
				new MacroEvent(
				MacroEventType.MouseMove,
				e,
				Environment.TickCount - lastTimeRecorded
				));

			lastTimeRecorded = Environment.TickCount;

		}

		void mouseHook_MouseDown(object sender, MouseEventArgs e)
		{

			events.Add(
				new MacroEvent(
				MacroEventType.MouseDown,
				e,
				Environment.TickCount - lastTimeRecorded
				));

			lastTimeRecorded = Environment.TickCount;

		}

		void mouseHook_MouseUp(object sender, MouseEventArgs e)
		{

			events.Add(
				new MacroEvent(
				MacroEventType.MouseUp,
				e,
				Environment.TickCount - lastTimeRecorded
				));

			lastTimeRecorded = Environment.TickCount;

		}

		void keyboardHook_KeyDown(object sender, KeyEventArgs e)
		{

			events.Add(
				new MacroEvent(
				MacroEventType.KeyDown,
				e,
				Environment.TickCount - lastTimeRecorded
				));

			lastTimeRecorded = Environment.TickCount;

		}

		void keyboardHook_KeyUp(object sender, KeyEventArgs e)
		{

			events.Add(
				new MacroEvent(
				MacroEventType.KeyUp,
				e,
				Environment.TickCount - lastTimeRecorded
				));

			lastTimeRecorded = Environment.TickCount;

		}

		private void recordStartButton_Click(object sender, EventArgs e)
		{
            events.Clear();
			lastTimeRecorded = Environment.TickCount;

			keyboardHook.Start();
			mouseHook.Start();

		}


		private void recordStopButton_Click(object sender, EventArgs e)
		{
            

			keyboardHook.Stop();
			mouseHook.Stop();

		}

		private void playBackMacroButton_Click(object sender, EventArgs e)
		{

			foreach (MacroEvent macroEvent in events)
			{

				Thread.Sleep(macroEvent.TimeSinceLastEvent);
				

				switch (macroEvent.MacroEventType)
				{
					case MacroEventType.MouseMove:
					{

						MouseEventArgs mouseArgs = (MouseEventArgs)macroEvent.EventArgs;

						MouseSimulator.X = mouseArgs.X;
						MouseSimulator.Y = mouseArgs.Y;

					}
						break;
					case MacroEventType.MouseDown:
					{

						MouseEventArgs mouseArgs = (MouseEventArgs)macroEvent.EventArgs;

						MouseSimulator.MouseDown(mouseArgs.Button);

					}
						break;
					case MacroEventType.MouseUp:
					{

						MouseEventArgs mouseArgs = (MouseEventArgs)macroEvent.EventArgs;

						MouseSimulator.MouseUp(mouseArgs.Button);

					}
						break;
					case MacroEventType.KeyDown:
					{

						KeyEventArgs keyArgs = (KeyEventArgs)macroEvent.EventArgs;

						KeyboardSimulator.KeyDown(keyArgs.KeyCode);

					}
						break;
					case MacroEventType.KeyUp:
					{

						KeyEventArgs keyArgs = (KeyEventArgs)macroEvent.EventArgs;

						KeyboardSimulator.KeyUp(keyArgs.KeyCode);

					}
						break;
					default:
						break;
				}

			}

		}


	}
}
