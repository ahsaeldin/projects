using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

namespace TrayIcons
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class Form1 : System.Windows.Forms.Form
	{
		internal System.Windows.Forms.GroupBox GroupBox1;
		internal System.Windows.Forms.Button FeedBackButton;
		internal System.Windows.Forms.Button HideButton;
		
		private AxMSFlexGridLib.AxMSFlexGrid MSFlexGrid1;
		private AxSTI.AxTrayIcons axTrayIcons1;
		
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
			this.FeedBackButton = new System.Windows.Forms.Button();
			this.HideButton = new System.Windows.Forms.Button();
			this.MSFlexGrid1 = new AxMSFlexGridLib.AxMSFlexGrid();
			this.axTrayIcons1 = new AxSTI.AxTrayIcons();
			this.GroupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.MSFlexGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.axTrayIcons1)).BeginInit();
			this.SuspendLayout();
			// 
			// GroupBox1
			// 
			this.GroupBox1.Controls.Add(this.FeedBackButton);
			this.GroupBox1.Controls.Add(this.HideButton);
			this.GroupBox1.Location = new System.Drawing.Point(8, 200);
			this.GroupBox1.Name = "GroupBox1";
			this.GroupBox1.Size = new System.Drawing.Size(168, 48);
			this.GroupBox1.TabIndex = 4;
			this.GroupBox1.TabStop = false;
			// 
			// FeedBackButton
			// 
			this.FeedBackButton.Location = new System.Drawing.Point(88, 16);
			this.FeedBackButton.Name = "FeedBackButton";
			this.FeedBackButton.Size = new System.Drawing.Size(72, 24);
			this.FeedBackButton.TabIndex = 1;
			this.FeedBackButton.Text = "&FeedBack";
			this.FeedBackButton.Click += new System.EventHandler(this.FeedBackButton_Click);
			// 
			// HideButton
			// 
			this.HideButton.Location = new System.Drawing.Point(8, 16);
			this.HideButton.Name = "HideButton";
			this.HideButton.Size = new System.Drawing.Size(72, 24);
			this.HideButton.TabIndex = 0;
			this.HideButton.Text = "&Hide";
			this.HideButton.Click += new System.EventHandler(this.HideButton_Click);
			// 
			// MSFlexGrid1
			// 
			this.MSFlexGrid1.Location = new System.Drawing.Point(8, 8);
			this.MSFlexGrid1.Name = "MSFlexGrid1";
			this.MSFlexGrid1.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("MSFlexGrid1.OcxState")));
			this.MSFlexGrid1.Size = new System.Drawing.Size(664, 192);
			this.MSFlexGrid1.TabIndex = 6;
			// 
			// axTrayIcons1
			// 
			this.axTrayIcons1.Enabled = true;
			this.axTrayIcons1.Location = new System.Drawing.Point(176, 208);
			this.axTrayIcons1.Name = "axTrayIcons1";
			this.axTrayIcons1.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("axTrayIcons1.OcxState")));
			this.axTrayIcons1.Size = new System.Drawing.Size(36, 36);
			this.axTrayIcons1.TabIndex = 7;
			// 
			// Form1
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(680, 254);
			this.Controls.Add(this.axTrayIcons1);
			this.Controls.Add(this.MSFlexGrid1);
			this.Controls.Add(this.GroupBox1);
			this.Name = "Form1";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "VC#TrayIcons";
			this.Load += new System.EventHandler(this.Form1_Load);
			this.GroupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.MSFlexGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.axTrayIcons1)).EndInit();
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



		private void Form1_Load(object sender, System.EventArgs e)
		{
		    FillFlexGrid();
		}
	    
		private void initFlexGrid()
		{ 
         
           MSFlexGrid1.set_ColWidth(0, 300);
           MSFlexGrid1.set_TextMatrix(0, 0, "I");
           MSFlexGrid1.set_ColWidth(1, 3000);
           MSFlexGrid1.set_TextMatrix(0, 1, "Application Path");
           MSFlexGrid1.set_ColWidth(2, 1300);
           MSFlexGrid1.set_TextMatrix(0, 2, "uID");
           MSFlexGrid1.set_ColWidth(3, 1300);
           MSFlexGrid1.set_TextMatrix(0, 3, "hWnd");
           MSFlexGrid1.set_ColWidth(4, 1300);
           MSFlexGrid1.set_TextMatrix(0, 4, "hIcon");
           MSFlexGrid1.set_ColWidth(5, 1700);
           MSFlexGrid1.set_TextMatrix(0, 5, "ToolTip");
           MSFlexGrid1.set_ColWidth(6, 1500);
           MSFlexGrid1.set_TextMatrix(0, 6, "uCallbackMessage");
           MSFlexGrid1.ForeColorFixed = Color.Blue;
           
        }

        private void FillFlexGrid()
		{  
		   
		   STI.TrayIcon[] TrayList = new STI.TrayIcon[100]; 
           initFlexGrid();

           //Get the System Tray Data
           TrayList = (STI.TrayIcon[])axTrayIcons1.GetSysTrayIcons();
            
		   //Fill the FlexGrid with Data
			
			for (int i = 0; i != (int)axTrayIcons1.IconCount; i++ )
			{
			    int count = i + 1;
                MSFlexGrid1.set_TextMatrix(i + 1, 0, count.ToString() );
				if (TrayList[i].APath == "")
				{
					MSFlexGrid1.set_TextMatrix(i + 1, 1, "N/A");
				}
				else
				{
					MSFlexGrid1.set_TextMatrix(i + 1, 1, TrayList[i].APath);
				}
				MSFlexGrid1.set_TextMatrix(i + 1, 2, TrayList[i].uId.ToString() );
                MSFlexGrid1.set_TextMatrix(i + 1, 3, TrayList[i].hwnd.ToString() );
                MSFlexGrid1.set_TextMatrix(i + 1, 4, TrayList[i].hIcon.ToString() );
                MSFlexGrid1.set_TextMatrix(i + 1, 5, TrayList[i].ToolTip.ToString() );
                MSFlexGrid1.set_TextMatrix(i + 1, 6, TrayList[i].ucallbackMessage.ToString());
            }						   
															   
         }

		private void HideButton_Click(object sender, System.EventArgs e)
		{   
			short RowSelected = (short)MSFlexGrid1.RowSel;
		    //Hide the selected Item from the System Tray area
            axTrayIcons1.CtlHide(ref RowSelected);
            //Clear the FlexGrid
            MSFlexGrid1.Clear();
            //fill the FlexGrid 
            FillFlexGrid();
          
		}

		private void FeedBackButton_Click(object sender, System.EventArgs e)
		{
			axTrayIcons1.sendreview();  
		}

		
			
	
	}
}
