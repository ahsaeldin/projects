#pragma once


namespace VCTrayIcons
{   
	using namespace Interop::STI; 
	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;

	/// <summary> 
	/// Summary for Form1
	///
	/// WARNING: If you change the name of this class, you will need to change the 
	///          'Resource File Name' property for the managed resource compiler tool 
	///          associated with all .resx files this class depends on.  Otherwise,
	///          the designers will not be able to interact properly with localized
	///          resources associated with this form.
	/// </summary>
	public __gc class Form1 : public System::Windows::Forms::Form
	{	
	public:
		Form1(void)
		{
			InitializeComponent();
		}
  
	protected:
		void Dispose(Boolean disposing)
		{
			if (disposing && components)
			{
				components->Dispose();
			}
			__super::Dispose(disposing);
		}
	public private: System::Windows::Forms::GroupBox *  GroupBox1;
	public private: System::Windows::Forms::Button *  FeedBackButton;
	public private: System::Windows::Forms::Button *  HideButton;
	
	private: AxInterop::STI::AxTrayIcons *  axTrayIcons1;
	private: AxInterop::MSFlexGridLib::AxMSFlexGrid *  MSFlexGrid1;




	private:
		/// <summary>
		/// Required designer variable.
		/// </summary>
		System::ComponentModel::Container * components;

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		void InitializeComponent(void)
		{
			System::Resources::ResourceManager *  resources = new System::Resources::ResourceManager(__typeof(VCTrayIcons::Form1));
			this->GroupBox1 = new System::Windows::Forms::GroupBox();
			this->FeedBackButton = new System::Windows::Forms::Button();
			this->HideButton = new System::Windows::Forms::Button();
			this->axTrayIcons1 = new AxInterop::STI::AxTrayIcons();
			this->MSFlexGrid1 = new AxInterop::MSFlexGridLib::AxMSFlexGrid();
			this->GroupBox1->SuspendLayout();
			(__try_cast<System::ComponentModel::ISupportInitialize *  >(this->axTrayIcons1))->BeginInit();
			(__try_cast<System::ComponentModel::ISupportInitialize *  >(this->MSFlexGrid1))->BeginInit();
			this->SuspendLayout();
			// 
			// GroupBox1
			// 
			this->GroupBox1->Controls->Add(this->FeedBackButton);
			this->GroupBox1->Controls->Add(this->HideButton);
			this->GroupBox1->Location = System::Drawing::Point(8, 200);
			this->GroupBox1->Name = S"GroupBox1";
			this->GroupBox1->Size = System::Drawing::Size(168, 48);
			this->GroupBox1->TabIndex = 5;
			this->GroupBox1->TabStop = false;
			// 
			// FeedBackButton
			// 
			this->FeedBackButton->Location = System::Drawing::Point(88, 16);
			this->FeedBackButton->Name = S"FeedBackButton";
			this->FeedBackButton->Size = System::Drawing::Size(72, 24);
			this->FeedBackButton->TabIndex = 1;
			this->FeedBackButton->Text = S"&FeedBack";
			this->FeedBackButton->Click += new System::EventHandler(this, FeedBackButton_Click);
			// 
			// HideButton
			// 
			this->HideButton->Location = System::Drawing::Point(8, 16);
			this->HideButton->Name = S"HideButton";
			this->HideButton->Size = System::Drawing::Size(72, 24);
			this->HideButton->TabIndex = 0;
			this->HideButton->Text = S"&Hide";
			this->HideButton->Click += new System::EventHandler(this, HideButton_Click);
			// 
			// axTrayIcons1
			// 
			this->axTrayIcons1->Enabled = true;
			this->axTrayIcons1->Location = System::Drawing::Point(176, 208);
			this->axTrayIcons1->Name = S"axTrayIcons1";
			this->axTrayIcons1->OcxState = (__try_cast<System::Windows::Forms::AxHost::State *  >(resources->GetObject(S"axTrayIcons1.OcxState")));
			this->axTrayIcons1->Size = System::Drawing::Size(36, 36);
			this->axTrayIcons1->TabIndex = 6;
			// 
			// MSFlexGrid1
			// 
			this->MSFlexGrid1->Location = System::Drawing::Point(8, 8);
			this->MSFlexGrid1->Name = S"MSFlexGrid1";
			this->MSFlexGrid1->OcxState = (__try_cast<System::Windows::Forms::AxHost::State *  >(resources->GetObject(S"MSFlexGrid1.OcxState")));
			this->MSFlexGrid1->Size = System::Drawing::Size(664, 192);
			this->MSFlexGrid1->TabIndex = 7;
			// 
			// Form1
			// 
			this->AutoScaleBaseSize = System::Drawing::Size(5, 13);
			this->ClientSize = System::Drawing::Size(680, 254);
			this->Controls->Add(this->MSFlexGrid1);
			this->Controls->Add(this->axTrayIcons1);
			this->Controls->Add(this->GroupBox1);
			this->Name = S"Form1";
			this->StartPosition = System::Windows::Forms::FormStartPosition::CenterScreen;
			this->Text = S"VC++.NET TrayIcons";
			this->Load += new System::EventHandler(this, Form1_Load);
			this->GroupBox1->ResumeLayout(false);
			(__try_cast<System::ComponentModel::ISupportInitialize *  >(this->axTrayIcons1))->EndInit();
			(__try_cast<System::ComponentModel::ISupportInitialize *  >(this->MSFlexGrid1))->EndInit();
			this->ResumeLayout(false);

		}	
    private: System::Void Form1_Load(System::Object *  sender, System::EventArgs *  e)
			 {
			    FillFlexGrid();
			 }
    private: System::Void initFlexGrid()
			 {
			   MSFlexGrid1->set_ColWidth(0, 300);
               MSFlexGrid1->set_TextMatrix(0, 0, "I");
               MSFlexGrid1->set_ColWidth(1, 3000);
               MSFlexGrid1->set_TextMatrix(0, 1, "Application Path");
               MSFlexGrid1->set_ColWidth(2, 1300);
               MSFlexGrid1->set_TextMatrix(0, 2, "uID");
               MSFlexGrid1->set_ColWidth(3, 1300);
               MSFlexGrid1->set_TextMatrix(0, 3, "hWnd");
               MSFlexGrid1->set_ColWidth(4, 1300);
               MSFlexGrid1->set_TextMatrix(0, 4, "hIcon");
               MSFlexGrid1->set_ColWidth(5, 1700);
               MSFlexGrid1->set_TextMatrix(0, 5, "ToolTip");
               MSFlexGrid1->set_ColWidth(6, 1500);
               MSFlexGrid1->set_TextMatrix(0, 6, "uCallbackMessage");
			   MSFlexGrid1->ForeColorFixed = Color::Blue;
			 }
    private: System::Void FillFlexGrid()
			 {   
				 TrayIcon TrayList[] = new  TrayIcon[100];  
				 initFlexGrid();
				 
				 //Get the System Tray Data
				 TrayList =  axTrayIcons1->GetSysTrayIcons();

				 int IconCount = System::Convert::ToInt32(axTrayIcons1->IconCount);
                 
				 //Fill the FlexGrid with Data
				 for (int i = 0; i != IconCount ; i++ )
				 {
				   	int count = i + 1; 
                     MSFlexGrid1->set_TextMatrix(i + 1, 0, count.ToString() );
					 if (TrayList[i].APath == "")
				     {
					   MSFlexGrid1->set_TextMatrix(i + 1, 1, "N/A");
				     }
				     else
				     {
					   MSFlexGrid1->set_TextMatrix(i + 1, 1, TrayList[i].APath);
				     }
				    
					 MSFlexGrid1->set_TextMatrix(i + 1, 2, TrayList[i].uId.ToString());
                     MSFlexGrid1->set_TextMatrix(i + 1, 3, TrayList[i].hwnd.ToString() );
                     MSFlexGrid1->set_TextMatrix(i + 1, 4, TrayList[i].hIcon.ToString() );
                     MSFlexGrid1->set_TextMatrix(i + 1, 5, TrayList[i].ToolTip);
                     MSFlexGrid1->set_TextMatrix(i + 1, 6, TrayList[i].ucallbackMessage.ToString());
				    
				 }

			 
			 }
    private: System::Void FeedBackButton_Click(System::Object *  sender, System::EventArgs *  e)
		 {
		    axTrayIcons1->sendreview(); 
		 }

    private: System::Void HideButton_Click(System::Object *  sender, System::EventArgs *  e)
		 {
            short RowSelected = (short)MSFlexGrid1->RowSel;
		    //Hide the selected Item from the System Tray area
            axTrayIcons1->CtlHide(&RowSelected);
            //Clear the FlexGrid
            MSFlexGrid1->Clear();
            //fill the FlexGrid 
            FillFlexGrid();   		       
		 }

};
}


