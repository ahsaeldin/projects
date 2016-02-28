#pragma once



namespace GetAllsystemTrayIconsInformationSample
{
	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;
	using namespace Interop::ASTC;  
	using namespace System::Runtime::InteropServices;
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
	private: AxInterop::ASTC::AxTrayList *  axTrayList1;
	private: AxInterop::MSFlexGridLib::AxMSFlexGrid *  MSFlexGrid1;

			 TrayIconInfo OldTrayInfo;
	protected:
		void Dispose(Boolean disposing)
		{
			if (disposing && components)
			{
				components->Dispose();
			}
			__super::Dispose(disposing);
		}
	public private: System::Windows::Forms::Button *  ButRefresh;

	public private: System::Windows::Forms::Button *  ButHide;

	
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
			System::Resources::ResourceManager *  resources = new System::Resources::ResourceManager(__typeof(GetAllsystemTrayIconsInformationSample::Form1));
			this->ButRefresh = new System::Windows::Forms::Button();
			this->ButHide = new System::Windows::Forms::Button();
			this->axTrayList1 = new AxInterop::ASTC::AxTrayList();
			this->MSFlexGrid1 = new AxInterop::MSFlexGridLib::AxMSFlexGrid();
			(__try_cast<System::ComponentModel::ISupportInitialize *  >(this->axTrayList1))->BeginInit();
			(__try_cast<System::ComponentModel::ISupportInitialize *  >(this->MSFlexGrid1))->BeginInit();
			this->SuspendLayout();
			// 
			// ButRefresh
			// 
			this->ButRefresh->Location = System::Drawing::Point(104, 232);
			this->ButRefresh->Name = S"ButRefresh";
			this->ButRefresh->Size = System::Drawing::Size(80, 23);
			this->ButRefresh->TabIndex = 8;
			this->ButRefresh->Text = S"Refresh";
			this->ButRefresh->Click += new System::EventHandler(this, ButRefresh_Click);
			// 
			// ButHide
			// 
			this->ButHide->Location = System::Drawing::Point(16, 232);
			this->ButHide->Name = S"ButHide";
			this->ButHide->Size = System::Drawing::Size(80, 23);
			this->ButHide->TabIndex = 6;
			this->ButHide->Text = S"Hide";
			this->ButHide->Click += new System::EventHandler(this, ButHide_Click);
			// 
			// axTrayList1
			// 
			this->axTrayList1->Enabled = true;
			this->axTrayList1->Location = System::Drawing::Point(664, 224);
			this->axTrayList1->Name = S"axTrayList1";
			this->axTrayList1->OcxState = (__try_cast<System::Windows::Forms::AxHost::State *  >(resources->GetObject(S"axTrayList1.OcxState")));
			this->axTrayList1->Size = System::Drawing::Size(36, 36);
			this->axTrayList1->TabIndex = 9;
			// 
			// MSFlexGrid1
			// 
			this->MSFlexGrid1->Location = System::Drawing::Point(8, 8);
			this->MSFlexGrid1->Name = S"MSFlexGrid1";
			this->MSFlexGrid1->OcxState = (__try_cast<System::Windows::Forms::AxHost::State *  >(resources->GetObject(S"MSFlexGrid1.OcxState")));
			this->MSFlexGrid1->Size = System::Drawing::Size(688, 216);
			this->MSFlexGrid1->TabIndex = 10;
			// 
			// Form1
			// 
			this->AutoScaleBaseSize = System::Drawing::Size(5, 13);
			this->ClientSize = System::Drawing::Size(704, 262);
			this->Controls->Add(this->MSFlexGrid1);
			this->Controls->Add(this->axTrayList1);
			this->Controls->Add(this->ButRefresh);
			this->Controls->Add(this->ButHide);
			this->Name = S"Form1";
			this->StartPosition = System::Windows::Forms::FormStartPosition::CenterScreen;
			this->Text = S"VC++.NET Tray List Sample";
			this->TopMost = true;
			this->Load += new System::EventHandler(this, Form1_Load);
			(__try_cast<System::ComponentModel::ISupportInitialize *  >(this->axTrayList1))->EndInit();
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
				 TrayIconInfo TrayList[] = new TrayIconInfo[100];

				 initFlexGrid();
				 
				 //Get the System Tray Data
				 TrayList =  axTrayList1->GetSysTrayIcons();

				 int IconCount = System::Convert::ToInt32(axTrayList1->IconCount);
                 
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


private: System::Void ButHide_Click(System::Object *  sender, System::EventArgs *  e)
		 {
			short RowSelected = (short)MSFlexGrid1->RowSel;

			TrayIconInfo TrayList[] = new TrayIconInfo[100];
            TrayList =  axTrayList1->GetSysTrayIcons();

            //Save the Icon Info before Remove it in order to use
            //when we restore the icon
			
            OldTrayInfo = TrayList[RowSelected - 1];

		    //Hide the selected Item from the System Tray area
            axTrayList1->HideIcon(&RowSelected);
            //Clear the FlexGrid
            MSFlexGrid1->Clear();
            //fill the FlexGrid 
            FillFlexGrid();  

		 }

private: System::Void ButRefresh_Click(System::Object *  sender, System::EventArgs *  e)
		 {
			 FillFlexGrid();
		 }

};
}

