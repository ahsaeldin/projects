/*
=========================================================================================
Publisher      CprinGold Software.
               http://www.cpringold.com
               support@cpringold.com


 Description:  A sample code demonstrates how to use Balloon Object.
==========================================================================================

Use the Optional parameter PHwnd to insure that the balloon tooltip will not appear outside a specified
Window or a child window region like a command button or a text box.
Note that:

1.you can't set this parameter to any external window i.e. you can't set to an external Apps window

2.you can set this parameter to the window which contains the current instance of the balloon control

i.e.  For the following code

balloon1.ShowBalloon From1.hwnd

The hwnd is a handle to the window in which contains the Balloon control and that's means that the balloon
Will only appear if the mouse is over the Form itself not any of its child controls [child windows] in which
It contains.

3.you can set this parameter to any control of the window that contains the current instance of the balloon
Control i.e if you have a Form called Form1 which has a command button called Command1 and a text box called
text1, you can set the parameter as the following

balloon1.ShowBalloon command1.hwnd

balloon1.ShowBalloon text1.hwnd

And that's means the balloon will only appear if the mouse is over the control itself

4. If you didn't pass this optional parameter, Balloon1 object will use its parent window automatically
I.e. if you have a form called Form1, Balloon object will Form1.Hwnd for PHwnd parameter silently but in
This case the balloon tooltip may appear at any part of the Form or its Controls

Use DelayTime parameter to set the Delay Time before the balloon appear, the default value for this
Parameter is 2000 milliseconds and the maximum, value for DelayTime Parameter is 65,535 milliseconds,
Is equivalent to just over 1 minute?
The maximum, value for Timeout Parameter is 65,535 milliseconds, is equivalent to just over 1 minute.
*/

#pragma once


namespace BalloonTooltipSample
{
	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;
	using namespace Interop::ASTC;  

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
	public private: System::Windows::Forms::Button *  ButBalloon3;
	public private: System::Windows::Forms::Button *  ButBalloon2;
	public private: System::Windows::Forms::Button *  ButBalloon1;
	public private: System::Windows::Forms::Label *  Label1;
	private: AxInterop::ASTC::AxBalloon *  Balloon2;
	private: AxInterop::ASTC::AxBalloon *  Balloon1;


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
			System::Resources::ResourceManager *  resources = new System::Resources::ResourceManager(__typeof(BalloonTooltipSample::Form1));
			this->GroupBox1 = new System::Windows::Forms::GroupBox();
			this->ButBalloon3 = new System::Windows::Forms::Button();
			this->ButBalloon2 = new System::Windows::Forms::Button();
			this->ButBalloon1 = new System::Windows::Forms::Button();
			this->Label1 = new System::Windows::Forms::Label();
			this->Balloon2 = new AxInterop::ASTC::AxBalloon();
			this->Balloon1 = new AxInterop::ASTC::AxBalloon();
			this->GroupBox1->SuspendLayout();
			(__try_cast<System::ComponentModel::ISupportInitialize *  >(this->Balloon2))->BeginInit();
			(__try_cast<System::ComponentModel::ISupportInitialize *  >(this->Balloon1))->BeginInit();
			this->SuspendLayout();
			// 
			// GroupBox1
			// 
			this->GroupBox1->Controls->Add(this->Balloon1);
			this->GroupBox1->Controls->Add(this->Balloon2);
			this->GroupBox1->Controls->Add(this->ButBalloon3);
			this->GroupBox1->Controls->Add(this->ButBalloon2);
			this->GroupBox1->Controls->Add(this->ButBalloon1);
			this->GroupBox1->Controls->Add(this->Label1);
			this->GroupBox1->Location = System::Drawing::Point(8, 0);
			this->GroupBox1->Name = S"GroupBox1";
			this->GroupBox1->Size = System::Drawing::Size(272, 112);
			this->GroupBox1->TabIndex = 2;
			this->GroupBox1->TabStop = false;
			// 
			// ButBalloon3
			// 
			this->ButBalloon3->Location = System::Drawing::Point(192, 56);
			this->ButBalloon3->Name = S"ButBalloon3";
			this->ButBalloon3->TabIndex = 7;
			this->ButBalloon3->Text = S"Balloon3";
			this->ButBalloon3->MouseMove += new System::Windows::Forms::MouseEventHandler(this, ButBalloon3_MouseMove);
			// 
			// ButBalloon2
			// 
			this->ButBalloon2->Location = System::Drawing::Point(96, 80);
			this->ButBalloon2->Name = S"ButBalloon2";
			this->ButBalloon2->TabIndex = 6;
			this->ButBalloon2->Text = S"Balloon2";
			this->ButBalloon2->MouseMove += new System::Windows::Forms::MouseEventHandler(this, ButBalloon2_MouseMove);
			// 
			// ButBalloon1
			// 
			this->ButBalloon1->Location = System::Drawing::Point(8, 56);
			this->ButBalloon1->Name = S"ButBalloon1";
			this->ButBalloon1->TabIndex = 5;
			this->ButBalloon1->Text = S"Balloon1";
			this->ButBalloon1->MouseMove += new System::Windows::Forms::MouseEventHandler(this, ButBalloon1_MouseMove);
			// 
			// Label1
			// 
			this->Label1->ForeColor = System::Drawing::Color::Blue;
			this->Label1->Location = System::Drawing::Point(8, 16);
			this->Label1->Name = S"Label1";
			this->Label1->Size = System::Drawing::Size(256, 88);
			this->Label1->TabIndex = 0;
			this->Label1->Text = S"Move the Cursor to every command button to see the balloon tooltip.";
			this->Label1->MouseMove += new System::Windows::Forms::MouseEventHandler(this, Label1_MouseMove);
			// 
			// Balloon2
			// 
			this->Balloon2->ContainingControl = this;
			this->Balloon2->Enabled = true;
			this->Balloon2->Location = System::Drawing::Point(144, 48);
			this->Balloon2->Name = S"Balloon2";
			this->Balloon2->OcxState = (__try_cast<System::Windows::Forms::AxHost::State *  >(resources->GetObject(S"Balloon2.OcxState")));
			this->Balloon2->Size = System::Drawing::Size(27, 27);
			this->Balloon2->TabIndex = 8;
			// 
			// Balloon1
			// 
			this->Balloon1->ContainingControl = this;
			this->Balloon1->Enabled = true;
			this->Balloon1->Location = System::Drawing::Point(104, 48);
			this->Balloon1->Name = S"Balloon1";
			this->Balloon1->OcxState = (__try_cast<System::Windows::Forms::AxHost::State *  >(resources->GetObject(S"Balloon1.OcxState")));
			this->Balloon1->Size = System::Drawing::Size(27, 27);
			this->Balloon1->TabIndex = 9;
			this->Balloon1->BalloonLeftClick += new System::EventHandler(this, Balloon1_BalloonLeftClick);
			// 
			// Form1
			// 
			this->AutoScaleBaseSize = System::Drawing::Size(5, 13);
			this->ClientSize = System::Drawing::Size(288, 118);
			this->Controls->Add(this->GroupBox1);
			this->Name = S"Form1";
			this->StartPosition = System::Windows::Forms::FormStartPosition::CenterScreen;
			this->Text = S"VC++.Net Balloon Tooltip";
			this->TopMost = true;
			this->MouseMove += new System::Windows::Forms::MouseEventHandler(this, Form1_MouseMove);
			this->GroupBox1->ResumeLayout(false);
			(__try_cast<System::ComponentModel::ISupportInitialize *  >(this->Balloon2))->EndInit();
			(__try_cast<System::ComponentModel::ISupportInitialize *  >(this->Balloon1))->EndInit();
			this->ResumeLayout(false);

		}	

		/*
		 ===============================================================================
		 Method:        Form1_MouseMove
		
		 Description:   Destroy the Balloon Tooltip.
	     ================================================================================
		*/
		private: System::Void Form1_MouseMove(System::Object *  sender, System::Windows::Forms::MouseEventArgs *  e)
			 {
				 Balloon1->Destroy(); 
		         Balloon2->Destroy();  
			 }

		/*
		 ===============================================================================
		 Method:        Label1_MouseMove
		
		 Description:   Destroy the Balloon Tooltip.
		 ================================================================================
		*/
		
private: System::Void Label1_MouseMove(System::Object *  sender, System::Windows::Forms::MouseEventArgs *  e)
		 {
		       Balloon1->Destroy(); 
		       Balloon2->Destroy();   
		 }


		 /*
		 ==============================================================================
		 Method:       ButBalloon1_MouseMove
		
		 Description:  Display a balloon tooltip for ButBalloon1.
		 ===============================================================================
		 */
		
private: System::Void ButBalloon1_MouseMove(System::Object *  sender, System::Windows::Forms::MouseEventArgs *  e)
		 {
		    int ButHandle;
			ButHandle = ButBalloon1->Handle.ToInt32();
		    Balloon1->Style = BalloonStyle::BalloonType; 
			Balloon1->ShowBalloon(&ButHandle,"Click Me","Balloon1",BIcoType::Info,1000,5000);
		 }

		 /*
		 ==============================================================================
		 Method:       ButBalloon2_MouseMove
		
		 Description:  Display a balloon tooltip for ButBalloon2.
		 ===============================================================================
	     */
private: System::Void ButBalloon2_MouseMove(System::Object *  sender, System::Windows::Forms::MouseEventArgs *  e)
		 {
		      int ButHandle;
			  ButHandle = ButBalloon2->Handle.ToInt32();
			  Balloon1->Style = BalloonStyle::RectangleType;
			  Balloon1->ShowBalloon(&ButHandle,"Rectangle Type","Balloon2",BIcoType::Info,1000,5000);
		 }

		 /*
		 ==============================================================================
		 Method:       ButBalloon3_MouseMove
		
		 Description:  Display a balloon tooltip for ButBalloon3.
		 ===============================================================================
		*/
private: System::Void ButBalloon3_MouseMove(System::Object *  sender, System::Windows::Forms::MouseEventArgs *  e)
		 {
		  	int ButHandle;
			ButHandle = ButBalloon3->Handle.ToInt32();
			Balloon2->CtlForeColor = 9843455;
			Balloon2->CtlBackColor = 16777215;
			Balloon2->ShowBalloon(&ButHandle,"Customizable balloon tooltip","Balloon3",BIcoType::Info,1000,5000);
		 }

		/*
		 ==============================================================================
		' Method:        Balloon1_BalloonLeftClick
		'
		' Description:   Display a message Box when the left click the balloon.
		 ==============================================================================
		*/
private: System::Void Balloon1_BalloonLeftClick(System::Object *  sender, System::EventArgs *  e)
		 {
			 MessageBox::Show("The Balloon Tooltip had been Clicked");  
		 }

};
}


