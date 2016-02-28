/*
=========================================================================================
Publisher      CprinGold Software.
               http://www.cpringold.com
               support@cpringold.com


 Description:  A sample code demonstrate how to add your software icon to system tray area
               As well as how to use the advanced features related to system tray area.
==========================================================================================
ASTC provide you with a very simple object, TrayIcon object which is the responsible to add the most of system
Tray functions to your applications.

With a few lines i can demonstrate how TrayIcon object can help you

Use AxTrayIcon1.Show method to add an icon to system tray area
Use AxTrayIcon1.Hide method to hide the icon you had add by TrayIcon.Show method from system tray
Use AxTrayIcon1. ChangeIcon method to change the icon you had added by another icon
Use AxTrayIcon1.Animate method to animate the icon in the system tray
Use AxTrayIcon1.StopAnimateing method to stop animating the icon
Use AxTrayIcon1.ShowBalloon to display a tooltip balloon
Use AxTrayIcon1.HideBalloon to hide the balloon
Use AxTrayIcon1.Popup to display a popup menu
Use AxTrayIcon1.TrackPopupmenu to track the popup menu whenever it lost focus in order to close it.
Use AxTrayIcon1.IsDisplayed to check if the icon is displayed in the system tray area

Note that you have a custom types of events associated with TrayIcon Object like
BalloonShow, BalloonHide, MedMouseUp, LeftMouseUp, RightMouseUp, MedMouseDown, LeftMouseDown
RightMouseDown, MedMouseDBLCLK, LeftMouseDBLCLK, BalloonLeftClick, BalloonRightClick, RightMouseDBLCLK ()

Note that TrayIcon Object doesn't subclass the window in which it contains TrayIcon control
because TrayIcon Object dynamically creates an internal window related to it's instance,
and destroy this window whenever you call Hide method or close your application,
so don 't worry about subclassing
If you have any further questions or need more sample code don't hesitate to contact us
At support@cpringold.com
*/
#pragma once


namespace TrayIconSample
{
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
	public private: System::Windows::Forms::GroupBox *  groupBox1;
	public private: System::Windows::Forms::Button *  butSTrack;
	public private: System::Windows::Forms::Button *  butTrack;
	public private: System::Windows::Forms::Button *  butCBalloon;
	public private: System::Windows::Forms::Button *  butShowBalloon;
	public private: System::Windows::Forms::Button *  butStAni;
	public private: System::Windows::Forms::Button *  butAni;
	public private: System::Windows::Forms::Button *  butHide;
	public private: System::Windows::Forms::Button *  butChgIcon;
	public private: System::Windows::Forms::Button *  butShow;
	private: AxInterop::ASTC::AxTrayIcon *  axTrayIcon1;
	private: AxInterop::MSComctlLib::AxImageList *  axImageList1;
	public private: System::Windows::Forms::ContextMenu *  ContextMenu1;
	public private: System::Windows::Forms::MenuItem *  MenuItem1;
	public private: System::Windows::Forms::MenuItem *  MenuItem2;

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
			System::Resources::ResourceManager *  resources = new System::Resources::ResourceManager(__typeof(TrayIconSample::Form1));
			this->groupBox1 = new System::Windows::Forms::GroupBox();
			this->axImageList1 = new AxInterop::MSComctlLib::AxImageList();
			this->butSTrack = new System::Windows::Forms::Button();
			this->butTrack = new System::Windows::Forms::Button();
			this->butCBalloon = new System::Windows::Forms::Button();
			this->butShowBalloon = new System::Windows::Forms::Button();
			this->butStAni = new System::Windows::Forms::Button();
			this->butAni = new System::Windows::Forms::Button();
			this->butHide = new System::Windows::Forms::Button();
			this->butChgIcon = new System::Windows::Forms::Button();
			this->butShow = new System::Windows::Forms::Button();
			this->axTrayIcon1 = new AxInterop::ASTC::AxTrayIcon();
			this->ContextMenu1 = new System::Windows::Forms::ContextMenu();
			this->MenuItem1 = new System::Windows::Forms::MenuItem();
			this->MenuItem2 = new System::Windows::Forms::MenuItem();
			this->groupBox1->SuspendLayout();
			(__try_cast<System::ComponentModel::ISupportInitialize *  >(this->axImageList1))->BeginInit();
			(__try_cast<System::ComponentModel::ISupportInitialize *  >(this->axTrayIcon1))->BeginInit();
			this->SuspendLayout();
			// 
			// groupBox1
			// 
			this->groupBox1->Controls->Add(this->axImageList1);
			this->groupBox1->Controls->Add(this->butSTrack);
			this->groupBox1->Controls->Add(this->butTrack);
			this->groupBox1->Controls->Add(this->butCBalloon);
			this->groupBox1->Controls->Add(this->butShowBalloon);
			this->groupBox1->Controls->Add(this->butStAni);
			this->groupBox1->Controls->Add(this->butAni);
			this->groupBox1->Controls->Add(this->butHide);
			this->groupBox1->Controls->Add(this->butChgIcon);
			this->groupBox1->Controls->Add(this->butShow);
			this->groupBox1->Controls->Add(this->axTrayIcon1);
			this->groupBox1->Location = System::Drawing::Point(8, 8);
			this->groupBox1->Name = S"groupBox1";
			this->groupBox1->Size = System::Drawing::Size(216, 176);
			this->groupBox1->TabIndex = 8;
			this->groupBox1->TabStop = false;
			// 
			// axImageList1
			// 
			this->axImageList1->ContainingControl = this;
			this->axImageList1->Enabled = true;
			this->axImageList1->Location = System::Drawing::Point(0, 136);
			this->axImageList1->Name = S"axImageList1";
			this->axImageList1->OcxState = (__try_cast<System::Windows::Forms::AxHost::State *  >(resources->GetObject(S"axImageList1.OcxState")));
			this->axImageList1->Size = System::Drawing::Size(38, 38);
			this->axImageList1->TabIndex = 15;
			// 
			// butSTrack
			// 
			this->butSTrack->Enabled = false;
			this->butSTrack->Location = System::Drawing::Point(112, 112);
			this->butSTrack->Name = S"butSTrack";
			this->butSTrack->Size = System::Drawing::Size(96, 24);
			this->butSTrack->TabIndex = 14;
			this->butSTrack->Text = S"Stop Track";
			this->butSTrack->Click += new System::EventHandler(this, butSTrack_Click);
			// 
			// butTrack
			// 
			this->butTrack->Location = System::Drawing::Point(112, 80);
			this->butTrack->Name = S"butTrack";
			this->butTrack->Size = System::Drawing::Size(96, 24);
			this->butTrack->TabIndex = 13;
			this->butTrack->Text = S"TrackPopUp";
			this->butTrack->Click += new System::EventHandler(this, butTrack_Click);
			// 
			// butCBalloon
			// 
			this->butCBalloon->Location = System::Drawing::Point(112, 48);
			this->butCBalloon->Name = S"butCBalloon";
			this->butCBalloon->Size = System::Drawing::Size(96, 24);
			this->butCBalloon->TabIndex = 12;
			this->butCBalloon->Text = S"Close Balloon";
			this->butCBalloon->Click += new System::EventHandler(this, butCBalloon_Click);
			// 
			// butShowBalloon
			// 
			this->butShowBalloon->Location = System::Drawing::Point(112, 16);
			this->butShowBalloon->Name = S"butShowBalloon";
			this->butShowBalloon->Size = System::Drawing::Size(96, 24);
			this->butShowBalloon->TabIndex = 11;
			this->butShowBalloon->Text = S"Show Balloon";
			this->butShowBalloon->Click += new System::EventHandler(this, butShowBalloon_Click);
			// 
			// butStAni
			// 
			this->butStAni->Location = System::Drawing::Point(56, 144);
			this->butStAni->Name = S"butStAni";
			this->butStAni->Size = System::Drawing::Size(104, 24);
			this->butStAni->TabIndex = 9;
			this->butStAni->Text = S"Stop Animation";
			this->butStAni->Click += new System::EventHandler(this, butStAni_Click);
			// 
			// butAni
			// 
			this->butAni->Location = System::Drawing::Point(8, 112);
			this->butAni->Name = S"butAni";
			this->butAni->Size = System::Drawing::Size(96, 24);
			this->butAni->TabIndex = 7;
			this->butAni->Text = S"Animate";
			this->butAni->Click += new System::EventHandler(this, butAni_Click);
			// 
			// butHide
			// 
			this->butHide->Location = System::Drawing::Point(8, 80);
			this->butHide->Name = S"butHide";
			this->butHide->Size = System::Drawing::Size(96, 24);
			this->butHide->TabIndex = 6;
			this->butHide->Text = S"Hide Icon";
			this->butHide->Click += new System::EventHandler(this, butHide_Click);
			// 
			// butChgIcon
			// 
			this->butChgIcon->Location = System::Drawing::Point(8, 48);
			this->butChgIcon->Name = S"butChgIcon";
			this->butChgIcon->Size = System::Drawing::Size(96, 24);
			this->butChgIcon->TabIndex = 4;
			this->butChgIcon->Text = S"Change Icon";
			this->butChgIcon->Click += new System::EventHandler(this, butChgIcon_Click);
			// 
			// butShow
			// 
			this->butShow->Location = System::Drawing::Point(8, 16);
			this->butShow->Name = S"butShow";
			this->butShow->Size = System::Drawing::Size(96, 24);
			this->butShow->TabIndex = 2;
			this->butShow->Text = S"&Show Icon";
			this->butShow->Click += new System::EventHandler(this, butShow_Click);
			// 
			// axTrayIcon1
			// 
			this->axTrayIcon1->ContainingControl = this;
			this->axTrayIcon1->Enabled = true;
			this->axTrayIcon1->Location = System::Drawing::Point(176, 136);
			this->axTrayIcon1->Name = S"axTrayIcon1";
			this->axTrayIcon1->OcxState = (__try_cast<System::Windows::Forms::AxHost::State *  >(resources->GetObject(S"axTrayIcon1.OcxState")));
			this->axTrayIcon1->Size = System::Drawing::Size(36, 36);
			this->axTrayIcon1->TabIndex = 9;
			this->axTrayIcon1->RightMouseUp += new System::EventHandler(this, axTrayIcon1_RightMouseUp);
			this->axTrayIcon1->BalloonLeftClick += new System::EventHandler(this, axTrayIcon1_BalloonLeftClick);
			// 
			// ContextMenu1
			// 
			System::Windows::Forms::MenuItem* __mcTemp__1[] = new System::Windows::Forms::MenuItem*[2];
			__mcTemp__1[0] = this->MenuItem1;
			__mcTemp__1[1] = this->MenuItem2;
			this->ContextMenu1->MenuItems->AddRange(__mcTemp__1);
			// 
			// MenuItem1
			// 
			this->MenuItem1->Index = 0;
			this->MenuItem1->Text = S"main";
			this->MenuItem1->Click += new System::EventHandler(this, MenuItem1_Click);
			// 
			// MenuItem2
			// 
			this->MenuItem2->Index = 1;
			this->MenuItem2->Text = S"Exit";
			this->MenuItem2->Click += new System::EventHandler(this, MenuItem2_Click);
			// 
			// Form1
			// 
			this->AutoScaleBaseSize = System::Drawing::Size(5, 13);
			this->ClientSize = System::Drawing::Size(232, 190);
			this->Controls->Add(this->groupBox1);
			this->Name = S"Form1";
			this->StartPosition = System::Windows::Forms::FormStartPosition::CenterScreen;
			this->Text = S"VC++.NET TrayIcon Sample";
			this->Closing += new System::ComponentModel::CancelEventHandler(this, Form1_Closing);
			this->groupBox1->ResumeLayout(false);
			(__try_cast<System::ComponentModel::ISupportInitialize *  >(this->axImageList1))->EndInit();
			(__try_cast<System::ComponentModel::ISupportInitialize *  >(this->axTrayIcon1))->EndInit();
			this->ResumeLayout(false);

		}	

	/*    
	 ==============================================================================
	 Method:        butShow_Click
		
	 Description:  Shows the icon in system tray.
	 ==============================================================================
    */
	private: System::Void butShow_Click(System::Object *  sender, System::EventArgs *  e)
			 {
				 /*
		          Unlike the previous version of ASTC, you don't need to pass a window handle to Show function because
                  ASTC dynamically creates an internal window related to ASTC instance, and destroy this window whenever
                  you call Hide method or close your application.
	         	 */
				 axTrayIcon1->CtlShow(this->get_Icon()->Handle.ToInt32(),"VC++.net TrayIcon");  
		     }
/*
 ==============================================================================
 Method:        butChgIcon_Click
		
 Description:  Changes the icon in system tray area.
 ==============================================================================
*/
private: System::Void butChgIcon_Click(System::Object *  sender, System::EventArgs *  e)
		 {
			   axTrayIcon1->ChangeIcon(this->get_Icon()->Handle.ToInt32());
		 }

/*
 ==============================================================================
 Method:       butHide_Click
		
 Description:  Removes the icon from system tray.
 ==============================================================================
*/
private: System::Void butHide_Click(System::Object *  sender, System::EventArgs *  e)
		 {
			 axTrayIcon1->CtlHide(); 
		 }

/*
 ==================================================================================
 ' Method:        butAni_Click
 '
 ' Description:  Animate the icon in system tray using in icons in imagelist control.
 '==================================================================================
*/
private: System::Void butAni_Click(System::Object *  sender, System::EventArgs *  e)
		 {
			 axTrayIcon1->Animate(axImageList1,1);   
         }

/*
 '==============================================================================
 ' Method:       butStAni_Click
 '
 ' Description:  Stop animating the icons in system tray.
 '==============================================================================
*/
private: System::Void butStAni_Click(System::Object *  sender, System::EventArgs *  e)
		 {
			 axTrayIcon1->StopAnimateing(this->get_Icon()->Handle.ToInt32()); 
		 }

/*
 '==============================================================================
 ' Method:       butShowBalloon_Click
 '
 ' Description:  Displays a balloon tooltip points to the icon.
 '==============================================================================
*/
private: System::Void butShowBalloon_Click(System::Object *  sender, System::EventArgs *  e)
		 {
			 axTrayIcon1->ShowBalloon("VC++.NET Balloon Tooltip.","VC++ TrayIcon", Interop::ASTC::BIcoType::Info,5000);
		 }
/*
'==============================================================================
' Method:       butCBalloon_Click
'
' Description:  Close the balloon tooltip.
'==============================================================================
*/
private: System::Void butCBalloon_Click(System::Object *  sender, System::EventArgs *  e)
		 {
			  axTrayIcon1->CloseBalloon();
		 }

/*
'==============================================================================
' Method:       butTrack_Click
'
' Description:  Start Tracking the popup menu in system tray area.
'==============================================================================
*/
private: System::Void butTrack_Click(System::Object *  sender, System::EventArgs *  e)
		 {
		   axTrayIcon1->set_TrackPopMenu(true);
		   butSTrack->set_Enabled(true);
		   MessageBox::Show("when you right click the icon a popupmenu will appear,ASTC will Track the popupmenu and close it when you left click again"); 
		 }

/*
 '==============================================================================
 ' Method:       butSTrack_Click
 '
 ' Description:  Stop track the popup menu in system tray area.
 '==============================================================================
*/
private: System::Void butSTrack_Click(System::Object *  sender, System::EventArgs *  e)
		 {
		   axTrayIcon1->set_TrackPopMenu(false);
		   butSTrack->set_Enabled(false);
		   MessageBox::Show("ASTC stop tracking the popupmenu,Now right click the icon again and you will see that popup menu willn't close unless you click an item."); 
         }

/*
 '==============================================================================
 ' Method:       Form1_Closing
 '
 ' Description:  Removes the icon from system tray whenever the
 '               terminates the program.
 '==============================================================================
*/
private: System::Void Form1_Closing(System::Object *  sender, System::ComponentModel::CancelEventArgs *  e)
		 {
			 if (axTrayIcon1->get_IsDisplayed())
                axTrayIcon1->CtlHide();
		 }

private: System::Void MenuItem1_Click(System::Object *  sender, System::EventArgs *  e)
		 {
		    MessageBox::Show("main menu Item had been clicked.");
		 }
/*
 '==============================================================================
 ' Method:        MenuItem2_Click
 '
 ' Description:  Removes the icon from system tray and ends the program.
 '==============================================================================
 */
private: System::Void MenuItem2_Click(System::Object *  sender, System::EventArgs *  e)
		 {
			  if (axTrayIcon1->get_IsDisplayed())
                axTrayIcon1->CtlHide();
			  this->Close();
		 }
/*
 '==============================================================================
 ' Method:       AxTrayIcon1_BalloonLeftClick
 '
 ' Description:  Display a message box whenever you left click the balloon tooltip.
 '==============================================================================
 */
private: System::Void axTrayIcon1_BalloonLeftClick(System::Object *  sender, System::EventArgs *  e)
		 {
			 MessageBox::Show("The Balloon Tooltip had been left clicked.");
		 }

/*     
 '====================================================================================
 ' Method:       AxTrayIcon1_RightMouseUp
 '
 ' Description:  Display a popup menu whenever you right click the icon in system tray.
 '====================================================================================
*/
private: System::Void axTrayIcon1_RightMouseUp(System::Object *  sender, System::EventArgs *  e)
		 {
			 axTrayIcon1->PopUp(ContextMenu1->Handle.ToInt32(),this->Handle.ToInt32());
		 }

};
}


