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

// BalloonDlg.cpp : implementation file
//

#include "stdafx.h"
#include "Balloon.h"
#include "BalloonDlg.h"
#include "DlgProxy.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CAboutDlg dialog used for App About

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// Dialog Data
	//{{AFX_DATA(CAboutDlg)
	enum { IDD = IDD_ABOUTBOX };
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CAboutDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	//{{AFX_MSG(CAboutDlg)
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
	//{{AFX_DATA_INIT(CAboutDlg)
	//}}AFX_DATA_INIT
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CAboutDlg)
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
	//{{AFX_MSG_MAP(CAboutDlg)
		// No message handlers
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CBalloonDlg dialog

IMPLEMENT_DYNAMIC(CBalloonDlg, CDialog);

CBalloonDlg::CBalloonDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CBalloonDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CBalloonDlg)
	//}}AFX_DATA_INIT
	// Note that LoadIcon does not require a subsequent DestroyIcon in Win32
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	m_pAutoProxy = NULL;
}

CBalloonDlg::~CBalloonDlg()
{
	// If there is an automation proxy for this dialog, set
	//  its back pointer to this dialog to NULL, so it knows
	//  the dialog has been deleted.
	if (m_pAutoProxy != NULL)
		m_pAutoProxy->m_pDialog = NULL;
}

void CBalloonDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CBalloonDlg)
	DDX_Control(pDX, IDC_BUTTON3, m_But3);
	DDX_Control(pDX, IDC_BUTTON2, m_But2);
	DDX_Control(pDX, IDC_BUTTON1, m_But1);
	DDX_Control(pDX, IDC_BALLOON1, m_Balloon1);
	DDX_Control(pDX, IDC_BALLOON2, m_Balloon2);
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CBalloonDlg, CDialog)
	//{{AFX_MSG_MAP(CBalloonDlg)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_WM_CLOSE()
	ON_BN_CLICKED(IDC_BUTTON1, OnButton1)
	ON_WM_MOUSEMOVE()
	ON_BN_CLICKED(IDC_BUTTON2, OnButton2)
	ON_BN_CLICKED(IDC_BUTTON3, OnButton3)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CBalloonDlg message handlers

BOOL CBalloonDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// Add "About..." menu item to system menu.

	// IDM_ABOUTBOX must be in the system command range.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		CString strAboutMenu;
		strAboutMenu.LoadString(IDS_ABOUTBOX);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon
	
	// TODO: Add extra initialization here
	
	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CBalloonDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CBalloonDlg::OnPaint() 
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, (WPARAM) dc.GetSafeHdc(), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

// The system calls this to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CBalloonDlg::OnQueryDragIcon()
{
	return (HCURSOR) m_hIcon;
}

// Automation servers should not exit when a user closes the UI
//  if a controller still holds on to one of its objects.  These
//  message handlers make sure that if the proxy is still in use,
//  then the UI is hidden but the dialog remains around if it
//  is dismissed.

void CBalloonDlg::OnClose() 
{
	if (CanExit())
		CDialog::OnClose();
}

void CBalloonDlg::OnOK() 
{
	if (CanExit())
		CDialog::OnOK();
}

void CBalloonDlg::OnCancel() 
{
	if (CanExit())
		CDialog::OnCancel();
}

BOOL CBalloonDlg::CanExit()
{
	// If the proxy object is still around, then the automation
	//  controller is still holding on to this application.  Leave
	//  the dialog around, but hide its UI.
	if (m_pAutoProxy != NULL)
	{
		ShowWindow(SW_HIDE);
		return FALSE;
	}

	return TRUE;
}


/*
   ==============================================================================
   Method:       CBalloonDlg::OnButton1
	
   Description:  Display a balloon tooltip for Balloon1.
   ===============================================================================
*/
void CBalloonDlg::OnButton1() 
{
	struct HWND__* ButHandle;
  
	ButHandle = m_But1.GetSafeHwnd();
	m_Balloon1.SetStyle(0); 
	m_Balloon1.ShowBalloon((long *)&ButHandle,"Click Me","Balloon1",1,500,5000);	

}


/*
  ===============================================================================
  Method:        CBalloonDlg::OnMouseMove
		
  Description:   Destroy the Balloon Tooltip.
  ================================================================================
*/
void CBalloonDlg::OnMouseMove(UINT nFlags, CPoint point) 
{
	
	m_Balloon1.Destroy();
	m_Balloon2.Destroy(); 
	CDialog::OnMouseMove(nFlags, point);
}


/*
  ==============================================================================
  Method:       CBalloonDlg::OnButton2
		
  Description:  Display a balloon tooltip for Balloon2.
  ===============================================================================
*/
void CBalloonDlg::OnButton2() 
{

    struct HWND__* ButHandle;
  
	ButHandle = m_But2.GetSafeHwnd();
	m_Balloon1.SetStyle(1); 
	m_Balloon1.ShowBalloon((long *)&ButHandle,"Rectangle Type","Balloon2",1,500,5000);	

}

/*
  ==============================================================================
  Method:       CBalloonDlg::OnButton3
		
  Description:  Display a balloon tooltip for Balloon3.
  ===============================================================================
*/
void CBalloonDlg::OnButton3() 
{

	struct HWND__* ButHandle;
  
	ButHandle = m_But3.GetSafeHwnd();
	m_Balloon2.SetStyle(0);
	m_Balloon2.SetBackColor(8389376);
	m_Balloon2.SetForeColor(52730);  
	m_Balloon2.ShowBalloon((long *)&ButHandle,"Customizable balloon tooltip","Balloon3",1,500,5000);	

}

BEGIN_EVENTSINK_MAP(CBalloonDlg, CDialog)
    //{{AFX_EVENTSINK_MAP(CBalloonDlg)
	ON_EVENT(CBalloonDlg, IDC_BALLOON1, 3 /* BalloonLeftClick */, OnBalloonLeftClickBalloon1, VTS_NONE)
	//}}AFX_EVENTSINK_MAP
END_EVENTSINK_MAP()

/*
  ==============================================================================
  ' Method:         CBalloonDlg::OnBalloonLeftClickBalloon1
  '
  ' Description:   Display a message Box when the left click the balloon.
  ==============================================================================
*/
void CBalloonDlg::OnBalloonLeftClickBalloon1() 
{
    AfxMessageBox("The Balloon Tooltip had been Clicked");  	
}
