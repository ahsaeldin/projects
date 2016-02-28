// TrayIconDlg.cpp : implementation file
//

#include "stdafx.h"
#include "TrayIcon.h"
#include "TrayIconDlg.h"
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
// CTrayIconDlg dialog

IMPLEMENT_DYNAMIC(CTrayIconDlg, CDialog);

CTrayIconDlg::CTrayIconDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CTrayIconDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CTrayIconDlg)
	//}}AFX_DATA_INIT
	// Note that LoadIcon does not require a subsequent DestroyIcon in Win32
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	m_pAutoProxy = NULL;
}

CTrayIconDlg::~CTrayIconDlg()
{
	// If there is an automation proxy for this dialog, set
	//  its back pointer to this dialog to NULL, so it knows
	//  the dialog has been deleted.
	if (m_pAutoProxy != NULL)
		m_pAutoProxy->m_pDialog = NULL;
}

void CTrayIconDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CTrayIconDlg)
	DDX_Control(pDX, IDC_StopTrack, m_strack);
	DDX_Control(pDX, IDC_TRPOP, m_ButTrack);
	DDX_Control(pDX, IDC_SBalloon2, m_ButCBalloon);
	DDX_Control(pDX, IDC_SBalloon, m_ButBalloon);
	DDX_Control(pDX, IDC_Hide, m_ButHide);
	DDX_Control(pDX, IDC_ChangeIcon, m_ChangeIcon);
	DDX_Control(pDX, IDC_BUTSHOW, m_ButShow);
	DDX_Control(pDX, IDC_TRAYICON1, m_TrayIcon);
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CTrayIconDlg, CDialog)
	//{{AFX_MSG_MAP(CTrayIconDlg)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_WM_CLOSE()
	ON_BN_CLICKED(IDC_BUTSHOW, OnButshow)
	ON_BN_CLICKED(IDC_ChangeIcon, OnChangeIcon)
	ON_BN_CLICKED(IDC_Hide, OnHide)
	ON_BN_CLICKED(IDC_SBalloon, OnSBalloon)
	ON_BN_CLICKED(IDC_SBalloon2, OnSBalloon2)
	ON_BN_CLICKED(IDC_TRPOP, OnTrpop)
	ON_BN_CLICKED(IDC_StopTrack, OnStopTrack)
	ON_COMMAND(ID_MAIN_SHOW, OnMainShow)
	ON_COMMAND(ID_MAIN_EXIT, OnMainExit)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CTrayIconDlg message handlers

BOOL CTrayIconDlg::OnInitDialog()
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

void CTrayIconDlg::OnSysCommand(UINT nID, LPARAM lParam)
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

void CTrayIconDlg::OnPaint() 
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
HCURSOR CTrayIconDlg::OnQueryDragIcon()
{
	return (HCURSOR) m_hIcon;
}

// Automation servers should not exit when a user closes the UI
//  if a controller still holds on to one of its objects.  These
//  message handlers make sure that if the proxy is still in use,
//  then the UI is hidden but the dialog remains around if it
//  is dismissed.

void CTrayIconDlg::OnClose() 
{   
    if (m_TrayIcon.GetIsDisplayed())
	m_TrayIcon.Hide();

	if (CanExit())
		CDialog::OnClose();
}

void CTrayIconDlg::OnOK() 
{
	if (CanExit())
		CDialog::OnOK();
}

void CTrayIconDlg::OnCancel() 
{
	if (CanExit())
		CDialog::OnCancel();
}

BOOL CTrayIconDlg::CanExit()
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
 Method:        CTrayIconDlg::OnButshow()
		
 Description:  Shows the icon in system tray.
 ==============================================================================
*/
void CTrayIconDlg::OnButshow() 
{
   /*
   Unlike the previous version of ASTC, you don't need to pass a window handle to Show function because
   ASTC dynamically creates an internal window related to ASTC instance, and destroy this window whenever
   you call Hide method or close your application.
   */
   m_TrayIcon.Show((long)m_hIcon,"VC++6.0TrayIcon"); 
}

/*
 ==============================================================================
 Method:        CTrayIconDlg::OnChangeIcon()
		
 Description:  Changes the icon in system tray area.
 ==============================================================================
*/
void CTrayIconDlg::OnChangeIcon() 
{
  m_TrayIcon.ChangeIcon((long)m_hIcon); 	
}


/*
 ==============================================================================
 Method:       CTrayIconDlg::OnHide()
		
 Description:  Removes the icon from system tray.
 ==============================================================================
*/
void CTrayIconDlg::OnHide() 
{
    m_TrayIcon.Hide();  	
}


/*
 '==============================================================================
 ' Method:       CTrayIconDlg::OnSBalloon()
 '
 ' Description:  Displays a balloon tooltip points to the icon.
 '==============================================================================
*/
void CTrayIconDlg::OnSBalloon() 
{
	m_TrayIcon.ShowBalloon("Click me","VC++ TrayIcon", 1,5000);
}

 /*
  '==============================================================================
  ' Method:       CTrayIconDlg::OnSBalloon2() 
  '
  ' Description:  Close the balloon tooltip.
  '==============================================================================
 */
void CTrayIconDlg::OnSBalloon2() 
{
	m_TrayIcon.CloseBalloon();
}

/*
 '==============================================================================
 ' Method:       CTrayIconDlg::OnTrpop() 
 '
 ' Description:  Start Tracking the popup menu in system tray area.
 '==============================================================================
 */
void CTrayIconDlg::OnTrpop() 
{
  m_TrayIcon.SetTrackPopMenu(true);
  m_strack.EnableWindow(true);
  AfxMessageBox( "when you right click the icon a popupmenu will appear,ASTC will Track the popupmenu and close it when you left click again"); 
}

/*
 '==============================================================================
 ' Method:       CTrayIconDlg::OnStopTrack() 
 '
 ' Description:  Stop track the popup menu in system tray area.
 '==============================================================================
*/
void CTrayIconDlg::OnStopTrack() 
{
    m_TrayIcon.SetTrackPopMenu(false);
   	m_strack.EnableWindow(false);
	AfxMessageBox( "ASTC stop tracking the popupmenu,Now right click the icon again and you will see that popupmenu willn't close unless you click an item."); 
}

BEGIN_EVENTSINK_MAP(CTrayIconDlg, CDialog)
    //{{AFX_EVENTSINK_MAP(CTrayIconDlg)
	ON_EVENT(CTrayIconDlg, IDC_TRAYICON1, 5 /* RightMouseUp */, OnRightMouseUpTrayicon1, VTS_NONE)
	ON_EVENT(CTrayIconDlg, IDC_TRAYICON1, 11 /* BalloonLeftClick */, OnBalloonLeftClickTrayicon1, VTS_NONE)
	//}}AFX_EVENTSINK_MAP
END_EVENTSINK_MAP()


/*     
 '====================================================================================
 ' Method:       CTrayIconDlg::OnRightMouseUpTrayicon1
 '
 ' Description:  Display a popup menu whenever you right click the icon in system tray.
 '====================================================================================
*/
void CTrayIconDlg::OnRightMouseUpTrayicon1() 
{
	CMenu menu1;
    CMenu * SubMenu1;
	menu1.LoadMenu(131); 
	SubMenu1 = menu1.GetSubMenu(0);
    m_TrayIcon.PopUp((long)SubMenu1->m_hMenu,(long)m_hWnd); 
}

/*
 '==============================================================================
 ' Method:       CTrayIconDlg::OnBalloonLeftClickTrayicon1
 '
 ' Description:  Display a message box whenever you left click the balloon tooltip.
 '==============================================================================
*/
void CTrayIconDlg::OnBalloonLeftClickTrayicon1() 
{
   AfxMessageBox( "The Balloon Tooltip had been left clicked."); 
}

void CTrayIconDlg::OnMainShow() 
{
	AfxMessageBox("show menu Item had been clicked.");
}

void CTrayIconDlg::OnMainExit() 
{
 
	m_TrayIcon.Hide();
    this->DestroyWindow(); 
	
}
