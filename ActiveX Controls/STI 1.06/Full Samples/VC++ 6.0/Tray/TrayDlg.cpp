// TrayDlg.cpp : implementation file
//

#include "stdafx.h"
#include "Tray.h"
#include "TrayDlg.h"

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
// CTrayDlg dialog

CTrayDlg::CTrayDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CTrayDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CTrayDlg)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
	// Note that LoadIcon does not require a subsequent DestroyIcon in Win32
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CTrayDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CTrayDlg)
	DDX_Control(pDX, IDC_TRAYICONS1, m_TrayIcons);
	DDX_Control(pDX, IDC_MSFLEXGRID1, m_Flexgrid);
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CTrayDlg, CDialog)
	//{{AFX_MSG_MAP(CTrayDlg)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BUTTON4, OnButton4)
	ON_BN_CLICKED(IDC_BUTTON3, OnButton3)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CTrayDlg message handlers

BOOL CTrayDlg::OnInitDialog()
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
	

	FillFlexGrid(); 
	// TODO: Add extra initialization here
	 
	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CTrayDlg::OnSysCommand(UINT nID, LPARAM lParam)
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

void CTrayDlg::OnPaint() 
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
HCURSOR CTrayDlg::OnQueryDragIcon()
{
	return (HCURSOR) m_hIcon;
}

void CTrayDlg::initFlexGrid()
{    
    m_Flexgrid.SetColWidth(0,300);
    m_Flexgrid.SetTextMatrix(0, 0, "I");
    m_Flexgrid.SetColWidth(1, 3000);
    m_Flexgrid.SetTextMatrix(0, 1, "Application Path");
    m_Flexgrid.SetColWidth(2, 1300);
    m_Flexgrid.SetTextMatrix(0, 2, "uID");
    m_Flexgrid.SetColWidth(3, 1300);
    m_Flexgrid.SetTextMatrix(0, 3, "hWnd");
    m_Flexgrid.SetColWidth(4, 1300);
    m_Flexgrid.SetTextMatrix(0, 4, "hIcon");
    m_Flexgrid.SetColWidth(5, 1700);
    m_Flexgrid.SetTextMatrix(0, 5, "ToolTip");
    m_Flexgrid.SetColWidth(6, 1500);
    m_Flexgrid.SetTextMatrix(0, 6, "uCallbackMessage");
}


void CTrayDlg::FillFlexGrid()
{    
    LPDISPATCH AHMED;
	initFlexGrid(); 
	m_Flexgrid.EnableAutomation();  
	AHMED = m_Flexgrid.GetIDispatch(false);
	m_TrayIcons.FillMSFlexGrid(AHMED);
}

void CTrayDlg::OnButton4() 
{
   m_TrayIcons.sendreview();  	
}

void CTrayDlg::OnButton3() 
{
	
}
