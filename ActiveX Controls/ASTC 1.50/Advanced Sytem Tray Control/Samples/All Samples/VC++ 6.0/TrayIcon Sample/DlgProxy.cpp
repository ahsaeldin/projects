// DlgProxy.cpp : implementation file
//

#include "stdafx.h"
#include "TrayIcon.h"
#include "DlgProxy.h"
#include "TrayIconDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CTrayIconDlgAutoProxy

IMPLEMENT_DYNCREATE(CTrayIconDlgAutoProxy, CCmdTarget)

CTrayIconDlgAutoProxy::CTrayIconDlgAutoProxy()
{
	EnableAutomation();
	
	// To keep the application running as long as an automation 
	//	object is active, the constructor calls AfxOleLockApp.
	AfxOleLockApp();

	// Get access to the dialog through the application's
	//  main window pointer.  Set the proxy's internal pointer
	//  to point to the dialog, and set the dialog's back pointer to
	//  this proxy.
	ASSERT (AfxGetApp()->m_pMainWnd != NULL);
	ASSERT_VALID (AfxGetApp()->m_pMainWnd);
	ASSERT_KINDOF(CTrayIconDlg, AfxGetApp()->m_pMainWnd);
	m_pDialog = (CTrayIconDlg*) AfxGetApp()->m_pMainWnd;
	m_pDialog->m_pAutoProxy = this;
}

CTrayIconDlgAutoProxy::~CTrayIconDlgAutoProxy()
{
	// To terminate the application when all objects created with
	// 	with automation, the destructor calls AfxOleUnlockApp.
	//  Among other things, this will destroy the main dialog
	if (m_pDialog != NULL)
		m_pDialog->m_pAutoProxy = NULL;
	AfxOleUnlockApp();
}

void CTrayIconDlgAutoProxy::OnFinalRelease()
{
	// When the last reference for an automation object is released
	// OnFinalRelease is called.  The base class will automatically
	// deletes the object.  Add additional cleanup required for your
	// object before calling the base class.

	CCmdTarget::OnFinalRelease();
}

BEGIN_MESSAGE_MAP(CTrayIconDlgAutoProxy, CCmdTarget)
	//{{AFX_MSG_MAP(CTrayIconDlgAutoProxy)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(CTrayIconDlgAutoProxy, CCmdTarget)
	//{{AFX_DISPATCH_MAP(CTrayIconDlgAutoProxy)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_DISPATCH_MAP
END_DISPATCH_MAP()

// Note: we add support for IID_ITrayIcon to support typesafe binding
//  from VBA.  This IID must match the GUID that is attached to the 
//  dispinterface in the .ODL file.

// {2646A23E-7E8C-4384-976A-68E4B45C082E}
static const IID IID_ITrayIcon =
{ 0x2646a23e, 0x7e8c, 0x4384, { 0x97, 0x6a, 0x68, 0xe4, 0xb4, 0x5c, 0x8, 0x2e } };

BEGIN_INTERFACE_MAP(CTrayIconDlgAutoProxy, CCmdTarget)
	INTERFACE_PART(CTrayIconDlgAutoProxy, IID_ITrayIcon, Dispatch)
END_INTERFACE_MAP()

// The IMPLEMENT_OLECREATE2 macro is defined in StdAfx.h of this project
// {70AACC8F-6DCA-498D-8DD2-B6F5F649EE60}
IMPLEMENT_OLECREATE2(CTrayIconDlgAutoProxy, "TrayIcon.Application", 0x70aacc8f, 0x6dca, 0x498d, 0x8d, 0xd2, 0xb6, 0xf5, 0xf6, 0x49, 0xee, 0x60)

/////////////////////////////////////////////////////////////////////////////
// CTrayIconDlgAutoProxy message handlers
