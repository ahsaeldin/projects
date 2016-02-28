// DlgProxy.cpp : implementation file
//

#include "stdafx.h"
#include "Balloon.h"
#include "DlgProxy.h"
#include "BalloonDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CBalloonDlgAutoProxy

IMPLEMENT_DYNCREATE(CBalloonDlgAutoProxy, CCmdTarget)

CBalloonDlgAutoProxy::CBalloonDlgAutoProxy()
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
	ASSERT_KINDOF(CBalloonDlg, AfxGetApp()->m_pMainWnd);
	m_pDialog = (CBalloonDlg*) AfxGetApp()->m_pMainWnd;
	m_pDialog->m_pAutoProxy = this;
}

CBalloonDlgAutoProxy::~CBalloonDlgAutoProxy()
{
	// To terminate the application when all objects created with
	// 	with automation, the destructor calls AfxOleUnlockApp.
	//  Among other things, this will destroy the main dialog
	if (m_pDialog != NULL)
		m_pDialog->m_pAutoProxy = NULL;
	AfxOleUnlockApp();
}

void CBalloonDlgAutoProxy::OnFinalRelease()
{
	// When the last reference for an automation object is released
	// OnFinalRelease is called.  The base class will automatically
	// deletes the object.  Add additional cleanup required for your
	// object before calling the base class.

	CCmdTarget::OnFinalRelease();
}

BEGIN_MESSAGE_MAP(CBalloonDlgAutoProxy, CCmdTarget)
	//{{AFX_MSG_MAP(CBalloonDlgAutoProxy)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(CBalloonDlgAutoProxy, CCmdTarget)
	//{{AFX_DISPATCH_MAP(CBalloonDlgAutoProxy)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_DISPATCH_MAP
END_DISPATCH_MAP()

// Note: we add support for IID_IBalloon to support typesafe binding
//  from VBA.  This IID must match the GUID that is attached to the 
//  dispinterface in the .ODL file.

// {04E62242-8FD3-4C9B-ABFE-BBD9D0D8BDB7}
static const IID IID_IBalloon =
{ 0x4e62242, 0x8fd3, 0x4c9b, { 0xab, 0xfe, 0xbb, 0xd9, 0xd0, 0xd8, 0xbd, 0xb7 } };

BEGIN_INTERFACE_MAP(CBalloonDlgAutoProxy, CCmdTarget)
	INTERFACE_PART(CBalloonDlgAutoProxy, IID_IBalloon, Dispatch)
END_INTERFACE_MAP()

// The IMPLEMENT_OLECREATE2 macro is defined in StdAfx.h of this project
// {541A05E5-213E-4A9E-BED7-2668E4BAF463}
IMPLEMENT_OLECREATE2(CBalloonDlgAutoProxy, "Balloon.Application", 0x541a05e5, 0x213e, 0x4a9e, 0xbe, 0xd7, 0x26, 0x68, 0xe4, 0xba, 0xf4, 0x63)

/////////////////////////////////////////////////////////////////////////////
// CBalloonDlgAutoProxy message handlers
