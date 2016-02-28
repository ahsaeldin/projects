// Tray.h : main header file for the TRAY application
//

#if !defined(AFX_TRAY_H__E52F4F6B_F546_4E3A_9126_68BF6E489A3B__INCLUDED_)
#define AFX_TRAY_H__E52F4F6B_F546_4E3A_9126_68BF6E489A3B__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// main symbols

/////////////////////////////////////////////////////////////////////////////
// CTrayApp:
// See Tray.cpp for the implementation of this class
//

class CTrayApp : public CWinApp
{
public:
	CTrayApp();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CTrayApp)
	public:
	virtual BOOL InitInstance();
	//}}AFX_VIRTUAL

// Implementation

	//{{AFX_MSG(CTrayApp)
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_TRAY_H__E52F4F6B_F546_4E3A_9126_68BF6E489A3B__INCLUDED_)
