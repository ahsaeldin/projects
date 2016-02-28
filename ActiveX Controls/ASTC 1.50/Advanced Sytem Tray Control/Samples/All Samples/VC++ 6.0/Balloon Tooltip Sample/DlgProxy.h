// DlgProxy.h : header file
//

#if !defined(AFX_DLGPROXY_H__9E278825_88A7_40D0_9922_6A3D8A333969__INCLUDED_)
#define AFX_DLGPROXY_H__9E278825_88A7_40D0_9922_6A3D8A333969__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class CBalloonDlg;

/////////////////////////////////////////////////////////////////////////////
// CBalloonDlgAutoProxy command target

class CBalloonDlgAutoProxy : public CCmdTarget
{
	DECLARE_DYNCREATE(CBalloonDlgAutoProxy)

	CBalloonDlgAutoProxy();           // protected constructor used by dynamic creation

// Attributes
public:
	CBalloonDlg* m_pDialog;

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CBalloonDlgAutoProxy)
	public:
	virtual void OnFinalRelease();
	//}}AFX_VIRTUAL

// Implementation
protected:
	virtual ~CBalloonDlgAutoProxy();

	// Generated message map functions
	//{{AFX_MSG(CBalloonDlgAutoProxy)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(CBalloonDlgAutoProxy)

	// Generated OLE dispatch map functions
	//{{AFX_DISPATCH(CBalloonDlgAutoProxy)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_DISPATCH
	DECLARE_DISPATCH_MAP()
	DECLARE_INTERFACE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_DLGPROXY_H__9E278825_88A7_40D0_9922_6A3D8A333969__INCLUDED_)
