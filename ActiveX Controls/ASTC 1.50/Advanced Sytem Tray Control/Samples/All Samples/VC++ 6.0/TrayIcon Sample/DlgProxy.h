// DlgProxy.h : header file
//

#if !defined(AFX_DLGPROXY_H__09CE3F03_960A_4C2D_BE74_38743886D061__INCLUDED_)
#define AFX_DLGPROXY_H__09CE3F03_960A_4C2D_BE74_38743886D061__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class CTrayIconDlg;

/////////////////////////////////////////////////////////////////////////////
// CTrayIconDlgAutoProxy command target

class CTrayIconDlgAutoProxy : public CCmdTarget
{
	DECLARE_DYNCREATE(CTrayIconDlgAutoProxy)

	CTrayIconDlgAutoProxy();           // protected constructor used by dynamic creation

// Attributes
public:
	CTrayIconDlg* m_pDialog;

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CTrayIconDlgAutoProxy)
	public:
	virtual void OnFinalRelease();
	//}}AFX_VIRTUAL

// Implementation
protected:
	virtual ~CTrayIconDlgAutoProxy();

	// Generated message map functions
	//{{AFX_MSG(CTrayIconDlgAutoProxy)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(CTrayIconDlgAutoProxy)

	// Generated OLE dispatch map functions
	//{{AFX_DISPATCH(CTrayIconDlgAutoProxy)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_DISPATCH
	DECLARE_DISPATCH_MAP()
	DECLARE_INTERFACE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_DLGPROXY_H__09CE3F03_960A_4C2D_BE74_38743886D061__INCLUDED_)
