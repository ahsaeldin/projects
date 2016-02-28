// TrayIconDlg.h : header file
//
//{{AFX_INCLUDES()
#include "_trayicon.h"
//}}AFX_INCLUDES

#if !defined(AFX_TRAYICONDLG_H__52EDA486_21E9_41AE_826B_FF74B9AAF7D5__INCLUDED_)
#define AFX_TRAYICONDLG_H__52EDA486_21E9_41AE_826B_FF74B9AAF7D5__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class CTrayIconDlgAutoProxy;

/////////////////////////////////////////////////////////////////////////////
// CTrayIconDlg dialog

class CTrayIconDlg : public CDialog
{
	DECLARE_DYNAMIC(CTrayIconDlg);
	friend class CTrayIconDlgAutoProxy;

// Construction
public:
	CTrayIconDlg(CWnd* pParent = NULL);	// standard constructor
	virtual ~CTrayIconDlg();

// Dialog Data
	//{{AFX_DATA(CTrayIconDlg)
	enum { IDD = IDD_TRAYICON_DIALOG };
	CButton	m_strack;
	CButton	m_ButTrack;
	CButton	m_ButCBalloon;
	CButton	m_ButBalloon;
	CButton	m_ButHide;
	CButton	m_ChangeIcon;
	CButton	m_ButShow;
	C_TrayIcon	m_TrayIcon;
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CTrayIconDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	CTrayIconDlgAutoProxy* m_pAutoProxy;
	HICON m_hIcon;

	BOOL CanExit();

	// Generated message map functions
	//{{AFX_MSG(CTrayIconDlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg void OnClose();
	virtual void OnOK();
	virtual void OnCancel();
	afx_msg void OnButshow();
	afx_msg void OnChangeIcon();
	afx_msg void OnHide();
	afx_msg void OnSBalloon();
	afx_msg void OnSBalloon2();
	afx_msg void OnTrpop();
	afx_msg void OnStopTrack();
	afx_msg void OnRightMouseUpTrayicon1();
	afx_msg void OnBalloonLeftClickTrayicon1();
	afx_msg void OnMainShow();
	afx_msg void OnMainExit();
	DECLARE_EVENTSINK_MAP()
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_TRAYICONDLG_H__52EDA486_21E9_41AE_826B_FF74B9AAF7D5__INCLUDED_)
