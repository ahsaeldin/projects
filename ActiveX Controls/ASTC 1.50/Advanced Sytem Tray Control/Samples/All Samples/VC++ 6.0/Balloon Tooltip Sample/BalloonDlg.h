// BalloonDlg.h : header file
//
//{{AFX_INCLUDES()
#include "_balloon.h"
#include "_balloon1.h"
//}}AFX_INCLUDES

#if !defined(AFX_BALLOONDLG_H__1696F1D7_DD2E_4C8B_9FAE_6ECFC6E50265__INCLUDED_)
#define AFX_BALLOONDLG_H__1696F1D7_DD2E_4C8B_9FAE_6ECFC6E50265__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class CBalloonDlgAutoProxy;

/////////////////////////////////////////////////////////////////////////////
// CBalloonDlg dialog

class CBalloonDlg : public CDialog
{
	DECLARE_DYNAMIC(CBalloonDlg);
	friend class CBalloonDlgAutoProxy;

// Construction
public:
	CBalloonDlg(CWnd* pParent = NULL);	// standard constructor
	virtual ~CBalloonDlg();

// Dialog Data
	//{{AFX_DATA(CBalloonDlg)
	enum { IDD = IDD_BALLOON_DIALOG };
	CButton	m_But3;
	CButton	m_But2;
	CButton	m_But1;
	C_Balloon1	m_Balloon1;
	C_Balloon1	m_Balloon2;
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CBalloonDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	CBalloonDlgAutoProxy* m_pAutoProxy;
	HICON m_hIcon;

	BOOL CanExit();

	// Generated message map functions
	//{{AFX_MSG(CBalloonDlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg void OnClose();
	virtual void OnOK();
	virtual void OnCancel();
	afx_msg void OnButton1();
	afx_msg void OnMouseMove(UINT nFlags, CPoint point);
	afx_msg void OnButton2();
	afx_msg void OnButton3();
	afx_msg void OnBalloonLeftClickBalloon1();
	DECLARE_EVENTSINK_MAP()
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_BALLOONDLG_H__1696F1D7_DD2E_4C8B_9FAE_6ECFC6E50265__INCLUDED_)
