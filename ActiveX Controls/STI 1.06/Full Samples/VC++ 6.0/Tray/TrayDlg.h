// TrayDlg.h : header file
//
//{{AFX_INCLUDES()
#include "_trayicons.h"
#include "msflexgrid.h"
//}}AFX_INCLUDES

#if !defined(AFX_TRAYDLG_H__5A7231D9_2899_46B8_9137_34D34E4141E3__INCLUDED_)
#define AFX_TRAYDLG_H__5A7231D9_2899_46B8_9137_34D34E4141E3__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

/////////////////////////////////////////////////////////////////////////////
// CTrayDlg dialog

class CTrayDlg : public CDialog
{
// Construction
public:
	CTrayDlg(CWnd* pParent = NULL);	// standard constructor

// Dialog Data
	//{{AFX_DATA(CTrayDlg)
	enum { IDD = IDD_TRAY_DIALOG };
	C_TrayIcons	m_TrayIcons;
	CMSFlexGrid	m_Flexgrid;
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CTrayDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	//{{AFX_MSG(CTrayDlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg void initFlexGrid();
	afx_msg void FillFlexGrid();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg void OnButton4();
	afx_msg void OnButton3();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_TRAYDLG_H__5A7231D9_2899_46B8_9137_34D34E4141E3__INCLUDED_)
