#if !defined(AFX_GRIDINPUTDLG_H__BF91B163_6A84_11D3_A7FE_0080C8763FA4__INCLUDED_)
#define AFX_GRIDINPUTDLG_H__BF91B163_6A84_11D3_A7FE_0080C8763FA4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// GridInputDlg.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CGridInputDlg dialog

class CGridInputDlg : public CDialog
{
// Construction
public:
	CGridInputDlg(CWnd* pParent = NULL);   // standard constructor
	CString m_strChoiceList;

// Dialog Data
	//{{AFX_DATA(CGridInputDlg)
	enum { IDD = IDD_DIALOG_GRIDINPUT };
		// NOTE: the ClassWizard will add data members here
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CGridInputDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	CGrid m_wndGrid;

	// Generated message map functions
	//{{AFX_MSG(CGridInputDlg)
	virtual BOOL OnInitDialog();
	virtual void OnOK();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_GRIDINPUTDLG_H__BF91B163_6A84_11D3_A7FE_0080C8763FA4__INCLUDED_)
