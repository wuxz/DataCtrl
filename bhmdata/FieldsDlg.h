#if !defined(AFX_FIELDSDLG_H__BF91B162_6A84_11D3_A7FE_0080C8763FA4__INCLUDED_)
#define AFX_FIELDSDLG_H__BF91B162_6A84_11D3_A7FE_0080C8763FA4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// FieldsDlg.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CFieldsDlg dialog

class CFieldsDlg : public CDialog
{
// Construction
public:
	CFieldsDlg(CWnd* pParent = NULL);   // standard constructor

public:
	CStringArray m_arField;

// Dialog Data
	//{{AFX_DATA(CFieldsDlg)
	enum { IDD = IDD_DIALOG_FIELDS };
		// NOTE: the ClassWizard will add data members here
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CFieldsDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CFieldsDlg)
	virtual BOOL OnInitDialog();
	virtual void OnOK();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_FIELDSDLG_H__BF91B162_6A84_11D3_A7FE_0080C8763FA4__INCLUDED_)
