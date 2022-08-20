#if !defined(AFX_ADDCOLUMNDLG_H__BF91B161_6A84_11D3_A7FE_0080C8763FA4__INCLUDED_)
#define AFX_ADDCOLUMNDLG_H__BF91B161_6A84_11D3_A7FE_0080C8763FA4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// AddColumnDlg.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CAddColumnDlg dialog
class CBHMDBColumnPropPage;
class CBHMDBGroupsPropPage;

class CAddColumnDlg : public CDialog
{
// Construction
public:
	void SetPagePtr(CBHMDBColumnPropPage * pPage);
	void SetPagePtr(CBHMDBGroupsPropPage * pPage);
	CAddColumnDlg(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CAddColumnDlg)
	enum { IDD = IDD_DIALOG_ADDCOLUMN };
	CString	m_strName;
		// NOTE: the ClassWizard will add data members here
	//}}AFX_DATA


protected:
	// The pointer of the column prop page.
	CBHMDBColumnPropPage * m_pColumnPage;

	// The pointer of the groups prop page.
	CBHMDBGroupsPropPage * m_pGroupsPage;

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CAddColumnDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CAddColumnDlg)
		// NOTE: the ClassWizard will add member functions here
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_ADDCOLUMNDLG_H__BF91B161_6A84_11D3_A7FE_0080C8763FA4__INCLUDED_)
