#if !defined(AFX_DROPDOWNREALWND_H__BF91B153_6A84_11D3_A7FE_0080C8763FA4__INCLUDED_)
#define AFX_DROPDOWNREALWND_H__BF91B153_6A84_11D3_A7FE_0080C8763FA4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// DropDownRealWnd.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CDropDownRealWnd window
class CBHMDBDropDownCtrl;
class CBHMDBComboCtrl;

class CDropDownRealWnd : public CGrid
{
// Construction
public:
	CDropDownRealWnd(CBHMDBDropDownCtrl * pCtrl);
	CDropDownRealWnd(CBHMDBComboCtrl * pCtrl);

// Attributes
public:
	// Callback pointers.
	CBHMDBDropDownCtrl * m_pDropDownCtrl;
	CBHMDBComboCtrl * m_pComboCtrl;

// Operations
public:
	// Overloaded function when navigate to other cell.
	virtual BOOL SetCurrentCell(int nRow, int nCol);
	
	// The helper function to show/hide the window and modify the status variable.
	void Showwindow(BOOL bShow);

	// Return text to callback control.
	void ReturnText(int nRow);

	// Overloaded function when left mouse button clicked one cell.
	virtual void OnLButtonClickedCell(int nRow, int nCol);

	// Overloaded function before drawing cell text.
	virtual void OnLoadCellText(int nRow, int nCol, CString &strText);
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CDropDownRealWnd)
	//}}AFX_VIRTUAL

// Implementation
public:
	virtual ~CDropDownRealWnd();

	// Generated message map functions
protected:
	//{{AFX_MSG(CDropDownRealWnd)
	afx_msg void OnKillFocus(CWnd* pNewWnd);
	afx_msg void OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);
	afx_msg void OnHScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar);
	afx_msg void OnVScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_DROPDOWNREALWND_H__BF91B153_6A84_11D3_A7FE_0080C8763FA4__INCLUDED_)
