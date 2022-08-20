#if !defined(AFX_COMBOEDIT_H__BF91B167_6A84_11D3_A7FE_0080C8763FA4__INCLUDED_)
#define AFX_COMBOEDIT_H__BF91B167_6A84_11D3_A7FE_0080C8763FA4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// ComboEdit.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CComboEdit window
class CBHMDBComboCtrl;

class CComboEdit : public CEdit
{
// Construction
public:
	CComboEdit(CBHMDBComboCtrl * pCtrl);

// Attributes
public:

protected:
	// Callback control pointer.
	CBHMDBComboCtrl * m_pCtrl;

// Operations
public:
// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CComboEdit)
	//}}AFX_VIRTUAL

// Implementation
public:
	virtual ~CComboEdit();

	// Generated message map functions
protected:
	// Move caret right, skip delimiters.
	void MoveRight();

	// Move caret left, skip delimiters.
	void MoveLeft();

	//{{AFX_MSG(CComboEdit)
	afx_msg void OnChar(UINT nChar, UINT nRepCnt, UINT nFlags);
	afx_msg void OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);
	afx_msg void OnSetFocus(CWnd* pOldWnd);
	afx_msg void OnLButtonDown(UINT nFlags, CPoint point);
	afx_msg	void OnPaste();
	afx_msg	void OnCut();
	afx_msg	void OnClear();
	afx_msg void OnKillFocus(CWnd* pNewWnd);
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_COMBOEDIT_H__BF91B167_6A84_11D3_A7FE_0080C8763FA4__INCLUDED_)
