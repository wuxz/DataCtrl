#if !defined(AFX_GROUP_H__0E9D7C61_7046_11D3_A7FE_0080C8763FA4__INCLUDED_)
#define AFX_GROUP_H__0E9D7C61_7046_11D3_A7FE_0080C8763FA4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// Group.h : header file
//



/////////////////////////////////////////////////////////////////////////////
// CGroup command target
class CBHMDBGridCtrl;
class CBHMDBDropDownCtrl;
class CBHMDBComboCtrl;

class CGroup : public CCmdTarget
{
	DECLARE_DYNCREATE(CGroup)

	CGroup();           // protected constructor used by dynamic creation

// Attributes
public:
	// Group index.
	int m_nGroupIndex;

	// Set callback pointers.
	void SetCtrlPtr(CBHMDBGridCtrl * pCtrl);
	void SetCtrlPtr(CBHMDBComboCtrl * pCtrl);
	void SetCtrlPtr(CBHMDBDropDownCtrl * pCtrl);

	// Callback grid window pointer.
	CGrid * m_pGrid;

protected:
	// Translate color.
	COLORREF TranslateColor(OLE_COLOR clr);

	// Throw error.
	void Error();

	virtual BOOL GetDispatchIID(IID* pIID);

	// Callback pointers.
	CBHMDBGridCtrl * m_pGridCtrl;
	CBHMDBDropDownCtrl * m_pDropDownCtrl;
	CBHMDBComboCtrl * m_pComboCtrl;

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CGroup)
	public:
	virtual void OnFinalRelease();
	//}}AFX_VIRTUAL

// Implementation
protected:
	void Error(SCODE sc, UINT nID);
	virtual ~CGroup();

	// Generated message map functions
	//{{AFX_MSG(CGroup)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()

	DECLARE_OLETYPELIB(CGroup)      // GetTypeInfo

	// Generated OLE dispatch map functions
	//{{AFX_DISPATCH(CGroup)
	afx_msg OLE_COLOR GetBackColor();
	afx_msg void SetBackColor(OLE_COLOR nNewValue);
	afx_msg OLE_COLOR GetForeColor();
	afx_msg void SetForeColor(OLE_COLOR nNewValue);
	afx_msg BSTR GetName();
	afx_msg void SetName(LPCTSTR lpszNewValue);
	afx_msg BSTR GetCaption();
	afx_msg void SetCaption(LPCTSTR lpszNewValue);
	afx_msg long GetWidth();
	afx_msg void SetWidth(long nNewValue);
	afx_msg BOOL GetVisible();
	afx_msg void SetVisible(BOOL bNewValue);
	afx_msg BOOL GetAllowSizing();
	afx_msg void SetAllowSizing(BOOL bNewValue);
	afx_msg LPDISPATCH GetColumns();
	//}}AFX_DISPATCH
	DECLARE_DISPATCH_MAP()
	DECLARE_INTERFACE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_GROUP_H__0E9D7C61_7046_11D3_A7FE_0080C8763FA4__INCLUDED_)
