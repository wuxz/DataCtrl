#if !defined(AFX_Column_H__BF91B155_6A84_11D3_A7FE_0080C8763FA4__INCLUDED_)
#define AFX_Column_H__BF91B155_6A84_11D3_A7FE_0080C8763FA4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// Column.h : header file
//



/////////////////////////////////////////////////////////////////////////////
// CColumn command target
class CBHMDBGridCtrl;
class CBHMDBDropDownCtrl;
class CBHMDBComboCtrl;

class CColumn : public CCmdTarget
{
	DECLARE_DYNCREATE(CColumn)

	CColumn();           // protected constructor used by dynamic creation

// Attributes
public:

// Operations
public:
	// Decides if one field can not use null value.
	BOOL IsColForceNotNullable(CString strField);

	// Set the callback control's pointer.
	SetCtrlPtr(CBHMDBGridCtrl * pCtrl);
	SetCtrlPtr(CBHMDBDropDownCtrl * pCtrl);
	SetCtrlPtr(CBHMDBComboCtrl * pCtrl);

	// The grid window.
	CGrid * m_pGrid;

	// Own properties.
	ColumnProp prop;

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CColumn)
	public:
	virtual void OnFinalRelease();
	//}}AFX_VIRTUAL

// Implementation
protected:
	// Get the data mode of callback control.
	int GetDataMode();

	// Decides if one field should be readonly.
	BOOL IsColForceLock(CString strField);

	// Update the choice list of one col.
	void CalcChoiceList(int nCol);

	// Translate color from OLE_COLOR to COLORREF.
	COLORREF TranslateColor(OLE_COLOR clr);

	// Throw error.
	void Error(SCODE sc, UINT nID);

	virtual ~CColumn();

	virtual BOOL GetDispatchIID(IID* pIID);

	// Get cell text in one row.
	BSTR GetCellText(long RowIndex);

	// Get cell value in one row.
	VARIANT GetCellValue(long RowIndex);

	// Recalc the choice list array.
	void CalcChoiceList(CString strList);

	// Callback control pointers.
	CBHMDBGridCtrl * m_pGridCtrl;
	CBHMDBDropDownCtrl * m_pDropDownCtrl;
	CBHMDBComboCtrl * m_pComboCtrl;

	// Choice list array.
	CStringArray m_arChoiceList;

	// Generated message map functions
	//{{AFX_MSG(CColumn)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()

	DECLARE_OLETYPELIB(CColumn)      // GetTypeInfo

	// Generated OLE dispatch map functions
	//{{AFX_DISPATCH(CColumn)
	afx_msg short GetFieldLen();
	afx_msg void SetFieldLen(short nNewValue);
	afx_msg BOOL GetAllowSizing();
	afx_msg void SetAllowSizing(BOOL bNewValue);
	afx_msg OLE_COLOR GetHeadForeColor();
	afx_msg void SetHeadForeColor(OLE_COLOR nNewValue);
	afx_msg OLE_COLOR GetHeadBackColor();
	afx_msg void SetHeadBackColor(OLE_COLOR nNewValue);
	afx_msg BSTR GetDataField();
	afx_msg void SetDataField(LPCTSTR lpszNewValue);
	afx_msg BOOL GetLocked();
	afx_msg void SetLocked(BOOL bNewValue);
	afx_msg BOOL GetVisible();
	afx_msg void SetVisible(BOOL bNewValue);
	afx_msg long GetWidth();
	afx_msg void SetWidth(long nNewValue);
	afx_msg long GetDataType();
	afx_msg void SetDataType(long nNewValue);
	afx_msg BOOL GetSelected();
	afx_msg void SetSelected(BOOL bNewValue);
	afx_msg long GetStyle();
	afx_msg void SetStyle(long nNewValue);
	afx_msg long GetCaptionAlignment();
	afx_msg void SetCaptionAlignment(long nNewValue);
	afx_msg BOOL GetColChanged();
	afx_msg BSTR GetName();
	afx_msg void SetName(LPCTSTR lpszNewValue);
	afx_msg BOOL GetNullable();
	afx_msg void SetNullable(BOOL bNewValue);
	afx_msg BSTR GetMask();
	afx_msg void SetMask(LPCTSTR lpszNewValue);
	afx_msg BSTR GetPromptChar();
	afx_msg void SetPromptChar(LPCTSTR lpszNewValue);
	afx_msg BSTR GetCaption();
	afx_msg void SetCaption(LPCTSTR lpszNewValue);
	afx_msg long GetAlignment();
	afx_msg void SetAlignment(long nNewValue);
	afx_msg OLE_COLOR GetForeColor();
	afx_msg void SetForeColor(OLE_COLOR nNewValue);
	afx_msg OLE_COLOR GetBackColor();
	afx_msg void SetBackColor(OLE_COLOR nNewValue);
	afx_msg long GetCase();
	afx_msg void SetCase(long nNewValue);
	afx_msg BSTR GetText();
	afx_msg void SetText(LPCTSTR lpszNewValue);
	afx_msg VARIANT GetValue();
	afx_msg void SetValue(const VARIANT FAR& newValue);
	afx_msg BOOL GetPromptInclude();
	afx_msg void SetPromptInclude(BOOL bNewValue);
	afx_msg short GetListCount();
	afx_msg short GetListIndex();
	afx_msg void SetListIndex(short nNewValue);
	afx_msg OLE_HANDLE GetDropDownHWnd();
	afx_msg void SetDropDownHWnd(OLE_HANDLE nNewValue);
	afx_msg BSTR CellText(const VARIANT FAR& Bookmark);
	afx_msg VARIANT CellValue(const VARIANT FAR& Bookmark);
	afx_msg BOOL IsCellValid();
	afx_msg void AddItem(LPCTSTR Item, const VARIANT FAR& Index);
	afx_msg void RemoveAll();
	afx_msg void RemoveItem(short Index);
	afx_msg BSTR GetList(short Index);
	afx_msg void SetList(short Index, LPCTSTR lpszNewValue);
	//}}AFX_DISPATCH
	DECLARE_DISPATCH_MAP()
	DECLARE_INTERFACE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_Column_H__BF91B155_6A84_11D3_A7FE_0080C8763FA4__INCLUDED_)
