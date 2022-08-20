#if !defined(AFX_ColumnS_H__BF91B158_6A84_11D3_A7FE_0080C8763FA4__INCLUDED_)
#define AFX_ColumnS_H__BF91B158_6A84_11D3_A7FE_0080C8763FA4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// Columns.h : header file
//



/////////////////////////////////////////////////////////////////////////////
// CColumns command target
class CBHMDBGridCtrl;
class CBHMDBDropDownCtrl;
class CBHMDBComboCtrl;

class CColumns : public CCmdTarget
{
	DECLARE_DYNCREATE(CColumns)

	CColumns();           // protected constructor used by dynamic creation

// Attributes
public:

// Operations
public:
	virtual BOOL GetDispatchIID(IID* pIID);

	// Get row/col ordinal from variant value.
	int GetRowColFromVariant(VARIANT va);

	// Set callback pointers.
	SetCtrlPtr(CBHMDBGridCtrl * pCtrl);
	SetCtrlPtr(CBHMDBDropDownCtrl * pCtrl);
	SetCtrlPtr(CBHMDBComboCtrl * pCtrl);
	
	// Group index in grid.
	int m_nGroupIndex;

	// Grid window.
	CGrid * m_pGrid;

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CColumns)
	public:
	virtual void OnFinalRelease();
	//}}AFX_VIRTUAL

// Implementation
protected:
	// Throw error.
	void Error(SCODE sc, UINT nID);

	virtual ~CColumns();

	// Callback pointers.
	CBHMDBGridCtrl * m_pGridCtrl;
	CBHMDBDropDownCtrl * m_pDropDownCtrl;
	CBHMDBComboCtrl * m_pComboCtrl;

	// Generated message map functions
	//{{AFX_MSG(CColumns)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_MSG

	DECLARE_OLETYPELIB(CColumns)      // GetTypeInfo

	DECLARE_MESSAGE_MAP()
	// Generated OLE dispatch map functions
	//{{AFX_DISPATCH(CColumns)
	afx_msg short GetCount();
	afx_msg void Remove(const VARIANT FAR& ColIndex);
	afx_msg void Add(const VARIANT FAR& ColIndex);
	afx_msg void RemoveAll();
	afx_msg LPDISPATCH Item(const VARIANT FAR& ColIndex);
	//}}AFX_DISPATCH
	afx_msg LPUNKNOWN _NewEnum();
	DECLARE_DISPATCH_MAP()

	BEGIN_INTERFACE_PART(EnumVARIANT, IEnumVARIANT)
		STDMETHOD(Next)(THIS_ unsigned long celt, VARIANT FAR* rgvar, 
			unsigned long FAR* pceltFetched);
	    STDMETHOD(Skip)(THIS_ unsigned long celt) ;
		STDMETHOD(Reset)(THIS) ;
	    STDMETHOD(Clone)(THIS_ IEnumVARIANT FAR* FAR* ppenum) ;
		XEnumVARIANT() ;        // constructor to set m_posCurrent
	    int m_nCurrent ; // Next() requires we keep track of our current item
	END_INTERFACE_PART(EnumVARIANT)    

	DECLARE_INTERFACE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_ColumnS_H__BF91B158_6A84_11D3_A7FE_0080C8763FA4__INCLUDED_)
