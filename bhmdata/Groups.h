#if !defined(AFX_GROUPS_H__AD466A41_703F_11D3_A7FE_0080C8763FA4__INCLUDED_)
#define AFX_GROUPS_H__AD466A41_703F_11D3_A7FE_0080C8763FA4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// Groups.h : header file
//



/////////////////////////////////////////////////////////////////////////////
// CGroups command target
class CBHMDBGridCtrl;
class CBHMDBDropDownCtrl;
class CBHMDBComboCtrl;

class CGroups : public CCmdTarget
{
	DECLARE_DYNCREATE(CGroups)

	CGroups();           // protected constructor used by dynamic creation

// Attributes
public:

// Operations
public:
	int GetGroupFromVariant(VARIANT va);
	virtual BOOL GetDispatchIID(IID* pIID);
	void SetCtrlPtr(CBHMDBGridCtrl * pCtrl);
	void SetCtrlPtr(CBHMDBComboCtrl * pCtrl);
	void SetCtrlPtr(CBHMDBDropDownCtrl * pCtrl);

	CGrid * m_pGrid;

protected:
	CBHMDBGridCtrl * m_pGridCtrl;
	CBHMDBDropDownCtrl * m_pDropDownCtrl;
	CBHMDBComboCtrl * m_pComboCtrl;

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CGroups)
	public:
	virtual void OnFinalRelease();
	//}}AFX_VIRTUAL

// Implementation
protected:
	int GetRowColFromVariant(VARIANT va);
	void Error(SCODE sc, UINT nID);
	virtual ~CGroups();

	// Generated message map functions
	//{{AFX_MSG(CGroups)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_MSG

	DECLARE_OLETYPELIB(CGroups)      // GetTypeInfo

	DECLARE_MESSAGE_MAP()
	// Generated OLE dispatch map functions
	//{{AFX_DISPATCH(CGroups)
	afx_msg short GetCount();
	afx_msg void Remove(const VARIANT FAR& GroupIndex);
	afx_msg void Add(const VARIANT FAR& GroupIndex);
	afx_msg void RemoveAll();
	afx_msg LPDISPATCH GetItem(const VARIANT FAR& GroupIndex);
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

#endif // !defined(AFX_GROUPS_H__AD466A41_703F_11D3_A7FE_0080C8763FA4__INCLUDED_)
