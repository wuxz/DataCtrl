#if !defined(AFX_SELBOOKMARKS_H__BF91B15F_6A84_11D3_A7FE_0080C8763FA4__INCLUDED_)
#define AFX_SELBOOKMARKS_H__BF91B15F_6A84_11D3_A7FE_0080C8763FA4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// SelBookmarks.h : header file
//



/////////////////////////////////////////////////////////////////////////////
// CSelBookmarks command target
class CBHMDBGridCtrl;

class CSelBookmarks : public CCmdTarget
{
	DECLARE_DYNCREATE(CSelBookmarks)

	CSelBookmarks();           // protected constructor used by dynamic creation

// Attributes
public:

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CSelBookmarks)
	public:
	virtual void OnFinalRelease();
	//}}AFX_VIRTUAL

// Implementation
protected:
	virtual ~CSelBookmarks();
	virtual BOOL GetDispatchIID(IID* pIID);

	// Set callback pointer.
	SetCtrlPtr(CBHMDBGridCtrl * pCtrl);

	// Callback pointer
	CBHMDBGridCtrl * m_pCtrl;

	// Generated message map functions
	//{{AFX_MSG(CSelBookmarks)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_MSG

	DECLARE_OLETYPELIB(CSelBookmarks)      // GetTypeInfo

	DECLARE_MESSAGE_MAP()
	// Generated OLE dispatch map functions
	//{{AFX_DISPATCH(CSelBookmarks)
	afx_msg long GetCount();
	afx_msg void Remove(long Index);
	afx_msg void Add(const VARIANT FAR& Bookmark);
	afx_msg void RemoveAll();
	afx_msg VARIANT Item(long Index);
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

	friend class CBHMDBGridCtrl;
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_SELBOOKMARKS_H__BF91B15F_6A84_11D3_A7FE_0080C8763FA4__INCLUDED_)
