#if !defined(AFX_BHMDBCOMBOPPGGENERAL_H__BF91B147_6A84_11D3_A7FE_0080C8763FA4__INCLUDED_)
#define AFX_BHMDBCOMBOPPGGENERAL_H__BF91B147_6A84_11D3_A7FE_0080C8763FA4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

// BHMDBComboPpgGeneral.h : Declaration of the CBHMDBComboPropPageGeneral property page class.

////////////////////////////////////////////////////////////////////////////
// CBHMDBComboPropPageGeneral : See BHMDBComboPpgGeneral.cpp.cpp for implementation.

class CBHMDBComboPropPageGeneral : public COlePropertyPage
{
	DECLARE_DYNCREATE(CBHMDBComboPropPageGeneral)
	DECLARE_OLECREATE_EX(CBHMDBComboPropPageGeneral)

// Constructor
public:
	CBHMDBComboPropPageGeneral();

// Dialog Data
	//{{AFX_DATA(CBHMDBComboPropPageGeneral)
	enum { IDD = IDD_PROPPAGE_BHMDBCOMBO };
	BOOL	m_bReadOnly;
	BOOL	m_bColumnHeader;
	BOOL	m_bListAutoPosition;
	BOOL	m_bListWidthAutoSize;
	int		m_nDividerStyle;
	int		m_nDividerType;
	int		m_nBorderStyle;
	int		m_nDataMode;
	CString	m_strFieldSeparator;
	long	m_nRows;
	long	m_nCols;
	long	m_n;
	long	m_nMinDropDownItems;
	long	m_nDefColWidth;
	long	m_nHeaderHeight;
	long	m_nRowHeight;
	long	m_nListWidth;
	CString	m_strFormat;
	long	m_nMaxLength;
	int		m_nStyle;
	BOOL	m_bGroupHeader;
	//}}AFX_DATA

// Implementation
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Message maps
protected:
	//{{AFX_MSG(CBHMDBComboPropPageGeneral)
		// NOTE - ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_BHMDBCOMBOPPGGENERAL_H__BF91B147_6A84_11D3_A7FE_0080C8763FA4__INCLUDED)
