#if !defined(AFX_BHMDBDROPDOWNPPGGENERAL_H__BF91B142_6A84_11D3_A7FE_0080C8763FA4__INCLUDED_)
#define AFX_BHMDBDROPDOWNPPGGENERAL_H__BF91B142_6A84_11D3_A7FE_0080C8763FA4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

// BHMDBDropDownPpgGeneral.h : Declaration of the CBHMDBDropDownPropPageGeneral property page class.

////////////////////////////////////////////////////////////////////////////
// CBHMDBDropDownPropPageGeneral : See BHMDBDropDownPpgGeneral.cpp.cpp for implementation.

class CBHMDBDropDownPropPageGeneral : public COlePropertyPage
{
	DECLARE_DYNCREATE(CBHMDBDropDownPropPageGeneral)
	DECLARE_OLECREATE_EX(CBHMDBDropDownPropPageGeneral)

// Constructor
public:
	CBHMDBDropDownPropPageGeneral();

// Dialog Data
	//{{AFX_DATA(CBHMDBDropDownPropPageGeneral)
	enum { IDD = IDD_PROPPAGE_BHMDBDROPDOWN };
	int		m_nDataMode;
	BOOL	m_bColumnHeader;
	BOOL	m_bListAutoPosition;
	BOOL	m_bListWidthAutoSize;
	int		m_nDividerStyle;
	int		m_nDividerType;
	int		m_nBorderStyle;
	CString	m_strFieldSeparator;
	long	m_nRows;
	long	m_nCols;
	long	m_nDefColWidth;
	long	m_nHeaderHeight;
	long	m_nRowHeight;
	long	m_nMaxDropDownItems;
	long	m_nMinDropDownItems;
	long	m_nListWidth;
	BOOL	m_bGroupHeader;
	//}}AFX_DATA

// Implementation
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Message maps
protected:
	//{{AFX_MSG(CBHMDBDropDownPropPageGeneral)
		// NOTE - ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_BHMDBDROPDOWNPPGGENERAL_H__BF91B142_6A84_11D3_A7FE_0080C8763FA4__INCLUDED)
