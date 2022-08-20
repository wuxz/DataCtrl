#if !defined(AFX_BHMDBGRIDPPGGENERAL_H__BF91B13D_6A84_11D3_A7FE_0080C8763FA4__INCLUDED_)
#define AFX_BHMDBGRIDPPGGENERAL_H__BF91B13D_6A84_11D3_A7FE_0080C8763FA4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

// BHMDBGridPpgGeneral.h : Declaration of the CBHMDBGridPropPageGeneral property page class.

////////////////////////////////////////////////////////////////////////////
// CBHMDBGridPropPageGeneral : See BHMDBGridPpgGeneral.cpp.cpp for implementation.

class CBHMDBGridPropPageGeneral : public COlePropertyPage
{
	DECLARE_DYNCREATE(CBHMDBGridPropPageGeneral)
	DECLARE_OLECREATE_EX(CBHMDBGridPropPageGeneral)

// Constructor
public:
	CBHMDBGridPropPageGeneral();

// Dialog Data
	//{{AFX_DATA(CBHMDBGridPropPageGeneral)
	enum { IDD = IDD_PROPPAGE_BHMDBGRID };
	int		m_nDataMode;
	BOOL	m_bAllowAddNew;
	BOOL	m_bAllowDelete;
	BOOL	m_bAllowUpdate;
	BOOL	m_bAllowRowSizing;
	BOOL	m_bAllowColMoving;
	BOOL	m_bAllowColSizing;
	BOOL	m_bColumnHeader;
	BOOL	m_bRecordSelectors;
	BOOL	m_bSelectByCell;
	CString	m_strCaption;
	CString	m_strFieldSeparator;
	int		m_nBorderStyle;
	int		m_nCaptionAlignment;
	int		m_nDividerStyle;
	int		m_nDividerType;
	long	m_nCols;
	long	m_nDefColWidth;
	long	m_nHeaderHeight;
	long	m_nRowHeight;
	long	m_nRows;
	BOOL	m_bGroupHeader;
	//}}AFX_DATA

// Implementation
protected:
	void UpdateControls();
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Message maps
protected:
	//{{AFX_MSG(CBHMDBGridPropPageGeneral)
	afx_msg void OnSelendokComboDatamode();
	afx_msg void OnCheckAllowupdate();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_BHMDBGRIDPPGGENERAL_H__BF91B13D_6A84_11D3_A7FE_0080C8763FA4__INCLUDED)
