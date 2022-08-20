#if !defined(AFX_BHMDBGROUPSPROPPAGE_H__05803602_6D0C_11D3_A7FE_0080C8763FA4__INCLUDED_)
#define AFX_BHMDBGROUPSPROPPAGE_H__05803602_6D0C_11D3_A7FE_0080C8763FA4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// BHMDBGroupsPropPage.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CBHMDBGroupsPropPage : Property page dialog
#include "GridInPpg.h"

class CBHMDBGridCtrl;
class CBHMDBDropDownCtrl;
class CBHMDBComboCtrl;

class CBHMDBGroupsPropPage : public COlePropertyPage
{
	DECLARE_DYNCREATE(CBHMDBGroupsPropPage)
	DECLARE_OLECREATE_EX(CBHMDBGroupsPropPage)

// Constructors
public:
	// Check button state in a radio button group.
	void CheckRadio(CButton * pButton, int nValue);

	CBHMDBGroupsPropPage();

// Dialog Data
	//{{AFX_DATA(CBHMDBGroupsPropPage)
	enum { IDD = IDD_PROPPAGE_BHMRICHGRID_GROUP };
		// NOTE - ClassWizard will add data members here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_DATA

// Implementation
protected:
	virtual void DoDataExchange(CDataExchange* pDX);        // DDX/DDV support

protected:
	// Pointers of underlying control.
	CBHMDBGridCtrl * m_pGridCtrl;
	CBHMDBDropDownCtrl * m_pDropDownCtrl;
	CBHMDBComboCtrl * m_pComboCtrl;

	CGridInPpg m_wndGrid;

// Message maps
protected:
	// Get the pointer of underlying control.
	void InitControlsVars();

	// Show/hide all controls in page.
	void HideControls(BOOL bHide = TRUE);

	// Init the grid.
	void InitGrid();
	//{{AFX_MSG(CBHMDBGroupsPropPage)
	afx_msg void OnButtonAddgroup();
	afx_msg void OnButtonBackcolor();
	afx_msg void OnButtonDeletegroup();
	afx_msg void OnButtonForecolor();
	afx_msg void OnButtonReset();
	afx_msg void OnCheckAllowsizing();
	afx_msg void OnCheckVisible();
	afx_msg void OnSelendokComboGroup();
	afx_msg void OnKillfocusEditTitle();
	afx_msg void OnKillfocusEditWidth();
	virtual BOOL OnInitDialog();
	afx_msg void OnKillfocusEditName();
	afx_msg void OnKillfocusEditLevels();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

	friend class CGridInPpg;
	friend class CAddColumnDlg;
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_BHMDBGROUPSPROPPAGE_H__05803602_6D0C_11D3_A7FE_0080C8763FA4__INCLUDED_)
