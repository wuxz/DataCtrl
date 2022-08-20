#if !defined(AFX_BHMDBCOLUMNPROPPAGE_H__BF91B166_6A84_11D3_A7FE_0080C8763FA4__INCLUDED_)
#define AFX_BHMDBCOLUMNPROPPAGE_H__BF91B166_6A84_11D3_A7FE_0080C8763FA4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// BHMDBColumnPropPage.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CBHMDBColumnPropPage : Property page dialog
#include "GridInPpg.h"

class CBHMDBGridCtrl;
class CBHMDBDropDownCtrl;
class CBHMDBComboCtrl;

class CBHMDBColumnPropPage : public COlePropertyPage
{
	DECLARE_DYNCREATE(CBHMDBColumnPropPage)
	DECLARE_OLECREATE_EX(CBHMDBColumnPropPage)

// Constructors
public:
	CBHMDBColumnPropPage();
	
	// If the field can not be null.
	BOOL IsFieldForceNotNullable(CString strField);

	// If the field is readonly.
	BOOL IsFieldForceLock(CString strField);

	// Calc the ordinal of the alignment ordinal in its group.
	int CalcAlignmentOrdinal(UINT align);

	// Insert one col into grid.
	int InsertCol();

	// Check the correct radio button in group.
	void CheckRadio(CButton * pButton, int nValue);


// Dialog Data
	//{{AFX_DATA(CBHMDBColumnPropPage)
	enum { IDD = IDD_PROPPAGE_BHMRICHGRID_COLUMN };
		// NOTE - ClassWizard will add data members here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_DATA

// Implementation
protected:
	virtual void DoDataExchange(CDataExchange* pDX);        // DDX/DDV support

protected:
	// The pointers to the underlying controls.
	CBHMDBGridCtrl * m_pGridCtrl;
	CBHMDBDropDownCtrl * m_pDropDownCtrl;
	CBHMDBComboCtrl * m_pComboCtrl;

	// The grid in page.
	CGridInPpg m_wndGrid;

// Message maps
protected:
	// Init the grid from underlying control.
	void InitGrid();

	// Calc the property ordinal in total properties.
	int CalcPropertyOrdinal(int nProperty);

	// Hide all of the controls in page.
	void HideControls();

	// Init the underlying control pointers.
	void InitControlsVars();

	//{{AFX_MSG(CBHMDBColumnPropPage)
	virtual BOOL OnInitDialog();
	afx_msg void OnSelchangeListProperty();
	afx_msg void OnSelendokComboColumn();
	afx_msg void OnButtonBackcolor();
	afx_msg void OnButtonForecolor();
	afx_msg void OnButtonHeadbackcolor();
	afx_msg void OnButtonHeadforecolor();
	afx_msg void OnCheckAllowsizing();
	afx_msg void OnCheckLocked();
	afx_msg void OnCheckNullable();
	afx_msg void OnCheckVisible();
	afx_msg void OnKillfocusEditName();
	afx_msg void OnKillfocusEditCaption();
	afx_msg void OnKillfocusEditFieldlen();
	afx_msg void OnKillfocusEditMask();
	afx_msg void OnKillfocusEditPromptchar();
	afx_msg void OnKillfocusEditWidth();
	afx_msg void OnRadioAlignmentCenter();
	afx_msg void OnRadioAlignmentLeft();
	afx_msg void OnRadioAlignmentRight();
	afx_msg void OnRadioCaptionalignmentAsalignment();
	afx_msg void OnRadioCaptionalignmentCenter();
	afx_msg void OnRadioCaptionalignmentLeft();
	afx_msg void OnRadioCaptionalignmentRight();
	afx_msg void OnRadioCaseLower();
	afx_msg void OnRadioCaseNone();
	afx_msg void OnRadioCaseUpper();
	afx_msg void OnRadioDatatypeBoolean();
	afx_msg void OnRadioDatatypeByte();
	afx_msg void OnRadioDatatypeCurrency();
	afx_msg void OnRadioDatatypeDate();
	afx_msg void OnRadioDatatypeInteger();
	afx_msg void OnRadioDatatypeLong();
	afx_msg void OnRadioDatatypeSingle();
	afx_msg void OnRadioDatatypeText();
	afx_msg void OnRadioStyleButton();
	afx_msg void OnRadioStyleCheckbox();
	afx_msg void OnRadioStyleCombobox();
	afx_msg void OnRadioStyleEdit();
	afx_msg void OnRadioStyleEditbutton();
	afx_msg void OnButtonSetupcombobox();
	afx_msg void OnButtonAddcolumn();
	afx_msg void OnButtonDeletecolumn();
	afx_msg void OnButtonReset();
	afx_msg void OnCheckPromptinclude();
	afx_msg void OnSelendokComboDatafield();
	afx_msg void OnButtonFields();
	afx_msg void OnSelendokComboDatemask();
	afx_msg void OnRadioStyleDatemask();
	afx_msg void OnRadioStyleCombolist();
	afx_msg void OnRadioStyleNumeric();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

	friend class CGridInPpg;
	friend class CAddColumnDlg;
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_BHMDBCOLUMNPROPPAGE_H__BF91B166_6A84_11D3_A7FE_0080C8763FA4__INCLUDED_)
