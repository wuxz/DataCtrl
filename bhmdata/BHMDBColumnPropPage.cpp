// BHMDBColumnPropPage.cpp : implementation file
//
/*

	This property page class is shared by all of the three controls. We use three pointer
	which point to the underlying controls to determine the actual control.

	The properties list box is shared also, but for some properties only exist in grid 
	control, there must be some extra work on it.

*/
#include "stdafx.h"
#include "bhmdata.h"
#include "BHMDBColumnPropPage.h"
#include "BHMDBGridCtl.h"
#include "GridInputDlg.h"
#include "AddColumnDlg.h"
#include "FieldsDlg.h"
#include "BHMDBDropDownCtl.h"
#include "BHMDBComboCtl.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

#define pComboColumn ((CComboBox *)GetDlgItem(IDC_COMBO_COLUMN))
#define pStaticFrame GetDlgItem(IDC_STATIC_FRAME)
#define pListProperty ((CListBox *)GetDlgItem(IDC_LIST_PROPERTY))
#define pFields GetDlgItem(IDC_BUTTON_FIELDS)
#define pDeleteColumn GetDlgItem(IDC_BUTTON_DELETECOLUMN)

#define	pAlignLeft ((CButton *)GetDlgItem(IDC_RADIO_ALIGNMENT_LEFT))
#define	pAlignRight ((CButton *)GetDlgItem(IDC_RADIO_ALIGNMENT_RIGHT))
#define	pAlignCenter ((CButton *)GetDlgItem(IDC_RADIO_ALIGNMENT_CENTER))
#define pAllowSizing ((CButton *)GetDlgItem(IDC_CHECK_ALLOWSIZING))

#define	pCaption GetDlgItem(IDC_EDIT_CAPTION)

#define pCapAlignLeft ((CButton *)GetDlgItem(IDC_RADIO_CAPTIONALIGNMENT_LEFT))
#define pCapAlignRight ((CButton *)GetDlgItem(IDC_RADIO_CAPTIONALIGNMENT_RIGHT))
#define pCapAlignCenter ((CButton *)GetDlgItem(IDC_RADIO_CAPTIONALIGNMENT_CENTER))
#define pCapAlignAsAlignment ((CButton *)GetDlgItem(IDC_RADIO_CAPTIONALIGNMENT_ASALIGNMENT))

#define pCaseNone ((CButton *)GetDlgItem(IDC_RADIO_CASE_NONE))
#define pCaseUpper ((CButton *)GetDlgItem(IDC_RADIO_CASE_UPPER))
#define pCaseLower ((CButton *)GetDlgItem(IDC_RADIO_CASE_LOWER))

#define pDataField ((CComboBox *)GetDlgItem(IDC_COMBO_DATAFIELD))

#define	pBackColor GetDlgItem(IDC_BUTTON_BACKCOLOR)
#define	pForeColor GetDlgItem(IDC_BUTTON_FORECOLOR)
#define pHeadBackColor GetDlgItem(IDC_BUTTON_HEADBACKCOLOR)
#define pHeadForeColor GetDlgItem(IDC_BUTTON_HEADFORECOLOR)

#define	pFieldLen GetDlgItem(IDC_EDIT_FIELDLEN)

#define pLocked ((CButton *)GetDlgItem(IDC_CHECK_LOCKED))

#define pText ((CButton *)GetDlgItem(IDC_RADIO_DATATYPE_TEXT))
#define pBoolean ((CButton *)GetDlgItem(IDC_RADIO_DATATYPE_BOOLEAN))
#define pByte ((CButton *)GetDlgItem(IDC_RADIO_DATATYPE_BYTE))
#define pInteger ((CButton *)GetDlgItem(IDC_RADIO_DATATYPE_INTEGER))
#define pLong ((CButton *)GetDlgItem(IDC_RADIO_DATATYPE_LONG))
#define pSingle ((CButton *)GetDlgItem(IDC_RADIO_DATATYPE_SINGLE))
#define pCurrency ((CButton *)GetDlgItem(IDC_RADIO_DATATYPE_CURRENCY))
#define pDate ((CButton *)GetDlgItem(IDC_RADIO_DATATYPE_DATE))
#define pDateMask GetDlgItem(IDC_COMBO_DATEMASK)
#define pStaticDateMask GetDlgItem(IDC_STATIC_DATEMASK)

#define	pMask GetDlgItem(IDC_EDIT_MASK)

#define pName GetDlgItem(IDC_EDIT_NAME)

#define pNullable ((CButton *)GetDlgItem(IDC_CHECK_NULLABLE))

#define pPromptChar GetDlgItem(IDC_EDIT_PROMPTCHAR)

#define pStyEdit ((CButton *)GetDlgItem(IDC_RADIO_STYLE_EDIT))
#define pStyEditButton ((CButton *)GetDlgItem(IDC_RADIO_STYLE_EDITBUTTON))
#define pStyCheckBox ((CButton *)GetDlgItem(IDC_RADIO_STYLE_CHECKBOX))
#define pStyComboBox ((CButton *)GetDlgItem(IDC_RADIO_STYLE_COMBOBOX))
#define pSetupComboBox (GetDlgItem(IDC_BUTTON_SETUPCOMBOBOX))
#define pStyButton ((CButton *)GetDlgItem(IDC_RADIO_STYLE_BUTTON))
#define pStyComboList ((CButton *)GetDlgItem(IDC_RADIO_STYLE_COMBOLIST))
#define pStyDateMask ((CButton *)GetDlgItem(IDC_RADIO_STYLE_DATEMASK))
#define pStyNumeric ((CButton *)GetDlgItem(IDC_RADIO_STYLE_NUMERIC))

#define pVisible ((CButton *)GetDlgItem(IDC_CHECK_VISIBLE))

#define pWidth GetDlgItem(IDC_EDIT_WIDTH)

#define pPromptInclude ((CButton *)GetDlgItem(IDC_CHECK_PROMPTINCLUDE))

/////////////////////////////////////////////////////////////////////////////
// CBHMDBColumnPropPage dialog

IMPLEMENT_DYNCREATE(CBHMDBColumnPropPage, COlePropertyPage)


/////////////////////////////////////////////////////////////////////////////
// Message map

BEGIN_MESSAGE_MAP(CBHMDBColumnPropPage, COlePropertyPage)
	//{{AFX_MSG_MAP(CBHMDBColumnPropPage)
	ON_LBN_SELCHANGE(IDC_LIST_PROPERTY, OnSelchangeListProperty)
	ON_CBN_SELENDOK(IDC_COMBO_COLUMN, OnSelendokComboColumn)
	ON_BN_CLICKED(IDC_BUTTON_BACKCOLOR, OnButtonBackcolor)
	ON_BN_CLICKED(IDC_BUTTON_FORECOLOR, OnButtonForecolor)
	ON_BN_CLICKED(IDC_BUTTON_HEADBACKCOLOR, OnButtonHeadbackcolor)
	ON_BN_CLICKED(IDC_BUTTON_HEADFORECOLOR, OnButtonHeadforecolor)
	ON_BN_CLICKED(IDC_CHECK_ALLOWSIZING, OnCheckAllowsizing)
	ON_BN_CLICKED(IDC_CHECK_LOCKED, OnCheckLocked)
	ON_BN_CLICKED(IDC_CHECK_NULLABLE, OnCheckNullable)
	ON_BN_CLICKED(IDC_CHECK_VISIBLE, OnCheckVisible)
	ON_EN_KILLFOCUS(IDC_EDIT_NAME, OnKillfocusEditName)
	ON_EN_KILLFOCUS(IDC_EDIT_CAPTION, OnKillfocusEditCaption)
	ON_EN_KILLFOCUS(IDC_EDIT_FIELDLEN, OnKillfocusEditFieldlen)
	ON_EN_KILLFOCUS(IDC_EDIT_MASK, OnKillfocusEditMask)
	ON_EN_KILLFOCUS(IDC_EDIT_PROMPTCHAR, OnKillfocusEditPromptchar)
	ON_EN_KILLFOCUS(IDC_EDIT_WIDTH, OnKillfocusEditWidth)
	ON_BN_CLICKED(IDC_RADIO_ALIGNMENT_CENTER, OnRadioAlignmentCenter)
	ON_BN_CLICKED(IDC_RADIO_ALIGNMENT_LEFT, OnRadioAlignmentLeft)
	ON_BN_CLICKED(IDC_RADIO_ALIGNMENT_RIGHT, OnRadioAlignmentRight)
	ON_BN_CLICKED(IDC_RADIO_CAPTIONALIGNMENT_ASALIGNMENT, OnRadioCaptionalignmentAsalignment)
	ON_BN_CLICKED(IDC_RADIO_CAPTIONALIGNMENT_CENTER, OnRadioCaptionalignmentCenter)
	ON_BN_CLICKED(IDC_RADIO_CAPTIONALIGNMENT_LEFT, OnRadioCaptionalignmentLeft)
	ON_BN_CLICKED(IDC_RADIO_CAPTIONALIGNMENT_RIGHT, OnRadioCaptionalignmentRight)
	ON_BN_CLICKED(IDC_RADIO_CASE_LOWER, OnRadioCaseLower)
	ON_BN_CLICKED(IDC_RADIO_CASE_NONE, OnRadioCaseNone)
	ON_BN_CLICKED(IDC_RADIO_CASE_UPPER, OnRadioCaseUpper)
	ON_BN_CLICKED(IDC_RADIO_DATATYPE_BOOLEAN, OnRadioDatatypeBoolean)
	ON_BN_CLICKED(IDC_RADIO_DATATYPE_BYTE, OnRadioDatatypeByte)
	ON_BN_CLICKED(IDC_RADIO_DATATYPE_CURRENCY, OnRadioDatatypeCurrency)
	ON_BN_CLICKED(IDC_RADIO_DATATYPE_DATE, OnRadioDatatypeDate)
	ON_BN_CLICKED(IDC_RADIO_DATATYPE_INTEGER, OnRadioDatatypeInteger)
	ON_BN_CLICKED(IDC_RADIO_DATATYPE_LONG, OnRadioDatatypeLong)
	ON_BN_CLICKED(IDC_RADIO_DATATYPE_SINGLE, OnRadioDatatypeSingle)
	ON_BN_CLICKED(IDC_RADIO_DATATYPE_TEXT, OnRadioDatatypeText)
	ON_BN_CLICKED(IDC_RADIO_STYLE_BUTTON, OnRadioStyleButton)
	ON_BN_CLICKED(IDC_RADIO_STYLE_CHECKBOX, OnRadioStyleCheckbox)
	ON_BN_CLICKED(IDC_RADIO_STYLE_COMBOBOX, OnRadioStyleCombobox)
	ON_BN_CLICKED(IDC_RADIO_STYLE_EDIT, OnRadioStyleEdit)
	ON_BN_CLICKED(IDC_RADIO_STYLE_EDITBUTTON, OnRadioStyleEditbutton)
	ON_BN_CLICKED(IDC_BUTTON_SETUPCOMBOBOX, OnButtonSetupcombobox)
	ON_BN_CLICKED(IDC_BUTTON_ADDCOLUMN, OnButtonAddcolumn)
	ON_BN_CLICKED(IDC_BUTTON_DELETECOLUMN, OnButtonDeletecolumn)
	ON_BN_CLICKED(IDC_BUTTON_RESET, OnButtonReset)
	ON_BN_CLICKED(IDC_CHECK_PROMPTINCLUDE, OnCheckPromptinclude)
	ON_CBN_SELENDOK(IDC_COMBO_DATAFIELD, OnSelendokComboDatafield)
	ON_BN_CLICKED(IDC_BUTTON_FIELDS, OnButtonFields)
	ON_CBN_SELENDOK(IDC_COMBO_DATEMASK, OnSelendokComboDatemask)
	ON_BN_CLICKED(IDC_RADIO_STYLE_DATEMASK, OnRadioStyleDatemask)
	ON_BN_CLICKED(IDC_RADIO_STYLE_COMBOLIST, OnRadioStyleCombolist)
	ON_BN_CLICKED(IDC_RADIO_STYLE_NUMERIC, OnRadioStyleNumeric)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// Initialize class factory and guid

// {BF91B165-6A84-11D3-A7FE-0080C8763FA4}
IMPLEMENT_OLECREATE_EX(CBHMDBColumnPropPage, "BHMData.CBHMDBColumnPropPage",
	0xbf91b165, 0x6a84, 0x11d3, 0xa7, 0xfe, 0x0, 0x80, 0xc8, 0x76, 0x3f, 0xa4)


/////////////////////////////////////////////////////////////////////////////
// CBHMDBColumnPropPage::CBHMDBColumnPropPageFactory::UpdateRegistry -
// Adds or removes system registry entries for CBHMDBColumnPropPage

BOOL CBHMDBColumnPropPage::CBHMDBColumnPropPageFactory::UpdateRegistry(BOOL bRegister)
{
	// TODO: Define string resource for page type; replace '0' below with ID.

	if (bRegister)
		return AfxOleRegisterPropertyPageClass(AfxGetInstanceHandle(),
			m_clsid, IDS_BHMDB_COLUMNS_PPG);
	else
		return AfxOleUnregisterClass(m_clsid, NULL);
}


/////////////////////////////////////////////////////////////////////////////
// CBHMDBColumnPropPage::CBHMDBColumnPropPage - Constructor

// TODO: Define string resource for page caption; replace '0' below with ID.

CBHMDBColumnPropPage::CBHMDBColumnPropPage() :
	COlePropertyPage(IDD, IDS_BHMDB_COLUMNS_PPG_CAPTION)
{
	//{{AFX_DATA_INIT(CBHMDBColumnPropPage)
	// NOTE: ClassWizard will add member initialization here
	//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_DATA_INIT
	m_pGridCtrl = NULL;
	m_pDropDownCtrl = NULL;
	m_pComboCtrl = NULL;

	// Set the callback pointer.
	m_wndGrid.SetPagePtr(this);
}

/////////////////////////////////////////////////////////////////////////////
// CBHMDBColumnPropPage::DoDataExchange - Moves data between page and properties

void CBHMDBColumnPropPage::DoDataExchange(CDataExchange* pDX)
{
	// NOTE: ClassWizard will add DDP, DDX, and DDV calls here
	//    DO NOT EDIT what you see in these blocks of generated code !
	//{{AFX_DATA_MAP(CBHMDBColumnPropPage)
	//}}AFX_DATA_MAP
	DDP_PostProcessing(pDX);

	if (!pDX->m_bSaveAndValidate)
	{
		// If is loading, init the grid.
		InitGrid();
	}
	else
	{
		// Calc the index of current itex in the properties list box.
		int nProperty = CalcPropertyOrdinal(pListProperty->GetCurSel());
		
		// Update the current property when saving.
		if (nProperty == 3)
			OnKillfocusEditCaption();
		else if(nProperty == 8)
			OnKillfocusEditFieldlen();
		else if(nProperty == 13)
			OnKillfocusEditMask();
		else if(nProperty == 14)
			OnKillfocusEditName();
		else if(nProperty == 16)
			OnKillfocusEditPromptchar();
		else if(nProperty == 19)
			OnKillfocusEditWidth();

		if (m_pGridCtrl)
		{
			// The control underlying is grid control.

			// Save the layout.
			m_wndGrid.SaveLayout(&m_pGridCtrl->m_arGroups, &m_pGridCtrl->m_arCols, &m_pGridCtrl->m_arCells);

			// Init the control.
			m_pGridCtrl->InitGridFromProp();

			// Update the other layout data not saved in prior function.
			m_pGridCtrl->SetRowHeight(m_wndGrid.GetRowHeight(1));
			m_pGridCtrl->SetHeaderHeight(m_wndGrid.GetRowHeight(0));

			// If it has layout data.
			m_pGridCtrl->m_bHasLayout = m_wndGrid.GetGroupCount() || m_wndGrid.GetColCount();
		}
		else if (m_pDropDownCtrl)
		{
			// The control underlying is dropdown control.

			// Save the layout.
			m_wndGrid.SaveLayout(&m_pDropDownCtrl->m_arGroups, &m_pDropDownCtrl->m_arCols, &m_pDropDownCtrl->m_arCells);

			// Init the control.
			m_pDropDownCtrl->InitGridFromProp();

			// Save the other layout data not saved in prior function.
			m_pDropDownCtrl->SetRowHeight(m_wndGrid.GetRowHeight(1));
			m_pDropDownCtrl->SetHeaderHeight(m_wndGrid.GetRowHeight(0));

			// If it has layout data.
			m_pDropDownCtrl->m_bHasLayout = m_wndGrid.GetGroupCount() || m_wndGrid.GetColCount();
		}
		else if (m_pComboCtrl)
		{
			// The control underlying is combo control.

			// Save the layout.
			m_wndGrid.SaveLayout(&m_pComboCtrl->m_arGroups, &m_pComboCtrl->m_arCols, &m_pComboCtrl->m_arCells);

			// Init the control.
			m_pComboCtrl->InitGridFromProp();

			// Save the unsaved layout data in prior function.			
			m_pComboCtrl->SetRowHeight(m_wndGrid.GetRowHeight(1));
			m_pComboCtrl->SetHeaderHeight(m_wndGrid.GetRowHeight(0));

			// If it has layout data.
			m_pComboCtrl->m_bHasLayout = m_wndGrid.GetGroupCount() || m_wndGrid.GetColCount();
		}
	}
}


/////////////////////////////////////////////////////////////////////////////
// CBHMDBColumnPropPage message handlers
BOOL CBHMDBColumnPropPage::OnInitDialog() 
{
	COlePropertyPage::OnInitDialog();
	
	// Init the controls in page.
	InitControlsVars();

	// Hide all controls default.
	HideControls();

	// If can not obtain the pointer of the underlying control, we can
	// not do anything.
	if (!m_pGridCtrl && !m_pDropDownCtrl && !m_pComboCtrl)
		return FALSE;

	m_wndGrid.ResetScrollBars();
	m_wndGrid.Invalidate();

	return FALSE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

void CBHMDBColumnPropPage::InitControlsVars()
/*
Routine Description:
	Init some variables and control before working.

Parameters:
	None.

Return Value:
	None.
*/
{
	// Subclass the grid control.
	m_wndGrid.SubclassDlgItem(IDC_GRID, this);
	
	// Get the dispatch pointer array.
	ULONG Ulong;
	LPDISPATCH FAR *m_lpDispatch = GetObjectArray(&Ulong);

	// Get the CCmdTarget object associated to any one of the above
	// array elements
	if (Ulong)
	{
		DISPPARAMS tParam;
		tParam.rgdispidNamedArgs = NULL;
		tParam.cArgs = 0;
		tParam.cNamedArgs = 0;
		tParam.rgvarg = 0;
		VARIANT va;
		VariantInit(&va);
		
		// Invoke my private method to get the type of the control.
		HRESULT hr = m_lpDispatch[0]->Invoke(255, IID_NULL, LOCALE_USER_DEFAULT,
			DISPATCH_METHOD | DISPATCH_PROPERTYGET, &tParam, &va, NULL, NULL);
		if (SUCCEEDED(hr))
		{
			if (va.iVal == 0)
				// It is the grid control.
				m_pGridCtrl = (CBHMDBGridCtrl *) CCmdTarget::FromIDispatch(m_lpDispatch[0]);
			else if (va.iVal == 1)
				// It is the dropdown control.
				m_pDropDownCtrl = (CBHMDBDropDownCtrl *) CCmdTarget::FromIDispatch(m_lpDispatch[0]);
			else if (va.iVal == 2)
				// It is the combo control.
				m_pComboCtrl = (CBHMDBComboCtrl *) CCmdTarget::FromIDispatch(m_lpDispatch[0]);
		}

		VariantClear(&va);
	}

	// Add the string into list box.
	// Note that some properties only exists in grid control.

	pListProperty->AddString("Alignment");
	if (m_pGridCtrl)
		pListProperty->AddString("AllowSizing");
	pListProperty->AddString("BackColor");
	pListProperty->AddString("Caption");
	pListProperty->AddString("CaptionAlignment");
	pListProperty->AddString("Case");
	pListProperty->AddString("DataField");
	if (m_pGridCtrl)
		pListProperty->AddString("DataType");
	if (m_pGridCtrl)
		pListProperty->AddString("FieldLen");
	pListProperty->AddString("ForeColor");
	pListProperty->AddString("HeadBackColor");
	pListProperty->AddString("HeadForeColor");
	if (m_pGridCtrl)
		pListProperty->AddString("Locked");
	if (m_pGridCtrl)
		pListProperty->AddString("Mask");
	pListProperty->AddString("Name");
	if (m_pGridCtrl)
		pListProperty->AddString("Nullable");
	if (m_pGridCtrl)
		pListProperty->AddString("PromptChar");
	if (m_pGridCtrl)
		pListProperty->AddString("Style");
	pListProperty->AddString("Visible");
	pListProperty->AddString("Width");
	if (m_pGridCtrl)
		pListProperty->AddString("PromptInclude");

	// If can not known the underlying control, we can't do anything.

	if (!m_pGridCtrl && !m_pDropDownCtrl && !m_pComboCtrl)
	{
		GetDlgItem(IDC_BUTTON_ADDCOLUMN)->EnableWindow(FALSE);
		pDeleteColumn->EnableWindow(FALSE);
		GetDlgItem(IDC_BUTTON_RESET)->EnableWindow(FALSE);
		pFields->EnableWindow(FALSE);
		pComboColumn->EnableWindow(FALSE);
		pListProperty->EnableWindow(FALSE);
		m_wndGrid.EnableWindow(FALSE);
		return;
	}
}

void CBHMDBColumnPropPage::HideControls()
/*
Routine Description:
	Hide all of the controls in page.

Parameters:
	None.

Return Value:
	None.
*/
{
	pAlignLeft->ShowWindow(SW_HIDE);
	pAlignRight->ShowWindow(SW_HIDE);
	pAlignCenter->ShowWindow(SW_HIDE);

	pAllowSizing->ShowWindow(SW_HIDE);

	pCaption->ShowWindow(SW_HIDE);

	pCapAlignLeft->ShowWindow(SW_HIDE);
	pCapAlignRight->ShowWindow(SW_HIDE);
	pCapAlignCenter->ShowWindow(SW_HIDE);
	pCapAlignAsAlignment->ShowWindow(SW_HIDE);

	pCaseNone->ShowWindow(SW_HIDE);
	pCaseUpper->ShowWindow(SW_HIDE);
	pCaseLower->ShowWindow(SW_HIDE);

	pBackColor->ShowWindow(SW_HIDE);
	pForeColor->ShowWindow(SW_HIDE);
	pHeadBackColor->ShowWindow(SW_HIDE);
	pHeadForeColor->ShowWindow(SW_HIDE);

	pDataField->ShowWindow(SW_HIDE);
	
	pFieldLen->ShowWindow(SW_HIDE);

	pLocked->ShowWindow(SW_HIDE);

	pText->ShowWindow(SW_HIDE);
	pBoolean->ShowWindow(SW_HIDE);
	pByte->ShowWindow(SW_HIDE);
	pInteger->ShowWindow(SW_HIDE);
	pLong->ShowWindow(SW_HIDE);
	pSingle->ShowWindow(SW_HIDE);
	pCurrency->ShowWindow(SW_HIDE);
	pDate->ShowWindow(SW_HIDE);
	pDateMask->ShowWindow(SW_HIDE);
	pStaticDateMask->ShowWindow(SW_HIDE);

	pMask->ShowWindow(SW_HIDE);

	pName->ShowWindow(SW_HIDE);

	pNullable->ShowWindow(SW_HIDE);

	pPromptChar->ShowWindow(SW_HIDE);

	pStyEdit->ShowWindow(SW_HIDE);
	pStyEditButton->ShowWindow(SW_HIDE);
	pStyCheckBox->ShowWindow(SW_HIDE);
	pStyComboBox->ShowWindow(SW_HIDE);
	pStyButton->ShowWindow(SW_HIDE);
	pSetupComboBox->ShowWindow(SW_HIDE);
	pStyComboList->ShowWindow(SW_HIDE);
	pStyDateMask->ShowWindow(SW_HIDE);
	pStyNumeric->ShowWindow(SW_HIDE);

	pVisible->ShowWindow(SW_HIDE);

	pWidth->ShowWindow(SW_HIDE);

	pPromptInclude->ShowWindow(SW_HIDE);

	pStaticFrame->SetWindowText("");
}

void CBHMDBColumnPropPage::OnSelchangeListProperty() 
{
	// The user select one property.

	// Calc the ordinal of current item.
	int nProperty = CalcPropertyOrdinal(pListProperty->GetCurSel());
	
	// Hide all controls first.
	HideControls();

	// The fields control is only valid when the datamode is 0;
	if (m_pGridCtrl)
		pFields->EnableWindow(m_pGridCtrl->m_nDataMode == 0);
	else if (m_pDropDownCtrl)
		pFields->EnableWindow(m_pDropDownCtrl->m_nDataMode == 0);
	else if (m_pComboCtrl)
		pFields->EnableWindow(m_pComboCtrl->m_nDataMode == 0);

	// Calc the current col number in grid.
	CString strCol;
	pComboColumn->GetWindowText(strCol);
	int nCurrentCol = m_wndGrid.GetColFromName(strCol);
	if (nCurrentCol == INVALID)
		return;

	// Clone the properties of current col.
	Col col;
	m_wndGrid.CloneColProp(nCurrentCol, &col);

	CString strText;

	switch (nProperty)
	{
	case 0:
		// Current property is "Alignment".

		// Enable some controls.
		pAlignLeft->ShowWindow(SW_SHOW);
		pAlignRight->ShowWindow(SW_SHOW);
		pAlignCenter->ShowWindow(SW_SHOW);
		pStaticFrame->SetWindowText("Alignment");

		CheckRadio(pAlignLeft, col.nAlignment == DT_LEFT ? 0 : col.nAlignment == DT_RIGHT ? 1 : 2);
		break;

	case 1:
		// Current property is "AllowSizing".

		// Enable some controls.
		pAllowSizing->ShowWindow(SW_SHOW);
		pStaticFrame->SetWindowText("AllowSizing");

		pAllowSizing->SetCheck(col.bAllowSizing ? 1 : 0);
		break;

	case 2:
		// Current property is "BackColor".

		// Enable some controls.
		pBackColor->ShowWindow(SW_SHOW);
		pStaticFrame->SetWindowText("BackColor");
		break;

	case 3:
		// Current property is "Caption".

		// Enable some controls.
		pCaption->ShowWindow(SW_SHOW);
		pStaticFrame->SetWindowText("Caption");

		pCaption->SetWindowText(col.strTitle);
		break;

	case 4:
		// Current property is "CaptionAlignment".

		// Enable some controls.
		pCapAlignLeft->ShowWindow(SW_SHOW);
		pCapAlignRight->ShowWindow(SW_SHOW);
		pCapAlignCenter->ShowWindow(SW_SHOW);
		pCapAlignAsAlignment->ShowWindow(SW_SHOW);
		pStaticFrame->SetWindowText("CaptionAlignment");

		CheckRadio(pCapAlignLeft, CalcAlignmentOrdinal(col.nHeaderAlignment));
		break;
	
	case 5:
		// Current property is "Case".

		// Enable some controls.
		pCaseNone->ShowWindow(SW_SHOW);
		pCaseUpper->ShowWindow(SW_SHOW);
		pCaseLower->ShowWindow(SW_SHOW);
		pStaticFrame->SetWindowText("Case");

		CheckRadio(pCaseNone, col.nCase);
		break;

	case 6:
		// Current property is "DataField".

		// Enable some controls.
		pDataField->ShowWindow(SW_SHOW);
		pStaticFrame->SetWindowText("DataField");
		
		if (m_pGridCtrl)
			pDataField->EnableWindow(m_pGridCtrl->m_nDataMode == 0);
		else if (m_pDropDownCtrl)
			pDataField->EnableWindow(m_pDropDownCtrl->m_nDataMode == 0);
		else if (m_pComboCtrl)
			pDataField->EnableWindow(m_pComboCtrl->m_nDataMode == 0);

		pDataField->SelectString(-1, col.arUserAttrib[0].IsEmpty() ? 
			_T(" ") : col.arUserAttrib[0]);
		break;

	case 7:
		// Current property is "DataType".

		// Enable some controls.
		pText->ShowWindow(SW_SHOW);
		pBoolean->ShowWindow(SW_SHOW);
		pByte->ShowWindow(SW_SHOW);
		pInteger->ShowWindow(SW_SHOW);
		pLong->ShowWindow(SW_SHOW);
		pSingle->ShowWindow(SW_SHOW);
		pCurrency->ShowWindow(SW_SHOW);
		pDate->ShowWindow(SW_SHOW);
		pStaticFrame->SetWindowText("DataType");

		CheckRadio(pText, col.nDataType);
		break;

	case 8:
		// Current property is "FieldLen".

		// Enable some controls.
		pFieldLen->ShowWindow(SW_SHOW);
		pStaticFrame->SetWindowText("FieldLen");

		strText.Format("%d", col.nMaxLength);
		pFieldLen->SetWindowText(strText);

		if (m_pGridCtrl)
			pFieldLen->EnableWindow(m_pGridCtrl->m_nDataMode != 0 || col.arUserAttrib[0].GetLength() == 0);
		break;

	case 9:
		// Current property is "ForeColor".

		// Enable some controls.
		pForeColor->ShowWindow(SW_SHOW);
		pStaticFrame->SetWindowText("ForeColor");
		break;

	case 10:
		// Current property is "HeadBackColor".

		// Enable some controls.
		pHeadBackColor->ShowWindow(SW_SHOW);
		pStaticFrame->SetWindowText("HeadBackColor");
		break;

	case 11:
		// Current property is "HeadForeColor".

		// Enable some controls.
		pHeadForeColor->ShowWindow(SW_SHOW);
		pStaticFrame->SetWindowText("HeadForeColor");
		break;

	case 12:
		// Current property is "Locked".

		// Enable some controls.
		pLocked->ShowWindow(SW_SHOW);
		pStaticFrame->SetWindowText("Locked");

		pLocked->SetCheck(col.bReadOnly ? 1 : 0);
		if (m_pGridCtrl)
			pLocked->EnableWindow(m_pGridCtrl->m_nDataMode != 0 || !IsFieldForceLock(col.arUserAttrib[0]));
		break;

	case 13:
		// Current property is "Mask".

		// Enable some controls.
		pMask->ShowWindow(SW_SHOW);
		pStaticFrame->SetWindowText("Mask");

		pMask->SetWindowText(col.strMask);
		break;

	case 14:
		// Current property is "Name".

		// Enable some controls.
		pName->ShowWindow(SW_SHOW);
		pStaticFrame->SetWindowText("Name");

		pName->SetWindowText(col.strName);
		break;

	case 15:
		// Current property is "Nullable".

		// Enable some controls.
		pNullable->ShowWindow(SW_SHOW);
		pStaticFrame->SetWindowText("Nullable");

		pNullable->SetCheck(col.arUserAttrib[1].Compare(_T("T")) ? 0 : 1);
		if (m_pGridCtrl)
			pNullable->EnableWindow(m_pGridCtrl->m_nDataMode != 0 || !IsFieldForceNotNullable(col.arUserAttrib[0]));
		break;

	case 16:
		// Current property is "PromptChar".

		// Enable some controls.
		pPromptChar->ShowWindow(SW_SHOW);
		pStaticFrame->SetWindowText("PromptChar");

		pPromptChar->SetWindowText(col.strPromptChar);
		break;

	case 17:
		// Current property is "Style".

		// Enable some controls.
		pStyEdit->ShowWindow(SW_SHOW);
		pStyEditButton->ShowWindow(SW_SHOW);
		pStyComboBox->ShowWindow(SW_SHOW);
		pStyCheckBox->ShowWindow(SW_SHOW);
		pStyButton->ShowWindow(SW_SHOW);
		pStyComboList->ShowWindow(SW_SHOW);
		pStyDateMask->ShowWindow(SW_SHOW);
		pStyNumeric->ShowWindow(SW_SHOW);
		pStaticFrame->SetWindowText("Style");

		CheckRadio(pStyEdit, col.nStyle);
		pSetupComboBox->ShowWindow(pStyComboBox->GetCheck() || 
			pStyComboList->GetCheck() ? SW_SHOW : SW_HIDE);
		pDateMask->ShowWindow(pStyDateMask->GetCheck() ? SW_SHOW : SW_HIDE);
		pStaticDateMask->ShowWindow(pStyDateMask->GetCheck() ? SW_SHOW : SW_HIDE);
		break;

	case 18:
		// Current property is "Visible".

		// Enable some controls.
		pVisible->ShowWindow(SW_SHOW);
		pStaticFrame->SetWindowText("Visible");

		pVisible->SetCheck(col.bVisible ? 1 : 0);
		break;

	case 19:
		// Current property is "Width".

		// Enable some controls.
		pWidth->ShowWindow(SW_SHOW);
		pStaticFrame->SetWindowText("Width");

		strText.Format("%d", col.nWidth);
		pWidth->SetWindowText(strText);
		break;

	case 20:
		// Current property is "PromptInclude".

		// Enable some controls.
		pPromptInclude->ShowWindow(SW_SHOW);
		pStaticFrame->SetWindowText("PromptInclude");

		pPromptInclude->SetCheck(col.bPromptInclude ? 1 : 0);
		break;
	}
}

void CBHMDBColumnPropPage::OnSelendokComboColumn() 
{
	// The user select on column.

	// Update controls on page.
	OnSelchangeListProperty();

	// Get the current name.
	CString strName;
	pComboColumn->GetWindowText(strName);

	// Get the correspond col number in grid.
	int nCol = m_wndGrid.GetColFromName(strName);
	if (nCol == INVALID)
	{
		pComboColumn->SetCurSel(-1);
		return;
	}

	int nRow, nColCurr;

	// Set current cell to new col.
	m_wndGrid.GetCurrentCell(nRow, nColCurr);
	m_wndGrid.SetCurrentCell(nRow, nCol);

	// Can delete column after selecing one col.
	pDeleteColumn->EnableWindow(TRUE);
}

void CBHMDBColumnPropPage::CheckRadio(CButton *pButton, int nValue)
/*
Routine Description:
	Check the radio button in one groups according to the give value.

Parameters:
	pButton		The pointer of the first button in group.
	
	nValue		The value of this radio button group.

Return Value:
	None.
*/
{
	CWnd * pCtrl = pButton;

	// It should be the first radio button in a group.
	ASSERT(::GetWindowLong(pCtrl->m_hWnd, GWL_STYLE) & WS_GROUP);
	ASSERT(::SendMessage(pCtrl->m_hWnd, WM_GETDLGCODE, 0, 0L) & DLGC_RADIOBUTTON);

	// walk all children in group
	int iButton = 0;
	do
	{
		if (::SendMessage(pCtrl->m_hWnd, WM_GETDLGCODE, 0, 0L) & DLGC_RADIOBUTTON)
		{
			// control in group is a radio button
			// select button
			::SendMessage(pCtrl->m_hWnd, BM_SETCHECK, (iButton == nValue), 0L);
			iButton++;
		}
		else
		{
			TRACE0("Warning: skipping non-radio button in group.\n");
		}
		pCtrl = pCtrl->GetWindow(GW_HWNDNEXT);

	} while (pCtrl != NULL && !(GetWindowLong(pCtrl->m_hWnd, GWL_STYLE) 
		& WS_GROUP));
}

void CBHMDBColumnPropPage::OnButtonBackcolor() 
{
	// The user want to set the "BackColor" property.

	// Calc the current col number.
	CString strName;
	pComboColumn->GetWindowText(strName);
	int nCol = m_wndGrid.GetColFromName(strName);
	if (nCol == INVALID)
	{
		AfxMessageBox(IDS_ERROR_NOCURRENTCOLINCOLUMNPPG, MB_ICONSTOP);
		return;
	}

	// Copy the column properties out.
	Col col;
	m_wndGrid.CloneColProp(nCol, &col);
	COLORREF clrBack;
	
	clrBack = (col.clrBack == DEFAULTCOLOR ? m_wndGrid.GetBackColor() : col.clrBack);

	// Update the grid.
	CColorDialog clrDlg(clrBack, CC_FULLOPEN | CC_RGBINIT | CC_SOLIDCOLOR);
	if (clrDlg.DoModal() == IDOK)
		m_wndGrid.SetColBackColor(nCol, clrDlg.GetColor());
}

void CBHMDBColumnPropPage::OnButtonForecolor() 
{
	// The user want to set the "ForeColor" property.

	// Calc the current col number.
	CString strName;
	pComboColumn->GetWindowText(strName);
	int nCol = m_wndGrid.GetColFromName(strName);
	if (nCol == INVALID)
	{
		AfxMessageBox(IDS_ERROR_NOCURRENTCOLINCOLUMNPPG, MB_ICONSTOP);
		return;
	}

	// Copy the column property out.
	Col col;
	m_wndGrid.CloneColProp(nCol, &col);
	COLORREF clrFore;
	
	clrFore = col.clrFore == DEFAULTCOLOR ? m_wndGrid.GetForeColor() : col.clrFore;

	// Update the grid.
	CColorDialog clrDlg(clrFore, CC_FULLOPEN | CC_RGBINIT | CC_SOLIDCOLOR);
	if (clrDlg.DoModal() == IDOK)
		m_wndGrid.SetColForeColor(nCol, clrDlg.GetColor());
}

void CBHMDBColumnPropPage::OnButtonHeadbackcolor() 
{
	// The user want to set the "HeadBackColor" property.

	// Calc the current col number.
	CString strName;
	pComboColumn->GetWindowText(strName);
	int nCol = m_wndGrid.GetColFromName(strName);
	if (nCol == INVALID)
	{
		AfxMessageBox(IDS_ERROR_NOCURRENTCOLINCOLUMNPPG, MB_ICONSTOP);
		return;
	}

	// Copy the column properties out.
	Col col;
	m_wndGrid.CloneColProp(nCol, &col);
	COLORREF clrHeaderBack;
	
	clrHeaderBack = col.clrHeaderBack == DEFAULTCOLOR ? m_wndGrid.GetHeaderBackColor() : col.clrHeaderBack;

	// Update the grid.
	CColorDialog clrDlg(clrHeaderBack, CC_FULLOPEN | CC_RGBINIT | CC_SOLIDCOLOR);
	if (clrDlg.DoModal() == IDOK)
		m_wndGrid.SetColHeaderBackColor(nCol, clrDlg.GetColor());
}

void CBHMDBColumnPropPage::OnButtonHeadforecolor() 
{
	// The user want to set the "HeadForeColor" property.

	// Calc the current col number.
	CString strName;
	pComboColumn->GetWindowText(strName);
	int nCol = m_wndGrid.GetColFromName(strName);
	if (nCol == INVALID)
	{
		AfxMessageBox(IDS_ERROR_NOCURRENTCOLINCOLUMNPPG, MB_ICONSTOP);
		return;
	}

	// Copy the column properties out.
	Col col;
	m_wndGrid.CloneColProp(nCol, &col);
	COLORREF clrHeaderFore;
	
	clrHeaderFore = col.clrHeaderFore == DEFAULTCOLOR ? m_wndGrid.GetHeaderForeColor() : col.clrHeaderFore;

	// Update the grid.
	CColorDialog clrDlg(clrHeaderFore, CC_FULLOPEN | CC_RGBINIT | CC_SOLIDCOLOR);
	if (clrDlg.DoModal() == IDOK)
		m_wndGrid.SetColHeaderForeColor(nCol, clrDlg.GetColor());
}

void CBHMDBColumnPropPage::OnCheckAllowsizing() 
{
	// The user modified the "AllowSizing" property.

	// Calc the current col number.
	CString strName;
	pComboColumn->GetWindowText(strName);
	int nCol = m_wndGrid.GetColFromName(strName);
	if (nCol == INVALID)
		return;

	// Update the grid.
	m_wndGrid.SetColAllowSizing(nCol, pAllowSizing->GetCheck());
}

void CBHMDBColumnPropPage::OnCheckLocked() 
{
	// The user modified the "Locked" property.

	// Calc the current col number.
	CString strName;
	pComboColumn->GetWindowText(strName);
	int nCol = m_wndGrid.GetColFromName(strName);
	if (nCol == INVALID)
		return;

	// Update the grid.
	m_wndGrid.SetColReadOnly(nCol, pLocked->GetCheck());
}

void CBHMDBColumnPropPage::OnCheckNullable() 
{
	// The user modified the "Nullable" property.

	// Calc the current col number.
	CString strName;
	pComboColumn->GetWindowText(strName);
	int nCol = m_wndGrid.GetColFromName(strName);
	if (nCol == INVALID)
		return;

	// Update the grid.
	m_wndGrid.GetColUserAttribRef(nCol).SetAt(1, pNullable->GetCheck() ? _T("T") : _T(""));
}

void CBHMDBColumnPropPage::OnCheckVisible() 
{
	// The user modified the "Visible" property.

	// Calc the current col number.
	CString strName;
	pComboColumn->GetWindowText(strName);
	int nCol = m_wndGrid.GetColFromName(strName);
	if (nCol == INVALID)
		return;

	// Update the grid.
	m_wndGrid.SetColVisible(nCol, pVisible->GetCheck());
}

void CBHMDBColumnPropPage::OnKillfocusEditName() 
{
	// The user modified the "Name" property.

	// Calc the current col number.
	CString strName;
	pComboColumn->GetWindowText(strName);
	int nCol = m_wndGrid.GetColFromName(strName);
	if (nCol == INVALID)
		return;

	// Update the grid.
	pName->GetWindowText(strName);
	if (!m_wndGrid.SetColName(nCol, strName))
		return;

	// Update the column combo box.
	int nCurrentCol = pComboColumn->GetCurSel();
	pComboColumn->InsertString(nCurrentCol, strName);
	pComboColumn->DeleteString(nCurrentCol + 1);
	pComboColumn->SelectString(-1, strName);

	// Update other controls.
	OnSelendokComboColumn();
}

void CBHMDBColumnPropPage::OnKillfocusEditCaption() 
{
	// The user modified the "Caption" property.

	// Calc the current col number.
	CString strName;
	pComboColumn->GetWindowText(strName);
	int nCol = m_wndGrid.GetColFromName(strName);
	if (nCol == INVALID)
		return;

	// Update the grid.
	CString strTitle;
	pCaption->GetWindowText(strTitle);
	m_wndGrid.SetColTitle(nCol, strTitle);
}

void CBHMDBColumnPropPage::OnKillfocusEditFieldlen() 
{
	// The user modified the "FieldLen" property.

	// Calc the current col number.
	CString strName;
	pComboColumn->GetWindowText(strName);
	int nCol = m_wndGrid.GetColFromName(strName);
	if (nCol == INVALID)
		return;

	// Update the grid.
	CString strFieldLen;
	pFieldLen->GetWindowText(strFieldLen);

	// The FieldLen has limitation.
	int nFieldLen = atoi(strFieldLen);
	if (nFieldLen <= 0 || nFieldLen > 255)
		return;

	m_wndGrid.SetColMaxLength(nCol, nFieldLen);
}

void CBHMDBColumnPropPage::OnKillfocusEditMask() 
{
	// The user modified the "Mask" property.

	// Calc the current col number.
	CString strName;
	pComboColumn->GetWindowText(strName);
	int nCol = m_wndGrid.GetColFromName(strName);
	if (nCol == INVALID)
		return;

	// Update the grid.
	CString strMask;
	pMask->GetWindowText(strMask);
	m_wndGrid.SetColMask(nCol, strMask);
}

void CBHMDBColumnPropPage::OnKillfocusEditPromptchar() 
{
	// The user modified the "PromptChar" property.

	// Calc the current col number.
	CString strName;
	pComboColumn->GetWindowText(strName);
	int nCol = m_wndGrid.GetColFromName(strName);
	if (nCol == INVALID)
		return;

	// The PromptChar can only be 1 char long.
	CString strPromptChar;
	pPromptChar->GetWindowText(strPromptChar);
	if (strPromptChar.GetLength() != 1)
		return;

	// Update the grid.
	m_wndGrid.SetColPromptChar(nCol, strPromptChar);
}

void CBHMDBColumnPropPage::OnKillfocusEditWidth() 
{
	// The user modified the "Width" property.

	// Calc the current col number.
	CString strName;
	pComboColumn->GetWindowText(strName);
	int nCol = m_wndGrid.GetColFromName(strName);
	if (nCol == INVALID)
		return;

	CString strWidth;
	pWidth->GetWindowText(strWidth);

	// Width can not less than 0.
	int nWidth = atoi(strWidth);
	if (nWidth < 0)
		return;

	// Update the grid.
	m_wndGrid.SetColWidth(nCol, nWidth);
}

void CBHMDBColumnPropPage::OnRadioAlignmentCenter() 
{
	// The user set the "Alignment" property to be "Center".

	// Calc the current col number.
	CString strName;
	pComboColumn->GetWindowText(strName);
	int nCol = m_wndGrid.GetColFromName(strName);
	if (nCol == INVALID)
		return;

	// Update the grid.
	m_wndGrid.SetColAlignment(nCol, DT_CENTER);
}

void CBHMDBColumnPropPage::OnRadioAlignmentLeft() 
{
	// The user set the "Alignment" property to be "Left".

	// Calc the current col number.
	CString strName;
	pComboColumn->GetWindowText(strName);
	int nCol = m_wndGrid.GetColFromName(strName);
	if (nCol == INVALID)
		return;

	// Update the grid.
	m_wndGrid.SetColAlignment(nCol, DT_LEFT);
}

void CBHMDBColumnPropPage::OnRadioAlignmentRight() 
{
	// The user set the "Alignment" property to be "Right".

	// Calc the current col number.
	CString strName;
	pComboColumn->GetWindowText(strName);
	int nCol = m_wndGrid.GetColFromName(strName);
	if (nCol == INVALID)
		return;

	// Update the grid.
	m_wndGrid.SetColAlignment(nCol, DT_RIGHT);
}

void CBHMDBColumnPropPage::OnRadioCaptionalignmentAsalignment() 
{
	// The user set the "CaptionAlignment" property to be "AsAlignment".

	// Calc the current col number.
	CString strName;
	pComboColumn->GetWindowText(strName);
	int nCol = m_wndGrid.GetColFromName(strName);
	if (nCol == INVALID)
		return;

	// Update the grid.
	m_wndGrid.SetColHeaderAlignment(nCol, -1);
}

void CBHMDBColumnPropPage::OnRadioCaptionalignmentCenter() 
{
	// The user set the "CaptionAlignment" property to be "Center".

	// Calc the current col number.
	CString strName;
	pComboColumn->GetWindowText(strName);
	int nCol = m_wndGrid.GetColFromName(strName);
	if (nCol == INVALID)
		return;

	// Update the grid.
	m_wndGrid.SetColHeaderAlignment(nCol, DT_CENTER);
}

void CBHMDBColumnPropPage::OnRadioCaptionalignmentLeft() 
{
	// The user set the "CaptionAlignment" property to be "Left".

	// Calc the current col number.
	CString strName;
	pComboColumn->GetWindowText(strName);
	int nCol = m_wndGrid.GetColFromName(strName);
	if (nCol == INVALID)
		return;

	// Update the grid.
	m_wndGrid.SetColHeaderAlignment(nCol, DT_LEFT);
}

void CBHMDBColumnPropPage::OnRadioCaptionalignmentRight() 
{
	// The user set the "Alignment" property to be "Right".

	// Calc the current col number.
	CString strName;
	pComboColumn->GetWindowText(strName);
	int nCol = m_wndGrid.GetColFromName(strName);
	if (nCol == INVALID)
		return;

	// Update the grid.
	m_wndGrid.SetColHeaderAlignment(nCol, DT_RIGHT);
}

void CBHMDBColumnPropPage::OnRadioCaseLower() 
{
	// The user set the "Case" property to be "Lower".

	// Calc the current col number.
	CString strName;
	pComboColumn->GetWindowText(strName);
	int nCol = m_wndGrid.GetColFromName(strName);
	if (nCol == INVALID)
		return;

	// Update the grid.
	m_wndGrid.SetColCase(nCol, CASE_LOWER);
}

void CBHMDBColumnPropPage::OnRadioCaseNone() 
{
	// The user set the "Case" property to be "None".

	// Calc the current col number.
	CString strName;
	pComboColumn->GetWindowText(strName);
	int nCol = m_wndGrid.GetColFromName(strName);
	if (nCol == INVALID)
		return;

	// Update the grid.
	m_wndGrid.SetColCase(nCol, CASE_NONE);
}

void CBHMDBColumnPropPage::OnRadioCaseUpper() 
{
	// The user set the "Case" property to be "Upper".

	// Calc the current col number.
	CString strName;
	pComboColumn->GetWindowText(strName);
	int nCol = m_wndGrid.GetColFromName(strName);
	if (nCol == INVALID)
		return;

	// Update the grid.
	m_wndGrid.SetColCase(nCol, CASE_UPPER);
}

void CBHMDBColumnPropPage::OnRadioDatatypeBoolean() 
{
	// The user set the "Case" property to be "None".

	// Calc the current col number.
	CString strName;
	pComboColumn->GetWindowText(strName);
	int nCol = m_wndGrid.GetColFromName(strName);
	if (nCol == INVALID)
		return;

	// Update the grid.
	m_wndGrid.SetColDataType(nCol, DataTypeVARIANT_BOOL);
}

void CBHMDBColumnPropPage::OnRadioDatatypeByte() 
{
	// The user set the "DataType" property to be "Byte".

	// Calc the current col number.
	CString strName;
	pComboColumn->GetWindowText(strName);
	int nCol = m_wndGrid.GetColFromName(strName);
	if (nCol == INVALID)
		return;

	// Update the grid.
	m_wndGrid.SetColDataType(nCol, DataTypeByte);
}

void CBHMDBColumnPropPage::OnRadioDatatypeCurrency() 
{
	// The user set the "DataType" property to be "Currency".

	// Calc the current col number.
	CString strName;
	pComboColumn->GetWindowText(strName);
	int nCol = m_wndGrid.GetColFromName(strName);
	if (nCol == INVALID)
		return;

	// Update the grid.
	m_wndGrid.SetColDataType(nCol, DataTypeCurrency);
}

void CBHMDBColumnPropPage::OnRadioDatatypeDate() 
{
	// The user set the "DataType" property to be "Date".

	// Calc the current col number.
	CString strName;
	pComboColumn->GetWindowText(strName);
	int nCol = m_wndGrid.GetColFromName(strName);
	if (nCol == INVALID)
		return;

	// Update the grid.
	m_wndGrid.SetColDataType(nCol, DataTypeDate);
}

void CBHMDBColumnPropPage::OnRadioDatatypeInteger() 
{
	// The user set the "DataType" property to be "Integer".

	// Calc the current col number.
	CString strName;
	pComboColumn->GetWindowText(strName);
	int nCol = m_wndGrid.GetColFromName(strName);
	if (nCol == INVALID)
		return;

	// Update the grid.
	m_wndGrid.SetColDataType(nCol, DataTypeInteger);
}

void CBHMDBColumnPropPage::OnRadioDatatypeLong() 
{
	// The user set the "DataType" property to be "Long".

	// Calc the current col number.
	CString strName;
	pComboColumn->GetWindowText(strName);
	int nCol = m_wndGrid.GetColFromName(strName);
	if (nCol == INVALID)
		return;

	// Update the grid.
	m_wndGrid.SetColDataType(nCol, DataTypeLong);
}

void CBHMDBColumnPropPage::OnRadioDatatypeSingle() 
{
	// The user set the "DataType" property to be "Single".

	// Calc the current col number.
	CString strName;
	pComboColumn->GetWindowText(strName);
	int nCol = m_wndGrid.GetColFromName(strName);
	if (nCol == INVALID)
		return;

	// Update the grid.
	m_wndGrid.SetColDataType(nCol, DataTypeSingle);
}

void CBHMDBColumnPropPage::OnRadioDatatypeText() 
{
	// The user set the "DataType" property to be "Text".

	// Calc the current col number.
	CString strName;
	pComboColumn->GetWindowText(strName);
	int nCol = m_wndGrid.GetColFromName(strName);
	if (nCol == INVALID)
		return;

	// Update the grid.
	m_wndGrid.SetColDataType(nCol, DataTypeText);
}

void CBHMDBColumnPropPage::OnRadioStyleButton() 
{
	// The user set the "Style" property to be "Button".

	// Calc the current col number.
	CString strName;
	pComboColumn->GetWindowText(strName);
	int nCol = m_wndGrid.GetColFromName(strName);
	if (nCol == INVALID)
		return;

	// Update the grid.
	m_wndGrid.SetColControl(nCol,  COLSTYLE_BUTTON);

	// Update other controls on page.
	pSetupComboBox->ShowWindow(SW_HIDE);
	pDateMask->ShowWindow(SW_HIDE);
	pStaticDateMask->ShowWindow(SW_HIDE);
}

void CBHMDBColumnPropPage::OnRadioStyleCheckbox() 
{
	// The user set the "Style" property to be "CheckBox".

	// Calc the current col number.
	CString strName;
	pComboColumn->GetWindowText(strName);
	int nCol = m_wndGrid.GetColFromName(strName);
	if (nCol == INVALID)
		return;

	// Update the grid.
	m_wndGrid.SetColControl(nCol, COLSTYLE_CHECKBOX);

	// Update other controls on page.
	pSetupComboBox->ShowWindow(SW_HIDE);
	pDateMask->ShowWindow(SW_HIDE);
	pStaticDateMask->ShowWindow(SW_HIDE);
}

void CBHMDBColumnPropPage::OnRadioStyleCombobox() 
{
	// The user set the "Style" property to be "ComboBox".

	// Calc the current col number.
	CString strName;
	pComboColumn->GetWindowText(strName);
	int nCol = m_wndGrid.GetColFromName(strName);
	if (nCol == INVALID)
		return;

	// Update the grid.
	m_wndGrid.SetColControl(nCol, COLSTYLE_COMBOBOX);

	// Update other controls on page.
	pSetupComboBox->ShowWindow(SW_SHOW);
	pDateMask->ShowWindow(SW_HIDE);
	pStaticDateMask->ShowWindow(SW_HIDE);
}

void CBHMDBColumnPropPage::OnRadioStyleEdit() 
{
	// The user set the "Style" property to be "Edit".

	// Calc the current col number.
	CString strName;
	pComboColumn->GetWindowText(strName);
	int nCol = m_wndGrid.GetColFromName(strName);
	if (nCol == INVALID)
		return;

	// Update the grid.
	m_wndGrid.SetColControl(nCol, COLSTYLE_EDIT);

	// Update other controls on page.
	pDateMask->ShowWindow(SW_HIDE);
	pStaticDateMask->ShowWindow(SW_HIDE);
	pSetupComboBox->ShowWindow(SW_HIDE);
}

void CBHMDBColumnPropPage::OnRadioStyleEditbutton() 
{
	// The user set the "Style" property to be "EditButton".

	// Calc the current col number.
	CString strName;
	pComboColumn->GetWindowText(strName);
	int nCol = m_wndGrid.GetColFromName(strName);
	if (nCol == INVALID)
		return;

	// Update the grid.
	m_wndGrid.SetColControl(nCol, COLSTYLE_EDITBUTTON);

	// Update other controls on page.
	pSetupComboBox->ShowWindow(SW_HIDE);
	pDateMask->ShowWindow(SW_HIDE);
	pStaticDateMask->ShowWindow(SW_HIDE);
}

void CBHMDBColumnPropPage::OnButtonSetupcombobox() 
{
	// The user want to seup the content of the combobox.

	// Calc the current col number.
	CString strName;
	pComboColumn->GetWindowText(strName);
	int nCol = m_wndGrid.GetColFromName(strName);
	if (nCol == INVALID)
		return;

	// Copy current content out.	
	Col col;
	m_wndGrid.CloneColProp(nCol, &col);

	// Init the variable of the input dialog.
	CGridInputDlg dlgInput;
	dlgInput.m_strChoiceList = col.strChoiceList;

	// Let the user input content in one dialog.
	if (dlgInput.DoModal() == IDOK)
	{
		// Update the grid.
		m_wndGrid.SetColChoiceList(nCol, dlgInput.m_strChoiceList);
	}
}

void CBHMDBColumnPropPage::OnButtonAddcolumn() 
{
	// The user want to add a new column.

	// Create a dialog to get the column name.
	CAddColumnDlg dlg;

	// Set the callback pointer.
	dlg.SetPagePtr(this);
	if (dlg.DoModal() == IDOK)
	{
		// Update the column combo box.
		pComboColumn->AddString(dlg.m_strName);
		
		// Update the grid.
		int nCol = InsertCol();
		m_wndGrid.SetColTitle(nCol, dlg.m_strName);
		m_wndGrid.SetColName(nCol, dlg.m_strName);
	}
}

void CBHMDBColumnPropPage::OnButtonDeletecolumn() 
{
	// The user want to delete current column.

	// Calc the current col number.
	CString strName;
	pComboColumn->GetWindowText(strName);
	int nCol = m_wndGrid.GetColFromName(strName);
	if (nCol == INVALID)
		return;

	// Delete the string from column combo box.
	pComboColumn->DeleteString(pComboColumn->GetCurSel());

	// Delete the column from grid.
	m_wndGrid.RemoveCol(nCol);
	
	// Update other controls in page.
	OnSelendokComboColumn();

	// Can not delete column without any column exists.
	if (m_wndGrid.GetColCount() == 0)
		pDeleteColumn->EnableWindow(FALSE);
}

int CBHMDBColumnPropPage::InsertCol()
/*
Routine Description:
	Insert one col into grid.

Parameters:
	None.

Return Value:
	The col number of the inserted col.
*/
{
	// Get the current col number.
	int nRow, nCol;
	m_wndGrid.GetCurrentCell(nRow, nCol);

	int nGroup = INVALID;

	for (int i = 0; i < m_wndGrid.GetGroupCount(); i ++)
	{
		if (m_wndGrid.IsGroupSelected(i + 1))
		{
			// One group is selected, we should insert column in this group.
			nGroup = i;
			nCol = 0;
			break;
		}
	}

//	if (nCol != 0)
//	{
//		nCol --;
//		nGroup = m_wndGrid.GetGroupFromCol(nCol);
//	}

	if (nGroup != INVALID)
	{
		// Insert col in group.
		m_wndGrid.InsertColInGroup(nGroup + 1);

		// Calc the col number of the inserted col.
		nCol = m_wndGrid.GetGroupColCount(nGroup + 1);
		nCol = m_wndGrid.GetColFromGroup(nGroup, nCol);
	}
	else
	{
		// Insert col to the end of the grid.
		m_wndGrid.InsertCol();

		// Calc the col number of the inserted col.
		nCol = m_wndGrid.GetColCount();
	}

	// Init the user attribute array.
	m_wndGrid.GetColUserAttribRef(nCol).SetSize(2);

	// Select the newly inserted col.
	pComboColumn->SetCurSel(nCol - 1);

	// Update other controls in page.
	OnSelendokComboColumn();

	return nCol;
}

void CBHMDBColumnPropPage::OnButtonReset() 
{
	// The user want to empty the grid.

	// Empty the column combo box.
	pComboColumn->ResetContent();

	// Remove all columns.
	m_wndGrid.SetColCount(0);

	// Can not delete column without any columns exists.
	pDeleteColumn->EnableWindow(FALSE);
	
	// Update other controls in page.
	OnSelendokComboColumn();
}

void CBHMDBColumnPropPage::OnCheckPromptinclude() 
{
	// The user modified the value of the "PromptInclude" property.

	// Calc the current col number.
	CString strName;
	pComboColumn->GetWindowText(strName);
	int nCol = m_wndGrid.GetColFromName(strName);
	if (nCol == INVALID)
		return;

	// Update the grid.
	m_wndGrid.SetColPromptInclude(nCol, pPromptInclude->GetCheck());
}

void CBHMDBColumnPropPage::OnSelendokComboDatafield() 
{
	// The user modified the value of the "DataField" property.

	// Calc the current col number.
	CString strName;
	pComboColumn->GetWindowText(strName);
	int nCol = m_wndGrid.GetColFromName(strName);
	if (nCol == INVALID)
		return;

	CString strField;
	pDataField->GetWindowText(strField);

	// If the string is " ", it means thatt empty.
	if (strField == _T(" "))
		strField.Empty();

	// Update the grid.
	m_wndGrid.GetColUserAttribRef(nCol).SetAt(0, strField);
}

void CBHMDBColumnPropPage::OnButtonFields() 
{
	// The user want to insert some fields into grid.

	// Create the dialog to choose fields.
	CFieldsDlg dlg;
	int i;

	// Init the fields array.
	dlg.m_arField.RemoveAll();
	
	if (m_pGridCtrl)
	// The control underlying is grid control.
	{
		// Fill the fields array with all of the fields in data source.
		for (i = 0; i < m_pGridCtrl->m_apColumnData.GetSize(); i ++)
			dlg.m_arField.Add(m_pGridCtrl->m_apColumnData[i]->strName);
	}
	else if (m_pDropDownCtrl)
	// The control underlying is dropdown control.
	{
		// Fill the fields array with all of the fields in data source.
		for (i = 0; i < m_pDropDownCtrl->m_apColumnData.GetSize(); i ++)
			dlg.m_arField.Add(m_pDropDownCtrl->m_apColumnData[i]->strName);
	}
	else
	// The control underlying is combo control.
	{
		// Fill the fields array with all of the fields in data source.
		for (i = 0; i < m_pComboCtrl->m_apColumnData.GetSize(); i ++)
			dlg.m_arField.Add(m_pComboCtrl->m_apColumnData[i]->strName);
	}

	if (dlg.DoModal() == IDOK)
	{
		// Prevent the grid from flickering.
		m_wndGrid.LockUpdate();

		for (i = 0; i < dlg.m_arField.GetSize(); i ++)
		// Insert one col into grid for every selected field, and bind them to the field.
		{
			// Insert one col into grid.
			int nCol = InsertCol();
			
			// Use the field name as the column title.
			m_wndGrid.SetColTitle(nCol, dlg.m_arField[i]);

			// Use the field name as the column name.
			m_wndGrid.SetColName(nCol, dlg.m_arField[i]);

			// If the field is readonly, set the column to readonly too.
			m_wndGrid.SetColReadOnly(nCol, IsFieldForceLock(dlg.m_arField[i]));

			// Use the field name as the value of "DataField" property.
			m_wndGrid.GetColUserAttribRef(nCol).SetAt(0, dlg.m_arField[i]);

			// If the field can not be null, set the column to not nullable too.
			m_wndGrid.GetColUserAttribRef(nCol).SetAt(1, IsFieldForceNotNullable(dlg.m_arField[i]) ? _T("T") : _T(""));
		}

		// Update the grid now.
		m_wndGrid.LockUpdate(FALSE);
		m_wndGrid.Invalidate();
	}
}

void CBHMDBColumnPropPage::OnSelendokComboDatemask() 
{
	// The user changed the value of the "Mask" property

	// Calc the current col number.
	CString strName;
	pComboColumn->GetWindowText(strName);
	int nCol = m_wndGrid.GetColFromName(strName);
	if (nCol == INVALID)
		return;

	// Update the grid.
	CString strMask;
	pDateMask->GetWindowText(strMask);
	m_wndGrid.SetColMask(nCol, strMask);
}

int CBHMDBColumnPropPage::CalcPropertyOrdinal(int nProperty)
/*
Routine Description:
	Calc the property number of the given item number in the "Properties" list box.

Parameters:
	nProperty		The item number in the list box.

Return Value:
	The ordinal of the property in the set of properties.
*/
{
	int nOrdinal = nProperty;

	if (m_pDropDownCtrl || m_pComboCtrl)
	{
		// If the underlying control is not the grid control, the item number is not 
		// equal to the property number, because some properties only exist in grid control.

		if (nOrdinal == 1)
			nOrdinal = 2;
		else if (nOrdinal == 2)
			nOrdinal = 3;
		else if (nOrdinal == 3)
			nOrdinal = 4;
		else if (nOrdinal == 4)
			nOrdinal = 5;
		else if (nOrdinal == 5)
			nOrdinal = 6;
		else if (nOrdinal == 6)
			nOrdinal = 9;
		else if (nOrdinal == 7)
			nOrdinal = 10;
		else if (nOrdinal == 8)
			nOrdinal = 11;
		else if (nOrdinal == 9)
			nOrdinal = 14;
		else if (nOrdinal == 10)
			nOrdinal = 18;
		else if (nOrdinal == 11)
			nOrdinal = 19;
	}

	return nOrdinal;
}

void CBHMDBColumnPropPage::OnRadioStyleDatemask() 
{
	// The user set the value of "Style" property to be "DateMask".

	// Calc the current col number.
	CString strName;
	pComboColumn->GetWindowText(strName);
	int nCol = m_wndGrid.GetColFromName(strName);
	if (nCol == INVALID)
		return;

	// Update the grid.
	m_wndGrid.SetColControl(nCol, COLSTYLE_DATEMASK);

	// Update other controls in page.
	pSetupComboBox->ShowWindow(SW_HIDE);
	pDateMask->ShowWindow(SW_SHOW);
	pStaticDateMask->ShowWindow(SW_SHOW);
}

int CBHMDBColumnPropPage::CalcAlignmentOrdinal(UINT align)
/*
Routine Description:
	Calc the ordinal of the correspond radio button in "Alignment" radio group or "CaptionAlignment"
	radio group.

Parameters:
	nalign		The alignment value.

Return Value:
	The ordinal of the correspond radio button.
*/
{
	// The first button is "Left", the second is "Right", the third is "Center", and the forth
	// is "AsAlignment" for "CaptionAlignment" property.

	if (align == DT_CENTER)
		return 2;
	if (align == DT_LEFT)
		return 0;
	else if (align == DT_RIGHT)
		return 1;
	else
		return 3;
}

void CBHMDBColumnPropPage::OnRadioStyleCombolist() 
{
	// The user modified the value of "Style" property to be "ComboList".

	// Calc the current col number.
	CString strName;
	pComboColumn->GetWindowText(strName);
	int nCol = m_wndGrid.GetColFromName(strName);
	if (nCol == INVALID)
		return;

	// Update the grid.
	m_wndGrid.SetColControl(nCol, COLSTYLE_COMBOBOX_DROPLIST);

	// Update other controls in page.
	pSetupComboBox->ShowWindow(SW_SHOW);
	pDateMask->ShowWindow(SW_HIDE);
	pStaticDateMask->ShowWindow(SW_HIDE);
}

void CBHMDBColumnPropPage::InitGrid()
/*
Routine Description:
	Init the grid from the data of the underlying control.

Parameters:
	None.

Return Value:
	None.
*/
{
	int i;

	// Prevent the grid from flickering.
	m_wndGrid.LockUpdate();

	// Clear the grid.
	m_wndGrid.SetRowCount(0);
	m_wndGrid.SetColCount(0);

	// Init the datafield combo box.
	pDataField->ResetContent();

	// Value " " means no binding data field.
	pDataField->AddString(_T(" "));

	// Init the column combo box.
	pComboColumn->ResetContent();
		
	int nColumns;

	if (m_pGridCtrl)
	{
		// The underlying control is grid control.

		// Load layout data from control.
		m_wndGrid.LoadLayout(&m_pGridCtrl->m_arGroups, &m_pGridCtrl->m_arCols, &m_pGridCtrl->m_arCells);

		// Get the column infomation from data source.
		m_pGridCtrl->GetColumnInfo();

		// Get the count of the fields in data source.
		nColumns = m_pGridCtrl->m_apColumnData.GetSize();
		
		// Insert every field into datafield combo box.
		for (i = 0; i < nColumns; i ++)
			pDataField->AddString(m_pGridCtrl->m_apColumnData[i]->strName);
		
		// Update other properties from underlying control.
		m_wndGrid.SetBackColor(m_pGridCtrl->TranslateColor(m_pGridCtrl->GetBackColor()));
		m_wndGrid.SetForeColor(m_pGridCtrl->TranslateColor(m_pGridCtrl->GetForeColor()));
		m_wndGrid.SetHeaderForeColor(m_pGridCtrl->TranslateColor(m_pGridCtrl->m_clrHeadFore));
		m_wndGrid.SetHeaderBackColor(m_pGridCtrl->TranslateColor(m_pGridCtrl->m_clrHeadBack));
		m_wndGrid.SetRowHeight(m_pGridCtrl->GetRowHeight());
		m_wndGrid.SetHeaderHeight(m_pGridCtrl->GetHeaderHeight());
	}
	else if (m_pDropDownCtrl)
	{
		// The underlying control is dropdown control.

		// Load layout data from control.
		m_wndGrid.LoadLayout(&m_pDropDownCtrl->m_arGroups, &m_pDropDownCtrl->m_arCols, &m_pDropDownCtrl->m_arCells);

		// Get the column infomation from data source.
		m_pDropDownCtrl->GetColumnInfo();

		// Get the count of the fields in data source.
		nColumns = m_pDropDownCtrl->m_apColumnData.GetSize();
		
		// Insert every field into datafield combo box.
		for (i = 0; i < nColumns; i ++)
			pDataField->AddString(m_pDropDownCtrl->m_apColumnData[i]->strName);
		
		// Update other properties from underlying control.
		m_wndGrid.SetBackColor(m_pDropDownCtrl->TranslateColor(m_pDropDownCtrl->GetBackColor()));
		m_wndGrid.SetForeColor(m_pDropDownCtrl->TranslateColor(m_pDropDownCtrl->GetForeColor()));
		m_wndGrid.SetHeaderForeColor(m_pDropDownCtrl->TranslateColor(m_pDropDownCtrl->m_clrHeadFore));
		m_wndGrid.SetHeaderBackColor(m_pDropDownCtrl->TranslateColor(m_pDropDownCtrl->m_clrHeadBack));
		m_wndGrid.SetRowHeight(m_pDropDownCtrl->GetRowHeight());
		m_wndGrid.SetHeaderHeight(m_pDropDownCtrl->GetHeaderHeight());
	}
	else if (m_pComboCtrl)
	{
		// The underlying control is combo control.

		// Load layout data from control.
		m_wndGrid.LoadLayout(&m_pComboCtrl->m_arGroups, &m_pComboCtrl->m_arCols, &m_pComboCtrl->m_arCells);

		// Get the column infomation from data source.
		m_pComboCtrl->GetColumnInfo();

		// Get the count of the fields in data source.
		nColumns = m_pComboCtrl->m_apColumnData.GetSize();
		
		// Insert every field into datafield combo box.
		for (i = 0; i < nColumns; i ++)
			pDataField->AddString(m_pComboCtrl->m_apColumnData[i]->strName);
		
		// Update other properties from underlying control.
		m_wndGrid.SetBackColor(m_pComboCtrl->TranslateColor(m_pComboCtrl->GetBackColor()));
		m_wndGrid.SetForeColor(m_pComboCtrl->TranslateColor(m_pComboCtrl->GetForeColor()));
		m_wndGrid.SetHeaderForeColor(m_pComboCtrl->TranslateColor(m_pComboCtrl->m_clrHeadFore));
		m_wndGrid.SetHeaderBackColor(m_pComboCtrl->TranslateColor(m_pComboCtrl->m_clrHeadBack));
		m_wndGrid.SetRowHeight(m_pComboCtrl->GetRowHeight());
		m_wndGrid.SetHeaderHeight(m_pComboCtrl->GetHeaderHeight());
	}

	// Insert every column name into column combo box.
	nColumns = m_wndGrid.GetColCount();

	for (i = 0; i < nColumns; i ++)
		pComboColumn->AddString(m_wndGrid.GetColName(i + 1));

	// Can not delete col without any col.
	if (nColumns == 0)
		pDeleteColumn->EnableWindow(FALSE);
	else
	// Display the first column default.
		pComboColumn->SetCurSel(0);
	
	// Update the grid.	
	m_wndGrid.LockUpdate(FALSE);
	
	if (IsWindow(m_wndGrid.m_hWnd))
	{
		m_wndGrid.ResetScrollBars();
		m_wndGrid.Invalidate();
	}

	// Update other controls in page.
	OnSelchangeListProperty();
}

BOOL CBHMDBColumnPropPage::IsFieldForceLock(CString strField)
/*
Routine Description:
	Determines if the given field is readonly.

Parameters:
	strField		The name of the field.

Return Value:
	If the field is readonly, return TRUE; Otherwise, return FALSE.
*/
{
	int nField;

	if (m_pGridCtrl)
	{
		// The control underlying is grid control.

		// Calc the field ordinal from name.
		nField = m_pGridCtrl->GetFieldFromStr(strField);

		if (nField == -1)
		// Field does not exist, so it is not readonly.
			return FALSE;
		else
		// Determine the result from the field flags.
			return m_pGridCtrl->m_pColumnInfo[nField].dwFlags & DBCOLUMNFLAGS_WRITE;
	}
	else if (m_pDropDownCtrl)
	{
		// The control underlying is dropdown control.

		// Calc the field ordinal from name.
		nField = m_pDropDownCtrl->GetFieldFromStr(strField);

		if (nField == -1)
		// Field does not exist, so it is not readonly.
			return FALSE;
		else
		// Determine the result from the field flags.
			return m_pDropDownCtrl->m_pColumnInfo[nField].dwFlags & DBCOLUMNFLAGS_WRITE;
	}
	else
	{
		// The control underlying is combo control.

		// Calc the field ordinal from name.
		nField = m_pComboCtrl->GetFieldFromStr(strField);

		if (nField == -1)
		// Field does not exist, so it is not readonly.
			return FALSE;
		else
		// Determine the result from the field flags.
			return m_pComboCtrl->m_pColumnInfo[nField].dwFlags & DBCOLUMNFLAGS_WRITE;
	}
}

BOOL CBHMDBColumnPropPage::IsFieldForceNotNullable(CString strField)
/*
Routine Description:
	Determines if the given field can fill with null value.

Parameters:
	strField		The name of the field.

Return Value:
	If the field is be null, return FALSE; Otherwise, return TRUE.
*/
{
	int nField;

	if (m_pGridCtrl)
	{
		// The control underlying is grid control.

		// Calc the field ordinal from name.
		nField = m_pGridCtrl->GetFieldFromStr(strField);

		if (nField == -1)
		// Field does not exist, so it can use null.
			return FALSE;
		else
		// Determine the result from the field flags.
			return m_pGridCtrl->m_pColumnInfo[nField].dwFlags & DBCOLUMNFLAGS_ISNULLABLE;
	}
	else if (m_pDropDownCtrl)
	{
		// The control underlying is dropdown control.

		// Calc the field ordinal from name.
		nField = m_pDropDownCtrl->GetFieldFromStr(strField);

		if (nField == -1)
		// Field does not exist, so it can use null.
			return FALSE;
		else
		// Determine the result from the field flags.
			return m_pDropDownCtrl->m_pColumnInfo[nField].dwFlags & DBCOLUMNFLAGS_ISNULLABLE;
	}
	else
	{
		// The control underlying is dropdown control.

		// Calc the field ordinal from name.
		nField = m_pComboCtrl->GetFieldFromStr(strField);

		if (nField == -1)
		// Field does not exist, so it can use null.
			return FALSE;
		else
		// Determine the result from the field flags.
			return m_pComboCtrl->m_pColumnInfo[nField].dwFlags & DBCOLUMNFLAGS_ISNULLABLE;
	}
}

void CBHMDBColumnPropPage::OnRadioStyleNumeric() 
{
	// The user modified the value of "Style" property to be "Numeric".

	// Calc the current col number.
	CString strName;
	pComboColumn->GetWindowText(strName);
	int nCol = m_wndGrid.GetColFromName(strName);
	if (nCol == INVALID)
		return;

	// Update the grid.
	m_wndGrid.SetColControl(nCol, COLSTYLE_NUM);

	// Update other controls in page.
	pSetupComboBox->ShowWindow(SW_HIDE);
	pDateMask->ShowWindow(SW_HIDE);
	pStaticDateMask->ShowWindow(SW_HIDE);
}
