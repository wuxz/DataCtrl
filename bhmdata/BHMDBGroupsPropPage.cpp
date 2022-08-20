// BHMDBGroupsPropPage.cpp : implementation file
//

#include "stdafx.h"
#include "bhmdata.h"
#include "BHMDBGroupsPropPage.h"
#include "BHMDBGridCtl.h"
#include "GridInputDlg.h"
#include "AddColumnDlg.h"
#include "BHMDBDropDownCtl.h"
#include "BHMDBComboCtl.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

#define pComboGroup ((CComboBox *)GetDlgItem(IDC_COMBO_GROUP))
#define pForeColor ((CButton *)GetDlgItem(IDC_BUTTON_FORECOLOR))
#define pBackColor ((CButton *)GetDlgItem(IDC_BUTTON_BACKCOLOR))
#define pAllowSizing ((CButton *)GetDlgItem(IDC_CHECK_ALLOWSIZING))
#define pVisible ((CButton *)GetDlgItem(IDC_CHECK_VISIBLE))
#define pName (GetDlgItem(IDC_EDIT_NAME))
#define pTitle (GetDlgItem(IDC_EDIT_TITLE))
#define pWidth (GetDlgItem(IDC_EDIT_WIDTH))
#define pLevels (GetDlgItem(IDC_EDIT_LEVELS))

#define pAddGroup ((CButton *)GetDlgItem(IDC_BUTTON_ADDGROUP))
#define pDeleteGroup ((CButton *)GetDlgItem(IDC_BUTTON_DELETEGROUP))
#define pReset ((CButton *)GetDlgItem(IDC_BUTTON_RESET))

/////////////////////////////////////////////////////////////////////////////
// CBHMDBGroupsPropPage dialog

IMPLEMENT_DYNCREATE(CBHMDBGroupsPropPage, COlePropertyPage)


/////////////////////////////////////////////////////////////////////////////
// Message map

BEGIN_MESSAGE_MAP(CBHMDBGroupsPropPage, COlePropertyPage)
	//{{AFX_MSG_MAP(CBHMDBGroupsPropPage)
	ON_BN_CLICKED(IDC_BUTTON_ADDGROUP, OnButtonAddgroup)
	ON_BN_CLICKED(IDC_BUTTON_BACKCOLOR, OnButtonBackcolor)
	ON_BN_CLICKED(IDC_BUTTON_DELETEGROUP, OnButtonDeletegroup)
	ON_BN_CLICKED(IDC_BUTTON_FORECOLOR, OnButtonForecolor)
	ON_BN_CLICKED(IDC_BUTTON_RESET, OnButtonReset)
	ON_BN_CLICKED(IDC_CHECK_ALLOWSIZING, OnCheckAllowsizing)
	ON_BN_CLICKED(IDC_CHECK_VISIBLE, OnCheckVisible)
	ON_CBN_SELENDOK(IDC_COMBO_GROUP, OnSelendokComboGroup)
	ON_EN_KILLFOCUS(IDC_EDIT_TITLE, OnKillfocusEditTitle)
	ON_EN_KILLFOCUS(IDC_EDIT_WIDTH, OnKillfocusEditWidth)
	ON_EN_KILLFOCUS(IDC_EDIT_NAME, OnKillfocusEditName)
	ON_EN_KILLFOCUS(IDC_EDIT_LEVELS, OnKillfocusEditLevels)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// Initialize class factory and guid

// {05803601-6D0C-11D3-A7FE-0080C8763FA4}
IMPLEMENT_OLECREATE_EX(CBHMDBGroupsPropPage, "BHMData.CBHMDBGroupsPropPage",
	0x5803601, 0x6d0c, 0x11d3, 0xa7, 0xfe, 0x0, 0x80, 0xc8, 0x76, 0x3f, 0xa4)


/////////////////////////////////////////////////////////////////////////////
// CBHMDBGroupsPropPage::CBHMDBGroupsPropPageFactory::UpdateRegistry -
// Adds or removes system registry entries for CBHMDBGroupsPropPage

BOOL CBHMDBGroupsPropPage::CBHMDBGroupsPropPageFactory::UpdateRegistry(BOOL bRegister)
{
	// TODO: Define string resource for page type; replace '0' below with ID.

	if (bRegister)
		return AfxOleRegisterPropertyPageClass(AfxGetInstanceHandle(),
			m_clsid, IDS_BHMDB_GROUPS_PPG);
	else
		return AfxOleUnregisterClass(m_clsid, NULL);
}


/////////////////////////////////////////////////////////////////////////////
// CBHMDBGroupsPropPage::CBHMDBGroupsPropPage - Constructor

// TODO: Define string resource for page caption; replace '0' below with ID.

CBHMDBGroupsPropPage::CBHMDBGroupsPropPage() :
	COlePropertyPage(IDD, IDS_BHMDB_GROUPS_PPG_CAPTION)
{
	//{{AFX_DATA_INIT(CBHMDBGroupsPropPage)
	// NOTE: ClassWizard will add member initialization here
	//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_DATA_INIT
	m_pGridCtrl = NULL;
	m_pDropDownCtrl = NULL;
	m_pComboCtrl = NULL;
	m_wndGrid.SetPagePtr(this);
}


/////////////////////////////////////////////////////////////////////////////
// CBHMDBGroupsPropPage::DoDataExchange - Moves data between page and properties

void CBHMDBGroupsPropPage::DoDataExchange(CDataExchange* pDX)
{
	// NOTE: ClassWizard will add DDP, DDX, and DDV calls here
	//    DO NOT EDIT what you see in these blocks of generated code !
	//{{AFX_DATA_MAP(CBHMDBGroupsPropPage)
	//}}AFX_DATA_MAP
	DDP_PostProcessing(pDX);

	if (!pDX->m_bSaveAndValidate)
	{
		// Loading properties from control.
		InitGrid();
	}
	else
	{
		// Saving properties into control.

		if (m_pGridCtrl)
		{
			// The underlying control is grid control.

			// Save the layout in page to grid control.
			m_wndGrid.SaveLayout(&m_pGridCtrl->m_arGroups, &m_pGridCtrl->m_arCols, &m_pGridCtrl->m_arCells);

			// Init the control from new data.
			m_pGridCtrl->InitGridFromProp();

			// Update other properties.
			m_pGridCtrl->SetRowHeight(m_wndGrid.GetRowHeight(1));
			m_pGridCtrl->SetHeaderHeight(m_wndGrid.GetRowHeight(0));

			m_pGridCtrl->m_bHasLayout = m_wndGrid.GetGroupCount() || m_wndGrid.GetColCount();
		}
		else if (m_pDropDownCtrl)
		{
			// The underlying control is dropdown control.

			// Save the layout in page to dropdown control.
			m_wndGrid.SaveLayout(&m_pDropDownCtrl->m_arGroups, &m_pDropDownCtrl->m_arCols, &m_pDropDownCtrl->m_arCells);

			// Init the control from new data.
			m_pDropDownCtrl->InitGridFromProp();

			// Update other properties.
			m_pDropDownCtrl->SetRowHeight(m_wndGrid.GetRowHeight(1));
			m_pDropDownCtrl->SetHeaderHeight(m_wndGrid.GetRowHeight(0));

			m_pDropDownCtrl->m_bHasLayout = m_wndGrid.GetGroupCount() || m_wndGrid.GetColCount();
		}
		else if (m_pComboCtrl)
		{
			// The underlying control is dropdown control.

			// Save the layout in page to dropdown control.
			m_wndGrid.SaveLayout(&m_pComboCtrl->m_arGroups, &m_pComboCtrl->m_arCols, &m_pComboCtrl->m_arCells);

			// Init the control from new data.
			m_pComboCtrl->InitGridFromProp();

			// Update other properties.
			m_pComboCtrl->SetRowHeight(m_wndGrid.GetRowHeight(1));
			m_pComboCtrl->SetHeaderHeight(m_wndGrid.GetRowHeight(0));
			
			m_pComboCtrl->m_bHasLayout = m_wndGrid.GetGroupCount() || m_wndGrid.GetColCount();
		}
	}
}


/////////////////////////////////////////////////////////////////////////////
// CBHMDBGroupsPropPage message handlers

void CBHMDBGroupsPropPage::OnButtonAddgroup() 
/*
Routine Description:
	Add one group into grid.

Parameters:
	None.

Return Value:
	None.
*/
{
	// Create a dialog to receive group name.
	CAddColumnDlg dlg;

	// Set the callback pointer.
	dlg.SetPagePtr(this);

	if (dlg.DoModal() == IDOK)
	{
		// Update the group name combo box.
		pComboGroup->AddString(dlg.m_strName);

		// Update the grid.
		m_wndGrid.InsertGroup();

		int nGroups = m_wndGrid.GetGroupCount();
		
		// Set name for the new group.
		m_wndGrid.SetGroupName(nGroups, dlg.m_strName);

		// Update other controls in page.
		pComboGroup->SetCurSel(nGroups - 1);
		OnSelendokComboGroup();
	}
}

void CBHMDBGroupsPropPage::OnButtonBackcolor() 
/*
Routine Description:
	Set the back color for current group.

Parameters:
	None.

Return Value:
	None.
*/
{
	// Get current group.
	CString strName;
	pComboGroup->GetWindowText(strName);

	int nGroup = m_wndGrid.GetGroupFromName(strName);

	if (nGroup == INVALID)
	{
		AfxMessageBox(IDS_ERROR_NOCURRENTCOLINCOLUMNPPG, MB_ICONSTOP);
		return;
	}

	// Choose a new color.
	COLORREF clrBack = m_wndGrid.GetGroupBackColor(nGroup);
	
	CColorDialog clrDlg(clrBack, CC_FULLOPEN | CC_RGBINIT | CC_SOLIDCOLOR);
	if (clrDlg.DoModal() == IDOK)
	// Update the grid.
		m_wndGrid.SetGroupBackColor(nGroup, clrDlg.GetColor());
}

void CBHMDBGroupsPropPage::OnButtonDeletegroup() 
/*
Routine Description:
	Delete current group.

Parameters:
	None.

Return Value:
	None.
*/
{
	// Get current group.
	CString strName;
	pComboGroup->GetWindowText(strName);

	int nGroup = m_wndGrid.GetGroupFromName(strName);

	if (nGroup == INVALID)
	{
		AfxMessageBox(IDS_ERROR_NOCURRENTCOLINCOLUMNPPG, MB_ICONSTOP);
		return;
	}

	// Update group name combo box.
	pComboGroup->DeleteString(pComboGroup->GetCurSel());
	
	// Update the grid.
	m_wndGrid.RemoveGroup(nGroup);

	if (m_wndGrid.GetGroupCount() == 0)
	// Can not delete group without any group exists.
		pDeleteGroup->EnableWindow(FALSE);

	// Update other controls in page.
	OnSelendokComboGroup();
}

void CBHMDBGroupsPropPage::OnButtonForecolor() 
/*
Routine Description:
	Change the fore color of current group.

Parameters:
	None.

Return Value:
	None.
*/
{
	// Get current group.
	CString strName;
	pComboGroup->GetWindowText(strName);

	int nGroup = m_wndGrid.GetGroupFromName(strName);

	if (nGroup == INVALID)
	{
		AfxMessageBox(IDS_ERROR_NOCURRENTCOLINCOLUMNPPG, MB_ICONSTOP);
		return;
	}

	// Choose a new color.
	COLORREF clrFore = m_wndGrid.GetGroupForeColor(nGroup);
	
	CColorDialog clrDlg(clrFore, CC_FULLOPEN | CC_RGBINIT | CC_SOLIDCOLOR);
	if (clrDlg.DoModal() == IDOK)
	// Update the grid.
		m_wndGrid.SetGroupForeColor(nGroup, clrDlg.GetColor());
}

void CBHMDBGroupsPropPage::OnButtonReset() 
/*
Routine Description:
	Remove all contents in grid.

Parameters:
	None.

Return Value:
	None.
*/
{
	// Update thte grid.
	m_wndGrid.SetGroupCount(0);

	// Update other controls in page.
	pDeleteGroup->EnableWindow(FALSE);
	
	OnSelendokComboGroup();
}

void CBHMDBGroupsPropPage::OnCheckAllowsizing() 
/*
Routine Description:
	The permission to resize group width will be changed.

Parameters:
	None.

Return Value:
	None.
*/
{
	// Get current group.
	CString strName;
	pComboGroup->GetWindowText(strName);

	int nGroup = m_wndGrid.GetGroupFromName(strName);

	if (nGroup == INVALID)
	{
		AfxMessageBox(IDS_ERROR_NOCURRENTCOLINCOLUMNPPG, MB_ICONSTOP);
		return;
	}

	// Update the grid.
	m_wndGrid.SetGroupAllowSizing(nGroup, pAllowSizing->GetCheck());
}

void CBHMDBGroupsPropPage::OnCheckVisible() 
/*
Routine Description:
	The visibility of current group will be changed.

Parameters:
	None.

Return Value:
	None.
*/
{
	// Get current group.
	CString strName;
	pComboGroup->GetWindowText(strName);

	int nGroup = m_wndGrid.GetGroupFromName(strName);

	if (nGroup == INVALID)
	{
		AfxMessageBox(IDS_ERROR_NOCURRENTCOLINCOLUMNPPG, MB_ICONSTOP);
		return;
	}

	// Update the grid.
	m_wndGrid.SetGroupVisible(nGroup, pVisible->GetCheck());
}

void CBHMDBGroupsPropPage::OnSelendokComboGroup() 
/*
Routine Description:
	The current group will be changed.

Parameters:
	None.

Return Value:
	None.
*/
{
	// Get current group.
	CString strName;
	pComboGroup->GetWindowText(strName);

	int nGroup = m_wndGrid.GetGroupFromName(strName);

	if (nGroup == INVALID)
	{
		HideControls();
		pComboGroup->SetCurSel(-1);
		return;
	}

	// Hide all controls default.
	HideControls(FALSE);

	// Update all controls in page.
	pAllowSizing->SetCheck(m_wndGrid.GetGroupAllowSizing(nGroup));
	pVisible->SetCheck(m_wndGrid.GetGroupVisible(nGroup));
	pName->SetWindowText(m_wndGrid.GetGroupName(nGroup));
	pTitle->SetWindowText(m_wndGrid.GetGroupTitle(nGroup));
	
	CString strWidth;
	strWidth.Format("%d", m_wndGrid.GetGroupWidth(nGroup));
	pWidth->SetWindowText(strWidth);

	CString strLevels;
	strLevels.Format("%d", m_wndGrid.GetLevelCount());
	pLevels->SetWindowText(strLevels);

	pDeleteGroup->EnableWindow(TRUE);
}

void CBHMDBGroupsPropPage::OnKillfocusEditTitle() 
/*
Routine Description:
	The title of current group will be changed.

Parameters:
	None.

Return Value:
	None.
*/
{
	// Get current group.
	CString strName;
	pComboGroup->GetWindowText(strName);

	int nGroup = m_wndGrid.GetGroupFromName(strName);

	if (nGroup == INVALID)
		return;

	// Update the grid.
	CString strTitle;
	pTitle->GetWindowText(strTitle);
	m_wndGrid.SetGroupTitle(nGroup, strTitle);
}

void CBHMDBGroupsPropPage::OnKillfocusEditWidth() 
/*
Routine Description:
	The width of current group will be changed.
	
Parameters:
	None.

Return Value:
	None.
*/
{
	// Get current group.
	CString strName;
	pComboGroup->GetWindowText(strName);

	int nGroup = m_wndGrid.GetGroupFromName(strName);

	if (nGroup == INVALID)
		return;

	// Update the grid.
	CString strWidth;
	pWidth->GetWindowText(strWidth);
	m_wndGrid.SetGroupWidth(nGroup, atoi(strWidth));
}

BOOL CBHMDBGroupsPropPage::OnInitDialog() 
{
	COlePropertyPage::OnInitDialog();
	
	// Init underlying control pointers.
	InitControlsVars();

	// Hide all controls in page default.
	HideControls();

	// Can not do anything without the pointer of underlying control.
	if (!m_pGridCtrl && !m_pDropDownCtrl && !m_pComboCtrl)
		return FALSE;

	// Update thte grid.
	m_wndGrid.ResetScrollBars();
	m_wndGrid.Invalidate();
	
	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

void CBHMDBGroupsPropPage::InitGrid()
/*
Routine Description:
	Init the grid.

Parameters:
	None.

Return Value:
	None.
*/
{
	if (IsWindow(m_wndGrid.m_hWnd))
	{
		m_wndGrid.LockUpdate(FALSE);
		m_wndGrid.SetReadOnly(TRUE);
	}
	else
		m_wndGrid.LockUpdate();

	// Reset the group combo box.
	pComboGroup->ResetContent();
		
	if (m_pGridCtrl)
	{
		// Load properties from grid control.
		m_wndGrid.LoadLayout(&m_pGridCtrl->m_arGroups, &m_pGridCtrl->m_arCols, &m_pGridCtrl->m_arCells);

		m_wndGrid.SetBackColor(m_pGridCtrl->TranslateColor(m_pGridCtrl->GetBackColor()));
		m_wndGrid.SetForeColor(m_pGridCtrl->TranslateColor(m_pGridCtrl->GetForeColor()));
		m_wndGrid.SetHeaderForeColor(m_pGridCtrl->TranslateColor(m_pGridCtrl->m_clrHeadFore));
		m_wndGrid.SetHeaderBackColor(m_pGridCtrl->TranslateColor(m_pGridCtrl->m_clrHeadBack));
		m_wndGrid.SetRowHeight(m_pGridCtrl->GetRowHeight());
		m_wndGrid.SetHeaderHeight(m_pGridCtrl->GetHeaderHeight());
	}
	else if (m_pDropDownCtrl)
	{
		// Load properties from dropdown control.
		m_wndGrid.LoadLayout(&m_pDropDownCtrl->m_arGroups, &m_pDropDownCtrl->m_arCols, &m_pDropDownCtrl->m_arCells);

		m_wndGrid.SetBackColor(m_pDropDownCtrl->TranslateColor(m_pDropDownCtrl->GetBackColor()));
		m_wndGrid.SetForeColor(m_pDropDownCtrl->TranslateColor(m_pDropDownCtrl->GetForeColor()));
		m_wndGrid.SetHeaderForeColor(m_pDropDownCtrl->TranslateColor(m_pDropDownCtrl->m_clrHeadFore));
		m_wndGrid.SetHeaderBackColor(m_pDropDownCtrl->TranslateColor(m_pDropDownCtrl->m_clrHeadBack));
		m_wndGrid.SetRowHeight(m_pDropDownCtrl->GetRowHeight());
		m_wndGrid.SetHeaderHeight(m_pDropDownCtrl->GetHeaderHeight());
	}
	else if (m_pComboCtrl)
	{
		// Load properties from combo control.
		m_wndGrid.LoadLayout(&m_pComboCtrl->m_arGroups, &m_pComboCtrl->m_arCols, &m_pComboCtrl->m_arCells);

		m_wndGrid.SetBackColor(m_pComboCtrl->TranslateColor(m_pComboCtrl->GetBackColor()));
		m_wndGrid.SetForeColor(m_pComboCtrl->TranslateColor(m_pComboCtrl->GetForeColor()));
		m_wndGrid.SetHeaderForeColor(m_pComboCtrl->TranslateColor(m_pComboCtrl->m_clrHeadFore));
		m_wndGrid.SetHeaderBackColor(m_pComboCtrl->TranslateColor(m_pComboCtrl->m_clrHeadBack));
		m_wndGrid.SetRowHeight(m_pComboCtrl->GetRowHeight());
		m_wndGrid.SetHeaderHeight(m_pComboCtrl->GetHeaderHeight());
	}

	// Insert each group name into group combo box.
	int nGroups = m_wndGrid.GetGroupCount();

	for (int i = 0; i < nGroups; i ++)
		pComboGroup->AddString(m_wndGrid.GetGroupName(i + 1));

	// Update other controls in page.
	if (nGroups == 0)
		pDeleteGroup->EnableWindow(FALSE);
	else
		pComboGroup->SetCurSel(0);
		
	if (IsWindow(m_wndGrid.m_hWnd))
	{
		m_wndGrid.ResetScrollBars();
		m_wndGrid.Invalidate();
	}

	OnSelendokComboGroup();
}

void CBHMDBGroupsPropPage::CheckRadio(CButton *pButton, int nValue)
/*
Routine Description:
	Check a radio in a radio button group.

Parameters:
	None.

Return Value:
	None.
*/
{
	CWnd * pCtrl = pButton;

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

void CBHMDBGroupsPropPage::HideControls(BOOL bHide)
/*
Routine Description:
	Show/hide all controls in page.

Parameters:
	None.

Return Value:
	None.
*/
{
	BOOL bShow = !bHide;

	pName->EnableWindow(bShow);
	pWidth->EnableWindow(bShow);
	pLevels->EnableWindow(bShow);
	pTitle->EnableWindow(bShow);
	pForeColor->EnableWindow(bShow);
	pBackColor->EnableWindow(bShow);
	pAllowSizing->EnableWindow(bShow);
	pVisible->EnableWindow(bShow);
}

void CBHMDBGroupsPropPage::InitControlsVars()
/*
Routine Description:
	Get the underlying control pointer.

Parameters:
	None.

Return Value:
	None.
*/
{
	// Attach grid object to target window.
	m_wndGrid.SubclassDlgItem(IDC_GRID, this);
	
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

		// Get the controltype property to decide the control type.
		HRESULT hr = m_lpDispatch[0]->Invoke(255, IID_NULL, LOCALE_USER_DEFAULT,
			DISPATCH_METHOD | DISPATCH_PROPERTYGET, &tParam, &va, NULL, NULL);
		if (SUCCEEDED(hr))
		{
			if (va.iVal == 0)
				m_pGridCtrl = (CBHMDBGridCtrl *) CCmdTarget::FromIDispatch(m_lpDispatch[0]);
			else if (va.iVal == 1)
				m_pDropDownCtrl = (CBHMDBDropDownCtrl *) CCmdTarget::FromIDispatch(m_lpDispatch[0]);
			else if (va.iVal == 2)
				m_pComboCtrl = (CBHMDBComboCtrl *) CCmdTarget::FromIDispatch(m_lpDispatch[0]);
		}
		VariantClear(&va);
	}

	// Can not do anything without underlying control pointer.
	if (!m_pGridCtrl && !m_pDropDownCtrl && !m_pComboCtrl)
	{
		pAddGroup->EnableWindow(FALSE);
		pDeleteGroup->EnableWindow(FALSE);
		pReset->EnableWindow(FALSE);
		pComboGroup->EnableWindow(FALSE);
		m_wndGrid.EnableWindow(FALSE);
		
		return;
	}
}

void CBHMDBGroupsPropPage::OnKillfocusEditName() 
/*
Routine Description:
	The name of current group will be changed.

Parameters:
	None.

Return Value:
	None.
*/
{
	// Get current group.
	CString strName;
	pComboGroup->GetWindowText(strName);

	int nGroup = m_wndGrid.GetGroupFromName(strName);

	if (nGroup == INVALID)
		return;

	// Update the grid.
	pName->GetWindowText(strName);

	m_wndGrid.SetGroupName(nGroup, strName);
}

void CBHMDBGroupsPropPage::OnKillfocusEditLevels() 
/*
Routine Description:
	The level count will be changed.

Parameters:
	None.

Return Value:
	None.
*/
{
	// Get the new level count.
	CString strLevels;
	pLevels->GetWindowText(strLevels);

	int nLevels = atoi(strLevels);

	if (nLevels <= 0 || nLevels > 255 || (nLevels > 1 && m_wndGrid.GetGroupCount() == 0))
		return;

	// Update the grid.
	m_wndGrid.SetLevelCount(nLevels);
}
