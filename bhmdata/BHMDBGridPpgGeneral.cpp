// BHMDBGridPpgGeneral.cpp : Implementation of the CBHMDBGridPropPageGeneral property page class.

#include "stdafx.h"
#include "BHMData.h"
#include "BHMDBGridPpgGeneral.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


IMPLEMENT_DYNCREATE(CBHMDBGridPropPageGeneral, COlePropertyPage)


/////////////////////////////////////////////////////////////////////////////
// Message map

BEGIN_MESSAGE_MAP(CBHMDBGridPropPageGeneral, COlePropertyPage)
	//{{AFX_MSG_MAP(CBHMDBGridPropPageGeneral)
	ON_CBN_SELENDOK(IDC_COMBO_DATAMODE, OnSelendokComboDatamode)
	ON_BN_CLICKED(IDC_CHECK_ALLOWUPDATE, OnCheckAllowupdate)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// Initialize class factory and guid

IMPLEMENT_OLECREATE_EX(CBHMDBGridPropPageGeneral, "BHMDATA.BHMDBGridGeneralPropPage.1",
	0xbf91b126, 0x6a84, 0x11d3, 0xa7, 0xfe, 0, 0x80, 0xc8, 0x76, 0x3f, 0xa4)


/////////////////////////////////////////////////////////////////////////////
// CBHMDBGridPropPageGeneral::CBHMDBGridPropPageGeneralFactory::UpdateRegistry -
// Adds or removes system registry entries for CBHMDBGridPropPageGeneral

BOOL CBHMDBGridPropPageGeneral::CBHMDBGridPropPageGeneralFactory::UpdateRegistry(BOOL bRegister)
{
	if (bRegister)
		return AfxOleRegisterPropertyPageClass(AfxGetInstanceHandle(),
			m_clsid, IDS_BHMDBGRID_PPG);
	else
		return AfxOleUnregisterClass(m_clsid, NULL);
}


/////////////////////////////////////////////////////////////////////////////
// CBHMDBGridPropPageGeneral::CBHMDBGridPropPageGeneral - Constructor

CBHMDBGridPropPageGeneral::CBHMDBGridPropPageGeneral() :
	COlePropertyPage(IDD, IDS_BHMDBGRID_PPG_CAPTION)
{
	//{{AFX_DATA_INIT(CBHMDBGridPropPageGeneral)
	m_nDataMode = -1;
	m_bAllowAddNew = FALSE;
	m_bAllowDelete = FALSE;
	m_bAllowUpdate = FALSE;
	m_bAllowRowSizing = FALSE;
	m_bAllowColMoving = FALSE;
	m_bAllowColSizing = FALSE;
	m_bColumnHeader = FALSE;
	m_bRecordSelectors = FALSE;
	m_bSelectByCell = FALSE;
	m_strCaption = _T("");
	m_strFieldSeparator = _T("");
	m_nBorderStyle = -1;
	m_nCaptionAlignment = -1;
	m_nDividerStyle = -1;
	m_nDividerType = -1;
	m_nCols = 0;
	m_nDefColWidth = 0;
	m_nHeaderHeight = 0;
	m_nRowHeight = 0;
	m_nRows = 0;
	m_bGroupHeader = FALSE;
	//}}AFX_DATA_INIT

	SetHelpInfo(_T("Names to appear in the control"), _T("BHMDATA.HLP"), 0);
}


/////////////////////////////////////////////////////////////////////////////
// CBHMDBGridPropPageGeneral::DoDataExchange - Moves data between page and properties

void CBHMDBGridPropPageGeneral::DoDataExchange(CDataExchange* pDX)
{
	//{{AFX_DATA_MAP(CBHMDBGridPropPageGeneral)
	DDP_CBIndex(pDX, IDC_COMBO_DATAMODE, m_nDataMode, _T("DataMode") );
	DDX_CBIndex(pDX, IDC_COMBO_DATAMODE, m_nDataMode);
	DDP_Check(pDX, IDC_CHECK_ALLOWADDNEW, m_bAllowAddNew, _T("AllowAddNew") );
	DDX_Check(pDX, IDC_CHECK_ALLOWADDNEW, m_bAllowAddNew);
	DDP_Check(pDX, IDC_CHECK_ALLOWDELETE, m_bAllowDelete, _T("AllowDelete") );
	DDX_Check(pDX, IDC_CHECK_ALLOWDELETE, m_bAllowDelete);
	DDP_Check(pDX, IDC_CHECK_ALLOWUPDATE, m_bAllowUpdate, _T("AllowUpdate") );
	DDX_Check(pDX, IDC_CHECK_ALLOWUPDATE, m_bAllowUpdate);
	DDP_Check(pDX, IDC_CHECK_ALLOWROWRESIZING, m_bAllowRowSizing, _T("AllowRowSizing") );
	DDX_Check(pDX, IDC_CHECK_ALLOWROWRESIZING, m_bAllowRowSizing);
	DDP_Check(pDX, IDC_CHECK_ALLOWCOLMOVING, m_bAllowColMoving, _T("AllowColumnMoving") );
	DDX_Check(pDX, IDC_CHECK_ALLOWCOLMOVING, m_bAllowColMoving);
	DDP_Check(pDX, IDC_CHECK_ALLOWCOLSIZING, m_bAllowColSizing, _T("AllowColumnSizing") );
	DDX_Check(pDX, IDC_CHECK_ALLOWCOLSIZING, m_bAllowColSizing);
	DDP_Check(pDX, IDC_CHECK_COLUMNHEADER, m_bColumnHeader, _T("ColumnHeader") );
	DDX_Check(pDX, IDC_CHECK_COLUMNHEADER, m_bColumnHeader);
	DDP_Check(pDX, IDC_CHECK_SHOWRECORDSELECTORS, m_bRecordSelectors, _T("RecordSelectors") );
	DDX_Check(pDX, IDC_CHECK_SHOWRECORDSELECTORS, m_bRecordSelectors);
	DDP_Check(pDX, IDC_CHECK_SELECTBYCELL, m_bSelectByCell, _T("SelectByCell") );
	DDX_Check(pDX, IDC_CHECK_SELECTBYCELL, m_bSelectByCell);
	DDP_Text(pDX, IDC_EDIT_CAPTION, m_strCaption, _T("Caption") );
	DDX_Text(pDX, IDC_EDIT_CAPTION, m_strCaption);
	DDP_Text(pDX, IDC_EDIT_FIELDSEPARATOR, m_strFieldSeparator, _T("FieldSeparator") );
	DDX_Text(pDX, IDC_EDIT_FIELDSEPARATOR, m_strFieldSeparator);
	DDP_CBIndex(pDX, IDC_COMBO_BORDERSTYLE, m_nBorderStyle, _T("BorderStyle") );
	DDX_CBIndex(pDX, IDC_COMBO_BORDERSTYLE, m_nBorderStyle);
	DDP_CBIndex(pDX, IDC_COMBO_CAPTIONALIGNMENT, m_nCaptionAlignment, _T("CaptionAlignment") );
	DDX_CBIndex(pDX, IDC_COMBO_CAPTIONALIGNMENT, m_nCaptionAlignment);
	DDP_CBIndex(pDX, IDC_COMBO_DIVIDERSTYLE, m_nDividerStyle, _T("DividerStyle") );
	DDX_CBIndex(pDX, IDC_COMBO_DIVIDERSTYLE, m_nDividerStyle);
	DDP_CBIndex(pDX, IDC_COMBO_DIVIDERTYPE, m_nDividerType, _T("DividerType") );
	DDX_CBIndex(pDX, IDC_COMBO_DIVIDERTYPE, m_nDividerType);
	DDP_Text(pDX, IDC_EDIT_COLS, m_nCols, _T("Cols") );
	DDX_Text(pDX, IDC_EDIT_COLS, m_nCols);
	DDV_MinMaxLong(pDX, m_nCols, 0, 255);
	DDP_Text(pDX, IDC_EDIT_DEFCOLWIDTH, m_nDefColWidth, _T("DefColWidth") );
	DDX_Text(pDX, IDC_EDIT_DEFCOLWIDTH, m_nDefColWidth);
	DDP_Text(pDX, IDC_EDIT_HEADERHEIGHT, m_nHeaderHeight, _T("HeaderHeight") );
	DDX_Text(pDX, IDC_EDIT_HEADERHEIGHT, m_nHeaderHeight);
	DDP_Text(pDX, IDC_EDIT_ROWHEIGHT, m_nRowHeight, _T("RowHeight") );
	DDX_Text(pDX, IDC_EDIT_ROWHEIGHT, m_nRowHeight);
	DDP_Text(pDX, IDC_EDIT_ROWS, m_nRows, _T("Rows") );
	DDX_Text(pDX, IDC_EDIT_ROWS, m_nRows);
	DDP_Check(pDX, IDC_CHECK_GROUPHEADER, m_bGroupHeader, _T("GroupHeader") );
	DDX_Check(pDX, IDC_CHECK_GROUPHEADER, m_bGroupHeader);
	//}}AFX_DATA_MAP
	DDP_PostProcessing(pDX);

	if (!pDX->m_bSaveAndValidate)
		UpdateControls();
}


/////////////////////////////////////////////////////////////////////////////
// CBHMDBGridPropPageGeneral message handlers
void CBHMDBGridPropPageGeneral::OnSelendokComboDatamode() 
{
	m_nDataMode = ((CComboBox *)GetDlgItem(IDC_COMBO_DATAMODE))->GetCurSel();
	UpdateControls();
}

void CBHMDBGridPropPageGeneral::OnCheckAllowupdate() 
{
	m_bAllowUpdate = ((CButton *)GetDlgItem(IDC_CHECK_ALLOWUPDATE))->GetCheck();
	UpdateControls();
}

void CBHMDBGridPropPageGeneral::UpdateControls()
{
	if (m_nDataMode == 0)
	{
		GetDlgItem(IDC_SPIN_ROWS)->EnableWindow(FALSE);
		GetDlgItem(IDC_SPIN_COLS)->EnableWindow(FALSE);
		GetDlgItem(IDC_EDIT_ROWS)->EnableWindow(FALSE);
		GetDlgItem(IDC_EDIT_COLS)->EnableWindow(FALSE);
	}
	else
	{
		GetDlgItem(IDC_SPIN_ROWS)->EnableWindow(TRUE);
		GetDlgItem(IDC_SPIN_COLS)->EnableWindow(TRUE);
		GetDlgItem(IDC_EDIT_ROWS)->EnableWindow(TRUE);
		GetDlgItem(IDC_EDIT_COLS)->EnableWindow(TRUE);
	}

	if (!m_bAllowUpdate)
	{
		((CButton *)GetDlgItem(IDC_CHECK_ALLOWADDNEW))->SetCheck(0);
		((CButton *)GetDlgItem(IDC_CHECK_ALLOWDELETE))->SetCheck(0);
		GetDlgItem(IDC_CHECK_ALLOWADDNEW)->EnableWindow(FALSE);
		GetDlgItem(IDC_CHECK_ALLOWDELETE)->EnableWindow(FALSE);
	}
	else
	{
		GetDlgItem(IDC_CHECK_ALLOWADDNEW)->EnableWindow(TRUE);
		GetDlgItem(IDC_CHECK_ALLOWDELETE)->EnableWindow(TRUE);
	}
}
