// BHMDBDropDownPpgGeneral.cpp : Implementation of the CBHMDBDropDownPropPageGeneral property page class.

#include "stdafx.h"
#include "BHMData.h"
#include "BHMDBDropDownPpgGeneral.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


IMPLEMENT_DYNCREATE(CBHMDBDropDownPropPageGeneral, COlePropertyPage)


/////////////////////////////////////////////////////////////////////////////
// Message map

BEGIN_MESSAGE_MAP(CBHMDBDropDownPropPageGeneral, COlePropertyPage)
	//{{AFX_MSG_MAP(CBHMDBDropDownPropPageGeneral)
	// NOTE - ClassWizard will add and remove message map entries
	//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// Initialize class factory and guid

IMPLEMENT_OLECREATE_EX(CBHMDBDropDownPropPageGeneral, "BHMDATA.BHMDBDropDownGeneralPropPage.1",
	0xbf91b12a, 0x6a84, 0x11d3, 0xa7, 0xfe, 0, 0x80, 0xc8, 0x76, 0x3f, 0xa4)


/////////////////////////////////////////////////////////////////////////////
// CBHMDBDropDownPropPageGeneral::CBHMDBDropDownPropPageGeneralFactory::UpdateRegistry -
// Adds or removes system registry entries for CBHMDBDropDownPropPageGeneral

BOOL CBHMDBDropDownPropPageGeneral::CBHMDBDropDownPropPageGeneralFactory::UpdateRegistry(BOOL bRegister)
{
	if (bRegister)
		return AfxOleRegisterPropertyPageClass(AfxGetInstanceHandle(),
			m_clsid, IDS_BHMDBDROPDOWN_PPG);
	else
		return AfxOleUnregisterClass(m_clsid, NULL);
}


/////////////////////////////////////////////////////////////////////////////
// CBHMDBDropDownPropPageGeneral::CBHMDBDropDownPropPageGeneral - Constructor

CBHMDBDropDownPropPageGeneral::CBHMDBDropDownPropPageGeneral() :
	COlePropertyPage(IDD, IDS_BHMDBDROPDOWN_PPG_CAPTION)
{
	//{{AFX_DATA_INIT(CBHMDBDropDownPropPageGeneral)
	m_nDataMode = -1;
	m_bColumnHeader = FALSE;
	m_bListAutoPosition = FALSE;
	m_bListWidthAutoSize = FALSE;
	m_nDividerStyle = -1;
	m_nDividerType = -1;
	m_nBorderStyle = -1;
	m_strFieldSeparator = _T("");
	m_nRows = 0;
	m_nCols = 0;
	m_nDefColWidth = 0;
	m_nHeaderHeight = 0;
	m_nRowHeight = 0;
	m_nMaxDropDownItems = 0;
	m_nMinDropDownItems = 0;
	m_nListWidth = 0;
	m_bGroupHeader = FALSE;
	//}}AFX_DATA_INIT

	SetHelpInfo(_T("Names to appear in the control"), _T("BHMDATA.HLP"), 0);
}


/////////////////////////////////////////////////////////////////////////////
// CBHMDBDropDownPropPageGeneral::DoDataExchange - Moves data between page and properties

void CBHMDBDropDownPropPageGeneral::DoDataExchange(CDataExchange* pDX)
{
	//{{AFX_DATA_MAP(CBHMDBDropDownPropPageGeneral)
	DDP_CBIndex(pDX, IDC_COMBO_DATAMODE, m_nDataMode, _T("DataMode") );
	DDX_CBIndex(pDX, IDC_COMBO_DATAMODE, m_nDataMode);
	DDP_Check(pDX, IDC_CHECK_COLUMNHEADER, m_bColumnHeader, _T("ColumnHeader") );
	DDX_Check(pDX, IDC_CHECK_COLUMNHEADER, m_bColumnHeader);
	DDP_Check(pDX, IDC_CHECK_LISTAUTOPOSITION, m_bListAutoPosition, _T("ListAutoPosition") );
	DDX_Check(pDX, IDC_CHECK_LISTAUTOPOSITION, m_bListAutoPosition);
	DDP_Check(pDX, IDC_CHECK_LISTWIDTHAUTOSIZE, m_bListWidthAutoSize, _T("ListWidthAutoSize") );
	DDX_Check(pDX, IDC_CHECK_LISTWIDTHAUTOSIZE, m_bListWidthAutoSize);
	DDP_CBIndex(pDX, IDC_COMBO_DIVIDERSTYLE, m_nDividerStyle, _T("DividerStyle") );
	DDX_CBIndex(pDX, IDC_COMBO_DIVIDERSTYLE, m_nDividerStyle);
	DDP_CBIndex(pDX, IDC_COMBO_DIVIDERTYPE, m_nDividerType, _T("DividerType") );
	DDX_CBIndex(pDX, IDC_COMBO_DIVIDERTYPE, m_nDividerType);
	DDP_CBIndex(pDX, IDC_COMBO_BORDERSTYLE, m_nBorderStyle, _T("BorderStyle") );
	DDX_CBIndex(pDX, IDC_COMBO_BORDERSTYLE, m_nBorderStyle);
	DDP_Text(pDX, IDC_EDIT_FIELDSEPARATOR, m_strFieldSeparator, _T("FieldSeparator") );
	DDX_Text(pDX, IDC_EDIT_FIELDSEPARATOR, m_strFieldSeparator);
	DDP_Text(pDX, IDC_EDIT_ROWS, m_nRows, _T("Rows") );
	DDX_Text(pDX, IDC_EDIT_ROWS, m_nRows);
	DDP_Text(pDX, IDC_EDIT_COLS, m_nCols, _T("Cols") );
	DDX_Text(pDX, IDC_EDIT_COLS, m_nCols);
	DDV_MinMaxLong(pDX, m_nCols, 0, 255);
	DDP_Text(pDX, IDC_EDIT_DEFCOLWIDTH, m_nDefColWidth, _T("DefColWidth") );
	DDX_Text(pDX, IDC_EDIT_DEFCOLWIDTH, m_nDefColWidth);
	DDP_Text(pDX, IDC_EDIT_HEADERHEIGHT, m_nHeaderHeight, _T("HeaderHeight") );
	DDX_Text(pDX, IDC_EDIT_HEADERHEIGHT, m_nHeaderHeight);
	DDP_Text(pDX, IDC_EDIT_ROWHEIGHT, m_nRowHeight, _T("RowHeight") );
	DDX_Text(pDX, IDC_EDIT_ROWHEIGHT, m_nRowHeight);
	DDP_Text(pDX, IDC_EDIT_MAXDROPDOWNITEMS, m_nMaxDropDownItems, _T("MaxDropDownItems") );
	DDX_Text(pDX, IDC_EDIT_MAXDROPDOWNITEMS, m_nMaxDropDownItems);
	DDV_MinMaxLong(pDX, m_nMaxDropDownItems, 1, 100);
	DDP_Text(pDX, IDC_EDIT_MINDROPDOWNITEMS, m_nMinDropDownItems, _T("MinDropDownItems") );
	DDX_Text(pDX, IDC_EDIT_MINDROPDOWNITEMS, m_nMinDropDownItems);
	DDV_MinMaxLong(pDX, m_nMinDropDownItems, 1, 100);
	DDP_Text(pDX, IDC_EDIT_LISTWIDTH, m_nListWidth, _T("ListWidth") );
	DDX_Text(pDX, IDC_EDIT_LISTWIDTH, m_nListWidth);
	DDP_Check(pDX, IDC_CHECK_GROUPHEADER, m_bGroupHeader, _T("GroupHeader") );
	DDX_Check(pDX, IDC_CHECK_GROUPHEADER, m_bGroupHeader);
	//}}AFX_DATA_MAP
	DDP_PostProcessing(pDX);
}


/////////////////////////////////////////////////////////////////////////////
// CBHMDBDropDownPropPageGeneral message handlers
