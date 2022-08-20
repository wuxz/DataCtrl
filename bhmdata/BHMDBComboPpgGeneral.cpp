// BHMDBComboPpgGeneral.cpp : Implementation of the CBHMDBComboPropPageGeneral property page class.

#include "stdafx.h"
#include "BHMData.h"
#include "BHMDBComboPpgGeneral.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


IMPLEMENT_DYNCREATE(CBHMDBComboPropPageGeneral, COlePropertyPage)


/////////////////////////////////////////////////////////////////////////////
// Message map

BEGIN_MESSAGE_MAP(CBHMDBComboPropPageGeneral, COlePropertyPage)
	//{{AFX_MSG_MAP(CBHMDBComboPropPageGeneral)
	// NOTE - ClassWizard will add and remove message map entries
	//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// Initialize class factory and guid

IMPLEMENT_OLECREATE_EX(CBHMDBComboPropPageGeneral, "BHMDATA.BHMDBComboGeneralPropPage.1",
	0xbf91b12e, 0x6a84, 0x11d3, 0xa7, 0xfe, 0, 0x80, 0xc8, 0x76, 0x3f, 0xa4)


/////////////////////////////////////////////////////////////////////////////
// CBHMDBComboPropPageGeneral::CBHMDBComboPropPageGeneralFactory::UpdateRegistry -
// Adds or removes system registry entries for CBHMDBComboPropPageGeneral

BOOL CBHMDBComboPropPageGeneral::CBHMDBComboPropPageGeneralFactory::UpdateRegistry(BOOL bRegister)
{
	if (bRegister)
		return AfxOleRegisterPropertyPageClass(AfxGetInstanceHandle(),
			m_clsid, IDS_BHMDBCOMBO_PPG);
	else
		return AfxOleUnregisterClass(m_clsid, NULL);
}


/////////////////////////////////////////////////////////////////////////////
// CBHMDBComboPropPageGeneral::CBHMDBComboPropPageGeneral - Constructor

CBHMDBComboPropPageGeneral::CBHMDBComboPropPageGeneral() :
	COlePropertyPage(IDD, IDS_BHMDBCOMBO_PPG_CAPTION)
{
	//{{AFX_DATA_INIT(CBHMDBComboPropPageGeneral)
	m_bReadOnly = FALSE;
	m_bColumnHeader = FALSE;
	m_bListAutoPosition = FALSE;
	m_bListWidthAutoSize = FALSE;
	m_nDividerStyle = -1;
	m_nDividerType = -1;
	m_nBorderStyle = -1;
	m_nDataMode = -1;
	m_strFieldSeparator = _T("");
	m_nRows = 0;
	m_nCols = 0;
	m_n = 0;
	m_nMinDropDownItems = 0;
	m_nDefColWidth = 0;
	m_nHeaderHeight = 0;
	m_nRowHeight = 0;
	m_nListWidth = 0;
	m_strFormat = _T("");
	m_nMaxLength = 0;
	m_nStyle = -1;
	m_bGroupHeader = FALSE;
	//}}AFX_DATA_INIT

	SetHelpInfo(_T("Names to appear in the control"), _T("BHMDATA.HLP"), 0);
}


/////////////////////////////////////////////////////////////////////////////
// CBHMDBComboPropPageGeneral::DoDataExchange - Moves data between page and properties

void CBHMDBComboPropPageGeneral::DoDataExchange(CDataExchange* pDX)
{
	//{{AFX_DATA_MAP(CBHMDBComboPropPageGeneral)
	DDP_Check(pDX, IDC_CHECK_READONLY, m_bReadOnly, _T("ReadOnly") );
	DDX_Check(pDX, IDC_CHECK_READONLY, m_bReadOnly);
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
	DDP_CBIndex(pDX, IDC_COMBO_DATAMODE, m_nDataMode, _T("DataMode") );
	DDX_CBIndex(pDX, IDC_COMBO_DATAMODE, m_nDataMode);
	DDP_Text(pDX, IDC_EDIT_FIELDSEPARATOR, m_strFieldSeparator, _T("FieldSeparator") );
	DDX_Text(pDX, IDC_EDIT_FIELDSEPARATOR, m_strFieldSeparator);
	DDP_Text(pDX, IDC_EDIT_ROWS, m_nRows, _T("Rows") );
	DDX_Text(pDX, IDC_EDIT_ROWS, m_nRows);
	DDP_Text(pDX, IDC_EDIT_COLS, m_nCols, _T("Cols") );
	DDX_Text(pDX, IDC_EDIT_COLS, m_nCols);
	DDP_Text(pDX, IDC_EDIT_MAXDROPDOWNITEMS, m_n, _T("MaxDropDownItems") );
	DDX_Text(pDX, IDC_EDIT_MAXDROPDOWNITEMS, m_n);
	DDP_Text(pDX, IDC_EDIT_MINDROPDOWNITEMS, m_nMinDropDownItems, _T("MinDropDownItems") );
	DDX_Text(pDX, IDC_EDIT_MINDROPDOWNITEMS, m_nMinDropDownItems);
	DDP_Text(pDX, IDC_EDIT_DEFCOLWIDTH, m_nDefColWidth, _T("DefColWidth") );
	DDX_Text(pDX, IDC_EDIT_DEFCOLWIDTH, m_nDefColWidth);
	DDP_Text(pDX, IDC_EDIT_HEADERHEIGHT, m_nHeaderHeight, _T("HeaderHeight") );
	DDX_Text(pDX, IDC_EDIT_HEADERHEIGHT, m_nHeaderHeight);
	DDP_Text(pDX, IDC_EDIT_ROWHEIGHT, m_nRowHeight, _T("RowHeight") );
	DDX_Text(pDX, IDC_EDIT_ROWHEIGHT, m_nRowHeight);
	DDP_Text(pDX, IDC_EDIT_LISTWIDTH, m_nListWidth, _T("ListWidth") );
	DDX_Text(pDX, IDC_EDIT_LISTWIDTH, m_nListWidth);
	DDP_Text(pDX, IDC_EDIT_FORMAT, m_strFormat, _T("Format") );
	DDX_Text(pDX, IDC_EDIT_FORMAT, m_strFormat);
	DDP_Text(pDX, IDC_EDIT_MAXLENGTH, m_nMaxLength, _T("MaxLength") );
	DDX_Text(pDX, IDC_EDIT_MAXLENGTH, m_nMaxLength);
	DDP_CBIndex(pDX, IDC_COMBO_STYLE, m_nStyle, _T("Style") );
	DDX_CBIndex(pDX, IDC_COMBO_STYLE, m_nStyle);
	DDP_Check(pDX, IDC_CHECK_GROUPHEADER, m_bGroupHeader, _T("GroupHeader") );
	DDX_Check(pDX, IDC_CHECK_GROUPHEADER, m_bGroupHeader);
	//}}AFX_DATA_MAP
	DDP_PostProcessing(pDX);
}


/////////////////////////////////////////////////////////////////////////////
// CBHMDBComboPropPageGeneral message handlers
