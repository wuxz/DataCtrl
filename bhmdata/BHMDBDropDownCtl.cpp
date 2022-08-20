// BHMDBDropDownCtl.cpp : Implementation of the CBHMDBDropDownCtrl ActiveX Control class.

#include "stdafx.h"
#include "BHMData.h"
#include "BHMDBDropDownCtl.h"
#include "BHMDBDropDownPpgGeneral.h"
#include <oledberr.h>
#include <msstkppg.h>
#include "Columns.h"
#include "DropDownRealWnd.h"
#include "GridInner.h"
#include "BHMDBColumnPropPage.h"
#include "BHMDBGroupsPropPage.h"
#include "Groups.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


IMPLEMENT_DYNCREATE(CBHMDBDropDownCtrl, COleControl)


BEGIN_INTERFACE_MAP(CBHMDBDropDownCtrl, COleControl)
	// Map the interfaces to their implementations
	INTERFACE_PART(CBHMDBDropDownCtrl, IID_IRowsetNotify, RowsetNotify)
END_INTERFACE_MAP()

STDMETHODIMP_(ULONG) CBHMDBDropDownCtrl::XRowsetNotify::AddRef(void)
{
	METHOD_PROLOGUE(CBHMDBDropDownCtrl, RowsetNotify)
	return pThis->ExternalAddRef();
}

STDMETHODIMP_(ULONG) CBHMDBDropDownCtrl::XRowsetNotify::Release(void)
{
	METHOD_PROLOGUE(CBHMDBDropDownCtrl, RowsetNotify)
	return pThis->ExternalRelease();
}

STDMETHODIMP CBHMDBDropDownCtrl::XRowsetNotify::QueryInterface(REFIID iid,
	LPVOID FAR* ppvObject)
{
	METHOD_PROLOGUE(CBHMDBDropDownCtrl, RowsetNotify)
	return pThis->ExternalQueryInterface(&iid, ppvObject);
}

STDMETHODIMP CBHMDBDropDownCtrl::XRowsetNotify::OnFieldChange(IRowset * prowset,
	HROW hRow, ULONG cColumns, ULONG rgColumns[], DBREASON eReason, DBEVENTPHASE ePhase, BOOL fCantDeny)
{
	METHOD_PROLOGUE(CBHMDBDropDownCtrl, RowsetNotify)

	if (DBEVENTPHASE_DIDEVENT == ePhase && prowset == pThis->m_Rowset.m_spRowset)
	{
		switch (eReason)
		{
		case DBREASON_COLUMN_SET:
		case DBREASON_COLUMN_RECALCULATED:
		{
			// Some columns have been modified.

			int nRow;

			// Get row ordinal from hrow
			nRow = pThis->GetRowFromHRow(hRow);
			if (nRow == INVALID || nRow > (int)pThis->m_pDropDownRealWnd->OnGetRecordCount())
			{
				// This is a new record, fetch it :)
				pThis->m_bEndReached = FALSE;
				pThis->FetchNewRows();
				return S_OK;
			}

			// Redraw the correspond cells.
			for (ULONG i = 0; i < cColumns; i ++)
				pThis->m_pDropDownRealWnd->RedrawCell(nRow, rgColumns[i]);
		}

		return S_OK;
		break;

		default:
			return DB_S_UNWANTEDREASON;
		}
	}

	return DB_S_UNWANTEDPHASE;
}

STDMETHODIMP CBHMDBDropDownCtrl::XRowsetNotify::OnRowChange(IRowset * prowset,
	ULONG cRows, const HROW rghRows[], DBREASON eReason, DBEVENTPHASE ePhase, BOOL fCantDeny)
{
	METHOD_PROLOGUE(CBHMDBDropDownCtrl, RowsetNotify)

	
	if (DBEVENTPHASE_ABOUTTODO == ePhase && prowset == pThis->m_Rowset.m_spRowset)
	{	
		switch (eReason)
		{
		case DBREASON_ROW_DELETE:
		case DBREASON_ROW_UNDOINSERT:
		{
			// Some rows have been deleted.
			int nRow;

			for (ULONG i = 0; i < cRows; i ++)
			{
				// Calc the row ordinal from hrow.
				nRow = pThis->GetRowFromHRow(rghRows[i]);

				// Update the grid.
				if (nRow != INVALID)
					pThis->DeleteRowFromSrc(nRow);
			}
		}

		default:
			return S_OK;
		}
	}

	if (DBEVENTPHASE_DIDEVENT == ePhase && prowset == pThis->m_Rowset.m_spRowset)
	{
		switch (eReason)
		{
		case DBREASON_ROW_ASYNCHINSERT:
		case DBREASON_ROW_INSERT:
		{
			// Some rows have been inserted.

			pThis->m_bEndReached = FALSE;
			pThis->FetchNewRows();
		}

		return S_OK;
		break;

		case DBREASON_ROW_UNDODELETE:
		{
			// Some records have been undeleted.

			// Restore the original rows.
			for (ULONG i = 0; i < cRows; i ++)
				pThis->UndeleteRecordFromHRow(rghRows[i]);
		}

		return S_OK;
		break;

		case DBREASON_ROW_UNDOCHANGE:
		case DBREASON_ROW_UPDATE:
		{
			// Some rows have been updated.

			int nRow;
			BOOL bUpdated = FALSE;
			for (ULONG i = 0; i < cRows; i ++)
			{
				// Calc the row ordinal.
				nRow = pThis->GetRowFromHRow(rghRows[i]);
				if (nRow != INVALID)
				{
					// Redraw the row.
					pThis->m_pDropDownRealWnd->RedrawRow(nRow);
					bUpdated = TRUE;
				}
			}
			if (!bUpdated)
			{
				// Some rows have been updated but I do not have them in hand,
				// so there must be some new rows to fetch :)

				// We have not reached the end of the rowset.
				pThis->m_bEndReached = FALSE;

				// Fetch new rows now.
				pThis->FetchNewRows();
			}

		}

		return S_OK;
		break;

		default:
			return S_OK;
		}
	}

	return DB_S_UNWANTEDPHASE;
}

STDMETHODIMP CBHMDBDropDownCtrl::XRowsetNotify::OnRowsetChange(IRowset * prowset,
	DBREASON eReason, DBEVENTPHASE ePhase, BOOL fCantDeny)
{
	METHOD_PROLOGUE(CBHMDBDropDownCtrl, RowsetNotify)

	if (DBEVENTPHASE_DIDEVENT == ePhase && prowset == pThis->m_Rowset.m_spRowset)
	{
		switch (eReason)
		{
		case DBREASON_ROWSET_RELEASE:
		case DBREASON_ROWSET_CHANGED:
			// Rowset has been changed.
			pThis->m_bDataSrcChanged = TRUE;

		return S_OK;
		break;
			
		default:
			return DB_S_UNWANTEDREASON;
		}
	}

	return DB_S_UNWANTEDPHASE;
}

STDMETHODIMP_(ULONG) CBHMDBDropDownCtrl::XDataSourceListener::AddRef(void)
{
	METHOD_PROLOGUE(CBHMDBDropDownCtrl, DataSourceListener)
	return pThis->ExternalAddRef();
}

STDMETHODIMP_(ULONG) CBHMDBDropDownCtrl::XDataSourceListener::Release(void)
{
	METHOD_PROLOGUE(CBHMDBDropDownCtrl, DataSourceListener)
	return pThis->ExternalRelease();
}

STDMETHODIMP CBHMDBDropDownCtrl::XDataSourceListener::QueryInterface(REFIID iid,
	LPVOID FAR* ppvObject)
{
	METHOD_PROLOGUE(CBHMDBDropDownCtrl, DataSourceListener)
	return pThis->ExternalQueryInterface(&iid, ppvObject);
}

STDMETHODIMP CBHMDBDropDownCtrl::XDataSourceListener::dataMemberChanged(BSTR bstrDM)
{
	METHOD_PROLOGUE(CBHMDBDropDownCtrl, RowsetNotify)

	USES_CONVERSION;
	if ((pThis->m_strDataMember.IsEmpty() && (bstrDM == NULL || bstrDM[0] == '\0')) || pThis->m_strDataMember == OLE2T(bstrDM))
	{
		// Our data source has been changed.

		// Clear the cached interface pointers.
		pThis->ClearInterfaces();

		// The data source has been changed.
		pThis->m_bDataSrcChanged = TRUE;

		// We have not reached the end of the rowset.
		pThis->m_bEndReached = FALSE;
	}
	
	return S_OK;
}

STDMETHODIMP CBHMDBDropDownCtrl::XDataSourceListener::dataMemberAdded(BSTR bstrDM)
{
	return S_OK;
}

STDMETHODIMP CBHMDBDropDownCtrl::XDataSourceListener::dataMemberRemoved(BSTR bstrDM)
{
	METHOD_PROLOGUE(CBHMDBDropDownCtrl, RowsetNotify)

	USES_CONVERSION;
	if ((pThis->m_strDataMember.IsEmpty() && (bstrDM == NULL || bstrDM[0] == '\0')) || pThis->m_strDataMember == OLE2T(bstrDM))
	{
		// Our data source has been changed.

		// Clear the cached interface pointers.
		pThis->ClearInterfaces();
		pThis->m_bDataSrcChanged = TRUE;
	}

	return S_OK;
}
/////////////////////////////////////////////////////////////////////////////
// Message map

BEGIN_MESSAGE_MAP(CBHMDBDropDownCtrl, COleControl)
	//{{AFX_MSG_MAP(CBHMDBDropDownCtrl)
	ON_WM_CREATE()
	ON_WM_KEYDOWN()
	//}}AFX_MSG_MAP
	ON_OLEVERB(AFX_IDS_VERB_PROPERTIES, OnProperties)
	ON_REGISTERED_MESSAGE(CBHMDataApp::m_nBringMsg, OnBringMsg)
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// Dispatch map

BEGIN_DISPATCH_MAP(CBHMDBDropDownCtrl, COleControl)
	//{{AFX_DISPATCH_MAP(CBHMDBDropDownCtrl)
	DISP_PROPERTY_NOTIFY(CBHMDBDropDownCtrl, "HeadBackColor", m_clrHeadBack, OnHeadBackColorChanged, VT_COLOR)
	DISP_PROPERTY_NOTIFY(CBHMDBDropDownCtrl, "HeadForeColor", m_clrHeadFore, OnHeadForeColorChanged, VT_COLOR)
	DISP_PROPERTY_NOTIFY(CBHMDBDropDownCtrl, "ListAutoPosition", m_bListAutoPosition, OnListAutoPositionChanged, VT_BOOL)
	DISP_PROPERTY_NOTIFY(CBHMDBDropDownCtrl, "ListWidthAutoSize", m_bListWidthAutoSize, OnListWidthAutoSizeChanged, VT_BOOL)
	DISP_PROPERTY_NOTIFY(CBHMDBDropDownCtrl, "DividerColor", m_clrDivider, OnDividerColorChanged, VT_COLOR)
	DISP_PROPERTY_EX(CBHMDBDropDownCtrl, "Font", GetFont, SetFont, VT_FONT)
	DISP_PROPERTY_EX(CBHMDBDropDownCtrl, "DefColWidth", GetDefColWidth, SetDefColWidth, VT_I4)
	DISP_PROPERTY_EX(CBHMDBDropDownCtrl, "RowHeight", GetRowHeight, SetRowHeight, VT_I4)
	DISP_PROPERTY_EX(CBHMDBDropDownCtrl, "DividerType", GetDividerType, SetDividerType, VT_I4)
	DISP_PROPERTY_EX(CBHMDBDropDownCtrl, "DividerStyle", GetDividerStyle, SetDividerStyle, VT_I4)
	DISP_PROPERTY_EX(CBHMDBDropDownCtrl, "DataMode", GetDataMode, SetDataMode, VT_I4)
	DISP_PROPERTY_EX(CBHMDBDropDownCtrl, "LeftCol", GetLeftCol, SetLeftCol, VT_I4)
	DISP_PROPERTY_EX(CBHMDBDropDownCtrl, "FirstRow", GetFirstRow, SetFirstRow, VT_I4)
	DISP_PROPERTY_EX(CBHMDBDropDownCtrl, "Row", GetRow, SetRow, VT_I4)
	DISP_PROPERTY_EX(CBHMDBDropDownCtrl, "Col", GetCol, SetCol, VT_I4)
	DISP_PROPERTY_EX(CBHMDBDropDownCtrl, "Rows", GetRows, SetRows, VT_I4)
	DISP_PROPERTY_EX(CBHMDBDropDownCtrl, "Cols", GetCols, SetCols, VT_I4)
	DISP_PROPERTY_EX(CBHMDBDropDownCtrl, "DroppedDown", GetDroppedDown, SetNotSupported, VT_BOOL)
	DISP_PROPERTY_EX(CBHMDBDropDownCtrl, "DataField", GetDataField, SetDataField, VT_BSTR)
	DISP_PROPERTY_EX(CBHMDBDropDownCtrl, "Columns", GetColumns, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX(CBHMDBDropDownCtrl, "VisibleRows", GetVisibleRows, SetNotSupported, VT_I4)
	DISP_PROPERTY_EX(CBHMDBDropDownCtrl, "VisibleCols", GetVisibleCols, SetNotSupported, VT_I4)
	DISP_PROPERTY_EX(CBHMDBDropDownCtrl, "ColumnHeader", GetColumnHeader, SetColumnHeader, VT_BOOL)
	DISP_PROPERTY_EX(CBHMDBDropDownCtrl, "HeadFont", GetHeadFont, SetHeadFont, VT_FONT)
	DISP_PROPERTY_EX(CBHMDBDropDownCtrl, "HeaderHeight", GetHeaderHeight, SetHeaderHeight, VT_I4)
	DISP_PROPERTY_EX(CBHMDBDropDownCtrl, "DataSource", GetDataSource, SetDataSource, VT_UNKNOWN)
	DISP_PROPERTY_EX(CBHMDBDropDownCtrl, "DataMember", GetDataMember, SetDataMember, VT_BSTR)
	DISP_PROPERTY_EX(CBHMDBDropDownCtrl, "FieldSeparator", GetFieldSeparator, SetFieldSeparator, VT_BSTR)
	DISP_PROPERTY_EX(CBHMDBDropDownCtrl, "MaxDropDownItems", GetMaxDropDownItems, SetMaxDropDownItems, VT_I2)
	DISP_PROPERTY_EX(CBHMDBDropDownCtrl, "MinDropDownItems", GetMinDropDownItems, SetMinDropDownItems, VT_I2)
	DISP_PROPERTY_EX(CBHMDBDropDownCtrl, "ListWidth", GetListWidth, SetListWidth, VT_I2)
	DISP_PROPERTY_EX(CBHMDBDropDownCtrl, "BackColor", GetBackColor, SetBackColor, VT_COLOR)
	DISP_PROPERTY_EX(CBHMDBDropDownCtrl, "Bookmark", GetBookmark, SetBookmark, VT_VARIANT)
	DISP_PROPERTY_EX(CBHMDBDropDownCtrl, "Groups", GetGroups, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX(CBHMDBDropDownCtrl, "GroupHeader", GetGroupHeader, SetGroupHeader, VT_BOOL)
	DISP_FUNCTION(CBHMDBDropDownCtrl, "AddItem", AddItem, VT_EMPTY, VTS_BSTR VTS_VARIANT)
	DISP_FUNCTION(CBHMDBDropDownCtrl, "RemoveItem", RemoveItem, VT_EMPTY, VTS_I4)
	DISP_FUNCTION(CBHMDBDropDownCtrl, "Scroll", Scroll, VT_EMPTY, VTS_I4 VTS_I4)
	DISP_FUNCTION(CBHMDBDropDownCtrl, "RemoveAll", RemoveAll, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CBHMDBDropDownCtrl, "MoveFirst", MoveFirst, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CBHMDBDropDownCtrl, "MovePrevious", MovePrevious, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CBHMDBDropDownCtrl, "MoveNext", MoveNext, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CBHMDBDropDownCtrl, "MoveLast", MoveLast, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CBHMDBDropDownCtrl, "MoveRecords", MoveRecords, VT_EMPTY, VTS_I4)
	DISP_FUNCTION(CBHMDBDropDownCtrl, "ReBind", ReBind, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CBHMDBDropDownCtrl, "RowBookmark", RowBookmark, VT_VARIANT, VTS_I4)
	DISP_STOCKPROP_FORECOLOR()
	DISP_STOCKPROP_BORDERSTYLE()
	DISP_STOCKPROP_HWND()
	//}}AFX_DISPATCH_MAP
	DISP_FUNCTION_ID(CBHMDBDropDownCtrl, "AboutBox", DISPID_ABOUTBOX, AboutBox, VT_EMPTY, VTS_NONE)
	DISP_PROPERTY_STOCK(CBHMDBDropDownCtrl, "CtrlType", 255, CBHMDBDropDownCtrl::GetCtrlType, COleControl::SetNotSupported, VT_I2)
END_DISPATCH_MAP()


/////////////////////////////////////////////////////////////////////////////
// Event map

BEGIN_EVENT_MAP(CBHMDBDropDownCtrl, COleControl)
	//{{AFX_EVENT_MAP(CBHMDBDropDownCtrl)
	EVENT_CUSTOM("Scroll", FireScroll, VTS_PI2)
	EVENT_CUSTOM("CloseUp", FireCloseUp, VTS_NONE)
	EVENT_CUSTOM("DropDown", FireDropDown, VTS_NONE)
	EVENT_CUSTOM("PositionList", FirePositionList, VTS_BSTR)
	EVENT_CUSTOM("ScrollAfter", FireScrollAfter, VTS_NONE)
	EVENT_STOCK_CLICK()
	//}}AFX_EVENT_MAP
END_EVENT_MAP()


/////////////////////////////////////////////////////////////////////////////
// Property pages

// TODO: Add more property pages as needed.  Remember to increase the count!
BEGIN_PROPPAGEIDS(CBHMDBDropDownCtrl, 5)
	PROPPAGEID(CBHMDBDropDownPropPageGeneral::guid)
	PROPPAGEID(CBHMDBGroupsPropPage::guid)
	PROPPAGEID(CBHMDBColumnPropPage::guid)
	PROPPAGEID(CLSID_StockColorPage)
	PROPPAGEID(CLSID_StockFontPage)
END_PROPPAGEIDS(CBHMDBDropDownCtrl)


/////////////////////////////////////////////////////////////////////////////
// Initialize class factory and guid

IMPLEMENT_OLECREATE_EX(CBHMDBDropDownCtrl, "BHMDATA.BHMDBDropDownCtrl.1",
	0xbf91b129, 0x6a84, 0x11d3, 0xa7, 0xfe, 0, 0x80, 0xc8, 0x76, 0x3f, 0xa4)


/////////////////////////////////////////////////////////////////////////////
// Type library ID and version

IMPLEMENT_OLETYPELIB(CBHMDBDropDownCtrl, _tlid, _wVerMajor, _wVerMinor)


/////////////////////////////////////////////////////////////////////////////
// Interface IDs

const IID BASED_CODE IID_DBHMDBDropDown =
		{ 0xbf91b127, 0x6a84, 0x11d3, { 0xa7, 0xfe, 0, 0x80, 0xc8, 0x76, 0x3f, 0xa4 } };
const IID BASED_CODE IID_DBHMDBDropDownEvents =
		{ 0xbf91b128, 0x6a84, 0x11d3, { 0xa7, 0xfe, 0, 0x80, 0xc8, 0x76, 0x3f, 0xa4 } };


/////////////////////////////////////////////////////////////////////////////
// Control type information

static const DWORD BASED_CODE _dwBHMDBDropDownOleMisc =
	OLEMISC_ACTIVATEWHENVISIBLE |
	OLEMISC_SETCLIENTSITEFIRST |
	OLEMISC_INSIDEOUT |
	OLEMISC_CANTLINKINSIDE |
	OLEMISC_RECOMPOSEONRESIZE;

IMPLEMENT_OLECTLTYPE(CBHMDBDropDownCtrl, IDS_BHMDBDROPDOWN, _dwBHMDBDropDownOleMisc)


/////////////////////////////////////////////////////////////////////////////
// CBHMDBDropDownCtrl::CBHMDBDropDownCtrlFactory::UpdateRegistry -
// Adds or removes system registry entries for CBHMDBDropDownCtrl

BOOL CBHMDBDropDownCtrl::CBHMDBDropDownCtrlFactory::UpdateRegistry(BOOL bRegister)
{
	// TODO: Verify that your control follows apartment-model threading rules.
	// Refer to MFC TechNote 64 for more information.
	// If your control does not conform to the apartment-model rules, then
	// you must modify the code below, changing the 6th parameter from
	// afxRegApartmentThreading to 0.

	if (bRegister)
		return AfxOleRegisterControlClass(
			AfxGetInstanceHandle(),
			m_clsid,
			m_lpszProgID,
			IDS_BHMDBDROPDOWN,
			IDB_BHMDBDROPDOWN,
			afxRegApartmentThreading,
			_dwBHMDBDropDownOleMisc,
			_tlid,
			_wVerMajor,
			_wVerMinor);
	else
		return AfxOleUnregisterClass(m_clsid, m_lpszProgID);
}


/////////////////////////////////////////////////////////////////////////////
// Licensing strings

static const TCHAR BASED_CODE _szLicFileName[] = _T("BHMData.lic");

static const WCHAR BASED_CODE _szLicString[] =
	L"Copyright (c) 1999 BHM";


/////////////////////////////////////////////////////////////////////////////
// CBHMDBDropDownCtrl::CBHMDBDropDownCtrlFactory::VerifyUserLicense -
// Checks for existence of a user license

BOOL CBHMDBDropDownCtrl::CBHMDBDropDownCtrlFactory::VerifyUserLicense()
{
	return AfxVerifyLicFile(AfxGetInstanceHandle(), _szLicFileName,
		_szLicString);
}


/////////////////////////////////////////////////////////////////////////////
// CBHMDBDropDownCtrl::CBHMDBDropDownCtrlFactory::GetLicenseKey -
// Returns a runtime licensing key

BOOL CBHMDBDropDownCtrl::CBHMDBDropDownCtrlFactory::GetLicenseKey(DWORD dwReserved,
	BSTR FAR* pbstrKey)
{
	if (pbstrKey == NULL)
		return FALSE;

	*pbstrKey = SysAllocString(_szLicString);
	return (*pbstrKey != NULL);
}


/////////////////////////////////////////////////////////////////////////////
// CBHMDBDropDownCtrl::CBHMDBDropDownCtrl - Constructor

CBHMDBDropDownCtrl::CBHMDBDropDownCtrl() : m_fntCell(NULL), m_fntHeader(NULL)
{
	InitializeIIDs(&IID_DBHMDBDropDown, &IID_DBHMDBDropDownEvents);

	m_bEndReached = TRUE;
	m_bDataSrcChanged = FALSE;
	m_dwCookieRN = 0;
	m_bHaveBookmarks = FALSE;
	m_pColumnInfo = NULL;
	m_pStrings = NULL;
	m_nColumns = m_nColumnsBind = 0;
	m_nBookmarkSize = 0;

	m_spDataSource = NULL;
	m_spRowPosition = NULL;

	m_pGridDraw = NULL;
	m_pDropDownRealWnd = NULL;
	m_bDataSrcChanged = FALSE;
	m_spDataSource = NULL;
	m_nCol = m_nRow = 0;
	m_nColsUsed = 0;
	m_bDroppedDown = FALSE;
}


/////////////////////////////////////////////////////////////////////////////
// CBHMDBDropDownCtrl::~CBHMDBDropDownCtrl - Destructor

CBHMDBDropDownCtrl::~CBHMDBDropDownCtrl()
{
	// Delete the internal grid.
	if (m_pGridDraw)
	{
		m_pGridDraw->DestroyWindow();
		delete m_pGridDraw;
		m_pGridDraw = NULL;
	}

	// Delete the dropdown window.
	if (m_pDropDownRealWnd)
	{
		m_pDropDownRealWnd->DestroyWindow();
		delete m_pDropDownRealWnd;
		m_pDropDownRealWnd = NULL;
	}

	// Clear the groups layout array.
	for (int i = 0; i < m_arGroups.GetSize(); i ++)
		delete m_arGroups[i];

	m_arGroups.RemoveAll();

	// Clear the cells layout array.
	for (i = 0; i < m_arCells.GetSize(); i ++)
		delete m_arCells[i];

	m_arCells.RemoveAll();


	// Clear the cols array.
	i = m_arCols.GetSize();

	for (; i > 0; i --)
		delete m_arCols[i - 1];

	m_arCols.RemoveAll();

	// Clear the cached bind information array.
	for (i = 0; i < m_arBindInfo.GetSize(); i ++)
		delete m_arBindInfo[i];

	m_arBindInfo.RemoveAll();

	// Clear the cached interface pointers.
	ClearInterfaces();
}


/////////////////////////////////////////////////////////////////////////////
// CBHMDBDropDownCtrl::OnDraw - Drawing function

void CBHMDBDropDownCtrl::OnDraw(
			CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid)
{
	// If the internal grid is not created, create it here.
	if (!m_pGridDraw && !AmbientUserMode())
		CreateDrawGrid();

	// At the run time, the end user can not see the main windows.
	if (AmbientUserMode())
	{
		// Shrink the main window to 1 point size.
		SetControlSize(1, 1);

		// Hide the main window.
		if (IsWindowVisible())
			ShowWindow(SW_HIDE);

		return;
	}
	else if (m_pGridDraw)
	{
		// It's at design time.

		CRect rt = rcBounds;

		// Let the grid draw in the given rect.
		m_pGridDraw->SetGridRect(&rt);

		// Let the grid draw itself.
		m_pGridDraw->LockUpdate(FALSE);
		m_pGridDraw->OnGridDraw(pdc);
		m_pGridDraw->LockUpdate();
	}
}


/////////////////////////////////////////////////////////////////////////////
// CBHMDBDropDownCtrl::DoPropExchange - Persistence support

void CBHMDBDropDownCtrl::DoPropExchange(CPropExchange* pPX)
{
	ExchangeVersion(pPX, MAKELONG(_wVerMinor, _wVerMajor));
	COleControl::DoPropExchange(pPX);

	PX_Bool(pPX, "ColumnHeader", m_bColumnHeader, TRUE);
	PX_Font(pPX, "Font", m_fntCell, &_BHMFontDescDefault);
	PX_Font(pPX, "HeaderFont", m_fntHeader, &_BHMFontDescDefault);
	PX_Long(pPX, "HeaderHeight", m_nHeaderHeight, 25);
	PX_Long(pPX, "RowHeight", m_nRowHeight, 25);
	PX_Long(pPX, "DataMode", m_nDataMode, 0);
	PX_Long(pPX, "DefColWidth", m_nDefColWidth, 100);
	PX_Long(pPX, "Rows", m_nRows, 2);
	PX_Long(pPX, "Cols", m_nCols, 2);
	PX_Long(pPX, "DividerType", m_nDividerType, 3);
	PX_Long(pPX, "DividerStyle", m_nDividerStyle, 0);
	PX_String(pPX, "FieldSeparator", m_strFieldSeparator, _T("(tab)"));
	PX_Color(pPX, "HeadForeColor", m_clrHeadFore, RGB(0, 0, 0));
	PX_Color(pPX, "HeadBackColor", m_clrHeadBack, GetSysColor(COLOR_3DFACE));
	PX_String(pPX, "DataField", m_strDataField);
	PX_String(pPX, "FieldSeparator", m_strFieldSeparator, _T("(tab)"));
	PX_Bool(pPX, "ListAutoPosition", m_bListAutoPosition, TRUE);
	PX_Bool(pPX, "ListWidthAutoSize", m_bListWidthAutoSize, TRUE);
	PX_Short(pPX, "MaxDropDownItems", m_nMaxDropDownItems, 8);
	PX_Short(pPX, "MinDropDownItems", m_nMinDropDownItems, 1);
	PX_Short(pPX, "ListWidth", m_nListWidth, 0);
	PX_Color(pPX, "BackColor", m_clrBack, GetSysColor(COLOR_WINDOW));
	PX_Color(pPX, "DividerColor", m_clrDivider, RGB(0, 0, 0));
	PX_Bool(pPX, "GroupHeader", m_bGroupHeader, TRUE);

    if(!pPX->IsLoading())
    {
		//  The control's properties are being saved.
		// Create a CMemFile
		CMemFile memFile;

		// construct an archive for saving
		CArchive ar( &memFile, CArchive::store, 512);

		SerializeColumnProp(ar, FALSE);
		// iterate through the tabs and serialize
		try
		{
			ar.Close();
		}
		catch(...)
		{
			TRACE0("Exception while writing data to file in DoPropExchange\n");
		}
		
		//BLOBDATA blobData;
		
		DWORD dwLength = memFile.GetLength();
		m_hBlob = GlobalAlloc(GMEM_FIXED, sizeof(long)+dwLength);
 
		if(m_hBlob != NULL)
		{
        
			BYTE * pMem = (BYTE*)GlobalLock(m_hBlob);
		
			ASSERT(pMem);

			if(pMem != NULL)
			{
          
				  //  Fill in the first part of your structure with the
				  //  sizeof(your property data).
				  BYTE* p = memFile.Detach();  
				  
				  memcpy(pMem, (DWORD*)&dwLength, sizeof(DWORD));
				  
				  
				  DWORD cb = *(DWORD*) pMem;

				  // If you are looking at this ASSERT then an incorrect number of
				  // bytes is being indicated. PX_Blob will ASSERT.
				  ASSERT(cb == dwLength);

				  pMem+=sizeof(DWORD);


				  //  Fill in the second part of your structure with the actual
				  //  data that represents the control's property value.
          
				  memcpy(pMem, p, dwLength);
				  
				  memFile.Attach(p, dwLength);
				  ((__CMemFile*)&memFile)->m_bAutoDelete = TRUE;


				  pMem-=sizeof(DWORD);
				  cb = *(DWORD*) pMem;

				  // If you are looking at this ASSERT then an incorrect number of
				  // bytes is being indicated. PX_Blob will ASSERT.
				  
				  // do a sanity check!
				  ASSERT(cb == dwLength);

          
				  //  Call PX_Blob, unlock and free your global memory.
				  PX_Blob(pPX, _T("BlobProp"), m_hBlob);
 
				  GlobalUnlock(m_hBlob);
				  
				  //m_hBlob = NULL;
			}
			else
			{
				  // The GlobalLock call failed. Pass in a NULL HGLOBAL for the
				  // third parameter to PX_Blob. This will cause it to write a
				  // value of zero for the BLOB data.
				  HGLOBAL hTmp = NULL;
 
				  PX_Blob(pPX, _T("BlobProp"), hTmp);
			}
 
			if(m_hBlob !=NULL)
				::GlobalFree(m_hBlob);
			
			m_hBlob = NULL;
		}
		else
			// The GlobalAlloc call failed. Pass in a NULL HGLOBAL for the
			// third parameter to PX_Blob. This will cause it to write a
			// value of zero for the BLOB data.
			PX_Blob(pPX, _T("BlobProp"), m_hBlob);
    }
    else
    {
      
		//  Properties are being loaded into the control.
		//  Call PX_Blob to get the handle to the memory containing
		//  the property data.
		//  We just use an arbitrary property name. Else it will fail 
		//  with prop bags
		PX_Blob(pPX, _T("BlobProp"), m_hBlob);
 
		if(m_hBlob != NULL)
		{
			BYTE* pMem = (BYTE*)GlobalLock(m_hBlob);
			  
			  // Create a CMemFile
			CMemFile memFile;

			DWORD cb = *(DWORD*) pMem;

			pMem+=sizeof(cb);

			memFile.Attach(pMem, cb);

			// construct an archive for loading data
			CArchive ar( &memFile, CArchive::load, 512);

			// Lock the memory to get a pointer to the actual property
			// data.
			if(pMem != NULL)
			{
				//  Set the value of the property, unlock and free the global
				//  memory.
         
				SerializeColumnProp(ar, TRUE);
				try
				{
					ar.Close();
				}
				catch(...)
				{
					TRACE0("Exception while reading data from file in DoPropExchange\n");
				}
			
				// unlock				
				GlobalUnlock(m_hBlob);
			}
			else
				TRACE0("there is no columns in persistent data.\n");
		}
	}		
 
	GlobalFree(m_hBlob);
	m_hBlob = NULL;
}


/////////////////////////////////////////////////////////////////////////////
// CBHMDBDropDownCtrl::OnResetState - Reset control to default state

void CBHMDBDropDownCtrl::OnResetState()
{
	COleControl::OnResetState();  // Resets defaults found in DoPropExchange

	// TODO: Reset any other control state here.
}


/////////////////////////////////////////////////////////////////////////////
// CBHMDBDropDownCtrl::AboutBox - Display an "About" box to the user

void CBHMDBDropDownCtrl::AboutBox()
{
	CDialog dlgAbout(IDD_ABOUTBOX_BHMDBDROPDOWN);
	dlgAbout.DoModal();
}


/////////////////////////////////////////////////////////////////////////////
// CBHMDBDropDownCtrl message handlers

// The backdoor for property page :)
short CBHMDBDropDownCtrl::GetCtrlType() 
{
	return 1;
}

int CBHMDBDropDownCtrl::GetColOrdinalFromIndex(int nColIndex)
/*
Routine Description:
	Get the col ordinal in bind info array from col index.

Parameters:
	nColIndex		The col index.

Return Value:
	If the col index is valid, return the col ordinal; Otherwise, return INVALID.
*/
{
	// Compare this index with the index of each col in array.

	for (int i = 0; i < m_arBindInfo.GetSize(); i ++)
		if (m_arBindInfo[i]->nColIndex == nColIndex)
			return i;

	return INVALID;
}

int CBHMDBDropDownCtrl::GetColOrdinalFromCol(int nCol)
/*
Routine Description:
	Get the col ordinal in bind info array from col number.

Parameters:
	nCol 		The col number.

Return Value:
	If the col number is valid, return the ordinal; Otherwise, return INVALID.
*/
{
	if (m_pGridDraw)
		return GetColOrdinalFromIndex(m_pGridDraw->GetColIndex(nCol));
	else if (m_pDropDownRealWnd)
		return GetColOrdinalFromIndex(m_pDropDownRealWnd->GetColIndex(nCol));

	return INVALID;
}

int CBHMDBDropDownCtrl::GetColOrdinalFromField(int nField)
/*
Routine Description:
	Get the col ordinal in bind info array from field number.

Parameters:
	nField		The field number.

Return Value:
	If the field number is valid, return the col ordinal; Otherwise, return INVALID.
*/
{
	// Compare the field with each field in array.


	for (int i = 0; i < m_arBindInfo.GetSize(); i ++)
		if (m_arBindInfo[i]->nDataField == nField)
			return i;

	return INVALID;
}

int CBHMDBDropDownCtrl::GetFieldFromStr(CString str)
/*
Routine Description:
	Calc the bounded field number from the column name.

Parameters:
	str		The field name.

Return Value:
	If the field name exists, return its field number; Otherwise, return -1.
*/
{
	// Field name can not be empty.

	if (str.IsEmpty())
		return -1;

	// Compare the given name with each col in grid.
	for (int i = 0; i < m_apColumnData.GetSize(); i ++)
		if (!str.CompareNoCase(m_apColumnData[i]->strName))
		// Found it. Return the bounded field number of this col.
			return m_apColumnData[i]->nColumn;

	// The column name is invalid.
	return -1;
}

void CBHMDBDropDownCtrl::CreateDrawGrid()
/*
Routine Description:
	Create the internal grid.

Parameters:
	None.
	
Return Value:
	None.
*/
{
	// If already created, just return.
	if (m_pGridDraw)
		return;

	// Get the parent window for internal grid.
	HWND hWnd = GetApplicationWindow(); 
	
	// Create the internal grid.
	m_pGridDraw = new CGridInner(NULL);
	m_pGridDraw->Create(WS_VISIBLE | WS_CHILD, CRect(100,100, 110, 110), CWnd::FromHandle(hWnd), 101);
	InitDrawGrid();
}

HWND CBHMDBDropDownCtrl::GetApplicationWindow()
/*
Routine Description:
	Get the parent window which will holds the created internal grid.

Parameters:
	None.
	
Return Value:
	The window handle found.
*/
{
	HWND  hwnd = NULL;
	HRESULT hr;
	
	
	if (m_pInPlaceSite != NULL)
	{
		// Get parent window from InPlaceSite.
		m_pInPlaceSite->GetWindow(&hwnd);
		if (hwnd!=NULL)
			return hwnd;
		else
			return AfxGetMainWnd( )->GetSafeHwnd();
	}
	
	
	LPOLECLIENTSITE pOleClientSite = GetClientSite();
	
	if ( pOleClientSite )
	{
		// Get parent window from OleClientSite.
		IOleWindow* pOleWindow;
		hr = pOleClientSite->QueryInterface( IID_IOleWindow, (LPVOID*) 
			&pOleWindow );
		
		if ( pOleWindow )
		{
			pOleWindow->GetWindow( &hwnd );
			pOleWindow->Release();
			if (hwnd!=NULL)
				return hwnd;
			else
				
				return AfxGetMainWnd( )->GetSafeHwnd();
		}
	}
	
	// I have no other choice except returning the main application window.
	return AfxGetMainWnd()->GetSafeHwnd();
}


int CBHMDBDropDownCtrl::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (COleControl::OnCreate(lpCreateStruct) == -1)
		return -1;
	
	CRect rt(0, 0, lpCreateStruct->cx, lpCreateStruct->cy);
	if (!AmbientUserMode())
	{
		// Create internal grid at design time.
		m_pGridDraw = new CGridInner(NULL);
		m_pGridDraw->Create(WS_VISIBLE | WS_CHILD, rt, this, 101);

		m_pGridDraw->LockUpdate();
		m_pGridDraw->ShowWindow(SW_HIDE);
		InitDrawGrid();
	}
	else
	{
		// Create dropdown window here at run time.
		m_pDropDownRealWnd = new CDropDownRealWnd(this);
		m_pDropDownRealWnd->CreateEx(WS_EX_TOOLWINDOW,
			BHMGRID_CLASSNAME, NULL, WS_CHILD | WS_BORDER, rt, GetDesktopWindow(), NULL, NULL);

		InitDropDownWnd();
	}
	
	return 0;
}

void CBHMDBDropDownCtrl::InitDrawGrid()
/*
Routine Description:
	Init the grid from cached prop data.

Parameters:
	None.

Return Value:
	None.
*/
{
	UpdateDivider();

	m_pGridDraw->SetForeColor(TranslateColor(GetForeColor()));
	m_pGridDraw->SetBackColor(TranslateColor(GetBackColor()));
	m_pGridDraw->SetHeaderForeColor(TranslateColor(m_clrHeadFore));
	m_pGridDraw->SetHeaderBackColor(TranslateColor(m_clrHeadBack));

	InitHeaderFont();
	InitCellFont();

	m_pGridDraw->SetReadOnly(FALSE);
	
	m_pGridDraw->SetHeaderHeight(m_nHeaderHeight);
	m_pGridDraw->SetRowHeight(m_nRowHeight);
	m_pGridDraw->SetDefColWidth(m_nDefColWidth);
	m_pGridDraw->SetShowGroupHeader(m_bGroupHeader);

	InitGridFromProp();
	
	m_pGridDraw->SetReadOnly(TRUE);
	m_pGridDraw->SetShowColHeader(m_bColumnHeader);

	InvalidateControl();
}

void CBHMDBDropDownCtrl::UpdateDivider()
/*
Routine Description:
	Update the divider lines.

Parameters:
	None.

Return Value:
	None.
*/
{
	if (m_pGridDraw)
	{
		m_pGridDraw->SetDividerType(m_nDividerType);
		m_pGridDraw->SetDividerStyle(m_nDividerStyle);
		m_pGridDraw->SetDividerColor(m_clrDivider);
	}
	else if (m_pDropDownRealWnd)
	{
		m_pDropDownRealWnd->SetDividerType(m_nDividerType);
		m_pDropDownRealWnd->SetDividerStyle(m_nDividerStyle);
		m_pDropDownRealWnd->SetDividerColor(m_clrDivider);
	}
}

void CBHMDBDropDownCtrl::InitCellFont()
/*
Routine Description:
	Init the cell font from font object.

Parameters:
	None.

Return Value:
	None.
*/
{
	// The font handle.

	HFONT hFont;
	
	if (m_fntCell.m_pFont->get_hFont(&hFont) == S_OK)
	{
		// Get the LOGFONT information.
		LOGFONT logCell;
		if (GetObject(hFont, sizeof(logCell), &logCell))
		{
			// Update the grid.
			if (m_pGridDraw)
				m_pGridDraw->SetCellFont(&logCell);
			else if (m_pDropDownRealWnd)
				m_pDropDownRealWnd->SetCellFont(&logCell);
		}
	}
}

LPFONTDISP CBHMDBDropDownCtrl::GetFont() 
/*
Routine Description:
	Get the font object.

Parameters:
	None.

Return Value:
	The pointer to the font object.
*/
{
	return m_fntCell.GetFontDispatch();
}

void CBHMDBDropDownCtrl::SetFont(LPFONTDISP newValue) 
/*
Routine Description:
	Set the font object.

Parameters:
	newValue		The new font object.

Return Value:
	None.
*/
{
	// Be careful in null pointer.

	ASSERT((newValue == NULL) ||
		   AfxIsValidAddress(newValue, sizeof(newValue), FALSE));

	// Init the font object.
	m_fntCell.InitializeFont(NULL, newValue);

	// Update the grid.
	InitCellFont();
	
	// Update the row height.
	if (m_pGridDraw)
		SetRowHeight(m_pGridDraw->GetRowHeight(0));
	else if (m_pDropDownRealWnd)
		SetRowHeight(m_pDropDownRealWnd->GetRowHeight(0));

	SetModifiedFlag();
	BoundPropertyChanged(dispidFont);
	InvalidateControl();
}

long CBHMDBDropDownCtrl::GetDefColWidth() 
/*
Routine Description:
	Get the default column width.

Parameters:
	None.

Return Value:
	The default column width.
*/
{
	return m_nDefColWidth;
}

void CBHMDBDropDownCtrl::SetDefColWidth(long nNewValue) 
/*
Routine Description:
	Set the default column width.

Parameters:
	nNewValue		The new value.

Return Value:
	None.
*/
{
	if (nNewValue < 0)
		ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDWIDTHORHEIGHT);

	// Save the new value.
	m_nDefColWidth = nNewValue;

	// Update the grid.
	if (m_pGridDraw)
		m_pGridDraw->SetDefColWidth(m_nDefColWidth);
	else if (m_pDropDownRealWnd)
		m_pDropDownRealWnd->SetDefColWidth(m_nDefColWidth);

	SetModifiedFlag();
	BoundPropertyChanged(dispidDefColWidth);
}

long CBHMDBDropDownCtrl::GetRowHeight() 
/*
Routine Description:
	Get the height of each content row.

Parameters:
	None.

Return Value:
	The height of each content row.
*/
{
	// If exists dropdown window, update the cached value from it.

	if (m_pDropDownRealWnd)
		m_nRowHeight = m_pDropDownRealWnd->GetRowHeight(1);
	else if (m_pGridDraw)
		m_nRowHeight = m_pGridDraw->GetRowHeight(1);

	return m_nRowHeight;
}

void CBHMDBDropDownCtrl::SetRowHeight(long nNewValue) 
/*
Routine Description:
	Set the height of each content row.

Parameters:
	nNewValue		The new value.

Return Value:
	None.
*/
{
	// Rowheight can not less than 0.

	if (nNewValue < 0)
		ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDWIDTHORHEIGHT);

	// Save the new value.
	m_nRowHeight = nNewValue;

	// Update the grid.
	if (m_pGridDraw)
		m_pGridDraw->SetRowHeight(m_nRowHeight);
	else if (m_pDropDownRealWnd)
		m_pDropDownRealWnd->SetRowHeight(m_nRowHeight);

	SetModifiedFlag();
	BoundPropertyChanged(dispidRowHeight);
	InvalidateControl();
}

long CBHMDBDropDownCtrl::GetDividerType() 
/*
Routine Description:
	Get the type of divider lines.

Parameters:
	None.

Return Value:
	The type of divider lines.
*/
{
	return m_nDividerType;
}

void CBHMDBDropDownCtrl::SetDividerType(long nNewValue) 
/*
Routine Description:
	Set the type of divider lines.

Parameters:
	nNewValue		The new value.

Return Value:
	None.
*/
{
	m_nDividerType = nNewValue;

	UpdateDivider();

	SetModifiedFlag();
	BoundPropertyChanged(dispidDividerType);
	InvalidateControl();
}

long CBHMDBDropDownCtrl::GetDividerStyle() 
/*
Routine Description:
	Get the style of divider lines.

Parameters:
	None.

Return Value:
	The style of divider lines.
*/
{
	return m_nDividerStyle;
}

void CBHMDBDropDownCtrl::SetDividerStyle(long nNewValue) 
/*
Routine Description:
	Set the style of divider lines.

Parameters:
	nNewValue		The new value.

Return Value:
	None.
*/
{
	// Save the new value.
	m_nDividerStyle = nNewValue;

	// Update the grid.
	UpdateDivider();

	SetModifiedFlag();
	BoundPropertyChanged(dispidDividerStyle);
	InvalidateControl();
}

long CBHMDBDropDownCtrl::GetDataMode() 
/*
Routine Description:
	Get the data mode.

Parameters:
	None.

Return Value:
	The data mode.
*/
{
	return m_nDataMode;
}

void CBHMDBDropDownCtrl::SetDataMode(long nNewValue) 
/*
Routine Description:
	Set the data mode.

Parameters:
	nNewValue		The new value.

Return Value:
	None.
*/
{
	// Can not change datamode in run mode.
	if (AmbientUserMode())
		ThrowError(CTL_E_SETNOTSUPPORTEDATRUNTIME);

	if (nNewValue == 0)
	// The new mode is bound mode.
		m_nDataMode = 0;
	else
	{
		// The new mode is manual mode.
		if (m_spDataSource)
		// Can not change to manual mode while having datasource.
			ThrowError(CTL_E_SETNOTSUPPORTED, IDS_ERROR_NOCHANGETOMANUALMODE);
		else
			m_nDataMode = 1;
	}

	SetModifiedFlag();
	BoundPropertyChanged(dispidDataMode);
}

long CBHMDBDropDownCtrl::GetLeftCol() 
/*
Routine Description:
	Get the first left column visible.

Parameters:
	None.

Return Value:
	The ordinal of that column.
*/
{
	// This property is unavailable at design time.
	if (!AmbientUserMode())
		GetNotSupported();

	if (m_pDropDownRealWnd)
		return m_pDropDownRealWnd->GetLeftCol();
	else
		return 0;
}

void CBHMDBDropDownCtrl::SetLeftCol(long nNewValue) 
/*
Routine Description:
	Set the first left column visible.

Parameters:
	nNewValue		The new value.

Return Value:
	None.
*/
{
	// This property is unavailable at design time.

	if (!AmbientUserMode())
		SetNotSupported();

	// Update the grid.
	if (m_pDropDownRealWnd)
		m_pDropDownRealWnd->SetLeftCol(nNewValue);
}

long CBHMDBDropDownCtrl::GetFirstRow() 
/*
Routine Description:
	Get the first row visible.

Parameters:
	None.

Return Value:
	The ordinal of that row.
*/
{
	// This property is unavailable at design time.

	if (!AmbientUserMode())
		GetNotSupported();

	return m_pDropDownRealWnd->GetTopRow();
}

void CBHMDBDropDownCtrl::SetFirstRow(long nNewValue) 
/*
Routine Description:
	Set the first row visible.

Parameters:
	nNewValue		The new value.

Return Value:
	None.
*/
{
	// This property is unavailable at design time.

	if (!AmbientUserMode())
		SetNotSupported();

	m_pDropDownRealWnd->SetTopRow(nNewValue);
}

long CBHMDBDropDownCtrl::GetRow() 
/*
Routine Description:
	Get the ordinal of current row.

Parameters:
	None.

Return Value:
	The the ordinal of current row.
*/
{
	// This property is unavailable at design time.

	if (!AmbientUserMode())
		GetNotSupported();

	int nCol;

	m_pDropDownRealWnd->GetCurrentCell(m_nRow, nCol);

	return m_nRow;
}

void CBHMDBDropDownCtrl::SetRow(long nNewValue) 
/*
Routine Description:
	Set the given row as current row.

Parameters:
	nNewValue		The new value.

Return Value:
	None.
*/
{
	// This property is unavailable at design time.
	if (!AmbientUserMode())
		SetNotSupported();

	// Save the new value.
	m_nRow = nNewValue;

	// Current row can not exceed the row count of the grid.
	if (m_pDropDownRealWnd->GetRowCount() >= (int)m_nRow)
	// Update the grid.
		m_pDropDownRealWnd->MoveTo(m_nRow);
	else
		ThrowError(CTL_E_INVALIDPROPERTYARRAYINDEX, AFX_IDP_E_INVALIDPROPERTYARRAYINDEX);

	SetModifiedFlag();
	BoundPropertyChanged(dispidRow);
}

long CBHMDBDropDownCtrl::GetCol() 
/*
Routine Description:
	Get the ordinal of current col.

Parameters:
	None.

Return Value:
	The ordinal of current col.
*/
{
	// This property is unavailable at design time.

	if (!AmbientUserMode())
		GetNotSupported();

	int nRow;
	
	m_pDropDownRealWnd->GetCurrentCell(nRow, m_nCol);
	return m_nCol;
}

void CBHMDBDropDownCtrl::SetCol(long nNewValue) 
/*
Routine Description:
	Set the given col as current col.

Parameters:
	nNewValue		The new value.

Return Value:
	None.
*/
{
	// This property is unavailable at design time.

	if (!AmbientUserMode())
		SetNotSupported();

	// Save the new value.
	m_nCol = nNewValue;

	// Current col can not exceed the col count of the grid.
	if (m_pDropDownRealWnd->GetColCount() >= (int)m_nCol)
	{
		int nRow, nCol;
		m_pDropDownRealWnd->GetCurrentCell(nRow, nCol);
		m_pDropDownRealWnd->SetCurrentCell(nRow, m_nCol);
	}
	else
		ThrowError(CTL_E_INVALIDPROPERTYARRAYINDEX, AFX_IDP_E_INVALIDPROPERTYARRAYINDEX);

	SetModifiedFlag();
	BoundPropertyChanged(dispidCol);
}

long CBHMDBDropDownCtrl::GetRows() 
/*
Routine Description:
	Get the row count.

Parameters:
	None.

Return Value:
	The row count.
*/
{
	// If the grid exists, return the value of the it; Otherwise, return the cached value.

	if (!m_pDropDownRealWnd && !m_pGridDraw)
		return m_nRows;

	if (m_pGridDraw)
		m_nRows = m_pGridDraw->GetRowCount();
	else
		m_nRows = m_pDropDownRealWnd->GetRowCount();

	return m_nRows;
}

void CBHMDBDropDownCtrl::SetRows(long nNewValue) 
/*
Routine Description:
	Set the row count.

Parameters:
	nNewValue		The new value.

Return Value:
	None.
*/
{
	// Can not set row count having layout.

	if (m_bHasLayout)
		ThrowError(CTL_E_SETNOTSUPPORTED, IDS_ERROR_NOROWSCOLSHAVINGCONTENT);

	// Rows can not less than 0.
	if (nNewValue < 0)
		return;

	if (m_nDataMode == 0)
	// Can not set row count in bound mode.
		ThrowError(CTL_E_SETNOTSUPPORTED, IDS_ERROR_NOSETROWSINBINDMODE);
	if (AmbientUserMode())
	// Can not set row count at run time.
		SetNotSupported();

	// Save the new value.
	m_nRows = nNewValue;

	m_pGridDraw->SetRowCount(m_nRows);

	SetModifiedFlag();
	BoundPropertyChanged(dispidRows);
	InvalidateControl();
}

long CBHMDBDropDownCtrl::GetCols() 
/*
Routine Description:
	Get the column count.

Parameters:
	None.

Return Value:
	The column count.
*/
{
	// Be careful in null pointer.

	if (!m_pDropDownRealWnd && !m_pGridDraw)
		return m_nCols;

	// Update the cached value.
	if (m_pGridDraw)
		m_nCols = m_pGridDraw->GetColCount();
	else
		m_nCols = m_pDropDownRealWnd->GetColCount();

	return m_nCols;
}

void CBHMDBDropDownCtrl::SetCols(long nNewValue) 
/*
Routine Description:
	Set the column count.

Parameters:
	nNewValue		The new value.

Return Value:
	None.
*/
{
	// Can not set col count having layout.

	if (m_bHasLayout)
		ThrowError(CTL_E_SETNOTSUPPORTED, IDS_ERROR_NOROWSCOLSHAVINGCONTENT);

	// Col count should between 0 and 255.
	if (nNewValue < 0)
		return;
	else if (nNewValue > 255)
		ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_MAX255COLS);

	if (m_nDataMode == 0)
	// Can not set col count in bound mode.
		ThrowError(CTL_E_SETNOTSUPPORTED, IDS_ERROR_NOSETCOLSINBINDMODE);
	if (AmbientUserMode())
	// Can not set col count at run time.
		SetNotSupported();

	// Save the new value.
	m_nCols = nNewValue;

	// Update the grid.	
	if (m_pGridDraw)
		m_pGridDraw->SetColCount(m_nCols);

	SetModifiedFlag();
	BoundPropertyChanged(dispidCols);
	InvalidateControl();
}

BOOL CBHMDBDropDownCtrl::GetDroppedDown() 
/*
Routine Description:
	Get the visibility of the dropdown window.

Parameters:
	None.

Return Value:
	If the dropdown window is visible, retutrn TRUE; Otherwise, return FALSE.
*/
{
	if (!AmbientUserMode())
		GetNotSupported();

	return m_bDroppedDown;
}

BSTR CBHMDBDropDownCtrl::GetDataField() 
/*
Routine Description:
	Get the column name from which to retrieve text.

Parameters:
	None.

Return Value:
	The column name.
*/
{
	return m_strDataField.AllocSysString();
}

void CBHMDBDropDownCtrl::SetDataField(LPCTSTR lpszNewValue) 
/*
Routine Description:
	Set the column name from which to retrieve text.

Parameters:
	lpszNewValue		The new value.

Return Value:
	None.
*/
{
	// Save the new value.

	m_strDataField = lpszNewValue;

	SetModifiedFlag();
	BoundPropertyChanged(dispidDataField);
}

LPDISPATCH CBHMDBDropDownCtrl::GetColumns() 
/*
Routine Description:
	Get the "Columns" object.

Parameters:
	None.

Return Value:
	The pointer of the object.
*/
{
	// Create a new object.

	CColumns * pColumns = new CColumns;
	pColumns->SetCtrlPtr(this);
	pColumns->m_pGrid = m_pDropDownRealWnd;

	return pColumns->GetIDispatch(FALSE);
}

long CBHMDBDropDownCtrl::GetVisibleRows() 
/*
Routine Description:
	Get the value of count of visible rows.

Parameters:
	None.

Return Value:
	The visible row count.
*/
{
	// This property is unavailable at design time.

	if (!AmbientUserMode())
		GetNotSupported();

	return m_pDropDownRealWnd->GetVisibleRows();
}

long CBHMDBDropDownCtrl::GetVisibleCols() 
/*
Routine Description:
	Get the count of visible columns.

Parameters:
	None.

Return Value:
	The visible column count.
*/
{
	// This property is unavailable at design time.
	if (!AmbientUserMode())
		GetNotSupported();

	return m_pDropDownRealWnd->GetVisibleCols();
}

void CBHMDBDropDownCtrl::SerializeColumnProp(CArchive &ar, BOOL bLoad)
/*
Routine Description:
	Serialize the props of each column.

Parameters:
	ar		The CArchive object to serialize into or from.

	bLoad		If TRUE, we are loading data; Otherwise, we are saving data.
	
Return Value:
	None.
*/
{
	if (!bLoad)
	{
		// We are saving.

		// Save the groups data.

		// The group count.
		int nGroups = m_arGroups.GetSize();
		Group * pGroup;

		// Save the group count first.
		ar << nGroups;

		// The level count.
		int nLevels = nGroups > 0 ? m_arGroups[0]->arLevels.GetSize() : 1;

		// Save the level count.
		ar << nLevels;

		for (int i = 0; i < nGroups; i ++)
		{
			// Save each group props.

			pGroup = m_arGroups[i];

			ar << pGroup->bAllowSizing << pGroup->bVisible << pGroup->clrBack;
			ar << pGroup->clrFore << pGroup->nCols << pGroup->nGroupIndex;
			ar << pGroup->nWidth << pGroup->strTitle << pGroup->strName;

			for (int j = 0; j < nLevels; j ++)
			{
				// Save each level props.
				Level * pLevel = pGroup->arLevels[j];

				ar << pLevel->nCols << pLevel->nColsVisible;
			}
		}
		
		// Save the column props.

		int nCols = m_arCols.GetSize();
		Col * pCol;

		// Save the column count first.
		ar << nCols;

		for (i = 0; i < nCols; i ++)
		{
			// Save each column props.

			pCol = m_arCols[i];

			ar << pCol->bAllowSizing;
			ar << pCol->bVisible;
			ar << pCol->clrBack << pCol->clrFore << pCol->clrHeaderBack << pCol->clrHeaderFore;
			ar << pCol->nAlignment << pCol->nHeaderAlignment << pCol->nCase;
			ar << pCol->nWidth << pCol->strTitle;
			ar << pCol->arUserAttrib[0];
			ar << pCol->strName;
		}

		// Save cell data.

		int nCells = m_arCells.GetSize();
		
		// Save cell count first.		
		ar << nCells;
		
		// Save each cell.
		for (i = 0; i < nCells; i ++)
			ar << m_arCells[i]->strText;
	}
	else
	{
		// We are loading.

		int nGroups, nLevels;
		Group * pGroup;

		// Retrieve group count and level count first.
		ar >> nGroups >> nLevels;

		// If group count is not zero, we have predefined layout.
		if (nGroups)
			m_bHasLayout = TRUE;

		// Group name.
		CString strGroupName;

		for (int i = 0; i < nGroups; i ++)
		{
			// Load group props.

			// Create a new group object.
			pGroup = new Group;

			ar >> pGroup->bAllowSizing >> pGroup->bVisible >> pGroup->clrBack;
			ar >> pGroup->clrFore >> pGroup->nCols >> pGroup->nGroupIndex;
			ar >> pGroup->nWidth >> pGroup->strTitle >> pGroup->strName;

			for (int j = 0; j < nLevels; j ++)
			{
				// Load level props.

				// Create a new level object.
				Level * pLevel = new Level;

				ar >> pLevel->nCols >> pLevel->nColsVisible;

				// Add new level object into group.
				pGroup->arLevels.Add(pLevel);
			}

			// Add the new group object into cached group data.
			m_arGroups.Add(pGroup);
		}

		// Load the column props.

		int nCols;
		Col * pCol;
		
		// Load column count first.
		ar >> nCols;

		// If column count is not zero, we have predefined layout.
		if (nCols)
			m_bHasLayout = TRUE;

		for (i = 0; i < nCols; i ++)
		{
			// Load each column props.

			// Create a new col object.
			pCol = new Col;
			pCol->arUserAttrib.SetSize(1);

			ar >> pCol->bAllowSizing;
			ar >> pCol->bVisible;
			ar >> pCol->clrBack >> pCol->clrFore >> pCol->clrHeaderBack >> pCol->clrHeaderFore;
			ar >> pCol->nAlignment >> pCol->nHeaderAlignment >> pCol->nCase;
			ar >> pCol->nWidth >> pCol->strTitle;
			ar >> pCol->arUserAttrib[0];
			ar >> pCol->strName;

			// Add the new col object into cached data.
			m_arCols.Add(pCol);
		}

		// Load cell data.

		CString strCell;
		int nCells;

		// Load cell count.
		ar >> nCells;
		Cell * pCell;

		while (nCells > 0)
		{
			// Create a new cell object.
			ar >> strCell;
			
			pCell = new Cell;
			pCell->strText = strCell;

			// Add the new cell object into cached data.
			m_arCells.Add(pCell);
			nCells --;
		}
	}
}

BOOL CBHMDBDropDownCtrl::OnSetObjectRects(LPCRECT lpRectPos, LPCRECT lpRectClip)
{
	CRect rt;
	rt.CopyRect(lpRectPos);
	rt.OffsetRect(-rt.left, -rt.top);

	// Move the internal grid.
	if (m_pGridDraw)
		m_pGridDraw->SetWindowPos(NULL, rt.left, rt.top, rt.Width(), rt.Height(), SWP_NOZORDER);
	
	return COleControl::OnSetObjectRects(lpRectPos, lpRectClip);
}

void CBHMDBDropDownCtrl::OnForeColorChanged() 
/*
Routine Description:
	The foreground color has been changed.

Parameters:
	None.

Return Value:
	None.
*/
{
	// Be careful in null pointer.

	if (m_pGridDraw)
		m_pGridDraw->SetForeColor(TranslateColor(GetForeColor()));
	else if (m_pDropDownRealWnd)
		m_pDropDownRealWnd->SetForeColor(TranslateColor(GetForeColor()));
	
	COleControl::OnForeColorChanged();
	BoundPropertyChanged(DISPID_FORECOLOR);
}

void CBHMDBDropDownCtrl::OnHeadBackColorChanged() 
/*
Routine Description:
	Update the control after the "HeadBackColor" property changed.

Parameters:
	None.

Return Value:
	None.
*/
{
	// Update the grid.

	if (m_pGridDraw)
	{
		m_pGridDraw->SetHeaderBackColor(TranslateColor(m_clrHeadFore));

		SetModifiedFlag();
		BoundPropertyChanged(dispidHeadForeColor);
		InvalidateControl();
	}
	else if (m_pDropDownRealWnd)
	{
		m_pDropDownRealWnd->SetHeaderBackColor(TranslateColor(m_clrHeadFore));
	}
}

void CBHMDBDropDownCtrl::OnHeadForeColorChanged() 
/*
Routine Description:
	Update the control after the "HeadForeColor" property changed.

Parameters:
	None.

Return Value:
	None.
*/
{
	if (m_pGridDraw)
	{
		m_pGridDraw->SetHeaderForeColor(TranslateColor(m_clrHeadBack));

		SetModifiedFlag();
		BoundPropertyChanged(dispidHeadBackColor);
		InvalidateControl();
	}
	else if (m_pDropDownRealWnd)
	{
		m_pDropDownRealWnd->SetHeaderForeColor(TranslateColor(m_clrHeadBack));
	}
}

void CBHMDBDropDownCtrl::ClearInterfaces()
/*
Routine Description:
	Clear the cached interface pointers.

Parameters:
	None.

Return Value:
	None.
*/
{
	// Free column props memory.

	FreeColumnMemory();
	FreeColumnData();

	// Disconnect the event sinks.
	if (m_dwCookieRN != 0)
	{
		AfxConnectionUnadvise(m_Rowset.m_spRowset, IID_IRowsetNotify, &m_xRowsetNotify, FALSE, m_dwCookieRN);
		m_dwCookieRN = 0;
	}
	
	if (m_spDataSource && AmbientUserMode())
		m_spDataSource->removeDataSourceListener((IDataSourceListener*)&m_xDataSourceListener);

	if (m_spRowPosition)
	{
		m_spRowPosition->Release();
		m_spRowPosition = NULL;
	}

	// Free column bind information.

	m_nColumnsBind = 0;
	
	for (int i = 0; i < m_arBindInfo.GetSize(); i ++)
		delete m_arBindInfo[i];

	m_arBindInfo.RemoveAll();
}

void CBHMDBDropDownCtrl::FreeColumnMemory()
/*
Routine Description:
	Free the cached data source field information.

Parameters:
	None.

Return Value:
	None.
*/
{
	// Free column info memory.

	if (m_pColumnInfo)
	{
		CoTaskMemFree(m_pColumnInfo);
		m_pColumnInfo = NULL;
	}

	// Free string memory.
	if (m_pStrings)
	{
		CoTaskMemFree(m_pStrings);
		m_pStrings = NULL;
	}

	// Free bookmark memory.
	FreeBookmarkMemory();
}

void CBHMDBDropDownCtrl::InitGridFromProp()
/*
Routine Description:
	Init the grid from cached prop data.

Parameters:
	None.

Return Value:
	None.
*/
{
	// Be careful in null pointer.

	BOOL bLockUpdate;
	if (m_pDropDownRealWnd)
	{
		// Save the update permission before changing it.
		bLockUpdate = m_pDropDownRealWnd->IsLockUpdate();
	
		// Prevent the grid from update.
		m_pDropDownRealWnd->LockUpdate();

		// Clear the grid.
		m_pDropDownRealWnd->SetColCount(0);
		m_pDropDownRealWnd->SetRowCount(0);
		m_pDropDownRealWnd->SetGroupCount(0);

		// Load the layout from cache.
		m_pDropDownRealWnd->LoadLayout(&m_arGroups, &m_arCols, m_nDataMode? &m_arCells : NULL);
	}
	else if (m_pGridDraw)
	{
		// Save the update permission before changing it.
		bLockUpdate = m_pGridDraw->IsLockUpdate();
	
		// Prevent the grid from update.
		m_pGridDraw->LockUpdate();

		// Clear the grid.
		m_pGridDraw->SetColCount(0);
		m_pGridDraw->SetRowCount(0);
		m_pGridDraw->SetGroupCount(0);

		// Load the layout from cache.
		m_pGridDraw->LoadLayout(&m_arGroups, &m_arCols, NULL);

		// If we have no predefined layout and the data mode is manual mode, insert some 
		// blank rows and cols.
		m_pGridDraw->SetRowCount(m_nRows);

		if (!m_pGridDraw->GetGroupCount() && !m_pGridDraw->GetColCount())
			m_pGridDraw->SetColCount(m_nCols);
	}
	
	if (m_pDropDownRealWnd)
	{
		// Restore the redraw permission.
		m_pDropDownRealWnd->LockUpdate(bLockUpdate);
	
		// Update the variables.
		m_nRows = m_pDropDownRealWnd->GetRowCount();
		m_nCols = m_pDropDownRealWnd->GetColCount();
	}
	else if (m_pGridDraw)
	{
		// Restore the redraw permission.
		m_pGridDraw->LockUpdate(bLockUpdate);
	
		// Update the variables.
		m_nRows = m_pGridDraw->GetRowCount();
		m_nCols = m_pGridDraw->GetColCount();
	}

	InvalidateControl();

	BoundPropertyChanged(dispidRows);
	BoundPropertyChanged(dispidCols);

	SetModifiedFlag();
}

BOOL CBHMDBDropDownCtrl::GetColumnInfo()
/*
Routine Description:
	Get the field information from data source. Note that we do not support some data type, 
	so filter them out.

Parameters:
	None.

Return Value:
	If succeeded, return TRUE; Otherwise, return FALSE.
*/
{
	USES_CONVERSION;

	// Clear cached pointers.
	ClearInterfaces();

	HRESULT hr;
	int nColumn = 0;

	// We can't do anything if the datasource hasn't been set
	if (m_spDataSource == NULL)
		return FALSE;

	m_nColumns = 0; // Reset the column count
	// Get the rowset from the data control
	hr = GetRowset();
	if (FAILED(hr) || !m_Rowset.m_spRowset)
		return FALSE;

	// Get the column information so we know the column names etc.
	hr = m_Rowset.CAccessorRowset<CManualAccessor>::GetColumnInfo((ULONG *)&m_nColumns, &m_pColumnInfo, &m_pStrings);
	if (FAILED(hr) || !m_pColumnInfo)
		return FALSE;

	// Check if we have bookmarks we can store
	if (m_pColumnInfo->iOrdinal == 0)
	{
		m_bHaveBookmarks = TRUE;
		FreeBookmarkMemory();
		m_nBookmarkSize = m_pColumnInfo->ulColumnSize;
		nColumn = 1;
	}
	else
	{
		m_bHaveBookmarks = FALSE;

		AfxMessageBox(IDS_ERROR_NOBOOKMARK, MB_ICONERROR | MB_OK);
		return FALSE;
	}

	// Now we can create the accessor and bind the data
	// First of all we need to find the column ordinal from the name
	int nColumnsBound = 0;
	for (; nColumn < m_nColumns; nColumn++)
	{
		switch (m_pColumnInfo[nColumn].wType)
		{
		case DBTYPE_I2:
		case DBTYPE_I4:
		case DBTYPE_R4:
		case DBTYPE_R8:
		case DBTYPE_CY:
		case DBTYPE_DATE:
		case DBTYPE_BSTR:
		case DBTYPE_BOOL:
		case DBTYPE_DECIMAL:
		case DBTYPE_UI1:
		case DBTYPE_I1:
		case DBTYPE_UI2:
		case DBTYPE_UI4:
		case DBTYPE_I8:
		case DBTYPE_UI8:
		case DBTYPE_FILETIME:
		case DBTYPE_STR:
		case DBTYPE_WSTR:
		case DBTYPE_NUMERIC:
		case DBTYPE_UDT:
		case DBTYPE_DBDATE:
		case DBTYPE_DBTIME:
		case DBTYPE_DBTIMESTAMP:
			nColumnsBound ++;
			if (m_nColumnsBind < 254)
			{
				ColumnData * pCld = new ColumnData;
				pCld->strName = OLE2T(m_pColumnInfo[nColumn].pwszName);
				pCld->nColumn = nColumn;
				pCld->vt = m_pColumnInfo[nColumn].wType;
				m_apColumnData.Add(pCld);
			}
			break;
		}
	}

	m_nColumnsBind = min(m_nColumnsBind, m_apColumnData.GetSize());

	return m_apColumnData.GetSize() > 0;
}

HRESULT CBHMDBDropDownCtrl::GetRowset()
/*
Routine Description:
	Get the rowset pointer from data source.

Parameters:
	None.

Return Value:
	Standard result code.
*/
{
	USES_CONVERSION;

	// Data source can not be null.
	ASSERT(m_spDataSource != NULL);
	HRESULT hr;

	// Close anything we had open and ClearInterfaces if necessary
	ClearInterfaces();

	// Close the cached rowset.
	m_Rowset.CAccessorRowset<CManualAccessor>::Close();

	// Get new rowset.

	hr = m_spDataSource->getDataMember(m_strDataMember.AllocSysString(), IID_IRowPosition,
		(IUnknown **)&m_spRowPosition);
	if (FAILED(hr) || !m_spRowPosition)
		return hr;

	return m_spRowPosition->GetRowset(IID_IRowset, (IUnknown**)&m_Rowset.m_spRowset);
}

void CBHMDBDropDownCtrl::InitDropDownWnd()
/*
Routine Description:
	Init the dropdown window before showing it.

Parameters:
	None.

Return Value:
	None.
*/
{
	// Update divider lines.
	UpdateDivider();

	// Update the list mode.
	m_pDropDownRealWnd->SetListMode(TRUE);

	InitHeaderFont();
	InitCellFont();
	
	m_pDropDownRealWnd->SetHeaderHeight(m_nHeaderHeight);
	m_pDropDownRealWnd->SetRowHeight(m_nRowHeight);
	m_pDropDownRealWnd->SetDefColWidth(m_nDefColWidth);
	m_pDropDownRealWnd->SetShowGroupHeader(m_bGroupHeader);

	InitGridFromProp();
	
	m_pDropDownRealWnd->SetReadOnly(TRUE);
	m_pDropDownRealWnd->SetShowColHeader(m_bColumnHeader);
	m_pDropDownRealWnd->SetShowRowHeader(FALSE);
}

BOOL CBHMDBDropDownCtrl::GetColumnHeader() 
/*
Routine Description:
	Get the visibility of column header.

Parameters:
	None.

Return Value:
	The visibility of column header.
*/
{
	return m_bColumnHeader;
}

void CBHMDBDropDownCtrl::SetColumnHeader(BOOL bNewValue) 
/*
Routine Description:
	Set the visibility of column header.

Parameters:
	bNewValue		The new value. If is TRUE, the grid should show column header;
	Otherwise, it should hide the column header.

Return Value:
	None.
*/
{
	// Save the new value.
	m_bColumnHeader = bNewValue;

	// Be careful in null pointer.
	if (m_pGridDraw)
		m_pGridDraw->SetShowColHeader(m_bColumnHeader);
	else if (m_pDropDownRealWnd)
	{
		m_pDropDownRealWnd->SetShowColHeader(m_bColumnHeader);
	}

	SetModifiedFlag();
	BoundPropertyChanged(dispidColumnHeader);
	InvalidateControl();
}

LPFONTDISP CBHMDBDropDownCtrl::GetHeadFont() 
/*
Routine Description:
	Get the font used in header.

Parameters:
	None.

Return Value:
	The pointer of IFontDispatch.
*/
{
	return m_fntHeader.GetFontDispatch();
}

void CBHMDBDropDownCtrl::SetHeadFont(LPFONTDISP newValue) 
/*
Routine Description:
	Set the font used in header.

Parameters:
	newValue		The new font object.

Return Value:
	None.
*/
{
	// Be careful in null pointer.
	ASSERT((newValue == NULL) ||
		   AfxIsValidAddress(newValue, sizeof(newValue), FALSE));

	// Init the cached value.
	m_fntHeader.InitializeFont(NULL, newValue);

	// Update the grid.
	InitHeaderFont();

	// Update the row height after changing font.
	if (m_pGridDraw)
		SetHeaderHeight(m_pGridDraw->GetRowHeight(1));
	else if (m_pDropDownRealWnd)
		SetHeaderHeight(m_pDropDownRealWnd->GetRowHeight(1));

	SetModifiedFlag();
	BoundPropertyChanged(dispidHeadFont);
	InvalidateControl();
}

long CBHMDBDropDownCtrl::GetHeaderHeight() 
/*
Routine Description:
	Get the height of header row.

Parameters:
	None.

Return Value:
	The height of header row.
*/
{
	if (m_pDropDownRealWnd)
		m_nHeaderHeight = m_pDropDownRealWnd->GetRowHeight(0);
	else if (m_pGridDraw)
		m_nHeaderHeight = m_pGridDraw->GetRowHeight(0);

	return m_nHeaderHeight;
}

void CBHMDBDropDownCtrl::SetHeaderHeight(long nNewValue) 
/*
Routine Description:
	Set the height of header row.

Parameters:
	nNewValue		The new value.

Return Value:
	None.
*/
{
	// Can not less than 0.
	if (nNewValue < 0)
		ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDWIDTHORHEIGHT);

	// Save the new value.
	m_nHeaderHeight = nNewValue;

	// Be careful in null pointer.
	if (m_pGridDraw)
		m_pGridDraw->SetHeaderHeight(m_nHeaderHeight);
	else if (m_pDropDownRealWnd)
		m_pDropDownRealWnd->SetHeaderHeight(m_nHeaderHeight);

	SetModifiedFlag();
	BoundPropertyChanged(dispidHeaderHeight);
	InvalidateControl();
}

void CBHMDBDropDownCtrl::InitHeaderFont()
/*
Routine Description:
	Init the font used in header row.

Parameters:
	None.

Return Value:
	None.
*/
{
	HFONT hFont;
	
	// Get the font handle from IFont pointer.
	if (m_fntHeader.m_pFont->get_hFont(&hFont) == S_OK)
	{
		// Get LOGFONT data from font handle.
		LOGFONT logHeader;
		if (GetObject(hFont, sizeof(logHeader), &logHeader))
		{
			// Update the grid.
			if (m_pGridDraw)
			{
				m_pGridDraw->SetHeaderFont(&logHeader);
			}
			else if (m_pDropDownRealWnd)
			{
				m_pDropDownRealWnd->SetHeaderFont(&logHeader);
			}
		}
	}
}

void CBHMDBDropDownCtrl::GetCellValue(int nRow, int nCol, CString& strText)
/*
Routine Description:
	Get the value of specified cell.

Parameters:
	nRow		The row ordinal.
	nCol		The col ordinal.
	strText		The buffer receives result value.

Return Value:
	None.
*/
{
	// This function has no means at design time.
	if (!AmbientUserMode())
		return;

	// This property has not means if the data mode is not bind mode.
	if (m_nDataMode != 0)
		return;
	
	// If we are loading the cell at the last row, we can check if we have new rows to fetch.
	if (!m_bEndReached && nRow == (int)m_pDropDownRealWnd->OnGetRecordCount())
		FetchNewRows();

	// The required row is unavailable yet, omit it.
	if ((int)m_apBookmark.GetSize() < nRow)
		return;

	// Get the bookmark of the desired row.

	CBookmark<0> bmk;
	HRESULT hr;

	GetBmkOfRow(nRow);
	if (!m_apBookmark[nRow - 1])
		return;

	// Get data from this record.	
	bmk.SetBookmark(m_nBookmarkSize, m_apBookmark[nRow - 1]);
	hr = m_Rowset.MoveToBookmark(bmk);
	
	if (FAILED(hr) || hr == DB_S_ENDOFROWSET || m_Rowset.m_hRow == NULL)
	{
		m_Rowset.FreeRecordMemory();
		return;
	}
	else if (hr != S_OK)
	// We got some error, but the row data is available anyway.
		m_Rowset.GetData();
	
	// The column prop object.
	ColumnProp * pCol;

	// The col ordinal of desired cell.
	int nColToSet;
	for (int j = 0; j < m_arBindInfo.GetSize(); j ++)
	{
		// If this col is not bound to any field in data source, omit it.
		pCol = m_arBindInfo[j];
		if (pCol->nDataField == -1)
			continue;
		
		// Calc the col ordinal.
		nColToSet = m_pDropDownRealWnd->GetColFromIndex(pCol->nColIndex);
		if (nColToSet != nCol)
			continue;

		// This col is the target.
		if (m_data.dwStatus[j] == DBSTATUS_S_OK)
		// The data is available.
			strText = m_data.strData[j];
		else if (m_data.dwStatus[j] == DBSTATUS_S_ISNULL)
		// The data is null.
			strText = _T("");
	}
	
	// Release current record.
	m_Rowset.FreeRecordMemory();
}

LPUNKNOWN CBHMDBDropDownCtrl::GetDataSource() 
/*
Routine Description:
	Get the data source object.

Parameters:
	None.

Return Value:
	The pointer of the datasource interface.
*/
{
	// AddRef() before return.
	if (m_spDataSource)
		m_spDataSource->AddRef();

	return m_spDataSource;
}

void CBHMDBDropDownCtrl::SetDataSource(LPUNKNOWN newValue) 
/*
Routine Description:
	Set the data source object.

Parameters:
	newValue		The new pointer.

Return Value:
	None.
*/
{
	// Clear the cached interface pointers.
	ClearInterfaces();

	if (m_spDataSource)
	{
		m_spDataSource->Release();
		m_spDataSource = NULL;
	}

	// Can not set datasource in manual mode.
	if (m_nDataMode != 0 && newValue)
	{
		ThrowError(CTL_E_SETNOTSUPPORTED, _T("Can not set datasource in manual mode"));
		return;
	}

	if (newValue == m_spDataSource)
		return;

	if (AmbientUserMode())
	{
		// Datasource has been changed.
		m_bDataSrcChanged = TRUE;
		m_bEndReached = FALSE;
	}

	if (!newValue)
	{
		m_spDataSource = NULL;
		return;
	}

	// Get IDataSource pointer.
	if (FAILED(newValue->QueryInterface(__uuidof(m_spDataSource), (void **)&m_spDataSource)))
		return;

	// Retrieve the fields information.
	if (!AmbientUserMode())
		GetColumnInfo();

	SetModifiedFlag();
}

BSTR CBHMDBDropDownCtrl::GetDataMember() 
/*
Routine Description:
	Get the datamember string.

Parameters:
	None.

Return Value:
	The data member string.
*/
{
	return m_strDataMember.AllocSysString();
}

void CBHMDBDropDownCtrl::SetDataMember(LPCTSTR lpszNewValue) 
/*
Routine Description:
	Set the data member string.

Parameters:
	lpszNewValue		The new value.

Return Value:
	None.
*/
{
	// Can not set datamember in manual mode.
	if (m_nDataMode != 0)
		ThrowError(CTL_E_SETNOTSUPPORTED, _T("Can not set datasource in manual mode"));

	if (m_strDataMember == lpszNewValue)
		return;

	// Save the new value.
	m_strDataMember = lpszNewValue;

	// Clear the cached interface pointers.
	ClearInterfaces();

	SetModifiedFlag();
	
	// If at design time, retrieve the fields information only; Otherwise, tell everyone that
	// the datasource has been changed.
	if (!AmbientUserMode())
		GetColumnInfo();
	else
		m_bDataSrcChanged = TRUE;
}

BSTR CBHMDBDropDownCtrl::GetFieldSeparator() 
/*
Routine Description:
	Get the string used as field separator in AddItem() funciton.

Parameters:
	None.

Return Value:
	The field separator string.
*/
{
	return m_strFieldSeparator.AllocSysString();
}

void CBHMDBDropDownCtrl::SetFieldSeparator(LPCTSTR lpszNewValue) 
/*
Routine Description:
	Set the string used as field separator in AddItem() function.

Parameters:
	lpszNewValue		The new value.

Return Value:
	None.
*/
{
	// The new value should be one char long or be one of the predefined string.
	if (lstrlen(lpszNewValue) == 0 || (lstrlen(lpszNewValue) > 1 && lstrcmpi(
		lpszNewValue, _T("(space)")) && lstrcmpi(lpszNewValue, _T("(tab)"))))
		ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_FIELDSEPARATOR);

	// Save the new value.
	m_strFieldSeparator = lpszNewValue;

	// Calc the real used value.
	m_strSeparator = lpszNewValue;

	// "(tab)" means "\t", "(space)" means " ".
	if (!m_strFieldSeparator.CompareNoCase(_T("(tab)")))
		m_strSeparator = _T("\t");
	else if (!m_strFieldSeparator.CompareNoCase(_T("(space)")))
		m_strSeparator = _T(" ");

	SetModifiedFlag();
	BoundPropertyChanged(dispidFieldSeparator);
}

void CBHMDBDropDownCtrl::AddItem(LPCTSTR Item, const VARIANT FAR& RowIndex) 
/*
Routine Description:
	Add one row into grid, fill the cells.

Parameters:
	Item		The data used to fill the new row.

	RowIndex	The row before which to insert the new row.

Return Value:
	None.
*/
{
	// Can not insert row in bound mode.
	if (m_nDataMode == 0)
		ThrowError(CTL_E_PERMISSIONDENIED, IDS_ERROR_NOAPPLICABLEINBINDMODE);

	// Calc the row ordinal.
	int nRow = GetRowColFromVariant(RowIndex);
	// Insert at the end default.
	if (nRow == INVALID)
		nRow = m_pDropDownRealWnd->GetRowCount();

	if (nRow == 0 || nRow > (int)m_pDropDownRealWnd->OnGetRecordCount() + 1)
		ThrowError(CTL_E_INVALIDPROPERTYARRAYINDEX, AFX_IDP_E_INVALIDPROPERTYARRAYINDEX);

	CString strText = Item, strField;
	int i = 1, j = 0;

	// Insert a row.
	m_pDropDownRealWnd->InsertRow(nRow);
	strField.Empty();

	while (j < strText.GetLength() && (int)i <= m_pDropDownRealWnd->GetColCount())
	{
		// Retrieve the value of each cell.

		if (strText[j] != m_strSeparator)
			strField += strText[j];
		else
		{
			// Fill this cell.
			m_pDropDownRealWnd->SetCellText(nRow, i, strField);
			strField.Empty();
			i ++;
		}

		j ++;
	}

	// Fill the last cell.
	if ((int)i <= m_pDropDownRealWnd->GetColCount())
		m_pDropDownRealWnd->SetCellText(nRow, i, strField);

	m_pDropDownRealWnd->RedrawRow(nRow);
}

void CBHMDBDropDownCtrl::RemoveItem(long RowIndex) 
/*
Routine Description:
	Remove one row from grid.

Parameters:
	RowIndex		The ordinal of the row to remove.

Return Value:
	None.
*/
{
	// Can not remove row in bound mode.
	if (m_nDataMode == 0)
		ThrowError(CTL_E_PERMISSIONDENIED, IDS_ERROR_NOAPPLICABLEINBINDMODE);

	int nRow = RowIndex;
	if (nRow == 0 || nRow > m_pDropDownRealWnd->GetRowCount())
		ThrowError(CTL_E_INVALIDPROPERTYARRAYINDEX, AFX_IDP_E_INVALIDPROPERTYARRAYINDEX);

	// Update the grid.
	m_pDropDownRealWnd->RemoveRow(nRow);
}

void CBHMDBDropDownCtrl::Scroll(long Rows, long Cols) 
/*
Routine Description:
	Scroll the grid.

Parameters:
	Rows		Rows to scroll up.
	Cols		Cols to scroll left.

Return Value:
	None.
*/
{
	// Send messages to grid to scroll.

	if (Rows > 0)
	{
		for (int i = 0; i < Rows; i ++)
			m_pDropDownRealWnd->SendMessage(WM_VSCROLL, SB_LINEDOWN);
	}
	else
	{
		for (int i = 0; i > Rows; i --)
			m_pDropDownRealWnd->SendMessage(WM_VSCROLL, SB_LINEUP);
	}
	if (Cols > 0)
	{
		for (int i = 0; i < Cols; i ++)
			m_pDropDownRealWnd->SendMessage(WM_HSCROLL, SB_LINERIGHT);
	}
	else
	{
		for (int i = 0; i > Cols; i --)
			m_pDropDownRealWnd->SendMessage(WM_HSCROLL, SB_LINELEFT);
	}
}

void CBHMDBDropDownCtrl::RemoveAll() 
/*
Routine Description:
	Remove all items from grid.

Parameters:
	None.

Return Value:
	None.
*/
{
	// Can not remove items in bound mode.
	if (m_nDataMode == 0)
		ThrowError(CTL_E_PERMISSIONDENIED, IDS_ERROR_NOAPPLICABLEINBINDMODE);

	// Update the grid.
	m_pDropDownRealWnd->SetRowCount(0);
}

void CBHMDBDropDownCtrl::MoveFirst() 
/*
Routine Description:
	Set the first row as current row.

Parameters:
	None.

Return Value:
	None.
*/
{
	if (m_bDataSrcChanged)
		UpdateDropDownWnd();

	// Update the grid.
	m_pDropDownRealWnd->MoveTo(1);
}

void CBHMDBDropDownCtrl::MovePrevious() 
/*
Routine Description:
	Set the prior row as the current row.

Parameters:
	None.

Return Value:
	None.
*/
{
	if (m_bDataSrcChanged)
		UpdateDropDownWnd();

	// Update the grid.
	int nRow, nCol;
	m_pDropDownRealWnd->GetCurrentCell(nRow, nCol);
	m_pDropDownRealWnd->MoveTo(nRow - 1);
}

void CBHMDBDropDownCtrl::MoveNext() 
/*
Routine Description:
	Set the next row as current row.

Parameters:
	None.

Return Value:
	None.
*/
{
	if (m_bDataSrcChanged)
		UpdateDropDownWnd();

	// Update the grid.
	int nRow, nCol;
	m_pDropDownRealWnd->GetCurrentCell(nRow, nCol);
	m_pDropDownRealWnd->MoveTo(nRow + 1);
}

void CBHMDBDropDownCtrl::MoveLast() 
/*
Routine Description:
	Set the last row as current row.

Parameters:
	None.

Return Value:
	None.
*/
{
	if (m_bDataSrcChanged)
		UpdateDropDownWnd();

	// Update the grid.
	m_pDropDownRealWnd->MoveTo(m_pDropDownRealWnd->OnGetRecordCount());
}

void CBHMDBDropDownCtrl::MoveRecords(long Rows) 
/*
Routine Description:
	Move current row by given rows.

Parameters:
	Rows		Rows to move.

Return Value:
	None.
*/
{
	if (m_bDataSrcChanged)
		UpdateDropDownWnd();

	// Update the grid.
	int nRow, nCol;
	m_pDropDownRealWnd->GetCurrentCell(nRow, nCol);
	m_pDropDownRealWnd->MoveTo(nRow + Rows);
}

void CBHMDBDropDownCtrl::ReBind() 
/*
Routine Description:
	Rebind to data source.

Parameters:
	None.

Return Value:
	None.
*/
{
	// Simply set the signal.
	m_bDataSrcChanged = TRUE;
}

int CBHMDBDropDownCtrl::GetRowColFromVariant(VARIANT va)
/*
Routine Description:
	Get row/col ordinal from give variant.

Parameters:
	va		The variant from which calc the row/col ordinal.

Return Value:
	None.
*/
{
	// We only accept string format data and simply convert it into long type.
	if (va.vt == VT_ERROR)
		return INVALID;

	COleVariant vaNew = va;
	vaNew.ChangeType(VT_BSTR);
	vaNew.ChangeType(VT_I4);

	return vaNew.lVal;
}

BOOL CBHMDBDropDownCtrl::OnGetPredefinedStrings(DISPID dispid, CStringArray* pStringArray, CDWordArray* pCookieArray) 
{
	int i;

	switch (dispid)
	{
	case dispidFieldSeparator:
	// Fill the "FieldSeparator" choice list.
		pStringArray->Add(_T("(tab)"));
		pStringArray->Add(_T("(space)"));
		pCookieArray->Add(0);
		pCookieArray->Add(1);

		return TRUE;
		break;

	case dispidDataField:
	// Fill the "DataField" choice list.
		if (m_arCols.GetSize() == 0)
		{
			// Has no layout.

			// Get the bounded fields information.
			GetColumnInfo();

			// Insert each field into list.
			for (i = 0; i < m_apColumnData.GetSize(); i ++)
			{
				pStringArray->Add(m_apColumnData[i]->strName);
				pCookieArray->Add(i);
			}
		}
		else
		{
			// Has layout.

			// Insert each column into list.
			for (i = 0; i < m_arCols.GetSize(); i ++)
			{
				pStringArray->Add(m_arCols[i]->strName);
				pCookieArray->Add(i);
			}
		}

		return TRUE;
		break;
	}
	
	return COleControl::OnGetPredefinedStrings(dispid, pStringArray, pCookieArray);
}

BOOL CBHMDBDropDownCtrl::OnGetPredefinedValue(DISPID dispid, DWORD dwCookie, VARIANT* lpvarOut) 
{
	switch (dispid)
	{
	case dispidFieldSeparator:
	// Get the value of "FieldSeparator".

		VariantInit(lpvarOut);
		lpvarOut->vt = VT_BSTR;
		if (dwCookie == 0)
			lpvarOut->bstrVal = CString(_T("(tab)")).AllocSysString();
		else if (dwCookie == 1)
			lpvarOut->bstrVal = CString(_T("(space)")).AllocSysString();

		return TRUE;
		break;

	case dispidDataField:
	// Get the value of "DataField"

		VariantClear(lpvarOut);
		V_VT(lpvarOut) = VT_BSTR;
		if (m_arCols.GetSize() == 0)
		{
			// The value comes from data source.

			// Retrieve the fields information.
			GetColumnInfo();
			if (dwCookie >= 0 && dwCookie < (DWORD)m_apColumnData.GetSize())
				V_BSTR(lpvarOut) = m_apColumnData[dwCookie]->strName.AllocSysString();
		}
		else
		{
			// The value comes from predefined cols.
			if (dwCookie >= 0 && dwCookie < (DWORD)m_arCols.GetSize())
				V_BSTR(lpvarOut) = m_arCols[dwCookie]->strName.AllocSysString();
		}
		
		return TRUE;
		break;
	}
	
	return COleControl::OnGetPredefinedValue(dispid, dwCookie, lpvarOut);
}

int CBHMDBDropDownCtrl::GetRowFromBmk(BYTE *bmk)
/*
Routine Description:
	Get the row ordinal from bookmark.

Parameters:
	bmk		The bookmark to search.

Return Value:
	If find the bookmark, return the row ordinal; Otherwise, return INVALID.
*/
{
	// Save the given bookmark.
	BYTE * bmkSearch = new BYTE[m_nBookmarkSize];
	memcpy(bmkSearch, bmk, sizeof(bmk));

	int i, nRows = m_apBookmark.GetSize();
	for (i = 0; i < nRows; i++)
	{
		// Compare each bookmark with the given one.

		// Get the bookmark of this row.
		GetBmkOfRow(i + 1);
		if (m_apBookmark[i] && memcmp(m_apBookmark[i], bmkSearch, m_pColumnInfo->ulColumnSize) == 0)
		{
			// We got it!
			delete [] bmkSearch;
			return i + 1;
		}
	}
	
	// The given bookmark does not exist.
	delete [] bmkSearch;
	return INVALID;
}

int CBHMDBDropDownCtrl::GetRowFromHRow(HROW hRow)
/*
Routine Description:
	Find the row ordinal from the given HROW data.

Parameters:
	nRow		The HROW data to search.

Return Value:
	If found, return the row ordinal; Otherwise, return INVALID.
*/
{
	// HROW can not be 0.
	if (hRow == 0)
		return INVALID;

	// Save current state.
	HROW hRowOld = m_Rowset.m_hRow;

	// Assign current HROW to the given data.
	m_Rowset.m_hRow = hRow;
	HRESULT hr;

	// Get the bookmark of current row.
	hr = m_Rowset.GetData();
	if (FAILED(hr))
	{
		// The given bookmark is invalid.

		// Restore cached state.
		m_Rowset.m_hRow = hRowOld;
		return INVALID;
	}

	// Calc the row ordinal from bookmark.
	int nRow;
	nRow = GetRowFromBmk(m_data.bookmark);

	// Restore cached state.
	m_Rowset.FreeRecordMemory();
	m_Rowset.m_hRow = hRowOld;

	return nRow;
}

void CBHMDBDropDownCtrl::FetchNewRows()
/*
Routine Description:
	Fecth new rows from datasource if available.

Parameters:
	None.

Return Value:
	None.
*/
{
	USES_CONVERSION;

	// Cancel any pending modification first.
	m_pDropDownRealWnd->CancelRecord();

	// This function is useful when bound with RDC datasource, because
	// the nRowsAvailable is just the number of the records in the local
	// cache.
	int i = 0, nFetchOnce;
	ULONG nRowsAvailable = 0;
	if (m_Rowset.GetApproximatePosition(NULL, NULL, &nRowsAvailable) == S_OK)
		nFetchOnce = nRowsAvailable - m_apBookmark.GetSize();
	else
		return;

	if (nFetchOnce == 0)
	{
		// No new rows.
		m_bEndReached = TRUE;
		return;
	}

	// Extend the bookmark array first, fill the new entries with null.
	for (i = 0; i < nFetchOnce; i ++)
		m_apBookmark.Add(NULL);

	// Calc the rows need to fetch.
	nFetchOnce = nRowsAvailable - m_pDropDownRealWnd->OnGetRecordCount();
	if (nFetchOnce <= 0)
		return;
	
	// Simply insert new blank rows into grid.	
	m_pDropDownRealWnd->SetRowCount(m_pDropDownRealWnd->GetRowCount() + nFetchOnce);

	m_pDropDownRealWnd->Invalidate();
	m_pDropDownRealWnd->ResetScrollBars();
}

void CBHMDBDropDownCtrl::UndeleteRecordFromHRow(HROW hRow)
/*
Routine Description:
	Undelete the record from give HROW data.

Parameters:
	hRow		The HROW of the record to be restored.

Return Value:
	None.
*/
{
	// Get the bookmark of this row.
	m_Rowset.ReleaseRows();
	m_Rowset.m_hRow = hRow;
	if (FAILED(m_Rowset.GetData()))
		return;

	ULONG nRow;
	CBookmark<0> bmk;
	bmk.SetBookmark(m_nBookmarkSize, m_data.bookmark);

	// Get the original ordinal of this record.
	if (FAILED(m_Rowset.GetApproximatePosition(&bmk, &nRow, NULL) || 
		(int)nRow > m_pDropDownRealWnd->GetRowCount()))
		// The original ordinal is beyond the range we hold, simply add this at the end.
		nRow = m_pDropDownRealWnd->GetRowCount();

	BYTE * pBookmarkCopy;
	if (m_bHaveBookmarks)
	{
		// We must have bookmark to work.

		pBookmarkCopy = new BYTE[m_nBookmarkSize];
		if (pBookmarkCopy != NULL)
			memcpy(pBookmarkCopy, &m_data.bookmark, m_nBookmarkSize);
		
		// Insert the new bookmark into cached array.		
		m_apBookmark.InsertAt(nRow - 1, pBookmarkCopy);
	}
	
	// Insert this row into grid at the original position.
	m_pDropDownRealWnd->InsertRow(nRow);
	
	m_Rowset.FreeRecordMemory();

	// Update the grid.
	m_pDropDownRealWnd->Invalidate();
}

void CBHMDBDropDownCtrl::GetBmkOfRow(int nRow)
/*
Routine Description:
	Get the bookmark of one row.

Parameters:
	nRow		The desired row ordinal.

Return Value:
	None.
*/
{
	// If the bookmark already exist or the given row ordial is invalid, just return.
	if (nRow < 1 || nRow > (int)m_apBookmark.GetSize() || m_apBookmark[nRow - 1])
		return;

	BYTE * pBookmarkCopy;
	DBBOOKMARK bmFirst = DBBMK_FIRST;
	CBookmark<0> bmk;
	
	// Get data from this record.	
	bmk.SetBookmark(m_nBookmarkSize, (BYTE *)&bmFirst);
	HRESULT hr = m_Rowset.MoveToBookmark(bmk, nRow - 1);
	if (FAILED(hr) || hr == DB_S_ENDOFROWSET || m_Rowset.m_hRow == NULL)
	{
		m_Rowset.FreeRecordMemory();
		return;
	}
	else if (hr != S_OK)
	// We got the row data anyway.
		m_Rowset.GetData();
	
	if (m_bHaveBookmarks)
	{
		// We need the bookmark.
		pBookmarkCopy = new BYTE[m_nBookmarkSize];
		if (pBookmarkCopy != NULL)
			memcpy(pBookmarkCopy, &m_data.bookmark, m_nBookmarkSize);
		
		// Update the cached bookmark.
		m_apBookmark[nRow - 1] = pBookmarkCopy;
	}
}

void CBHMDBDropDownCtrl::OnBringMsg(long wParam, long lParam)
/*
Routine Description:
	Process the message comes from dropdown control.

Parameters:
	wParam		Standard param.

	lParam		Standard param.

Return Value:
	None.
*/
{
	// Retrieve the struct from wParam.
	BringInfo * pBi = (BringInfo *)wParam;
	if (!pBi)
		return;

	if (!pBi->hWnd)
	{
		// The comsumer tell me that I should close the dropdown window
		delete pBi;

		if (m_pDropDownRealWnd)
			HideDropDownWnd();

		m_bDroppedDown = FALSE;
		return;
	}

	// Encryption :)
	WORD x = LOWORD(pBi->wParam);
	WORD y = HIWORD(pBi->wParam);
	x ^= 235;
	y ^= 235;

	m_strText = pBi->strText;
	m_hConsumer = pBi->hWnd;
	if (x + y == (pBi->lParam ^ 235) && !m_bDroppedDown)
	// Show the dropdown window.
		Bring(x, y);
		
	// Free the memory allocated by message sender
	delete pBi;
}


void CBHMDBDropDownCtrl::HideDropDownWnd(LPCTSTR pCharValue)
/*
Routine Description:
	Hide the dropdown window.

Parameters:
	None.

Return Value:
	None.
*/
{
	if (::IsWindow(m_hConsumer))
	{
		// Send message to grid control, tell it that the window will be invisible.
		BringInfo * pBi = new BringInfo;

		pBi->wParam = FALSE;
		
		// Non null hWnd indicates that this message indicates the visibility of dropdown
		// window.
		pBi->hWnd = (HWND)0xff; 

		if (pCharValue)
		{
			// Return the result text back together.
			CString strValue = pCharValue;
			lstrcpy(pBi->strText, strValue);
			pBi->hWnd = NULL;
		}

		// Send the message to grid control.
		::SendMessage(m_hConsumer, CBHMDataApp::m_nBringMsg, (WPARAM)pBi, NULL);
	}

	// Hide the dropdown window.
	m_pDropDownRealWnd->Showwindow(FALSE);
}

void CBHMDBDropDownCtrl::Bring(int cx, int cy)
/*
Routine Description:
	Show the dropdown window.

Parameters:
	None.

Return Value:
	None.
*/
{
	// If the data source has been changed, update the binding information and grid now.
	if(m_bDataSrcChanged)
		UpdateDropDownWnd();
	
	// We will show the dropdown window from the left bottom point.

	int x = 0, y = 0, yWndMin = 0, yWndMax = 0, xTotal, yTotal;

	// Can not exceed the border of screen.
	int nMaxWidth = GetSystemMetrics(SM_CXSCREEN) - cx;
	int nMaxHeight = GetSystemMetrics(SM_CYSCREEN) - cy;

	// The rect the window will occupies.
	CRect rcBounds(cx, cy, 0, 0);

	yWndMin = yWndMax = yTotal = m_pDropDownRealWnd->GetVirtualHeaderHeight() + 2;

	// Calc the total height.
	yTotal += m_nRowHeight * m_pDropDownRealWnd->GetRowCount();

	// Take the max and min dropdown items property into account.
	yWndMax += m_nRowHeight * m_nMaxDropDownItems;
	yWndMin += m_nRowHeight * m_nMinDropDownItems;

	// Calc the total width.
	xTotal = m_pDropDownRealWnd->GetVirtualWidth() + 2;

	y = min(yTotal, yWndMax);
	y = max(y, yWndMin);
	y = min(y, nMaxHeight);
	if (y < yTotal)
	{
		// We need to show the vertical scroll bar, reserve width for it.
		xTotal += GetSystemMetrics(SM_CXVSCROLL);
	}

	// Take the listautosize and listwidth property into acount.
	x = (m_bListWidthAutoSize || m_nListWidth == 0) ? xTotal : m_nListWidth;
	x = min(x, nMaxWidth);
	if (x < xTotal)
	{
		// We need to show the horizontal scroll bar, reserve height for it.
		y += GetSystemMetrics(SM_CYHSCROLL);
	}
	y = min(y, nMaxHeight);

	rcBounds.right = rcBounds.left + x;
	rcBounds.bottom = rcBounds.top + y;

	// Modify the size and position of the dropdown window.
	m_pDropDownRealWnd->MoveWindow(&rcBounds, FALSE);

	// Update the scroll bars.
	m_pDropDownRealWnd->ResetScrollBars();

	if (::IsWindow(m_hConsumer))
	{
		// Tell the grid control that the dropdown window will be shown.
		BringInfo * pBi = new BringInfo;

		// The dropdown window handle.
		pBi->hWnd = m_pDropDownRealWnd->m_hWnd;

		pBi->wParam = TRUE;
		::SendMessage(m_hConsumer, CBHMDataApp::m_nBringMsg, (WPARAM)pBi, NULL);
	}

	// Search the text being shown.
	SearchText(m_strText);

	if (m_pDropDownRealWnd->GetRowCount() > 0 && m_pDropDownRealWnd-> GetColCount() > 0)
		m_pDropDownRealWnd->MoveTo(m_nRow);

	// Bring the window to topmost.
	m_pDropDownRealWnd->SetWindowPos(&wndTopMost, 0, 0, 0, 0, SWP_NOSIZE | SWP_NOMOVE);

	// Show the window.
	m_pDropDownRealWnd->SetForegroundWindow();
	m_pDropDownRealWnd->Showwindow(TRUE);
}

void CBHMDBDropDownCtrl::UpdateDropDownWnd()
/*
Routine Description:
	UPdate the binding information.

Parameters:
	None.

Return Value:
	None.
*/
{
	USES_CONVERSION;

	// Do nothing at design time.
	if (!AmbientUserMode())
		return;

	m_pDropDownRealWnd->LockUpdate(TRUE);

	// We processed the change of data source.
	m_bDataSrcChanged = FALSE;

	// Retrieve field information.
	if (!GetColumnInfo())
	{
		m_pDropDownRealWnd->LockUpdate(FALSE);
		m_pDropDownRealWnd->Invalidate();

		return;
	}

	int i;

	// If we have not predefine layout, simply bind every field.
	if (!m_bHasLayout)
	{
		// Remove column binding array.

		i = m_arCols.GetSize();
		
		for (; i > 0; i --)
			delete m_arCols[i - 1];
		
		m_arCols.RemoveAll();
	}

	if (m_arCols.GetSize() == 0)
	{
		// We have not layout.

		Col * pCol;
		m_nColumnsBind = m_apColumnData.GetSize();

		for (i = 0; i < m_nColumnsBind; i ++)
		{
			// Bind every field.

			pCol = new Col;

			pCol->strName = pCol->strTitle = m_apColumnData[i]->strName;
			pCol->arUserAttrib.SetAtGrow(0, pCol->strTitle);
			pCol->arUserAttrib.SetAtGrow(1, _T(""));
			pCol->nWidth = m_nDefColWidth;

			m_arCols.Add(pCol);
		}
	}

	InitDropDownWnd();

	// Do the binding job.
	Bind();

	// Now fetch new rows according new binding information.
	m_Rowset.SetupOptionalRowsetInterfaces();
	FetchNewRows();
	
	// Set up the sink so we know when the rowset is repositioned.
	if (AmbientUserMode())
	{
		AfxConnectionAdvise(m_Rowset.m_spRowset, IID_IRowsetNotify, &m_xRowsetNotify, FALSE, &m_dwCookieRN);
		m_spDataSource->addDataSourceListener((IDataSourceListener*)&m_xDataSourceListener);
	}

	m_pDropDownRealWnd->LockUpdate(FALSE);
	m_pDropDownRealWnd->Invalidate();
}

void CBHMDBDropDownCtrl::OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags) 
{
	if (m_pDropDownRealWnd && m_pDropDownRealWnd->IsWindowVisible())
	// Send keyboard messages to dropdown window.
		m_pDropDownRealWnd->SendMessage(WM_KEYDOWN, nChar, 1);
	
	COleControl::OnKeyDown(nChar, nRepCnt, nFlags);
}

void CBHMDBDropDownCtrl::DeleteRowFromSrc(int nRow)
/*
Routine Description:
	One row has been deleted from data source by another user, remove this record from our grid
	also.

Parameters:
	nRow 		The row ordinal.

Return Value:
	None.
*/
{
	// Update the grid.
	m_pDropDownRealWnd->RemoveRow(nRow);

	// Remove bookmark from cached data.

	delete [] m_apBookmark[nRow - 1];
	m_apBookmark.RemoveAt(nRow - 1);
}

void CBHMDBDropDownCtrl::SearchText(CString strSearch)
/*
Routine Description:
	Search the given string in data source.

Parameters:
	strSearch		The string to search.

	bAutoPosition		If be TRUE, move current row to the found row.

Return Value:
	If found, return the ordinal of result row; Otherwise, return INVALID.
*/
{
	if (!m_bListAutoPosition)
		return;

	// Fire the event.
	FirePositionList(strSearch);

	// Calc the controling col ordinal.
	int nCol = m_pDropDownRealWnd->GetColFromName(m_strDataField);
	if (nCol == INVALID)
		return;

	// Calc the ordinal in bind information array.
	CString strValue;
	int nColIndex = m_pDropDownRealWnd->GetColIndex(nCol);

	for (int i = 0; i < m_arBindInfo.GetSize(); i ++)
		if (m_arBindInfo[i]->nColIndex == nColIndex)
			break;

	if (i >= m_arBindInfo.GetSize() || m_arBindInfo[i]->nDataField == -1)
	{
		// The controling is not bound to data source. Search the text manully.

		for (i = 1; (int)i <= m_pDropDownRealWnd->GetRowCount(); i ++)
		{
			strValue = m_pDropDownRealWnd->GetCellText(i, nCol);
			if (!strValue.Left(strSearch.GetLength()).CompareNoCase(strSearch))
			{
				m_nRow = i;
				break;
			}
		}
		
		return;
	}
	else
	{
		// Search the text from data source.

		if (m_Rowset.m_spRowset == NULL)
			return;

		// Get the IRowsetFind pointer.
		CComQIPtr<IRowsetFind, &IID_IRowsetFind> pFind = m_Rowset.m_spRowset;
		if (!pFind)
			return;

		DBBOOKMARK bmFirst = DBBMK_FIRST;
		CAccessorRowset<CManualAccessor> rac;
		rac.m_spRowset = m_Rowset.m_spRowset;
		RowData bindData;
	
		// Construct binding information for searching.
		rac.CreateAccessor(1, &bindData, sizeof(bindData));
		rac.AddBindEntry(m_arBindInfo[i]->nDataField, DBTYPE_STR, 255 * sizeof(TCHAR), 
			&(m_data.strData[0]), NULL, &(m_data.dwStatus[0]));

		// Fill the searching entry.
		if (strSearch.IsEmpty())
			m_data.dwStatus[0] = DBSTATUS_S_ISNULL;
		else
		{
			m_data.dwStatus[0] = DBSTATUS_S_OK;
			lstrcpy(m_data.strData[0], strSearch);
		}

		if (FAILED(rac.Bind()))
			return;

		// First use partly match search. Ff fails, use exactly match search.
		// Not efficient but simple :)
		ULONG nRowsOb = 0;
		HROW* phRow = &rac.m_hRow;
		pFind->FindNextRow(NULL, rac.m_pAccessor->GetHAccessor(0), rac.
			m_pAccessor->GetBuffer(), DBCOMPAREOPS_EQ | DBCOMPAREOPS_CASEINSENSITIVE, 
			m_nBookmarkSize, (BYTE *)&bmFirst, 0, 1, &nRowsOb, &phRow);

		if (rac.m_hRow == 0)
			pFind->FindNextRow(NULL, rac.m_pAccessor->GetHAccessor(0), rac.
				m_pAccessor->GetBuffer(), DBCOMPAREOPS_BEGINSWITH | DBCOMPAREOPS_CASEINSENSITIVE, 
				m_nBookmarkSize, (BYTE *)&bmFirst, 0, 1, &nRowsOb, &phRow);

		if (rac.m_hRow == 0)
			return;

		// Do the clearing work.
		HROW hRow = rac.m_hRow;
		rac.ReleaseRows();
		m_Rowset.ReleaseRows();
		m_Rowset.m_hRow = hRow;
		if (FAILED(m_Rowset.GetData()))
			return;

		ULONG nRow;
		CBookmark<0> bmk;
		bmk.SetBookmark(m_nBookmarkSize, m_data.bookmark);

		// Get the row ordinal.
		if (FAILED(m_Rowset.GetApproximatePosition(&bmk, &nRow, NULL) || 
			(int)nRow > m_pDropDownRealWnd->GetRowCount()))
		nRow = m_pDropDownRealWnd->GetRowCount();

		// Move to this row if required.
		m_nRow = nRow;

		return;
	}
}

void CBHMDBDropDownCtrl::OnListAutoPositionChanged() 
/*
Routine Description:
	The user changed the value of "ListAutoPosition" property.

Parameters:
	None.

Return Value:
	None.
*/
{
	SetModifiedFlag();
	BoundPropertyChanged(dispidListAutoPosition);
}

void CBHMDBDropDownCtrl::OnListWidthAutoSizeChanged() 
/*
Routine Description:
	The user changed the value of "ListWidthAutoSize" property.

Parameters:
	None.

Return Value:
	None.
*/
{
	SetModifiedFlag();
	BoundPropertyChanged(dispidListWidthAutoSize);
}

short CBHMDBDropDownCtrl::GetMaxDropDownItems() 
/*
Routine Description:
	Get the max items can be visible in dropdown window simultaneously.

Parameters:
	None.

Return Value:
	The max value.
*/
{
	return m_nMaxDropDownItems;
}

void CBHMDBDropDownCtrl::SetMaxDropDownItems(short nNewValue) 
/*
Routine Description:
	Set the max items can be visible in dropdown window simultaneously.

Parameters:
	nNewValue		The new value.

Return Value:
	None.
*/
{
	if (nNewValue < 1 || nNewValue > 100 || nNewValue < m_nMinDropDownItems)
		ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDITEMSLIMITATION);

	m_nMaxDropDownItems = nNewValue;
	SetModifiedFlag();
	BoundPropertyChanged(dispidMaxDropDownItems);
}

short CBHMDBDropDownCtrl::GetMinDropDownItems() 
/*
Routine Description:
	Get the min items can be visible in dropdown window simultaneously.

Parameters:
	None.

Return Value:
	The min count.
*/
{
	return m_nMinDropDownItems;
}

void CBHMDBDropDownCtrl::SetMinDropDownItems(short nNewValue) 
/*
Routine Description:
	Set the min items can be visible in dropdown window simultaneously.

Parameters:
	nNewValue		The new value.

Return Value:
	None.
*/
{
	// Verify the new value.
	if (nNewValue < 1 || nNewValue > 100 || nNewValue > m_nMaxDropDownItems)
		ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDITEMSLIMITATION);

	// Save the new value.
	m_nMinDropDownItems = nNewValue;
	SetModifiedFlag();
	BoundPropertyChanged(dispidMinDropDownItems);
}

short CBHMDBDropDownCtrl::GetListWidth() 
/*
Routine Description:
	Get the width of dropdown window.

Parameters:
	None.

Return Value:
	The width of dropdown window.
*/
{
	return m_nListWidth;
}

void CBHMDBDropDownCtrl::SetListWidth(short nNewValue) 
/*
Routine Description:
	Set the width of dropdown window.

Parameters:
	nNewValue		The new value.

Return Value:
	None.
*/
{
	// Negative value has no meaning.
	if (nNewValue < 0)
		nNewValue = 0;

	m_nListWidth = nNewValue;
	SetModifiedFlag();
	BoundPropertyChanged(dispidListWidth);
}

OLE_COLOR CBHMDBDropDownCtrl::GetBackColor() 
/*
Routine Description:
	Get the background color.

Parameters:
	None.

Return Value:
	The background color.
*/
{
	return m_clrBack;
}

void CBHMDBDropDownCtrl::SetBackColor(OLE_COLOR nNewValue) 
/*
Routine Description:
	Set the background color.

Parameters:
	nNewValue		The new value.

Return Value:
	None.
*/
{
	// Save the new value.
	m_clrBack = nNewValue;

	// Update the grid.
	if (m_pGridDraw)
		m_pGridDraw->SetBackColor(TranslateColor(GetBackColor()));
	else if (m_pDropDownRealWnd)
		m_pDropDownRealWnd->SetBackColor(TranslateColor(GetBackColor()));
	
	SetModifiedFlag();
	BoundPropertyChanged(dispidBackColor);
	InvalidateControl();
}

void CBHMDBDropDownCtrl::Bind()
/*
Routine Description:
	Update the binding entries in data source.

Parameters:
	None.

Return Value:
	None.
*/
{
	m_Rowset.ReleaseRows();

	// Release current accessors.
	if (FAILED(m_Rowset.ReleaseAccessors(m_Rowset.m_spRowset)))
		return;

	// Remove cached binding information.
	m_nColumnsBind = 0;
	
	for (int i = 0; i < m_arBindInfo.GetSize(); i ++)
		delete m_arBindInfo[i];

	m_arBindInfo.RemoveAll();

	// Construct new binding information.
	m_arBindInfo.SetSize(m_pDropDownRealWnd->GetColCount());

	for (i = 0; m_nDataMode == 0 && i < m_pDropDownRealWnd->GetColCount(); i ++)
	{
		// Fill each binding entry.

		// Create a new entry object.
		ColumnProp * pCol = new ColumnProp;

		// Set the col index of this entry.
		pCol->nColIndex = m_pDropDownRealWnd->GetColIndex(i + 1);

		// Set the data field value.
		CString strField = m_pDropDownRealWnd->GetColUserAttribRef(i + 1).ElementAt(0);
		pCol->nDataField = GetFieldFromStr(strField);

		if (pCol->nDataField != -1)
			m_nColumnsBind ++;

		m_arBindInfo[i] = pCol;
	}

	// Include the bookmark field.
	int nEntries = m_nColumnsBind + (m_bHaveBookmarks ? 1 : 0);

	// Bind the new entries with data source.

	m_Rowset.CreateAccessor(nEntries, &m_data, sizeof(m_data));
	if (m_bHaveBookmarks)
		m_Rowset.AddBindEntry(0, DBTYPE_BYTES, sizeof(m_data.bookmark), &m_data.bookmark);

	for (i = 0; i < m_arBindInfo.GetSize(); i ++)
	{
		if (m_arBindInfo[i]->nDataField != -1)
			m_Rowset.AddBindEntry(m_arBindInfo[i]->nDataField, DBTYPE_STR, 255 * sizeof(TCHAR), 
				&(m_data.strData[i]), NULL, &(m_data.dwStatus[i]));
	}

	HRESULT hr = m_Rowset.Bind();
	if (FAILED(hr))
	{
		m_pDropDownRealWnd->LockUpdate(FALSE);
		m_pDropDownRealWnd->Invalidate();
	}
}

VARIANT CBHMDBDropDownCtrl::GetBookmark() 
/*
Routine Description:
	Get the bookmark of current row.

Parameters:
	None.

Return Value:
	The bookmark object.
*/
{
	// This property is unavailable at design time.
	if (!AmbientUserMode())
		GetNotSupported();

	// Get the ordinal of current row.
	int nRow, nCol;
	m_pDropDownRealWnd->GetCurrentCell(nRow, nCol);

	// Return the bookmark.
	VARIANT sa;
	VariantInit(&sa);
	GetBmkOfRow(nRow, &sa);

	return sa;
}

void CBHMDBDropDownCtrl::SetBookmark(const VARIANT FAR& newValue) 
/*
Routine Description:
	Move current row to given bookmark.

Parameters:
	newValue		The new bookmark.

Return Value:
	None.
*/
{
	// This property is unavailable at design time.
	if (!AmbientUserMode())
		SetNotSupported();

	SetRow(GetRowFromBmk(&newValue));
}

VARIANT CBHMDBDropDownCtrl::RowBookmark(long RowIndex) 
/*
Routine Description:
	Get the bookmark of the given row.

Parameters:
	RowIndex		The row to fecth.

Return Value:
	The bookmark of this row.
*/
{
	VARIANT vaResult;
	VariantInit(&vaResult);

	// Get the bookmark using helper function.
	GetBmkOfRow(RowIndex, &vaResult);

	return vaResult;
}

void CBHMDBDropDownCtrl::GetBmkOfRow(int nRow, VARIANT *pVaRet)
/*
Routine Description:
	Get bookmark of one row.

Parameters:
	nRow		The row ordinal.

	pvaRet		The buffer to receive result bookmark.

Return Value:
	None.
*/
{
	if (m_nDataMode)
	{
		// In manual mode, row bookmark is just the row index.

		pVaRet->vt = VT_I4;
		pVaRet->lVal = m_pDropDownRealWnd->GetRowIndex(nRow);

		return;
	}
	
	if (nRow > 0 && nRow <= (int)m_apBookmark.GetSize())
	{
		// Get the bookmark from data source.

		GetBmkOfRow(nRow);
		
		// Init the receive buffer.
		pVaRet->vt = VT_ARRAY | VT_UI1;
		SAFEARRAYBOUND rgsabound;
		rgsabound.cElements = m_nBookmarkSize;
		rgsabound.lLbound = 0;
		
		SAFEARRAY * parray = SafeArrayCreate(VT_UI1, 1, &rgsabound);
		if (parray == NULL)
			AfxThrowMemoryException();
		
		// Copy the bookmark into receive buffer.		
		void* pvDestData;
		AfxCheckError(SafeArrayAccessData(parray, &pvDestData));
		memcpy(pvDestData, m_apBookmark[nRow - 1], m_nBookmarkSize);
		AfxCheckError(SafeArrayUnaccessData(parray));
		
		pVaRet->parray = parray;
	}
}

int CBHMDBDropDownCtrl::GetRowFromIndex(int nIndex)
/*
Routine Description:
	Get the row ordinal from row index.

Parameters:
	nIndex		The row index to search.

Return Value:
	If found, return row ordinal; Otherwise, return INVALID.
*/
{
	// Compare the index of each row with the given index.
	for (int nRow = 1; nRow <= m_pDropDownRealWnd->GetRowCount(); nRow ++)
		if (m_pDropDownRealWnd->GetRowIndex(nRow) == nIndex)
			return nRow;

	// Did not find.
	return INVALID;
}

int CBHMDBDropDownCtrl::GetRowFromBmk(const VARIANT *pBmk)
/*
Routine Description:
	Get row ordinal from bookmark.

Parameters:
	pBmk		The bookmark object to search.

Return Value:
	The row ordinal.
*/
{
	if (m_nDataMode)
	{
		// In manual mode, row bookmark is just the row index.

		COleVariant va = *pBmk;
		va.ChangeType(VT_I4);

		return GetRowFromIndex(va.lVal);
	}

	// Validate the given bookmark object.
	if (pBmk->vt != (VT_UI1 | VT_ARRAY) || !pBmk->parray->pvData)
		ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDBOOKMARK);

	LONG nLBound = 0, nUBound = 0;
	AfxCheckError(SafeArrayGetLBound(pBmk->parray, 1, &nLBound));
	AfxCheckError(SafeArrayGetUBound(pBmk->parray, 1, &nUBound));

	if (nUBound - nLBound + 1 != m_nBookmarkSize)
		ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDBOOKMARK);

	BYTE * bookmark = NULL;

	AfxCheckError(SafeArrayAccessData(pBmk->parray, (void **)&bookmark));
	int nRow = GetRowFromBmk(&bookmark[nLBound]);
	AfxCheckError(SafeArrayUnaccessData(pBmk->parray));

	return nRow;
}

LPDISPATCH CBHMDBDropDownCtrl::GetGroups() 
/*
Routine Description:
	Get the groups object.

Parameters:
	None.

Return Value:
	The groups object.
*/
{
	// Create a new groups object.
	CGroups * pGroups = new CGroups;

	// Set the callback pointer.
	pGroups->SetCtrlPtr(this);
	pGroups->m_pGrid = m_pDropDownRealWnd;

	return pGroups->GetIDispatch(FALSE);
}

void CBHMDBDropDownCtrl::OnDividerColorChanged() 
/*
Routine Description:
	The dividercolor has been changed.

Parameters:
	None.

Return Value:
	None.
*/
{
	// Update the divider lines.
{
	UpdateDivider();
}

BOOL CBHMDBDropDownCtrl::GetGroupHeader() 
/*
Routine Description:
	Get the visibility of group header.

Parameters:
	None.

Return Value:
	If the group header should be shown, return TRUE; Otherwise, return FALSE.
*/
{
	if (m_pDropDownRealWnd)
		m_bGroupHeader = m_pDropDownRealWnd->GetShowGroupHeader();
	else if (m_pGridDraw)
		m_bGroupHeader = m_pGridDraw->GetShowGroupHeader();

	return m_bGroupHeader;
}

void CBHMDBDropDownCtrl::SetGroupHeader(BOOL bNewValue) 
/*
Routine Description:
	Set the visibility of group header.

Parameters:
	bNewValue		The new value.

Return Value:
	None.
*/
{
	// Save the new value.
	m_bGroupHeader = bNewValue;

	// Update the grid.
	if (m_pDropDownRealWnd)
		m_pDropDownRealWnd->SetShowGroupHeader(m_bGroupHeader);
	else if (m_pGridDraw)
		m_pGridDraw->SetShowGroupHeader(m_bGroupHeader);

	SetModifiedFlag();
	BoundPropertyChanged(dispidGroupHeader);
	InvalidateControl();
}
