// BHMDBComboCtl.cpp : Implementation of the CBHMDBComboCtrl ActiveX Control class.

#include "stdafx.h"
#include "BHMData.h"
#include "BHMDBComboCtl.h"
#include "BHMDBComboPpgGeneral.h"
#include <oledberr.h>
#include <msstkppg.h>
#include "Columns.h"
#include "DropDownRealWnd.h"
#include "BHMDBColumnPropPage.h"
#include "BHMDBGroupsPropPage.h"
#include "ComboEdit.h"
#include "Groups.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


IMPLEMENT_DYNCREATE(CBHMDBComboCtrl, COleControl)

BEGIN_INTERFACE_MAP(CBHMDBComboCtrl, COleControl)
	// Map the interfaces to their implementations
	INTERFACE_PART(CBHMDBComboCtrl, IID_IRowsetNotify, RowsetNotify)
END_INTERFACE_MAP()

STDMETHODIMP_(ULONG) CBHMDBComboCtrl::XRowsetNotify::AddRef(void)
{
	METHOD_PROLOGUE(CBHMDBComboCtrl, RowsetNotify)
	return pThis->ExternalAddRef();
}

STDMETHODIMP_(ULONG) CBHMDBComboCtrl::XRowsetNotify::Release(void)
{
	METHOD_PROLOGUE(CBHMDBComboCtrl, RowsetNotify)
	return pThis->ExternalRelease();
}

STDMETHODIMP CBHMDBComboCtrl::XRowsetNotify::QueryInterface(REFIID iid,
	LPVOID FAR* ppvObject)
{
	METHOD_PROLOGUE(CBHMDBComboCtrl, RowsetNotify)
	return pThis->ExternalQueryInterface(&iid, ppvObject);
}

STDMETHODIMP CBHMDBComboCtrl::XRowsetNotify::OnFieldChange(IRowset * prowset,
	HROW hRow, ULONG cColumns, ULONG rgColumns[], DBREASON eReason, DBEVENTPHASE ePhase, BOOL fCantDeny)
{
	METHOD_PROLOGUE(CBHMDBComboCtrl, RowsetNotify)

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
			if (nRow == INVALID || nRow > pThis->m_pDropDownRealWnd->OnGetRecordCount())
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

STDMETHODIMP CBHMDBComboCtrl::XRowsetNotify::OnRowChange(IRowset * prowset,
	ULONG cRows, const HROW rghRows[], DBREASON eReason, DBEVENTPHASE ePhase, BOOL fCantDeny)
{
	METHOD_PROLOGUE(CBHMDBComboCtrl, RowsetNotify)

	
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

STDMETHODIMP CBHMDBComboCtrl::XRowsetNotify::OnRowsetChange(IRowset * prowset,
	DBREASON eReason, DBEVENTPHASE ePhase, BOOL fCantDeny)
{
	METHOD_PROLOGUE(CBHMDBComboCtrl, RowsetNotify)

	if (DBEVENTPHASE_DIDEVENT == ePhase && prowset == pThis->m_Rowset.m_spRowset)
	{
		switch (eReason)
		{
		// Rowset has been changed.
		case DBREASON_ROWSET_RELEASE:
		case DBREASON_ROWSET_CHANGED:
			pThis->m_bDataSrcChanged = TRUE;

		return S_OK;
		break;
			
		default:
			return DB_S_UNWANTEDREASON;
		}
	}

	return DB_S_UNWANTEDPHASE;
}

STDMETHODIMP_(ULONG) CBHMDBComboCtrl::XDataSourceListener::AddRef(void)
{
	METHOD_PROLOGUE(CBHMDBComboCtrl, DataSourceListener)
	return pThis->ExternalAddRef();
}

STDMETHODIMP_(ULONG) CBHMDBComboCtrl::XDataSourceListener::Release(void)
{
	METHOD_PROLOGUE(CBHMDBComboCtrl, DataSourceListener)
	return pThis->ExternalRelease();
}

STDMETHODIMP CBHMDBComboCtrl::XDataSourceListener::QueryInterface(REFIID iid,
	LPVOID FAR* ppvObject)
{
	METHOD_PROLOGUE(CBHMDBComboCtrl, DataSourceListener)
	return pThis->ExternalQueryInterface(&iid, ppvObject);
}

STDMETHODIMP CBHMDBComboCtrl::XDataSourceListener::dataMemberChanged(BSTR bstrDM)
{
	METHOD_PROLOGUE(CBHMDBComboCtrl, RowsetNotify)

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

STDMETHODIMP CBHMDBComboCtrl::XDataSourceListener::dataMemberAdded(BSTR bstrDM)
{
	return S_OK;
}

STDMETHODIMP CBHMDBComboCtrl::XDataSourceListener::dataMemberRemoved(BSTR bstrDM)
{
	METHOD_PROLOGUE(CBHMDBComboCtrl, RowsetNotify)

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

BEGIN_MESSAGE_MAP(CBHMDBComboCtrl, COleControl)
	//{{AFX_MSG_MAP(CBHMDBComboCtrl)
	ON_WM_CREATE()
	ON_WM_KILLFOCUS()
	ON_WM_SETFOCUS()
	ON_WM_LBUTTONDOWN()
	ON_WM_LBUTTONUP()
	ON_WM_CTLCOLOR()
	ON_WM_MOUSEACTIVATE()
	//}}AFX_MSG_MAP
	ON_OLEVERB(AFX_IDS_VERB_PROPERTIES, OnProperties)
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// Dispatch map

BEGIN_DISPATCH_MAP(CBHMDBComboCtrl, COleControl)
	//{{AFX_DISPATCH_MAP(CBHMDBComboCtrl)
	DISP_PROPERTY_NOTIFY(CBHMDBComboCtrl, "HeadBackColor", m_clrHeadBack, OnHeadBackColorChanged, VT_COLOR)
	DISP_PROPERTY_NOTIFY(CBHMDBComboCtrl, "HeadForeColor", m_clrHeadFore, OnHeadForeColorChanged, VT_COLOR)
	DISP_PROPERTY_NOTIFY(CBHMDBComboCtrl, "ListAutoPosition", m_bListAutoPosition, OnListAutoPositionChanged, VT_BOOL)
	DISP_PROPERTY_NOTIFY(CBHMDBComboCtrl, "ListWidthAutoSize", m_bListWidthAutoSize, OnListWidthAutoSizeChanged, VT_BOOL)
	DISP_PROPERTY_NOTIFY(CBHMDBComboCtrl, "DividerColor", m_clrDivider, OnDividerColorChanged, VT_COLOR)
	DISP_PROPERTY_EX(CBHMDBComboCtrl, "BackColor", GetBackColor, SetBackColor, VT_COLOR)
	DISP_PROPERTY_EX(CBHMDBComboCtrl, "Format", GetFormat, SetFormat, VT_BSTR)
	DISP_PROPERTY_EX(CBHMDBComboCtrl, "MaxLength", GetMaxLength, SetMaxLength, VT_I2)
	DISP_PROPERTY_EX(CBHMDBComboCtrl, "ReadOnly", GetReadOnly, SetReadOnly, VT_BOOL)
	DISP_PROPERTY_EX(CBHMDBComboCtrl, "DataSource", GetDataSource, SetDataSource, VT_UNKNOWN)
	DISP_PROPERTY_EX(CBHMDBComboCtrl, "DataMember", GetDataMember, SetDataMember, VT_BSTR)
	DISP_PROPERTY_EX(CBHMDBComboCtrl, "DefColWidth", GetDefColWidth, SetDefColWidth, VT_I4)
	DISP_PROPERTY_EX(CBHMDBComboCtrl, "RowHeight", GetRowHeight, SetRowHeight, VT_I4)
	DISP_PROPERTY_EX(CBHMDBComboCtrl, "DividerType", GetDividerType, SetDividerType, VT_I4)
	DISP_PROPERTY_EX(CBHMDBComboCtrl, "DividerStyle", GetDividerStyle, SetDividerStyle, VT_I4)
	DISP_PROPERTY_EX(CBHMDBComboCtrl, "DataMode", GetDataMode, SetDataMode, VT_I4)
	DISP_PROPERTY_EX(CBHMDBComboCtrl, "LeftCol", GetLeftCol, SetLeftCol, VT_I4)
	DISP_PROPERTY_EX(CBHMDBComboCtrl, "FirstRow", GetFirstRow, SetFirstRow, VT_I4)
	DISP_PROPERTY_EX(CBHMDBComboCtrl, "Row", GetRow, SetRow, VT_I4)
	DISP_PROPERTY_EX(CBHMDBComboCtrl, "Col", GetCol, SetCol, VT_I4)
	DISP_PROPERTY_EX(CBHMDBComboCtrl, "Rows", GetRows, SetRows, VT_I4)
	DISP_PROPERTY_EX(CBHMDBComboCtrl, "Cols", GetCols, SetCols, VT_I4)
	DISP_PROPERTY_EX(CBHMDBComboCtrl, "DroppedDown", GetDroppedDown, SetDroppedDown, VT_BOOL)
	DISP_PROPERTY_EX(CBHMDBComboCtrl, "DataField", GetDataField, SetDataField, VT_BSTR)
	DISP_PROPERTY_EX(CBHMDBComboCtrl, "Columns", GetColumns, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX(CBHMDBComboCtrl, "VisibleRows", GetVisibleRows, SetNotSupported, VT_I4)
	DISP_PROPERTY_EX(CBHMDBComboCtrl, "VisibleCols", GetVisibleCols, SetNotSupported, VT_I4)
	DISP_PROPERTY_EX(CBHMDBComboCtrl, "ColumnHeader", GetColumnHeader, SetColumnHeader, VT_BOOL)
	DISP_PROPERTY_EX(CBHMDBComboCtrl, "Font", GetFont, SetFont, VT_FONT)
	DISP_PROPERTY_EX(CBHMDBComboCtrl, "HeadFont", GetHeadFont, SetHeadFont, VT_FONT)
	DISP_PROPERTY_EX(CBHMDBComboCtrl, "FieldSeparator", GetFieldSeparator, SetFieldSeparator, VT_BSTR)
	DISP_PROPERTY_EX(CBHMDBComboCtrl, "MaxDropDownItems", GetMaxDropDownItems, SetMaxDropDownItems, VT_I2)
	DISP_PROPERTY_EX(CBHMDBComboCtrl, "MinDropDownItems", GetMinDropDownItems, SetMinDropDownItems, VT_I2)
	DISP_PROPERTY_EX(CBHMDBComboCtrl, "ListWidth", GetListWidth, SetListWidth, VT_I2)
	DISP_PROPERTY_EX(CBHMDBComboCtrl, "HeaderHeight", GetHeaderHeight, SetHeaderHeight, VT_I4)
	DISP_PROPERTY_EX(CBHMDBComboCtrl, "Text", GetText, SetText, VT_BSTR)
	DISP_PROPERTY_EX(CBHMDBComboCtrl, "Bookmark", GetBookmark, SetBookmark, VT_VARIANT)
	DISP_PROPERTY_EX(CBHMDBComboCtrl, "Style", GetStyle, SetStyle, VT_I4)
	DISP_PROPERTY_EX(CBHMDBComboCtrl, "Groups", GetGroups, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX(CBHMDBComboCtrl, "GroupHeader", GetGroupHeader, SetGroupHeader, VT_BOOL)
	DISP_FUNCTION(CBHMDBComboCtrl, "AddItem", AddItem, VT_EMPTY, VTS_BSTR VTS_VARIANT)
	DISP_FUNCTION(CBHMDBComboCtrl, "RemoveItem", RemoveItem, VT_EMPTY, VTS_I4)
	DISP_FUNCTION(CBHMDBComboCtrl, "Scroll", Scroll, VT_EMPTY, VTS_I4 VTS_I4)
	DISP_FUNCTION(CBHMDBComboCtrl, "RemoveAll", RemoveAll, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CBHMDBComboCtrl, "MoveFirst", MoveFirst, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CBHMDBComboCtrl, "MovePrevious", MovePrevious, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CBHMDBComboCtrl, "MoveNext", MoveNext, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CBHMDBComboCtrl, "MoveLast", MoveLast, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CBHMDBComboCtrl, "MoveRecords", MoveRecords, VT_EMPTY, VTS_I4)
	DISP_FUNCTION(CBHMDBComboCtrl, "ReBind", ReBind, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CBHMDBComboCtrl, "IsItemInList", IsItemInList, VT_BOOL, VTS_NONE)
	DISP_FUNCTION(CBHMDBComboCtrl, "RowBookmark", RowBookmark, VT_VARIANT, VTS_I4)
	DISP_DEFVALUE(CBHMDBComboCtrl, "Text")
	DISP_STOCKPROP_FORECOLOR()
	//}}AFX_DISPATCH_MAP
	DISP_FUNCTION_ID(CBHMDBComboCtrl, "AboutBox", DISPID_ABOUTBOX, AboutBox, VT_EMPTY, VTS_NONE)
	DISP_PROPERTY_STOCK(CBHMDBComboCtrl, "CtrlType", 255, CBHMDBComboCtrl::GetCtrlType, COleControl::SetNotSupported, VT_I2)
END_DISPATCH_MAP()


/////////////////////////////////////////////////////////////////////////////
// Event map

BEGIN_EVENT_MAP(CBHMDBComboCtrl, COleControl)
	//{{AFX_EVENT_MAP(CBHMDBComboCtrl)
	EVENT_CUSTOM("Scroll", FireScroll, VTS_PI2)
	EVENT_CUSTOM("ScrollAfter", FireScrollAfter, VTS_NONE)
	EVENT_CUSTOM("CloseUp", FireCloseUp, VTS_NONE)
	EVENT_CUSTOM("DropDown", FireDropDown, VTS_NONE)
	EVENT_CUSTOM("PositionList", FirePositionList, VTS_BSTR)
	EVENT_CUSTOM("Change", FireChange, VTS_NONE)
	EVENT_STOCK_CLICK()
	EVENT_STOCK_KEYDOWN()
	EVENT_STOCK_KEYPRESS()
	EVENT_STOCK_KEYUP()
	//}}AFX_EVENT_MAP
END_EVENT_MAP()


/////////////////////////////////////////////////////////////////////////////
// Property pages

// TODO: Add more property pages as needed.  Remember to increase the count!
BEGIN_PROPPAGEIDS(CBHMDBComboCtrl, 5)
	PROPPAGEID(CBHMDBComboPropPageGeneral::guid)
	PROPPAGEID(CBHMDBGroupsPropPage::guid)
	PROPPAGEID(CBHMDBColumnPropPage::guid)
	PROPPAGEID(CLSID_StockColorPage)
	PROPPAGEID(CLSID_StockFontPage)
END_PROPPAGEIDS(CBHMDBComboCtrl)


/////////////////////////////////////////////////////////////////////////////
// Initialize class factory and guid

IMPLEMENT_OLECREATE_EX(CBHMDBComboCtrl, "BHMDATA.BHMDBComboCtrl.1",
	0xbf91b12d, 0x6a84, 0x11d3, 0xa7, 0xfe, 0, 0x80, 0xc8, 0x76, 0x3f, 0xa4)


/////////////////////////////////////////////////////////////////////////////
// Type library ID and version

IMPLEMENT_OLETYPELIB(CBHMDBComboCtrl, _tlid, _wVerMajor, _wVerMinor)


/////////////////////////////////////////////////////////////////////////////
// Interface IDs

const IID BASED_CODE IID_DBHMDBCombo =
		{ 0xbf91b12b, 0x6a84, 0x11d3, { 0xa7, 0xfe, 0, 0x80, 0xc8, 0x76, 0x3f, 0xa4 } };
const IID BASED_CODE IID_DBHMDBComboEvents =
		{ 0xbf91b12c, 0x6a84, 0x11d3, { 0xa7, 0xfe, 0, 0x80, 0xc8, 0x76, 0x3f, 0xa4 } };


/////////////////////////////////////////////////////////////////////////////
// Control type information

static const DWORD BASED_CODE _dwBHMDBComboOleMisc =
	OLEMISC_ACTIVATEWHENVISIBLE |
	OLEMISC_SETCLIENTSITEFIRST |
	OLEMISC_INSIDEOUT |
	OLEMISC_CANTLINKINSIDE |
	OLEMISC_RECOMPOSEONRESIZE;

IMPLEMENT_OLECTLTYPE(CBHMDBComboCtrl, IDS_BHMDBCOMBO, _dwBHMDBComboOleMisc)


/////////////////////////////////////////////////////////////////////////////
// CBHMDBComboCtrl::CBHMDBComboCtrlFactory::UpdateRegistry -
// Adds or removes system registry entries for CBHMDBComboCtrl

BOOL CBHMDBComboCtrl::CBHMDBComboCtrlFactory::UpdateRegistry(BOOL bRegister)
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
			IDS_BHMDBCOMBO,
			IDB_BHMDBCOMBO,
			afxRegApartmentThreading,
			_dwBHMDBComboOleMisc,
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
// CBHMDBComboCtrl::CBHMDBComboCtrlFactory::VerifyUserLicense -
// Checks for existence of a user license

BOOL CBHMDBComboCtrl::CBHMDBComboCtrlFactory::VerifyUserLicense()
{
	return AfxVerifyLicFile(AfxGetInstanceHandle(), _szLicFileName,
		_szLicString);
}


/////////////////////////////////////////////////////////////////////////////
// CBHMDBComboCtrl::CBHMDBComboCtrlFactory::GetLicenseKey -
// Returns a runtime licensing key

BOOL CBHMDBComboCtrl::CBHMDBComboCtrlFactory::GetLicenseKey(DWORD dwReserved,
	BSTR FAR* pbstrKey)
{
	if (pbstrKey == NULL)
		return FALSE;

	*pbstrKey = SysAllocString(_szLicString);
	return (*pbstrKey != NULL);
}


/////////////////////////////////////////////////////////////////////////////
// CBHMDBComboCtrl::CBHMDBComboCtrl - Constructor

CBHMDBComboCtrl::CBHMDBComboCtrl() : m_fntCell(NULL), m_fntHeader(NULL)
{
	InitializeIIDs(&IID_DBHMDBCombo, &IID_DBHMDBComboEvents);

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

	m_pDropDownRealWnd = NULL;
	m_bDataSrcChanged = FALSE;
	m_spDataSource = NULL;
	m_nCol = m_nRow = 0;
	m_nColsUsed = 0;

	m_strMask.Empty();
	m_strFormat.Empty();
	m_nMaxLength = 255;
	m_nFontHeight = BUTTONHEIGHT;
	m_hFont = NULL;

	for (int i = 0; i < 255; i ++)
		m_bIsDelimiter[i] = FALSE;

	m_pEdit = NULL;
	m_hSystemHook = NULL;
	m_bLButtonDown = FALSE;

	m_pBrhBack = NULL;
}


/////////////////////////////////////////////////////////////////////////////
// CBHMDBComboCtrl::~CBHMDBComboCtrl - Destructor

CBHMDBComboCtrl::~CBHMDBComboCtrl()
{
	// Remove the edit window if exists.
	if (m_pEdit)
	{
		m_pEdit->DestroyWindow();
		delete m_pEdit;
		m_pEdit = NULL;
	}

	// Remove the dropdown window if exists.
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

	// Released the held IFont pointer.
	IFont * pIfont = m_fntCell.m_pFont;
	if (m_hFont)
	{
		pIfont->ReleaseHfont(m_hFont);
		m_hFont = NULL;
	}
 
	// Release the mouse hook.
	if (m_hSystemHook)
	{
		UnhookWindowsHookEx(m_hSystemHook);
		m_hSystemHook = NULL;
	}

	// Clear the brush object.
	if (m_pBrhBack)
	{
		delete m_pBrhBack;
		m_pBrhBack = NULL;
	}
}


/////////////////////////////////////////////////////////////////////////////
// CBHMDBComboCtrl::OnDraw - Drawing function

void CBHMDBComboCtrl::OnDraw(
			CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid)
{
	// Init the font object if hasn't done.
	if (m_hFont == NULL)
		InitCellFont();

	// Draw the background first.
	pdc->FillSolidRect(&rcBounds, TranslateColor(GetBackColor()));

	// Reserve the space for border.
	CRect rt = rcBounds;
	rt.DeflateRect(2, 2, 2, 2);

	// Calc the rect of the button.
	m_rectButton = rt;
	m_rectButton.left = m_rectButton.right - BUTTONWIDTH;
	m_rectButton.bottom = m_rectButton.top + m_nFontHeight - 4;

	// Draw the border of the button.
	pdc->Draw3dRect(rcBounds.left, rcBounds.top, rcBounds.Width(),
		rcBounds.Height(), GetSysColor(COLOR_3DSHADOW), RGB(255, 255, 255));
	pdc->Draw3dRect(rcBounds.left + 1, rcBounds.top + 1, rcBounds.Width() - 2,
		rcBounds.Height() - 2, RGB(0, 0, 0), GetSysColor(COLOR_3DFACE));

	if (!AmbientUserMode())
	{
		// Draw the text at design time.

		// The rect to draw text.
		CRect rectText = rt;

		// Reserve some space at bottom line.
		rectText.DeflateRect(0, 2, 0, 2);

		// Draw the text.
		CFont * pOldFont = m_fntCell.Select(pdc, rcBounds.Height(), m_cyExtent);
		pdc->SetBkColor(TranslateColor(GetBackColor()));
		pdc->SetTextColor(TranslateColor(GetForeColor()));
		pdc->ExtTextOut(rectText.left, rectText.top, ETO_CLIPPED | ETO_OPAQUE, &rectText, 
			AmbientDisplayName(), NULL);
		pdc->SelectObject(pOldFont);
	}
	
	// The brush to draw button.
	CBrush brh;

	if (!m_bLButtonDown)
	{
		// The button is not pressed down.
		pdc->Draw3dRect(&m_rectButton, GetSysColor(COLOR_3DLIGHT), GetSysColor(COLOR_3DDKSHADOW));
		m_rectButton.DeflateRect(1, 1);
		pdc->Draw3dRect(&m_rectButton, GetSysColor(COLOR_3DHILIGHT), GetSysColor(COLOR_3DSHADOW));
		m_rectButton.DeflateRect(1, 1);
		brh.CreateSolidBrush(GetSysColor(COLOR_3DFACE));
		pdc->FillSolidRect(&m_rectButton, GetSysColor(COLOR_3DFACE));
		m_rectButton.InflateRect(2, 2);
	}
	else
	{
		// The button is pressed down.
		brh.CreateSolidBrush(GetSysColor(COLOR_3DFACE));
		pdc->FillRect(&m_rectButton, &brh);
	}

	// Calc the point array of the triangle image.
	POINT pt[3];
	pt[0].x = m_rectButton.left + m_rectButton.Width() / 2 - 4;
	pt[0].y = m_rectButton.top + m_rectButton.Height() / 2 - 2;
	pt[1].x = pt[0].x + 8;
	pt[1].y = pt[0].y;
	pt[2].x = pt[0].x + 4;
	pt[2].y = pt[0].y + 4;
	
	// Draw the triangle in the center of the button.
	CBrush brhNew;
	brhNew.CreateSolidBrush(RGB(0, 0, 0));
	CBrush * pBrhOld = pdc->SelectObject(&brhNew);
	pdc->Polygon(pt, 3);
	pdc->SelectObject(pBrhOld);
}


/////////////////////////////////////////////////////////////////////////////
// CBHMDBComboCtrl::DoPropExchange - Persistence support

void CBHMDBComboCtrl::DoPropExchange(CPropExchange* pPX)
{
	ExchangeVersion(pPX, MAKELONG(_wVerMinor, _wVerMajor));
	COleControl::DoPropExchange(pPX);

	PX_String(pPX, _T("Format"), m_strFormat, "");
	PX_String(pPX, _T("Text"), m_strText, "");
	PX_Bool(pPX, _T("ReadOnly"), m_bReadOnly, FALSE);
	PX_Short(pPX, _T("MaxLength"), m_nMaxLength, 255);
	PX_Long(pPX, _T("RowHeight"), m_nRowHeight, 20);
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
	PX_Long(pPX, "Style", m_nStyle, 0);
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
		CalcMask();
      
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
// CBHMDBComboCtrl::OnResetState - Reset control to default state

void CBHMDBComboCtrl::OnResetState()
{
	COleControl::OnResetState();  // Resets defaults found in DoPropExchange

	// TODO: Reset any other control state here.
}


/////////////////////////////////////////////////////////////////////////////
// CBHMDBComboCtrl::AboutBox - Display an "About" box to the user

void CBHMDBComboCtrl::AboutBox()
{
	CDialog dlgAbout(IDD_ABOUTBOX_BHMDBCOMBO);
	dlgAbout.DoModal();
}

// The backdoor for property page :)
short CBHMDBComboCtrl::GetCtrlType() 
{
	return 2;
}

/////////////////////////////////////////////////////////////////////////////
// CBHMDBComboCtrl message handlers
short CBHMDBComboCtrl::GetMaxLength() 
/*
Routine Description:
	Get the max chars the text can hold.

Parameters:
	None.

Return Value:
	The max length.
*/
{
	return m_nMaxLength;
}

void CBHMDBComboCtrl::SetMaxLength(short nNewValue) 
/*
Routine Description:
	Set the max chars the text can hold.

Parameters:
	nNewValue		The new value.

Return Value:
	None.
*/
{
	// The value should between 0 and 255.
	if (nNewValue <0 || nNewValue > 255)
		return;

	// Save the new value.
	m_nMaxLength = nNewValue;

	SetModifiedFlag();

	BoundPropertyChanged(dispidMaxLength);
}

void CBHMDBComboCtrl::SetRowHeight(long nNewValue)
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

	// Update the dropdown window if exists.
	if (m_pDropDownRealWnd)
		m_pDropDownRealWnd->SetRowHeight(m_nRowHeight);

	SetModifiedFlag();
	BoundPropertyChanged(dispidRowHeight);
	InvalidateControl();
}

BOOL CBHMDBComboCtrl::GetReadOnly() 
/*
Routine Description:
	Get the permission for modifying the text shown in control.

Parameters:
	None.

Return Value:
	If is readonly, return TRUE; Otherwise, return FALSE.
*/
{
	return m_bReadOnly;
}

void CBHMDBComboCtrl::SetReadOnly(BOOL bNewValue) 
/*
Routine Description:
	Set the permission for modifying the text shown in control.

Parameters:
	bNewValue		The new value.

Return Value:
	None.
*/
{
	// Save the new value.
	m_bReadOnly = bNewValue;
	
	SetModifiedFlag();
	BoundPropertyChanged(dispidReadOnly);
}

LPUNKNOWN CBHMDBComboCtrl::GetDataSource() 
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

void CBHMDBComboCtrl::SetDataSource(LPUNKNOWN newValue) 
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

BSTR CBHMDBComboCtrl::GetDataMember() 
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

void CBHMDBComboCtrl::SetDataMember(LPCTSTR lpszNewValue) 
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

long CBHMDBComboCtrl::GetDefColWidth() 
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

void CBHMDBComboCtrl::SetDefColWidth(long nNewValue) 
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
	
	// Update the dropdown window if exists.
	if (m_pDropDownRealWnd)
		m_pDropDownRealWnd->SetDefColWidth(m_nDefColWidth);

	SetModifiedFlag();
	BoundPropertyChanged(dispidDefColWidth);
}

long CBHMDBComboCtrl::GetRowHeight() 
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

	return m_nRowHeight;
}

long CBHMDBComboCtrl::GetDividerType() 
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

void CBHMDBComboCtrl::SetDividerType(long nNewValue) 
/*
Routine Description:
	Set the type of divider lines.

Parameters:
	nNewValue		The new value.

Return Value:
	None.
*/
{
	// Save the new value.
	m_nDividerType = nNewValue;

	// Update the grid being used.
	UpdateDivider();

	SetModifiedFlag();
	BoundPropertyChanged(dispidDividerType);
}

long CBHMDBComboCtrl::GetDividerStyle() 
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

void CBHMDBComboCtrl::SetDividerStyle(long nNewValue) 
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
}

long CBHMDBComboCtrl::GetDataMode() 
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

void CBHMDBComboCtrl::SetDataMode(long nNewValue) 
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

long CBHMDBComboCtrl::GetLeftCol() 
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

	// If dropdown window exists, return the property from this window; Otherwise return 0.
	if (m_pDropDownRealWnd)
		return m_pDropDownRealWnd->GetLeftCol();
	else
		return 0;
}

void CBHMDBComboCtrl::SetLeftCol(long nNewValue) 
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

long CBHMDBComboCtrl::GetFirstRow() 
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

void CBHMDBComboCtrl::SetFirstRow(long nNewValue) 
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

long CBHMDBComboCtrl::GetRow() 
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

void CBHMDBComboCtrl::SetRow(long nNewValue) 
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
	if (m_pDropDownRealWnd->GetRowCount() >= m_nRow)
	// Update the grid.
		m_pDropDownRealWnd->MoveTo(m_nRow);
	else
		ThrowError(CTL_E_INVALIDPROPERTYARRAYINDEX, AFX_IDP_E_INVALIDPROPERTYARRAYINDEX);

	SetModifiedFlag();
	BoundPropertyChanged(dispidRow);
}

long CBHMDBComboCtrl::GetCol() 
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

void CBHMDBComboCtrl::SetCol(long nNewValue) 
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
	if (m_pDropDownRealWnd->GetColCount() >= m_nCol)
	{
		// Update the grid.
		int nRow, nCol;
		m_pDropDownRealWnd->GetCurrentCell(nRow, nCol);
		m_pDropDownRealWnd->SetCurrentCell(nRow, m_nCol);
	}
	else
		ThrowError(CTL_E_INVALIDPROPERTYARRAYINDEX, AFX_IDP_E_INVALIDPROPERTYARRAYINDEX);

	SetModifiedFlag();
	BoundPropertyChanged(dispidCol);
}

long CBHMDBComboCtrl::GetRows() 
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
	if (!m_pDropDownRealWnd)
		return m_nRows;

	m_nRows = m_pDropDownRealWnd->GetRowCount();

	return m_nRows;
}

void CBHMDBComboCtrl::SetRows(long nNewValue) 
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

	SetModifiedFlag();
	BoundPropertyChanged(dispidRows);
}

long CBHMDBComboCtrl::GetCols() 
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
	if (!m_pDropDownRealWnd)
		return m_nCols;

	// Update the cached value.
	m_nCols = m_pDropDownRealWnd->GetColCount();

	return m_nCols;
}

void CBHMDBComboCtrl::SetCols(long nNewValue) 
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

	SetModifiedFlag();
	BoundPropertyChanged(dispidCols);
}

BOOL CBHMDBComboCtrl::GetDroppedDown() 
/*
Routine Description:
	Get the visibility of the dropdown window.

Parameters:
	None.

Return Value:
	If that window is visible, return TRUE; Otherwise, return FALSE.
*/
{
	// This property is unavailable at design time.
	if (!AmbientUserMode())
		GetNotSupported();

	// Be careful in null pointer.
	if (!m_pDropDownRealWnd)
		return FALSE;

	return m_pDropDownRealWnd->IsWindowVisible();
}

void CBHMDBComboCtrl::SetDroppedDown(BOOL bNewValue) 
/*
Routine Description:
	Set the visibility of the dropdown window.

Parameters:
	bNewValue		The new value. If is TRUE, we should show the dropdown window; 
	Otherwise, we should hide the dropdown window.

Return Value:
	None.
*/
{
	// This property is unavailable at design time.
	if (!AmbientUserMode())
		SetNotSupported();

	// Be careful in null pointer.
	if (!m_pDropDownRealWnd)
		return;

	if (m_pDropDownRealWnd->IsWindowVisible() && !bNewValue)
		HideDropDownWnd();
	else if (!m_pDropDownRealWnd->IsWindowVisible() && bNewValue)
		ShowDropDownWnd();
}

BSTR CBHMDBComboCtrl::GetDataField() 
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

void CBHMDBComboCtrl::SetDataField(LPCTSTR lpszNewValue) 
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

LPDISPATCH CBHMDBComboCtrl::GetColumns() 
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

	// Set the callback pointers.
	pColumns->SetCtrlPtr(this);
	pColumns->m_pGrid = m_pDropDownRealWnd;

	return pColumns->GetIDispatch(FALSE);
}

long CBHMDBComboCtrl::GetVisibleRows() 
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

long CBHMDBComboCtrl::GetVisibleCols() 
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

void CBHMDBComboCtrl::OnHeadBackColorChanged() 
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
	if (m_pDropDownRealWnd)
		m_pDropDownRealWnd->SetHeaderBackColor(TranslateColor(m_clrHeadFore));
}

void CBHMDBComboCtrl::OnHeadForeColorChanged() 
/*
Routine Description:
	Update the control after the "HeadForeColor" property changed.

Parameters:
	None.

Return Value:
	None.
*/
{
	if (m_pDropDownRealWnd)
		m_pDropDownRealWnd->SetHeaderForeColor(TranslateColor(m_clrHeadBack));
}

BOOL CBHMDBComboCtrl::GetColumnHeader() 
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

void CBHMDBComboCtrl::SetColumnHeader(BOOL bNewValue) 
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
	if (m_pDropDownRealWnd)
	{
		m_pDropDownRealWnd->SetShowColHeader(m_bColumnHeader);
	}

	SetModifiedFlag();
	BoundPropertyChanged(dispidColumnHeader);
}

LPFONTDISP CBHMDBComboCtrl::GetHeadFont() 
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

void CBHMDBComboCtrl::SetHeadFont(LPFONTDISP newValue) 
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
	if (m_pDropDownRealWnd)
		SetHeaderHeight(m_pDropDownRealWnd->GetRowHeight(1));

	SetModifiedFlag();
	BoundPropertyChanged(dispidHeadFont);
}

BSTR CBHMDBComboCtrl::GetFieldSeparator() 
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

void CBHMDBComboCtrl::SetFieldSeparator(LPCTSTR lpszNewValue) 
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

void CBHMDBComboCtrl::OnListAutoPositionChanged() 
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

void CBHMDBComboCtrl::OnListWidthAutoSizeChanged() 
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

short CBHMDBComboCtrl::GetMaxDropDownItems() 
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

void CBHMDBComboCtrl::SetMaxDropDownItems(short nNewValue) 
/*
Routine Description:
	Set the max items can be visible in dropdown window simultaneously.

Parameters:
	nNewValue		The new value.

Return Value:
	None.
*/
{
	// Verify the new value.
	if (nNewValue < 1 || nNewValue > 100 || nNewValue < m_nMinDropDownItems)
		ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDITEMSLIMITATION);

	// Save the new value.
	m_nMaxDropDownItems = nNewValue;

	SetModifiedFlag();
	BoundPropertyChanged(dispidMaxDropDownItems);
}

short CBHMDBComboCtrl::GetMinDropDownItems() 
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

void CBHMDBComboCtrl::SetMinDropDownItems(short nNewValue) 
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

short CBHMDBComboCtrl::GetListWidth() 
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

void CBHMDBComboCtrl::SetListWidth(short nNewValue) 
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

	// Save the new value.
	m_nListWidth = nNewValue;

	SetModifiedFlag();
	BoundPropertyChanged(dispidListWidth);
}

long CBHMDBComboCtrl::GetHeaderHeight() 
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

	return m_nHeaderHeight;
}

void CBHMDBComboCtrl::SetHeaderHeight(long nNewValue) 
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
	if (m_pDropDownRealWnd)
		m_pDropDownRealWnd->SetHeaderHeight(m_nHeaderHeight);

	SetModifiedFlag();
	BoundPropertyChanged(dispidHeaderHeight);
}

void CBHMDBComboCtrl::OnForeColorChanged() 
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
	if (m_pDropDownRealWnd)
		m_pDropDownRealWnd->SetForeColor(TranslateColor(GetForeColor()));
	
	COleControl::OnForeColorChanged();
	BoundPropertyChanged(DISPID_FORECOLOR);
}

int CBHMDBComboCtrl::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (COleControl::OnCreate(lpCreateStruct) == -1)
		return -1;
	
	// Get the rect this control occupies.
	CRect rect;
	GetClientRect(&rect);

	if (AmbientUserMode())
	{
		// Only create child window at run time.

		// Create the edit window.
		rect.DeflateRect(2, 4, BUTTONWIDTH, 4);
		m_pEdit = new CComboEdit(this);
		rect.right -= BUTTONWIDTH;
		m_pEdit->Create(ES_AUTOHSCROLL | WS_VISIBLE | WS_CHILD, rect, this, IDC_EDIT);
		m_pEdit->SetWindowText("");

		// Create thte dropdown window.
		m_pDropDownRealWnd = new CDropDownRealWnd(this);
		m_pDropDownRealWnd->CreateEx(WS_EX_TOOLWINDOW,
			BHMGRID_CLASSNAME, NULL, WS_CHILD | WS_BORDER, rect, GetDesktopWindow(), NULL, NULL);

		// Init the dropdown window.
		InitDropDownWnd();

		// Init the edit window.
		SetText(m_strText);
	}
	
	return 0;
}

void CBHMDBComboCtrl::OnKillFocus(CWnd* pNewWnd) 
{
	COleControl::OnKillFocus(pNewWnd);
	
	// Hide the dropdown window when kill focus.
	HideDropDownWnd();
}

BOOL CBHMDBComboCtrl::OnGetPredefinedStrings(DISPID dispid, CStringArray* pStringArray, CDWordArray* pCookieArray) 
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

BOOL CBHMDBComboCtrl::OnGetPredefinedValue(DISPID dispid, DWORD dwCookie, VARIANT* lpvarOut) 
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

// force the height to fit the font height
BOOL CBHMDBComboCtrl::OnSetObjectRects(LPCRECT lpRectPos, LPCRECT lpRectClip) 
{
	// Out height is cacl out according to the font being used and can not be change by the
	// user.
	if (lpRectPos->right - lpRectPos->left != m_nFontHeight)
	{
		// The height is not desired.

		// Set the height to the desired value.
		SetControlSize(lpRectPos->right - lpRectPos->left, m_nFontHeight);

		// Update the position of th edit window.
		if (AmbientUserMode() && m_pEdit)
			m_pEdit->SetWindowPos(&wndTop, 2, 4, lpRectPos->right - lpRectPos->left - 4 - BUTTONWIDTH, 
				m_nFontHeight - 8, SWP_NOZORDER);
	}
	
	return COleControl::OnSetObjectRects(lpRectPos, lpRectClip);
}

// force the height to fit the font height
BOOL CBHMDBComboCtrl::OnSetExtent(LPSIZEL lpSizeL) 
{
	// Out height is cacl out according to the font being used and can not be change by the
	// user.
	if (IsWindow(m_hWnd))
	{
		// calc and set the right size myself
		CDC * pDC = GetDC();
		pDC->HIMETRICtoLP(lpSizeL);

		// Modify the height as our will.
		lpSizeL->cy = m_nFontHeight;

		// The width can not be too short.
		lpSizeL->cx = max(lpSizeL->cx , BUTTONWIDTH);

		pDC->LPtoHIMETRIC(lpSizeL);
		ReleaseDC(pDC);
	}

	BOOL bRet = COleControl::OnSetExtent(lpSizeL);
	InvalidateControl();
	
	return bRet;
}

void CBHMDBComboCtrl::OnSetFocus(CWnd* pOldWnd) 
{
	COleControl::OnSetFocus(pOldWnd);
	
	// If have edit window, set the focus to it.
	if (m_pEdit)
	{
		m_pEdit->SetFocus();
		m_pEdit->SetSel(0, 0xff);
	}
}

void CBHMDBComboCtrl::OnLButtonDown(UINT nFlags, CPoint point) 
{
	COleControl::OnLButtonDown(nFlags, point);

	if (m_rectButton.PtInRect(point))
	{
		// The button is clicked.

		m_bLButtonDown = TRUE;

		CDC * pDC = GetDC();
		CRect rect;
		GetClientRect(&rect);

		// Redraw the control to show pressed button.
		OnDraw(pDC, &rect, &rect);

		ReleaseDC(pDC);
		
		// Do the work as combo box.
		if (m_pEdit)
			OnDropDown();
	}
}

void CBHMDBComboCtrl::OnLButtonUp(UINT nFlags, CPoint point) 
{
	// The button is not pressed.

	m_bLButtonDown = FALSE;
	
	CDC * pDC = GetDC();
	CRect rect;
	GetClientRect(&rect);

	// Redraw the button.
	OnDraw(pDC, &rect, &rect);

	ReleaseDC(pDC);
	
	COleControl::OnLButtonUp(nFlags, point);
}

BOOL CBHMDBComboCtrl::PreTranslateMessage(MSG* pMsg) 
{
	// Fire the events.

	USHORT nCharShort;
	short nShiftState = _AfxShiftState();

	if (pMsg->hwnd == m_hWnd || ::IsChild(m_hWnd, pMsg->hwnd))
	switch (pMsg->message)
	{
	case WM_KEYDOWN:
		nCharShort = (USHORT)pMsg->wParam;
		FireKeyDown(&nCharShort, nShiftState);
		if (nCharShort == 0)
			return TRUE;

		pMsg->wParam = MAKEWPARAM(nCharShort, HIWORD(pMsg->wParam));

		if (pMsg->wParam == VK_DOWN || pMsg->wParam == VK_UP 
			|| pMsg->wParam == VK_ESCAPE || pMsg->wParam == VK_RETURN ||
			pMsg->wParam == VK_NEXT || pMsg->wParam == VK_PRIOR)
		{
			if (m_pDropDownRealWnd && m_pDropDownRealWnd->IsWindowVisible())
			{
				// The dropdown window is shown, send the keyboard message to it.
				m_pDropDownRealWnd->SendMessage(WM_KEYDOWN, pMsg->wParam, pMsg->lParam);
				return TRUE;
			}
		}
		
		if (nCharShort == VK_LEFT || nCharShort == VK_RIGHT || nCharShort == VK_UP || 
			nCharShort == VK_DOWN || nCharShort == VK_DELETE || nCharShort == VK_TAB)
		{
			// Send these special keyboard message to the window which has focus.
			GetFocus()->SendMessage(pMsg->message, pMsg->wParam, pMsg->lParam);
			return TRUE;
		}
		else if (nCharShort == VK_F4)
		{
			// F4 key toggles the visibility of dropdown window.
			OnDropDown();
			return TRUE;
		}

		break;

	case WM_KEYUP:
		nCharShort = (USHORT)pMsg->wParam;
		FireKeyUp(&nCharShort, nShiftState);
		if (nCharShort == 0)
			return TRUE;

		pMsg->wParam = MAKEWPARAM(nCharShort, HIWORD(pMsg->wParam));

		if (nCharShort == VK_LEFT || nCharShort == VK_RIGHT || nCharShort == VK_UP || 
			nCharShort == VK_DOWN || nCharShort == VK_DELETE || nCharShort == VK_TAB)
		{
			// The dropdown window is shown, so send the keyboard message to it.
			GetFocus()->SendMessage(pMsg->message, pMsg->wParam, pMsg->lParam);
			return TRUE;
		}
		
		break;

	case WM_CHAR:
		nCharShort = (USHORT)pMsg->wParam;
		FireKeyPress(&nCharShort);
		if (nCharShort == 0)
			return TRUE;

		pMsg->wParam = MAKEWPARAM(nCharShort, HIWORD(pMsg->wParam));
		
		break;

	case WM_SYSKEYDOWN:
		if (pMsg->wParam == VK_DOWN || pMsg->wParam == VK_UP)
		{
			// These key toggles the visibility of dropdown window.
			OnDropDown();
			return TRUE;
		}	
	}
	
	return COleControl::PreTranslateMessage(pMsg);
}

BSTR CBHMDBComboCtrl::GetFormat() 
/*
Routine Description:
	Get the text input mask.

Parameters:
	None.

Return Value:
	The text input mask.
*/
{
	return m_strFormat.AllocSysString();
}

void CBHMDBComboCtrl::SetFormat(LPCTSTR lpszNewValue) 
/*
Routine Description:
	Set the text input mask.

Parameters:
	lpszNewValue		The new value.

Return Value:
	None.
*/
{
	// Verify the new value. first.
	if (IsValidFormat(lpszNewValue))
	{
		// Save the new value.
		m_strFormat = lpszNewValue;

		if (AmbientUserMode())
		{
			// Calc the correspond mask value only at run time.
			CalcMask();
			return;
		}

		SetModifiedFlag();
		BoundPropertyChanged(dispidFormat);
	}
}

void CBHMDBComboCtrl::CalcMask()
/*
Routine Description:
	Calc the real mask we use from the one the user gave.

Parameters:
	None.

Return Value:
	None.
*/
{
	// If the prior char is delimiter.
	BOOL bDelimiter = FALSE;

	// The current char.
	char cCurrChr;

	int i, j;

	// The default string displayed if current value is invalid.
	CString strInit;

	// Empty the mask first.
	m_strMask.Empty();

	// Init the array.
	for (i = 0; i < 255; i++)
		m_bIsDelimiter[i] = FALSE;

	if (m_strFormat.IsEmpty())
	{
		// If format is empty, the mask is empty too.
		m_strMask.Empty();
		return;
	}

	for (i = 0, j = 0; i < m_strFormat.GetLength(); i++)
	{
		// Get the current char.
		cCurrChr = m_strFormat[i];

		if (!bDelimiter && cCurrChr == '/')
		{
			// '/' is a special char which means that the next char to it will be used
			// as delimiter anyway.
			bDelimiter = TRUE;
			continue;
		}

		// Add current char to the mask.
		m_strMask += cCurrChr;
		if (bDelimiter)
		// current char is delimiter.
			m_bIsDelimiter[j] = TRUE;

		// If current position is delimter, use it in default string. Otherwise, use '_' 
		// instead.
		strInit += m_bIsDelimiter[j] ? cCurrChr : '_';

		j ++;

		// The delimiter is processed.
		bDelimiter = FALSE;
	}
	
	// validate current text.
	if (!IsValidText(m_strText))
	// If current text is invalid, use the default text.
		SetText(strInit);
}

BOOL CBHMDBComboCtrl::IsValidText(CString strNewText)
/*
Routine Description:
	Determines if the given text matches the mask.

Parameters:
	strNewText		The text to be verified.

Return Value:
	If the given string is valid, return TRUE; Otherwise, return FALSE.
*/
{
	if (m_strMask.IsEmpty())
	// If mask is empty, the string is invalid only if its length is too long.
		return strNewText.GetLength() <= m_nMaxLength;

	if (strNewText.GetLength() != m_strMask.GetLength())
	// If the length of the string does not equal to that of mask, it must be invalid.
		return FALSE;

	int i;
	for (i = 0; i < m_strMask.GetLength(); i ++)
	// Check every char, once found invalid char, the string is invalid.
		if (!IsValidChar(strNewText[i], i) && strNewText[i] != ' '
			&& strNewText[i] != '_' && strNewText[i] != ' ')
			return FALSE;
	
	// Now the string is verified to be valid.
	return TRUE;
}

BOOL CBHMDBComboCtrl::IsValidFormat(CString strNewFormat)
/*
Routine Description:
	Determines if the given format is valid.

Parameters:
	strNewFormat		The format to be verified.

Return Value:
	If the format is valid, return TRUE; Otherwise, return FALSE.
*/
{
	BOOL bDelimiter = FALSE;
	char cCurrChr;

	if (strNewFormat.IsEmpty())
	// Empty format is always valid.
		return TRUE;

	if (strNewFormat.GetLength() > 255)
	// The length can not too long.
		return FALSE;

	for (int i = 0; i < strNewFormat.GetLength(); i++)
	{
		// Each char in format should be in the predefined char set.

		cCurrChr = strNewFormat[i];
		if (!bDelimiter && cCurrChr != '#' && cCurrChr != '9' &&
			cCurrChr != '?' && cCurrChr != 'C' && cCurrChr != 'A' &&
			cCurrChr != 'a' && cCurrChr != '/')
			return FALSE;

		bDelimiter = (cCurrChr == '/') ? TRUE : FALSE;
	}

	return TRUE;
}

BOOL CBHMDBComboCtrl::IsValidChar(UINT nChar, int nPosition)
/*
Routine Description:
	Determines if the given char at give position matches the mask.

Parameters:
	nChar		The char.
	nPosition	The position.

Return Value:
	If the char is valid, return TRUE; Otherwise, return FALSE.
*/
{
	if (nPosition >= m_strMask.GetLength())
	// The position should be valid first.
		return FALSE;

	if (m_bIsDelimiter[nPosition])
	// If this position containes delimiter, check it with the delimiter.
		return (char)nChar == m_strMask[nPosition];

	switch (m_strMask[nPosition])
	{
	case '#':
		return nChar >= '0' && nChar <= '9';
		break;

		case '9':
		return nChar == ' ' || (nChar >= '0' && nChar <= '9');
		break;

	case '?':
		return (nChar >= 'a' && nChar <= 'z') ||
			(nChar >= 'A' && nChar <= 'Z');
		break;

	case 'C':
		return nChar == ' ' || (nChar >= 'a' && nChar <= 'z') 
			|| (nChar >= 'A' && nChar <= 'Z');
		break;

	case 'A':
		return (nChar >= 'a' && nChar <= 'z') ||
			(nChar >= 'A' && nChar <= 'Z') ||
			(nChar >= '0' && nChar <= '9');
		break;

	case 'a':
		return (nChar >= 32 && nChar <= 126) ||
			(nChar >= 128 && nChar <= 255);
		break;
	};

	return FALSE;
}

int CBHMDBComboCtrl::GetColOrdinalFromIndex(int nColIndex)
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

	// The col index is invalid.
	return INVALID;
}

int CBHMDBComboCtrl::GetColOrdinalFromCol(int nCol)
/*
Routine Description:
	Get the col ordinal in bind info array from col number.

Parameters:
	nCol 		The col number.

Return Value:
	If the col number is valid, return the ordinal; Otherwise, return INVALID.
*/
{
	if (m_pDropDownRealWnd)
		return GetColOrdinalFromIndex(m_pDropDownRealWnd->GetColIndex(nCol));

	return INVALID;
}

int CBHMDBComboCtrl::GetColOrdinalFromField(int nField)
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

	// The field number is invalid.
	return INVALID;
}

int CBHMDBComboCtrl::GetFieldFromStr(CString str)
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

void CBHMDBComboCtrl::UpdateDivider()
/*
Routine Description:
	Update the divider lines.

Parameters:
	None.

Return Value:
	None.
*/
{
	if (m_pDropDownRealWnd)
	{
		m_pDropDownRealWnd->SetDividerType(m_nDividerType);
		m_pDropDownRealWnd->SetDividerStyle(m_nDividerStyle);
		m_pDropDownRealWnd->SetDividerColor(TranslateColor(m_clrDivider));
	}
}

void CBHMDBComboCtrl::InitCellFont()
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
		LOGFONT logCell;

		// Get the LOGFONT information.
		if (GetObject(hFont, sizeof(logCell), &logCell))
		{
			// Update the grid.
			if (m_pDropDownRealWnd)
				m_pDropDownRealWnd->SetCellFont(&logCell);
		}

		IFont * pIfont = m_fntCell.m_pFont;
		
		//release cached IFont pointer.
		if (m_hFont)
			pIfont->ReleaseHfont(m_hFont);
		
		CY cy;
		SIZEL size;
		int x, y;
		pIfont->get_Size(&cy);

		// Save the new font handle.
		m_hFont = hFont;

		// prevent the IFont interface from releasing cached font handle.
		pIfont->AddRefHfont(m_hFont);
		
		if (IsWindow(m_hWnd))
		{
			CDC * pDC = GetDC();
			GetControlSize(&x, &y);
			
			// Calc the font height.
			m_nFontHeight = cy.Lo * GetDeviceCaps(pDC->m_hDC, LOGPIXELSY) / 720000 + 12;
			
			size.cx = x;
			size.cy = m_nFontHeight;
			pDC->LPtoHIMETRIC(&size);

			// Resize the control according to the new font.
			OnSetExtent(&size);
//			SetControlSize(x, m_nFontHeight);
			
			ReleaseDC(pDC);
		}

		if (m_pEdit)
		// Set the font to edit window.
			m_pEdit->SendMessage(WM_SETFONT, (WPARAM)m_hFont, MAKELPARAM(TRUE, 0));
	}
}

LPFONTDISP CBHMDBComboCtrl::GetFont() 
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

void CBHMDBComboCtrl::SetFont(LPFONTDISP newValue) 
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
	if (m_pDropDownRealWnd)
		SetRowHeight(m_pDropDownRealWnd->GetRowHeight(0));
	else
		SetRowHeight(m_nFontHeight);

	SetModifiedFlag();
	BoundPropertyChanged(dispidFont);
}

void CBHMDBComboCtrl::SerializeColumnProp(CArchive &ar, BOOL bLoad)
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

		int nCells;

		// Load cell count.
		ar >> nCells;

		Cell * pCell;

		while (nCells > 0)
		{
			// Create a new cell object.
			pCell = new Cell;

			ar >> pCell->strText;
			
			// Add the new cell object into cached data.
			m_arCells.Add(pCell);

			nCells --;
		}
	}
}

void CBHMDBComboCtrl::ClearInterfaces()
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

void CBHMDBComboCtrl::FreeColumnMemory()
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

void CBHMDBComboCtrl::InitGridFromProp()
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
	if (!m_pDropDownRealWnd)
	{
		SetModifiedFlag();

		return;
	}

	// Save the update permission before changing it.
	BOOL bLockUpdate;
	bLockUpdate = m_pDropDownRealWnd->IsLockUpdate();

	// Prevent the grid from update.
	m_pDropDownRealWnd->LockUpdate();

	// Clear the grid.
	m_pDropDownRealWnd->SetColCount(0);
	m_pDropDownRealWnd->SetRowCount(0);
	m_pDropDownRealWnd->SetGroupCount(0);

	// Load the layout from cache.
	m_pDropDownRealWnd->LoadLayout(&m_arGroups, &m_arCols, m_nDataMode? &m_arCells : NULL);
	
	// If we have no predefined layout and the data mode is manual mode, insert some blank 
	// rows and cols.
	if (!m_pDropDownRealWnd->GetGroupCount() && !m_pDropDownRealWnd->GetColCount() && m_nDataMode)
	{
		m_pDropDownRealWnd->SetRowCount(m_nRows);
		m_pDropDownRealWnd->SetColCount(m_nCols);
	}

	InvalidateControl();

	// Restore the redraw permission.
	m_pDropDownRealWnd->LockUpdate(bLockUpdate);

	// Update the variables.
	m_nRows = m_pDropDownRealWnd->GetRowCount();
	m_nCols = m_pDropDownRealWnd->GetColCount();

	BoundPropertyChanged(dispidRows);
	BoundPropertyChanged(dispidCols);

	SetModifiedFlag();
}


BOOL CBHMDBComboCtrl::GetColumnInfo()
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

	// We can't do anything if the data source hasn't been set.
	if (m_spDataSource == NULL)
		return FALSE;

	// Reset the column count.
	m_nColumns = 0;

	// Get rowset from data source.
	hr = GetRowset();
	if (FAILED(hr) || !m_Rowset.m_spRowset)
		return FALSE;

	// Get the column information so we know the column names etc.
	hr = m_Rowset.CAccessorRowset<CManualAccessor>::GetColumnInfo((ULONG *)&m_nColumns, &m_pColumnInfo, &m_pStrings);
	if (FAILED(hr) || !m_pColumnInfo)
		return FALSE;

	// Check if we have bookmarks can be stored.
	if (m_pColumnInfo->iOrdinal == 0)
	{
		// The ordinal 0 represents bookmark field.
		m_bHaveBookmarks = TRUE;

		// Free old bookmark.
		FreeBookmarkMemory();

		// The size of each bookmark.
		m_nBookmarkSize = m_pColumnInfo->ulColumnSize;

		// There is one column to bind.
		nColumn = 1;
	}
	else
	{
		// We have not bookmark.
		m_bHaveBookmarks = FALSE;

		AfxMessageBox(IDS_ERROR_NOBOOKMARK, MB_ICONERROR | MB_OK);
		return FALSE;
	}

	// Now we can create the accessor and bind the data.
	// First of all we need to find the column ordinal from the name.

	// The total columns can be bound.
	int nColumnsBound = 0;

	for (; nColumn < m_nColumns; nColumn++)
	{
		// We only support these data type.
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

			// The count of bindable field is limited.
			if (m_nColumnsBind < 254)
			{
				// Create a new column data object.
				ColumnData * pCld = new ColumnData;

				// Init the column data object.
				pCld->strName = OLE2T(m_pColumnInfo[nColumn].pwszName);
				pCld->nColumn = nColumn;
				pCld->vt = m_pColumnInfo[nColumn].wType;

				// Add the new column data object into cache.
				m_apColumnData.Add(pCld);
			}

			break;
		}
	}

	m_nColumnsBind = min(m_nColumnsBind, m_apColumnData.GetSize());

	return m_apColumnData.GetSize() > 0;
}

HRESULT CBHMDBComboCtrl::GetRowset()
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

	// Close cached pointers.
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

void CBHMDBComboCtrl::InitDropDownWnd()
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

	// Update the colors.
	m_pDropDownRealWnd->SetForeColor(TranslateColor(GetForeColor()));
	m_pDropDownRealWnd->SetBackColor(TranslateColor(GetBackColor()));
	m_pDropDownRealWnd->SetHeaderForeColor(TranslateColor(m_clrHeadFore));
	m_pDropDownRealWnd->SetHeaderBackColor(TranslateColor(m_clrHeadBack));

	// Init the fonts.
	InitHeaderFont();
	InitCellFont();
	
	// Init the width and height.
	m_pDropDownRealWnd->SetHeaderHeight(m_nHeaderHeight);
	m_pDropDownRealWnd->SetRowHeight(m_nRowHeight);
	m_pDropDownRealWnd->SetDefColWidth(m_nDefColWidth);
	m_pDropDownRealWnd->SetShowGroupHeader(m_bGroupHeader);

	// Init cell data.
	InitGridFromProp();
	
	// This grid is readonly.
	m_pDropDownRealWnd->SetReadOnly(TRUE);

	m_pDropDownRealWnd->SetShowColHeader(m_bColumnHeader);
	m_pDropDownRealWnd->SetShowRowHeader(FALSE);
}

void CBHMDBComboCtrl::InitHeaderFont()
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
			if (m_pDropDownRealWnd)
				m_pDropDownRealWnd->SetHeaderFont(&logHeader);
		}
	}
}

void CBHMDBComboCtrl::GetCellValue(int nRow, int nCol, CString & strText)
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

int CBHMDBComboCtrl::GetRowColFromVariant(VARIANT va)
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

int CBHMDBComboCtrl::GetRowFromBmk(BYTE *bmk)
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

int CBHMDBComboCtrl::GetRowFromHRow(HROW hRow)
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

void CBHMDBComboCtrl::FetchNewRows()
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

	int i = 0, nFetchOnce;
	ULONG nRowsAvailable = 0;

	// This function is useful when bound with RDC datasource, because
	// the nRowsAvailable is just the number of the records in the local
	// cache.
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

	// Enlarge the bookmark array first, fill the new entries with null.
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

void CBHMDBComboCtrl::UndeleteRecordFromHRow(HROW hRow)
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

void CBHMDBComboCtrl::GetBmkOfRow(int nRow)
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
	if (nRow < 1 || nRow > m_apBookmark.GetSize() || m_apBookmark[nRow - 1])
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

void CBHMDBComboCtrl::HideDropDownWnd()
/*
Routine Description:
	Hide the dropdown window.

Parameters:
	None.

Return Value:
	None.
*/
{
	if (!m_pDropDownRealWnd || !m_pDropDownRealWnd->IsWindowVisible())
		return;

	// Fire the event.
	FireCloseUp();

	// Release the mouse hook.
	UnhookWindowsHookEx(m_hSystemHook);
	m_hSystemHook = NULL;

	// Hide the window.
	m_pDropDownRealWnd->ShowWindow(SW_HIDE);
}

void CBHMDBComboCtrl::UpdateDropDownWnd()
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

	// If we have not predefine layout, simply bind every field.

	int i;
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

void CBHMDBComboCtrl::DeleteRowFromSrc(int nRow)
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

	m_pDropDownRealWnd->Invalidate();
}

int CBHMDBComboCtrl::SearchText(CString strSearch, BOOL bAutoPosition)
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
	if (bAutoPosition && !m_bListAutoPosition)
		return INVALID;

	// Fire the event.
	FirePositionList(strSearch);

	// Calc the controling col ordinal.
	int nCol = m_pDropDownRealWnd->GetColFromName(m_strDataField);
	if (nCol == INVALID)
		return INVALID;

	// Calc the ordinal in bind information array.
	CString strValue;
	int nColIndex = m_pDropDownRealWnd->GetColIndex(nCol);

	for (int i = 0; i < m_arBindInfo.GetSize(); i ++)
		if (m_arBindInfo[i]->nColIndex == nColIndex)
			break;

	if (i >= m_arBindInfo.GetSize() || m_arBindInfo[i]->nDataField == -1)
	{
		// The controling is not bound to data source. Search the text manully.

		for (i = 1; i <= m_pDropDownRealWnd->GetRowCount(); i ++)
		{
			strValue = m_pDropDownRealWnd->GetCellText(i, nCol);
			if (!strValue.Left(strSearch.GetLength()).CompareNoCase(strSearch))
			{
				m_nRow = i;
				break;
			}
		}
		
		if (i > m_pDropDownRealWnd->GetRowCount())
			return INVALID;
	}
	else
	{
		// Search the text from data source.

		if (m_Rowset.m_spRowset == NULL)
			return INVALID;

		// Get the IRowsetFind pointer.
		CComQIPtr<IRowsetFind, &IID_IRowsetFind> pFind = m_Rowset.m_spRowset;
		if (!pFind)
			return INVALID;

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
			return INVALID;

		ULONG nRowsOb = 0;
		HROW* phRow = &rac.m_hRow;

		// First use partly match search. Ff fails, use exactly match search.
		// Not efficient but simple :)
		pFind->FindNextRow(NULL, rac.m_pAccessor->GetHAccessor(0), rac.
			m_pAccessor->GetBuffer(), DBCOMPAREOPS_EQ | DBCOMPAREOPS_CASEINSENSITIVE, 
			m_nBookmarkSize, (BYTE *)&bmFirst, 0, 1, &nRowsOb, &phRow);

		if (rac.m_hRow == 0)
			pFind->FindNextRow(NULL, rac.m_pAccessor->GetHAccessor(0), rac.
				m_pAccessor->GetBuffer(), DBCOMPAREOPS_BEGINSWITH | DBCOMPAREOPS_CASEINSENSITIVE, 
				m_nBookmarkSize, (BYTE *)&bmFirst, 0, 1, &nRowsOb, &phRow);

		if (rac.m_hRow == 0)
			return INVALID;

		// Do the clearing work.
		HROW hRow = rac.m_hRow;
		rac.ReleaseRows();
		m_Rowset.ReleaseRows();
		m_Rowset.m_hRow = hRow;
		if (FAILED(m_Rowset.GetData()))
			return INVALID;

		// Get the row ordinal.

		ULONG nRow;
		CBookmark<0> bmk;
		bmk.SetBookmark(m_nBookmarkSize, m_data.bookmark);
		if (FAILED(m_Rowset.GetApproximatePosition(&bmk, &nRow, NULL) || 
			(int)nRow > m_pDropDownRealWnd->GetRowCount()))
		nRow = m_pDropDownRealWnd->GetRowCount();

		m_nRow = nRow;
	}

	if (m_nRow > 0)
	{
		// Move to this row if required.
		if (bAutoPosition && m_bListAutoPosition)
			m_pDropDownRealWnd->MoveTo(m_nRow);

		return m_nRow;
	}
	else
		return INVALID;
}

void CBHMDBComboCtrl::Bind()
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

void CBHMDBComboCtrl::ShowDropDownWnd()
/*
Routine Description:
	Show the dropdown window.

Parameters:
	None.

Return Value:
	None.
*/
{
	// Fire the event first.
	FireDropDown();

	// If the data source has been changed, update the binding information and grid now.
	if(m_bDataSrcChanged)
		UpdateDropDownWnd();

	// We will show the dropdown window from the left bottom point.
	
	CPoint pt;
	CRect rect;
	GetClientRect(&rect);
	pt.x = rect.left;
	pt.y = rect.bottom;
	ClientToScreen(&pt);
	int cx = pt.x;
	int cy = pt.y;

	int x = 0, y = 0, yWndMin = 0, yWndMax = 0, xTotal = 2, yTotal = 0;
	
	// Can not exceed the border of screen.
	int nMaxWidth = GetSystemMetrics(SM_CXSCREEN) - cx;
	int nMaxHeight = GetSystemMetrics(SM_CYSCREEN) - cy;

	// The rect the window will occupies.
	CRect rcBounds(cx, cy, 0, 0);

	GetRowHeight();
	GetHeaderHeight();

	// Include the space of column header.
	if (m_bColumnHeader)
	{
		yWndMin += m_nHeaderHeight;
		yWndMax += m_nHeaderHeight;
		yTotal += m_nHeaderHeight;
	}

	// Calc the total height.
	yTotal += m_nRowHeight * m_pDropDownRealWnd->GetRowCount();

	// Take the max and min dropdown items property into account.
	yWndMax += m_nRowHeight * m_nMaxDropDownItems;
	yWndMin += m_nRowHeight * m_nMinDropDownItems;

	// Calc the total width.
	for (int i = 1; i <= m_pDropDownRealWnd->GetColCount(); i ++)
		xTotal += m_pDropDownRealWnd->GetColWidth(i);

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

	CString strText;
	m_pEdit->GetWindowText(strText);

	// Search the text being shown.
	SearchText(strText);
	if (m_nRow < 1)
	{
		// If not found, move to the first row.
		m_nRow = 1;
		m_pDropDownRealWnd->MoveTo(m_nRow);
	}

	// Bring the window to topmost.
	m_pDropDownRealWnd->SetWindowPos(&wndTopMost, 0, 0, 0, 0, SWP_NOSIZE | SWP_NOMOVE);
	m_pDropDownRealWnd->SetForegroundWindow();

	// Show the window.
	m_pDropDownRealWnd->ShowWindow(SW_SHOW);
	
	// Install the mouse hook.
	HookMouse(m_pDropDownRealWnd->m_hWnd);
}

void CBHMDBComboCtrl::OnDropDown()
/*
Routine Description:
	Process the dropdown message.

Parameters:
	None.

Return Value:
	None.
*/
{
	if (!m_pDropDownRealWnd)
		return;

	// Toggle the visibility of dropdown window.
	if (m_pDropDownRealWnd->IsWindowVisible())
		HideDropDownWnd();
	else
		ShowDropDownWnd();
}

void CBHMDBComboCtrl::HookMouse(HWND hWnd)
/*
Routine Description:
	Install the mouse hook.

Parameters:
	hWnd		The window for the hook procedure to monitor.

Return Value:
	None.
*/
{
	// Set the global variable used in hook procedure.
	SetHost(hWnd, m_pDropDownRealWnd->m_hWnd);

	// Install mouse hook.
	m_hSystemHook = SetWindowsHookEx(WH_MOUSE, MouseProc, (HINSTANCE)GetCurrentThread(), 0);
	
	if (!m_hSystemHook)
	{
#ifdef _DEBUG
		AfxMessageBox("Failed to install hook");
#endif
	}
}

BSTR CBHMDBComboCtrl::GetText() 
/*
Routine Description:
	Get the text being shown.

Parameters:
	None.

Return Value:
	The text being shown.
*/
{
	// If the edit window exists, get text from it; Otherwise, return the cached text.

	CString strResult;
	if (m_pEdit)
		m_pEdit->GetWindowText(m_strText);

	strResult = m_strText;

	return strResult.AllocSysString();
}

void CBHMDBComboCtrl::SetText(LPCTSTR lpszNewValue) 
/*
Routine Description:
	Set the text being shown.

Parameters:
	lpszNewValue		The new value.

Return Value:
	None.
*/
{
	// Validate the value first.
	if (m_pEdit && IsValidText(lpszNewValue))
	{
		if (!AmbientUserMode())
		{
			// Save the new value.
			m_strText = lpszNewValue;

			InvalidateControl();
		}

		// Update the edit window at run time.

		int nStart, nEnd;
		m_pEdit->GetSel(nStart, nEnd);
		m_pEdit->SetWindowText(lpszNewValue);
		m_pEdit->SetSel(nEnd, nEnd);
	}
}

OLE_COLOR CBHMDBComboCtrl::GetBackColor() 
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

void CBHMDBComboCtrl::SetBackColor(OLE_COLOR nNewValue) 
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
	if (m_pDropDownRealWnd)
		m_pDropDownRealWnd->SetBackColor(TranslateColor(GetBackColor()));
	
	SetModifiedFlag();
	BoundPropertyChanged(dispidBackColor);
	InvalidateControl();
}

HBRUSH CBHMDBComboCtrl::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor)
{
	if (m_pEdit && pWnd->m_hWnd == m_pEdit->m_hWnd && nCtlColor == CTLCOLOR_EDIT)
	{
		// Set the background and foreground color for edit window.

		// Delete cached brush.
		if (m_pBrhBack)
		{
			delete m_pBrhBack;
			m_pBrhBack = NULL;
		}

		// Create brush.
		m_pBrhBack = new CBrush(TranslateColor(m_clrBack));

		// Update the colors.
		pDC->SetTextColor(TranslateColor(GetForeColor()));
		pDC->SetBkColor(TranslateColor(m_clrBack));

		return (HBRUSH)*m_pBrhBack;
	}

	HBRUSH hbr = COleControl::OnCtlColor(pDC, pWnd, nCtlColor);

	return hbr;
}

void CBHMDBComboCtrl::AddItem(LPCTSTR Item, const VARIANT FAR& RowIndex) 
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
	if (nRow == INVALID)
	// Insert at the end default.
		nRow = m_pDropDownRealWnd->GetRowCount();

	if (nRow == 0 || nRow > m_pDropDownRealWnd->OnGetRecordCount() + 1)
		ThrowError(CTL_E_INVALIDPROPERTYARRAYINDEX, AFX_IDP_E_INVALIDPROPERTYARRAYINDEX);

	CString strText = Item, strField;
	int i = 1, j = 0;

	// Insert a row.
	m_pDropDownRealWnd->InsertRow(nRow);
	strField.Empty();

	while (j < strText.GetLength() && i <= m_pDropDownRealWnd->GetColCount())
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
	if (i <= m_pDropDownRealWnd->GetColCount())
		m_pDropDownRealWnd->SetCellText(nRow, i, strField);

	m_pDropDownRealWnd->RedrawRow(nRow);
}

void CBHMDBComboCtrl::RemoveItem(long RowIndex) 
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

void CBHMDBComboCtrl::Scroll(long Rows, long Cols) 
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

void CBHMDBComboCtrl::RemoveAll() 
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

void CBHMDBComboCtrl::MoveFirst() 
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

void CBHMDBComboCtrl::MovePrevious() 
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

void CBHMDBComboCtrl::MoveNext() 
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

void CBHMDBComboCtrl::MoveLast() 
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

void CBHMDBComboCtrl::MoveRecords(long Rows) 
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

void CBHMDBComboCtrl::ReBind() 
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

BOOL CBHMDBComboCtrl::IsItemInList() 
/*
Routine Description:
	Determines if the current text is in the item list.

Parameters:
	None.

Return Value:
	If is in, return TRUE; Otherwise, return FALSE.
*/
{
	if (m_bDataSrcChanged)
		UpdateDropDownWnd();

	// Search the text.
	return SearchText(GetText(), FALSE) != INVALID;
}

VARIANT CBHMDBComboCtrl::RowBookmark(long RowIndex) 
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

VARIANT CBHMDBComboCtrl::GetBookmark() 
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

void CBHMDBComboCtrl::SetBookmark(const VARIANT FAR& newValue) 
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

int CBHMDBComboCtrl::GetRowFromIndex(int nIndex)
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
		// Got it!
			return nRow;

	// Did not find.
	return INVALID;
}

void CBHMDBComboCtrl::GetBmkOfRow(int nRow, VARIANT *pVaRet)
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
	
	if (nRow > 0 && nRow <= m_apBookmark.GetSize())
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

int CBHMDBComboCtrl::GetRowFromBmk(const VARIANT *pBmk)
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
		// In manual mode, bookmark is just the row index.
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

long CBHMDBComboCtrl::GetStyle() 
/*
Routine Description:
	Get the control style.

Parameters:
	None.

Return Value:
	The style.
*/
{
	return m_nStyle;
}

void CBHMDBComboCtrl::SetStyle(long nNewValue) 
/*
Routine Description:
	Set the control style.

Parameters:
	nNewValue		The new value.

Return Value:
	None.
*/
{
	// Save the new value.
	m_nStyle = nNewValue;

	SetModifiedFlag();
	BoundPropertyChanged(dispidStyle);
}

int CBHMDBComboCtrl::OnMouseActivate(CWnd* pDesktopWnd, UINT nHitTest, UINT message) 
{
	// Inplace activate on mouse activate.
	if (AmbientUserMode())
		OnActivateInPlace (TRUE, NULL);
	
	return COleControl::OnMouseActivate(pDesktopWnd, nHitTest, message);
}

LPDISPATCH CBHMDBComboCtrl::GetGroups() 
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

void CBHMDBComboCtrl::OnDividerColorChanged() 
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
	UpdateDivider();
}

BOOL CBHMDBComboCtrl::GetGroupHeader() 
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

	return m_bGroupHeader;
}

void CBHMDBComboCtrl::SetGroupHeader(BOOL bNewValue) 
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

	SetModifiedFlag();
	BoundPropertyChanged(dispidGroupHeader);
}
