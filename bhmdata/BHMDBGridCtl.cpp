// BHMDBGridCtl.cpp : Implementation of the CBHMDBGridCtrl ActiveX Control class.

#include "stdafx.h"
#include "BHMData.h"
#include "BHMDBGridCtl.h"
#include "BHMDBGridPpgGeneral.h"
#include "BHMDBColumnPropPage.h"
#include "BHMDBGroupsPropPage.h"
#include "GridInner.h"
#include <oledberr.h>
#include <msstkppg.h>
#include "Columns.h"
#include "Group.h"
#include "Groups.h"
#include "SelBookmarks.h"
#include "oleimpl2.h"
#include "Groups.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


IMPLEMENT_DYNCREATE(CBHMDBGridCtrl, COleControl)

BEGIN_INTERFACE_MAP(CBHMDBGridCtrl, COleControl)
	// Map the interfaces to their implementations
	INTERFACE_PART(CBHMDBGridCtrl, IID_IRowPositionChange, RowPositionChange)
	INTERFACE_PART(CBHMDBGridCtrl, IID_IRowsetNotify, RowsetNotify)
END_INTERFACE_MAP()

STDMETHODIMP_(ULONG) CBHMDBGridCtrl::XRowPositionChange::AddRef(void)
{
	METHOD_PROLOGUE(CBHMDBGridCtrl, RowPositionChange)
	return pThis->ExternalAddRef();
}

STDMETHODIMP_(ULONG) CBHMDBGridCtrl::XRowPositionChange::Release(void)
{
	METHOD_PROLOGUE(CBHMDBGridCtrl, RowPositionChange)
	return pThis->ExternalRelease();
}

STDMETHODIMP CBHMDBGridCtrl::XRowPositionChange::QueryInterface(REFIID iid,
	LPVOID FAR* ppvObject)
{
	METHOD_PROLOGUE(CBHMDBGridCtrl, RowPositionChange)
	return pThis->ExternalQueryInterface(&iid, ppvObject);
}

STDMETHODIMP CBHMDBGridCtrl::XRowPositionChange::OnRowPositionChange(
	DBREASON eReason, DBEVENTPHASE ePhase, BOOL fCantDeny)
{
	METHOD_PROLOGUE(CBHMDBGridCtrl, RowPositionChange)

	if (ePhase == DBEVENTPHASE_DIDEVENT)
	{
		// Sync the current row position when that of data source changed.

		DBPOSITIONFLAGS dwPositionFlags;
		HRESULT hr;
		HCHAPTER hChapter = DB_NULL_HCHAPTER;

		// Get the new row position from data source.		
		hr = pThis->m_spRowPosition->GetRowPosition(&hChapter, &pThis->m_Rowset.m_hRow, &dwPositionFlags);
		if (FAILED(hr) || pThis->m_Rowset.m_hRow == DB_NULL_HROW)
		{
			return S_OK;
		}
		
		// Get row ordinal from HROW.
		int i = pThis->GetRowFromHRow(pThis->m_Rowset.m_hRow);
		int nRow, nCol;
		if (i != INVALID)
		{
			// Update the grid.
			pThis->m_pGridInner->GetCurrentCell(nRow, nCol);
			if (i != nRow)
			{
				pThis->m_pGridInner->MoveTo(i);
			}
		}

		pThis->m_Rowset.FreeRecordMemory();
		pThis->m_Rowset.ReleaseRows();
		
		return S_OK;
	}
	
	return DB_S_UNWANTEDPHASE;
}

STDMETHODIMP_(ULONG) CBHMDBGridCtrl::XRowsetNotify::AddRef(void)
{
	METHOD_PROLOGUE(CBHMDBGridCtrl, RowsetNotify)
	return pThis->ExternalAddRef();
}

STDMETHODIMP_(ULONG) CBHMDBGridCtrl::XRowsetNotify::Release(void)
{
	METHOD_PROLOGUE(CBHMDBGridCtrl, RowsetNotify)
	return pThis->ExternalRelease();
}

STDMETHODIMP CBHMDBGridCtrl::XRowsetNotify::QueryInterface(REFIID iid,
	LPVOID FAR* ppvObject)
{
	METHOD_PROLOGUE(CBHMDBGridCtrl, RowsetNotify)
	return pThis->ExternalQueryInterface(&iid, ppvObject);
}

STDMETHODIMP CBHMDBGridCtrl::XRowsetNotify::OnFieldChange(IRowset * prowset,
	HROW hRow, ULONG cColumns, ULONG rgColumns[], DBREASON eReason, DBEVENTPHASE ePhase, BOOL fCantDeny)
{
	METHOD_PROLOGUE(CBHMDBGridCtrl, RowsetNotify)

	if (DBEVENTPHASE_DIDEVENT == ePhase)// && prowset == pThis->m_Rowset.m_spRowset)
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
			if (nRow == INVALID || nRow > (int)pThis->m_pGridInner->GetRecordCount())
			{
				// This is a new record, fetch it :)
				pThis->m_bEndReached = FALSE;
				pThis->FetchNewRows();
				return S_OK;
			}

			// Redraw the correspond cells.
			for (ULONG i = 0; i < cColumns; i ++)
			{
				pThis->m_pGridInner->RedrawCell(nRow, rgColumns[i]);
			}
		}

		return S_OK;
		break;

		default:
			return DB_S_UNWANTEDREASON;
		}
	}

	return DB_S_UNWANTEDPHASE;
}

STDMETHODIMP CBHMDBGridCtrl::XRowsetNotify::OnRowChange(IRowset * prowset,
	ULONG cRows, const HROW rghRows[], DBREASON eReason, DBEVENTPHASE ePhase, BOOL fCantDeny)
{
	METHOD_PROLOGUE(CBHMDBGridCtrl, RowsetNotify)
	
	if (DBEVENTPHASE_ABOUTTODO == ePhase)// && pThis->m_Rowset.m_spRowset == prowset)
	{	
		switch (eReason)
		{
		case DBREASON_ROW_DELETE:
		case DBREASON_ROW_UNDOINSERT:
		{
			// Some rows have been deleted.
			int nRow;

			// We should remember the record ordinal for in some cases
			// we can not get the ordinal of the deleted records using
			// oledb interface.

			// This method may be buggy but I do not know why the datagrid
			// of MS can work fine. Maybe he cached this information like me?
			for (ULONG i = 0; i < cRows; i ++)
			{
				// Calc the row ordinal from hrow.
				nRow = pThis->GetRowFromHRow(rghRows[i]);
				if (nRow != INVALID)
				{
					// Cache the row info.
					int j;
					for (j = 0; j < pThis->m_arRowToDelete.GetSize() && pThis->
						m_arRowToDelete[i] > nRow; j ++);

					pThis->m_arHRowToDelete.InsertAt(j, rghRows[i]);
					pThis->m_arRowToDelete.InsertAt(j, nRow);
				}
			}
		}

		default:
			return S_OK;
		}
	}

	if (DBEVENTPHASE_DIDEVENT == ePhase)// && pThis->m_Rowset.m_spRowset == prowset)
	{
		switch (eReason)
		{
		case DBREASON_ROW_DELETE:
		case DBREASON_ROW_UNDOINSERT:
		{
			for (ULONG i = 0; i < cRows; i ++)
			{
				for (int j = 0; j < pThis->m_arHRowToDelete.GetSize(); j ++)
				{
					// Update the grid.
					if (pThis->m_arHRowToDelete[j] == rghRows[i])
						pThis->DeleteRowFromSrc(pThis->m_arRowToDelete[j]);
				}
			}
		}

		return S_OK;
		break;

		case DBREASON_ROW_ASYNCHINSERT:
		case DBREASON_ROW_INSERT:
		{
			// Some rows have been inserted.

			pThis->m_bEndReached = FALSE;
			pThis->FetchNewRows();
		}

		return S_OK;
		break;

		// In some cases I can not get the currect record ordinal, so I
		// have to cache the information for undeletion.
		case DBREASON_ROW_UNDODELETE:
		{
			for (int j = 0; j < pThis->m_arHRowToDelete.GetSize(); j ++)
			{
				for (ULONG i = 0; i < cRows; i ++)
				{
					// Restore the original rows.
					if (pThis->m_arHRowToDelete[j] == rghRows[i])
						pThis->UndeleteRecordFromHRow(rghRows[i], pThis->m_arRowToDelete[j]);
				}
			}
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
					pThis->m_pGridInner->RedrawRow(nRow);
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

STDMETHODIMP CBHMDBGridCtrl::XRowsetNotify::OnRowsetChange(IRowset * prowset,
	DBREASON eReason, DBEVENTPHASE ePhase, BOOL fCantDeny)
{
	METHOD_PROLOGUE(CBHMDBGridCtrl, RowsetNotify)

	if (DBEVENTPHASE_DIDEVENT == ePhase)// && prowset == pThis->m_Rowset.m_spRowset)
	{
		switch (eReason)
		{
		case DBREASON_ROWSET_RELEASE:
		case DBREASON_ROWSET_CHANGED:
		{
			// Rowset has been changed.
			pThis->m_bDataSrcChanged = TRUE;
			pThis->InvalidateControl();
		}

		return S_OK;
		break;
			
		default:
			return DB_S_UNWANTEDREASON;
		}
	}

	return DB_S_UNWANTEDPHASE;
}

STDMETHODIMP_(ULONG) CBHMDBGridCtrl::XDataSourceListener::AddRef(void)
{
	METHOD_PROLOGUE(CBHMDBGridCtrl, DataSourceListener)
	return pThis->ExternalAddRef();
}

STDMETHODIMP_(ULONG) CBHMDBGridCtrl::XDataSourceListener::Release(void)
{
	METHOD_PROLOGUE(CBHMDBGridCtrl, DataSourceListener)
	return pThis->ExternalRelease();
}

STDMETHODIMP CBHMDBGridCtrl::XDataSourceListener::QueryInterface(REFIID iid,
	LPVOID FAR* ppvObject)
{
	METHOD_PROLOGUE(CBHMDBGridCtrl, DataSourceListener)
	return pThis->ExternalQueryInterface(&iid, ppvObject);
}

STDMETHODIMP CBHMDBGridCtrl::XDataSourceListener::dataMemberChanged(BSTR bstrDM)
{
	METHOD_PROLOGUE(CBHMDBGridCtrl, RowsetNotify)

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
		pThis->UpdateControl();
	}
	
	return S_OK;
}

STDMETHODIMP CBHMDBGridCtrl::XDataSourceListener::dataMemberAdded(BSTR bstrDM)
{
	return S_OK;
}

STDMETHODIMP CBHMDBGridCtrl::XDataSourceListener::dataMemberRemoved(BSTR bstrDM)
{
	METHOD_PROLOGUE(CBHMDBGridCtrl, RowsetNotify)

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

BEGIN_MESSAGE_MAP(CBHMDBGridCtrl, COleControl)
	//{{AFX_MSG_MAP(CBHMDBGridCtrl)
	ON_WM_CREATE()
	ON_WM_KILLFOCUS()
	ON_WM_MOUSEACTIVATE()
	//}}AFX_MSG_MAP
	ON_OLEVERB(AFX_IDS_VERB_PROPERTIES, OnProperties)
	ON_REGISTERED_MESSAGE(CBHMDataApp::m_nBringMsg, OnBringMsg)
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// Dispatch map

BEGIN_DISPATCH_MAP(CBHMDBGridCtrl, COleControl)
	//{{AFX_DISPATCH_MAP(CBHMDBGridCtrl)
	DISP_PROPERTY_NOTIFY(CBHMDBGridCtrl, "HeadForeColor", m_clrHeadFore, OnHeadForeColorChanged, VT_COLOR)
	DISP_PROPERTY_NOTIFY(CBHMDBGridCtrl, "HeadBackColor", m_clrHeadBack, OnHeadBackColorChanged, VT_COLOR)
	DISP_PROPERTY_NOTIFY(CBHMDBGridCtrl, "DividerColor", m_clrDivider, OnDividerColorChanged, VT_COLOR)
	DISP_PROPERTY_EX(CBHMDBGridCtrl, "CaptionAlignment", GetCaptionAlignment, SetCaptionAlignment, VT_I4)
	DISP_PROPERTY_EX(CBHMDBGridCtrl, "DividerType", GetDividerType, SetDividerType, VT_I4)
	DISP_PROPERTY_EX(CBHMDBGridCtrl, "AllowAddNew", GetAllowAddNew, SetAllowAddNew, VT_BOOL)
	DISP_PROPERTY_EX(CBHMDBGridCtrl, "AllowDelete", GetAllowDelete, SetAllowDelete, VT_BOOL)
	DISP_PROPERTY_EX(CBHMDBGridCtrl, "AllowUpdate", GetAllowUpdate, SetAllowUpdate, VT_BOOL)
	DISP_PROPERTY_EX(CBHMDBGridCtrl, "AllowRowSizing", GetAllowRowSizing, SetAllowRowSizing, VT_BOOL)
	DISP_PROPERTY_EX(CBHMDBGridCtrl, "RecordSelectors", GetRecordSelectors, SetRecordSelectors, VT_BOOL)
	DISP_PROPERTY_EX(CBHMDBGridCtrl, "AllowColumnSizing", GetAllowColumnSizing, SetAllowColumnSizing, VT_BOOL)
	DISP_PROPERTY_EX(CBHMDBGridCtrl, "ColumnHeader", GetColumnHeader, SetColumnHeader, VT_BOOL)
	DISP_PROPERTY_EX(CBHMDBGridCtrl, "HeadFont", GetHeadFont, SetHeadFont, VT_FONT)
	DISP_PROPERTY_EX(CBHMDBGridCtrl, "Font", GetFont, SetFont, VT_FONT)
	DISP_PROPERTY_EX(CBHMDBGridCtrl, "Col", GetCol, SetCol, VT_I4)
	DISP_PROPERTY_EX(CBHMDBGridCtrl, "Row", GetRow, SetRow, VT_I4)
	DISP_PROPERTY_EX(CBHMDBGridCtrl, "DataMode", GetDataMode, SetDataMode, VT_I4)
	DISP_PROPERTY_EX(CBHMDBGridCtrl, "VisibleCols", GetVisibleCols, SetNotSupported, VT_I4)
	DISP_PROPERTY_EX(CBHMDBGridCtrl, "VisibleRows", GetVisibleRows, SetNotSupported, VT_I4)
	DISP_PROPERTY_EX(CBHMDBGridCtrl, "LeftCol", GetLeftCol, SetLeftCol, VT_I4)
	DISP_PROPERTY_EX(CBHMDBGridCtrl, "Cols", GetCols, SetCols, VT_I4)
	DISP_PROPERTY_EX(CBHMDBGridCtrl, "Rows", GetRows, SetRows, VT_I4)
	DISP_PROPERTY_EX(CBHMDBGridCtrl, "FieldSeparator", GetFieldSeparator, SetFieldSeparator, VT_BSTR)
	DISP_PROPERTY_EX(CBHMDBGridCtrl, "DividerStyle", GetDividerStyle, SetDividerStyle, VT_I4)
	DISP_PROPERTY_EX(CBHMDBGridCtrl, "FirstRow", GetFirstRow, SetFirstRow, VT_I4)
	DISP_PROPERTY_EX(CBHMDBGridCtrl, "DefColWidth", GetDefColWidth, SetDefColWidth, VT_I4)
	DISP_PROPERTY_EX(CBHMDBGridCtrl, "RowHeight", GetRowHeight, SetRowHeight, VT_I4)
	DISP_PROPERTY_EX(CBHMDBGridCtrl, "Redraw", GetRedraw, SetRedraw, VT_BOOL)
	DISP_PROPERTY_EX(CBHMDBGridCtrl, "AllowColumnMoving", GetAllowColumnMoving, SetAllowColumnMoving, VT_BOOL)
	DISP_PROPERTY_EX(CBHMDBGridCtrl, "SelectByCell", GetSelectByCell, SetSelectByCell, VT_BOOL)
	DISP_PROPERTY_EX(CBHMDBGridCtrl, "Columns", GetColumns, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX(CBHMDBGridCtrl, "DataSource", GetDataSource, SetDataSource, VT_UNKNOWN)
	DISP_PROPERTY_EX(CBHMDBGridCtrl, "DataMember", GetDataMember, SetDataMember, VT_BSTR)
	DISP_PROPERTY_EX(CBHMDBGridCtrl, "RowChanged", GetRowChanged, SetRowChanged, VT_BOOL)
	DISP_PROPERTY_EX(CBHMDBGridCtrl, "HeaderHeight", GetHeaderHeight, SetHeaderHeight, VT_I4)
	DISP_PROPERTY_EX(CBHMDBGridCtrl, "BackColor", GetBackColor, SetBackColor, VT_COLOR)
	DISP_PROPERTY_EX(CBHMDBGridCtrl, "Bookmark", GetBookmark, SetBookmark, VT_VARIANT)
	DISP_PROPERTY_EX(CBHMDBGridCtrl, "SelBookmarks", GetSelBookmarks, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX(CBHMDBGridCtrl, "FrozenCols", GetFrozenCols, SetFrozenCols, VT_I4)
	DISP_PROPERTY_EX(CBHMDBGridCtrl, "Groups", GetGroups, SetNotSupported, VT_DISPATCH)
	DISP_PROPERTY_EX(CBHMDBGridCtrl, "GroupHeader", GetGroupHeader, SetGroupHeader, VT_BOOL)
	DISP_FUNCTION(CBHMDBGridCtrl, "AddItem", AddItem, VT_EMPTY, VTS_BSTR VTS_VARIANT)
	DISP_FUNCTION(CBHMDBGridCtrl, "RemoveItem", RemoveItem, VT_EMPTY, VTS_I4)
	DISP_FUNCTION(CBHMDBGridCtrl, "Scroll", Scroll, VT_EMPTY, VTS_I4 VTS_I4)
	DISP_FUNCTION(CBHMDBGridCtrl, "RemoveAll", RemoveAll, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CBHMDBGridCtrl, "MoveFirst", MoveFirst, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CBHMDBGridCtrl, "MovePrevious", MovePrevious, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CBHMDBGridCtrl, "MoveNext", MoveNext, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CBHMDBGridCtrl, "MoveLast", MoveLast, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CBHMDBGridCtrl, "MoveRecords", MoveRecords, VT_EMPTY, VTS_I4)
	DISP_FUNCTION(CBHMDBGridCtrl, "IsAddRow", IsAddRow, VT_BOOL, VTS_NONE)
	DISP_FUNCTION(CBHMDBGridCtrl, "Update", Update, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CBHMDBGridCtrl, "CancelUpdate", CancelUpdate, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CBHMDBGridCtrl, "ReBind", ReBind, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CBHMDBGridCtrl, "RowBookmark", RowBookmark, VT_VARIANT, VTS_I4)
	DISP_STOCKFUNC_REFRESH()
	DISP_STOCKPROP_FORECOLOR()
	DISP_STOCKPROP_HWND()
	DISP_STOCKPROP_CAPTION()
	DISP_STOCKPROP_BORDERSTYLE()
	//}}AFX_DISPATCH_MAP
	DISP_FUNCTION_ID(CBHMDBGridCtrl, "AboutBox", DISPID_ABOUTBOX, AboutBox, VT_EMPTY, VTS_NONE)
	DISP_PROPERTY_STOCK(CBHMDBGridCtrl, "CtrlType", 255, CBHMDBGridCtrl::GetCtrlType, COleControl::SetNotSupported, VT_I2)
END_DISPATCH_MAP()


/////////////////////////////////////////////////////////////////////////////
// Event map

BEGIN_EVENT_MAP(CBHMDBGridCtrl, COleControl)
	//{{AFX_EVENT_MAP(CBHMDBGridCtrl)
	EVENT_CUSTOM("BtnClick", FireBtnClick, VTS_NONE)
	EVENT_CUSTOM("AfterColUpdate", FireAfterColUpdate, VTS_I2)
	EVENT_CUSTOM("BeforeColUpdate", FireBeforeColUpdate, VTS_I2  VTS_BSTR  VTS_PI2)
	EVENT_CUSTOM("BeforeInsert", FireBeforeInsert, VTS_PI2)
	EVENT_CUSTOM("BeforeUpdate", FireBeforeUpdate, VTS_PI2)
	EVENT_CUSTOM("ColResize", FireColResize, VTS_I2  VTS_PI2)
	EVENT_CUSTOM("RowResize", FireRowResize, VTS_PI2)
	EVENT_CUSTOM("RowColChange", FireRowColChange, VTS_NONE)
	EVENT_CUSTOM("Scroll", FireScroll, VTS_PI2)
	EVENT_CUSTOM("Change", FireChange, VTS_NONE)
	EVENT_CUSTOM("AfterDelete", FireAfterDelete, VTS_PI2)
	EVENT_CUSTOM("BeforeDelete", FireBeforeDelete, VTS_PI2  VTS_PI2)
	EVENT_CUSTOM("AfterUpdate", FireAfterUpdate, VTS_PI2)
	EVENT_CUSTOM("AfterInsert", FireAfterInsert, VTS_PI2)
	EVENT_CUSTOM("ColSwap", FireColSwap, VTS_I2  VTS_I2  VTS_PI2)
	EVENT_CUSTOM("UpdateError", FireUpdateError, VTS_NONE)
	EVENT_CUSTOM("BeforeRowColChange", FireBeforeRowColChange, VTS_PI2)
	EVENT_CUSTOM("ScrollAfter", FireScrollAfter, VTS_NONE)
	EVENT_CUSTOM("BeforeDrawCell", FireBeforeDrawCell, VTS_I4  VTS_I4  VTS_PVARIANT  VTS_PCOLOR  VTS_PI2)
	EVENT_STOCK_KEYDOWN()
	EVENT_STOCK_KEYPRESS()
	EVENT_STOCK_CLICK()
	EVENT_STOCK_DBLCLICK()
	EVENT_STOCK_KEYUP()
	EVENT_STOCK_MOUSEDOWN()
	EVENT_STOCK_MOUSEMOVE()
	EVENT_STOCK_MOUSEUP()
	//}}AFX_EVENT_MAP
END_EVENT_MAP()


/////////////////////////////////////////////////////////////////////////////
// Property pages

// TODO: Add more property pages as needed.  Remember to increase the count!
BEGIN_PROPPAGEIDS(CBHMDBGridCtrl, 5)
	PROPPAGEID(CBHMDBGridPropPageGeneral::guid)
	PROPPAGEID(CBHMDBGroupsPropPage::guid)
	PROPPAGEID(CBHMDBColumnPropPage::guid)
	PROPPAGEID(CLSID_StockColorPage)
	PROPPAGEID(CLSID_StockFontPage)
END_PROPPAGEIDS(CBHMDBGridCtrl)


/////////////////////////////////////////////////////////////////////////////
// Initialize class factory and guid

IMPLEMENT_OLECREATE_EX(CBHMDBGridCtrl, "BHMDATA.BHMDBGridCtrl.1",
	0xbf91b125, 0x6a84, 0x11d3, 0xa7, 0xfe, 0, 0x80, 0xc8, 0x76, 0x3f, 0xa4)


/////////////////////////////////////////////////////////////////////////////
// Type library ID and version

IMPLEMENT_OLETYPELIB(CBHMDBGridCtrl, _tlid, _wVerMajor, _wVerMinor)


/////////////////////////////////////////////////////////////////////////////
// Interface IDs

const IID BASED_CODE IID_DBHMDBGrid =
		{ 0xbf91b123, 0x6a84, 0x11d3, { 0xa7, 0xfe, 0, 0x80, 0xc8, 0x76, 0x3f, 0xa4 } };
const IID BASED_CODE IID_DBHMDBGridEvents =
		{ 0xbf91b124, 0x6a84, 0x11d3, { 0xa7, 0xfe, 0, 0x80, 0xc8, 0x76, 0x3f, 0xa4 } };
const IID BASED_CODE IID_IDataSource =
		{ 0x7c0ffab3, 0xcd84, 0x11d0, { 0x94, 0x9a, 0, 0xa0, 0xc9, 0x11, 0x10, 0xed } };


/////////////////////////////////////////////////////////////////////////////
// Control type information

static const DWORD BASED_CODE _dwBHMDBGridOleMisc =
	OLEMISC_ACTIVATEWHENVISIBLE |
	OLEMISC_SETCLIENTSITEFIRST |
	OLEMISC_INSIDEOUT |
	OLEMISC_CANTLINKINSIDE |
	OLEMISC_RECOMPOSEONRESIZE;

IMPLEMENT_OLECTLTYPE(CBHMDBGridCtrl, IDS_BHMDBGRID, _dwBHMDBGridOleMisc)


/////////////////////////////////////////////////////////////////////////////
// CBHMDBGridCtrl::CBHMDBGridCtrlFactory::UpdateRegistry -
// Adds or removes system registry entries for CBHMDBGridCtrl

BOOL CBHMDBGridCtrl::CBHMDBGridCtrlFactory::UpdateRegistry(BOOL bRegister)
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
			IDS_BHMDBGRID,
			IDB_BHMDBGRID,
			afxRegApartmentThreading,
			_dwBHMDBGridOleMisc,
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
// CBHMDBGridCtrl::CBHMDBGridCtrlFactory::VerifyUserLicense -
// Checks for existence of a user license

BOOL CBHMDBGridCtrl::CBHMDBGridCtrlFactory::VerifyUserLicense()
{
	return AfxVerifyLicFile(AfxGetInstanceHandle(), _szLicFileName,
		_szLicString);
}


/////////////////////////////////////////////////////////////////////////////
// CBHMDBGridCtrl::CBHMDBGridCtrlFactory::GetLicenseKey -
// Returns a runtime licensing key

BOOL CBHMDBGridCtrl::CBHMDBGridCtrlFactory::GetLicenseKey(DWORD dwReserved,
	BSTR FAR* pbstrKey)
{
	if (pbstrKey == NULL)
		return FALSE;

	*pbstrKey = SysAllocString(_szLicString);
	return (*pbstrKey != NULL);
}


/////////////////////////////////////////////////////////////////////////////
// CBHMDBGridCtrl::CBHMDBGridCtrl - Constructor

CBHMDBGridCtrl::CBHMDBGridCtrl() : m_fntCell(NULL), m_fntHeader(NULL)
{
	InitializeIIDs(&IID_DBHMDBGrid, &IID_DBHMDBGridEvents);

	m_bHasLayout = FALSE;

	m_bIsPosSync = TRUE;
	m_bEndReached = TRUE;
	m_bDataSrcChanged = FALSE;
	m_dwCookieRPC = m_dwCookieRN = 0;
	m_bHaveBookmarks = FALSE;
	m_pColumnInfo = NULL;
	m_pStrings = NULL;
	m_nColumns = m_nColumnsBind = 0;
	m_nBookmarkSize = 0;

	m_spDataSource = NULL;
	m_spRowPosition = NULL;

	m_pGridInner = NULL;
	m_bDataSrcChanged = FALSE;
	m_spDataSource = NULL;
	m_nCol = m_nRow = 0;
	m_bRedraw = TRUE;
	m_nColsUsed = 0;
	m_bDroppedDown = FALSE;
	m_hSystemHook = NULL;
	m_hExternalDropDownWnd = NULL;
}


/////////////////////////////////////////////////////////////////////////////
// CBHMDBGridCtrl::~CBHMDBGridCtrl - Destructor

CBHMDBGridCtrl::~CBHMDBGridCtrl()
{
	// Release the mouse hook.
	if (m_hSystemHook)
	{
		UnhookWindowsHookEx(m_hSystemHook);
		m_hSystemHook = NULL;
	}

	// Delete the internal grid.
	if (m_pGridInner)
	{
		m_pGridInner->DestroyWindow();
		delete m_pGridInner;
		m_pGridInner = NULL;
	}

	// Clear the groups layout array.
	for (int i = 0; i < m_arGroups.GetSize(); i ++)
		delete m_arGroups[i];

	m_arGroups.RemoveAll();

	// Clear the cells layout array.
	for (i = 0; i < m_arCells.GetSize(); i ++)
		delete m_arCells[i];

	m_arCells.RemoveAll();

	// Clear the cached bind information array.
	for (i = 0; i < m_arBindInfo.GetSize(); i ++)
		delete m_arBindInfo[i];

	m_arBindInfo.RemoveAll();

	// Clear the cols array.
	i = m_arCols.GetSize();

	for (; i > 0; i --)
		delete m_arCols[i - 1];

	m_arCols.RemoveAll();

	// Clear the cached interface pointers.
	ClearInterfaces();
}


/////////////////////////////////////////////////////////////////////////////
// CBHMDBGridCtrl::OnDraw - Drawing function

void CBHMDBGridCtrl::OnDraw(
			CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid)
{
	if (m_bDataSrcChanged)
		UpdateControl();

	CRect rtGrid = rcBounds, rtCaption = rcBounds;
	CFont * pFontOld;
	CSize szCaption;

	int nSave = pdc->SaveDC();

	if (InternalGetText().GetLength())
	{
		// Draw the title bar, use the level height as the bar height.
		int nHeaderHeight = m_pGridInner->GetLevelHeight();

		// Save the old font handle.
		pFontOld = SelectFontObject(pdc, m_fntHeader);

		// Draw 3D bar.
		rtCaption.bottom = rtCaption.top + nHeaderHeight;
		pdc->Draw3dRect(&rtCaption, GetSysColor(COLOR_3DLIGHT), GetSysColor(COLOR_3DDKSHADOW));
		rtCaption.DeflateRect(1, 1);
		pdc->Draw3dRect(&rtCaption, GetSysColor(COLOR_3DHILIGHT), GetSysColor(COLOR_3DSHADOW));
		rtCaption.DeflateRect(1, 1);
		pdc->FillRect(&rtCaption, &CBrush(TranslateColor(m_clrHeadBack)));

		pdc->SetBkColor(TranslateColor(m_clrHeadBack));
		pdc->SetTextColor(TranslateColor(m_clrHeadFore));
		
		rtCaption.DeflateRect(3, 0);

		// Draw title text.
		switch (m_nCapitonAlignment)
		{
		case 0:
			pdc->DrawText(InternalGetText(), rtCaption, DT_VCENTER | DT_SINGLELINE | DT_LEFT);
			break;

		case 1:
			pdc->DrawText(InternalGetText(), rtCaption, DT_VCENTER | DT_SINGLELINE | DT_RIGHT);
			break;

		case 2:
			pdc->DrawText(InternalGetText(), rtCaption, DT_VCENTER | DT_SINGLELINE | DT_CENTER);
		}

		pdc->SelectObject(pFontOld);

		rtCaption.InflateRect(5, 2);
		rtGrid.top += nHeaderHeight;
	}

	pdc->RestoreDC(nSave);

	// If the internal grid is not created, create it here.
	if (!m_pGridInner) 
		CreateDrawGrid();		

	if (!AmbientUserMode())
	{
		// At design time, draw the grid by myself.
		m_pGridInner->SetGridRect(&rtGrid);
		m_pGridInner->LockUpdate(FALSE);
		m_pGridInner->OnGridDraw(pdc);
		m_pGridInner->LockUpdate();
	}
	else
	{
		// At run time, let the grid draw itself.
		m_pGridInner->Invalidate();
	}
}


/////////////////////////////////////////////////////////////////////////////
// CBHMDBGridCtrl::DoPropExchange - Persistence support

void CBHMDBGridCtrl::DoPropExchange(CPropExchange* pPX)
{
	ExchangeVersion(pPX, MAKELONG(_wVerMinor, _wVerMajor));
	ASSERT_POINTER(pPX, CPropExchange);

	ExchangeExtent(pPX);
	ExchangeStockProps(pPX);

	PX_Font(pPX, "Font", m_fntCell, &_BHMFontDescDefault);
	PX_Font(pPX, "HeaderFont", m_fntHeader, &_BHMFontDescDefault);
	PX_Long(pPX, "HeaderHeight", m_nHeaderHeight, 20);
	PX_Long(pPX, "RowHeight", m_nRowHeight, 20);
	PX_Long(pPX, "DataMode", m_nDataMode, 0);
	PX_Color(pPX, "HeadForeColor", m_clrHeadFore, RGB(0, 0, 0));
	PX_Color(pPX, "HeadBackColor", m_clrHeadBack, GetSysColor(COLOR_3DFACE));
	PX_Long(pPX, "DefColWidth", m_nDefColWidth, 100);
	PX_Bool(pPX, "AllowAddNew", m_bAllowAddNew, FALSE);
	PX_Bool(pPX, "AllowDelete", m_bAllowDelete, FALSE);
	PX_Bool(pPX, "AllowUpdate", m_bAllowUpdate, FALSE);
	PX_Bool(pPX, "RecordSelectors", m_bRecordSelectors, TRUE);
	PX_Long(pPX, "Rows", m_nRows, 2);
	PX_Long(pPX, "Cols", m_nCols, 2);
	PX_Long(pPX, "CapitonAlignment", m_nCapitonAlignment, 2);
	PX_Long(pPX, "DividerType", m_nDividerType, 3);
	PX_Long(pPX, "DividerStyle", m_nDividerStyle, 0);
	PX_Bool(pPX, "AllowRowSizing", m_bAllowRowSizing, TRUE);
	PX_Bool(pPX, "AllowColumnSizing", m_bAllowColumnSizing, TRUE);
	PX_Bool(pPX, "ColumnHeader", m_bColumnHeader, TRUE);
	PX_Bool(pPX, "AllowColumnMoving", m_bAllowColumnMoving, TRUE);
	PX_Bool(pPX, "SelectByCell", m_bSelectByCell, FALSE);
	PX_String(pPX, "FieldSeparator", m_strFieldSeparator, _T("(tab)"));
	PX_Color(pPX, "BackColor", m_clrBack, GetSysColor(COLOR_WINDOW));
	PX_Color(pPX, "DividerColor", m_clrDivider, RGB(0, 0, 0));
	PX_Long(pPX, "FrozenCols", m_nFrozenCols, 0);
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

		InvalidateControl();
	}
 
	GlobalFree(m_hBlob);
	m_hBlob = NULL;
}


/////////////////////////////////////////////////////////////////////////////
// CBHMDBGridCtrl::OnResetState - Reset control to default state

void CBHMDBGridCtrl::OnResetState()
{
	COleControl::OnResetState();  // Resets defaults found in DoPropExchange

	// TODO: Reset any other control state here.
}


/////////////////////////////////////////////////////////////////////////////
// CBHMDBGridCtrl::AboutBox - Display an "About" box to the user

void CBHMDBGridCtrl::AboutBox()
{
	CDialog dlgAbout(IDD_ABOUTBOX_BHMDBGRID);
	dlgAbout.DoModal();
}


/////////////////////////////////////////////////////////////////////////////
// CBHMDBGridCtrl message handlers
int CBHMDBGridCtrl::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (COleControl::OnCreate(lpCreateStruct) == -1)
		return -1;
	
	CRect rt(0, 0, lpCreateStruct->cx, lpCreateStruct->cy);

	// If has caption, reserve space for it.
	if (InternalGetText().GetLength())
		rt.top += m_nHeaderHeight;

	// Create internal grid here.
	m_pGridInner = new CGridInner(this);
	m_pGridInner->Create(WS_VISIBLE | WS_CHILD, rt, this, 101);

	// Init the grid.
	InitGridHeader();
	
	return 0;
}

long CBHMDBGridCtrl::GetCaptionAlignment() 
/*
Routine Description:
	Get the caption alignment.

Parameters:
	None.

Return Value:
	The caption alignment.
*/
{
	return m_nCapitonAlignment;
}

void CBHMDBGridCtrl::SetCaptionAlignment(long nNewValue) 
/*
Routine Description:
	Set the caption alignment.

Parameters:
	nNewValue		The new value.

Return Value:
	None.
*/
{
	// Save the new value.
	m_nCapitonAlignment = nNewValue;

	// Update the grid.
	SetModifiedFlag();
	BoundPropertyChanged(dispidCaptionAlignment);
	InvalidateControl();
}

long CBHMDBGridCtrl::GetDividerType() 
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

void CBHMDBGridCtrl::SetDividerType(long nNewValue) 
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

BOOL CBHMDBGridCtrl::GetAllowAddNew() 
/*
Routine Description:
	Determines if the grid allow adding new row manually.

Parameters:
	None.

Return Value:
	The permission.
*/
{
	return m_bAllowAddNew;
}

void CBHMDBGridCtrl::SetAllowAddNew(BOOL bNewValue) 
/*
Routine Description:
	Set if the grid allow adding new row manually.

Parameters:
	bNewValue		The new value.

Return Value:
	None.
*/
{
	// If is readonly, then can not add new.
	if (!m_bAllowUpdate && bNewValue)
		return;

	// Save the new value.
	m_bAllowAddNew = bNewValue;

	// Update the grid.
	if (m_pGridInner)
		m_pGridInner->SetAllowAddNew(m_bAllowAddNew);

	InvalidateControl();
	SetModifiedFlag();
	BoundPropertyChanged(dispidAllowAddNew);
}

BOOL CBHMDBGridCtrl::GetAllowDelete() 
/*
Routine Description:
	Determines if the grid allow deleting row manually.

Parameters:
	None.

Return Value:
	The permission.
*/
{
	return m_bAllowDelete;
}

void CBHMDBGridCtrl::SetAllowDelete(BOOL bNewValue) 
/*
Routine Description:
	Set if the grid allow deleting row manually.

Parameters:
	bNewValue		The new value.

Return Value:
	None.
*/
{
	if (!m_bAllowUpdate && bNewValue)
		return;

	m_bAllowDelete = bNewValue;

	SetModifiedFlag();
	BoundPropertyChanged(dispidAllowDelete);
}

BOOL CBHMDBGridCtrl::GetAllowUpdate() 
/*
Routine Description:
	Determines if the grid allow modification manually.

Parameters:
	None.

Return Value:
	The permission.
*/
{
	return m_bAllowUpdate;
}

void CBHMDBGridCtrl::SetAllowUpdate(BOOL bNewValue) 
/*
Routine Description:
	Set if the grid allow modification manually.

Parameters:
	bNewValue		The new value.

Return Value:
	None.
*/
{
	// Save the new value.
	m_bAllowUpdate = bNewValue;

	// Update the grid.
	if (!m_bAllowUpdate)
	{
		SetAllowDelete(FALSE);
		SetAllowAddNew(FALSE);
	}
	if (m_pGridInner && AmbientUserMode())
		m_pGridInner->SetReadOnly(!m_bAllowUpdate);

	SetModifiedFlag();
	BoundPropertyChanged(dispidAllowUpdate);
}

BOOL CBHMDBGridCtrl::GetAllowRowSizing() 
/*
Routine Description:
	Determines if the grid allow resizing row height manually.

Parameters:
	None.

Return Value:
	The permission.
*/
{
	return m_bAllowRowSizing;
}

void CBHMDBGridCtrl::SetAllowRowSizing(BOOL nNewValue) 
/*
Routine Description:
	Set if the grid allow resizing row height manually.

Parameters:
	nNewValue		The new value.

Return Value:
	None.
*/
{
	// Save the new value.
	m_bAllowRowSizing = nNewValue;

	// Update the grid.
	if (AmbientUserMode())
		m_pGridInner->SetAllowRowResize(m_bAllowRowSizing);

	SetModifiedFlag();
	BoundPropertyChanged(dispidAllowRowSizing);
}

BOOL CBHMDBGridCtrl::GetRecordSelectors() 
/*
Routine Description:
	Determines if the grid show row header.

Parameters:
	None.

Return Value:
	The permission.
*/
{
	return m_bRecordSelectors;
}

void CBHMDBGridCtrl::SetRecordSelectors(BOOL bNewValue) 
/*
Routine Description:
	Determines if the grid show row header.

Parameters:
	bNewValue		The new value.

Return Value:
	None.
*/
{
	// Save the new value.
	m_bRecordSelectors = bNewValue;

	// Update the grid.
	if (m_pGridInner)
		m_pGridInner->SetShowRowHeader(m_bRecordSelectors);

	SetModifiedFlag();
	BoundPropertyChanged(dispidRecordSelectors);
	InvalidateControl();
}

BOOL CBHMDBGridCtrl::GetAllowColumnSizing() 
/*
Routine Description:
	Determines if the grid allow resizing col width manually.

Parameters:
	None.

Return Value:
	The permission.
*/
{
	return m_bAllowColumnSizing;
}

void CBHMDBGridCtrl::SetAllowColumnSizing(BOOL bNewValue) 
/*
Routine Description:
	Set if the grid allow adding resizing col width manually.

Parameters:
	bNewValue		The new value.

Return Value:
	None.
*/
{
	// Save the new value.
	m_bAllowColumnSizing = bNewValue;

	// Update the grid.
	if (AmbientUserMode())
		m_pGridInner->SetAllowColResize(m_bAllowColumnSizing);

	SetModifiedFlag();
	BoundPropertyChanged(dispidAllowColumnSizing);
}

BOOL CBHMDBGridCtrl::GetColumnHeader() 
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

void CBHMDBGridCtrl::SetColumnHeader(BOOL bNewValue) 
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

	// Update the grid.
	if (m_pGridInner)
		m_pGridInner->SetShowColHeader(m_bColumnHeader);

	SetModifiedFlag();
	BoundPropertyChanged(dispidColumnHeader);
	InvalidateControl();
}

LPFONTDISP CBHMDBGridCtrl::GetHeadFont() 
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

void CBHMDBGridCtrl::SetHeadFont(LPFONTDISP newValue) 
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
	if (!AmbientUserMode() && m_pGridInner)
		SetHeaderHeight(m_pGridInner->GetRowHeight(1));

	SetModifiedFlag();
	BoundPropertyChanged(dispidHeadFont);
	InvalidateControl();
}

LPFONTDISP CBHMDBGridCtrl::GetFont() 
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

void CBHMDBGridCtrl::SetFont(LPFONTDISP newValue) 
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
	if (!AmbientUserMode() && m_pGridInner)
		SetRowHeight(m_pGridInner->GetRowHeight(0));

	SetModifiedFlag();
	BoundPropertyChanged(dispidFont);
	InvalidateControl();
}

long CBHMDBGridCtrl::GetCol() 
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
	m_pGridInner->GetCurrentCell(nRow, m_nCol);
	return m_nCol;
}

void CBHMDBGridCtrl::SetCol(long nNewValue) 
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
	if (m_pGridInner && m_pGridInner->GetColCount() > (int)m_nCol)
	{
		int nRow, nCol;
		m_pGridInner->GetCurrentCell(nRow, nCol);
		m_pGridInner->SetCurrentCell(nRow, m_nCol);
	}
	else
		ThrowError(CTL_E_INVALIDPROPERTYARRAYINDEX, AFX_IDP_E_INVALIDPROPERTYARRAYINDEX);

	SetModifiedFlag();
	BoundPropertyChanged(dispidCol);
}

long CBHMDBGridCtrl::GetRow() 
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
	m_pGridInner->GetCurrentCell(m_nRow, nCol);
	return m_nRow;
}

void CBHMDBGridCtrl::SetRow(long nNewValue) 
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
	if (m_pGridInner && m_pGridInner->GetRowCount() > (int)m_nRow)
	// Update the grid.
		m_pGridInner->MoveTo(m_nRow);
	else
		ThrowError(CTL_E_INVALIDPROPERTYARRAYINDEX, AFX_IDP_E_INVALIDPROPERTYARRAYINDEX);

	SetModifiedFlag();
	BoundPropertyChanged(dispidRow);
}

long CBHMDBGridCtrl::GetDataMode() 
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

void CBHMDBGridCtrl::SetDataMode(long nNewValue) 
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

long CBHMDBGridCtrl::GetVisibleCols() 
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

	return m_pGridInner->GetVisibleCols();
}

long CBHMDBGridCtrl::GetVisibleRows() 
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

	return m_pGridInner->GetVisibleRows();
}

long CBHMDBGridCtrl::GetLeftCol() 
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

	return m_pGridInner->GetLeftCol();
}

void CBHMDBGridCtrl::SetLeftCol(long nNewValue) 
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
	m_pGridInner->SetLeftCol(nNewValue);
}

long CBHMDBGridCtrl::GetCols() 
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
	if (!m_pGridInner)
		return m_nCols;

	// Update the cached value.
	m_nCols = m_pGridInner->GetColCount();

	return m_nCols;
}

void CBHMDBGridCtrl::SetCols(long nNewValue) 
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
	if (m_pGridInner)
		m_pGridInner->SetColCount(m_nCols);

	SetModifiedFlag();
	BoundPropertyChanged(dispidCols);
	InvalidateControl();
}

long CBHMDBGridCtrl::GetRows() 
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
	if (!m_pGridInner)
		return m_nRows;

	m_nRows = m_pGridInner->GetRowCount();

	return m_nRows;
}

void CBHMDBGridCtrl::SetRows(long nNewValue) 
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

	if (nNewValue < 0)
		return;

	// Rows can not less than 0.
	if (m_nDataMode == 0)
	// Can not set row count in bound mode.
		ThrowError(CTL_E_SETNOTSUPPORTED, IDS_ERROR_NOSETROWSINBINDMODE);
	if (AmbientUserMode())
	// Can not set row count at run time.
		SetNotSupported();

	// Save the new value.
	m_nRows = nNewValue;

	m_pGridInner->SetRowCount(m_nRows);

	SetModifiedFlag();
	BoundPropertyChanged(dispidRows);
	InvalidateControl();
}

BSTR CBHMDBGridCtrl::GetFieldSeparator() 
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

void CBHMDBGridCtrl::SetFieldSeparator(LPCTSTR lpszNewValue) 
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

long CBHMDBGridCtrl::GetDividerStyle() 
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

void CBHMDBGridCtrl::SetDividerStyle(long nNewValue) 
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

long CBHMDBGridCtrl::GetFirstRow() 
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

	return m_pGridInner->GetTopRow();
}

void CBHMDBGridCtrl::SetFirstRow(long nNewValue) 
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

	m_pGridInner->SetTopRow(nNewValue);
}

long CBHMDBGridCtrl::GetDefColWidth() 
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

void CBHMDBGridCtrl::SetDefColWidth(long nNewValue) 
/*
Routine Description:
	Set the default column width.

Parameters:
	nNewValue		The new value.

Return Value:
	None.
*/
{
	if (nNewValue < 0 || nNewValue > 15000)
		ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDWIDTHORHEIGHT);

	// Save the new value.
	m_nDefColWidth = nNewValue;

	// Update the grid.
	if (m_pGridInner)
	{
		m_pGridInner->SetDefColWidth(m_nDefColWidth);
	}

	SetModifiedFlag();
	BoundPropertyChanged(dispidDefColWidth);
	InvalidateControl();
}

long CBHMDBGridCtrl::GetRowHeight() 
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
	if (m_pGridInner)
		m_nRowHeight = m_pGridInner->GetRowHeight(1);

	return m_nRowHeight;
}

void CBHMDBGridCtrl::SetRowHeight(long nNewValue) 
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
	if (nNewValue < 0 || nNewValue > 15000)
		ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDWIDTHORHEIGHT);

	// Save the new value.
	m_nRowHeight = nNewValue;

	// Update the grid.
	if (m_pGridInner)
	{
		m_pGridInner->SetRowHeight(m_nRowHeight);
	}

	SetModifiedFlag();
	BoundPropertyChanged(dispidRowHeight);
	InvalidateControl();
}

BOOL CBHMDBGridCtrl::GetRedraw() 
/*
Routine Description:
	Determines if the grid can redraw.

Parameters:
	None.

Return Value:
	The permission.
*/
{
	if (!AmbientUserMode())
		GetNotSupported();

	return m_bRedraw;
}

void CBHMDBGridCtrl::SetRedraw(BOOL bNewValue) 
/*
Routine Description:
	Set if the grid can redraw.

Parameters:
	bNewValue		The new value.

Return Value:
	None.
*/
{
	// This property is unavailable at design time.
	if (!AmbientUserMode())
		SetNotSupported();

	// Save the new value.
	m_bRedraw = bNewValue;

	// Update the grid.
	m_pGridInner->LockUpdate(!m_bRedraw);
}

BOOL CBHMDBGridCtrl::GetAllowColumnMoving() 
/*
Routine Description:
	Determines if the user can move the column position.

Parameters:
	None.

Return Value:
	The permission.
*/
{
	return m_bAllowColumnMoving;
}

void CBHMDBGridCtrl::SetAllowColumnMoving(BOOL bNewValue) 
/*
Routine Description:
	Set if the user can move the column position.

Parameters:
	bNewValue		The new value.

Return Value:
	None.
*/
{
	// Save the new value.
	m_bAllowColumnMoving = bNewValue;

	// Update the grid.
	if (AmbientUserMode())
		m_pGridInner->SetAllowMoveCol(m_bAllowColumnMoving);

	SetModifiedFlag();
	BoundPropertyChanged(dispidAllowColumnMoving);
}

BOOL CBHMDBGridCtrl::GetSelectByCell() 
/*
Routine Description:
	Determines if the grid acts like a list box.

Parameters:
	None.

Return Value:
	The permission.
*/
{
	return m_bSelectByCell;
}

void CBHMDBGridCtrl::SetSelectByCell(BOOL bNewValue) 
/*
Routine Description:
	Determines if the grid acts like a list box.

Parameters:
	bNewValue		The new value.

Return Value:
	None.
*/
{
	// Save the new value.
	m_bSelectByCell = bNewValue;

	// Update the grid at run time.
	if (AmbientUserMode())
	{
		if (m_bSelectByCell)
		{
			m_pGridInner->SetListMode(TRUE);
			SetAllowColumnMoving(FALSE);
			SetAllowColumnSizing(FALSE);
			SetAllowRowSizing(FALSE);
		}
		else
		{
			m_pGridInner->SetListMode(FALSE);
		}
	}
	
	SetModifiedFlag();
	BoundPropertyChanged(dispidSelectByCell);
}

LPDISPATCH CBHMDBGridCtrl::GetColumns() 
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
	CColumns * pColumns = new CColumns();

	// Set the callback pointers.
	pColumns->SetCtrlPtr(this);
	pColumns->m_pGrid = m_pGridInner;

	return pColumns->GetIDispatch(FALSE);
}

LPUNKNOWN CBHMDBGridCtrl::GetDataSource() 
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

void CBHMDBGridCtrl::SetDataSource(LPUNKNOWN newValue) 
/*
Routine Description:
	Set the data source object.

Parameters:
	newValue		The new pointer.

Return Value:
	None.
*/
{
	// Can not set datasource in manual mode.
	if (m_nDataMode != 0 && newValue)
	{
		ThrowError(CTL_E_SETNOTSUPPORTED, _T("Can not set datasource in manual mode"));
		goto exit;
	}

	if (newValue == m_spDataSource)
		return;

	// Clear the cached interface pointers.
	ClearInterfaces();
	if (m_spDataSource)
	{
		m_spDataSource->Release();
		m_spDataSource = NULL;
	}

	if (AmbientUserMode())
	{
		// Datasource has been changed.
		m_bDataSrcChanged = TRUE;
		m_bEndReached = FALSE;
	}

	if (!newValue)
	{
		m_spDataSource = NULL;
		goto exit;
	}

	// Get IDataSource pointer.
	if (FAILED(newValue->QueryInterface(IID_IDataSource, (void **)&m_spDataSource)))
		goto exit;

	// Retrieve the fields information.
	if (!AmbientUserMode())
		GetColumnInfo();

	SetModifiedFlag();

exit:
	UpdateControl();
}

BSTR CBHMDBGridCtrl::GetDataMember() 
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

void CBHMDBGridCtrl::SetDataMember(LPCTSTR lpszNewValue) 
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

	InvalidateControl();
}

void CBHMDBGridCtrl::AddItem(LPCTSTR Item, const VARIANT FAR& RowIndex) 
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
		nRow = m_pGridInner->OnGetRecordCount() + 1;

	if (nRow == 0 || nRow > (int)m_pGridInner->GetRecordCount() + 1)
		ThrowError(CTL_E_INVALIDPROPERTYARRAYINDEX, AFX_IDP_E_INVALIDPROPERTYARRAYINDEX);

	CString strText = Item, strField;
	int i = 1, j = 0;

	// Insert a row.
	m_pGridInner->InsertRow(nRow);
	strField.Empty();

	while (j < strText.GetLength() && (int)i <= m_pGridInner->GetColCount())
	{
		// Retrieve the value of each cell.

		if (strText[j] != m_strSeparator)
			strField += strText[j];
		else
		{
			// Fill this cell.
			m_pGridInner->SetCellText(nRow, i, strField);
			strField.Empty();
			i ++;
		}

		j ++;
	}

	// Fill the last cell.
	if ((int)i <= m_pGridInner->GetColCount())
		m_pGridInner->SetCellText(nRow, i, strField);

	m_pGridInner->RedrawRow(nRow);
}

void CBHMDBGridCtrl::RemoveItem(long RowIndex) 
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
	if (nRow == 0 || nRow > m_pGridInner->GetRowCount())
		ThrowError(CTL_E_INVALIDPROPERTYARRAYINDEX, AFX_IDP_E_INVALIDPROPERTYARRAYINDEX);

	m_pGridInner->RemoveRow(nRow);

	// Update the grid.
	int nCol;
	m_pGridInner->GetCurrentCell(nRow, nCol);
	if (nRow > m_pGridInner->GetRecordCount())
		nRow = m_pGridInner->GetRecordCount();
	if (nCol > m_pGridInner->GetColCount())
		nCol = m_pGridInner->GetColCount();
	m_pGridInner->MoveTo(nRow);
}

void CBHMDBGridCtrl::Scroll(long Rows, long Cols) 
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
			m_pGridInner->SendMessage(WM_VSCROLL, SB_LINEDOWN);
	}
	else
	{
		for (int i = 0; i > Rows; i --)
			m_pGridInner->SendMessage(WM_VSCROLL, SB_LINEUP);
	}
	if (Cols > 0)
	{
		for (int i = 0; i < Cols; i ++)
			m_pGridInner->SendMessage(WM_HSCROLL, SB_LINERIGHT);
	}
	else
	{
		for (int i = 0; i > Cols; i --)
			m_pGridInner->SendMessage(WM_HSCROLL, SB_LINELEFT);
	}
}

void CBHMDBGridCtrl::RemoveAll() 
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
	m_pGridInner->SetRowCount(0);
}

void CBHMDBGridCtrl::MoveFirst() 
/*
Routine Description:
	Set the first row as current row.

Parameters:
	None.

Return Value:
	None.
*/
{
	// Update the grid.
	m_pGridInner->MoveTo(1);
}

void CBHMDBGridCtrl::MovePrevious() 
/*
Routine Description:
	Set the prior row as the current row.

Parameters:
	None.

Return Value:
	None.
*/
{
	// Update the grid.
	int nRow, nCol;
	m_pGridInner->GetCurrentCell(nRow, nCol);
	m_pGridInner->MoveTo(nRow - 1);
}

void CBHMDBGridCtrl::MoveNext() 
/*
Routine Description:
	Set the next row as current row.

Parameters:
	None.

Return Value:
	None.
*/
{
	// Update the grid.
	int nRow, nCol;
	m_pGridInner->GetCurrentCell(nRow, nCol);
	m_pGridInner->MoveTo(nRow + 1);
}

void CBHMDBGridCtrl::MoveLast() 
/*
Routine Description:
	Set the last row as current row.

Parameters:
	None.

Return Value:
	None.
*/
{
	// Update the grid.
	m_pGridInner->MoveTo(m_pGridInner->OnGetRecordCount());
}

void CBHMDBGridCtrl::MoveRecords(long Rows) 
/*
Routine Description:
	Move current row by given rows.

Parameters:
	Rows		Rows to move.

Return Value:
	None.
*/
{
	// Update the grid.
	int nRow, nCol;
	m_pGridInner->GetCurrentCell(nRow, nCol);
	m_pGridInner->MoveTo(nRow + Rows);
}

BOOL CBHMDBGridCtrl::IsAddRow() 
/*
Routine Description:
	Determines if current row is the pending new row.

Parameters:
	None.

Return Value:
	If is, return TRUE; Otherwise, return FALSE.
*/
{
	return m_pGridInner->IsAddRow();
}

void CBHMDBGridCtrl::Update() 
/*
Routine Description:
	Write all pending modification back into data source.

Parameters:
	None.

Return Value:
	None.
*/
{
	m_pGridInner->FlushRecord();
}

void CBHMDBGridCtrl::CancelUpdate() 
/*
Routine Description:
	Cancel all pending modification.

Parameters:
	Rows		Rows to move.

Return Value:
	None.
*/
{
	m_pGridInner->CancelRecord();
}

void CBHMDBGridCtrl::ReBind() 
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

	InvalidateControl();
}

BOOL CBHMDBGridCtrl::GetRowChanged() 
/*
Routine Description:
	Determines if current row is dirty.

Parameters:
	None.

Return Value:
	If is, return TRUE; Otherwise, return FALSE.
*/
{
	if (!AmbientUserMode())
		GetNotSupported();

	return m_pGridInner->IsRecordDirty();
}

void CBHMDBGridCtrl::SetRowChanged(BOOL bNewValue) 
/*
Routine Description:
	Set if current row is dirty, only FALSE value is valid.

Parameters:
	bNewValue		The new value.

Return Value:
	None.
*/
{
	if (!AmbientUserMode())
		SetNotSupported();

	if (m_pGridInner->IsRecordDirty() && !bNewValue)
		m_pGridInner->CancelRecord();
}

BOOL CBHMDBGridCtrl::OnSetObjectRects(LPCRECT lpRectPos, LPCRECT lpRectClip) 
{
	// Reserve space for capiton.
	CRect rt;
	rt.CopyRect(lpRectPos);
	rt.OffsetRect(-rt.left, -rt.top);
	rt.top += InternalGetText().GetLength() ? (m_pGridInner ? m_pGridInner->GetLevelHeight() : m_nHeaderHeight) : 0;

	if (m_pGridInner)
	{
		m_pGridInner->SetWindowPos(NULL, rt.left, rt.top, rt.Width(), rt.Height(), SWP_NOZORDER);
	}
	
	return COleControl::OnSetObjectRects(lpRectPos, lpRectClip);
}

void CBHMDBGridCtrl::InitHeaderFont()
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
		if (GetObject(hFont, sizeof(logHeader), &logHeader) && m_pGridInner)
		{
			// Update the grid.
			m_pGridInner->SetHeaderFont(&logHeader);
		}
	}
}

void CBHMDBGridCtrl::InitCellFont()
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
		if (GetObject(hFont, sizeof(logCell), &logCell) && m_pGridInner)
		{
			// Update the grid.
			m_pGridInner->SetCellFont(&logCell);
		}
	}
}

void CBHMDBGridCtrl::InitGridHeader()
/*
Routine Description:
	Init the grid from cached prop data.

Parameters:
	None.

Return Value:
	None.
*/
{
	// Update field separator.
	SetFieldSeparator(m_strFieldSeparator);

	// Update divider lines.
	UpdateDivider();

	m_pGridInner->SetAllowColResize(m_bAllowColumnSizing);
	m_pGridInner->SetAllowRowResize(m_bAllowRowSizing);
	m_pGridInner->SetFrozenCols(m_nFrozenCols);
	
	if(AmbientUserMode())
	{
		if (m_bSelectByCell)
		{
			m_pGridInner->SetListMode(TRUE);
			SetAllowColumnMoving(FALSE);
			SetAllowColumnSizing(FALSE);
			SetAllowRowSizing(FALSE);
		}
		else
		{
			m_pGridInner->SetListMode(FALSE);
		}
	}

	// Update colors.	
	m_pGridInner->SetForeColor(TranslateColor(GetForeColor()));
	m_pGridInner->SetBackColor(TranslateColor(GetBackColor()));
	m_pGridInner->SetHeaderForeColor(TranslateColor(m_clrHeadFore));
	m_pGridInner->SetHeaderBackColor(TranslateColor(m_clrHeadBack));

	// Update fonts.
	InitHeaderFont();
	InitCellFont();
	if (!AmbientUserMode())
	{
		m_pGridInner->SetReadOnly(FALSE);
		m_pGridInner->ShowWindow(SW_HIDE);
		m_pGridInner->LockUpdate();
	}
	
	// Update heights.
	m_pGridInner->SetHeaderHeight(m_nHeaderHeight);
	m_pGridInner->SetRowHeight(m_nRowHeight);
	m_pGridInner->SetDefColWidth(m_nDefColWidth);
	m_pGridInner->SetShowGroupHeader(m_bGroupHeader);

	m_pGridInner->SetAllowAddNew(m_bAllowAddNew);
	InitGridFromProp();

	if (AmbientUserMode())
		m_pGridInner->SetReadOnly(!m_bAllowUpdate);

	m_pGridInner->SetShowColHeader(m_bColumnHeader);
	m_pGridInner->SetShowRowHeader(m_bRecordSelectors);

	InvalidateControl();
}

void CBHMDBGridCtrl::OnForeColorChanged() 
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
	if (m_pGridInner)
		m_pGridInner->SetForeColor(TranslateColor(GetForeColor()));
	
	COleControl::OnForeColorChanged();
	BoundPropertyChanged(DISPID_FORECOLOR);
}

void CBHMDBGridCtrl::UpdateControl()
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

	m_pGridInner->LockUpdate(TRUE);

	// We processed the change of data source.
	m_bDataSrcChanged = FALSE;

	// Retrieve field information.
	if (!GetColumnInfo())
	{
		m_pGridInner->LockUpdate(FALSE);
		m_pGridInner->Invalidate();

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

	m_pGridInner->LockUpdate(TRUE);
	InitGridHeader();

	// Do the binding job.
	Bind();

	m_pGridInner->LockUpdate(FALSE);

	// Now fetch new rows according new binding information.
	m_Rowset.SetupOptionalRowsetInterfaces();
	FetchNewRows();
	
	// Set up the sink so we know when the rowset is repositioned
	if (AmbientUserMode())
	{
		AfxConnectionAdvise(m_spRowPosition, IID_IRowPositionChange, &m_xRowPositionChange, FALSE, &m_dwCookieRPC);
		AfxConnectionAdvise(m_Rowset.m_spRowset, IID_IRowsetNotify, &m_xRowsetNotify, FALSE, &m_dwCookieRN);
		m_spDataSource->addDataSourceListener((IDataSourceListener*)&m_xDataSourceListener);
	}

	m_pGridInner->LockUpdate(FALSE);
	m_pGridInner->Invalidate();
}

long CBHMDBGridCtrl::GetHeaderHeight()
/*
Routine Description:
	Get the height of header row.

Parameters:
	None.

Return Value:
	The height of header row.
*/
{
	if (m_pGridInner)
		m_nHeaderHeight = m_pGridInner->GetRowHeight(0);

	return m_nHeaderHeight;
}

void CBHMDBGridCtrl::SetHeaderHeight(long nNewValue) 
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
	if (nNewValue < 0 || nNewValue > 15000)
		ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDWIDTHORHEIGHT);

	// Save the new value.
	m_nHeaderHeight = nNewValue;
	CRect rt;

	// Be careful in null pointer.
	if (m_pGridInner)
	{
		m_pGridInner->SetHeaderHeight(m_nHeaderHeight);
		
		int nHeaderHeight = m_pGridInner->GetLevelHeight();
		GetClientRect(&rt);
		rt.top += InternalGetText().GetLength() ? nHeaderHeight : 0;
		m_pGridInner->SetWindowPos(NULL, rt.left, rt.top, rt.Width(), rt.Height(), SWP_NOZORDER);
	}

	SetModifiedFlag();
	BoundPropertyChanged(dispidHeaderHeight);
	InvalidateControl();
}

void CBHMDBGridCtrl::OnHeadForeColorChanged() 
/*
Routine Description:
	Update the control after the "HeadForeColor" property changed.

Parameters:
	None.

Return Value:
	None.
*/
{
	// Update the grid.
	if (m_pGridInner)
	{
		m_pGridInner->SetHeaderForeColor(TranslateColor(m_clrHeadFore));
	}

	SetModifiedFlag();
	BoundPropertyChanged(dispidHeadForeColor);
	InvalidateControl();
}

void CBHMDBGridCtrl::OnHeadBackColorChanged() 
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
	if (m_pGridInner)
	{
		m_pGridInner->SetHeaderBackColor(TranslateColor(m_clrHeadBack));
	}

	SetModifiedFlag();
	BoundPropertyChanged(dispidHeadBackColor);
	InvalidateControl();
}

void CBHMDBGridCtrl::UpdateDivider()
/*
Routine Description:
	Update the divider lines.

Parameters:
	None.

Return Value:
	None.
*/
{
	if (m_pGridInner)
	{
		m_pGridInner->SetDividerType(m_nDividerType);
		m_pGridInner->SetDividerStyle(m_nDividerStyle);
		m_pGridInner->SetDividerColor(TranslateColor(m_clrDivider));
	}
}

int CBHMDBGridCtrl::GetRowColFromVariant(VARIANT va)
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

BOOL CBHMDBGridCtrl::OnGetPredefinedStrings(DISPID dispid, CStringArray* pStringArray, CDWordArray* pCookieArray) 
{
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
	}
	
	return COleControl::OnGetPredefinedStrings(dispid, pStringArray, pCookieArray);
}

BOOL CBHMDBGridCtrl::OnGetPredefinedValue(DISPID dispid, DWORD dwCookie, VARIANT* lpvarOut) 
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
	}
	
	return COleControl::OnGetPredefinedValue(dispid, dwCookie, lpvarOut);
}

void CBHMDBGridCtrl::InitGridFromProp()
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
	BOOL bLockUpdate = m_pGridInner->IsLockUpdate();

	// Prevent the grid from update.
	m_pGridInner->LockUpdate();

	// Clear the grid.
	m_pGridInner->SetColCount(0);
	m_pGridInner->SetRowCount(0);
	m_pGridInner->SetGroupCount(0);
	
	// Load the layout from cache.
	m_pGridInner->LoadLayout(&m_arGroups, &m_arCols, AmbientUserMode() ? &m_arCells : NULL);

	// Recalc binding info of each column.
	for (int i = 0; m_nDataMode == 0 && i < m_pGridInner->GetColCount(); i++)
	{
		CString strField = m_pGridInner->GetColUserAttribRef(i + 1).ElementAt(0);

		int nDataField = GetFieldFromStr(strField);
		BOOL bForceNoNullable = m_pColumnInfo && !(m_pColumnInfo[nDataField].dwFlags & DBCOLUMNFLAGS_ISNULLABLE);
		BOOL bForceLock = m_pColumnInfo && !(m_pColumnInfo[nDataField].dwFlags & DBCOLUMNFLAGS_WRITE);

		if (bForceLock)
			m_pGridInner->SetColReadOnly(i + 1, TRUE);
		if (bForceNoNullable)
			m_pGridInner->GetColUserAttribRef(i + 1).ElementAt(1) = _T("T");
	}

	// If we have no predefined layout and the data mode is manual mode, insert some blank 
	// rows and cols.
	if (!AmbientUserMode())
	{
		m_pGridInner->SetRowCount(m_nRows);

		if (!m_pGridInner->GetGroupCount() && !m_pGridInner->GetColCount())
			m_pGridInner->SetColCount(m_nCols);
	}	
	else if (!m_pGridInner->GetGroupCount() && !m_pGridInner->GetColCount() && m_nDataMode)
	{
		m_pGridInner->SetRowCount(m_nRows);
		m_pGridInner->SetColCount(m_nCols);
	}

	m_pGridInner->LockUpdate(bLockUpdate);
	InvalidateControl();

	// Update the variables.
	m_nRows = m_pGridInner->GetRowCount();
	m_nCols = m_pGridInner->GetColCount();

	BoundPropertyChanged(dispidRows);
	BoundPropertyChanged(dispidCols);

	SetModifiedFlag();
}

void CBHMDBGridCtrl::SerializeColumnProp(CArchive &ar, BOOL bLoad)
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

		// Save the column count first.
		int nCols = m_arCols.GetSize();
		Col * pCol;

		ar << nCols;

		for (i = 0; i < nCols; i ++)
		{
			// Save each column props.

			pCol = m_arCols[i];

			ar << pCol->bAllowSizing;
			ar << pCol->bReadOnly << pCol->bVisible;
			ar << pCol->clrBack << pCol->clrFore << pCol->clrHeaderBack << pCol->clrHeaderFore;
			ar << pCol->nAlignment << pCol->nHeaderAlignment << pCol->nCase;
			ar << pCol->nDataType << pCol->nMaxLength << pCol->nStyle;
			ar << pCol->nWidth << pCol->strTitle << pCol->strChoiceList;
			ar << pCol->arUserAttrib[1] << pCol->strMask << pCol->arUserAttrib[0];
			ar << pCol->strPromptChar << pCol->bPromptInclude << pCol->strName;
		}

		// Save cell data.

		// Save cell count first.		
		int nCells = m_arCells.GetSize();
		
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
			pCol->arUserAttrib.SetSize(2);

			ar >> pCol->bAllowSizing;
			ar >> pCol->bReadOnly >> pCol->bVisible;
			ar >> pCol->clrBack >> pCol->clrFore >> pCol->clrHeaderBack >> pCol->clrHeaderFore;
			ar >> pCol->nAlignment >> pCol->nHeaderAlignment >> pCol->nCase;
			ar >> pCol->nDataType >> pCol->nMaxLength >> pCol->nStyle;
			ar >> pCol->nWidth >> pCol->strTitle >> pCol->strChoiceList;
			ar >> pCol->arUserAttrib[1] >> pCol->strMask >> pCol->arUserAttrib[0];
			ar >> pCol->strPromptChar >> pCol->bPromptInclude >> pCol->strName;

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
			pCell = new Cell;
			ar >> pCell->strText;

			// Add the new cell object into cached data.
			m_arCells.Add(pCell);
			nCells --;
		}
	}
}

BOOL CBHMDBGridCtrl::PreTranslateMessage(MSG* pMsg) 
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
			if (m_hWndDropDown && m_bDroppedDown)
			{
				// The dropdown window is shown, send the keyboard message to it.
				::SendMessage((HWND)m_hWndDropDown, WM_KEYDOWN, pMsg->wParam, pMsg->lParam);
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
	}
	
	return COleControl::PreTranslateMessage(pMsg);
}

void CBHMDBGridCtrl::ButtonDown(UINT nFlags, CPoint point)
{
	FireMouseDown(nFlags, _AfxShiftState(), point.x, point.y);
}

void CBHMDBGridCtrl::ButtonUp(UINT nFlags, CPoint point)
{
	FireMouseUp(nFlags, _AfxShiftState(), point.x, point.y);
	
	CRect rect;
	GetClientRect(&rect);
	if (rect.PtInRect(point))
		FireClick();
}

void CBHMDBGridCtrl::ButtonDblClick(UINT nFlags, CPoint point)
{
	FireDblClick();
}

HWND CBHMDBGridCtrl::GetApplicationWindow()
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

void CBHMDBGridCtrl::CreateDrawGrid()
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
	if (m_pGridInner)
		return;

	// Get the parent window for internal grid.
	HWND hWnd = GetApplicationWindow(); 
	
	// Create the internal grid.
	m_pGridInner = new CGridInner(this);
	m_pGridInner->Create(WS_VISIBLE | WS_CHILD, CRect(100,100, 110, 110), CWnd::FromHandle(hWnd), 101);
//	InitGridHeader();
}

int CBHMDBGridCtrl::GetFieldFromStr(CString str)
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

int CBHMDBGridCtrl::GetRowFromBmk(BYTE *bmk)
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
			m_bIsPosSync = TRUE;
			delete [] bmkSearch;
			return i + 1;
		}
	}
	
	// The given bookmark does not exist.
	delete [] bmkSearch;
	m_bIsPosSync = FALSE;
	return INVALID;
}

HRESULT CBHMDBGridCtrl::GetRowset()
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

void CBHMDBGridCtrl::GetCellValue(int nRow, int nCol, CString & strText)
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
	if (!m_bEndReached && nRow == (int)m_pGridInner->GetRecordCount())
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

	// Move to this row.
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
		nColToSet = m_pGridInner->GetColFromIndex(pCol->nColIndex);
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

void CBHMDBGridCtrl::SetCurrentCell(int nRow, int nCol)
/*
Routine Description:
	Sync the current row in data source when the user navigate to other row.

Parameters:
	nRow		The row ordinal.

	nCol		The col ordinal.

Return Value:
	None.
*/
{
	// Do nothing in manual mode.
	if (m_nDataMode != 0)
		return;

	// The row ordinal should be valid.	
	if ((int)nRow > m_apBookmark.GetSize())
		return;

	// Get the bookmark of this row.
	CBookmark<0> bmk;
	bmk.SetBookmark(m_nBookmarkSize, m_apBookmark[nRow - 1]);
	
	HRESULT hr = m_Rowset.MoveToBookmark(bmk);
	if (FAILED(hr))
		return;

	// Move the data source to new bookmark.	
	HROW hRowNow = 0;
	DBPOSITIONFLAGS dwPositionFlags;
	hr = m_spRowPosition->GetRowPosition(NULL, &hRowNow, &dwPositionFlags);
	if (FAILED(hr) || hRowNow == DB_NULL_HROW || hRowNow == m_Rowset.m_hRow)
		return;

	m_spRowPosition->ClearRowPosition();
	m_spRowPosition->SetRowPosition(NULL, m_Rowset.m_hRow, DBPOSITION_OK);
	m_Rowset.FreeRecordMemory();
}

int CBHMDBGridCtrl::GetRowFromHRow(HROW hRow)
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

void CBHMDBGridCtrl::DeleteRowFromSrc(int nRow)
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
	if (m_pGridInner->IsAddRow())
		m_pGridInner->CancelRecord();
	else
	{
		m_pGridInner->CancelRecord();
		m_pGridInner->RemoveRow(nRow);

		// Remove bookmark from cached data.
		delete [] m_apBookmark[nRow - 1];
		m_apBookmark.RemoveAt(nRow - 1);
	}
}

void CBHMDBGridCtrl::ClearInterfaces()
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
	if (m_pGridInner && AmbientUserMode())
	{
		m_pGridInner->SetColCount(0);
		m_pGridInner->SetRowCount(0);
	}

	if (m_dwCookieRPC != 0)
	{
		AfxConnectionUnadvise(m_spRowPosition, IID_IRowPositionChange, &m_xRowPositionChange, FALSE, m_dwCookieRPC);
		m_dwCookieRPC = 0;
	}

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

void CBHMDBGridCtrl::FreeColumnMemory()
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

void CBHMDBGridCtrl::FetchNewRows()
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
	m_pGridInner->CancelRecord();

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

	// Enlarge the bookmark array first, fill the new entries with null.
	for (i = 0; i < nFetchOnce; i ++)
		m_apBookmark.Add(NULL);

	// Calc the rows need to fetch.
	nFetchOnce = nRowsAvailable - m_pGridInner->GetRecordCount();
	if (nFetchOnce <= 0)
		return;
	
	// Simply insert new blank rows into grid.	
	m_pGridInner->SetRowCount(m_pGridInner->GetRowCount() + nFetchOnce);

	m_pGridInner->Invalidate();
	m_pGridInner->ResetScrollBars();
}

BOOL CBHMDBGridCtrl::GetColumnInfo()
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

	return m_apColumnData.GetSize() > 0;
}

HRESULT CBHMDBGridCtrl::UpdateRowData(int nRow)
/*
Routine Description:
	Write the modification back into datat source.

Parameters:
	The row to be written back.

Return Value:
	Standard HResult.
*/
{
	// We can not do anything without rowset object.
	if (m_Rowset.m_spRowset == NULL)
		return E_FAIL;

	// Do we should insert new row?
	BOOL bInsert = m_pGridInner->IsAddRow();

	// Fire events.
	short bCancel = 0;
	short bDispMsg = 0;

	if (!bInsert)
		FireBeforeUpdate(&bCancel);
	if (bCancel)
		return E_FAIL;

	if (nRow > (int)m_apBookmark.GetSize() + bInsert ? 1: 0)
		return E_FAIL;

	// Create accessor.
	CAccessorRowset<CManualAccessor> rac;
	RowData bindData;
	int i;

	// Init bind entries.
	for (i = 0; i < 255; i ++)
		bindData.dwStatus[i] = DBSTATUS_S_OK;

	HRESULT hr;
	CString str;

	// Init the accessor.
	rac.m_spRowset = m_Rowset.m_spRowset;

	// Is this column writable?
	int nColumnsWritable = 0;

	int nColOrdinal;
	ColumnProp * pCol;

	for (i = 1; i <= m_pGridInner->GetColCount(); i ++)
	{
		// Skip the cols which are not dirty.
		if (!m_pGridInner->m_bColDirty[i - 1] && !bInsert)
			continue;

		// Skip the cols which are not bouned to a field.		
		pCol = m_arBindInfo[GetColOrdinalFromCol(i)];
		if (pCol->nDataField == -1)
			continue;

		// Add this field into binding entry.
		if (m_pColumnInfo[pCol->nDataField].dwFlags & DBCOLUMNFLAGS_WRITE)
			nColumnsWritable ++;
	}

	// If no field need update, return.
	if (nColumnsWritable == 0)
		return S_OK;

	// Bind the entries with accessor.	
	rac.CreateAccessor(nColumnsWritable, &bindData, sizeof(bindData));
	for (i = 1; i <= m_pGridInner->GetColCount(); i ++)
	{
		// Add each col needed into binding entries.
		if (!m_pGridInner->m_bColDirty[i - 1] && !bInsert)
			continue;
		
		nColOrdinal = GetColOrdinalFromCol(i);
		pCol = m_arBindInfo[nColOrdinal];
		if (pCol->nDataField == -1)
			continue;

		if (m_pColumnInfo[pCol->nDataField].dwFlags & DBCOLUMNFLAGS_WRITE)
			rac.AddBindEntry(pCol->nDataField, DBTYPE_STR, 255 * sizeof(TCHAR), 
				&(bindData.strData[nColOrdinal]), NULL, &(bindData.dwStatus[nColOrdinal]));
	}

	// Bind.
	hr = rac.Bind();
	if (FAILED(hr))
	{
		ReportError();
		return E_FAIL;
	}

	if (!bInsert)
	{
		// Update existing record.

		// Move to this record.
		CBookmark<0> bmk;
		bmk.SetBookmark(m_nBookmarkSize, m_apBookmark[nRow - 1]);
		hr = rac.MoveToBookmark(bmk);
		if (FAILED(hr))
		{
			// Fire the event.
			FireUpdateError();
			
			return E_FAIL;
		}

		rac.FreeRecordMemory();
	}

	rac.SetupOptionalRowsetInterfaces();
	
	// Fill the data to be written back.
	for (i = 1; i <= m_pGridInner->GetColCount(); i ++)
	{
		if (!m_pGridInner->m_bColDirty[i - 1] && !bInsert)
			continue;
		
		nColOrdinal = GetColOrdinalFromCol(i);
		pCol = m_arBindInfo[nColOrdinal];
		if (pCol->nDataField == -1)
			continue;
		
		if (m_pColumnInfo[pCol->nDataField].dwFlags & DBCOLUMNFLAGS_WRITE)
		{
			str = m_pGridInner->m_arContent[i - 1];

			// Can we use null value?			
			int nCol = m_pGridInner->GetColFromIndex(pCol->nColIndex);
			BOOL bNullable = !m_pGridInner->GetColUserAttribRef(nCol).ElementAt(1).Compare(_T("T"));
			if (str.IsEmpty() && bNullable)
			{
				// Use null value.
				bindData.dwStatus[nColOrdinal] = DBSTATUS_S_ISNULL;
			}
			else
			{
				lstrcpy(bindData.strData[nColOrdinal], str);
				bindData.dwStatus[nColOrdinal] = DBSTATUS_S_OK;
			}
		}
	}
	
	if (!bInsert)
	{
		// Update existing record.
		try
		{
			hr = rac.SetData();
		}
		catch (CException * e)
		{
			e->ReportError();
			e->Delete();
			hr = E_FAIL;
		}
	}
	else
	{
		// Insert new record.
		try
		{
			hr = rac.Insert(0, true);
		}
		catch (CException * e)
		{
			e->ReportError();
			e->Delete();
			hr = E_FAIL;
		}
	}

	if (FAILED(hr))
	{
		// Report the error.
		ReportError();
		
		for (i = 1; i <= m_pGridInner->GetColCount(); i ++)
		{
			nColOrdinal = GetColOrdinalFromCol(i);
			pCol = m_arBindInfo[nColOrdinal];
			
			if (pCol->nDataField == -1)
				continue;

			// Locate the col which causes failure.			
			if (bindData.dwStatus[nColOrdinal] != DBSTATUS_S_OK && bindData.dwStatus[nColOrdinal]
				!= DBSTATUS_S_ISNULL && bindData.dwStatus[nColOrdinal] != DBSTATUS_S_DEFAULT)
				break;
		}
		
		if (i <= m_pGridInner->GetColCount())
		{
			m_pGridInner->SetCurrentCell(nRow, i);
		}
		
		// Fire event.
		FireUpdateError();
		
		return E_FAIL;
	}
	
	try
	{
		hr = rac.Update();
	}
	catch (CException * e)
	{
		e->ReportError();
		e->Delete();
		hr = E_FAIL;
	}
	if (FAILED(hr) && hr != E_NOINTERFACE)
	{
		// Report error.
		ReportError();

		// Undo the modification.
		rac.Undo();
		
		// Fire event.
		FireUpdateError();
		
		return E_FAIL;
	}
	
	// Fire event.
	if (bInsert)
	{
		FireAfterInsert(&bDispMsg);
		if (bDispMsg)
			AfxMessageBox(IDS_AFTERINSERT);
	}
	else
	{
		FireAfterUpdate(&bDispMsg);
		if (bDispMsg)
			AfxMessageBox(IDS_AFTERUPDATE);
	}

	return S_OK;
}


int CBHMDBGridCtrl::GetColOrdinalFromIndex(int nColIndex)
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

int CBHMDBGridCtrl::GetColOrdinalFromCol(int nCol)
/*
Routine Description:
	Get the col ordinal in bind info array from col number.

Parameters:
	nCol 		The col number.

Return Value:
	If the col number is valid, return the ordinal; Otherwise, return INVALID.
*/
{
	return GetColOrdinalFromIndex(m_pGridInner->GetColIndex(nCol));
}

int CBHMDBGridCtrl::GetColOrdinalFromField(int nField)
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

void CBHMDBGridCtrl::DeleteRowData(int nRow)
/*
Routine Description:
	Delete one row from data source.

Parameters:
	nRow		The row to delete.

Return Value:
	None.
*/
{
	short bCancel = 0, bDispMsg = 0;
	nRow = nRow >= 0 ? nRow : 0;

	// Row ordinal should be valid.
	if (m_apBookmark.GetSize() < (int)nRow)
		return;

	// Fire event.
	FireBeforeDelete(&bCancel, &bDispMsg);
	if (bCancel)
		return;
	if (bDispMsg && AfxMessageBox(IDS_BEFOREDELETE, MB_YESNO) != IDYES)
		return;

	// Get the HROW of this row.
	CBookmark<0> bmk;
	bmk.SetBookmark(m_nBookmarkSize, m_apBookmark[nRow - 1]);
	HRESULT hr = m_Rowset.MoveToBookmark(bmk);
	if (FAILED(hr))
		return;

	m_Rowset.FreeRecordMemory();

	// Save HROW first.
	HROW hRow = m_Rowset.m_hRow;
	
	// Delete this row.
	try
	{
		hr = m_Rowset.Delete();
	}
	catch (CException * e)
	{
		e->ReportError();
		e->Delete();
		hr = E_FAIL;
	}
	if (FAILED(hr))
	{
		ReportError();
		return;
	}

	// Restore the HROW.
	m_Rowset.m_hRow = hRow;

	// Update the deletion.
	try
	{
		hr = m_Rowset.Update();
	}
	catch (CException * e)
	{
		e->ReportError();
		e->Delete();
		hr = E_FAIL;
	}
	if (FAILED(hr) && hr != E_NOINTERFACE)
	{
		// Failed.
		ReportError();

		// Restore the HROW again.
		m_Rowset.m_hRow = hRow;

		// Undo the deletion.
		m_Rowset.Undo();

		return;
	}
	else
	{
		// Ok, now fire event.
		bDispMsg = 0;
		FireAfterDelete(&bDispMsg);
		if (bDispMsg)
			AfxMessageBox(IDS_AFTERDELETE);

		return;
	}
}

int CBHMDBGridCtrl::GetDataTypeFromStr(CString strField)
/*
Routine Description:
	Get the data type of one field from its name.

Parameters:
	strField		The field name.

Return Value:
	The data type of this field.
*/
{
	if (strField.IsEmpty())
		return 0;

	for (int i = 0; i < m_apColumnData.GetSize(); i ++)
		if (!strField.CompareNoCase(m_apColumnData[i]->strName))
		{
			switch (m_apColumnData[i]->vt)
			{
			case DBTYPE_UI1:
			case DBTYPE_I1:
				return 2;
				break;

			case DBTYPE_I2:
			case DBTYPE_UI2:
				return 3;
				break;

			case DBTYPE_I4:
			case DBTYPE_UI4:
				return 4;
				break;

			case DBTYPE_R4:
			case DBTYPE_NUMERIC:
			case DBTYPE_DECIMAL:
				return 5;
				break;

			case DBTYPE_CY:
				return 6;
				break;

			case DBTYPE_DATE:
			case DBTYPE_FILETIME:
			case DBTYPE_DBDATE:
			case DBTYPE_DBTIME:
			case DBTYPE_DBTIMESTAMP:
				return 7;
				break;

			case DBTYPE_BOOL:
				return 1;
				break;

			default:
				return 0;
				break;
			}
		}

	return 0;
}

void CBHMDBGridCtrl::UndeleteRecordFromHRow(HROW hRow, int nRow)
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

	if (nRow > m_pGridInner->GetRowCount())
		nRow = m_pGridInner->GetRowCount();
	
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
	m_pGridInner->InsertRow(nRow);
	
	// Update the grid.
	m_Rowset.FreeRecordMemory();

	m_pGridInner->Invalidate();
}

void CBHMDBGridCtrl::GetBmkOfRow(int nRow)
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

	// Get data from this record.	
	BYTE * pBookmarkCopy;
	DBBOOKMARK bmFirst = DBBMK_FIRST;
	CBookmark<0> bmk;
	
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

// The backdoor for property page :)
short CBHMDBGridCtrl::GetCtrlType() 
{
	return 0;
}

void CBHMDBGridCtrl::OnDropDown(int nRow, int nCol)
/*
Routine Description:
	Process the dropdown message.

Parameters:
	None.

Return Value:
	None.
*/
{
	if (nRow <= 0 || nCol <= 0)
		return;

	// Retrieve the callback window handle of current column.
	m_hWndDropDown = NULL;
	
	if (m_pGridInner->GetColUserAttribRef(nCol).GetSize() == 3)
		m_hWndDropDown = (HWND)atoi(m_pGridInner->GetColUserAttribRef(nCol).ElementAt(2));

	// Current column should be EditButton.
	if (m_pGridInner->GetColControl(nCol) != COLSTYLE_EDITBUTTON || !IsWindow(m_hWndDropDown))
		return;

	// Toggle the visibility of dropdown window.
	if (m_bDroppedDown)
	{
		HideDropDownWnd();
		return;
	}

	// Calc the point where show the dropdown window.
	CPoint pt;
	CRect rect;
	m_pGridInner->GetCellRect(nRow, nCol, rect);
	pt.x = rect.left;
	pt.y = rect.bottom;
	m_pGridInner->ClientToScreen(&pt);

	// Encryption :)
	long lParam = pt.x + pt.y;
	lParam ^= 235;
	pt.x ^= 235;
	pt.y ^= 235;
	long wParam = MAKEWPARAM(pt.x, pt.y);

	// Send message to dropdown control.

	BringInfo * pBi = new BringInfo;
	pBi->wParam = wParam;
	pBi->lParam = lParam;
	pBi->hWnd = m_hWnd;
	
	CString strText;
	strText = m_pGridInner->GetCellText(nRow, nCol);

	lstrcpy(pBi->strText, strText);

	::SendMessage(m_hWndDropDown, CBHMDataApp::m_nBringMsg, (WPARAM)pBi, NULL);
}

void CBHMDBGridCtrl::HideDropDownWnd()
/*
Routine Description:
	Hide the dropdown window.

Parameters:
	None.

Return Value:
	None.
*/
{
	if (IsWindow(m_hWndDropDown) && m_bDroppedDown)
	{
		// Send message to dropdown control.
		BringInfo * pBi = new BringInfo;
		pBi->hWnd = NULL;
		
		::SendMessage(m_hWndDropDown, CBHMDataApp::m_nBringMsg, (WPARAM)pBi, NULL);
		m_bDroppedDown = FALSE;
	}
}

void CBHMDBGridCtrl::OnBringMsg(long wParam, long lParam)
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
	
	if ((pBi->hWnd))
		// This message tell me the dropdown status and the dropdown
		// window's handle.
	{
		// Update the cached variable.
		m_bDroppedDown = pBi->wParam;
		TRACE1("DroppedDown is %d\t", m_bDroppedDown);
		TRACE1("hookhandle is %x\n", m_hSystemHook);

		if (m_bDroppedDown)
		{
			// The dropdown appears.
			m_hExternalDropDownWnd = pBi->hWnd;

			// install the mouse hook.
			HookMouse(m_hExternalDropDownWnd);
		}
		else if (m_hSystemHook)
		{
			// The dropdown is disappears.

			// Release the mouse hook.
			UnhookWindowsHookEx(m_hSystemHook);
			m_hSystemHook = NULL;
		}
	}
	else
	{
		// The user select one text, we should update the text in current cell.

		int nRow, nCol;
		m_pGridInner->GetCurrentCell(nRow, nCol);

		m_pGridInner->SetCellText(nRow, nCol, pBi->strText);

		m_bDroppedDown = FALSE;

		// The dropdown is disappears.
		// Release the mouse hook.

		if (m_hSystemHook)
		{
			UnhookWindowsHookEx(m_hSystemHook);
			m_hSystemHook = NULL;
		}
	}

	// free the memory previously allocated by message sender
	delete pBi;
}

void CBHMDBGridCtrl::HookMouse(HWND hWnd)
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
	SetHost(m_hWndDropDown, hWnd);

	// Install mouse hook.
	m_hSystemHook = SetWindowsHookEx(WH_MOUSE, MouseProc, (HINSTANCE)GetCurrentThread(), 0);
	
	if (!m_hSystemHook)
	{
#ifdef _DEBUG
		AfxMessageBox("Failed to install hook");
#endif
	}
}

void CBHMDBGridCtrl::OnKillFocus(CWnd* pNewWnd) 
{
	COleControl::OnKillFocus(pNewWnd);

	// Hide the dropdown window when kill focus.
	HideDropDownWnd();
}

OLE_COLOR CBHMDBGridCtrl::GetBackColor() 
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

void CBHMDBGridCtrl::SetBackColor(OLE_COLOR nNewValue) 
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
	if (m_pGridInner)
		m_pGridInner->SetBackColor(TranslateColor(GetBackColor()));

	SetModifiedFlag();
	BoundPropertyChanged(dispidBackColor);
	InvalidateControl();
}

void CBHMDBGridCtrl::OnTextChanged() 
/*
Routine Description:
	Update the caption.

Parameters:
	None.

Return Value:
	None.
*/
{
	CRect rt;

	if (m_pGridInner)
	{
		GetClientRect(&rt);
		
		// The height of caption is the same as that of each level.
		int nHeaderHeight = m_pGridInner->GetLevelHeight();
		
		// Reserve space for caption.
		rt.top += InternalGetText().GetLength() ? nHeaderHeight : 0;
		m_pGridInner->SetWindowPos(NULL, rt.left, rt.top, rt.Width(), rt.Height(), 
			SWP_NOZORDER);
	}

	SetModifiedFlag();
	COleControl::OnTextChanged();
}

void CBHMDBGridCtrl::OnDividerColorChanged() 
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

	SetModifiedFlag();
	BoundPropertyChanged(dispidDividerColor);
	InvalidateControl();
}

void CBHMDBGridCtrl::Bind()
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
	m_arBindInfo.SetSize(m_pGridInner->GetColCount());

	for (i = 0; i < m_pGridInner->GetColCount(); i ++)
	{
		// Fill each binding entry.

		// Create a new entry object.
		ColumnProp * pCol = new ColumnProp;

		// Set the col index of this entry.
		pCol->nColIndex = m_pGridInner->GetColIndex(i + 1);

		// Set the data field value.
		CString strField = m_pGridInner->GetColUserAttribRef(i + 1).ElementAt(0);
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
		m_pGridInner->LockUpdate(FALSE);
		m_pGridInner->Invalidate();
	}
}

void CBHMDBGridCtrl::ExchangeStockProps(CPropExchange *pPX)
/*
Routine Description:
	Exchange stock pros, replace the default implementation for "Appearance" and "BorderStyle".
	This code segment is copied from MFC source.

Parameters:
	pPX

Return Value:
	None.
*/
{
	BOOL bLoading = pPX->IsLoading();
	DWORD dwStockPropMask = GetStockPropMask();
	DWORD dwPersistMask = dwStockPropMask;

	PX_ULong(pPX, _T("_StockProps"), dwPersistMask);

	if (dwStockPropMask & (STOCKPROP_CAPTION | STOCKPROP_TEXT))
	{
		CString strText;

		if (dwPersistMask & (STOCKPROP_CAPTION | STOCKPROP_TEXT))
		{
			if (!bLoading)
				strText = InternalGetText();
			if (dwStockPropMask & STOCKPROP_CAPTION)
				PX_String(pPX, _T("Caption"), strText, _T(""));
			if (dwStockPropMask & STOCKPROP_TEXT)
				PX_String(pPX, _T("Text"), strText, _T(""));
		}
		if (bLoading)
		{
			TRY
				SetText(strText);
			END_TRY
		}
	}

	if (dwStockPropMask & STOCKPROP_FORECOLOR)
	{
		if (dwPersistMask & STOCKPROP_FORECOLOR)
			PX_Color(pPX, _T("ForeColor"), m_clrForeColor, AmbientForeColor());
		else if (bLoading)
			m_clrForeColor = AmbientForeColor();
	}

	if (dwStockPropMask & STOCKPROP_BACKCOLOR)
	{
		if (dwPersistMask & STOCKPROP_BACKCOLOR)
			PX_Color(pPX, _T("BackColor"), m_clrBackColor, AmbientBackColor());
		else if (bLoading)
			m_clrBackColor = AmbientBackColor();
	}

	if (dwStockPropMask & STOCKPROP_FONT)
	{
		LPFONTDISP pFontDispAmbient = AmbientFont();
		BOOL bChanged = TRUE;

		if (dwPersistMask & STOCKPROP_FONT)
			bChanged = PX_Font(pPX, _T("Font"), m_font, NULL, pFontDispAmbient);
		else if (bLoading)
			m_font.InitializeFont(NULL, pFontDispAmbient);

		if (bLoading && bChanged)
			OnFontChanged();

		RELEASE(pFontDispAmbient);
	}

	if (dwStockPropMask & STOCKPROP_BORDERSTYLE)
	{
		short sBorderStyle = m_sBorderStyle;

		if (dwPersistMask & STOCKPROP_BORDERSTYLE)
			PX_Short(pPX, _T("BorderStyle"), m_sBorderStyle, 1);
		else if (bLoading)
			m_sBorderStyle = 0;

		if (sBorderStyle != m_sBorderStyle)
			_AfxToggleBorderStyle(this);
	}

	if (dwStockPropMask & STOCKPROP_ENABLED)
	{
		BOOL bEnabled = m_bEnabled;

		if (dwPersistMask & STOCKPROP_ENABLED)
			PX_Bool(pPX, _T("Enabled"), m_bEnabled, TRUE);
		else if (bLoading)
			m_bEnabled = TRUE;

		if ((bEnabled != m_bEnabled) && (m_hWnd != NULL))
			::EnableWindow(m_hWnd, m_bEnabled);
	}

	if (dwStockPropMask & STOCKPROP_APPEARANCE)
	{
		short sAppearance = m_sAppearance;

		if (dwPersistMask & STOCKPROP_APPEARANCE)
			PX_Short(pPX, _T("Appearance"), m_sAppearance, 0);
		else if (bLoading)
			m_sAppearance = AmbientAppearance();

		if (sAppearance != m_sAppearance)
			_AfxToggleAppearance(this);
	}
}

VARIANT CBHMDBGridCtrl::GetBookmark() 
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
	m_pGridInner->GetCurrentCell(nRow, nCol);

	// Return the bookmark.
	VARIANT sa;
	VariantInit(&sa);
	GetBmkOfRow(nRow, &sa);

	return sa;
}

void CBHMDBGridCtrl::SetBookmark(const VARIANT FAR& newValue) 
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

int CBHMDBGridCtrl::GetRowFromIndex(int nIndex)
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
	for (int nRow = 1; nRow <= m_pGridInner->GetRowCount(); nRow ++)
		if (m_pGridInner->GetRowIndex(nRow) == nIndex)
			return nRow;

	// Did not find.
	return INVALID;
}

void CBHMDBGridCtrl::GetBmkOfRow(int nRow, VARIANT *pVaRet)
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
		pVaRet->lVal = m_pGridInner->GetRowIndex(nRow);

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

int CBHMDBGridCtrl::GetRowFromBmk(const VARIANT *pBmk)
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

LPDISPATCH CBHMDBGridCtrl::GetSelBookmarks()
/*
Routine Description:
	Get the "SelBookmarks" object.

Parameters:
	None.

Return Value:
	The result object.
*/
{
	// Create a new object.
	CSelBookmarks * pBmks = new CSelBookmarks;

	// Set the callback pointer.
	pBmks->SetCtrlPtr(this);

	return pBmks->GetIDispatch(FALSE);
}

VARIANT CBHMDBGridCtrl::RowBookmark(long RowIndex) 
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

long CBHMDBGridCtrl::GetFrozenCols() 
/*
Routine Description:
	Get the number of frozen cols.

Parameters:
	None.

Return Value:
	The number of frozen cols.
*/
{
	// This property is unavailable at design time.
	if (!AmbientUserMode())
		GetNotSupported();

	return m_nFrozenCols;
}

void CBHMDBGridCtrl::SetFrozenCols(long nNewValue) 
/*
Routine Description:
	Set the number of frozen cols.

Parameters:
	nNewValue		The new value.

Return Value:
	None.
*/
{
	// This property is unavailable at design time.
	if (!AmbientUserMode())
		SetNotSupported();

	if (nNewValue < 0 || nNewValue > m_nCols)
		ThrowError(CTL_E_INVALIDPROPERTYVALUE);

	// Save the new value.
	m_nFrozenCols = nNewValue;

	// Update the grid.
	if (m_pGridInner)
	{
		m_pGridInner->SetFrozenCols(m_nFrozenCols);
	}

	SetModifiedFlag();
	BoundPropertyChanged(dispidFrozenCols);
}

int CBHMDBGridCtrl::OnMouseActivate(CWnd* pDesktopWnd, UINT nHitTest, UINT message) 
{
	// Inplace activate on mouse activate.
	if (AmbientUserMode())
		OnActivateInPlace (TRUE, NULL);

	return COleControl::OnMouseActivate(pDesktopWnd, nHitTest, message);
}

LPDISPATCH CBHMDBGridCtrl::GetGroups() 
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
{
	CGroups * pGroups = new CGroups;

	// Set the callback pointer.
	pGroups->SetCtrlPtr(this);
	pGroups->m_pGrid = m_pGridInner;

	return pGroups->GetIDispatch(FALSE);
}

BOOL CBHMDBGridCtrl::GetGroupHeader() 
/*
Routine Description:
	Get the visibility of group header.

Parameters:
	None.

Return Value:
	If the group header should be shown, return TRUE; Otherwise, return FALSE.
*/
{
	if (m_pGridInner)
		m_bGroupHeader = m_pGridInner->GetShowGroupHeader();

	return m_bGroupHeader;
}

void CBHMDBGridCtrl::SetGroupHeader(BOOL bNewValue) 
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
	if (m_pGridInner)
		m_pGridInner->SetShowGroupHeader(m_bGroupHeader);

	SetModifiedFlag();
	BoundPropertyChanged(dispidGroupHeader);
	InvalidateControl();
}
