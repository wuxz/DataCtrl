#if !defined(AFX_BHMDBGRIDCTL_H__BF91B13B_6A84_11D3_A7FE_0080C8763FA4__INCLUDED_)
#define AFX_BHMDBGRIDCTL_H__BF91B13B_6A84_11D3_A7FE_0080C8763FA4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

// BHMDBGridCtl.h : Declaration of the CBHMDBGridCtrl ActiveX Control class.

/////////////////////////////////////////////////////////////////////////////
// CBHMDBGridCtrl : See BHMDBGridCtl.cpp for implementation.
class CGridInner;
class CGroup;

class CBHMDBGridCtrl : public COleControl
{
	DECLARE_DYNCREATE(CBHMDBGridCtrl)

// Constructor
public:
	CBHMDBGridCtrl();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CBHMDBGridCtrl)
	public:
	virtual void OnDraw(CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid);
	virtual void DoPropExchange(CPropExchange* pPX);
	virtual void OnResetState();
	virtual BOOL OnSetObjectRects(LPCRECT lpRectPos, LPCRECT lpRectClip);
	virtual void OnForeColorChanged();
	virtual BOOL OnGetPredefinedStrings(DISPID dispid, CStringArray* pStringArray, CDWordArray* pCookieArray);
	virtual BOOL OnGetPredefinedValue(DISPID dispid, DWORD dwCookie, VARIANT* lpvarOut);
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	virtual void OnTextChanged();
	//}}AFX_VIRTUAL

// Implementation
protected:
	// Font object used for grid cell.
	CFontHolder m_fntCell;

	// Font object used for header row and caption.
	CFontHolder m_fntHeader;

	// Row height.
	long m_nRowHeight;

	// Header row height.
	long m_nHeaderHeight;

	// Data mode.
	long m_nDataMode;

	// Data source object.
	IDataSource * m_spDataSource;

	// Data member.
	CString m_strDataMember;

	// Default col width.
	long m_nDefColWidth;

	// The permission for deleting record manually.
	BOOL m_bAllowDelete;

	// The permission for inserting record manually.
	BOOL m_bAllowAddNew;

	// The permission for modifying data in grid manually.
	BOOL m_bAllowUpdate;

	// The permission for showing row header.
	BOOL m_bRecordSelectors;

	// Row count.
	long m_nRows;

	// Col count.
	long m_nCols;

	// Level count.
	long m_nLevels;

	// Caption alignment.
	long m_nCapitonAlignment;

	// Divider lines type.
	long m_nDividerType;

	// Divider lines style.
	long m_nDividerStyle;

	// The permission for resizing row height manually.
	BOOL m_bAllowRowSizing;

	// The permission for resizing col width manually.
	BOOL m_bAllowColumnSizing;

	// The permission for showing column header.
	BOOL m_bColumnHeader;

	// The current col ordinal.
	int m_nCol;

	// The current row ordinal.
	int m_nRow;

	// The permission for the grid to redrawing itself.
	BOOL m_bRedraw;

	// The permission for moving column manually.
	BOOL m_bAllowColumnMoving;

	// If the grid acts like a list box.
	BOOL m_bSelectByCell;

	// If the row is dirty.
	BOOL m_bRowChanged;

	// The string used as field separator.
	CString m_strFieldSeparator, m_strSeparator;

	// The background color.
	OLE_COLOR m_clrBack;

	// The frozen col count.
	long m_nFrozenCols;

protected:
	// The cached column binding info.
	CArray<ColumnProp * , ColumnProp * > m_arBindInfo;

	// The permission for showing group header.
	BOOL m_bGroupHeader;

	// If we have predefined layout?
	BOOL m_bHasLayout;

	// The system mouse hook handle.
	HHOOK m_hSystemHook;

	// If the data source has been changed?
	BOOL m_bDataSrcChanged;

	// Have we reached the end of the rowset?
	BOOL m_bEndReached;

	// The rowset object.
	CAccessorRowset<CManualAccessor> m_Rowset;

	// The cols have been named.
	int m_nColsUsed;

	// The global handle for exchanging props.
	HGLOBAL m_hBlob;

	// Column private property array.
	CArray<ColumnData*, ColumnData*> m_apColumnData;

	// Cached bookmark array.
	CArray<BYTE*, BYTE*> m_apBookmark;

	IRowPosition * m_spRowPosition;

	// Sink cookies.
	DWORD m_dwCookieRPC, m_dwCookieRN;

	// If we have bookmarks available?
	BOOL m_bHaveBookmarks;

	int m_nColumns, m_nColumnsBind;

	// Cached field info.
	DBCOLUMNINFO * m_pColumnInfo;
	LPOLESTR m_pStrings;

	// Current row data.
	BindData m_data;

	// The size of bookmark object.
	int m_nBookmarkSize;

	BOOL m_bIsPosSync;

	// The cached HROW array for undeletion.
	CArray<HROW, HROW> m_arHRowToDelete;
	CArray<int, int> m_arRowToDelete;

protected:
	// Given bookmark, calc the row ordinal
	int GetRowFromBmk(const VARIANT * pBmk);

	// Get bookmark from row.
	void GetBmkOfRow(int nRow, VARIANT * pVaRet);

	// Get row ordinal from row index.
	int GetRowFromIndex(int nIndex);

	// Exchange stock props.
	void ExchangeStockProps(CPropExchange* pPX);

	// Bind to datasource.
	void Bind();

	// The handle to dropdown control.
	HWND m_hWndDropDown;

	// Install the mouse hook proc.
	void HookMouse(HWND hWnd);

	// The handle to dropdown window.
	HWND m_hExternalDropDownWnd;

	// Process the message comes from dropdown control.
	void OnBringMsg(long wParam, long lParam);

	// Hide the dropdow window.
	void HideDropDownWnd();

	// The visibility of dropdown window.
	BOOL m_bDroppedDown;

	// Handle the dropdown message.
	void OnDropDown(int nRow, int nCol);

	// Get the bookmark of one row.
	void GetBmkOfRow(int nRow);

	// When one record was undeleted from datasource, update the grid.
	void UndeleteRecordFromHRow(HROW hRow, int nRow);

	// Serialize the column props.
	void SerializeColumnProp(CArchive & ar, BOOL bLoad);

	// Update the grid divider lines.
	void UpdateDivider();

	// Init the grid inner.
	void InitGridHeader();

	// Init the cell font property.
	void InitCellFont();

	// Init the headerfont property.
	void InitHeaderFont();

	// Get field information from data source.
	BOOL GetColumnInfo();

	// Write dirty record back into data source.
	HRESULT UpdateRowData(int nRow);

	// Delete record from data source.
	void DeleteRowData(int nRow);

	// Free the cached field information.
	void FreeColumnMemory();

	// Clearup cached interfaces.
	void ClearInterfaces();

	// When one record was deleted from datasource, update the grid.
	void DeleteRowFromSrc(int nRow);

	// Fetch new records to fill the grid
	void FetchNewRows();

	// Load the data from datasource.
	void GetCellValue(int nRow, int nCol, CString & strText);

	// Get the rowset object.
	HRESULT GetRowset();

	// Given bookmark, calc the row ordinal
	int GetRowFromBmk(BYTE * bmk);

	// Given hrow, calc the row ordinal
	int GetRowFromHRow(HROW hRow);

	// Update control after data source has been changed.
	void UpdateControl();

	// Sync the current row in data source.
	void SetCurrentCell(int nRow, int nCol);

// Implementation
protected:
	~CBHMDBGridCtrl();

	BEGIN_OLEFACTORY(CBHMDBGridCtrl)        // Class factory and guid
		virtual BOOL VerifyUserLicense();
		virtual BOOL GetLicenseKey(DWORD, BSTR FAR*);
	END_OLEFACTORY(CBHMDBGridCtrl)

	DECLARE_OLETYPELIB(CBHMDBGridCtrl)      // GetTypeInfo
	DECLARE_PROPPAGEIDS(CBHMDBGridCtrl)     // Property page IDs
	DECLARE_OLECTLTYPE(CBHMDBGridCtrl)		// Type name and misc status

// Message maps
	//{{AFX_MSG(CBHMDBGridCtrl)
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnKillFocus(CWnd* pNewWnd);
	afx_msg int OnMouseActivate(CWnd* pDesktopWnd, UINT nHitTest, UINT message);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

// Dispatch maps
	//{{AFX_DISPATCH(CBHMDBGridCtrl)
	OLE_COLOR m_clrHeadFore;
	afx_msg void OnHeadForeColorChanged();
	OLE_COLOR m_clrHeadBack;
	afx_msg void OnHeadBackColorChanged();
	OLE_COLOR m_clrDivider;
	afx_msg void OnDividerColorChanged();
	afx_msg long GetCaptionAlignment();
	afx_msg void SetCaptionAlignment(long nNewValue);
	afx_msg long GetDividerType();
	afx_msg void SetDividerType(long nNewValue);
	afx_msg BOOL GetAllowAddNew();
	afx_msg void SetAllowAddNew(BOOL bNewValue);
	afx_msg BOOL GetAllowDelete();
	afx_msg void SetAllowDelete(BOOL bNewValue);
	afx_msg BOOL GetAllowUpdate();
	afx_msg void SetAllowUpdate(BOOL bNewValue);
	afx_msg BOOL GetAllowRowSizing();
	afx_msg void SetAllowRowSizing(BOOL bNewValue);
	afx_msg BOOL GetRecordSelectors();
	afx_msg void SetRecordSelectors(BOOL bNewValue);
	afx_msg BOOL GetAllowColumnSizing();
	afx_msg void SetAllowColumnSizing(BOOL bNewValue);
	afx_msg BOOL GetColumnHeader();
	afx_msg void SetColumnHeader(BOOL bNewValue);
	afx_msg LPFONTDISP GetHeadFont();
	afx_msg void SetHeadFont(LPFONTDISP newValue);
	afx_msg LPFONTDISP GetFont();
	afx_msg void SetFont(LPFONTDISP newValue);
	afx_msg long GetCol();
	afx_msg void SetCol(long nNewValue);
	afx_msg long GetRow();
	afx_msg void SetRow(long nNewValue);
	afx_msg long GetDataMode();
	afx_msg void SetDataMode(long nNewValue);
	afx_msg long GetVisibleCols();
	afx_msg long GetVisibleRows();
	afx_msg long GetLeftCol();
	afx_msg void SetLeftCol(long nNewValue);
	afx_msg long GetCols();
	afx_msg void SetCols(long nNewValue);
	afx_msg long GetRows();
	afx_msg void SetRows(long nNewValue);
	afx_msg BSTR GetFieldSeparator();
	afx_msg void SetFieldSeparator(LPCTSTR lpszNewValue);
	afx_msg long GetDividerStyle();
	afx_msg void SetDividerStyle(long nNewValue);
	afx_msg long GetFirstRow();
	afx_msg void SetFirstRow(long nNewValue);
	afx_msg long GetDefColWidth();
	afx_msg void SetDefColWidth(long nNewValue);
	afx_msg long GetRowHeight();
	afx_msg void SetRowHeight(long nNewValue);
	afx_msg BOOL GetRedraw();
	afx_msg void SetRedraw(BOOL bNewValue);
	afx_msg BOOL GetAllowColumnMoving();
	afx_msg void SetAllowColumnMoving(BOOL bNewValue);
	afx_msg BOOL GetSelectByCell();
	afx_msg void SetSelectByCell(BOOL bNewValue);
	afx_msg LPDISPATCH GetColumns();
	afx_msg LPUNKNOWN GetDataSource();
	afx_msg void SetDataSource(LPUNKNOWN newValue);
	afx_msg BSTR GetDataMember();
	afx_msg void SetDataMember(LPCTSTR lpszNewValue);
	afx_msg BOOL GetRowChanged();
	afx_msg void SetRowChanged(BOOL bNewValue);
	afx_msg long GetHeaderHeight();
	afx_msg void SetHeaderHeight(long nNewValue);
	afx_msg OLE_COLOR GetBackColor();
	afx_msg void SetBackColor(OLE_COLOR nNewValue);
	afx_msg VARIANT GetBookmark();
	afx_msg void SetBookmark(const VARIANT FAR& newValue);
	afx_msg LPDISPATCH GetSelBookmarks();
	afx_msg long GetFrozenCols();
	afx_msg void SetFrozenCols(long nNewValue);
	afx_msg LPDISPATCH GetGroups();
	afx_msg BOOL GetGroupHeader();
	afx_msg void SetGroupHeader(BOOL bNewValue);
	afx_msg void AddItem(LPCTSTR Item, const VARIANT FAR& RowIndex);
	afx_msg void RemoveItem(long RowIndex);
	afx_msg void Scroll(long Rows, long Cols);
	afx_msg void RemoveAll();
	afx_msg void MoveFirst();
	afx_msg void MovePrevious();
	afx_msg void MoveNext();
	afx_msg void MoveLast();
	afx_msg void MoveRecords(long Rows);
	afx_msg BOOL IsAddRow();
	afx_msg void Update();
	afx_msg void CancelUpdate();
	afx_msg void ReBind();
	afx_msg VARIANT RowBookmark(long RowIndex);
	//}}AFX_DISPATCH
	afx_msg short GetCtrlType();
	afx_msg void AboutBox();

	DECLARE_DISPATCH_MAP()

	void FreeBookmarkMemory()
	/*
	Routine Description:
		Free the cached bookmarks.

	Parameters:
		None.

	Return Value:
		None.
	*/
	{
		if (m_apBookmark.GetSize() == 0)
			return;

		int i, nSize = m_apBookmark.GetSize();
		for (i=0; i<nSize; i++)
		{
			delete [] m_apBookmark[i];
		}
		m_apBookmark.RemoveAll();
	}

	void FreeColumnData()
	/*
	Routine Description:
		Clear the cached column data.

	Parameters:
		None.

	Return Value:
		None.
	*/
	{
		if (m_apColumnData.GetSize() == 0)
			return;

		int i, nSize = m_apColumnData.GetSize();
		for (i=0; i<nSize; i++)
		{
			delete m_apColumnData[i];
		}
		m_apColumnData.RemoveAll();
	}

// Event maps
	//{{AFX_EVENT(CBHMDBGridCtrl)
	void FireBtnClick()
		{FireEvent(eventidBtnClick,EVENT_PARAM(VTS_NONE));}
	void FireAfterColUpdate(short ColIndex)
		{FireEvent(eventidAfterColUpdate,EVENT_PARAM(VTS_I2), ColIndex);}
	void FireBeforeColUpdate(short ColIndex, LPCTSTR OldValue, short FAR* Cancel)
		{FireEvent(eventidBeforeColUpdate,EVENT_PARAM(VTS_I2  VTS_BSTR  VTS_PI2), ColIndex, OldValue, Cancel);}
	void FireBeforeInsert(short FAR* Cancel)
		{FireEvent(eventidBeforeInsert,EVENT_PARAM(VTS_PI2), Cancel);}
	void FireBeforeUpdate(short FAR* Cancel)
		{FireEvent(eventidBeforeUpdate,EVENT_PARAM(VTS_PI2), Cancel);}
	void FireColResize(short ColIndex, short FAR* Cancel)
		{FireEvent(eventidColResize,EVENT_PARAM(VTS_I2  VTS_PI2), ColIndex, Cancel);}
	void FireRowResize(short FAR* Cancel)
		{FireEvent(eventidRowResize,EVENT_PARAM(VTS_PI2), Cancel);}
	void FireRowColChange()
		{FireEvent(eventidRowColChange,EVENT_PARAM(VTS_NONE));}
	void FireScroll(short FAR* Cancel)
		{FireEvent(eventidScroll,EVENT_PARAM(VTS_PI2), Cancel);}
	void FireChange()
		{FireEvent(eventidChange,EVENT_PARAM(VTS_NONE));}
	void FireAfterDelete(short FAR* RtnDispErrMsg)
		{FireEvent(eventidAfterDelete,EVENT_PARAM(VTS_PI2), RtnDispErrMsg);}
	void FireBeforeDelete(short FAR* Cancel, short FAR* DispPromptMsg)
		{FireEvent(eventidBeforeDelete,EVENT_PARAM(VTS_PI2  VTS_PI2), Cancel, DispPromptMsg);}
	void FireAfterUpdate(short FAR* RtnDispErrMsg)
		{FireEvent(eventidAfterUpdate,EVENT_PARAM(VTS_PI2), RtnDispErrMsg);}
	void FireAfterInsert(short FAR* RtnDispErrMsg)
		{FireEvent(eventidAfterInsert,EVENT_PARAM(VTS_PI2), RtnDispErrMsg);}
	void FireColSwap(short FromCol, short ToCol, short FAR* Cancel)
		{FireEvent(eventidColSwap,EVENT_PARAM(VTS_I2  VTS_I2  VTS_PI2), FromCol, ToCol, Cancel);}
	void FireUpdateError()
		{FireEvent(eventidUpdateError,EVENT_PARAM(VTS_NONE));}
	void FireBeforeRowColChange(short FAR* Cancel)
		{FireEvent(eventidBeforeRowColChange,EVENT_PARAM(VTS_PI2), Cancel);}
	void FireScrollAfter()
		{FireEvent(eventidScrollAfter,EVENT_PARAM(VTS_NONE));}
	void FireBeforeDrawCell(long Row, long Col, VARIANT FAR* Picture, OLE_COLOR* BkColor, short FAR* DrawText)
		{FireEvent(eventidBeforeDrawCell,EVENT_PARAM(VTS_I4  VTS_I4  VTS_PVARIANT  VTS_PCOLOR  VTS_PI2), Row, Col, Picture, BkColor, DrawText);}
	//}}AFX_EVENT
	DECLARE_EVENT_MAP()

// Dispatch and event IDs
public:
	int GetDataTypeFromStr(CString strField);
	int GetColOrdinalFromField(int nField);
	int GetColOrdinalFromCol(int nCol);
	int GetColOrdinalFromIndex(int nColIndex);
	int GetFieldFromStr(CString str);
	void CreateDrawGrid();
	HWND GetApplicationWindow();
	void ButtonDblClick(UINT nFlags, CPoint point);
	void ButtonDown(UINT nFlags, CPoint point);
	void ButtonUp(UINT nFlags, CPoint point);
	void InitGridFromProp();
	int GetRowColFromVariant(VARIANT va);
	
	CColArray m_arCols;
	CGroupArray m_arGroups;
	CCellArray m_arCells;
	
	CGridInner * m_pGridInner;

	enum {
	//{{AFX_DISP_ID(CBHMDBGridCtrl)
	dispidCaptionAlignment = 4L,
	dispidDividerType = 5L,
	dispidAllowAddNew = 6L,
	dispidAllowDelete = 7L,
	dispidAllowUpdate = 8L,
	dispidAllowRowSizing = 9L,
	dispidRecordSelectors = 10L,
	dispidAllowColumnSizing = 11L,
	dispidColumnHeader = 12L,
	dispidHeadFont = 13L,
	dispidFont = 14L,
	dispidCol = 15L,
	dispidRow = 16L,
	dispidDataMode = 17L,
	dispidVisibleCols = 18L,
	dispidVisibleRows = 19L,
	dispidLeftCol = 20L,
	dispidCols = 21L,
	dispidRows = 22L,
	dispidFieldSeparator = 23L,
	dispidDividerStyle = 24L,
	dispidFirstRow = 25L,
	dispidDefColWidth = 26L,
	dispidRowHeight = 27L,
	dispidRedraw = 28L,
	dispidAllowColumnMoving = 29L,
	dispidSelectByCell = 30L,
	dispidColumns = 31L,
	dispidDataSource = 32L,
	dispidDataMember = 33L,
	dispidRowChanged = 34L,
	dispidHeaderHeight = 35L,
	dispidHeadForeColor = 1L,
	dispidHeadBackColor = 2L,
	dispidBackColor = 36L,
	dispidDividerColor = 3L,
	dispidBookmark = 37L,
	dispidSelBookmarks = 38L,
	dispidFrozenCols = 39L,
	dispidGroups = 40L,
	dispidGroupHeader = 41L,
	dispidAddItem = 42L,
	dispidRemoveItem = 43L,
	dispidScroll = 44L,
	dispidRemoveAll = 45L,
	dispidMoveFirst = 46L,
	dispidMovePrevious = 47L,
	dispidMoveNext = 48L,
	dispidMoveLast = 49L,
	dispidMoveRecords = 50L,
	dispidIsAddRow = 51L,
	dispidUpdate = 52L,
	dispidCancelUpdate = 53L,
	dispidReBind = 54L,
	dispidRowBookmark = 55L,
	eventidBtnClick = 1L,
	eventidAfterColUpdate = 2L,
	eventidBeforeColUpdate = 3L,
	eventidBeforeInsert = 4L,
	eventidBeforeUpdate = 5L,
	eventidColResize = 6L,
	eventidRowResize = 7L,
	eventidRowColChange = 8L,
	eventidScroll = 9L,
	eventidChange = 10L,
	eventidAfterDelete = 11L,
	eventidBeforeDelete = 12L,
	eventidAfterUpdate = 13L,
	eventidAfterInsert = 14L,
	eventidColSwap = 15L,
	eventidUpdateError = 16L,
	eventidBeforeRowColChange = 17L,
	eventidScrollAfter = 18L,
	eventidBeforeDrawCell = 19L,
	//}}AFX_DISP_ID
	dispidCtrlType = 255L,
	};
	BEGIN_INTERFACE_PART(RowPositionChange, IRowPositionChange)
		INIT_INTERFACE_PART(CBHMDBGridCtrl, RowPositionChange)
		STDMETHOD(OnRowPositionChange)(DBREASON, DBEVENTPHASE, BOOL);
	END_INTERFACE_PART(RowPositionChange)

	BEGIN_INTERFACE_PART(RowsetNotify, IRowsetNotify)
		INIT_INTERFACE_PART(CBHMDBGridCtrl, RowsetNotify)
		STDMETHOD(OnFieldChange)(IRowset *, HROW, ULONG, ULONG[], DBREASON, DBEVENTPHASE, BOOL);
		STDMETHOD(OnRowChange)(IRowset *, ULONG, const HROW[], DBREASON, DBEVENTPHASE, BOOL);
		STDMETHOD(OnRowsetChange)(IRowset *, DBREASON, DBEVENTPHASE, BOOL);
	END_INTERFACE_PART(RowsetNotify)

	DECLARE_INTERFACE_MAP()

	BEGIN_INTERFACE_PART(DataSourceListener, IDataSourceListener)
		INIT_INTERFACE_PART(CBHMDBGridCtrl, DataSourceListener)
		STDMETHOD(dataMemberChanged)(BSTR bstrDM);
		STDMETHOD(dataMemberAdded)(BSTR bstrDM);
		STDMETHOD(dataMemberRemoved)(BSTR bstrDM);
	END_INTERFACE_PART(DataSourceListener)

	friend class CColumn;
	friend class CColumns;
	friend class CGridInner;
	friend class CBHMDBColumnPropPage;
	friend class CSelBookmarks;
	friend class CGroups;
	friend class CBHMDBGroupsPropPage;
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_BHMDBGRIDCTL_H__BF91B13B_6A84_11D3_A7FE_0080C8763FA4__INCLUDED)
