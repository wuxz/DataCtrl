#if !defined(AFX_BHMDBDROPDOWNCTL_H__BF91B140_6A84_11D3_A7FE_0080C8763FA4__INCLUDED_)
#define AFX_BHMDBDROPDOWNCTL_H__BF91B140_6A84_11D3_A7FE_0080C8763FA4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

// BHMDBDropDownCtl.h : Declaration of the CBHMDBDropDownCtrl ActiveX Control class.

/////////////////////////////////////////////////////////////////////////////
// CBHMDBDropDownCtrl : See BHMDBDropDownCtl.cpp for implementation.
class CDropDownColumn;
class CDropDownRealWnd;
class CGridInner;

class CBHMDBDropDownCtrl : public COleControl
{
	DECLARE_DYNCREATE(CBHMDBDropDownCtrl)

// Constructor
public:
	CBHMDBDropDownCtrl();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CBHMDBDropDownCtrl)
	public:
	virtual void OnDraw(CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid);
	virtual void DoPropExchange(CPropExchange* pPX);
	virtual void OnResetState();
	virtual void OnForeColorChanged();
	virtual BOOL OnGetPredefinedStrings(DISPID dispid, CStringArray* pStringArray, CDWordArray* pCookieArray);
	virtual BOOL OnGetPredefinedValue(DISPID dispid, DWORD dwCookie, VARIANT* lpvarOut);
	//}}AFX_VIRTUAL

// Implementation
protected:
	~CBHMDBDropDownCtrl();

protected:
	// If show group header.
	BOOL m_bGroupHeader;

	// The font currently used.
	CFontHolder m_fntCell;

	// The font used in grid header row.
	CFontHolder m_fntHeader;

	// The row height.
	long m_nRowHeight;

	// The header row height.
	long m_nHeaderHeight;

	// The data mode.
	long m_nDataMode;

	// The data source pointer.
	IDataSource * m_spDataSource;

	// The data member.
	CString m_strDataMember;

	// Default col width.
	long m_nDefColWidth;

	// The row count.
	long m_nRows;

	// The col count.
	long m_nCols;

	// The divider type.
	long m_nDividerType;

	// The divider style.
	long m_nDividerStyle;

	// The current col.
	int m_nCol;

	// The current row.
	int m_nRow;

	// Field separator.
	CString m_strFieldSeparator, m_strSeparator;

	// Controling data field.
	CString m_strDataField;

	// Whether show the column header or not.
	BOOL m_bColumnHeader;

	// Max items on screen.
	short m_nMaxDropDownItems;

	// Min items on screen.
	short m_nMinDropDownItems;

	// The width of the dropdown window.
	short m_nListWidth;

	// The backgroud color.
	OLE_COLOR m_clrBack;

protected:
	// Get row ordinal from bookmark object.
	int GetRowFromBmk(const VARIANT *pBmk);

	// Get row ordinal from row index.
	int GetRowFromIndex(int nIndex);

	// Get bookmark from row.
	void GetBmkOfRow(int nRow, VARIANT *pVaRet);

	// Bind to datasource.
	void Bind();

	// Search the text in datasource or manually input data.
	void SearchText(CString strSearch);

	// When one record was deleted from datasource, update the grid.
	void DeleteRowFromSrc(int nRow);

	// Update the dropdown window.
	void UpdateDropDownWnd();

	// Show the dropdown window
	void Bring(int cx, int cy);

	// Hide the dropdown window.
	void HideDropDownWnd(LPCTSTR pCharValue = NULL);

	// The text being shown.
	CString m_strText;

	// Process messages from grid control.
	void OnBringMsg(long wParam, long lParam);

	// Get the bookmark of one row.
	void GetBmkOfRow(int nRow);

	// When one record was undeleted from datasource, update the grid.
	void UndeleteRecordFromHRow(HROW hRow);

	// Fetch new records to fill the grid
	void FetchNewRows();

	// Given hrow, calc the row ordinal
	int GetRowFromHRow(HROW hRow);

	// Given bookmark, calc the row ordinal
	int GetRowFromBmk(BYTE *bmk);

	// Init the headerfont property.
	void InitHeaderFont();

	// init the dropdown window.
	void InitDropDownWnd();

	// Get the rowset object.
	HRESULT GetRowset();

	// Get field information from data source.
	BOOL GetColumnInfo();

	// Free the cached field information.
	void FreeColumnMemory();

	// Clearup cached interfaces.
	void ClearInterfaces();

	virtual BOOL OnSetObjectRects(LPCRECT lpRectPos, LPCRECT lpRectClip);

	// Serialize the column props.
	void SerializeColumnProp(CArchive & ar, BOOL bLoad);

	// Init the cell font property.
	void InitCellFont();

	// Update the grid divider lines.
	void UpdateDivider();

	// Init the internal grid.
	void InitDrawGrid();

	// The internal grid window.
	CGridInner * m_pGridDraw;

	// The dropdown window.
	CDropDownRealWnd * m_pDropDownRealWnd;

	// Cached bind info.
	CArray<ColumnProp * , ColumnProp * > m_arBindInfo;

	// Whether we have predefined layout.
	BOOL m_bHasLayout;

	// The handle of the blob data serialize columns props.
	HGLOBAL m_hBlob;

	// Whether the datasource was changed or not.
	BOOL m_bDataSrcChanged;

	// Have we reached the end of the rowset?
	BOOL m_bEndReached;

	// The rowset object.
	CAccessorRowset<CManualAccessor> m_Rowset;

	// The used number of the col name.
	int m_nColsUsed;

	// The filtered field info.
	CArray<ColumnData*, ColumnData*> m_apColumnData;

	// The cached bookmarks.
	CArray<BYTE*, BYTE*> m_apBookmark;

	IRowPosition * m_spRowPosition;

	// Sink cookie.
	DWORD m_dwCookieRN;

	// Whether the datasource support bookmark or not.
	BOOL m_bHaveBookmarks;

	int m_nColumns, m_nColumnsBind;

	// The raw column info data.
	DBCOLUMNINFO * m_pColumnInfo;
	LPOLESTR m_pStrings;

	// The memory to hold the data of current record.
	BindData m_data;

	// The size of bookmark.
	int m_nBookmarkSize;

	// The handle of the grid control window.
	HWND m_hConsumer;

	BEGIN_OLEFACTORY(CBHMDBDropDownCtrl)        // Class factory and guid
		virtual BOOL VerifyUserLicense();
		virtual BOOL GetLicenseKey(DWORD, BSTR FAR*);
	END_OLEFACTORY(CBHMDBDropDownCtrl)

	DECLARE_OLETYPELIB(CBHMDBDropDownCtrl)      // GetTypeInfo
	DECLARE_PROPPAGEIDS(CBHMDBDropDownCtrl)     // Property page IDs
	DECLARE_OLECTLTYPE(CBHMDBDropDownCtrl)		// Type name and misc status

// Message maps
	//{{AFX_MSG(CBHMDBDropDownCtrl)
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

// Dispatch maps
	//{{AFX_DISPATCH(CBHMDBDropDownCtrl)
	OLE_COLOR m_clrHeadBack;
	afx_msg void OnHeadBackColorChanged();
	OLE_COLOR m_clrHeadFore;
	afx_msg void OnHeadForeColorChanged();
	BOOL m_bListAutoPosition;
	afx_msg void OnListAutoPositionChanged();
	BOOL m_bListWidthAutoSize;
	afx_msg void OnListWidthAutoSizeChanged();
	OLE_COLOR m_clrDivider;
	afx_msg void OnDividerColorChanged();
	afx_msg LPFONTDISP GetFont();
	afx_msg void SetFont(LPFONTDISP newValue);
	afx_msg long GetDefColWidth();
	afx_msg void SetDefColWidth(long nNewValue);
	afx_msg long GetRowHeight();
	afx_msg void SetRowHeight(long nNewValue);
	afx_msg long GetDividerType();
	afx_msg void SetDividerType(long nNewValue);
	afx_msg long GetDividerStyle();
	afx_msg void SetDividerStyle(long nNewValue);
	afx_msg long GetDataMode();
	afx_msg void SetDataMode(long nNewValue);
	afx_msg long GetLeftCol();
	afx_msg void SetLeftCol(long nNewValue);
	afx_msg long GetFirstRow();
	afx_msg void SetFirstRow(long nNewValue);
	afx_msg long GetRow();
	afx_msg void SetRow(long nNewValue);
	afx_msg long GetCol();
	afx_msg void SetCol(long nNewValue);
	afx_msg long GetRows();
	afx_msg void SetRows(long nNewValue);
	afx_msg long GetCols();
	afx_msg void SetCols(long nNewValue);
	afx_msg BOOL GetDroppedDown();
	afx_msg BSTR GetDataField();
	afx_msg void SetDataField(LPCTSTR lpszNewValue);
	afx_msg LPDISPATCH GetColumns();
	afx_msg long GetVisibleRows();
	afx_msg long GetVisibleCols();
	afx_msg BOOL GetColumnHeader();
	afx_msg void SetColumnHeader(BOOL bNewValue);
	afx_msg LPFONTDISP GetHeadFont();
	afx_msg void SetHeadFont(LPFONTDISP newValue);
	afx_msg long GetHeaderHeight();
	afx_msg void SetHeaderHeight(long nNewValue);
	afx_msg LPUNKNOWN GetDataSource();
	afx_msg void SetDataSource(LPUNKNOWN newValue);
	afx_msg BSTR GetDataMember();
	afx_msg void SetDataMember(LPCTSTR lpszNewValue);
	afx_msg BSTR GetFieldSeparator();
	afx_msg void SetFieldSeparator(LPCTSTR lpszNewValue);
	afx_msg short GetMaxDropDownItems();
	afx_msg void SetMaxDropDownItems(short nNewValue);
	afx_msg short GetMinDropDownItems();
	afx_msg void SetMinDropDownItems(short nNewValue);
	afx_msg short GetListWidth();
	afx_msg void SetListWidth(short nNewValue);
	afx_msg OLE_COLOR GetBackColor();
	afx_msg void SetBackColor(OLE_COLOR nNewValue);
	afx_msg VARIANT GetBookmark();
	afx_msg void SetBookmark(const VARIANT FAR& newValue);
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
	//{{AFX_EVENT(CBHMDBDropDownCtrl)
	void FireScroll(short FAR* Cancel)
		{FireEvent(eventidScroll,EVENT_PARAM(VTS_PI2), Cancel);}
	void FireCloseUp()
		{FireEvent(eventidCloseUp,EVENT_PARAM(VTS_NONE));}
	void FireDropDown()
		{FireEvent(eventidDropDown,EVENT_PARAM(VTS_NONE));}
	void FirePositionList(LPCTSTR Text)
		{FireEvent(eventidPositionList,EVENT_PARAM(VTS_BSTR), Text);}
	void FireScrollAfter()
		{FireEvent(eventidScrollAfter,EVENT_PARAM(VTS_NONE));}
	//}}AFX_EVENT
	DECLARE_EVENT_MAP()

// Dispatch and event IDs
public:
	// If the dropdown window is visible.
	BOOL m_bDroppedDown;

	// Load the data from datasource.
	void GetCellValue(int nRow, int nCol, CString & strText);

	// Init the grid from props saved.
	void InitGridFromProp();

	// Get the ordinal in columns array.
	int GetColOrdinalFromField(int nField);
	int GetColOrdinalFromCol(int nCol);
	int GetColOrdinalFromIndex(int nColIndex);

	// Get the field ordinal of the datasource.
	int GetFieldFromStr(CString str);

	// Create the internal grid window.
	void CreateDrawGrid();

	// Get the parent window for hold the internal grid.
	HWND GetApplicationWindow();

	// Convert variant to integer value.
	int GetRowColFromVariant(VARIANT va);

	// Cached layout data.
	CColArray m_arCols;
	CGroupArray m_arGroups;
	CCellArray m_arCells;

	enum {
	//{{AFX_DISP_ID(CBHMDBDropDownCtrl)
	dispidFont = 6L,
	dispidDefColWidth = 7L,
	dispidRowHeight = 8L,
	dispidDividerType = 9L,
	dispidDividerStyle = 10L,
	dispidDataMode = 11L,
	dispidLeftCol = 12L,
	dispidFirstRow = 13L,
	dispidRow = 14L,
	dispidCol = 15L,
	dispidRows = 16L,
	dispidCols = 17L,
	dispidDroppedDown = 18L,
	dispidDataField = 19L,
	dispidColumns = 20L,
	dispidVisibleRows = 21L,
	dispidVisibleCols = 22L,
	dispidHeadBackColor = 1L,
	dispidHeadForeColor = 2L,
	dispidColumnHeader = 23L,
	dispidHeadFont = 24L,
	dispidHeaderHeight = 25L,
	dispidDataSource = 26L,
	dispidDataMember = 27L,
	dispidFieldSeparator = 28L,
	dispidListAutoPosition = 3L,
	dispidListWidthAutoSize = 4L,
	dispidMaxDropDownItems = 29L,
	dispidMinDropDownItems = 30L,
	dispidListWidth = 31L,
	dispidBackColor = 32L,
	dispidBookmark = 33L,
	dispidGroups = 34L,
	dispidDividerColor = 5L,
	dispidGroupHeader = 35L,
	dispidAddItem = 36L,
	dispidRemoveItem = 37L,
	dispidScroll = 38L,
	dispidRemoveAll = 39L,
	dispidMoveFirst = 40L,
	dispidMovePrevious = 41L,
	dispidMoveNext = 42L,
	dispidMoveLast = 43L,
	dispidMoveRecords = 44L,
	dispidReBind = 45L,
	dispidRowBookmark = 46L,
	eventidScroll = 1L,
	eventidCloseUp = 2L,
	eventidDropDown = 3L,
	eventidPositionList = 4L,
	eventidScrollAfter = 5L,
	//}}AFX_DISP_ID
	dispidCtrlType = 255L,
	};
	DECLARE_INTERFACE_MAP()

	BEGIN_INTERFACE_PART(RowsetNotify, IRowsetNotify)
		INIT_INTERFACE_PART(CBHMDBDropDownCtrl, RowsetNotify)
		STDMETHOD(OnFieldChange)(IRowset *, HROW, ULONG, ULONG[], DBREASON, DBEVENTPHASE, BOOL);
		STDMETHOD(OnRowChange)(IRowset *, ULONG, const HROW[], DBREASON, DBEVENTPHASE, BOOL);
		STDMETHOD(OnRowsetChange)(IRowset *, DBREASON, DBEVENTPHASE, BOOL);
	END_INTERFACE_PART(RowsetNotify)

	BEGIN_INTERFACE_PART(DataSourceListener, IDataSourceListener)
		INIT_INTERFACE_PART(CBHMDBDropDownCtrl, DataSourceListener)
		STDMETHOD(dataMemberChanged)(BSTR bstrDM);
		STDMETHOD(dataMemberAdded)(BSTR bstrDM);
		STDMETHOD(dataMemberRemoved)(BSTR bstrDM);
	END_INTERFACE_PART(DataSourceListener)

	friend class CBHMDBColumnPropPage;
	friend class CDropDownRealWnd;
	friend class CColumns;
	friend class CColumn;
	friend class CGroups;
	friend class CBHMDBGroupsPropPage;
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_BHMDBDROPDOWNCTL_H__BF91B140_6A84_11D3_A7FE_0080C8763FA4__INCLUDED)
