#if !defined(AFX_BHMDBCOMBOCTL_H__BF91B145_6A84_11D3_A7FE_0080C8763FA4__INCLUDED_)
#define AFX_BHMDBCOMBOCTL_H__BF91B145_6A84_11D3_A7FE_0080C8763FA4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

// BHMDBComboCtl.h : Declaration of the CBHMDBComboCtrl ActiveX Control class.

/////////////////////////////////////////////////////////////////////////////
// CBHMDBComboCtrl : See BHMDBComboCtl.cpp for implementation.
class CDropDownColumn;
class CDropDownRealWnd;

#define IDC_EDIT 1001
#define IDC_BUTTON 1002
#define BUTTONWIDTH 18
#define BUTTONHEIGHT 18
#define PIXELRATIO 40

// the edit part
class CComboEdit;

class CBHMDBComboCtrl : public COleControl
{
	DECLARE_DYNCREATE(CBHMDBComboCtrl)

// Constructor
public:
	CBHMDBComboCtrl();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CBHMDBComboCtrl)
	public:
	virtual void OnForeColorChanged();
	virtual void OnDraw(CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid);
	virtual void DoPropExchange(CPropExchange* pPX);
	virtual void OnResetState();
	virtual BOOL OnGetPredefinedStrings(DISPID dispid, CStringArray* pStringArray, CDWordArray* pCookieArray);
	virtual BOOL OnGetPredefinedValue(DISPID dispid, DWORD dwCookie, VARIANT* lpvarOut);
	virtual BOOL OnSetObjectRects(LPCRECT lpRectPos, LPCRECT lpRectClip);
	virtual BOOL OnSetExtent(LPSIZEL lpSizeL);
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	//}}AFX_VIRTUAL

// Implementation
protected:
	~CBHMDBComboCtrl();

	// Install the mouse hook proc.
	void HookMouse(HWND hWnd);

	// Show the dropdown window.
	void ShowDropDownWnd();

	// Hide the dropdow window.
	void HideDropDownWnd();

	// Handle the dropdown message.
	void OnDropDown();
	
	// Bind to datasource.
	void Bind();

	// Search the text in datasource or manually input data.
	int SearchText(CString strSearch, BOOL bAutoPosition = TRUE);

	// When one record was deleted from datasource, update the grid.
	void DeleteRowFromSrc(int nRow);
	
	// Update the dropdown window.
	void UpdateDropDownWnd();

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

	// Serialize the column props.
	void SerializeColumnProp(CArchive & ar, BOOL bLoad);

	// Init the cell font property.
	void InitCellFont();

	// Update the grid divider lines.
	void UpdateDivider();

	// Decide if the char is valid.
	BOOL IsValidChar(UINT nChar, int nPosition);
	
	// Decide if the given text format is valid.
	BOOL IsValidFormat(CString strNewFormat);

	// Decide if the given text format is valid.
	BOOL IsValidText(CString strNewText);

	// Calc input mask.
	void CalcMask();

protected:

	// If show group header.
	BOOL m_bGroupHeader;

	// Control style.
	long m_nStyle;

	// The max length of the text of the editbox.
	short m_nMaxLength;

	// The text being shown.
	CString m_strText;

	// The current text format.
	CString m_strFormat;

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

	// Whether the editbox is readonly or not.
	BOOL m_bReadOnly;

protected:
	// Get row ordinal from bookmark object.
	int GetRowFromBmk(const VARIANT *pBmk);

	// Get bookmark from row.
	void GetBmkOfRow(int nRow, VARIANT *pVaRet);

	// Get row ordinal from row index.
	int GetRowFromIndex(int nIndex);
	
	// Cached bind info.
	CArray<ColumnProp * , ColumnProp * > m_arBindInfo;

	// Whether we have predefined layout.
	BOOL m_bHasLayout;

	// The dropdown window.
	CDropDownRealWnd * m_pDropDownRealWnd;

	// The brush to draw the background.
	CBrush * m_pBrhBack;

	// The system mouse hook handle.
	HHOOK m_hSystemHook;

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
	
	// Sink cookie
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
	
	// Is the LButton pressed?
	BOOL m_bLButtonDown;

	// The rect to draw the button.
	CRect m_rectButton;

	// The current text mask.
	CString m_strMask;


	// The signal to mark whether current char is delimiter or data.
	BOOL m_bIsDelimiter[255];

	// The pointer of the editbox window.
	CComboEdit * m_pEdit;

	// The handle of the using font for editbox
	HFONT m_hFont;

	// The height of the using font.
	int m_nFontHeight;

	BEGIN_OLEFACTORY(CBHMDBComboCtrl)        // Class factory and guid
		virtual BOOL VerifyUserLicense();
		virtual BOOL GetLicenseKey(DWORD, BSTR FAR*);
	END_OLEFACTORY(CBHMDBComboCtrl)

	DECLARE_OLETYPELIB(CBHMDBComboCtrl)      // GetTypeInfo
	DECLARE_PROPPAGEIDS(CBHMDBComboCtrl)     // Property page IDs
	DECLARE_OLECTLTYPE(CBHMDBComboCtrl)		// Type name and misc status

// Message maps
	//{{AFX_MSG(CBHMDBComboCtrl)
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnKillFocus(CWnd* pNewWnd);
	afx_msg void OnSetFocus(CWnd* pOldWnd);
	afx_msg void OnLButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnLButtonUp(UINT nFlags, CPoint point);
	afx_msg HBRUSH OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor);
	afx_msg int OnMouseActivate(CWnd* pDesktopWnd, UINT nHitTest, UINT message);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

// Dispatch maps
	//{{AFX_DISPATCH(CBHMDBComboCtrl)
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
	afx_msg OLE_COLOR GetBackColor();
	afx_msg void SetBackColor(OLE_COLOR nNewValue);
	afx_msg BSTR GetFormat();
	afx_msg void SetFormat(LPCTSTR lpszNewValue);
	afx_msg short GetMaxLength();
	afx_msg void SetMaxLength(short nNewValue);
	afx_msg BOOL GetReadOnly();
	afx_msg void SetReadOnly(BOOL bNewValue);
	afx_msg LPUNKNOWN GetDataSource();
	afx_msg void SetDataSource(LPUNKNOWN newValue);
	afx_msg BSTR GetDataMember();
	afx_msg void SetDataMember(LPCTSTR lpszNewValue);
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
	afx_msg void SetDroppedDown(BOOL bNewValue);
	afx_msg BSTR GetDataField();
	afx_msg void SetDataField(LPCTSTR lpszNewValue);
	afx_msg LPDISPATCH GetColumns();
	afx_msg long GetVisibleRows();
	afx_msg long GetVisibleCols();
	afx_msg BOOL GetColumnHeader();
	afx_msg void SetColumnHeader(BOOL bNewValue);
	afx_msg LPFONTDISP GetFont();
	afx_msg void SetFont(LPFONTDISP newValue);
	afx_msg LPFONTDISP GetHeadFont();
	afx_msg void SetHeadFont(LPFONTDISP newValue);
	afx_msg BSTR GetFieldSeparator();
	afx_msg void SetFieldSeparator(LPCTSTR lpszNewValue);
	afx_msg short GetMaxDropDownItems();
	afx_msg void SetMaxDropDownItems(short nNewValue);
	afx_msg short GetMinDropDownItems();
	afx_msg void SetMinDropDownItems(short nNewValue);
	afx_msg short GetListWidth();
	afx_msg void SetListWidth(short nNewValue);
	afx_msg long GetHeaderHeight();
	afx_msg void SetHeaderHeight(long nNewValue);
	afx_msg BSTR GetText();
	afx_msg void SetText(LPCTSTR lpszNewValue);
	afx_msg VARIANT GetBookmark();
	afx_msg void SetBookmark(const VARIANT FAR& newValue);
	afx_msg long GetStyle();
	afx_msg void SetStyle(long nNewValue);
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
	afx_msg BOOL IsItemInList();
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
	//{{AFX_EVENT(CBHMDBComboCtrl)
	void FireScroll(short FAR* Cancel)
		{FireEvent(eventidScroll,EVENT_PARAM(VTS_PI2), Cancel);}
	void FireScrollAfter()
		{FireEvent(eventidScrollAfter,EVENT_PARAM(VTS_NONE));}
	void FireCloseUp()
		{FireEvent(eventidCloseUp,EVENT_PARAM(VTS_NONE));}
	void FireDropDown()
		{FireEvent(eventidDropDown,EVENT_PARAM(VTS_NONE));}
	void FirePositionList(LPCTSTR Text)
		{FireEvent(eventidPositionList,EVENT_PARAM(VTS_BSTR), Text);}
	void FireChange()
		{FireEvent(eventidChange,EVENT_PARAM(VTS_NONE));}
	//}}AFX_EVENT
	DECLARE_EVENT_MAP()

// Dispatch and event IDs
public:
	// Load the data from datasource.
	void GetCellValue(int nRow, int nCol, CString& strText);

	// Init the grid from props saved.
	void InitGridFromProp();

	// Get the ordinal in columns array.
	int GetColOrdinalFromField(int nField);
	int GetColOrdinalFromCol(int nCol);
	int GetColOrdinalFromIndex(int nColIndex);

	// Get the field ordinal of the datasource.
	int GetFieldFromStr(CString str);

	// Convert variant to integer value.
	int GetRowColFromVariant(VARIANT va);
	
	// Cached layout data.
	CColArray m_arCols;
	CGroupArray m_arGroups;
	CCellArray m_arCells;

	enum {
	//{{AFX_DISP_ID(CBHMDBComboCtrl)
	dispidHeadBackColor = 1L,
	dispidHeadForeColor = 2L,
	dispidBackColor = 6L,
	dispidFormat = 7L,
	dispidMaxLength = 8L,
	dispidReadOnly = 9L,
	dispidDataSource = 10L,
	dispidDataMember = 11L,
	dispidDefColWidth = 12L,
	dispidRowHeight = 13L,
	dispidDividerType = 14L,
	dispidDividerStyle = 15L,
	dispidDataMode = 16L,
	dispidLeftCol = 17L,
	dispidFirstRow = 18L,
	dispidRow = 19L,
	dispidCol = 20L,
	dispidRows = 21L,
	dispidCols = 22L,
	dispidDroppedDown = 23L,
	dispidDataField = 24L,
	dispidColumns = 25L,
	dispidVisibleRows = 26L,
	dispidVisibleCols = 27L,
	dispidColumnHeader = 28L,
	dispidFont = 29L,
	dispidHeadFont = 30L,
	dispidFieldSeparator = 31L,
	dispidListAutoPosition = 3L,
	dispidListWidthAutoSize = 4L,
	dispidMaxDropDownItems = 32L,
	dispidMinDropDownItems = 33L,
	dispidListWidth = 34L,
	dispidHeaderHeight = 35L,
	dispidText = 36L,
	dispidBookmark = 37L,
	dispidStyle = 38L,
	dispidGroups = 39L,
	dispidDividerColor = 5L,
	dispidGroupHeader = 40L,
	dispidAddItem = 41L,
	dispidRemoveItem = 42L,
	dispidScroll = 43L,
	dispidRemoveAll = 44L,
	dispidMoveFirst = 45L,
	dispidMovePrevious = 46L,
	dispidMoveNext = 47L,
	dispidMoveLast = 48L,
	dispidMoveRecords = 49L,
	dispidReBind = 50L,
	dispidIsItemInList = 51L,
	dispidRowBookmark = 52L,
	eventidScroll = 1L,
	eventidScrollAfter = 2L,
	eventidCloseUp = 3L,
	eventidDropDown = 4L,
	eventidPositionList = 5L,
	eventidChange = 6L,
	//}}AFX_DISP_ID
	dispidCtrlType = 255L,
	};

	DECLARE_INTERFACE_MAP()

	BEGIN_INTERFACE_PART(RowsetNotify, IRowsetNotify)
		INIT_INTERFACE_PART(CBHMDBComboCtrl, RowsetNotify)
		STDMETHOD(OnFieldChange)(IRowset *, HROW, ULONG, ULONG[], DBREASON, DBEVENTPHASE, BOOL);
		STDMETHOD(OnRowChange)(IRowset *, ULONG, const HROW[], DBREASON, DBEVENTPHASE, BOOL);
		STDMETHOD(OnRowsetChange)(IRowset *, DBREASON, DBEVENTPHASE, BOOL);
	END_INTERFACE_PART(RowsetNotify)

	BEGIN_INTERFACE_PART(DataSourceListener, IDataSourceListener)
		INIT_INTERFACE_PART(CBHMDBComboCtrl, DataSourceListener)
		STDMETHOD(dataMemberChanged)(BSTR bstrDM);
		STDMETHOD(dataMemberAdded)(BSTR bstrDM);
		STDMETHOD(dataMemberRemoved)(BSTR bstrDM);
	END_INTERFACE_PART(DataSourceListener)

	friend class CBHMDBColumnPropPage;
	friend class CDropDownRealWnd;
	friend class CComboEdit;
	friend class CColumns;
	friend class CColumn;
	friend class CGroups;
	friend class CBHMDBGroupsPropPage;
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_BHMDBCOMBOCTL_H__BF91B145_6A84_11D3_A7FE_0080C8763FA4__INCLUDED)
