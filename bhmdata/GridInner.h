#if !defined(AFX_GRIDINNER_H__BF91B15B_6A84_11D3_A7FE_0080C8763FA4__INCLUDED_)
#define AFX_GRIDINNER_H__BF91B15B_6A84_11D3_A7FE_0080C8763FA4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// GridInner.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CGridInner window
class CBHMDBGridCtrl;

class CGridInner : public CGrid
{
// Construction
public:
	CGridInner(CBHMDBGridCtrl * pCtrl);

// Attributes
public:
	// The dirty state for each col.
	BOOL m_bColDirty[255];

	// The content of the modified fields of current row.
	CStringArray m_arContent;

protected:
	// The callback control pointer.
	CBHMDBGridCtrl * m_pCtrl;

// Operations
public:
	// Get the record count.
	int GetRecordCount();
	
	// Cancel the pending modification in current row.
	virtual void CancelRecord();

	// Navigate to new cell.
	virtual BOOL SetCurrentCell(int nRow, int nCol);

	// Load the cell text.
	virtual void OnLoadCellText(int nRow, int nCol, CString &strText);

	// Move col to new position.
	virtual BOOL MoveCol(int nFromCol, int nToCol);

	// Begin mouse tracking.
	virtual BOOL OnStartTracking(int nRow, int nCol, int nTrackingMode);

	// Add new row.
	virtual BOOL AddNew();

	// Store cell text.
	virtual BOOL StoreCellValue(int nRow, int nCol, CString strNewValue, CString strOldValue);

	// Button in cell is clicked.
	virtual void OnClickedCellButton(int nRow, int nCol);

	// Delete one row.
	virtual BOOL DeleteRecord(int nRow);

	// Save pending modification in current row.
	virtual BOOL FlushRecord();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CGridInner)
	//}}AFX_VIRTUAL

// Implementation
public:
	// Draw text in cell.
	virtual void OnDrawCellText(CDC * pDC, CellStyle * pStyle);

	virtual ~CGridInner();

	// Generated message map functions
protected:
	//{{AFX_MSG(CGridInner)
	afx_msg void OnLButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnLButtonDblClk(UINT nFlags, CPoint point);
	afx_msg void OnMouseMove(UINT nFlags, CPoint point);
	afx_msg void OnMButtonDblClk(UINT nFlags, CPoint point);
	afx_msg void OnMButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnRButtonDblClk(UINT nFlags, CPoint point);
	afx_msg void OnRButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnLButtonUp(UINT nFlags, CPoint point);
	afx_msg void OnMButtonUp(UINT nFlags, CPoint point);
	afx_msg void OnRButtonUp(UINT nFlags, CPoint point);
	afx_msg void OnKillFocus(CWnd* pNewWnd);
	afx_msg void OnSetFocus(CWnd* pOldWnd);
	afx_msg void OnHScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar);
	afx_msg void OnVScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar);
	afx_msg void OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);
	afx_msg void OnSysKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_GRIDINNER_H__BF91B15B_6A84_11D3_A7FE_0080C8763FA4__INCLUDED_)
