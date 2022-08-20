#if !defined(AFX_GRIDINPPG_H__BF91B15C_6A84_11D3_A7FE_0080C8763FA4__INCLUDED_)
#define AFX_GRIDINPPG_H__BF91B15C_6A84_11D3_A7FE_0080C8763FA4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// GridInPpg.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CGridInPpg window
class CBHMDBColumnPropPage;
class CBHMDBGroupsPropPage;

class CGridInPpg : public CGrid
{
// Construction
public:
	CGridInPpg();
	SetPagePtr(CBHMDBColumnPropPage * pPage);
	SetPagePtr(CBHMDBGroupsPropPage * pPage);

// Attributes
public:

protected:

	// Callback page pointers.
	CBHMDBColumnPropPage * m_pColumnPage;
	CBHMDBGroupsPropPage * m_pGroupsPage;

// Operations
public:
	// Delete row from grid.
	virtual BOOL DeleteRecord(int nRow);

	// Store new col width.
	virtual void StoreColWidth(int nCol, int nWidth);

	// Store new cell text.
	virtual BOOL StoreCellValue(int nRow, int nCol, CString strNewValue, CString strOldValue);

	// Navigate to new cell.
	virtual BOOL SetCurrentCell(int nRow, int nCol);

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CGridInPpg)
	//}}AFX_VIRTUAL

// Implementation
public:
	virtual ~CGridInPpg();

	// Generated message map functions
protected:
	// Store new header height.
	virtual void StoreHeaderHeight(int nHeight);

	// Store new row height.
	virtual void StoreRowHeight(int nHeight);

	// Move col to new position.
	virtual BOOL MoveCol(int nFromCol, int nToCol);

	// Store new group width.
	virtual void StoreGroupWidth(int nGroup, int nWidth);

	// Group header cell has been clicked.
	virtual void OnGroupHeaderClick(int nGroup);

	//{{AFX_MSG(CGridInPpg)
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_GRIDINPPG_H__BF91B15C_6A84_11D3_A7FE_0080C8763FA4__INCLUDED_)
