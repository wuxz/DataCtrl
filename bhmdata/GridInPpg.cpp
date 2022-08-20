// GridInPpg.cpp : implementation file
//

#include "stdafx.h"
#include "bhmdata.h"
#include "GridInPpg.h"
#include "BHMDBColumnPropPage.h"
#include "BHMDBGroupsPropPage.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CGridInPpg

CGridInPpg::CGridInPpg()
{
	m_pColumnPage = NULL;
	m_pGroupsPage = NULL;
}

CGridInPpg::~CGridInPpg()
{
}

CGridInPpg::SetPagePtr(CBHMDBColumnPropPage * pPage)
{
	// Save callback pointer.
	m_pColumnPage = pPage;
}

CGridInPpg::SetPagePtr(CBHMDBGroupsPropPage * pPage)
{
	// Save callback pointer.
	m_pGroupsPage = pPage;
}

BEGIN_MESSAGE_MAP(CGridInPpg, CGrid)
	//{{AFX_MSG_MAP(CGridInPpg)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// CGridInPpg message handlers
BOOL CGridInPpg::SetCurrentCell(int nRow, int nCol)
/*
Routine Description:
	Overloaded function. Navigate to new cell.

Parameters:
	nRow		The row ordinal.

	nCol		The col ordinal.

Return Value:
	If succeeded, return TRUE; Otherwise, return FALSE.
*/
{
	if (!CGrid::SetCurrentCell(nRow, nCol))
		return FALSE;

	if (m_pColumnPage)
	{
		// Get the current column in combo box.
		CString strName;
		CComboBox * pComboColumn = (CComboBox *)m_pColumnPage->GetDlgItem(IDC_COMBO_COLUMN);
		pComboColumn->GetWindowText(strName);

		if (!strName.CompareNoCase(m_arCols[nCol - 1]->strName))
			return TRUE;

		// The current col changed, select a new item in combo box.
		pComboColumn->SelectString(-1, m_arCols[nCol - 1]->strName);
		m_pColumnPage->OnSelendokComboColumn();
	}
	else if (m_pGroupsPage)
	{
		// Get current group name in combo box.
		nCol --;
		int nGroup = GetGroupFromCol(nCol);
		if (nGroup == INVALID)
			return TRUE;

		CString strName;
		CComboBox * pComboGroup = (CComboBox *)m_pGroupsPage->GetDlgItem(IDC_COMBO_GROUP);
		pComboGroup->GetWindowText(strName);

		if (!strName.CompareNoCase(m_arGroups[nGroup]->strName))
			return TRUE;

		// The current group changed, select a new item in combo box.
		pComboGroup->SelectString(-1, m_arGroups[nGroup]->strName);
		m_pGroupsPage->OnSelendokComboGroup();
	}

	return TRUE;
}

BOOL CGridInPpg::StoreCellValue(int nRow, int nCol, CString strNewValue, CString strOldValue)
/*
Routine Description:
	Overloaded function, save pending modified data.

Parameters:
	nRow		The row ordinal.

	nCol		The col ordina.

	strNewValue	The new value.

	strOldValue	The old value

Return Value:
	If succeeded, return TRUE; Otherwise, return FALSE.
*/
{
	if (!m_pColumnPage)
		return FALSE;

	if (!CGrid::StoreCellValue(nRow, nCol, strNewValue, strOldValue))
		return FALSE;

	// Mark the page as dirty.
	if (!m_pColumnPage->IsModified())
		m_pColumnPage->SetModifiedFlag(TRUE);

	return TRUE;
}

void CGridInPpg::StoreColWidth(int nCol, int nWidth)
/*
Routine Description:
	Overloaded function, save new col width.

Parameters:
	nCol		The col ordinal.

	nWidth		The new col width.

Return Value:
	None.
*/
{
	CGrid::StoreColWidth(nCol, nWidth);

	if (nCol == 0)
		return;

	// Mark page as dirty.
	if (m_pColumnPage)
		m_pColumnPage->SetModifiedFlag(TRUE);
}

BOOL CGridInPpg::DeleteRecord(int nRow)
/*
Routine Description:
	Overloaded function, delete row in grid.

Parameters:
	nRow		The row ordinal.

Return Value:
	If succeeded, return TRUE; Otherwise, return FALSE.
*/
{
	// Can not delete row in group prop page.
	if (!m_pColumnPage)
		return FALSE;

	if (!CGrid::DeleteRecord(nRow))
		return FALSE;

	// Mark page as dirty.
	m_pColumnPage->SetModifiedFlag(TRUE);

	return TRUE;
}

void CGridInPpg::StoreGroupWidth(int nGroup, int nWidth)
/*
Routine Description:
	Overloaded function, save new group width.

Parameters:
	nGroup		The group ordinal.

	nWidth		The width.

Return Value:
	None.
*/
{
	CGrid::StoreGroupWidth(nGroup, nWidth);

	CString str;
	str.Format("%d", nWidth);

	if (m_pGroupsPage)
	{
		// Update the width edit box in page.
		CString strName;
		m_pGroupsPage->GetDlgItem(IDC_COMBO_GROUP)->GetWindowText(strName);

		if (m_arGroups[nGroup]->strName.CompareNoCase(strName))
			return;

		m_pGroupsPage->GetDlgItem(IDC_EDIT_WIDTH)->SetWindowText(str);

		// Mark the page as dirty.
		m_pGroupsPage->SetModifiedFlag(TRUE);
	}
}

BOOL CGridInPpg::MoveCol(int nFromCol, int nToCol)
/*
Routine Description:
	Overloaded function, move col to new position.

Parameters:
	nFromCol		The old position.

	nToCol			The new position

Return Value:
	If succeeded, return TRUE; Otherwise, return FALSE.
*/
{
	if (!m_pColumnPage)
		return FALSE;

	if (!CGrid::MoveCol(nFromCol, nToCol))
		return FALSE;

	// Mark the page as dirty.
	if (!m_pColumnPage->IsModified())
		m_pColumnPage->SetModifiedFlag(TRUE);

	return TRUE;
}

void CGridInPpg::OnGroupHeaderClick(int nGroup)
/*
Routine Description:
	Overloaded function, toggle the selection state of one group.

Parameters:
	nGroup 		The group ordinal.

Return Value:
	If succeeded, return TRUE; Otherwise, return FALSE.
*/
{
	CGrid::OnGroupHeaderClick(nGroup);

	if (!m_pGroupsPage)
		return;

	// Select correspond item in group combo box.
	CString strName;
	CComboBox * pComboGroup = (CComboBox *)m_pGroupsPage->GetDlgItem(IDC_COMBO_GROUP);
	pComboGroup->GetWindowText(strName);

	if (m_arGroups[nGroup]->strName.CompareNoCase(strName))
		return;

	pComboGroup->SelectString(-1, m_arGroups[nGroup]->strName);
	m_pGroupsPage->OnSelendokComboGroup();
}

void CGridInPpg::StoreRowHeight(int nHeight)
/*
Routine Description:
	Overloaded function, save row height.

Parameters:
	nHeight		The row height.

Return Value:
	None.
*/
{
	CGrid::StoreRowHeight(nHeight);

	// Mark the page as dirty.
	if (m_pColumnPage)
		m_pColumnPage->SetModifiedFlag(TRUE);
	else
		m_pGroupsPage->SetModifiedFlag(TRUE);
}

void CGridInPpg::StoreHeaderHeight(int nHeight)
/*
Routine Description:
	Overloaded function, save header height.

Parameters:
	nHeight		The new height.

Return Value:
	None.
*/
{
	CGrid::StoreHeaderHeight(nHeight);

	// Mark the page as dirty.
	if (m_pColumnPage)
		m_pColumnPage->SetModifiedFlag(TRUE);
	else
		m_pGroupsPage->SetModifiedFlag(TRUE);
}
