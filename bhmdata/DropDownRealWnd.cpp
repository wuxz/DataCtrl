// DropDownRealWnd.cpp : implementation file
//

#include "stdafx.h"
#include "BHMData.h"
#include "DropDownRealWnd.h"
#include "BHMDBDropDownCtl.h"
#include "BHMDBComboCtl.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CDropDownRealWnd

CDropDownRealWnd::CDropDownRealWnd(CBHMDBDropDownCtrl * pCtrl)
/*
Routine Description:
	Set callback dropdown control pointer.

Parameters:
	pCtrl		The new value.

Return Value:
	None.
*/
{
	// Save the new value.
	m_pDropDownCtrl = pCtrl;

	// Clear other pointer.
	m_pComboCtrl = NULL;
}

CDropDownRealWnd::CDropDownRealWnd(CBHMDBComboCtrl * pCtrl)
/*
Routine Description:
	Set callback dropdown control pointer.

Parameters:
	pCtrl		The new value.

Return Value:
	None.
*/
{
	// Save the new value.
	m_pDropDownCtrl = NULL;

	// Clear other pointer.
	m_pComboCtrl = pCtrl;
}

CDropDownRealWnd::~CDropDownRealWnd()
{
}


BEGIN_MESSAGE_MAP(CDropDownRealWnd, CGrid)
	//{{AFX_MSG_MAP(CDropDownRealWnd)
	ON_WM_KILLFOCUS()
	ON_WM_KEYDOWN()
	ON_WM_HSCROLL()
	ON_WM_VSCROLL()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// CDropDownRealWnd message handlers
void CDropDownRealWnd::OnLoadCellText(int nRow, int nCol, CString &strText)
/*
Routine Description:
	Overloaded function to load cell text before drawing it.

Parameters:
	nRow		The row ordinal.

	nCol		The col ordinal.

	strText		Should save new text here.

Return Value:
	None.
*/
{
	// Let default calss to get cached text.
	CGrid::OnLoadCellText(nRow, nCol, strText);

	// If have no callback pointer, simply return.
	if (!m_pDropDownCtrl && !m_pComboCtrl)
		return;

	if (nRow > 0 && nCol > 0 && nRow <= (int)OnGetRecordCount())
	{
		// We are loading data from data source.

		// Let callback pointers process it.
		if (m_pDropDownCtrl)
			m_pDropDownCtrl->GetCellValue(nRow, nCol, strText);
		else if (m_pComboCtrl)
			m_pComboCtrl->GetCellValue(nRow, nCol, strText);
	}

	return;
}

void CDropDownRealWnd::Showwindow(BOOL bShow)
/*
Routine Description:
	Show/hide the dropdown window.

Parameters:
	bShow		Show command. TRUE for showing, FALSE for hiding.

Return Value:
	None.
*/
{
	if (m_pDropDownCtrl)
	{
		// Only work with dropdown control.

		if (bShow)
		{
			// Show the window.

			// Fire event.
			m_pDropDownCtrl->FireDropDown();

			// Show window.
			ShowWindow(SW_SHOW);

			// Mark the new state.
			m_pDropDownCtrl->m_bDroppedDown = TRUE;
		}
		else
		{
			// Hide the window.

			// Fire event.
			ShowWindow(SW_HIDE);

			// Hide window.
			m_pDropDownCtrl->m_bDroppedDown = FALSE;

			// Mark the new state.
			m_pDropDownCtrl->FireCloseUp();
		}
	}
}

void CDropDownRealWnd::OnKillFocus(CWnd* pNewWnd) 
{
	CGrid::OnKillFocus(pNewWnd);
	
	if (pNewWnd == this)
		return;
	
	// Lost focus, should hide the dropdown window.

	if (::IsWindowVisible(m_hWnd))
	{
		// The window is visible, hide it.
		if (m_pDropDownCtrl)
		{
			// Use callback control hide this window.
			m_pDropDownCtrl->HideDropDownWnd();
			
			if (::IsWindow(m_pDropDownCtrl->m_hConsumer))
			{
				// Send message to grid control this new state.
				BringInfo * pBi = new BringInfo;
				pBi->hWnd = (HWND)0xff;
				pBi->wParam = FALSE;
				::PostMessage(m_pDropDownCtrl->m_hConsumer, CBHMDataApp::m_nBringMsg, (WPARAM)pBi, NULL);
			}
		}
		else if (m_pComboCtrl)
		{
			// Use callback control hide this window.
			m_pComboCtrl->HideDropDownWnd();
		}
	}
}

void CDropDownRealWnd::OnLButtonClickedCell(int nRow, int nCol)
/*
Routine Description:
	Overloaded function when the user clicked one cell using left mouse button.

Parameters:
	nRow		The row ordinal.

	nCol		The col ordinal.

Return Value:
	None.
*/
{
	// Fire event.
	if (m_pDropDownCtrl)
		m_pDropDownCtrl->FireClick();
	else if (m_pComboCtrl)
		m_pComboCtrl->FireClick();

	if (nRow == 0 || nCol == 0)
		return;

	// Hide dropdown window, return text.
	ReturnText(nRow);
	if (m_pComboCtrl)
		m_pComboCtrl->HideDropDownWnd();

	return;
}

void CDropDownRealWnd::ReturnText(int nRow)
/*
Routine Description:
	Return text to callback control.

Parameters:
	nRow		The new current row position.

Return Value:
	None.
*/
{
	if (m_pDropDownCtrl)
	{
		// Update variable in callback control.
		m_pDropDownCtrl->m_nRow = nRow;


		// Get new text.
		CString strValue;
		
		int i = GetColFromName(m_pDropDownCtrl->m_strDataField);
		
		if (i != INVALID)
		{
			strValue = GetCellText(nRow, i);

			// Return text to dropdown control and hide this window.
			m_pDropDownCtrl->HideDropDownWnd(strValue);
		}
	}
	else if (m_pComboCtrl)
	{
		// Update current row variable in control.
		m_pComboCtrl->m_nRow = nRow;

		// Get new text.
		CString strValue;
		
		int i = GetColFromName(m_pComboCtrl->m_strDataField);
		
		if (i != INVALID)
		{
			strValue = GetCellText(nRow, i);

			// Return new text.
			m_pComboCtrl->SetText(strValue);
		}
	}
}

BOOL CDropDownRealWnd::SetCurrentCell(int nRow, int nCol)
/*
Routine Description:
	Overloaded function when navigate to a new cell.

Parameters:
	nRow		The row ordinal.

	nCol		The col ordinal.

Return Value:
	None.
*/
{
	// Left base class process it first.
	if (!CGrid::SetCurrentCell(nRow, nCol))
		return FALSE;

	if (nRow > 0 && m_pComboCtrl)
	{
		// If the callback control is combo control, update its text with new cell text.
		ReturnText(nRow);
	}

	return TRUE;
}

void CDropDownRealWnd::OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags) 
{
	switch (nChar)
	{
	case VK_RETURN:
	{
		// Return text and hide this window.
		int nRow, nCol;
			
		GetCurrentCell(nRow, nCol);
		ReturnText(nRow);
	}
		
	case VK_ESCAPE:
	{
		// Hide this window.
		if (m_pDropDownCtrl)
			m_pDropDownCtrl->HideDropDownWnd();
		else if (m_pComboCtrl)
			m_pComboCtrl->HideDropDownWnd();
			
		return;
	}
	}
	
	CGrid::OnKeyDown(nChar, nRepCnt, nFlags);
}

void CDropDownRealWnd::OnHScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar) 
/*
Routine Description:
	Overloaded function when receive horizontal scroll message.

Parameters:
	nSBCode

	nPos

	pScrollBar

Return Value:
	None.
*/
{
	short bCancel = 0;

	// Fire event.
	if (m_pDropDownCtrl)
		m_pDropDownCtrl->FireScroll(&bCancel);
	else
		m_pComboCtrl->FireScroll(&bCancel);

	if (bCancel)
		return;

	// Scroll.
	CGrid::OnHScroll(nSBCode, nPos, pScrollBar);

	// Fire event.
	if (m_pDropDownCtrl)
		m_pDropDownCtrl->FireScrollAfter();
	else
		m_pComboCtrl->FireScrollAfter();
}

void CDropDownRealWnd::OnVScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar) 
/*
Routine Description:
	Overloaded function when receive vertical scroll message.

Parameters:
	nSBCode

	nPos

	pScrollBar

Return Value:
	None.
*/
{
	// Fire event.
	short bCancel = 0;

	if (m_pDropDownCtrl)
		m_pDropDownCtrl->FireScroll(&bCancel);
	else
		m_pComboCtrl->FireScroll(&bCancel);

	if (bCancel)
		return;

	// Scroll.
	CGrid::OnVScroll(nSBCode, nPos, pScrollBar);

	// Fire event.
	if (m_pDropDownCtrl)
		m_pDropDownCtrl->FireScrollAfter();
	else
		m_pComboCtrl->FireScrollAfter();
}
