// ComboEdit.cpp : implementation file
//

#include "stdafx.h"
#include "bhmdata.h"
#include "ComboEdit.h"
#include "BHMDBComboCtl.h"
#include "DropDownRealWnd.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CComboEdit

CComboEdit::CComboEdit(CBHMDBComboCtrl * pCtrl)
/*
Routine Description:
	Set the callback control pointer.

Parameters:
	pCtrl		The new value.

Return Value:
	None.
*/
{
	// Save the new value.
	m_pCtrl = pCtrl;
}

CComboEdit::~CComboEdit()
{
}


BEGIN_MESSAGE_MAP(CComboEdit, CEdit)
	//{{AFX_MSG_MAP(CComboEdit)
	ON_WM_CHAR()
	ON_WM_KEYDOWN()
	ON_WM_SETFOCUS()
	ON_WM_LBUTTONDOWN()
	ON_MESSAGE_VOID(WM_PASTE, OnPaste)
	ON_MESSAGE_VOID(WM_CUT, OnCut)
	ON_MESSAGE_VOID(WM_CLEAR, OnClear)
	ON_WM_KILLFOCUS()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CComboEdit message handlers
void CComboEdit::OnChar(UINT nChar, UINT nRepCnt, UINT nFlags) 
{
	// If it's combo list or readonly, do nothing.
	if (m_pCtrl->m_nStyle || m_pCtrl->GetReadOnly())
		return;

	// Get the caret position.
	int nStart, nEnd;
	GetSel(nStart, nEnd);

	// Get current text.
	CString strText;
	GetWindowText(strText);

	if (!(m_pCtrl->m_strMask.IsEmpty()))
	{
		// There is input mask.

		// Can not insert more char.
		if (nEnd == strText.GetLength() && nChar != VK_BACK)
			return;
		
		else if (nChar == VK_BACK)
		{
			// Backspace.

			// Move the caret left.
			MoveLeft();

			// Replace current char with '_'.
			GetSel(nStart, nEnd);
			strText.SetAt(nEnd, '_');

			// Update window text.
			m_pCtrl->SetText(strText);

			// Fire event.
			m_pCtrl->FireChange();
		}
		else if (m_pCtrl->IsValidChar(nChar, nEnd))
		{
			// It is a valid char, replace current char with new char.
			MoveRight();
			strText.SetAt(nEnd, nChar);
			m_pCtrl->SetText(strText);

			// Fire event.
			m_pCtrl->FireChange();
		}
		return;
	}
	
	// let the default window procedure to update the text and caret :)
	//		DefWindowProc(WM_CHAR, nCharRet, MAKELONG(nRepCnt, nFlags));
	CEdit::OnChar(nChar, nRepCnt, nFlags);
}

void CComboEdit::OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags) 
{
	// Get caret position.
	int nStart, nEnd;
	GetSel(nStart, nEnd);

	// Get current text.
	CString strText;
	GetWindowText(strText);

	// Delete current char, replace current char with '_'.
	if (!(m_pCtrl->m_strMask.IsEmpty()) && nChar == VK_DELETE)
	{
		// If it is drop list or readonly, do nothing.
		if (m_pCtrl->m_nStyle || m_pCtrl->GetReadOnly())
			return;
		
		// The caret should in valid position.
		if (nEnd == strText.GetLength())
			return;

		// Replace current char with '_'.
		strText.SetAt(nEnd, '_');
		m_pCtrl->SetText(strText);

		// Fire event.
		m_pCtrl->FireChange();
		return;
	}
	
	int nRow;

	if (m_pCtrl->GetReadOnly())
		return;

	switch (nChar)
	{
	// Move left or right.
	case VK_LEFT:
		MoveLeft();
		break;

	case VK_RIGHT:
		MoveRight();
		break;

	// Just simulate the auto search function of normal combobox.
	case VK_UP:
		// Search current text in items.
		nRow = m_pCtrl->SearchText(strText, FALSE);

		// Locate first item default.
		if (nRow == INVALID)
			nRow = 2;
		nRow --;

		// Locate founded item.
		m_pCtrl->m_pDropDownRealWnd->MoveTo(nRow);

		break;

	case VK_DOWN:
		// Search current text in itetms.
		nRow = m_pCtrl->SearchText(strText, FALSE);

		// Locate first item default.
		if (nRow == INVALID)
			nRow = 0;
		nRow ++;

		// Locate founded item.
		m_pCtrl->m_pDropDownRealWnd->MoveTo(nRow);

		break;

	default:
		//	DefWindowProc(WM_KEYDOWN, nCharRet, MAKELONG(nRepCnt, nFlags));
		CEdit::OnKeyDown(nChar, nRepCnt, nFlags);
	}
}

void CComboEdit::OnPaste()
{
	// Do nothing for we are mask :)
}

void CComboEdit::OnClear()
{
	// Do nothing for we have mask :)
}

void CComboEdit::OnCut()
{
	// Do nothing for we have mask.
}

void CComboEdit::MoveLeft()
/*
Routine Description:
	Move the caret left, skip delimiters.

Parameters:
	None.

Return Value:
	None.
*/
{
	// Get current caret.
	int nStart, nEnd;
	GetSel(nStart, nEnd);

	if (m_pCtrl->m_strMask.IsEmpty())
	{
		// If have no mask, simply move caret left.
		SetSel(nEnd - 1, nEnd - 1);
		return;
	}

	// Skip delimiters.
	for (int i = nEnd - 1; i >= 0 && m_pCtrl->m_bIsDelimiter[i]; i--);

	if (i >= 0 && !m_pCtrl->m_bIsDelimiter[i])
		nEnd = i;

	// Update caret position.
	SetSel(nEnd, nEnd);
}

void CComboEdit::MoveRight()
/*
Routine Description:
	Move the caret right, skip delimiters.

Parameters:
	None.

Return Value:
	None.
*/
{
	int nStart, nEnd;
	CString strText;

	// Get current caret position.
	GetWindowText(strText);
	GetSel(nStart, nEnd);
	if (nEnd >= 0 && nEnd < strText.GetLength() && strText[nEnd] < 0)
		nEnd ++;

	if (m_pCtrl->m_strMask.IsEmpty())
	{
		// Simply move right if have no mask.
		SetSel(nEnd + 1, nEnd + 1);
		return;
	}

	// Skip delimiters.
	for (int i = nEnd + 1; i < m_pCtrl->m_strMask.GetLength() && 
		m_pCtrl->m_bIsDelimiter[i]; i++);

	if (!m_pCtrl->m_bIsDelimiter[i])
		nEnd = i;

	// Update the caret position.
	SetSel(nEnd, nEnd);
}

// Can not put the caret at thte delimiter.
void CComboEdit::OnSetFocus(CWnd* pOldWnd) 
{
	CEdit::OnSetFocus(pOldWnd);
	
	int nStart, nEnd;
	GetSel(nStart, nEnd);

	if (m_pCtrl->m_bIsDelimiter[nEnd])
		MoveRight();
	GetSel(nStart, nEnd);
	if (m_pCtrl->m_bIsDelimiter[nEnd])
		MoveLeft();
}

void CComboEdit::OnLButtonDown(UINT nFlags, CPoint point) 
{
	CEdit::OnLButtonDown(nFlags, point);

	int nStart, nEnd;
	GetSel(nStart, nEnd);
	if (m_pCtrl->m_bIsDelimiter[nEnd])
		MoveRight();
	GetSel(nStart, nEnd);
	if (m_pCtrl->m_bIsDelimiter[nEnd])
		MoveLeft();
}

void CComboEdit::OnKillFocus(CWnd* pNewWnd) 
{
	CEdit::OnKillFocus(pNewWnd);
	
	// Hide dropdown window when kill focus.
	m_pCtrl->HideDropDownWnd();
}
