// GridInner.cpp : implementation file
//

#include "stdafx.h"
#include "bhmdata.h"
#include "GridInner.h"
#include "BHMDBGridCtl.h"
#include "Columns.h"
#include "Column.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CGridInner

CGridInner::CGridInner(CBHMDBGridCtrl * pCtrl)
{
	m_pCtrl = pCtrl;
}

CGridInner::~CGridInner()
{
}


BEGIN_MESSAGE_MAP(CGridInner, CGrid)
	//{{AFX_MSG_MAP(CGridInner)
	ON_WM_LBUTTONDOWN()
	ON_WM_LBUTTONDBLCLK()
	ON_WM_MOUSEMOVE()
	ON_WM_MBUTTONDBLCLK()
	ON_WM_MBUTTONDOWN()
	ON_WM_RBUTTONDBLCLK()
	ON_WM_RBUTTONDOWN()
	ON_WM_LBUTTONUP()
	ON_WM_MBUTTONUP()
	ON_WM_RBUTTONUP()
	ON_WM_KILLFOCUS()
	ON_WM_SETFOCUS()
	ON_WM_HSCROLL()
	ON_WM_VSCROLL()
	ON_WM_KEYDOWN()
	ON_WM_SYSKEYDOWN()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// CGridInner message handlers
BOOL CGridInner::FlushRecord()
/*
Routine Description:
	Overloaded function, save pending modified data.

Parameters:
	None.

Return Value:
	If succeeded, return TRUE; Otherwise, return FALSE.
*/
{
	// If is not dirty, just return.
	if (!m_bRowDirty)
		return TRUE;

	// If have no callback control, use base class to process it.
	if (!m_pCtrl)
		return CGrid::FlushRecord();

	short bCancel = 0;
	
	if (m_pCtrl->m_nDataMode != 0)
	{
		// Data mode is manual mode.

		// Fire event.
		m_pCtrl->FireBeforeUpdate(&bCancel);
		if (bCancel)
			return FALSE;

		// Let base class process it.
		if (!CGrid::FlushRecord())
			return FALSE;
	}

	short bDispMsg = 0;
	if (m_pCtrl->m_nDataMode != 0)
	{
		// Data mode is manual mode.

		if (IsAddRow())
		{
			// There is pending new row, insert it into new grid.

			// Fire event.
			m_pCtrl->FireAfterInsert(&bDispMsg);
			if (bDispMsg)
				AfxMessageBox(IDS_AFTERINSERT);
		}
		else
		{
			// Fire event.
			m_pCtrl->FireAfterUpdate(&bDispMsg);
			if (bDispMsg)
				AfxMessageBox(IDS_AFTERUPDATE);
		}
		
		return TRUE;
	}
	else
	{
		// It is in bound mode.

		int nCols = GetColCount();
		m_arContent.SetSize(nCols);

		// Fill the modification information to update the data source.
		for (int i = 0; i < GetColCount(); i ++)
		{
			m_bColDirty[i] = IsColDirty(i + 1);

			m_arContent[i] = GetCell(m_nRow, i + 1)->strText;
		}

		// Save new data into data source.
		if (FAILED(m_pCtrl->UpdateRowData(m_nRow)))
			return FALSE;
		else
			return CGrid::FlushRecord();
	}

	return TRUE;
}

BOOL CGridInner::DeleteRecord(int nRow)
/*
Routine Description:
	Overloaded function. Delete current row.

Parameters:
	nRow		The row ordinal.

Return Value:
	If succeeded, return TRUE; Otherwise, return FALSE.
*/
{
	if (!m_pCtrl)
	{
		// If have no callback control, use base class to process it.
		return CGrid::DeleteRecord(nRow);
	}

	// Should allow deletion and not be read only.
	if (!m_pCtrl->m_bAllowDelete || GetReadOnly())
		return FALSE;

	if (m_pCtrl->m_nDataMode)
	{
		// It is in manual mode.
		short bCancel = 0, bDispMsg = 0;

		// Fire event.
		m_pCtrl->FireBeforeDelete(&bCancel, &bDispMsg);
		if (bCancel)
			return FALSE;
		if (bDispMsg && AfxMessageBox(IDS_BEFOREDELETE, MB_YESNO) != IDYES)
			return FALSE;

		// Let base class process it.
		if (!CGrid::DeleteRecord(nRow))
			return FALSE;

		// Fire event.
		bDispMsg = 0;
		m_pCtrl->FireAfterDelete(&bDispMsg);
		if (bDispMsg)
			AfxMessageBox(IDS_AFTERDELETE);

		return TRUE;
	}

	// Let callback control delete row from data source.
	m_pCtrl->DeleteRowData(m_nRow);

	return TRUE;
}

// catch these messages only for firing events :(
void CGridInner::OnLButtonDblClk(UINT nFlags, CPoint point) 
{
	CGrid::OnLButtonDblClk(nFlags, point);

	if (m_pCtrl)
		m_pCtrl->ButtonDblClick(LEFT_BUTTON, point);
}

void CGridInner::OnLButtonDown(UINT nFlags, CPoint point) 
{
	CGrid::OnLButtonDown(nFlags, point);

	if (m_pCtrl)
		m_pCtrl->ButtonDown(LEFT_BUTTON, point);
}

void CGridInner::OnMouseMove(UINT nFlags, CPoint point) 
{
	CGrid::OnMouseMove(nFlags, point);

	if (m_pCtrl)
		m_pCtrl->OnMouseMove(nFlags, point);
}

void CGridInner::OnMButtonDblClk(UINT nFlags, CPoint point) 
{
	CGrid::OnMButtonDblClk(nFlags, point);

	if (m_pCtrl)
		m_pCtrl->ButtonDblClick(MIDDLE_BUTTON, point);
}

void CGridInner::OnMButtonDown(UINT nFlags, CPoint point) 
{
	CGrid::OnMButtonDown(nFlags, point);

	if (m_pCtrl)
		m_pCtrl->ButtonDown(MIDDLE_BUTTON, point);
}

void CGridInner::OnRButtonDblClk(UINT nFlags, CPoint point) 
{
	CGrid::OnRButtonDblClk(nFlags, point);

	if (m_pCtrl)
		m_pCtrl->ButtonDblClick(RIGHT_BUTTON, point);
}

void CGridInner::OnRButtonDown(UINT nFlags, CPoint point) 
{
	CGrid::OnRButtonDown(nFlags, point);

	if (m_pCtrl)
		m_pCtrl->ButtonDown(RIGHT_BUTTON, point);
}

void CGridInner::OnLButtonUp(UINT nFlags, CPoint point) 
{
	CGrid::OnLButtonUp(nFlags, point);

	if (m_pCtrl)
		m_pCtrl->ButtonUp(LEFT_BUTTON, point);
}

void CGridInner::OnMButtonUp(UINT nFlags, CPoint point) 
{
	CGrid::OnMButtonUp(nFlags, point);

	if (m_pCtrl)
		m_pCtrl->ButtonUp(MIDDLE_BUTTON, point);
}

void CGridInner::OnRButtonUp(UINT nFlags, CPoint point) 
{
	CGrid::OnRButtonUp(nFlags, point);

	if (m_pCtrl)
		m_pCtrl->ButtonUp(RIGHT_BUTTON, point);
}

BOOL CGridInner::SetCurrentCell(int nRow, int nCol)
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
	int nRowCurr = m_nRow, nColCurr = m_nCol;

	// Hide dropdown window.
	m_pCtrl->HideDropDownWnd();

	short bCancel = 0;

	// Fire event.
	if (IsWindowVisible())
		m_pCtrl->FireBeforeRowColChange(&bCancel);
	if (bCancel)
		return FALSE;

	// Navigate in grid.
	if (CGrid::SetCurrentCell(nRow, nCol))
	{
		if (IsWindowVisible())
		// Fire event.
			m_pCtrl->FireRowColChange();
	}
	else
		return FALSE;

	// Let callback control sync current row position in data source.
	if (nRow != nRowCurr && m_pCtrl && IsWindowVisible())
		m_pCtrl->SetCurrentCell(nRow, nCol);

	return TRUE;
}

void CGridInner::OnClickedCellButton(int nRow, int nCol)
/*
Routine Description:
	Overloaded function. The user clicked the button in one cell.

Parameters:
	nRow		The row ordinal.

	nCol		The col ordina.

Return Value:
	None.
*/
{
	CGrid::OnClickedCellButton(nRow, nCol);

	// Inform the callback control.
	m_pCtrl->OnDropDown(nRow, nCol);

	// Fire event.
	if (m_pCtrl)
		m_pCtrl->FireBtnClick();
}

BOOL CGridInner::StoreCellValue(int nRow, int nCol, CString strNewValue, CString strOldValue)
/*
Routine Description:
	Overloaded function. Save text in current cell.
Parameters:
	None.

Return Value:
	If succeeded, return TRUE; Otherwise, return FALSE.
*/
{
	if (!m_pCtrl)
		return CGrid::StoreCellValue(nRow, nCol, strNewValue, strOldValue);

	short bCancel = 0;
	
	// Fire event.
	m_pCtrl->FireBeforeColUpdate((short)nCol, strOldValue, &bCancel);
	if (bCancel)
	{
		return FALSE;
	}

	if (!CGrid::StoreCellValue(nRow, nCol, strNewValue, strOldValue))
		return FALSE;

	// Fire event.
	m_pCtrl->FireAfterColUpdate((short)nCol);
	m_pCtrl->FireChange();

	return TRUE;
}

BOOL CGridInner::AddNew()
/*
Routine Description:
	Overloaded function. Add new row to the end.

Parameters:
	None.

Return Value:
	If succeeded, return TRUE; Otherwise, return FALSE;.
*/
{
	if (!m_pCtrl)
		return CGrid::AddNew();

	short bCancel = 0;
	short bDispMsg = 0;

	// Fire event.
	m_pCtrl->FireBeforeInsert(&bCancel);
	if (bCancel)
	{
		return FALSE;
	}

	return CGrid::AddNew();
}

BOOL CGridInner::OnStartTracking(int nRow, int nCol, int nTrackingMode)
/*
Routine Description:
	Overloaded funtion. Begin mouse tracking.

Parameters:
	nRow		The row ordinal.

	nCol		The col ordinal.

	nTrackingMode	The tracking mode.

Return Value:
	If succeeded, return TRUE; Otherwise, return FALSE.
*/
{
	if (!m_pCtrl)
		return CGrid::OnStartTracking(nRow, nCol, nTrackingMode);

	short bCancel = 0;

	// Fire event.
	if (nTrackingMode == TRACKCOLWIDTH)
		m_pCtrl->FireColResize((short)nCol, &bCancel);
	else if (nTrackingMode == TRACKROWHEIGHT)
		m_pCtrl->FireRowResize(&bCancel);

	if (bCancel)
		return FALSE;

	return CGrid::OnStartTracking(nRow, nCol, nTrackingMode);
}

BOOL CGridInner::MoveCol(int nFromCol, int nToCol)
/*
Routine Description:
	Overloaded function. Move one column to new position.

Parameters:
	nFromCol	The original position.

	nToCol		The new position.

Return Value:
	If succeeded, return TRUE; Otherwise, return FALSE.
*/
{
	short bCancel = 0;

	// Fire event.
	if (m_pCtrl)
	{
		m_pCtrl->FireColSwap((short)nFromCol, (short)nToCol, &bCancel);
		if (bCancel)
			return FALSE;
	}

	return CGrid::MoveCol(nFromCol, nToCol);
}

// events! events! events!
// update the cell content in bound mode
void CGridInner::OnLoadCellText(int nRow, int nCol, CString &strText)
{
	if (!m_pCtrl)
		return;

	CGrid::OnLoadCellText(nRow, nCol, strText);

	if (nRow > 0 && nCol > 0 && nRow <= OnGetRecordCount())
		m_pCtrl->GetCellValue(nRow, nCol, strText);
}

// cancel the modification, if this row is a "addnew" row, delete this row
// and update the bookmark array
void CGridInner::CancelRecord()
{
	if (!m_pCtrl)
	{
		CGrid::CancelRecord();
		return;
	}

	if (m_pCtrl->m_nDataMode)
	{
		CGrid::CancelRecord();
		return;
	}

	if (IsAddRow() && m_pCtrl->m_apBookmark.GetSize() >= m_nRow)
	{
		delete [] m_pCtrl->m_apBookmark[m_nRow - 1];
		m_pCtrl->m_apBookmark.RemoveAt(m_nRow - 1);
	}

	CGrid::CancelRecord();
}

int CGridInner::GetRecordCount()
/*
Routine Description:
	Get total record count.

Parameters:
	None.

Return Value:
	The count.
*/
{
	// Take the pending new row into account.
	return OnGetRecordCount() + (IsAddRow() ? 1 : 0);
}

void CGridInner::OnKillFocus(CWnd* pNewWnd) 
{
	if (m_pCtrl && !IsChild(pNewWnd))
		m_pCtrl->OnKillFocus(pNewWnd);

	CGrid::OnKillFocus(pNewWnd);
}

void CGridInner::OnSetFocus(CWnd* pOldWnd) 
{
	if (m_pCtrl)
		m_pCtrl->OnSetFocus(pOldWnd);

	CGrid::OnSetFocus(pOldWnd);
}

void CGridInner::OnHScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar) 
{
	if (!m_pCtrl)
	{
		CGrid::OnHScroll(nSBCode, nPos, pScrollBar);
		return;
	}

	short bCancel = 0;

	m_pCtrl->FireScroll(&bCancel);
	if (bCancel)
		return;

	CGrid::OnHScroll(nSBCode, nPos, pScrollBar);
	m_pCtrl->FireScrollAfter();
}

void CGridInner::OnVScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar) 
{
	if (!m_pCtrl)
	{
		CGrid::OnVScroll(nSBCode, nPos, pScrollBar);
		return;
	}

	short bCancel = 0;

	m_pCtrl->FireScroll(&bCancel);
	if (bCancel)
		return;

	CGrid::OnVScroll(nSBCode, nPos, pScrollBar);
	m_pCtrl->FireScrollAfter();
}

void CGridInner::OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags) 
{
	if (nChar == VK_F4)
	{
		// Show/hide the dropdown window.
		if (m_nRow > 0 && m_nCol > 0)
			OnClickedCellButton(m_nRow, m_nCol);

		return;
	}
	
	CGrid::OnKeyDown(nChar, nRepCnt, nFlags);
}

void CGridInner::OnSysKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags) 
{
	if (nChar == VK_DOWN || nChar == VK_UP)
	{
		// Show/hide dropdown window.
		if (m_nRow > 0 && m_nCol > 0)
			OnClickedCellButton(m_nRow, m_nCol);
		
		return;
	}
	
	CGrid::OnSysKeyDown(nChar, nRepCnt, nFlags);
}

void CGridInner::OnDrawCellText(CDC * pDC, CellStyle *pStyle)
/*
Routine Description:
	Overloaded function. Draw the text in cell.
Parameters:
	None.

Return Value:
	None.
*/
{
	HBITMAP hBmp = NULL;
	
	// Should we draw the text?
	short bDrawText = 1;
	
	if (m_pCtrl)
	{
		// Fire event to get picture handle.
		VARIANT vaPic;
		VariantInit(&vaPic);
		OLE_COLOR clrBack = (OLE_COLOR)pStyle->clrBack;
		
		m_pCtrl->FireBeforeDrawCell(pStyle->nRow, pStyle->nCol, &vaPic, &clrBack, &bDrawText);
		
		pStyle->clrBack = m_pCtrl->TranslateColor(clrBack);
		
		if (vaPic.vt == VT_I4)
			hBmp = (HBITMAP)vaPic.lVal;
	}
	
	// Draw background.
	DrawCellBackGround(pDC, pStyle);
	
	if (hBmp)
	{
		if (!bDrawText)
		// Should not draw text.
			pStyle->strText.Empty();
		
		// Get the bitmap information.
		BITMAPINFO bmpinfo;
		bmpinfo.bmiHeader.biBitCount = 0;
		bmpinfo.bmiHeader.biSize = sizeof(bmpinfo.bmiHeader);
		
		GetDIBits(pDC->m_hDC, hBmp, 0, 0, NULL, &bmpinfo, DIB_RGB_COLORS);

		// Calc the space should reserve for text.		
		CSize sz;
		sz = pDC->GetOutputTextExtent(pStyle->strText);
		
		// Create memory DC.
		HDC memdc = CreateCompatibleDC(pDC->m_hDC);
		HBITMAP hBmpOld = (HBITMAP)SelectObject(memdc, hBmp);
		
		int nShiftX, nShiftY;
		
		if (bmpinfo.bmiHeader.biHeight < pStyle->rect.Height())
		// The cell rect is high enough, put the picture center vertically.
			nShiftY = (pStyle->rect.Height() - bmpinfo.bmiHeader.biHeight) / 2;
		else
		// The cell rect is not high enough, draw from to top.
			nShiftY = 0;
		
		// The cell width is big enough to hold both the text and picture.
		if (sz.cx + bmpinfo.bmiHeader.biWidth < pStyle->rect.Width() && pStyle->strText.Find(_T("\n"), 0) == -1)
		{
			// Take the alignment into account.
			switch (pStyle->nAlignment)
			{
			case DT_CENTER:
				nShiftX = (pStyle->rect.Width() - sz.cx - bmpinfo.bmiHeader.biWidth) / 2;
				BitBlt(pDC->m_hDC, pStyle->rect.left + nShiftX, pStyle->rect.top + nShiftY, bmpinfo.bmiHeader.biWidth,
					bmpinfo.bmiHeader.biHeight, memdc, 0, 0, SRCCOPY);
				pStyle->rect.left += nShiftX + bmpinfo.bmiHeader.biWidth;
				pStyle->nAlignment = DT_LEFT;
				
				break;
				
			case DT_LEFT:
				BitBlt(pDC->m_hDC, pStyle->rect.left, pStyle->rect.top + nShiftY, bmpinfo.bmiHeader.biWidth,
					bmpinfo.bmiHeader.biHeight, memdc, 0, 0, SRCCOPY);
				pStyle->rect.left += bmpinfo.bmiHeader.biWidth;
				
				break;
				
			case DT_RIGHT:
				BitBlt(pDC->m_hDC, pStyle->rect.right - bmpinfo.bmiHeader.biWidth, pStyle->rect.top + nShiftY, 
					bmpinfo.bmiHeader.biWidth, bmpinfo.bmiHeader.biHeight, memdc, 0, 0, SRCCOPY);
				pStyle->rect.right -= bmpinfo.bmiHeader.biWidth;
				
				break;
			}
		}
		// The cell width is not big enough to hold both the text and picture
		// so put the picture at left unless the alighment is right.
		else
		{
			int nWidth = min(pStyle->rect.Width(), bmpinfo.bmiHeader.biWidth);
			
			// Take the alignment into account.
			switch (pStyle->nAlignment)
			{
			case DT_CENTER:
			case DT_LEFT:
				BitBlt(pDC->m_hDC, pStyle->rect.left, pStyle->rect.top + nShiftY, nWidth, bmpinfo.bmiHeader.biHeight, 
					memdc, 0, 0, SRCCOPY);
				pStyle->rect.left += nWidth;
				pStyle->nAlignment = DT_LEFT;
				
				break;
				
			case DT_RIGHT:
				BitBlt(pDC->m_hDC, pStyle->rect.right - nWidth, pStyle->rect.top + nShiftY, nWidth, bmpinfo.bmiHeader.biHeight, 
					memdc, 0, 0, SRCCOPY);
				pStyle->rect.right -= nWidth;
				
				break;
			}
		}

		// Clear temp objects.		
		SelectObject(memdc, hBmpOld);
		DeleteDC(memdc);
	}
}
