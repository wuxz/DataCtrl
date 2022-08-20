// Column.cpp : implementation file
//

#include "stdafx.h"
#include "BHMData.h"
#include "Column.h"
#include "BHMDBGridCtl.h"
#include "BHMDBDropDownCtl.h"
#include "BHMDBComboCtl.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

// Validate the given col index.
#define VERIFYCOLINDEX \
	int nCol = m_pGrid->GetColFromIndex(prop.nColIndex); \
	if (nCol == INVALID) \
		Error(CTL_E_INVALIDPROPERTYARRAYINDEX, AFX_IDP_E_INVALIDPROPERTYARRAYINDEX);

/////////////////////////////////////////////////////////////////////////////
// CColumn

IMPLEMENT_DYNCREATE(CColumn, CCmdTarget)

IMPLEMENT_OLETYPELIB(CColumn, _tlid, _wVerMajor, _wVerMinor)

CColumn::CColumn()
{
	EnableAutomation();
	EnableTypeLib();

	m_pGridCtrl = NULL;
	m_pDropDownCtrl = NULL;
	m_pComboCtrl = NULL;
	m_pGrid = NULL;
}

CColumn::SetCtrlPtr(CBHMDBGridCtrl * pCtrl)
/*
Routine Description:
	Set call back pointer of grid control.

Parameters:
	pCtrl		The new value.

Return Value:
	None.
*/
{
	// Save the new value.
	m_pGridCtrl = pCtrl;
}

CColumn::SetCtrlPtr(CBHMDBDropDownCtrl * pCtrl)
/*
Routine Description:
	Set call back pointer of dropdown control.

Parameters:
	pCtrl		The new value.

Return Value:
	None.
*/
{
	// Save the new value.
	m_pDropDownCtrl = pCtrl;
}

CColumn::SetCtrlPtr(CBHMDBComboCtrl * pCtrl)
/*
Routine Description:
	Set call back pointer of combo control.

Parameters:
	pCtrl		The new value.

Return Value:
	None.
*/
{
	m_pComboCtrl = pCtrl;
}

CColumn::~CColumn()
{
}


void CColumn::OnFinalRelease()
{
	// When the last reference for an automation object is released
	// OnFinalRelease is called.  The base class will automatically
	// deletes the object.  Add additional cleanup required for your
	// object before calling the base class.

	CCmdTarget::OnFinalRelease();
}


BEGIN_MESSAGE_MAP(CColumn, CCmdTarget)
	//{{AFX_MSG_MAP(CColumn)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(CColumn, CCmdTarget)
	//{{AFX_DISPATCH_MAP(CColumn)
	DISP_PROPERTY_EX(CColumn, "FieldLen", GetFieldLen, SetFieldLen, VT_I2)
	DISP_PROPERTY_EX(CColumn, "AllowSizing", GetAllowSizing, SetAllowSizing, VT_BOOL)
	DISP_PROPERTY_EX(CColumn, "HeadForeColor", GetHeadForeColor, SetHeadForeColor, VT_COLOR)
	DISP_PROPERTY_EX(CColumn, "HeadBackColor", GetHeadBackColor, SetHeadBackColor, VT_COLOR)
	DISP_PROPERTY_EX(CColumn, "DataField", GetDataField, SetDataField, VT_BSTR)
	DISP_PROPERTY_EX(CColumn, "Locked", GetLocked, SetLocked, VT_BOOL)
	DISP_PROPERTY_EX(CColumn, "Visible", GetVisible, SetVisible, VT_BOOL)
	DISP_PROPERTY_EX(CColumn, "Width", GetWidth, SetWidth, VT_I4)
	DISP_PROPERTY_EX(CColumn, "DataType", GetDataType, SetDataType, VT_I4)
	DISP_PROPERTY_EX(CColumn, "Selected", GetSelected, SetSelected, VT_BOOL)
	DISP_PROPERTY_EX(CColumn, "Style", GetStyle, SetStyle, VT_I4)
	DISP_PROPERTY_EX(CColumn, "CaptionAlignment", GetCaptionAlignment, SetCaptionAlignment, VT_I4)
	DISP_PROPERTY_EX(CColumn, "ColChanged", GetColChanged, SetNotSupported, VT_BOOL)
	DISP_PROPERTY_EX(CColumn, "Name", GetName, SetName, VT_BSTR)
	DISP_PROPERTY_EX(CColumn, "Nullable", GetNullable, SetNullable, VT_BOOL)
	DISP_PROPERTY_EX(CColumn, "Mask", GetMask, SetMask, VT_BSTR)
	DISP_PROPERTY_EX(CColumn, "PromptChar", GetPromptChar, SetPromptChar, VT_BSTR)
	DISP_PROPERTY_EX(CColumn, "Caption", GetCaption, SetCaption, VT_BSTR)
	DISP_PROPERTY_EX(CColumn, "Alignment", GetAlignment, SetAlignment, VT_I4)
	DISP_PROPERTY_EX(CColumn, "ForeColor", GetForeColor, SetForeColor, VT_COLOR)
	DISP_PROPERTY_EX(CColumn, "BackColor", GetBackColor, SetBackColor, VT_COLOR)
	DISP_PROPERTY_EX(CColumn, "Case", GetCase, SetCase, VT_I4)
	DISP_PROPERTY_EX(CColumn, "Text", GetText, SetText, VT_BSTR)
	DISP_PROPERTY_EX(CColumn, "Value", GetValue, SetValue, VT_VARIANT)
	DISP_PROPERTY_EX(CColumn, "PromptInclude", GetPromptInclude, SetPromptInclude, VT_BOOL)
	DISP_PROPERTY_EX(CColumn, "ListCount", GetListCount, SetNotSupported, VT_I2)
	DISP_PROPERTY_EX(CColumn, "ListIndex", GetListIndex, SetListIndex, VT_I2)
	DISP_PROPERTY_EX(CColumn, "DropDownHWnd", GetDropDownHWnd, SetDropDownHWnd, VT_HANDLE)
	DISP_FUNCTION(CColumn, "CellText", CellText, VT_BSTR, VTS_VARIANT)
	DISP_FUNCTION(CColumn, "CellValue", CellValue, VT_VARIANT, VTS_VARIANT)
	DISP_FUNCTION(CColumn, "IsCellValid", IsCellValid, VT_BOOL, VTS_NONE)
	DISP_FUNCTION(CColumn, "AddItem", AddItem, VT_EMPTY, VTS_BSTR VTS_VARIANT)
	DISP_FUNCTION(CColumn, "RemoveAll", RemoveAll, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CColumn, "RemoveItem", RemoveItem, VT_EMPTY, VTS_I2)
	DISP_PROPERTY_PARAM(CColumn, "List", GetList, SetList, VT_BSTR, VTS_I2)
	DISP_DEFVALUE(CColumn, "Text")
	//}}AFX_DISPATCH_MAP
END_DISPATCH_MAP()

// Note: we add support for IID_IColumn to support typesafe binding
//  from VBA.  This IID must match the GUID that is attached to the 
//  dispinterface in the .ODL file.

// {BF91B154-6A84-11D3-A7FE-0080C8763FA4}
static const IID IID_IColumn =
{ 0xbf91b154, 0x6a84, 0x11d3, { 0xa7, 0xfe, 0x0, 0x80, 0xc8, 0x76, 0x3f, 0xa4 } };

BEGIN_INTERFACE_MAP(CColumn, CCmdTarget)
	INTERFACE_PART(CColumn, IID_IColumn, Dispatch)
END_INTERFACE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CColumn message handlers
BSTR CColumn::GetCaption() 
/*
Routine Description:
	Get the caption.

Parameters:
	None.

Return Value:
	The caption.
*/
{
	VERIFYCOLINDEX;

	return m_pGrid->GetColTitle(nCol).AllocSysString();
}

void CColumn::SetCaption(LPCTSTR lpszNewValue) 
/*
Routine Description:
	Set the caption.
Parameters:
	lpszNewValue		The new value.

Return Value:
	None.
*/
{
	int nCol = m_pGrid->GetColFromIndex(prop.nColIndex);

	// Update the grid.
	m_pGrid->SetColTitle(nCol, lpszNewValue);
}

long CColumn::GetAlignment() 
/*
Routine Description:
	Get the alignment.

Parameters:
	None.

Return Value:
	The alignment.
*/
{
	VERIFYCOLINDEX;

	// Convert the property of grid to our own format.
	switch (m_pGrid->GetColAlignment(nCol))
	{
	case DT_CENTER:
		return 2;
		
	case DT_LEFT:
		return 0;

	default:
		return 1;
	}
}

void CColumn::SetAlignment(long nNewValue) 
/*
Routine Description:
	Set the alignment.

Parameters:
	nNewValue		The new value.

Return Value:
	None.
*/
{
	// Can only be 0 to 2.
	if (nNewValue < 0 || nNewValue > 2)
		return;

	VERIFYCOLINDEX;

	DWORD dwAlign;

	// Convert the value to the format of the grid.
	switch (nNewValue)
	{
	case 0:
		dwAlign = DT_LEFT;
		break;

	case 1:
		dwAlign = DT_RIGHT;
		break;

	case 2:
		dwAlign = DT_CENTER;
		break;
	}

	// Update the grid.
	m_pGrid->SetColAlignment(nCol, dwAlign);
}

OLE_COLOR CColumn::GetForeColor() 
/*
Routine Description:
	Get the fore color.

Parameters:
	None.

Return Value:
	The color.
*/
{
	VERIFYCOLINDEX;

	return m_pGrid->GetColForeColor(nCol);
}

void CColumn::SetForeColor(OLE_COLOR nNewValue) 
/*
Routine Description:
	Set the fore color.

Parameters:
	nNewValue		The new value.

Return Value:
	None.
*/
{
	VERIFYCOLINDEX;

	// Update the grid.
	m_pGrid->SetColForeColor(nCol, TranslateColor(nNewValue));
}

OLE_COLOR CColumn::GetBackColor() 
/*
Routine Description:
	Get the back color.

Parameters:
	None.

Return Value:
	The color.
*/
{
	VERIFYCOLINDEX;

	return m_pGrid->GetColBackColor(nCol);
}

void CColumn::SetBackColor(OLE_COLOR nNewValue) 
/*
Routine Description:
	Set the back color.

Parameters:
	nNewValue		The new value.

Return Value:
	None.
*/
{
	VERIFYCOLINDEX;

	// Update the grid.
	m_pGrid->SetColBackColor(nCol, TranslateColor(nNewValue));
}

long CColumn::GetCase() 
/*
Routine Description:
	Get the text case.

Parameters:
	None.

Return Value:
	The case.
*/
{
	VERIFYCOLINDEX;

	return m_pGrid->GetColCase(nCol);
}

void CColumn::SetCase(long nNewValue) 
/*
Routine Description:
	Set the text case.

Parameters:
	nNewValue		The new value.

Return Value:
	None.
*/
{
	VERIFYCOLINDEX;

	// Update the grid.
	m_pGrid->SetColCase(nCol, nNewValue);
}

short CColumn::GetFieldLen() 
/*
Routine Description:
	Get the max length of text.

Parameters:
	None.

Return Value:
	The max length.
*/
{
	// Only be support in grid control.
	if (!m_pGridCtrl)
		GetNotSupported();

	VERIFYCOLINDEX;

	return m_pGrid->GetColMaxLength(nCol);
}

void CColumn::SetFieldLen(short nNewValue) 
/*
Routine Description:
	Set the max text length.

Parameters:
	nNewValue		The new value.

Return Value:
	None.
*/
{
	// Only be supported in grid control.
	if (!m_pGridCtrl)
		SetNotSupported();

	VERIFYCOLINDEX;

	// Update the grid.
	m_pGrid->SetColMaxLength(nCol, nNewValue);
}

BOOL CColumn::GetAllowSizing() 
/*
Routine Description:
	Get the permission to resize col width manually.

Parameters:
	None.

Return Value:
	The permission.
*/
{
	VERIFYCOLINDEX;

	return m_pGrid->GetColAllowSizing(nCol);
}

void CColumn::SetAllowSizing(BOOL bNewValue) 
/*
Routine Description:
	Set the permission to resize col width manually.

Parameters:
	bNewValue		The new value.

Return Value:
	None.
*/
{
	VERIFYCOLINDEX;

	// Update the grid.
	m_pGrid->SetColAllowSizing(nCol, bNewValue);
}

OLE_COLOR CColumn::GetHeadForeColor() 
/*
Routine Description:
	Get the header fore color.

Parameters:
	None.

Return Value:
	The color.
*/
{
	VERIFYCOLINDEX;

	return m_pGrid->GetColHeaderForeColor(nCol);
}

void CColumn::SetHeadForeColor(OLE_COLOR nNewValue) 
/*
Routine Description:
	Set the header fore color.

Parameters:
	nNewValue		The new value.

Return Value:
	None.
*/
{
	VERIFYCOLINDEX;

	// Update thte grid.
	m_pGrid->SetColHeaderForeColor(nCol, TranslateColor(nNewValue));
}

OLE_COLOR CColumn::GetHeadBackColor() 
/*
Routine Description:
	Get the header back color.

Parameters:
	None.

Return Value:
	The color.
*/
{
	VERIFYCOLINDEX;

	return m_pGrid->GetColHeaderBackColor(nCol);
}

void CColumn::SetHeadBackColor(OLE_COLOR nNewValue) 
/*
Routine Description:
	Set the head back color.

Parameters:
	nNewValue		The new value.

Return Value:
	None.
*/
{
	VERIFYCOLINDEX;

	// Update the grid.
	m_pGrid->SetColHeaderBackColor(nCol, TranslateColor(nNewValue));
}

BSTR CColumn::GetDataField() 
/*
Routine Description:
	Get the binding data field.

Parameters:
	None.

Return Value:
	The field name.
*/
{
	VERIFYCOLINDEX;

	CString strResult;

	if (m_pGrid->GetColUserAttribRef(nCol).GetSize() >= 1)
		strResult = m_pGrid->GetColUserAttribRef(nCol).ElementAt(0);

	return strResult.AllocSysString();
}

BOOL CColumn::GetLocked() 
/*
Routine Description:
	Get the permission to modify cell text manually.

Parameters:
	None.

Return Value:
	The permission.
*/
{
	if (!m_pGridCtrl)
		GetNotSupported();

	VERIFYCOLINDEX;

	return m_pGrid->GetColReadOnly(nCol);
}

void CColumn::SetLocked(BOOL bNewValue) 
/*
Routine Description:
	Set the permission to modify cell text manually.

Parameters:
	bNewValue		The new value.

Return Value:
	None.
*/
{
	// Only be supporte in grid control.
	if (!m_pGridCtrl)
		SetNotSupported();

	VERIFYCOLINDEX;

	CString strField;

	// Update the grid.
	if (m_pGrid->GetColUserAttribRef(nCol).GetSize() >= 1)
		strField = m_pGrid->GetColUserAttribRef(nCol).ElementAt(0);

	if (!IsColForceLock(strField))
		m_pGrid->SetColReadOnly(nCol, bNewValue);
}

BOOL CColumn::GetVisible() 
/*
Routine Description:
	Get the visibility.

Parameters:
	None.

Return Value:
	The visibility.
*/
{
	VERIFYCOLINDEX;

	return m_pGrid->GetColVisible(nCol);
}

void CColumn::SetVisible(BOOL bNewValue) 
/*
Routine Description:
	Set the visibility.

Parameters:
	bNewValue		The new value.

Return Value:
	None.
*/
{
	VERIFYCOLINDEX;

	// Update the grid.
	m_pGrid->SetColVisible(nCol, bNewValue);
}

long CColumn::GetWidth() 
/*
Routine Description:
	Get the width.

Parameters:
	None.

Return Value:
	The width.
*/
{
	VERIFYCOLINDEX;

	return m_pGrid->GetColWidth(nCol);
}

void CColumn::SetWidth(long nNewValue) 
/*
Routine Description:
	Set the width.

Parameters:
	nNewValue		The new value.

Return Value:
	None.
*/
{
	VERIFYCOLINDEX;

	// Update the grid.
	m_pGrid->SetColWidth(nCol, nNewValue);
}

long CColumn::GetDataType() 
/*
Routine Description:
	Get the type of underlying data.

Parameters:
	None.

Return Value:
	The data type.
*/
{
	// Only be supported in grid control.
	if (!m_pGridCtrl)
		GetNotSupported();

	VERIFYCOLINDEX;

	return m_pGrid->GetColDataType(nCol);
}

void CColumn::SetDataField(LPCTSTR lpszNewValue) 
/*
Routine Description:

Parameters:
	None.

Return Value:
	None.
*/
{
	VERIFYCOLINDEX;

	// Not be supported in manual mode.
	if (GetDataMode())
		Error(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_NOAPPLICABLEINMANUALMODE);

	m_pGrid->GetColUserAttribRef(nCol).SetAtGrow(0, lpszNewValue);
	
	// Update the readonly and nullable properties.
	if (IsColForceLock(lpszNewValue))
		m_pGrid->SetColReadOnly(nCol, TRUE);
	if (IsColForceNotNullable(lpszNewValue))
		SetNullable(TRUE);

	// Rebind to data source
	if (m_pGridCtrl)
		m_pGridCtrl->Bind();
	else if (m_pDropDownCtrl)
		m_pDropDownCtrl->Bind();
	else
		m_pComboCtrl->Bind();

	// Update the grid.
	m_pGrid->RedrawCol(nCol + 1);
}

void CColumn::SetDataType(long nNewValue) 
/*
Routine Description:
	Set the type of underlying data.

Parameters:
	nNewValue		The new value.

Return Value:
	None.
*/
{
	// Only be supported in grid control.
	if (!m_pGridCtrl)
		SetNotSupported();

	// Validate the new value.
	if (nNewValue < 0 || nNewValue > 7)
		return;

	VERIFYCOLINDEX;

	// Update the grid.
	m_pGrid->SetColDataType(nCol, nNewValue);
}

BOOL CColumn::GetSelected() 
/*
Routine Description:
	Get if this col is selected.

Parameters:
	None.

Return Value:
	If is selected, return TRUE; Otherwise, return FALSE.
*/
{
	// Only be supported in grid control.
	if (!m_pGridCtrl)
		GetNotSupported();

	VERIFYCOLINDEX;

	return m_pGrid->IsColSelected(nCol);
}

void CColumn::SetSelected(BOOL bNewValue) 
/*
Routine Description:
	Set if this col is selected.

Parameters:
	bNewValue		The new value.

Return Value:
	None.
*/
{
	// Only be supported in grid control.
	if (!m_pGridCtrl)
		SetNotSupported();

	VERIFYCOLINDEX;

	// Update the grid.
	m_pGrid->SelectCol(nCol, bNewValue);
}

long CColumn::GetStyle() 
/*
Routine Description:
	Get the style.

Parameters:
	None.

Return Value:
	The style.
*/
{
	// Only be supported in grid control.
	if (!m_pGridCtrl)
		GetNotSupported();

	VERIFYCOLINDEX;

	return m_pGrid->GetColControl(nCol);
}

void CColumn::SetStyle(long nNewValue)
/*
Routine Description:
	Set the style.

Parameters:
	nNewValue		The new value.

Return Value:
	None.
*/
{
	// Only be supported in grid control.
	if (!m_pGridCtrl)
		SetNotSupported();

	// Validate the new value.
	if (nNewValue < 0 || nNewValue > 7)
		return;

	VERIFYCOLINDEX;

	// Update the grid.
	m_pGrid->SetColControl(nCol, nNewValue);
}

long CColumn::GetCaptionAlignment() 
/*
Routine Description:
	Get the caption alignment.

Parameters:
	None.

Return Value:
	The alignment.
*/
{
	VERIFYCOLINDEX;

	// Convert the gotten value into our own format.
	switch (m_pGrid->GetColHeaderAlignment(nCol))
	{
	case DT_LEFT:
		return 0;

	case DT_RIGHT:
		return 1;

	case DT_CENTER:
		return 2;

	default:
		return 3;
	}
}

void CColumn::SetCaptionAlignment(long nNewValue) 
/*
Routine Description:
	Set caption alignment.

Parameters:
	nNewValue		The new value.

Return Value:
	None.
*/
{
	VERIFYCOLINDEX;

	DWORD dwAlign;

	// Convert into the format of grid.
	switch (nNewValue)
	{
	case 0:
		dwAlign = DT_LEFT;
		break;

	case 1:
		dwAlign = DT_RIGHT;
		break;

	case 2:
		dwAlign = DT_CENTER;
		break;

	case 3:
		dwAlign = -1;
		break;
	}

	// Update the grid.
	m_pGrid->SetColHeaderAlignment(nCol, dwAlign);
}

BOOL CColumn::GetColChanged() 
/*
Routine Description:
	Get if the text in current cell is been modified.

Parameters:
	None.

Return Value:
	If is modified, return TRUE; Otherwise, return FALSE.
*/
{
	// Only be supported in grid control.
	if (!m_pGridCtrl)
		GetNotSupported();

	VERIFYCOLINDEX;

	return m_pGrid->IsColDirty(nCol);
}

BSTR CColumn::GetName() 
/*
Routine Description:
	Get the col name.

Parameters:
	None.

Return Value:
	The name.
*/
{
	VERIFYCOLINDEX;

	return m_pGrid->GetColName(nCol).AllocSysString();
}

void CColumn::SetName(LPCTSTR lpszNewValue) 
/*
Routine Description:
	Set the col name.

Parameters:
	lpszNewValue		The new value.

Return Value:
	None.
*/
{
	VERIFYCOLINDEX;

	// Update the grid.
	if (!m_pGrid->SetColName(nCol, lpszNewValue))
		Error(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_NAMENOTUNIQUE);
}

BOOL CColumn::GetNullable() 
/*
Routine Description:
	Get if can use null value in binding field.

Parameters:
	None.

Return Value:
	The permission.
*/
{
	// Only be supported in grid control.
	if (!m_pGridCtrl)
		GetNotSupported();

	VERIFYCOLINDEX;

	CString strResult;

	if (m_pGrid->GetColUserAttribRef(nCol).GetSize() >= 2)
		strResult = m_pGrid->GetColUserAttribRef(nCol).ElementAt(1);

	return strResult.Compare(_T("T"));
}

void CColumn::SetNullable(BOOL bNewValue) 
/*
Routine Description:
	Set the permission to use null value in binding field.

Parameters:
	bNewValue		The new value.

Return Value:
	None.
*/
{
	// Only be supported in grid control.
	if (!m_pGridCtrl)
		SetNotSupported();

	// Not support in manual mode.
	if (GetDataMode())
		Error(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_NOAPPLICABLEINMANUALMODE);

	VERIFYCOLINDEX;

	CString strField;

	if (m_pGrid->GetColUserAttribRef(nCol).GetSize() >= 1)
		strField = m_pGrid->GetColUserAttribRef(nCol).ElementAt(0);

	if (!IsColForceNotNullable(strField))
		m_pGrid->GetColUserAttribRef(nCol).SetAtGrow(1, bNewValue ? _T("T") : _T(""));
}

BSTR CColumn::GetMask() 
/*
Routine Description:
	Get the input mask.

Parameters:
	None.

Return Value:
	The mask string.
*/
{
	// Only be supported in grid control.
	if (!m_pGridCtrl)
		GetNotSupported();

	VERIFYCOLINDEX;

	return m_pGrid->GetColMask(nCol).AllocSysString();
}

void CColumn::SetMask(LPCTSTR lpszNewValue) 
/*
Routine Description:
	Set the input mask.

Parameters:
	lpszNewValue		The new value.

Return Value:
	None.
*/
{
	// Only be supported in grid control.
	if (!m_pGridCtrl)
		SetNotSupported();

	VERIFYCOLINDEX;

	// Update the grid.
	m_pGrid->SetColMask(nCol, lpszNewValue);
}

BSTR CColumn::GetPromptChar() 
/*
Routine Description:
	Get the prompt char.

Parameters:
	None.

Return Value:
	The string.
*/
{
	// Only be supported in grid control.
	if (m_pGridCtrl)
		GetNotSupported();

	VERIFYCOLINDEX;

	return m_pGrid->GetColPromptChar(nCol).AllocSysString();
}

void CColumn::SetPromptChar(LPCTSTR lpszNewValue) 
/*
Routine Description:
	Set the prompt char.

Parameters:
	lpszNewValue		The new value.

Return Value:
	None.
*/
{
	// Only be supported in grid control.
	if (!m_pGridCtrl)
		SetNotSupported();

	VERIFYCOLINDEX;

	// Update the grid.
	m_pGrid->SetColPromptChar(nCol, lpszNewValue);
}

BOOL CColumn::GetDispatchIID(IID *pIID)
{
	*pIID = IID_IColumn;

	return TRUE;
}

BSTR CColumn::GetText() 
/*
Routine Description:
	Get the text of current cell.

Parameters:
	None.

Return Value:
	The text.
*/
{
	// Get current cell.
	int nRow, nCol;
	m_pGrid->GetCurrentCell(nRow, nCol);

	return GetCellText(nRow);
}

void CColumn::SetText(LPCTSTR lpszNewValue) 
/*
Routine Description:
	Set the text in current cell.

Parameters:
	lpszNewValue		The new value.

Return Value:
	None.
*/
{
	// Only be supported in grid control.
	if (!m_pGridCtrl)
		SetNotSupported();

	VERIFYCOLINDEX;

	// Get current cell.
	int nRow, nCurrCol;
	m_pGrid->GetCurrentCell(nRow, nCurrCol);

	// Update the grid.
	m_pGrid->SetCellText(nRow, nCol, lpszNewValue);
}

VARIANT CColumn::GetValue() 
/*
Routine Description:
	Get the value in current cell.

Parameters:
	None.

Return Value:
	The value.
*/
{
	// Only be supported in grid control.
	if (!m_pGridCtrl)
		GetNotSupported();

	// Get current cell.
	int nRow, nCol;
	m_pGrid->GetCurrentCell(nRow, nCol);

	return GetCellValue(nRow);
}

void CColumn::SetValue(const VARIANT FAR& newValue) 
/*
Routine Description:
	Set the value in current cell.

Parameters:
	newValue		The new value.

Return Value:
	None.
*/
{
	// Only be supported in grid control.
	if (!m_pGridCtrl)
		SetNotSupported();

	COleVariant vaResult = newValue;
	vaResult.ChangeType(VT_BSTR);
	CString strValue = vaResult.bstrVal;

	SetText(strValue);
}

BSTR CColumn::CellText(const VARIANT FAR& Bookmark)
/*
Routine Description:
	Get the text in specified row.

Parameters:
	Bookmark		The desired bookmark.

Return Value:
	The Value.
*/
{
	if (m_pGridCtrl)
		return GetCellText(m_pGridCtrl->GetRowFromBmk(&Bookmark));
	else if (m_pDropDownCtrl)
		return GetCellText(m_pDropDownCtrl->GetRowFromBmk(&Bookmark));
	else
		return GetCellText(m_pComboCtrl->GetRowFromBmk(&Bookmark));
}

VARIANT CColumn::CellValue(const VARIANT FAR& Bookmark)
/*
Routine Description:
	Get the value at specified bookmark.

Parameters:
	Bookmark		The desired bookmark.

Return Value:
	None.
*/
{
	// Only be supported in grid control.
	if (!m_pGridCtrl)
		GetNotSupported();

	return GetCellValue(m_pGridCtrl->GetRowFromBmk(&Bookmark));
}

BSTR CColumn::GetCellText(long RowIndex) 
/*
Routine Description:
	Get the cell text in specified row.

Parameters:
	RowIndex		The desired row.

Return Value:
	The string.
*/
{
	VERIFYCOLINDEX;

	// Validate the given row ordinal.
	if (RowIndex > m_pGrid->GetRowCount() || RowIndex < 1)
		Error(CTL_E_INVALIDPROPERTYARRAYINDEX, AFX_IDP_E_INVALIDPROPERTYARRAYINDEX);

	// Get the text.
	CString strResult = m_pGrid->GetCellText(RowIndex, nCol);

	// Take the text case into account.
	int nCase = m_pGrid->GetColCase(nCol);
	if (nCase == CASE_UPPER)
		strResult.MakeUpper();
	else if (nCase == CASE_LOWER)
		strResult.MakeLower();

	return strResult.AllocSysString();
}

VARIANT CColumn::GetCellValue(long RowIndex) 
/*
Routine Description:
	Get the cell value in specified row.

Parameters:
	RowIndex		The desired row.

Return Value:
	The value.
*/
{
	VERIFYCOLINDEX;

	COleVariant vaResult;

	// Get cell text from grid.
	vaResult = m_pGrid->GetCellText(RowIndex, nCol);
	
	int nDataType = m_pGrid->GetColDataType(nCol);
	CString strMask = m_pGrid->GetColMask(nCol);

	// If the date mask is mingguo format, the value will be text type
	if (nDataType != DataTypeDate || strMask.Find(_T("ggg")) == -1)
		vaResult.ChangeType(arrVt[nDataType]);

	return vaResult;
}

BOOL CColumn::IsCellValid() 
/*
Routine Description:
	Get if the text in current cell is valid.

Parameters:
	None.

Return Value:
	Allways return TRUE :)
*/
{
	// Only be supported in grid control.
	if (!m_pGridCtrl)
		GetNotSupported();

	return TRUE;
}

BOOL CColumn::GetPromptInclude() 
/*
Routine Description:
	Get if the prompt char should be included in text.

Parameters:
	None.

Return Value:
	The permission.
*/
{
	// Only be supported in grid control.
	if (m_pGridCtrl)
		GetNotSupported();

	VERIFYCOLINDEX;

	return m_pGrid->GetColPromptInclude(nCol);
}

void CColumn::SetPromptInclude(BOOL bNewValue) 
/*
Routine Description:
	Set if the prompt char should be included in text.

Parameters:
	bNewValue		The new value.

Return Value:
	None.
*/
{
	// Only be supported in grid control.
	if (m_pGridCtrl)
		SetNotSupported();

	VERIFYCOLINDEX;

	// Update the grid.
	m_pGrid->SetColPromptInclude(nCol, bNewValue);
}

short CColumn::GetListCount() 
/*
Routine Description:
	Get the item count in combo box.

Parameters:
	None.

Return Value:
	The count.
*/
{
	// Only be supported in grid control.
	if (!m_pGridCtrl)
		GetNotSupported();

	VERIFYCOLINDEX;

	// Recalc the choice list.
	CalcChoiceList(m_pGrid->GetColChoiceList(nCol));

	return m_arChoiceList.GetSize();
}

short CColumn::GetListIndex() 
/*
Routine Description:
	Get the index of current cell text in combo box.

Parameters:
	None.

Return Value:
	The index.
*/
{
	// Only be supported in grid control.
	if (!m_pGridCtrl)
		GetNotSupported();

	VERIFYCOLINDEX;

	// Only be significant in combo box or combo list style col.
	int nStyle = m_pGrid->GetColControl(nCol);
	if (nStyle != COLSTYLE_COMBOBOX && nStyle != COLSTYLE_COMBOBOX_DROPLIST)
		return CB_ERR;

	int nRow, nColCurr;

	m_pGrid->GetCurrentCell(nRow, nColCurr);
	if (nRow > m_pGrid->GetRowCount() || nRow < 1)
		return CB_ERR;

	// Get current cell value.
	CString strValue = m_pGrid->GetCellText(nRow, nCol);

	// Recalc the choice list.
	CalcChoiceList(m_pGrid->GetColChoiceList(nCol));

	// Get the index.
	for (int i = 0; i < m_arChoiceList.GetSize(); i ++)
		if (!m_arChoiceList[i].CompareNoCase(strValue))
			return i;

	return CB_ERR;
}

void CColumn::SetListIndex(short nNewValue) 
/*
Routine Description:
	Replace current cell text as the specified item in choice list.

Parameters:
	nNewValue		The new value.

Return Value:
	None.
*/
{
	// Only be supported in grid control.
	if (!m_pGridCtrl)
		SetNotSupported();

	VERIFYCOLINDEX;

	// Only be significant in combo box or combo list style col.
	int nStyle = m_pGrid->GetColControl(nCol);
	if (nStyle != COLSTYLE_COMBOBOX && nStyle != COLSTYLE_COMBOBOX_DROPLIST)
		return;

	// Recalc the choice list.
	CalcChoiceList(m_pGrid->GetColChoiceList(nCol));

	if (nNewValue < 0 || nNewValue >= m_arChoiceList.GetSize())
		return;

	// Get current cell.
	int nRow, nColCurr;

	m_pGrid->GetCurrentCell(nRow, nColCurr);
	if (nRow > m_pGrid->GetRowCount() || nRow < 1)
		return;

	// Update the grid.
	m_pGrid->SetCellText(nRow, nCol, m_arChoiceList[nNewValue]);
}

void CColumn::AddItem(LPCTSTR Item, const VARIANT FAR& Index) 
/*
Routine Description:
	Add an item to choice list.

Parameters:
	Item		The new item string.

	Index		The position to insert.

Return Value:
	None.
*/
{
	// Only be supported in grid control.
	if (!m_pGridCtrl)
		SetNotSupported();

	VERIFYCOLINDEX;

	// Recalc choice list.
	CalcChoiceList(m_pGrid->GetColChoiceList(nCol));

	int nIndex;

	if (Index.vt == VT_ERROR)
	// Insert at the end default.
		nIndex = m_arChoiceList.GetSize();
	else
	{
		// Convert give index into integer.
		COleVariant vaIndex = Index;
		vaIndex.ChangeType(VT_I4);

		// Calc a valid index.
		if (vaIndex.lVal < 0)
			nIndex = 0;
		else if(vaIndex.lVal > m_arChoiceList.GetSize())
			nIndex = m_arChoiceList.GetSize();
		else
			nIndex = vaIndex.lVal;
	}

	// Insert new item into choice list.
	m_arChoiceList.InsertAt(nIndex, Item);

	// Update the grid.
	CalcChoiceList(nCol);
}

void CColumn::RemoveAll() 
/*
Routine Description:
	Remove all items in choice list.

Parameters:
	None.

Return Value:
	None.
*/
{
	// Only be supported in grid control.
	if (!m_pGridCtrl)
		SetNotSupported();

	VERIFYCOLINDEX;

	// Update the choice list.
	m_arChoiceList.RemoveAll();

	// Update the grid.
	CalcChoiceList(nCol);
}

void CColumn::CalcChoiceList(CString strList)
/*
Routine Description:
	Convert given string into string array.

Parameters:
	strList.

Return Value:
	None.
*/
{
	// Clear result array first.
	m_arChoiceList.RemoveAll();

	char ch;
	CString strChoice;

	for (int i = 0; i < strList.GetLength(); i ++)
	{
		ch = strList[i]; 
		if (ch == '\n')
		{
			// The end of current string.

			// Add it to result array.
			m_arChoiceList.Add(strChoice);

			// Empty current string.
			strChoice.Empty();
		}
		else
			strChoice += ch;
	}

	// Add the last string.
	if (ch != '\n')
		m_arChoiceList.Add(strChoice);
}

void CColumn::RemoveItem(short Index) 
/*
Routine Description:
	Remove one item from choice list.

Parameters:
	Index 		The item index.

Return Value:
	None.
*/
{
	// Only be supported in grid control.
	if (!m_pGridCtrl)
		SetNotSupported();

	VERIFYCOLINDEX;

	// Calc choice list.
	CalcChoiceList(m_pGrid->GetColChoiceList(nCol));

	if (Index < 0 || Index >= m_arChoiceList.GetSize())
		return;

	// Update the choice list.
	m_arChoiceList.RemoveAt(Index);

	// Update the grid.
	CalcChoiceList(nCol);
}

BSTR CColumn::GetList(short Index) 
/*
Routine Description:
	Get one item in choice list.

Parameters:
	Index		The item index.

Return Value:
	The item string.
*/
{
	// Only be supported in grid control.
	if (!m_pGridCtrl)
		GetNotSupported();

	VERIFYCOLINDEX;

	CString strResult;
	
	// Calc choice list array.
	CalcChoiceList(m_pGrid->GetColChoiceList(nCol));

	if (Index >= 0 && Index < m_arChoiceList.GetSize())
		strResult = m_arChoiceList[Index];

	return strResult.AllocSysString();
}

void CColumn::SetList(short Index, LPCTSTR lpszNewValue) 
/*
Routine Description:
	Set the text of one item in choice list.

Parameters:
	Index		The item index.

	lpszNewValue	The new value.

Return Value:
	None.
*/
{
	// Only be supported in grid control.
	if (!m_pGridCtrl)
		SetNotSupported();

	VERIFYCOLINDEX;

	// Calc choice list array.
	CalcChoiceList(m_pGrid->GetColChoiceList(nCol));

	if (Index < 0 || Index >= m_arChoiceList.GetSize())
		return;
	
	// Update the choice list array.
	m_arChoiceList[Index] = lpszNewValue;

	// Update the grid.
	CalcChoiceList(nCol);
}

void CColumn::Error(SCODE sc, UINT nID)
/*
Routine Description:
	Throw error using callback control.

Parameters:
	sc

	nID

Return Value:
	None.
*/
{
	if (m_pGridCtrl)
		m_pGridCtrl->ThrowError(sc, nID);
	else if (m_pDropDownCtrl)
		m_pDropDownCtrl->ThrowError(sc, nID);
	else
		m_pComboCtrl->ThrowError(sc, nID);
}

COLORREF CColumn::TranslateColor(OLE_COLOR clr)
/*
Routine Description:
	Translate color using callback control.

Parameters:
	clr		The color in OLE_COLOR format.

Return Value:
	The translated COLORREF value.
*/
{
	if (m_pGridCtrl)
		return m_pGridCtrl->TranslateColor(clr);
	else if (m_pDropDownCtrl)
		return m_pDropDownCtrl->TranslateColor(clr);
	else
		return m_pComboCtrl->TranslateColor(clr);
}

void CColumn::CalcChoiceList(int nCol)
/*
Routine Description:
	Set the choice list in grid.

Parameters:
	nCol		The col index.

Return Value:
	None.
*/
{
	CString strList;

	// Make up the choice list string.
	for (int i = 0; i < m_arChoiceList.GetSize(); i ++)
	{
		strList += _T('\n');
		strList += m_arChoiceList[i];
	}

	// Update the grid.
	m_pGrid->SetColChoiceList(nCol, strList);
}

OLE_HANDLE CColumn::GetDropDownHWnd() 
/*
Routine Description:
	Get the dropdown control handle.

Parameters:
	None.

Return Value:
	The window handle.
*/
{
	// Only be supported in grid control.
	if (!m_pGridCtrl)
		GetNotSupported();

	VERIFYCOLINDEX;

	CStringArray & arUA = m_pGrid->GetColUserAttribRef(nCol);
	if (arUA.GetSize() < 3)
		return NULL;
	else
		return (OLE_HANDLE)atoi(arUA[2]);
}

void CColumn::SetDropDownHWnd(OLE_HANDLE nNewValue) 
/*
Routine Description:
	Set the dropdown control handle.

Parameters:
	None.

Return Value:
	None.
*/
{
	if (!m_pGridCtrl)
		SetNotSupported();

	VERIFYCOLINDEX;

	m_pGrid->GetColUserAttribRef(nCol).SetSize(3);
	m_pGrid->GetColUserAttribRef(nCol).ElementAt(2).Format("%d", nNewValue);

	// If external dropdown is removed, mark it.
	if (!nNewValue)
		m_pGridCtrl->m_bDroppedDown = FALSE;
}

BOOL CColumn::IsColForceLock(CString strField)
/*
Routine Description:
	Decides if one field can only be readonly.

Parameters:
	strField		The field name.

Return Value:
	If can only be readonly, return TRUE; Otherwise, return FALSE.
*/
{
	int nField;

	if (m_pGridCtrl)
	{
		nField = m_pGridCtrl->GetFieldFromStr(strField);

		if (nField == -1)
			return FALSE;
		else
			return m_pGridCtrl->m_pColumnInfo[nField].dwFlags & DBCOLUMNFLAGS_WRITE;
	}
	else if (m_pDropDownCtrl)
	{
		nField = m_pDropDownCtrl->GetFieldFromStr(strField);

		if (nField == -1)
			return FALSE;
		else
			return m_pDropDownCtrl->m_pColumnInfo[nField].dwFlags & DBCOLUMNFLAGS_WRITE;
	}
	else
	{
		nField = m_pComboCtrl->GetFieldFromStr(strField);

		if (nField == -1)
			return FALSE;
		else
			return m_pComboCtrl->m_pColumnInfo[nField].dwFlags & DBCOLUMNFLAGS_WRITE;
	}
}

BOOL CColumn::IsColForceNotNullable(CString strField)
/*
Routine Description:
	Decides if one field can not use null value.

Parameters:
	strField		The field name.

Return Value:
	If can not use null value, return TRUE; Otherwise, return FALSE.
*/
{
	int nField;

	if (m_pGridCtrl)
	{
		nField = m_pGridCtrl->GetFieldFromStr(strField);

		if (nField == -1)
			return FALSE;
		else
			return m_pGridCtrl->m_pColumnInfo[nField].dwFlags & DBCOLUMNFLAGS_ISNULLABLE;
	}
	else if (m_pDropDownCtrl)
	{
		nField = m_pDropDownCtrl->GetFieldFromStr(strField);

		if (nField == -1)
			return FALSE;
		else
			return m_pDropDownCtrl->m_pColumnInfo[nField].dwFlags & DBCOLUMNFLAGS_ISNULLABLE;
	}
	else
	{
		nField = m_pComboCtrl->GetFieldFromStr(strField);

		if (nField == -1)
			return FALSE;
		else
			return m_pComboCtrl->m_pColumnInfo[nField].dwFlags & DBCOLUMNFLAGS_ISNULLABLE;
	}
}

int CColumn::GetDataMode()
/*
Routine Description:
	Get the data mode of callback control.

Parameters:
	None.

Return Value:
	The data mode.
*/
{
	if (m_pGridCtrl)
		return m_pGridCtrl->m_nDataMode;
	else if (m_pDropDownCtrl)
		return m_pDropDownCtrl->m_nDataMode;
	else
		return m_pComboCtrl->m_nDataMode;
}
