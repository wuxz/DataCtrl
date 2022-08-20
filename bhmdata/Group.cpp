// Group.cpp : implementation file
//

#include "stdafx.h"
#include "bhmdata.h"
#include "Group.h"
#include "BHMDBDropDownCtl.h"
#include "BHMDBComboCtl.h"
#include "BHMDBGridCtl.h"
#include "Columns.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

// Verify group index, throw error if is invalid.
#define VERIFYGROUPINDEX \
	int nGroup = m_pGrid->GetGroupFromIndex(m_nGroupIndex); \
	if (nGroup == INVALID) \
		Error(CTL_E_INVALIDPROPERTYARRAYINDEX, AFX_IDP_E_INVALIDPROPERTYARRAYINDEX); \

/////////////////////////////////////////////////////////////////////////////
// CGroup

IMPLEMENT_DYNCREATE(CGroup, CCmdTarget)

IMPLEMENT_OLETYPELIB(CGroup, _tlid, _wVerMajor, _wVerMinor)

CGroup::CGroup()
{
	EnableAutomation();
	EnableTypeLib();

	m_pDropDownCtrl = NULL;
	m_pComboCtrl = NULL;
	m_pGridCtrl = NULL;
	m_pGrid = NULL;
}

// Set callback control pointers.
void CGroup::SetCtrlPtr(CBHMDBDropDownCtrl * pCtrl)
{
	m_pDropDownCtrl = pCtrl;
}

void CGroup::SetCtrlPtr(CBHMDBComboCtrl * pCtrl)
{
	m_pComboCtrl = pCtrl;
}

void CGroup::SetCtrlPtr(CBHMDBGridCtrl * pCtrl)
{
	m_pGridCtrl = pCtrl;
}

CGroup::~CGroup()
{
}


void CGroup::OnFinalRelease()
{
	// When the last reference for an automation object is released
	// OnFinalRelease is called.  The base class will automatically
	// deletes the object.  Add additional cleanup required for your
	// object before calling the base class.

	CCmdTarget::OnFinalRelease();
}


BEGIN_MESSAGE_MAP(CGroup, CCmdTarget)
	//{{AFX_MSG_MAP(CGroup)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(CGroup, CCmdTarget)
	//{{AFX_DISPATCH_MAP(CGroup)
	DISP_PROPERTY_EX(CGroup, "BackColor", GetBackColor, SetBackColor, VT_COLOR)
	DISP_PROPERTY_EX(CGroup, "ForeColor", GetForeColor, SetForeColor, VT_COLOR)
	DISP_PROPERTY_EX(CGroup, "Name", GetName, SetName, VT_BSTR)
	DISP_PROPERTY_EX(CGroup, "Caption", GetCaption, SetCaption, VT_BSTR)
	DISP_PROPERTY_EX(CGroup, "Width", GetWidth, SetWidth, VT_I4)
	DISP_PROPERTY_EX(CGroup, "Visible", GetVisible, SetVisible, VT_BOOL)
	DISP_PROPERTY_EX(CGroup, "AllowSizing", GetAllowSizing, SetAllowSizing, VT_BOOL)
	DISP_PROPERTY_EX(CGroup, "Columns", GetColumns, SetNotSupported, VT_DISPATCH)
	//}}AFX_DISPATCH_MAP
END_DISPATCH_MAP()

// Note: we add support for IID_IGroup to support typesafe binding
//  from VBA.  This IID must match the GUID that is attached to the 
//  dispinterface in the .ODL file.

// {0E9D7C60-7046-11D3-A7FE-0080C8763FA4}
static const IID IID_IGroup =
{ 0xe9d7c60, 0x7046, 0x11d3, { 0xa7, 0xfe, 0x0, 0x80, 0xc8, 0x76, 0x3f, 0xa4 } };

BEGIN_INTERFACE_MAP(CGroup, CCmdTarget)
	INTERFACE_PART(CGroup, IID_IGroup, Dispatch)
END_INTERFACE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CGroup message handlers

BOOL CGroup::GetDispatchIID(IID *pIID)
{
	*pIID = IID_IGroup;

	return TRUE;
}

OLE_COLOR CGroup::GetBackColor() 
/*
Routine Description:
	Get the back color.

Parameters:
	None.

Return Value:
	The color.
*/
{
	VERIFYGROUPINDEX;

	return m_pGrid->GetGroupBackColor(nGroup);
}

void CGroup::SetBackColor(OLE_COLOR nNewValue) 
/*
Routine Description:
	Set the back color.

Parameters:
	nNewValue		The new value.

Return Value:
	None.
*/
{
	VERIFYGROUPINDEX;

	// Update the grid.
	m_pGrid->SetGroupBackColor(nGroup, TranslateColor(nNewValue));
}

OLE_COLOR CGroup::GetForeColor() 
/*
Routine Description:
	Get the fore color.

Parameters:
	None.

Return Value:
	The color.
*/
{
	VERIFYGROUPINDEX;

	return m_pGrid->GetGroupForeColor(nGroup);
}

void CGroup::SetForeColor(OLE_COLOR nNewValue) 
/*
Routine Description:
	Set the fore color.

Parameters:
	nNewValue		The new value.

Return Value:
	None.
*/
{
	VERIFYGROUPINDEX;

	// Update the grid.
	m_pGrid->SetGroupForeColor(nGroup, TranslateColor(nNewValue));
}

BSTR CGroup::GetName() 
/*
Routine Description:
	Get the name.

Parameters:
	None.

Return Value:
	The name string.
*/
{
	VERIFYGROUPINDEX;

	return m_pGrid->GetGroupName(nGroup).AllocSysString();
}

void CGroup::SetName(LPCTSTR lpszNewValue) 
/*
Routine Description:
	Set the name.

Parameters:
	lpszNewValue		The new value.

Return Value:
	None.
*/
{
	// Can not be empty.
	if (lstrlen(lpszNewValue) == 0)
		Error(CTL_E_INVALIDPROPERTYVALUE, AFX_IDP_E_INVALIDPROPERTYVALUE);

	VERIFYGROUPINDEX;

	// Update the grid.
	if (!m_pGrid->SetGroupName(nGroup, lpszNewValue))
		Error(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_NAMENOTUNIQUE);
}

BSTR CGroup::GetCaption() 
/*
Routine Description:
	Get the caption.

Parameters:
	None.

Return Value:
	The caption string.
*/
{
	VERIFYGROUPINDEX;

	return m_pGrid->GetGroupTitle(nGroup).AllocSysString();
}

void CGroup::SetCaption(LPCTSTR lpszNewValue) 
/*
Routine Description:
	Set the caption.

Parameters:
	lpszNewValue		The new value.

Return Value:
	None.
*/
{
	VERIFYGROUPINDEX;

	// Update the grid.
	m_pGrid->SetGroupTitle(nGroup, lpszNewValue);
}

long CGroup::GetWidth() 
/*
Routine Description:
	Get the width.

Parameters:
	None.

Return Value:
	The width.
*/
{
	VERIFYGROUPINDEX;

	return m_pGrid->GetGroupWidth(nGroup);
}

void CGroup::SetWidth(long nNewValue) 
/*
Routine Description:
	Set the width.

Parameters:
	nNewValue		The new value.

Return Value:
	None.
*/
{
	// Width can not less than 0.
	if (nNewValue < 0)
		Error(CTL_E_INVALIDPROPERTYVALUE, AFX_IDP_E_INVALIDPROPERTYVALUE);

	VERIFYGROUPINDEX;

	// Update the grid.
	m_pGrid->SetGroupWidth(nGroup, nNewValue);
}

BOOL CGroup::GetVisible() 
/*
Routine Description:
	Get the visibility.

Parameters:
	None.

Return Value:
	If is visible, return TRUE; Otherwise, return FALSE.
*/
{
	VERIFYGROUPINDEX;


	return m_pGrid->GetGroupVisible(nGroup);
}

void CGroup::SetVisible(BOOL bNewValue) 
/*
Routine Description:
	Set the visibility.

Parameters:
	bNewValue		The new value.

Return Value:
	None.
*/
{
	VERIFYGROUPINDEX;

	// Update the grid.
	m_pGrid->SetGroupVisible(nGroup, bNewValue);
}

BOOL CGroup::GetAllowSizing() 
/*
Routine Description:
	Get the permission to resize group width manually.

Parameters:
	None.

Return Value:
	The permission.
*/
{
	VERIFYGROUPINDEX;

	return m_pGrid->GetGroupAllowSizing(nGroup);
}

void CGroup::SetAllowSizing(BOOL bNewValue) 
/*
Routine Description:
	Get the permission to resize group width manually.

Parameters:
	The permission.

Return Value:
	None.
*/
{
	VERIFYGROUPINDEX;

	// Update the grid.
	m_pGrid->SetGroupAllowSizing(nGroup, bNewValue);
}

LPDISPATCH CGroup::GetColumns() 
/*
Routine Description:
	Get the columns object.

Parameters:
	None.

Return Value:
	The columns object.
*/
{
	// Create a new columns object.
	CColumns * pColumns = new CColumns;
	
	// Set callback pointers.
	if (m_pGridCtrl)
		pColumns->SetCtrlPtr(m_pGridCtrl);
	else if (m_pDropDownCtrl)
		pColumns->SetCtrlPtr(m_pDropDownCtrl);
	else
		pColumns->SetCtrlPtr(m_pComboCtrl);

	// Set the grid window.
	pColumns->m_pGrid = m_pGrid;

	// Set the group index.
	pColumns->m_nGroupIndex = m_nGroupIndex;

	return pColumns->GetIDispatch(FALSE);
}

void CGroup::Error(SCODE sc, UINT nID)
/*
Routine Description:
	Throw error.

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

COLORREF CGroup::TranslateColor(OLE_COLOR clr)
/*
Routine Description:
	Translate color.

Parameters:
	clr		The color of OLE_COLOR type.

Return Value:
	The COLORREF type color.
*/
{
	if (m_pGridCtrl)
		return m_pGridCtrl->TranslateColor(clr);
	else if (m_pDropDownCtrl)
		return m_pDropDownCtrl->TranslateColor(clr);
	else
		return m_pComboCtrl->TranslateColor(clr);
}