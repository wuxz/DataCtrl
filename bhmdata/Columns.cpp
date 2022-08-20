// Columns.cpp : implementation file
//

#include "stdafx.h"
#include "BHMData.h"
#include "Columns.h"
#include "BHMDBGridCtl.h"
#include "Column.h"
#include "GridInner.h"
#include "BHMDBDropDownCtl.h"
#include "BHMDBComboCtl.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CColumns

IMPLEMENT_DYNCREATE(CColumns, CCmdTarget)

IMPLEMENT_OLETYPELIB(CColumns, _tlid, _wVerMajor, _wVerMinor)

CColumns::CColumns()
{
	EnableAutomation();
	EnableTypeLib();
	
	m_pGridCtrl = NULL;
	m_pDropDownCtrl = NULL;
	m_pComboCtrl = NULL;
	m_nGroupIndex = -1;
}

// Set the callback control pointers.

CColumns::SetCtrlPtr(CBHMDBGridCtrl * pCtrl)
{
	// The control is grid control.
	m_pGridCtrl = pCtrl;
}

CColumns::SetCtrlPtr(CBHMDBDropDownCtrl * pCtrl)
{
	// The control is dropdown control.
	m_pDropDownCtrl = pCtrl;
}

CColumns::SetCtrlPtr(CBHMDBComboCtrl * pCtrl)
{
	// The control is combo control.
	m_pComboCtrl = pCtrl;
}

CColumns::~CColumns()
{
}


void CColumns::OnFinalRelease()
{
	// When the last reference for an automation object is released
	// OnFinalRelease is called.  The base class will automatically
	// deletes the object.  Add additional cleanup required for your
	// object before calling the base class.

	CCmdTarget::OnFinalRelease();
}


BEGIN_MESSAGE_MAP(CColumns, CCmdTarget)
	//{{AFX_MSG_MAP(CColumns)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(CColumns, CCmdTarget)
	//{{AFX_DISPATCH_MAP(CColumns)
	DISP_PROPERTY_EX(CColumns, "Count", GetCount, SetNotSupported, VT_I2)
	DISP_FUNCTION(CColumns, "Remove", Remove, VT_EMPTY, VTS_VARIANT)
	DISP_FUNCTION(CColumns, "Add", Add, VT_EMPTY, VTS_VARIANT)
	DISP_FUNCTION(CColumns, "RemoveAll", RemoveAll, VT_EMPTY, VTS_NONE)
	DISP_PROPERTY_PARAM(CColumns, "Item", Item, SetNotSupported, VT_DISPATCH, VTS_VARIANT)
	DISP_DEFVALUE(CColumns, "Item")
	//}}AFX_DISPATCH_MAP
	DISP_PROPERTY_EX_ID(CColumns, "_NewEnum", DISPID_NEWENUM, _NewEnum, SetNotSupported, VT_UNKNOWN)
END_DISPATCH_MAP()

// Note: we add support for IID_IColumns to support typesafe binding
//  from VBA.  This IID must match the GUID that is attached to the 
//  dispinterface in the .ODL file.

// {BF91B157-6A84-11D3-A7FE-0080C8763FA4}
static const IID IID_IColumns =
{ 0xbf91b157, 0x6a84, 0x11d3, { 0xa7, 0xfe, 0x0, 0x80, 0xc8, 0x76, 0x3f, 0xa4 } };

BEGIN_INTERFACE_MAP(CColumns, CCmdTarget)
	INTERFACE_PART(CColumns, IID_IColumns, Dispatch)
	INTERFACE_PART(CColumns, IID_IEnumVARIANT, EnumVARIANT)
END_INTERFACE_MAP()

CColumns::XEnumVARIANT::XEnumVARIANT()
{
	m_nCurrent = 0;
}

STDMETHODIMP_(ULONG) CColumns::XEnumVARIANT::AddRef()
{   
	METHOD_PROLOGUE(CColumns, EnumVARIANT)
	return pThis->ExternalAddRef();
}   

STDMETHODIMP_(ULONG) CColumns::XEnumVARIANT::Release()
{   
	METHOD_PROLOGUE(CColumns, EnumVARIANT)
	return pThis->ExternalRelease();
}   

STDMETHODIMP CColumns::XEnumVARIANT::QueryInterface(REFIID iid, void FAR* FAR* ppvObj)
{   
	METHOD_PROLOGUE(CColumns, EnumVARIANT)
	return (HRESULT)pThis->ExternalQueryInterface((void FAR*)&iid, ppvObj);
}   

// IEnumVARIANT::Next
// 
STDMETHODIMP CColumns::XEnumVARIANT::Next(ULONG celt, VARIANT FAR* rgvar, ULONG FAR* pceltFetched)
{
	METHOD_PROLOGUE(CColumns, EnumVARIANT)

	CColumn * pItem = NULL;

	// pceltFetched can legally == 0
	//                                           
	if (pceltFetched != NULL)
		*pceltFetched = 0;
	else if (celt > 1)
	{   
		return ResultFromScode(E_INVALIDARG); 
	}

	// Retrieve the next celt elements.
	int nColCount;
	int nGroup = INVALID;

	if (pThis->m_nGroupIndex != -1)
	{
		// Navigate inside a group.
		nGroup = pThis->m_pGrid->GetGroupFromIndex(pThis->m_nGroupIndex);
		if (nGroup == INVALID)
			pThis->Error(CTL_E_INVALIDPROPERTYARRAYINDEX, AFX_IDP_E_INVALIDPROPERTYARRAYINDEX);

		nColCount = pThis->m_pGrid->GetGroupColCount(nGroup);
	}
	else
	// Navigate inside grid.
		nColCount = pThis->m_pGrid->GetColCount();

	if (m_nCurrent < nColCount)
	{
		VariantInit(&rgvar[0]);

		// Create new column object.
		pItem = new CColumn;
		
		// Set callback pointers.
		if (pThis->m_pGridCtrl)
			pItem->SetCtrlPtr(pThis->m_pGridCtrl);
		else if (pThis->m_pDropDownCtrl)
			pItem->SetCtrlPtr(pThis->m_pDropDownCtrl);
		else
			pItem->SetCtrlPtr(pThis->m_pComboCtrl);

		pItem->m_pGrid = pThis->m_pGrid;

		// Set the col index.
		if (nGroup == INVALID)
			pItem->prop.nColIndex = pThis->m_pGrid->GetColIndex(++ m_nCurrent);
		else
			pItem->prop.nColIndex = pThis->m_pGrid->GetColIndex(pThis->m_pGrid->GetColFromGroup(nGroup - 1, ++ m_nCurrent));

		rgvar[0].vt = VT_DISPATCH;
		rgvar[0].pdispVal = pItem->GetIDispatch(FALSE);
		if (pceltFetched != NULL)
			*pceltFetched = 1;

		return NOERROR;
	}

	return S_FALSE;
}

// IEnumVARIANT::Skip
//
STDMETHODIMP CColumns::XEnumVARIANT::Skip(unsigned long celt)
{
    METHOD_PROLOGUE(CColumns, EnumVARIANT)

	int nColCount;

	if (pThis->m_nGroupIndex != -1)
	{
		// Navigate inside a group.
		int nGroup = pThis->m_pGrid->GetGroupFromIndex(pThis->m_nGroupIndex);
		if (nGroup == INVALID)
			pThis->Error(CTL_E_INVALIDPROPERTYARRAYINDEX, AFX_IDP_E_INVALIDPROPERTYARRAYINDEX);

		nColCount = pThis->m_pGrid->GetGroupColCount(nGroup);
	}
	else
	// Navigate inside grid.
		nColCount = pThis->m_pGrid->GetColCount();

	// Skip desired steps.
	while (m_nCurrent < nColCount && celt --)
        m_nCurrent ++;
    
    return celt == 0 ? NOERROR : S_FALSE;
}

STDMETHODIMP CColumns::XEnumVARIANT::Reset()
{
    METHOD_PROLOGUE(CColumns, EnumVARIANT)
 
	m_nCurrent = 0;
	
	return NOERROR;
}

STDMETHODIMP CColumns::XEnumVARIANT::Clone(IEnumVARIANT FAR* FAR* ppenum)
{
    METHOD_PROLOGUE(CColumns, EnumVARIANT)   

	// Create a clone of myself.
	CColumns* p = new CColumns;
	
	// Set callback pointers.
	if (pThis->m_pGridCtrl)
		p->SetCtrlPtr(pThis->m_pGridCtrl);
	else if (pThis->m_pDropDownCtrl)
		p->SetCtrlPtr(pThis->m_pDropDownCtrl);
	else
		p->SetCtrlPtr(pThis->m_pComboCtrl);

	p->m_pGrid = pThis->m_pGrid;
	p->m_nGroupIndex = pThis->m_nGroupIndex;

	if (p)
	{
		// Copy the current position.
		p->m_xEnumVARIANT.m_nCurrent = m_nCurrent ;
		return NOERROR ;    
	}
	else
		return ResultFromScode(E_OUTOFMEMORY);
}

/////////////////////////////////////////////////////////////////////////////
// CColumns message handlers

short CColumns::GetCount() 
/*
Routine Description:
	Get the col count.

Parameters:
	None.

Return Value:
	The count.
*/
{
	int nColCount;
	int nGroup = INVALID;

	if (m_nGroupIndex != -1)
	{
		// Operate inside a group.
		nGroup = m_pGrid->GetGroupFromIndex(m_nGroupIndex);
		if (nGroup == INVALID)
			Error(CTL_E_INVALIDPROPERTYARRAYINDEX, AFX_IDP_E_INVALIDPROPERTYARRAYINDEX);

		nColCount = m_pGrid->GetGroupColCount(nGroup);
	}
	else
	// Operate inside a grid.
		nColCount = m_pGrid->GetColCount();

	return nColCount;
}

// Remember that the index begins from 1

void CColumns::Remove(const VARIANT FAR& ColIndex) 
/*
Routine Description:
	Remove one col from grid.

Parameters:
	ColIndex		The col index.

Return Value:
	None.
*/
{
	// Data mode can not be bound mode.
	if (m_pGridCtrl)
	{
		if(m_pGridCtrl->m_nDataMode == 0)
			m_pGridCtrl->ThrowError(CTL_E_PERMISSIONDENIED, IDS_ERROR_NOAPPLICABLEINBINDMODE);
	}
	else if (m_pDropDownCtrl)
	{
		if(m_pDropDownCtrl->m_nDataMode == 0)
			m_pDropDownCtrl->ThrowError(CTL_E_PERMISSIONDENIED, IDS_ERROR_NOAPPLICABLEINBINDMODE);
	}
	else
	{
		if(m_pComboCtrl->m_nDataMode == 0)
			m_pComboCtrl->ThrowError(CTL_E_PERMISSIONDENIED, IDS_ERROR_NOAPPLICABLEINBINDMODE);
	}

	int nColCount;
	int nGroup = INVALID;

	// Get the col count.

	if (m_nGroupIndex != -1)
	{
		// Operate inside a group.

		nGroup = m_pGrid->GetGroupFromIndex(m_nGroupIndex);
		if (nGroup == INVALID)
			Error(CTL_E_INVALIDPROPERTYARRAYINDEX, AFX_IDP_E_INVALIDPROPERTYARRAYINDEX);

		nColCount = m_pGrid->GetGroupColCount(nGroup);
	}
	else
	// Operate inside grid.
		nColCount = m_pGrid->GetColCount();

	// Get the col index.
	int nColIndex = GetRowColFromVariant(ColIndex);

	// Validate the col index.
	if (nColIndex < 1 || nColIndex > nColCount)
		Error(CTL_E_INVALIDPROPERTYARRAYINDEX, AFX_IDP_E_INVALIDPROPERTYARRAYINDEX);

	if (nGroup != INVALID)
		nColIndex = m_pGrid->GetColFromGroup(nGroup - 1, nColIndex);

	// Update the grid.
	m_pGrid->RemoveCol(nColIndex);
}

void CColumns::Add(const VARIANT FAR& ColIndex) 
/*
Routine Description:
	Add a column to grid.

Parameters:
	ColIndex		The position to insert at.

Return Value:
	None.
*/
{
	// Data mode can not be bound mode.

	if (m_pGridCtrl)
	{
		if(m_pGridCtrl->m_nDataMode == 0)
			m_pGridCtrl->ThrowError(CTL_E_PERMISSIONDENIED, IDS_ERROR_NOAPPLICABLEINBINDMODE);
	}
	else if (m_pDropDownCtrl)
	{
		if(m_pDropDownCtrl->m_nDataMode == 0)
			m_pDropDownCtrl->ThrowError(CTL_E_PERMISSIONDENIED, IDS_ERROR_NOAPPLICABLEINBINDMODE);
	}
	else
	{
		if(m_pComboCtrl->m_nDataMode == 0)
			m_pComboCtrl->ThrowError(CTL_E_PERMISSIONDENIED, IDS_ERROR_NOAPPLICABLEINBINDMODE);
	}

	int nColCount;
	int nGroup = INVALID;

	// Calc the col count.

	if (m_nGroupIndex != -1)
	{
		// Operate inside a group.
		nGroup = m_pGrid->GetGroupFromIndex(m_nGroupIndex);
		if (nGroup == INVALID)
			Error(CTL_E_INVALIDPROPERTYARRAYINDEX, AFX_IDP_E_INVALIDPROPERTYARRAYINDEX);

		nColCount = m_pGrid->GetGroupColCount(nGroup);
	}
	else
	// Operate inside grid.
		nColCount = m_pGrid->GetColCount();

	// Validate the index.
	if (nColCount >= 255)
		Error(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_MAX255COLS);

	// Calc the position.
	int nColIndex = GetRowColFromVariant(ColIndex);

	// Validate the position.
	if (nColIndex < 1 || nColIndex > nColCount + 1)
		Error(CTL_E_INVALIDPROPERTYARRAYINDEX, AFX_IDP_E_INVALIDPROPERTYARRAYINDEX);

	// Update the grid.
	if (nGroup == INVALID)
	// Insert at the end default.
		m_pGrid->InsertCol(nColIndex);
	else
		m_pGrid->InsertColInGroup(nGroup, nColIndex);
}

void CColumns::RemoveAll() 
/*
Routine Description:
	Remove all cols in grid.

Parameters:
	None.

Return Value:
	None.
*/
{
	// Can not in bound mode.

	if (m_pGridCtrl)
	{
		if(m_pGridCtrl->m_nDataMode == 0)
			m_pGridCtrl->ThrowError(CTL_E_PERMISSIONDENIED, IDS_ERROR_NOAPPLICABLEINBINDMODE);
	}
	else if (m_pDropDownCtrl)
	{
		if(m_pDropDownCtrl->m_nDataMode == 0)
			m_pDropDownCtrl->ThrowError(CTL_E_PERMISSIONDENIED, IDS_ERROR_NOAPPLICABLEINBINDMODE);
	}
	else
	{
		if(m_pComboCtrl->m_nDataMode == 0)
			m_pComboCtrl->ThrowError(CTL_E_PERMISSIONDENIED, IDS_ERROR_NOAPPLICABLEINBINDMODE);
	}

	// Calc the col count.

	int nColCount;
	int nGroup = INVALID;

	if (m_nGroupIndex != -1)
	{
		// Operate inside a group.
		nGroup = m_pGrid->GetGroupFromIndex(m_nGroupIndex);
		if (nGroup == INVALID)
			Error(CTL_E_INVALIDPROPERTYARRAYINDEX, AFX_IDP_E_INVALIDPROPERTYARRAYINDEX);

		nColCount = m_pGrid->GetGroupColCount(nGroup);
	}

	// Update the grid.

	if (nGroup == INVALID)
	// Operate inside grid.
		m_pGrid->SetColCount(0);
	else
	{
		// Operate inside a group.

		// Get the col ordinal inside this group.
		int nColIndex = m_pGrid->GetColFromGroup(nGroup - 1, 1);
		while (nColCount > 0)
		{
			m_pGrid->RemoveCol(nColIndex);
			nColCount --;
		}
	}
}

LPUNKNOWN CColumns::_NewEnum()
{
	return GetIDispatch(TRUE);
}

int CColumns::GetRowColFromVariant(VARIANT va)
/*
Routine Description:
	Convert variant value into row/col ordinal.

Parameters:
	va		Variant value to be converted.

Return Value:
	None.
*/
{
	USES_CONVERSION;

	COleVariant vaNew = va;
	CString str;

	// Change to string format.
	vaNew.ChangeType(VT_BSTR);
	str = OLE2T(vaNew.bstrVal);
	if (str.GetLength() == 0)
	// Empty value is invalid.
		return INVALID;

	if (str[0] >= (TCHAR)'0' && str[0] <= (TCHAR)'9')
	{
		// The given data contains ordinal value.
		vaNew.ChangeType(VT_I4);
		return vaNew.lVal;
	}

	for (int i = 0; i < m_pGrid->GetColCount(); i ++)
	{
		// The given data contains col name.

		// Search this name in grid.
		if (!m_pGrid->GetColName(i + 1).CompareNoCase(str))
			return i;
	}

	return INVALID;
}

LPDISPATCH CColumns::Item(const VARIANT FAR& ColIndex) 
/*
Routine Description:
	Get the one column object in grid.

Parameters:
	ColIndex		The col ordinal.

Return Value:
	The result column object.
*/
{
	int nColCount;
	int nGroup = INVALID;

	// Calc the col count.

	if (m_nGroupIndex != -1)
	{
		// Operate inside a group.
		nGroup = m_pGrid->GetGroupFromIndex(m_nGroupIndex);
		if (nGroup == INVALID)
			Error(CTL_E_INVALIDPROPERTYARRAYINDEX, AFX_IDP_E_INVALIDPROPERTYARRAYINDEX);

		nColCount = m_pGrid->GetGroupColCount(nGroup);
	}
	else
	// Operate inside grid.
		nColCount = m_pGrid->GetColCount();

	// Calc the col ordinal.
	int nColIndex = GetRowColFromVariant(ColIndex);

	// Validate the col ordinal.
	if (nColIndex < 1 || nColIndex > nColCount)
		Error(CTL_E_INVALIDPROPERTYARRAYINDEX, AFX_IDP_E_INVALIDPROPERTYARRAYINDEX);

	if (nGroup != INVALID)
	// Operate inside a group.
		nColIndex = m_pGrid->GetColFromGroup(nGroup - 1, nColIndex);

	// Create a new column object.
	CColumn * pItem = new CColumn;

	// Set the callback pointers.
	if (m_pGridCtrl)
		pItem->SetCtrlPtr(m_pGridCtrl);
	else if (m_pDropDownCtrl)
		pItem->SetCtrlPtr(m_pDropDownCtrl);
	else
		pItem->SetCtrlPtr(m_pComboCtrl);

	pItem->m_pGrid = m_pGrid;

	// Set the col index.
	pItem->prop.nColIndex = m_pGrid->GetColIndex(nColIndex);

	return pItem->GetIDispatch(FALSE);
}

BOOL CColumns::GetDispatchIID(IID *pIID)
{
	*pIID = IID_IColumns;

	return TRUE;
}

void CColumns::Error(SCODE sc, UINT nID)
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
