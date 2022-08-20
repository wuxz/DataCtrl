// Groups.cpp : implementation file
//

#include "stdafx.h"
#include "bhmdata.h"
#include "Groups.h"
#include "BHMDBDropDownCtl.h"
#include "BHMDBComboCtl.h"
#include "BHMDBGridCtl.h"
#include "Group.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CGroups

IMPLEMENT_DYNCREATE(CGroups, CCmdTarget)

IMPLEMENT_OLETYPELIB(CGroups, _tlid, _wVerMajor, _wVerMinor)

CGroups::CGroups()
{
	EnableAutomation();
	EnableTypeLib();
	
	m_pDropDownCtrl = NULL;
	m_pComboCtrl = NULL;
	m_pGridCtrl = NULL;
	m_pGrid = NULL;
}

// Set callback pointers.
void CGroups::SetCtrlPtr(CBHMDBDropDownCtrl * pCtrl)
{
	m_pDropDownCtrl = pCtrl;
}

void CGroups::SetCtrlPtr(CBHMDBComboCtrl * pCtrl)
{
	m_pComboCtrl = pCtrl;
}

void CGroups::SetCtrlPtr(CBHMDBGridCtrl * pCtrl)
{
	m_pGridCtrl = pCtrl;
}

CGroups::~CGroups()
{
}


void CGroups::OnFinalRelease()
{
	// When the last reference for an automation object is released
	// OnFinalRelease is called.  The base class will automatically
	// deletes the object.  Add additional cleanup required for your
	// object before calling the base class.

	CCmdTarget::OnFinalRelease();
}


BEGIN_MESSAGE_MAP(CGroups, CCmdTarget)
	//{{AFX_MSG_MAP(CGroups)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(CGroups, CCmdTarget)
	//{{AFX_DISPATCH_MAP(CGroups)
	DISP_PROPERTY_EX(CGroups, "Count", GetCount, SetNotSupported, VT_I2)
	DISP_FUNCTION(CGroups, "Remove", Remove, VT_EMPTY, VTS_VARIANT)
	DISP_FUNCTION(CGroups, "Add", Add, VT_EMPTY, VTS_VARIANT)
	DISP_FUNCTION(CGroups, "RemoveAll", RemoveAll, VT_EMPTY, VTS_NONE)
	DISP_PROPERTY_PARAM(CGroups, "Item", GetItem, SetNotSupported, VT_DISPATCH, VTS_VARIANT)
	DISP_DEFVALUE(CGroups, "Item")
	//}}AFX_DISPATCH_MAP
	DISP_PROPERTY_EX_ID(CGroups, "_NewEnum", DISPID_NEWENUM, _NewEnum, SetNotSupported, VT_UNKNOWN)
END_DISPATCH_MAP()

// Note: we add support for IID_IGroups to support typesafe binding
//  from VBA.  This IID must match the GUID that is attached to the 
//  dispinterface in the .ODL file.

// {AD466A40-703F-11D3-A7FE-0080C8763FA4}
static const IID IID_IGroups =
{ 0xad466a40, 0x703f, 0x11d3, { 0xa7, 0xfe, 0x0, 0x80, 0xc8, 0x76, 0x3f, 0xa4 } };

BEGIN_INTERFACE_MAP(CGroups, CCmdTarget)
	INTERFACE_PART(CGroups, IID_IGroups, Dispatch)
END_INTERFACE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CGroups message handlers
CGroups::XEnumVARIANT::XEnumVARIANT()
{
	m_nCurrent = 0;
}

STDMETHODIMP_(ULONG) CGroups::XEnumVARIANT::AddRef()
{   
	METHOD_PROLOGUE(CGroups, EnumVARIANT)
	return pThis->ExternalAddRef();
}   

STDMETHODIMP_(ULONG) CGroups::XEnumVARIANT::Release()
{   
	METHOD_PROLOGUE(CGroups, EnumVARIANT)
	return pThis->ExternalRelease();
}   

STDMETHODIMP CGroups::XEnumVARIANT::QueryInterface(REFIID iid, void FAR* FAR* ppvObj)
{   
	METHOD_PROLOGUE(CGroups, EnumVARIANT)
	return (HRESULT)pThis->ExternalQueryInterface((void FAR*)&iid, ppvObj);
}   

// IEnumVARIANT::Next
// 
STDMETHODIMP CGroups::XEnumVARIANT::Next(ULONG celt, VARIANT FAR* rgvar, ULONG FAR* pceltFetched)
{
	METHOD_PROLOGUE(CGroups, EnumVARIANT)

	CGroup * pItem = NULL;

	// pceltFetched can legally == 0
	//                                           
	if (pceltFetched != NULL)
		*pceltFetched = 0;
	else if (celt > 1)
	{   
		return ResultFromScode(E_INVALIDARG); 
	}

    // Retrieve the next celt elements.
	if (m_nCurrent < pThis->m_pGrid->GetGroupCount())
	{
		VariantInit(&rgvar[0]);

		// Create a new group object.			
		pItem = new CGroup;

		// Set the callback pointers.
		if (pThis->m_pGridCtrl)
			pItem->SetCtrlPtr(pThis->m_pGridCtrl);
		else if (pThis->m_pDropDownCtrl)
			pItem->SetCtrlPtr(pThis->m_pDropDownCtrl);
		else
			pItem->SetCtrlPtr(pThis->m_pComboCtrl);

		pItem->m_pGrid = pThis->m_pGrid;

		// Set the group index.
		pItem->m_nGroupIndex = pThis->m_pGrid->GetGroupIndex(++ m_nCurrent);

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
STDMETHODIMP CGroups::XEnumVARIANT::Skip(unsigned long celt)
{
    METHOD_PROLOGUE(CGroups, EnumVARIANT)

	while (m_nCurrent < pThis->m_pGrid->GetGroupCount() && celt --)
       m_nCurrent ++;

    return celt == 0 ? NOERROR : S_FALSE;
}

STDMETHODIMP CGroups::XEnumVARIANT::Reset()
{
    METHOD_PROLOGUE(CGroups, EnumVARIANT)
 
	m_nCurrent = 0;
	
	return NOERROR;
}

STDMETHODIMP CGroups::XEnumVARIANT::Clone(IEnumVARIANT FAR* FAR* ppenum)
{
    METHOD_PROLOGUE(CGroups, EnumVARIANT)   

	// Create new groups object.
    CGroups* p = new CGroups;
	
	// Set callback pointers.
	if (pThis->m_pGridCtrl)
		p->SetCtrlPtr(pThis->m_pGridCtrl);
	else if (pThis->m_pDropDownCtrl)
		p->SetCtrlPtr(pThis->m_pDropDownCtrl);
	else
		p->SetCtrlPtr(pThis->m_pComboCtrl);

	p->m_pGrid = pThis->m_pGrid;

    if (p)
    {
	// Copy current position.
        p->m_xEnumVARIANT.m_nCurrent = m_nCurrent ;
        return NOERROR ;    
    }
    else
        return ResultFromScode(E_OUTOFMEMORY);
}

LPUNKNOWN CGroups::_NewEnum()
{
	return GetIDispatch(TRUE);
}

short CGroups::GetCount() 
{
	return m_pGrid->GetGroupCount();
}

void CGroups::Remove(const VARIANT FAR& GroupIndex) 
{
	// Can not remove group in bound mode.
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

	// Calc the group ordinal.
	int nGroupIndex = GetRowColFromVariant(GroupIndex);
	if (nGroupIndex < 1 || nGroupIndex > m_pGrid->GetGroupCount())
		Error(CTL_E_INVALIDPROPERTYARRAYINDEX, AFX_IDP_E_INVALIDPROPERTYARRAYINDEX);

	// Update the grid.
	m_pGrid->RemoveGroup(nGroupIndex);
}

void CGroups::Add(const VARIANT FAR& GroupIndex) 
{
	// Can not add group in bound mode.
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

	// Calc group ordinal.
	int nGroupIndex = GetRowColFromVariant(GroupIndex);
	if (nGroupIndex < 1 || nGroupIndex > m_pGrid->GetGroupCount() + 1)
		Error(CTL_E_INVALIDPROPERTYARRAYINDEX, AFX_IDP_E_INVALIDPROPERTYARRAYINDEX);

	// Update the grid.
	m_pGrid->InsertGroup(nGroupIndex);
}

void CGroups::RemoveAll() 
{
	// Can not remove group in bound mode.
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

	// Update the grid.
	m_pGrid->SetGroupCount(0);
}

LPDISPATCH CGroups::GetItem(const VARIANT FAR& GroupIndex) 
{
/*	if (m_pGridCtrl)
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
	}*/

	// Calc the group index.
	int nGroupIndex = GetRowColFromVariant(GroupIndex);
	if (nGroupIndex < 1 || nGroupIndex > m_pGrid->GetGroupCount())
		Error(CTL_E_INVALIDPROPERTYARRAYINDEX, AFX_IDP_E_INVALIDPROPERTYARRAYINDEX);

	// Create a new group object.
	CGroup * pItem = new CGroup;

	// Set callback pointers.
	if (m_pGridCtrl)
		pItem->SetCtrlPtr(m_pGridCtrl);
	else if (m_pDropDownCtrl)
		pItem->SetCtrlPtr(m_pDropDownCtrl);
	else
		pItem->SetCtrlPtr(m_pComboCtrl);

	pItem->m_pGrid = m_pGrid;

	// Copy group index.
	pItem->m_nGroupIndex = m_pGrid->GetGroupIndex(nGroupIndex);

	return pItem->GetIDispatch(FALSE);
}

void CGroups::Error(SCODE sc, UINT nID)
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

int CGroups::GetRowColFromVariant(VARIANT va)
/*
Routine Description:
	Calc ordinal from variant value.

Parameters:
	va		THe variant value.

Return Value:
	The result ordinal.
*/
{
	USES_CONVERSION;

	COleVariant vaNew = va;
	CString str;

	vaNew.ChangeType(VT_BSTR);
	str = OLE2T(vaNew.bstrVal);

	// Can not be empty.
	if (str.GetLength() == 0)
		return INVALID;

	if (str[0] >= (TCHAR)'0' && str[0] <= (TCHAR)'9')
	{
		// The variant value contains group ordinal.
		vaNew.ChangeType(VT_I4);
		return vaNew.lVal;
	}

	for (int i = 0; i < m_pGrid->GetColCount(); i ++)
	{
		// The variant value contains group name.
		if (!m_pGrid->GetColName(i + 1).CompareNoCase(str))
			return i;
	}

	return INVALID;
}

BOOL CGroups::GetDispatchIID(IID *pIID)
{
	*pIID = IID_IGroups;

	return TRUE;
}
