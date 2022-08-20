// SelBookmarks.cpp : implementation file
//

#include "stdafx.h"
#include "bhmdata.h"
#include "SelBookmarks.h"
#include "BHMDBGridCtl.h"
#include "GridInner.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CSelBookmarks

IMPLEMENT_DYNCREATE(CSelBookmarks, CCmdTarget)

IMPLEMENT_OLETYPELIB(CSelBookmarks, _tlid, _wVerMajor, _wVerMinor)

CSelBookmarks::CSelBookmarks()
{
	EnableAutomation();
	EnableTypeLib();
	
	m_pCtrl = NULL;
}

// Set callback pointer.
CSelBookmarks::SetCtrlPtr(CBHMDBGridCtrl * pCtrl)
{
	m_pCtrl = pCtrl;
}

CSelBookmarks::~CSelBookmarks()
{
}


void CSelBookmarks::OnFinalRelease()
{
	// When the last reference for an automation object is released
	// OnFinalRelease is called.  The base class will automatically
	// deletes the object.  Add additional cleanup required for your
	// object before calling the base class.

	CCmdTarget::OnFinalRelease();
}


BEGIN_MESSAGE_MAP(CSelBookmarks, CCmdTarget)
	//{{AFX_MSG_MAP(CSelBookmarks)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(CSelBookmarks, CCmdTarget)
	//{{AFX_DISPATCH_MAP(CSelBookmarks)
	DISP_PROPERTY_EX(CSelBookmarks, "Count", GetCount, SetNotSupported, VT_I2)
	DISP_FUNCTION(CSelBookmarks, "Remove", Remove, VT_EMPTY, VTS_I4)
	DISP_FUNCTION(CSelBookmarks, "Add", Add, VT_EMPTY, VTS_VARIANT)
	DISP_FUNCTION(CSelBookmarks, "RemoveAll", RemoveAll, VT_EMPTY, VTS_NONE)
	DISP_PROPERTY_PARAM(CSelBookmarks, "Item", Item, SetNotSupported, VT_DISPATCH, VTS_I4)
	DISP_DEFVALUE(CSelBookmarks, "Item")
	//}}AFX_DISPATCH_MAP
END_DISPATCH_MAP()

// Note: we add support for IID_ISelBookmarks to support typesafe binding
//  from VBA.  This IID must match the GUID that is attached to the 
//  dispinterface in the .ODL file.

// {BF91B15E-6A84-11D3-A7FE-0080C8763FA4}
static const IID IID_ISelBookmarks =
{ 0xbf91b15e, 0x6a84, 0x11d3, { 0xa7, 0xfe, 0x0, 0x80, 0xc8, 0x76, 0x3f, 0xa4 } };

BEGIN_INTERFACE_MAP(CSelBookmarks, CCmdTarget)
	INTERFACE_PART(CSelBookmarks, IID_ISelBookmarks, Dispatch)
	INTERFACE_PART(CSelBookmarks, IID_IEnumVARIANT, EnumVARIANT)
END_INTERFACE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CSelBookmarks message handlers
CSelBookmarks::XEnumVARIANT::XEnumVARIANT()
{
	m_nCurrent = 0;
}

STDMETHODIMP_(ULONG) CSelBookmarks::XEnumVARIANT::AddRef()
{   
	METHOD_PROLOGUE(CSelBookmarks, EnumVARIANT)
	return pThis->ExternalAddRef();
}   

STDMETHODIMP_(ULONG) CSelBookmarks::XEnumVARIANT::Release()
{   
	METHOD_PROLOGUE(CSelBookmarks, EnumVARIANT)
	return pThis->ExternalRelease();
}   

STDMETHODIMP CSelBookmarks::XEnumVARIANT::QueryInterface(REFIID iid, void FAR* FAR* ppvObj)
{   
	METHOD_PROLOGUE(CSelBookmarks, EnumVARIANT)
	return (HRESULT)pThis->ExternalQueryInterface((void FAR*)&iid, ppvObj);
}   

// IEnumVARIANT::Next
// 
STDMETHODIMP CSelBookmarks::XEnumVARIANT::Next(ULONG celt, VARIANT FAR* rgvar, ULONG FAR* pceltFetched)
{
	METHOD_PROLOGUE(CSelBookmarks, EnumVARIANT)

	// pceltFetched can legally == 0
	//                                           
	if (pceltFetched != NULL)
		*pceltFetched = 0;
	else if (celt > 1)
	{   
		return ResultFromScode(E_INVALIDARG); 
	}

    CArray<int, int> arRows;
	int nRows = pThis->m_pCtrl->m_pGridInner->GetSelectedRows(arRows);

	// Retrieve the next celt elements.
	if (m_nCurrent < nRows)
	{
		VariantInit(&rgvar[0]);

		// Copy bookmark from grid.
		pThis->m_pCtrl->GetBmkOfRow[arRows[m_nCurrent ++], rgvar];
		if (pceltFetched != NULL)
			*pceltFetched = 1;

		return NOERROR;
	}

	return S_FALSE;
}

// IEnumVARIANT::Skip
//
STDMETHODIMP CSelBookmarks::XEnumVARIANT::Skip(unsigned long celt)
{
    METHOD_PROLOGUE(CSelBookmarks, EnumVARIANT)

    CArray<int, int> arRows;
	int nRows = pThis->m_pCtrl->m_pGridInner->GetSelectedRows(arRows);

	while (m_nCurrent < nRows && celt --)
        m_nCurrent ++;
    
    return celt == 0 ? NOERROR : S_FALSE;
}

STDMETHODIMP CSelBookmarks::XEnumVARIANT::Reset()
{
    METHOD_PROLOGUE(CSelBookmarks, EnumVARIANT)
 
	m_nCurrent = 0;
	
	return NOERROR;
}

STDMETHODIMP CSelBookmarks::XEnumVARIANT::Clone(IEnumVARIANT FAR* FAR* ppenum)
{
    METHOD_PROLOGUE(CSelBookmarks, EnumVARIANT)   

	// Create a new object.
    CSelBookmarks* p = new CSelBookmarks;

	// Set callback pointer.
	p->SetCtrlPtr(pThis->m_pCtrl);
    if (p)
    {
	// Copy current position.
        p->m_xEnumVARIANT.m_nCurrent = m_nCurrent ;
        return NOERROR ;    
    }
    else
        return ResultFromScode(E_OUTOFMEMORY);
}

/////////////////////////////////////////////////////////////////////////////
// CSelBookmarks message handlers

long CSelBookmarks::GetCount() 
{
	// Get the selection count.
    CArray<int, int> arRows;
	int nRows = m_pCtrl->m_pGridInner->GetSelectedRows(arRows);

	return nRows;
}

void CSelBookmarks::Remove(long Index)
{
	// Get the selection array.
    CArray<int, int> arRows;
	int nRows = m_pCtrl->m_pGridInner->GetSelectedRows(arRows);
	
	if (Index < 1 || Index > nRows)
		m_pCtrl->ThrowError(CTL_E_INVALIDPROPERTYARRAYINDEX, AFX_IDP_E_INVALIDPROPERTYARRAYINDEX);

	// Update the grid.
	m_pCtrl->m_pGridInner->SelectRow(arRows[Index - 1], FALSE);
}

void CSelBookmarks::Add(const VARIANT FAR& Bookmark) 
{
	// Get the desired row ordinal.
	int nRow = m_pCtrl->GetRowFromBmk(&Bookmark);
	if (nRow == INVALID)
		m_pCtrl->ThrowError(CTL_E_INVALIDPROPERTYVALUE, IDS_ERROR_INVALIDBOOKMARK);

	// Update the grid.
	m_pCtrl->m_pGridInner->SelectRow(nRow, TRUE);
}

void CSelBookmarks::RemoveAll() 
{
	// Update the grid.
    CArray<int, int> arRows;
	int nRows = m_pCtrl->m_pGridInner->GetSelectedRows(arRows);
	if (nRows > 0)
		m_pCtrl->m_pGridInner->ResetSelection();
}

LPUNKNOWN CSelBookmarks::_NewEnum()
{
	return GetIDispatch(TRUE);
}

VARIANT CSelBookmarks::Item(long Index)
{
	// Get the selection array.
    CArray<int, int> arRows;
	int nRows = m_pCtrl->m_pGridInner->GetSelectedRows(arRows);
	
	if (Index < 1 || Index > nRows)
		m_pCtrl->ThrowError(CTL_E_INVALIDPROPERTYARRAYINDEX, AFX_IDP_E_INVALIDPROPERTYARRAYINDEX);

	// Get the bookmark from desired row.
	VARIANT va;
	VariantInit(&va);
	m_pCtrl->GetBmkOfRow(arRows[Index - 1], &va);

	return va;
}

BOOL CSelBookmarks::GetDispatchIID(IID *pIID)
{
	*pIID = IID_ISelBookmarks;

	return TRUE;
}
