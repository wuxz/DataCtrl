#if !defined(AFX_BHMDATA_H__BF91B133_6A84_11D3_A7FE_0080C8763FA4__INCLUDED_)
#define AFX_BHMDATA_H__BF91B133_6A84_11D3_A7FE_0080C8763FA4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

// BHMData.h : main header file for BHMDATA.DLL

#if !defined( __AFXCTL_H__ )
	#error include 'afxctl.h' before including this file
#endif

#include "resource.h"       // main symbols

#include "..\testgrid\grid.h"

// The property of column object.
struct ColumnProp
{
	ColumnProp()
	{
		bForceLock = FALSE;
		bForceNoNullable = FALSE;
		nDataField = -1;
		nColIndex = -1;
	};

	// The ordinal of the field bounded.
	int nDataField;

	// whether force this column to be locked because it is readonly in datasource
	BOOL bForceLock;

	// whether force this column to be not nullable because it is needed by datasource
	BOOL bForceNoNullable;

	// The col index in grid.
	int nColIndex;
};

/////////////////////////////////////////////////////////////////////////////
// CBHMDataApp : See BHMData.cpp for implementation.

class CBHMDataApp : public COleControlModule
{
public:
	// the common registered message
	static int m_nBringMsg;

	BOOL InitInstance();
	int ExitInstance();
};

extern const GUID CDECL _tlid;
extern const WORD _wVerMajor;
extern const WORD _wVerMinor;

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_BHMDATA_H__BF91B133_6A84_11D3_A7FE_0080C8763FA4__INCLUDED)
