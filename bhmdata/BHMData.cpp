// BHMData.cpp : Implementation of CBHMDataApp and DLL registration.

#include "stdafx.h"
#include "BHMData.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


CBHMDataApp NEAR theApp;

// The window handle used in mouse hook procedure.
HWND m_hWndMain, m_hWndDropDown;

const GUID CDECL BASED_CODE _tlid =
		{ 0xbf91b122, 0x6a84, 0x11d3, { 0xa7, 0xfe, 0, 0x80, 0xc8, 0x76, 0x3f, 0xa4 } };
const WORD _wVerMajor = 1;
const WORD _wVerMinor = 0;

// Register the dropdown message
int CBHMDataApp::m_nBringMsg = RegisterWindowMessage(_T("Bring BHM DBDropDown"));

////////////////////////////////////////////////////////////////////////////
// CBHMDataApp::InitInstance - DLL initialization

BOOL CBHMDataApp::InitInstance()
{
	BOOL bInit = COleControlModule::InitInstance();

	if (bInit)
	{
		// TODO: Add your own module initialization code here.
	}

	return bInit;
}


////////////////////////////////////////////////////////////////////////////
// CBHMDataApp::ExitInstance - DLL termination

int CBHMDataApp::ExitInstance()
{
	// TODO: Add your own module termination code here.

	return COleControlModule::ExitInstance();
}


/////////////////////////////////////////////////////////////////////////////
// DllRegisterServer - Adds entries to the system registry

STDAPI DllRegisterServer(void)
{
	AFX_MANAGE_STATE(_afxModuleAddrThis);

	if (!AfxOleRegisterTypeLib(AfxGetInstanceHandle(), _tlid))
		return ResultFromScode(SELFREG_E_TYPELIB);

	if (!COleObjectFactoryEx::UpdateRegistryAll(TRUE))
		return ResultFromScode(SELFREG_E_CLASS);

	return NOERROR;
}


/////////////////////////////////////////////////////////////////////////////
// DllUnregisterServer - Removes entries from the system registry

STDAPI DllUnregisterServer(void)
{
	AFX_MANAGE_STATE(_afxModuleAddrThis);

	if (!AfxOleUnregisterTypeLib(_tlid, _wVerMajor, _wVerMinor))
		return ResultFromScode(SELFREG_E_TYPELIB);

	if (!COleObjectFactoryEx::UpdateRegistryAll(FALSE))
		return ResultFromScode(SELFREG_E_CLASS);

	return NOERROR;
}

void ReportError()
/*
Routine Description:
	Get the text message of the last com error and report it.

Parameters:
	None.

Return Value:
	None.
*/
{
	IErrorInfo *pErrorInfo = NULL;
	CString strError;
	BSTR bstrDescription = NULL;
	HINSTANCE hInstResources = AfxGetResourceHandle();

	// go and get the error information
	//
	HRESULT hr = GetErrorInfo(0, &pErrorInfo);
	if (FAILED(hr) || !pErrorInfo)
		strError.LoadString(IDS_ERROR_GENERAL);
	else
	{
		// get the source and the description
		//
		hr = pErrorInfo->GetDescription(&bstrDescription);
		pErrorInfo->Release();

		USES_CONVERSION;
		if (SUCCEEDED(hr))
		{
			strError = W2T(bstrDescription);
			SysFreeString(bstrDescription);
		}
	}

	// show the error!
	//
	AfxMessageBox(strError, MB_ICONEXCLAMATION|MB_OK|MB_TASKMODAL);
}

#ifdef _AFXDLL

// These codes are copied from MFC source code. 
// They must be here if use AFXDLL.

short AFXAPI _AfxShiftState()
{
	BOOL bShift = (GetKeyState(VK_SHIFT) < 0);
	BOOL bCtrl  = (GetKeyState(VK_CONTROL) < 0);
	BOOL bAlt   = (GetKeyState(VK_MENU) < 0);

	return (short)(bShift + (bCtrl << 1) + (bAlt << 2));
}

void AFXAPI _AfxToggleBorderStyle(CWnd* pWnd)
{
	if (pWnd->m_hWnd != NULL)
	{
		// toggle border style and force redraw of border
		::SetWindowLong(pWnd->m_hWnd, GWL_STYLE, pWnd->GetStyle() ^ WS_BORDER);
		::SetWindowPos(pWnd->m_hWnd, NULL, 0, 0, 0, 0,
			SWP_DRAWFRAME | SWP_NOSIZE | SWP_NOMOVE | SWP_NOACTIVATE);
	}
}

void AFXAPI _AfxToggleAppearance(CWnd* pWnd)
{
	if (pWnd->m_hWnd != NULL)
	{
		// toggle border style and force redraw of border
		::SetWindowLong(pWnd->m_hWnd, GWL_EXSTYLE, pWnd->GetExStyle() ^
			WS_EX_CLIENTEDGE);
		::SetWindowPos(pWnd->m_hWnd, NULL, 0, 0, 0, 0,
			SWP_DRAWFRAME | SWP_NOSIZE | SWP_NOMOVE | SWP_NOACTIVATE);
	}
}

#endif

// The mouse hook procedure.
__declspec(dllexport) LRESULT CALLBACK MouseProc(int nCode, WPARAM wParam, LPARAM lParam)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	if (nCode >=0 && nCode != HC_NOREMOVE)
	{	
//		if (wParam != WM_MOUSEMOVE)
//		{
//			TRACE("hwnd is %x", ((MOUSEHOOKSTRUCT*)lParam)->hwnd);
//			TRACE("message is %x \n", wParam);
//		}
		// if the button clicked outside the dropdown window, close it
		if ((wParam == WM_LBUTTONDOWN || wParam == WM_NCLBUTTONDOWN) && ((MOUSEHOOKSTRUCT*)lParam)->hwnd != m_hWndDropDown)
		{
			::SendMessage(m_hWndMain, WM_KEYDOWN, VK_ESCAPE, 1);
		}
	}

	return 0;
}

// Set the value of the global variable.
__declspec(dllexport) void SetHost(HWND hWndMain, HWND hWndDropDown)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	
	// Set the main window handle.
	m_hWndMain = hWndMain;

	// Set the dropdown window handle.
	m_hWndDropDown = hWndDropDown;
}
