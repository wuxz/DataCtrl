// AddColumnDlg.cpp : implementation file
//

#include "stdafx.h"
#include "bhmdata.h"
#include "AddColumnDlg.h"
#include "BHMDBColumnPropPage.h"
#include "BHMDBGroupsPropPage.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CAddColumnDlg dialog


CAddColumnDlg::CAddColumnDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CAddColumnDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CAddColumnDlg)
	m_strName = _T("");
	//}}AFX_DATA_INIT

	m_pColumnPage = NULL;
	m_pGroupsPage = NULL;
}


void CAddColumnDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CAddColumnDlg)
	DDX_Text(pDX, IDC_EDIT_COLUMNNAME, m_strName);
	//}}AFX_DATA_MAP
	if (pDX->m_bSaveAndValidate)
	{
		// We are saving data.

		if (m_strName.IsEmpty())
		{
			// Name can not be empty.
			pDX->Fail();
			return;
		}
		
		if (m_pColumnPage)
		{
			// The caller is column prop page.

			if (!m_pColumnPage->m_wndGrid.IsColNameUnique(m_strName))
			{
				// Name already exists.
				AfxMessageBox(IDS_ERROR_NAMENOTUNIQUE);
				pDX->Fail();
				return;
			}
		}
		else if (m_pGroupsPage)
		{
			// The caller is groups prop page.

			if (!m_pGroupsPage->m_wndGrid.IsGroupNameUnique(m_strName))
			{
				// The name already exists.
				AfxMessageBox(IDS_ERROR_NAMENOTUNIQUE);
				pDX->Fail();
				return;
			}
		}
	}
}


BEGIN_MESSAGE_MAP(CAddColumnDlg, CDialog)
	//{{AFX_MSG_MAP(CAddColumnDlg)
		// NOTE: the ClassWizard will add message map macros here
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CAddColumnDlg message handlers

void CAddColumnDlg::SetPagePtr(CBHMDBColumnPropPage *pPage)
/*
Routine Description:
	Set the pointer of the caller page.

Parameters:
	pPage		The pointer of the caller page.

Return Values:
	None.
*/
{
	// Save the given pointer.
	m_pColumnPage = pPage;
}

void CAddColumnDlg::SetPagePtr(CBHMDBGroupsPropPage *pPage)
/*
Routine Description:
	Set the pointer of the caller page.

Parameters:
	pPage		The pointer of the caller page.

Return Values:
	None.
*/
{
	// Save the given pointer.
	m_pGroupsPage = pPage;
}
