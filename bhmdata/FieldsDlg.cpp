// FieldsDlg.cpp : implementation file
//

#include "stdafx.h"
#include "bhmdata.h"
#include "FieldsDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CFieldsDlg dialog


CFieldsDlg::CFieldsDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CFieldsDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CFieldsDlg)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
}


void CFieldsDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CFieldsDlg)
		// NOTE: the ClassWizard will add DDX and DDV calls here
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CFieldsDlg, CDialog)
	//{{AFX_MSG_MAP(CFieldsDlg)
		// NOTE: the ClassWizard will add message map macros here
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CFieldsDlg message handlers
BOOL CFieldsDlg::OnInitDialog() 
{
	CDialog::OnInitDialog();
	
	CListBox * pList = (CListBox *)GetDlgItem(IDC_LIST_FIELDS);
	pList->ResetContent();

	for (int i = 0; i < m_arField.GetSize(); i ++)
		pList->AddString(m_arField[i]);
	
	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

void CFieldsDlg::OnOK() 
{
	m_arField.RemoveAll();

	CListBox * pList = (CListBox *)GetDlgItem(IDC_LIST_FIELDS);
	int nItems, nIndex[255];

	nItems = pList->GetSelItems(255, nIndex);
	CString strItem;
	for (int i = 0; i < nItems; i ++)
	{
		pList->GetText(nIndex[i], strItem);
		m_arField.Add(strItem);
	}

	CDialog::OnOK();
}
