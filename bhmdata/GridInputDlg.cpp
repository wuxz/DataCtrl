// GridInputDlg.cpp : implementation file
//

#include "stdafx.h"
#include "bhmdata.h"
#include "GridInputDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CGridInputDlg dialog


CGridInputDlg::CGridInputDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CGridInputDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CGridInputDlg)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
}


void CGridInputDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CGridInputDlg)
		// NOTE: the ClassWizard will add DDX and DDV calls here
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CGridInputDlg, CDialog)
	//{{AFX_MSG_MAP(CGridInputDlg)
		// NOTE: the ClassWizard will add message map macros here
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CGridInputDlg message handlers
BOOL CGridInputDlg::OnInitDialog() 
{
	CDialog::OnInitDialog();
	
	m_wndGrid.SubclassDlgItem(IDC_GRID_INPUT, this);
	m_wndGrid.SetColCount(1);

	CString strChoice;
	TCHAR ch;
	m_wndGrid.InsertRow();
	for (int i = 0, j = 0; i < m_strChoiceList.GetLength(); i ++)
	{
		ch = m_strChoiceList[i]; 
		if (ch == '\n')
		{
			m_wndGrid.SetCellText(j + 1, 1, strChoice);
			strChoice.Empty();
			m_wndGrid.InsertRow(j + 2);
			j ++;
		}
		else
			strChoice += ch;
	}
	m_wndGrid.SetCellText(j + 1, 1, strChoice);
	m_wndGrid.FlushRecord();
	
	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

void CGridInputDlg::OnOK() 
{
	CString strText;

	m_strChoiceList.Empty();

	if (m_wndGrid.GetRowCount() == 0)
		return;

	for (int i = 1; i <= m_wndGrid.GetRowCount(); i ++)
	{
		strText = m_wndGrid.GetCellText(i, 1);
		if (i < m_wndGrid.GetRowCount())
			strText += _T("\n");

		m_strChoiceList += strText;
	}
	
	CDialog::OnOK();
}
