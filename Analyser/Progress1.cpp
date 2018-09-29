// Progress.cpp : implementation file
//

#include "stdafx.h"
#include "Analyser.h"
#include "Progress.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CProgress dialog


CProgress::CProgress(CWnd* pParent /*=NULL*/)
	: CDialog(CProgress::IDD, pParent)
	, m_Titles(_T(""))
{
	//{{AFX_DATA_INIT(CProgress)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
	Create(CProgress::IDD,pParent);
	ShowWindow(SW_SHOW);
}


void CProgress::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CProgress)
	DDX_Control(pDX, IDC_PROGRESS2, m_Progress);
	//}}AFX_DATA_MAP
	DDX_Text(pDX, IDC_STATIC1, m_Titles);
}


BEGIN_MESSAGE_MAP(CProgress, CDialog)
	//{{AFX_MSG_MAP(CProgress)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CProgress message handlers

BOOL CProgress::OnInitDialog() 
{
	CDialog::OnInitDialog();
	
	
	CenterWindow();
	m_Progress.SetRange(0,101);
	m_Progress.SetPos(0);

	m_Progress.ShowPercent();

	m_Progress.SetStartColor(RGB(255,0,0));
	m_Progress.SetEndColor(RGB(255,255,0));
	m_Progress.SetTextColor(RGB(0,0,185));
	m_Progress.SetBkColor(RGB(155,155,155));

	SetWindowText( "Please wait processing recognition..." );
	
	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

void CProgress::PostNcDestroy() 
{
	// Free the C++ class.
	delete this;
}
//
//void CProgress::DestroyProgress()
//{
//	DestroyWindow();
//}

void CProgress::SetRange(int iStart, int iEnd )
{
	m_Progress.SetRange32(iStart, iEnd);
	MSG msg;
	if (::PeekMessage((LPMSG)&msg,(HWND)NULL,(WORD)NULL,(WORD)NULL,TRUE))
	{
		TranslateMessage((LPMSG)&msg);
		DispatchMessage((LPMSG)&msg);
	}
}

void CProgress::SetPos(int site)
{
	m_Progress.SetPos(site);
	MSG msg;
	if (::PeekMessage((LPMSG)&msg,(HWND)NULL,(WORD)NULL,(WORD)NULL,TRUE))
	{
		TranslateMessage((LPMSG)&msg);
		DispatchMessage((LPMSG)&msg);
	}
}

void CProgress::Step()
{
	m_Progress.StepIt();
	MSG msg;
	if (::PeekMessage((LPMSG)&msg,(HWND)NULL,(WORD)NULL,(WORD)NULL,TRUE))
	{
		TranslateMessage((LPMSG)&msg);
		DispatchMessage((LPMSG)&msg);
	}
}

void CProgress::SetTitle(CString strTitle)
{
	m_Titles	= strTitle;
	UpdateData(FALSE);
}