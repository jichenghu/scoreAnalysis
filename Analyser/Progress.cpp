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
{
	//{{AFX_DATA_INIT(CProgress)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
	Create(CProgress::IDD,pParent);
}


void CProgress::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CProgress)
	DDX_Control(pDX, IDC_PROGRESS2, m_Progress2);
	DDX_Control(pDX, IDC_PROGRESS1, m_Progress1);
	DDX_Control(pDX, IDC_EDIT2, m_Title2);
	DDX_Control(pDX, IDC_EDIT1, m_Title1);
	//}}AFX_DATA_MAP
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
	m_Progress1.SetRange(0,101);
	m_Progress2.SetRange(0,101);
	m_Progress1.SetPos(0);
	m_Progress2.SetPos(0);

	m_Progress1.ShowPercent();
	m_Progress2.ShowPercent();

	m_Progress1.SetStartColor(RGB(255,0,0));
	m_Progress1.SetEndColor(RGB(255,255,0));
	m_Progress1.SetTextColor(RGB(0,0,185));
	m_Progress1.SetBkColor(RGB(155,155,155));

	m_Progress2.SetStartColor(RGB(100,255,0));
	m_Progress2.SetEndColor(RGB(255,255,0));
	m_Progress2.SetTextColor(RGB(0,0,185));
	m_Progress2.SetBkColor(RGB(200,200,200));

	SetWindowText( "Please wait while processing score analyzation..." );
	
	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

void CProgress::PostNcDestroy() 
{
	// Free the C++ class.
	delete this;
}

void CProgress::SetRange1(int iStart, int iEnd )
{
	m_Progress1.SetRange32(iStart, iEnd);
	MSG msg;
	if (::PeekMessage((LPMSG)&msg,(HWND)NULL,(WORD)NULL,(WORD)NULL,TRUE))
	{
		TranslateMessage((LPMSG)&msg);
		DispatchMessage((LPMSG)&msg);
	}
}

void CProgress::SetRange2(int iStart, int iEnd )
{
	m_Progress2.SetRange32(iStart, iEnd);
	MSG msg;
	if (::PeekMessage((LPMSG)&msg,(HWND)NULL,(WORD)NULL,(WORD)NULL,TRUE))
	{
		TranslateMessage((LPMSG)&msg);
		DispatchMessage((LPMSG)&msg);
	}
}

void CProgress::SetPos1(int site)
{
	m_Progress1.SetPos(site);
	MSG msg;
	if (::PeekMessage((LPMSG)&msg,(HWND)NULL,(WORD)NULL,(WORD)NULL,TRUE))
	{
		TranslateMessage((LPMSG)&msg);
		DispatchMessage((LPMSG)&msg);
	}
}

void CProgress::SetPos2(int site)
{
	m_Progress2.SetPos(site);
	MSG msg;
	if (::PeekMessage((LPMSG)&msg,(HWND)NULL,(WORD)NULL,(WORD)NULL,TRUE))
	{
		TranslateMessage((LPMSG)&msg);
		DispatchMessage((LPMSG)&msg);
	}
}

void CProgress::SetTitle1 ( CString strTitle )
{
	m_Title1.SetWindowText ( strTitle )	;
}

void CProgress::SetTitle2 ( CString strTitle )
{
	m_Title2.SetWindowText ( strTitle )	;
}

void CProgress::Step1()
{
	m_Progress1.StepIt();
	MSG msg;
	if (::PeekMessage((LPMSG)&msg,(HWND)NULL,(WORD)NULL,(WORD)NULL,TRUE))
	{
		TranslateMessage((LPMSG)&msg);
		DispatchMessage((LPMSG)&msg);
	}
}

void CProgress::Step2()
{
	m_Progress2.StepIt();
	MSG msg;
	if (::PeekMessage((LPMSG)&msg,(HWND)NULL,(WORD)NULL,(WORD)NULL,TRUE))
	{
		TranslateMessage((LPMSG)&msg);
		DispatchMessage((LPMSG)&msg);
	}
}
