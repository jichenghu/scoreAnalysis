// AnalyserDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "Analyser.h"
#include "AnalyserDlg.h"

#include "Analyzer.h"
#include "Api.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

// ����Ӧ�ó��򡰹��ڡ��˵���� CAboutDlg �Ի���

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// �Ի�������
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

// ʵ��
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
END_MESSAGE_MAP()


// CAnalyserDlg �Ի���



CAnalyserDlg::CAnalyserDlg(CWnd* pParent /*=NULL*/)
//	: CDialog(CAnalyserDlg::IDD, pParent)
	: CDialogResize(CAnalyserDlg::IDD, pParent)
{
	m_bInitialized	= FALSE	;
	m_lstItemHead	= NULL	;

	m_pAnalyzer	= new CAnalyzer	;

	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME)	;
}

void CAnalyserDlg::DoDataExchange(CDataExchange* pDX)
{
	//	CDialog::DoDataExchange(pDX);
	CDialogResize::DoDataExchange(pDX)	;

	DDX_Control(pDX, IDC_TAB, m_tabCtrl);
}

BEGIN_MESSAGE_MAP(CAnalyserDlg, CDialogResize)
	ON_WM_SYSCOMMAND()
	ON_WM_DESTROY()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_WM_SIZE()

	ON_NOTIFY(TCN_SELCHANGE, IDC_TAB, OnTcnSelchangeTab)
END_MESSAGE_MAP()


BEGIN_DLGRESIZE_MAP(CAnalyserDlg)
	//DLGRESIZE_CONTROL(IDC_OUTLOOKBAR, DLSZ_SIZE_Y)
	DLGRESIZE_CONTROL(IDC_DIALOG_AREA, DLSZ_SIZE_X | DLSZ_SIZE_Y)
	//DLGRESIZE_CONTROL(AFX_IDW_STATUS_BAR, DLSZ_MOVE_Y | DLSZ_SIZE_X)
	//DLGRESIZE_CONTROL(AFX_IDW_TOOLBAR, DLSZ_SIZE_X)
	//DLGRESIZE_CONTROL(IDC_INFOBAR, DLSZ_SIZE_X)
END_DLGRESIZE_MAP()

// CAnalyserDlg ��Ϣ�������

BOOL CAnalyserDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// ��\������...\���˵�����ӵ�ϵͳ�˵��С�

	// IDM_ABOUTBOX ������ϵͳ���Χ�ڡ�
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		CString strAboutMenu;
		strAboutMenu.LoadString(IDS_ABOUTBOX);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// ���ô˶Ի����ͼ�ꡣ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��
	
	m_tabCtrl.InsertItem ( 0, "���÷�������" )	;
	m_tabCtrl.InsertItem ( 1, "���Ƴɼ�����" )	;
	m_tabCtrl.InsertItem ( 2, "��Ƴɼ�����" )	;

	// Associate a 'tag'with the window so we can locate it later
	::SetProp( GetSafeHwnd(), AfxGetApp()->m_pszExeName, (HANDLE)1 );

	InitResizing(FALSE);

	m_lstItemHead	= NULL	;
	m_lstItemTail	= NULL	;

	// create property pages
	m_pageConfig.Create(IDD_PROPPAGE_CONFIG, this )			;
	m_pageSingle.Create(IDD_PROPPAGE_SINGLE_SUBJECT, this)	;
	m_pageMulti .Create(IDD_PROPPAGE_MULTI_SUBJECTS, this)	;
	// activate main page 
	m_iCurrentItem	= 0;
	ActivatePage( m_iCurrentItem );

	m_bInitialized	= TRUE	;

	ShowWindow(SW_MAXIMIZE)	;

	//TestCopy();

	return TRUE;  // ���������˿ؼ��Ľ��㣬���򷵻� TRUE
}

/********************************************************************/
/*																	*/
/* Function name : CApplicationDlg::MoveChilds						*/
/* Description   : Move child windows into place holder area.		*/
/*																	*/
/********************************************************************/
void CAnalyserDlg::MoveChilds()
{
	// position property pages 
	CRect rcDlgs;
	
	// get dialog area rect
	GetDlgItem(IDC_DIALOG_AREA)->GetWindowRect(rcDlgs);
	
	ScreenToClient(rcDlgs);
	
	m_tabCtrl.MoveWindow (rcDlgs)	;
	m_pageConfig.MoveWindow(rcDlgs)	;
	m_pageSingle.MoveWindow(rcDlgs)	; 
	m_pageMulti.MoveWindow(rcDlgs)	; 
}

/********************************************************************/
/*																	*/
/* Function name : CApplicationDlg::ActivatePage					*/
/* Description   : Called when an icon on the outlookbar is pressed.*/
/*																	*/
/********************************************************************/
void CAnalyserDlg::ActivatePage(int nIndex)
{

	switch(nIndex)
	{
		case 0:
			m_pageConfig.ShowWindow(SW_SHOW);
			m_pageSingle.ShowWindow(SW_HIDE); 
			m_pageMulti. ShowWindow(SW_HIDE); 
			break;
		case 1:
			m_pageConfig.ShowWindow(SW_HIDE);
			m_pageSingle.ShowWindow(SW_SHOW);
			m_pageMulti. ShowWindow(SW_HIDE);
			break;
		case 2:
			m_pageConfig.ShowWindow(SW_HIDE);
			m_pageSingle.ShowWindow(SW_HIDE);
			m_pageMulti. ShowWindow(SW_SHOW);
			//m_OnlineUsersPage.ShowWindow(SW_HIDE);	
			//m_TracePage.ShowWindow(SW_HIDE);	
			//m_MemoryUsagePage.ShowWindow(SW_HIDE);	
			//m_UserGroupPage.ShowWindow(SW_SHOW);	
			//m_ConfigurationPage.ShowWindow(SW_HIDE);	
			//m_PaperManagerPage.ShowWindow(SW_HIDE);
			//m_InfobarCtrl.SetText("�û���");
			//UpdateChannelServerState();
			break;
		case 3:
			//m_OnlineUsersPage.ShowWindow(SW_HIDE);	
			//m_TracePage.ShowWindow(SW_HIDE);	
			//m_MemoryUsagePage.ShowWindow(SW_HIDE);	
			//m_UserGroupPage.ShowWindow(SW_HIDE);	
			//m_ConfigurationPage.ShowWindow(SW_HIDE);	
			//m_PaperManagerPage.ShowWindow(SW_SHOW);
			//m_InfobarCtrl.SetText("�Ծ�״̬");
			break;
		case 4:
			//m_OnlineUsersPage.ShowWindow(SW_HIDE);	
			//m_TracePage.ShowWindow(SW_HIDE);	
			//m_MemoryUsagePage.ShowWindow(SW_SHOW);	
			//m_UserGroupPage.ShowWindow(SW_HIDE);	
			//m_ConfigurationPage.ShowWindow(SW_HIDE);	
			//m_PaperManagerPage.ShowWindow(SW_HIDE);
			//m_InfobarCtrl.SetText("�ڴ�״̬");
			break;
		case 5:
			//m_OnlineUsersPage.ShowWindow(SW_HIDE);	
			//m_TracePage.ShowWindow(SW_HIDE);	
			//m_MemoryUsagePage.ShowWindow(SW_HIDE);	
			//m_UserGroupPage.ShowWindow(SW_HIDE);	
			//m_ConfigurationPage.ShowWindow(SW_SHOW);	
			//m_PaperManagerPage.ShowWindow(SW_HIDE);
			//m_InfobarCtrl.SetText("����������");
			break;
		default:
			break;
	}

	MoveChilds();
}

void CAnalyserDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		//CDialog::OnSysCommand(nID, lParam);
		CDialogResize::OnSysCommand(nID, lParam);
	}
}

void CAnalyserDlg::OnDestroy()
{
	WinHelp(0L, HELP_QUIT);
	CDialog::OnDestroy();

	if ( NULL != m_pAnalyzer )
		delete	m_pAnalyzer	;

	DestroyAllItems();
}

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void CAnalyserDlg::OnPaint() 
{
	if (IsIconic())
	{
		CPaintDC dc(this); // ���ڻ��Ƶ��豸������

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// ʹͼ���ڹ��������о���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// ����ͼ��
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
//		CDialog::OnPaint();
		CDialogResize::OnPaint();
	}
}

//���û��϶���С������ʱϵͳ���ô˺���ȡ�ù����ʾ��
HCURSOR CAnalyserDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

/********************************************************************/
/*																	*/
/* Function name : OnSize											*/		
/* Description   : Handle WM_SIZE message							*/
/*																	*/
/********************************************************************/
void CAnalyserDlg::OnSize(UINT nType, int cx, int cy) 
{
	CDialogResize::OnSize(nType, cx, cy);
	
	if ( m_bInitialized )
		MoveChilds();
}

void CAnalyserDlg::SetPage ( int nIndex )
{
	int	iOldItem;

	iOldItem		= m_iCurrentItem		;

	switch( iOldItem )
	{
	case 0:
		m_pageConfig.OnCbnSelchangeComboRegion()	;
		break;

	case 1:
		break;

	case 2:
		break;

	default:
		break;
	}

	m_tabCtrl.SetCurSel ( nIndex )	;
	ActivatePage ( nIndex )			;

	m_iCurrentItem	= m_tabCtrl.GetCurSel()	;
	switch( m_iCurrentItem )
	{
	case 0:
		break;

	case 1:
		m_pageSingle.UpdateUI ( )	;
		m_pageSingle.UpdateData ( )	;
		break;

	case 2:
		break;

	default:
		break;
	}
}

void CAnalyserDlg::OnTcnSelchangeTab(NMHDR *pNMHDR, LRESULT *pResult)
{
	int	iOldItem;

	iOldItem		= m_iCurrentItem		;
	m_iCurrentItem	= m_tabCtrl.GetCurSel()	;

	switch(iOldItem)
	{
	case 0:
		m_pageConfig.OnCbnSelchangeComboRegion()	;
		break;

	case 1:
		break;

	case 2:
		break;

	default:
		break;
	}

	ActivatePage( m_iCurrentItem )	;

	switch(m_iCurrentItem)
	{
	case 0:
		break;

	case 1:
		m_pageSingle.UpdateUI ( )	;
		break;

	case 2:
		break;

	default:
		break;
	}

	*pResult = 0;
}

void CAnalyserDlg::AddItem ( LPITEM lpItem )
{
	ASSERT ( NULL != lpItem )	;

	lpItem->lpNext	= NULL	;

	if ( NULL == m_lstItemHead )
	{
		m_lstItemHead	= lpItem;
		m_lstItemTail	= lpItem;
	}
	else
	{
		m_lstItemTail->lpNext	= lpItem;
		m_lstItemTail			= lpItem;
	}
}

void CAnalyserDlg::DestroyAllItems ( )
{
	LPITEM		lpItem		;
	LPSUBITEM	lpSubItem	;
	LPSUBITEM	lpSubHead	;

	lpItem	= m_lstItemHead	;
	while ( NULL != lpItem )
	{
		m_lstItemHead	= lpItem->lpNext;

		lpSubHead	= lpItem->lpFirstSub;
		lpSubItem	= lpSubHead			;
		while ( NULL != lpSubItem )
		{
			lpSubHead	= lpSubItem->m_lpNext	;

			delete	lpSubItem	;

			lpSubItem	= lpSubHead	;
		}

		delete	lpItem	;

		lpItem	= m_lstItemHead	;
	}
	m_lstItemTail	= NULL	;
}

void CAnalyserDlg::SetPathOutput ( CString strPath ) 
{
	// ���Ը�·���Ƿ����
	CString	strHint	;

	if ( !TestDir (strPath, strHint) )
	{
		AfxMessageBox ( strHint )	;
	}

	m_strPathOutput	= strPath	;

	m_pAnalyzer->SetPathOutput ( strPath )	;
}
