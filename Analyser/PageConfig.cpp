// PageConfig.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "Analyser.h"
#include "PageConfig.h"
#include "AnalyserDlg.h"
#include "Bind.h"
#include "Analyzer.h"
//#include "vCnt.h"

// CPageConfig �Ի���

IMPLEMENT_DYNAMIC(CPageConfig, CDialog)
CPageConfig::CPageConfig(CWnd* pParent /*=NULL*/)
	: CDialogResize(CPageConfig::IDD, pParent)
{
	m_iRegionID		= -1	;

	m_pParentDlg	= NULL	;

	m_bReverse		= FALSE	;
}

CPageConfig::~CPageConfig()
{
}

void CPageConfig::DoDataExchange(CDataExchange* pDX)
{
	CDialogResize::DoDataExchange(pDX)	;

	DDX_Control(pDX, IDC_COMBO_REGION, m_ctlRegion)	;
	DDX_Control(pDX, IDC_CHECK1, m_chkForm1)		;
	DDX_Control(pDX, IDC_CHECK2, m_chkForm2)		;
	DDX_Control(pDX, IDC_CHECK3, m_chkForm3)		;
	DDX_Control(pDX, IDC_CHECK4, m_chkForm4)		;
	DDX_Control(pDX, IDC_CHECK5, m_chkForm5)		;
	DDX_Control(pDX, IDC_CHECK6, m_chkForm6)		;
	DDX_Control(pDX, IDC_CHECK7, m_chkForm7)		;
	DDX_Control(pDX, IDC_CHECK8, m_chkForm8)		;

	DDX_Control(pDX, IDC_REVERSE_CHECK, m_chkReverse )	;
	
	DDX_Control(pDX, IDC_EDIT_PATH, m_edPathConfig)		;
	DDX_Control(pDX, IDC_BTN_PATH, m_btnPathConfig)		;
	DDX_Control(pDX, IDC_EDIT_PATH1, m_edPathResults)	;
	DDX_Control(pDX, IDC_BTN_PATH1, m_btnPathResults)	;
}


BEGIN_MESSAGE_MAP(CPageConfig, CDialogResize)
	ON_CBN_SELCHANGE(IDC_COMBO_REGION, OnCbnSelchangeComboRegion)
	ON_WM_DESTROY()
	ON_BN_CLICKED(IDC_BTN_PATH, OnBnClickedBtnPathConfig)
	ON_BN_CLICKED(IDC_BTN_PATH1, OnBnClickedBtnPathResults)
	ON_BN_CLICKED(IDC_CHECK1, OnBnClickedCheck1)
	ON_BN_CLICKED(IDC_CHECK2, OnBnClickedCheck2)
	ON_BN_CLICKED(IDC_CHECK3, OnBnClickedCheck3)
	ON_BN_CLICKED(IDC_CHECK4, OnBnClickedCheck4)
	ON_BN_CLICKED(IDC_CHECK5, OnBnClickedCheck5)
	ON_BN_CLICKED(IDC_CHECK6, OnBnClickedCheck6)
	ON_BN_CLICKED(IDC_CHECK7, OnBnClickedCheck7)
	ON_BN_CLICKED(IDC_CHECK8, OnBnClickedCheck8)
	ON_BN_CLICKED(IDC_REVERSE_CHECK, OnBnClickedReverse)
	ON_BN_CLICKED(IDC_BTN_STEPSCORE, OnBnClickedBtnStepscore)
	ON_BN_CLICKED(IDC_BTN_CHECK_CNT, OnBnClickedBtnCheckCnt)
	ON_BN_CLICKED(IDC_BTN_STEPSCORE3, OnBnClickedBtnStepscore3)
	ON_WM_PAINT()
END_MESSAGE_MAP()

#include "setSelectID.h"
//////////////////////////////////////////////////////////////
// BindSQL()
// ��SQL����Դ���б�����Ͽ�
// ���� wnd:     �б�ؼ�������Ͽ�ؼ�����Ҫ��֧��ResetContent()��AddString()����
//      sql:     �󶨵�sql���(ע���ú���������sql�����Ч��У�飬������ȷ������ȷ��)
//               sql����������"select distinct [field] from tablename where ..."����ʽ
//      attach:  ���ӹ��ܿ���(�������'-')
//////////////////////////////////////////////////////////////
template <class T> bool BindSQL(T &wnd, CString sql, CMapStringToPtr* pMapID, bool bAttach)
{
	CString			str		;
	CSetSelectID	setID	;

	setID.m_pDatabase	= &theApp.m_database;

	wnd.ResetContent ( )	;

	POSITION	pos		;
	CString		key		;
	int *		piID	;

	for( pos = pMapID->GetStartPosition(); pos != NULL; )
	{
		pMapID->GetNextAssoc( pos, key, (void*&)piID );
#ifdef _DEBUG
//		afxDump << key << " : " << lpLevel << "\n";
#endif
		delete	piID;
	}
	pMapID->RemoveAll ( )	;

	try
	{
		setID.Open ( CRecordset::snapshot, sql, CRecordset::readOnly )	;

		while ( !setID.IsEOF() )
		{
			wnd.AddString ( setID.m_strName )	;

			piID	= new int	;

			*piID	= setID.m_iID	;

			pMapID->SetAt( setID.m_strName, piID );

			setID.MoveNext ( )	;
		}

		setID.Close ( )	;
	}
	catch ( CDBException * e )
	{
		AfxMessageBox ( e->m_strError )	;

		return	false;
	}

	//ѡ���һ��
	if( wnd.GetCount() > 0 ) 
		wnd.SetCurSel( 0 )	;

	return true;
}



// CPageConfig ��Ϣ�������

BOOL CPageConfig::OnInitDialog()
{
	CDialogResize::OnInitDialog();

	for ( int i=IDC_CHECK1; i<=IDC_CHECK8; i++ )
	{
		GetDlgItem ( i )->ShowWindow ( SW_HIDE );
	}

	::BindSQL(m_ctlRegion,   "select regionID, regionName from vRegion ORDER BY regionID", &m_mapRegionID, false)		;

	CString	strPath	;
	strPath	= theApp.GetProfileString("Config", "m_strPathConfig");
	m_edPathConfig.SetWindowText( strPath )	;
	if ( 0 != strPath.GetLength() )
	{
//		m_pParentDlg->DestroyAllItems ();
		if ( ReadConfig ( strPath ) )
			ResetConfig ( )	;	// ����ʧ������
	}

	strPath	= theApp.GetProfileString("Config", "m_strPathResults");
	m_edPathResults.SetWindowText( strPath )	;
	if ( 0 != strPath.GetLength() )
		m_pParentDlg->SetPathOutput ( strPath )	;

	OnCbnSelchangeComboRegion();

	return TRUE;  // return TRUE unless you set the focus to a control
	// �쳣: OCX ����ҳӦ���� FALSE
}

void CPageConfig::OnDestroy()
{
	CDialogResize::OnDestroy();

	// free m_mapRegionID
	POSITION	pos		;
	CString		key		;
	int *		lpiID	;
	// Iterate through the entire map, dumping both name and age.
	for( pos = m_mapRegionID.GetStartPosition(); pos != NULL; )
	{
		m_mapRegionID.GetNextAssoc( pos, key, (void*&)lpiID );
#ifdef _DEBUG
//		afxDump << key << " : " << lpLevel << "\n";
#endif
		delete	lpiID;
	}
}

void CPageConfig::OnCbnSelchangeComboRegion()
{
	int *	lpiID		;

	m_ctlRegion.GetLBText(m_ctlRegion.GetCurSel(), m_strRegion);

	m_mapRegionID.Lookup( m_strRegion, (void*&)lpiID )	;

	m_iRegionID	= *lpiID	;
}

void CPageConfig::OnBnClickedBtnPathConfig()
{
	// szFilters is a text string that includes two file name filters:
	// "*.my" for "MyType Files" and "*.*' for "All Files."
	char szFilters[]=
		"ConfigType Files (*.ini)|*.ini|All Files (*.*)|*.*||";

   // ���������ļ���·��
	// Create an Open dialog; the default file name extension is ".my".
	CFileDialog fileDlg (TRUE, "��ѡ���ʼ���ļ�", "*.ini",
		OFN_FILEMUSTEXIST| OFN_HIDEREADONLY, szFilters, this);

	// Display the file dialog. When user clicks OK, fileDlg.DoModal() 
	// returns IDOK.
	if( fileDlg.DoModal ()==IDOK )
	{
		if ( IDYES != MessageBox("ȷ�ϵ��������ļ�", "��ȷ��", MB_YESNOCANCEL) )
			return	;
			
		CString pathName = fileDlg.GetPathName();

		// Implement opening and reading file in here.
		m_edPathConfig.SetWindowText( pathName )	;
		
		//Change the window's title to the opened file's title.
//		CString fileName = fileDlg.GetFileTitle ();

//		SetWindowText(fileName);

		m_pParentDlg->DestroyAllItems ();

		ReadConfig ( pathName )	;
		ResetConfig ( )			;

		theApp.WriteProfileString("Config", "m_strPathConfig", pathName);
	}

}

void CPageConfig::OnBnClickedBtnPathResults()
{
	CString		str;
	BROWSEINFO	bi;
	char		name[MAX_PATH];

	ZeroMemory( &bi, sizeof(BROWSEINFO) )	;
	bi.hwndOwner		= GetSafeHwnd()		;
	bi.pszDisplayName	= name				;
	bi.lpszTitle		= "��ѡ�񴢴����������ļ��У�";
	bi.ulFlags			= BIF_USENEWUI;

	LPITEMIDLIST idl	= SHBrowseForFolder(&bi);
	if( idl == NULL )
		return;
	SHGetPathFromIDList(idl,str.GetBuffer(MAX_PATH));
	str.ReleaseBuffer();

	if( str.GetLength()<=0 )
	{
		AfxMessageBox( "·����Ч��\n  ʧ�ܣ�" )	;
		return;
	}

	m_edPathResults.SetWindowText( str )	;

	m_pParentDlg->SetPathOutput ( str )	;

	theApp.WriteProfileString("Config", "m_strPathResults", str);
}

BOOL CPageConfig::ReadConfig ( CString strPath )
{
	BYTE		yItemNumber	;	// ��Ŀ���
	CStdioFile	file		;
	CString		strLine		;
	CString		strRight	;
	int			iPos		;

	//GetCurrentDirectory( _MAX_PATH, cPath )	;
	//strPath.Format ( "%s", cPath )			;
	//strPath	+= "\\Config.ini"				;
	
	CFileException	e;
	if ( !file.Open ( strPath, CFile::modeRead | CFile::typeText, &e ) )	
	{
#ifdef _DEBUG
		afxDump << "File could not be opened " << e.m_cause << "\n";
#else
		AfxMessageBox ( strPath + " ��ʧ��\n" )	;
#endif

		return	FALSE;
	}

	// ���ļ������ж������ݽ��н���
	yItemNumber	= 0	;
	while ( file.ReadString ( strLine ) )
	{
		strLine.TrimLeft(); strLine.TrimRight();
		strLine.MakeLower ( );

		// ���˵�����
		if ( 0 == strLine.GetLength () )
			continue;

		// ���˵�ע����
		if ( 0 == strLine.Find ( "//" ) )
			continue;

		// �ҵ�һ����
		if ( 0 == strLine.Find ( "begin" ) )
		{
			// �������
			strRight	= strLine.Right ( strLine.GetLength() - 5 )	;
			iPos		= strRight.Find ( "{" )	;
			if ( -1 == iPos )
			{
				AfxMessageBox ( "�������ļ� Config.ini ʱ��ǰ�ñ�ʶ { ʧ��" )	;
				file.Close ()	;
				return	FALSE	;
			}

			strRight	= strRight.Right ( strRight.GetLength() - iPos - 1 );

			iPos		= strRight.Find ( "}" )	;
			if ( -1 == iPos )
			{
				AfxMessageBox ( "�������ļ� Config.ini ʱ��ǰ�ñ�ʶ } ʧ��" )	;
				file.Close ()	;
				return	FALSE	;
			}

			strRight	= strRight.Left ( iPos );
			strRight.TrimLeft();	strRight.TrimRight();

			if ( 0 == strRight.Compare ( "form" ) )
			{
				ReadItem ( &file, yItemNumber )	;
				yItemNumber ++	;
			}

			continue;
		}
	}

	file.Close ( )	;

	return	TRUE;
}

BOOL CPageConfig::ReadItem ( CStdioFile * pFile, BYTE yItemNumber )
{
	CString		strLine		;
	CString		strRight	;
	int			iPos		;
	LPITEM		lpItem		;

	lpItem	= new ITEM	;

	lpItem->bSelected	= TRUE			;
	lpItem->yItemNumber	= yItemNumber	;
	lpItem->lpFirstSub	= NULL			;

	while ( pFile->ReadString ( strLine ) )
	{
		strLine.TrimLeft(); strLine.TrimRight();
		strLine.MakeLower ( );

		// ���˵�����
		if ( 0 == strLine.GetLength () )
			continue;

		// ���˵�ע����
		if ( 0 == strLine.Find ( "//" ) )
			continue;

		// �ҵ�һ����
		if ( 0 == strLine.Find ( "end" ) )
		{
			m_pParentDlg->AddItem ( lpItem )	;

			return	TRUE;
		}

		if ( 0 == strLine.Find ( "formclass" ) )
		{
			// �������
			iPos		= strLine.Find ( "=" )	;
			if ( -1 == iPos )
			{
				AfxMessageBox ( "�������ļ� Config.ini ʱTITLE ���δ���� = " )	;
				delete	lpItem	;
				return	FALSE	;
			}

			strRight	= strLine.Right ( strLine.GetLength() - iPos - 1 );

			strRight.TrimLeft ( )	;

			if ( 0 == strRight.Compare ("1") )
			{
				// ���Ʒ���
				lpItem->yFormClass	= 1	;
			}
			else if ( 0 == strRight.Compare ("2") )
			{
				// ��Ʒ���
				lpItem->yFormClass	= 2	;
			}
			else	// ȱʡ
				lpItem->yFormClass	= 1	;

			continue;
		}

		// �����Ĳ㼶(����У����)
		if ( 0 == strLine.Find ( "layers" ) )
		{
			// �������
			iPos		= strLine.Find ( "=" )	;
			if ( -1 == iPos )
			{
				AfxMessageBox ( "�������ļ� Config.ini ʱ LAYERS ���δ���� = " )	;
				delete	lpItem	;
				return	FALSE	;
			}

			strRight	= strLine.Right ( strLine.GetLength() - iPos - 1 );

			strRight.TrimLeft ( )	;
			strRight.TrimRight ( )	;

			// ��:=0  ��У:=1  У:=2  У��:=3  ��:=4  ��У��:=5
			if ( 0 == strRight.Compare ("��") )
			{
				lpItem->yLayers	= 0	;
			}
			else if ( 0 == strRight.Compare ("��У") )
			{
				lpItem->yLayers	= 1	;
			}
			else if ( 0 == strRight.Compare ("У") )
			{
				lpItem->yLayers	= 2	;
			}
			else if ( 0 == strRight.Compare ("У��") )
			{
				lpItem->yLayers	= 3	;
			}
			else if ( 0 == strRight.Compare ("��") )
			{
				lpItem->yLayers	= 4	;
			}
			else if ( 0 == strRight.Compare ("��У��") )
			{
				lpItem->yLayers	= 5	;
			}
			else	// ȱʡ
				lpItem->yLayers	= 0	;

			continue;
		}

		// �ҵ�һ���ӿ�
		if ( 0 == strLine.Find ( "begin" ) )
		{
			// �����ӿ�
			strRight	= strLine.Right ( strLine.GetLength() - 5 )	;
			iPos		= strRight.Find ( "{" )	;
			if ( -1 == iPos )
			{
				AfxMessageBox ( "�������ļ� Config.ini ʱ�ӿ���ñ�ʶ { ʧ��" )	;
				return	FALSE	;
			}

			strRight	= strRight.Right ( strRight.GetLength() - iPos - 1 );

			iPos		= strRight.Find ( "}" )	;
			if ( -1 == iPos )
			{
				AfxMessageBox ( "�������ļ� Config.ini ʱ�ӿ���ñ�ʶ } ʧ��" )	;
				return	FALSE	;
			}

			strRight	= strRight.Left ( iPos );
			strRight.TrimLeft();	strRight.TrimRight();

			if ( 0 == strRight.Compare ( "sub" ) )
			{
				ReadSub ( pFile, lpItem )	;
			}

			continue;
		}

		if ( 0 == strLine.Find ( "title" ) )
		{
			// �������
			iPos		= strLine.Find ( "=" )	;
			if ( -1 == iPos )
			{
				AfxMessageBox ( "�������ļ� Config.ini ʱTITLE ���δ���� = " )	;
				delete	lpItem	;
				return	FALSE	;
			}

			strRight	= strLine.Right ( strLine.GetLength() - iPos - 1 );

			strRight.TrimLeft ( )	;

			::strcpy ( lpItem->cTitle, strRight )	;

			continue;
		}
	}

	delete	lpItem	;
	return	FALSE	;
}

#define	MAIN		1
#define	SUB_FIRST	2
#define	SUB			4
BOOL CPageConfig::ReadSel ( CStdioFile *pFile, paraStruct ** lppPara, CWordAction ** lppAction, LPSUBITEM lpSubItem )
{
	CString		strLine		;
	CString		strRight	;
	CString		strLeft		;
	int			iPos		;
	int			iState		;

	CWordAction *	lpPrevAction	;
	CWordAction *	lpPrevSub		;
	CWordAction *	lpAction		;

	lpPrevAction	= *lppAction	;

	iState	= MAIN	;
	while ( pFile->ReadString ( strLine ) )
	{
		strLine.TrimLeft(); strLine.TrimRight();
		strLine.MakeLower ( );

		// ���˵�����
		if ( 0 == strLine.GetLength () )
			continue;

		// ���˵�ע����
		if ( 0 == strLine.Find ( "//" ) )
			continue;

		// ���˵�������ʾ��
		if ( 0 == strLine.Find ( "{" ) )
			continue;

		// �ҵ�������־ }
		if ( 0 == strLine.Find ( "}" ) )
		{
			// end sub

			return	TRUE;
		}

		// �ҵ��µ�����
		if ( 0 == strLine.Find ( "[" ) )
		{
			if ( (SUB_FIRST == iState) || (SUB == iState) )
			{
				AfxMessageBox ( "��������2��[����" )	;
				return	FALSE;
			}
		
			iState	= SUB_FIRST;
			continue;
		}

		// ���������־
		if ( 0 == strLine.Find ( "]" ) )
		{
			if ( (SUB_FIRST != iState) && (SUB != iState) )
			{
				AfxMessageBox ( "]���Ź���" )	;
				return	FALSE;
			}
		
			iState	= MAIN	;
			continue;
		}

		// �ҵ� HomeKey ����
		if ( 0 == strLine.Find ( "homekey" ) )
		{
			lpAction	= new CWordHomeKey(lppPara);

			iPos		= 8	;	// ���˵���HomeKey
			strRight	= strLine.Right ( strLine.GetLength() - iPos )	;

			iPos		= strRight.Find ( "," )	;
			strLeft		= strRight.Left ( iPos );
			strRight	= strRight.Right ( strRight.GetLength() - iPos - 1 )	;

			lpAction->m_vVar1	= COleVariant( (short)atoi( strLeft ) )	;
			lpAction->m_vVar2	= COleVariant( (short)atoi( strRight ) );

			lpAction->m_pNext		= NULL	;
			lpAction->m_pSubNext	= NULL	;

			if ( NULL == lpPrevAction )
			{
				*lppAction	= lpAction	;

				if ( SUB_FIRST == iState )
				{
					iState		= SUB		;
					lpPrevSub	= lpAction	;
				}
				
				lpPrevAction	= lpAction	;
			}
			else
			{
				switch ( iState )
				{
				case	SUB_FIRST:
					iState					= SUB		;
					lpPrevSub				= lpAction	;

					lpPrevAction->m_pNext	= lpAction	;
					lpPrevAction			= lpAction	;
					break;

				case	SUB:
					lpPrevSub->m_pSubNext	= lpAction	;
					lpPrevSub				= lpAction	;
					break;

				case	MAIN:
					lpPrevAction->m_pNext	= lpAction	;
					lpPrevAction			= lpAction	;
					break;

				default:
					break;
				}
			}

			continue;
		}

		// �ҵ� HomeKey ����
		if ( 0 == strLine.Find ( "formatprint" ) )
		{
			lpAction	= new CWordFormatPrint(lppPara);

			iPos		= 12	;	// ���˵���FormatPrint
			strRight	= strLine.Right ( strLine.GetLength() - iPos )	;

			strRight.TrimLeft ();

			if ( 0 == strRight.Find ( "mark" ) )
			{
				iPos	= 5		;	// ���˵� Mark
				strRight= strRight.Right ( strRight.GetLength() - iPos );

				strRight.TrimLeft();
				strRight= strRight.Right ( 7 );	
				strRight.TrimRight();
				//strcpy ( (*lppPara)->cFormat, strRight ); 
				strcpy ( ((CWordFormatPrint*)lpAction)->m_cFormat, strRight ); 
				
				lpAction->m_dwVariant	= MARK	;
			}

			lpAction->m_pNext		= NULL	;
			lpAction->m_pSubNext	= NULL	;

			if ( NULL == lpPrevAction )
			{
				*lppAction	= lpAction	;

				if ( SUB_FIRST == iState )
				{
					iState		= SUB		;
					lpPrevSub	= lpAction	;
				}
				
				lpPrevAction	= lpAction	;
			}
			else
			{
				switch ( iState )
				{
				case	SUB_FIRST:
					iState					= SUB		;
					lpPrevSub				= lpAction	;

					lpPrevAction->m_pNext	= lpAction	;
					lpPrevAction			= lpAction	;
					break;

				case	SUB:
					lpPrevSub->m_pSubNext	= lpAction	;
					lpPrevSub				= lpAction	;
					break;

				case	MAIN:
					lpPrevAction->m_pNext	= lpAction	;
					lpPrevAction			= lpAction	;
					break;

				default:
					break;
				}
			}

			continue;
		}

		// �ҵ� TypeText ����
		if ( 0 == strLine.Find ( "typetext" ) )
		{
			lpAction	= new CWordTypeText(lppPara);

			iPos		= 9	;	// ���˵� TypeText
			strRight	= strLine.Right ( strLine.GetLength() - iPos )	;

			strRight.TrimLeft();	strRight.TrimRight();
			
			if ( 0 == strRight.Find ( "classname" ) )
			{
				lpAction->m_dwVariant	= CLASS_NAME	;
			}
			else if ( 0 == strRight.Find ( "schoolname" ) )
			{
				lpAction->m_dwVariant	= SCHOOL_NAME	;
			}
			else if ( 0 == strRight.Find ( "regionname" ) )
			{
				lpAction->m_dwVariant	= REGION_NAME	;
			}
			else if ( 0 == strRight.Find ( "enrollid" ) )
			{
				lpAction->m_dwVariant	= ENROLL_ID		;
			}
			else if ( 0 == strRight.Find ( "studentname" ) )
			{
				lpAction->m_dwVariant	= STUDENT_NAME	;
			}
			else
			{
				lpAction->m_dwVariant	= NOTHING		;
			}

			lpAction->m_pNext		= NULL	;
			lpAction->m_pSubNext	= NULL	;

			if ( NULL == lpPrevAction )
			{
				*lppAction	= lpAction	;

				if ( SUB_FIRST == iState )
				{
					iState		= SUB		;
					lpPrevSub	= lpAction	;
				}
				
				lpPrevAction	= lpAction	;
			}
			else
			{
				switch ( iState )
				{
				case	SUB_FIRST:
					iState					= SUB		;
					lpPrevSub				= lpAction	;

					lpPrevAction->m_pNext	= lpAction	;
					lpPrevAction			= lpAction	;
					break;

				case	SUB:
					lpPrevSub->m_pSubNext	= lpAction	;
					lpPrevSub				= lpAction	;
					break;

				case	MAIN:
					lpPrevAction->m_pNext	= lpAction	;
					lpPrevAction			= lpAction	;
					break;

				default:
					break;
				}
			}

			continue;
		}

		// �ҵ� MoveDown ����
		if ( 0 == strLine.Find ( "movedown" ) )
		{
			lpAction	= new CWordMoveDown(lppPara);

			iPos		= 9	;	// ���˵� MoveDown
			strRight	= strLine.Right ( strLine.GetLength() - iPos )	;

			iPos		= strRight.Find ( "," )	;
			strLeft		= strRight.Left ( iPos );
			strRight	= strRight.Right ( strRight.GetLength() - iPos - 1 )	;

			lpAction->m_vVar1	= COleVariant( (short)atoi( strLeft ) )	;

			iPos		= strRight.Find ( "," )	;
			strLeft		= strRight.Left ( iPos );
			strRight	= strRight.Right ( strRight.GetLength() - iPos - 1 )	;

			lpAction->m_vVar2	= COleVariant( (short)atoi( strLeft ) )	;
			lpAction->m_vVar3	= COleVariant( (short)atoi( strRight ) );

			lpAction->m_pNext	= NULL	;
			lpAction->m_pSubNext= NULL	;

			if ( NULL == lpPrevAction )
			{
				*lppAction	= lpAction	;

				if ( SUB_FIRST == iState )
				{
					iState		= SUB		;
					lpPrevSub	= lpAction	;
				}
				
				lpPrevAction	= lpAction	;
			}
			else
			{
				switch ( iState )
				{
				case	SUB_FIRST:
					iState					= SUB		;
					lpPrevSub				= lpAction	;

					lpPrevAction->m_pNext	= lpAction	;
					lpPrevAction			= lpAction	;
					break;

				case	SUB:
					lpPrevSub->m_pSubNext	= lpAction	;
					lpPrevSub				= lpAction	;
					break;

				case	MAIN:
					lpPrevAction->m_pNext	= lpAction	;
					lpPrevAction			= lpAction	;
					break;

				default:
					break;
				}
			}

			continue;
		}

		// �ҵ� MoveRight ����
		if ( 0 == strLine.Find ( "moveright" ) )
		{
			lpAction	= new CWordMoveRight(lppPara);

			iPos		= 10	;	// ���˵� MoveRight
			strRight	= strLine.Right ( strLine.GetLength() - iPos )	;

			iPos		= strRight.Find ( "," )	;
			strLeft		= strRight.Left ( iPos );
			strRight	= strRight.Right ( strRight.GetLength() - iPos - 1 )	;

			lpAction->m_vVar1	= COleVariant( (short)atoi( strLeft ) )	;

			iPos		= strRight.Find ( "," )	;
			strLeft		= strRight.Left ( iPos );
			strRight	= strRight.Right ( strRight.GetLength() - iPos - 1 )	;

			lpAction->m_vVar2	= COleVariant( (short)atoi( strLeft ) )	;
			lpAction->m_vVar3	= COleVariant( (short)atoi( strRight ) );

			lpAction->m_pNext	= NULL	;
			lpAction->m_pSubNext= NULL	;

			if ( NULL == lpPrevAction )
			{
				*lppAction	= lpAction	;

				if ( SUB_FIRST == iState )
				{
					iState		= SUB		;
					lpPrevSub	= lpAction	;
				}
				
				lpPrevAction	= lpAction	;
			}
			else
			{
				switch ( iState )
				{
				case	SUB_FIRST:
					iState					= SUB		;
					lpPrevSub				= lpAction	;

					lpPrevAction->m_pNext	= lpAction	;
					lpPrevAction			= lpAction	;
					break;

				case	SUB:
					lpPrevSub->m_pSubNext	= lpAction	;
					lpPrevSub				= lpAction	;
					break;

				case	MAIN:
					lpPrevAction->m_pNext	= lpAction	;
					lpPrevAction			= lpAction	;
					break;

				default:
					break;
				}
			}

			continue;
		}

		// �ҵ���������
		if ( 0 == strLine.Find ( "function" ) )
		{
			lpAction	= new CWordActFunction(lppPara);

			iPos		= 9	;
			strRight	= strLine.Right ( strLine.GetLength() - iPos )	;
			strRight.TrimLeft();	strRight.TrimRight();

			// �ҵ��ú���
			CActFunction *	lpFunc	;
			lpFunc	= lpSubItem->m_funcList	;
			while ( NULL != lpFunc )
			{
				if ( 0 == lpFunc->m_strFuncName.Compare ( strRight ) )
				{
					lpAction->m_pFunc	= lpFunc	;

					break	;
				}
				lpFunc	= lpFunc->m_lpNext	;
			}

			if ( NULL == lpFunc )
			{
				AfxMessageBox ( strRight + "\n�޷��ڽű����ҵ��ú���" )	;
			}

			lpAction->m_pNext	= NULL	;
			lpAction->m_pSubNext= NULL	;

			if ( NULL == lpPrevAction )
			{
				*lppAction	= lpAction	;

				if ( SUB_FIRST == iState )
				{
					iState		= SUB		;
					lpPrevSub	= lpAction	;
				}
				
				lpPrevAction	= lpAction	;
			}
			else
			{
				switch ( iState )
				{
				case	SUB_FIRST:
					iState					= SUB		;
					lpPrevSub				= lpAction	;

					lpPrevAction->m_pNext	= lpAction	;
					lpPrevAction			= lpAction	;
					break;

				case	SUB:
					lpPrevSub->m_pSubNext	= lpAction	;
					lpPrevSub				= lpAction	;
					break;

				case	MAIN:
					lpPrevAction->m_pNext	= lpAction	;
					lpPrevAction			= lpAction	;
					break;

				default:
					break;
				}
			}

			continue;
		}
	}

	return	FALSE;
}

BOOL CPageConfig::ReadMaxRowOfBlock ( CStdioFile *pFile,
									 LPSUBITEM lpSubItem )
{
	CString		strLine		;
	CString		strLeft		;
	int			iPos		;


	while ( pFile->ReadString ( strLine ) )
	{
		strLine.TrimLeft(); strLine.TrimRight();
		strLine.MakeLower ( );

		// ���˵�����
		if ( 0 == strLine.GetLength () )
			continue;

		// ���˵�ע����
		if ( 0 == strLine.Find ( "//" ) )
			continue;

		// ���˵�������ʾ��
		if ( 0 == strLine.Find ( "{" ) )
			continue;

		// �ҵ�������־ }
		if ( 0 == strLine.Find ( "}" ) )
		{
			// end sub

			return	TRUE;
		}

		// ÿ��block�е�����
		if ( 0 == strLine.Find ( "|" ) )
		{
			int	iCnt	= 0	;
			int	*lpiMax		;

			lpiMax	= lpSubItem->m_iMaxColOfBlock	;

			iPos	= strLine.Find ( '|' )	;
			strLine	= strLine.Right ( strLine.GetLength() - iPos - 1 )	;
			iPos	= strLine.Find ( ',' )	;

			while ( -1 != iPos )
			{
				iCnt++	;
				if ( 8 == iCnt )	// ���ֻ�ܷ�8���� m_iMaxColwOfBlock[8] ��
					break	;

				strLeft	= strLine.Left ( iPos )								;
				strLine	= strLine.Right ( strLine.GetLength() - iPos - 1 )	;

				*lpiMax++	= atoi ( strLeft )	;

				iPos	= strLine.Find ( ',' )	;
			}

			continue;
		}

		// ÿ������Ӧ��������
		if ( 0 == strLine.Find ( ":" ) )
		{
			int	iCnt	= 0	;
			int	*lpiMax		;

			lpiMax	= lpSubItem->m_iMaxSubrowOfRow	;

			iPos	= strLine.Find ( ':' )	;
			strLine	= strLine.Right ( strLine.GetLength() - iPos - 1 )	;
			iPos	= strLine.Find ( ',' )	;

			while ( -1 != iPos )
			{
				iCnt++	;
				if ( 256 == iCnt )	// ���ֻ�ܷ�256���� m_iMaxRowOfBlock[256] ��
					break	;

				strLeft	= strLine.Left ( iPos )								;
				strLine	= strLine.Right ( strLine.GetLength() - iPos - 1 )	;

				*lpiMax++	= atoi ( strLeft )	;

				iPos	= strLine.Find ( ',' )	;
			}

			continue;
		}

		// block�ĸ���
		if ( 0 == strLine.Find ( ";" ) )
		{
			iPos	= strLine.Find ( ';' )	;
			strLine	= strLine.Right ( strLine.GetLength() - iPos - 1 )	;
			iPos	= strLine.Find ( ',' )	;

			strLeft	= strLine.Left ( iPos )								;
			lpSubItem->m_iCntBlocks	= atoi ( strLeft )	;


			continue;
		}

		// ÿblock�е�����
		if ( 0 != strLine.GetLength () )
		{
			int	iCnt	= 0	;
			int	*lpiMax		;

			lpiMax	= lpSubItem->m_iMaxRowOfBlock	;
			iPos	= strLine.Find ( ',' )			;

			while ( -1 != iPos )
			{
				iCnt++	;
				if ( 8 == iCnt )	// ���8���� m_iMaxRowOfBlock[8] ��
					break	;

				strLeft	= strLine.Left ( iPos )								;
				strLine	= strLine.Right ( strLine.GetLength() - iPos - 1 )	;

				*lpiMax++	= atoi ( strLeft )	;

				iPos	= strLine.Find ( ',' )	;
			}
		}

	}
	return	FALSE;
}

BOOL CPageConfig::ReadSub ( CStdioFile * pFile, LPITEM lpItem )
{
	CString		strLine		;
	CString		strRight	;
	int			iPos		;
	LPSUBITEM	lpSubItem	;
	LPSUBITEM	lpLastSub	;

	lpSubItem	= new CSubItem	;

	lpSubItem->m_lpNext			= NULL	;
	lpSubItem->m_listlpTitle	= NULL	;
	lpSubItem->m_listlpBody		= NULL	;
	lpSubItem->m_listlpColumn	= NULL	;
	lpSubItem->m_listlpRow		= NULL	;

	lpSubItem->m_funcList		= NULL	;

	lpSubItem->cSQL1[0]	= '0'	;
	lpSubItem->cSQL2[0]	= '0'	;

	if ( NULL == lpItem->lpFirstSub )
		lpItem->lpFirstSub	= lpSubItem	;
	else
	{
		lpLastSub	= lpItem->lpFirstSub;
		while ( NULL != lpLastSub->m_lpNext )
			lpLastSub	= lpLastSub->m_lpNext	;

		lpLastSub->m_lpNext	= lpSubItem	;
	}

	while ( pFile->ReadString ( strLine ) )
	{
		strLine.TrimLeft(); strLine.TrimRight();
		strLine.MakeLower ( );

		// ���˵�����
		if ( 0 == strLine.GetLength () )
			continue;

		// ���˵�ע����
		if ( 0 == strLine.Find ( "//" ) )
			continue;

		// �ҵ�һ����
		if ( 0 == strLine.Find ( "end" ) )
		{
			// end sub

			return	TRUE;
		}

		if ( 0 == strLine.Find ( "subject" ) )
		{
			// �������
			iPos		= strLine.Find ( "=" )	;
			if ( -1 == iPos )
			{
				AfxMessageBox ( "�������ļ� Config.ini ʱ Subject ���δ���� = " )	;
				delete	lpItem	;
				return	FALSE	;
			}

			strRight	= strLine.Right ( strLine.GetLength() - iPos - 1 );

			strRight.TrimLeft ( )	;
//			strRight.TrimRight ( )	;
			strRight	= strRight.Left ( 4 );

			::strcpy ( lpSubItem->cSubject, strRight )	;

			continue;
		}

		if ( 0 == strLine.Find ( "doctitle" ) )
		{
			// �������
			iPos		= strLine.Find ( "=" )	;
			if ( -1 == iPos )
			{
				AfxMessageBox ( "�������ļ� Config.ini ʱ DOCTITLE ���δ���� = " )	;
				delete	lpItem	;
				return	FALSE	;
			}

			strRight	= strLine.Right ( strLine.GetLength() - iPos - 1 );

			strRight.TrimLeft ( )	;

			::strcpy ( lpSubItem->cTitle, strRight )	;

			continue;
		}

		// �� Sub ��ʹ�õ� function
		if ( 0 == strLine.Find ( "function" ) )
		{
			iPos	= 9;
			strRight	= strLine.Right ( strLine.GetLength() - iPos )	;
			strRight.TrimLeft ( )	;

			ReadFunction ( pFile, lpSubItem, strRight )	;
		}

		if ( 0 == strLine.Find ( "title" ) )
		{
			ReadSel ( pFile, &lpSubItem->m_pPara, &lpSubItem->m_listlpTitle, lpSubItem );

			continue;
		}

		if ( 0 == strLine.Find ( "head" ) )
		{
			ReadSel ( pFile, &lpSubItem->m_pPara, &lpSubItem->m_listlpHead, lpSubItem );

			continue;
		}

		if ( 0 == strLine.Find ( "body" ) )
		{
			ReadSel ( pFile, &lpSubItem->m_pPara, &lpSubItem->m_listlpBody, lpSubItem )	;

			continue;
		}

		if ( 0 == strLine.Find ( "row" ) )
		{
			ReadSel ( pFile, &lpSubItem->m_pPara, &lpSubItem->m_listlpRow, lpSubItem )	;

			continue;
		}

		if ( 0 == strLine.Find ( "block" ) )
		{
			ReadMaxRowOfBlock ( pFile, lpSubItem )	;
			continue;
		}

		if ( 0 == strLine.Find ( "column" ) )
		{
			ReadSel ( pFile, &lpSubItem->m_pPara, &lpSubItem->m_listlpColumn, lpSubItem )	;

			continue;
		}

		if ( 0 == strLine.Find ( "sql2" ) )
		{
			// �������
			iPos		= strLine.Find ( "=" )	;
			if ( -1 == iPos )
			{
				AfxMessageBox ( "�������ļ� Config.ini ʱ SQL ���δ���� = " )	;
				delete	lpItem	;
				return	FALSE	;
			}

			strRight	= strLine.Right ( strLine.GetLength() - iPos - 1 );

			strRight.TrimLeft ( )	;
			strRight.TrimRight ( )	;
			strRight	= strRight.Left ( _MAX_SQL );

			::strcpy ( lpSubItem->cSQL2, strRight )	;

			lpSubItem->cSort2[0]	= '0'	;
			lpSubItem->cFilter2[0]	= '0'	;

			continue;
		}

		if ( 0 == strLine.Find ( "sort2" ) )
		{
			// �������
			iPos		= strLine.Find ( "=" )	;
			if ( -1 == iPos )
			{
				AfxMessageBox ( "�������ļ� Config.ini ʱ SQL ���δ���� = " )	;
				delete	lpItem	;
				return	FALSE	;
			}

			strRight	= strLine.Right ( strLine.GetLength() - iPos - 1 );

			strRight.TrimLeft ( )	;
			strRight.TrimRight ( )	;
			strRight	= strRight.Left ( _MAX_SORT );

			::strcpy ( lpSubItem->cSort2, strRight )	;

			continue;
		}

		if ( 0 == strLine.Find ( "filter2" ) )
		{
			// �������
			iPos		= strLine.Find ( "=" )	;
			if ( -1 == iPos )
			{
				AfxMessageBox ( "�������ļ� Config.ini ʱ SQL ���δ���� = " )	;
				delete	lpItem	;
				return	FALSE	;
			}

			strRight	= strLine.Right ( strLine.GetLength() - iPos - 1 );

			strRight.TrimLeft ( )	;
			strRight.TrimRight ( )	;
			strRight	= strRight.Left ( _MAX_SORT );

			::strcpy ( lpSubItem->cFilter2, strRight )	;

			continue;
		}

		if ( 0 == strLine.Find ( "sql1" ) )
		{
			// �������
			iPos		= strLine.Find ( "=" )	;
			if ( -1 == iPos )
			{
				AfxMessageBox ( "�������ļ� Config.ini ʱ SQL ���δ���� = " )	;
				delete	lpItem	;
				return	FALSE	;
			}

			strRight	= strLine.Right ( strLine.GetLength() - iPos - 1 );

			strRight.TrimLeft ( )	;
			strRight.TrimRight ( )	;
			strRight	= strRight.Left ( _MAX_SQL );

			::strcpy ( lpSubItem->cSQL1, strRight )	;

			lpSubItem->cSort1[0]	= '0'	;
			lpSubItem->cFilter1[0]	= '0'	;

			continue;
		}

		if ( 0 == strLine.Find ( "sort1" ) )
		{
			// �������
			iPos		= strLine.Find ( "=" )	;
			if ( -1 == iPos )
			{
				AfxMessageBox ( "�������ļ� Config.ini ʱ SQL ���δ���� = " )	;
				delete	lpItem	;
				return	FALSE	;
			}

			strRight	= strLine.Right ( strLine.GetLength() - iPos - 1 );

			strRight.TrimLeft ( )	;
			strRight.TrimRight ( )	;
			strRight	= strRight.Left ( _MAX_SORT );

			::strcpy ( lpSubItem->cSort1, strRight )	;

			continue;
		}

		if ( 0 == strLine.Find ( "filter1" ) )
		{
			// �������
			iPos		= strLine.Find ( "=" )	;
			if ( -1 == iPos )
			{
				AfxMessageBox ( "�������ļ� Config.ini ʱ SQL ���δ���� = " )	;
				delete	lpItem	;
				return	FALSE	;
			}

			strRight	= strLine.Right ( strLine.GetLength() - iPos - 1 );

			strRight.TrimLeft ( )	;
			strRight.TrimRight ( )	;
			strRight	= strRight.Left ( _MAX_SORT );

			::strcpy ( lpSubItem->cFilter1, strRight )	;

			continue;
		}

		if ( 0 == strLine.Find ( "sql" ) )
		{
			// �������
			iPos		= strLine.Find ( "=" )	;
			if ( -1 == iPos )
			{
				AfxMessageBox ( "�������ļ� Config.ini ʱ SQL ���δ���� = " )	;
				delete	lpItem	;
				return	FALSE	;
			}

			strRight	= strLine.Right ( strLine.GetLength() - iPos - 1 );

			strRight.TrimLeft ( )	;
			strRight.TrimRight ( )	;
			strRight	= strRight.Left ( _MAX_SQL );

			::strcpy ( lpSubItem->cSQL, strRight )	;

			lpSubItem->cSort[0]		= '0'	;
			lpSubItem->cFilter[0]	= '0'	;

			continue;
		}

		if ( 0 == strLine.Find ( "sort" ) )
		{
			// �������
			iPos		= strLine.Find ( "=" )	;
			if ( -1 == iPos )
			{
				AfxMessageBox ( "�������ļ� Config.ini ʱ SQL ���δ���� = " )	;
				delete	lpItem	;
				return	FALSE	;
			}

			strRight	= strLine.Right ( strLine.GetLength() - iPos - 1 );

			strRight.TrimLeft ( )	;
			strRight.TrimRight ( )	;
			strRight	= strRight.Left ( _MAX_SORT );

			::strcpy ( lpSubItem->cSort, strRight )	;

			continue;
		}

		if ( 0 == strLine.Find ( "filter" ) )
		{
			// �������
			iPos		= strLine.Find ( "=" )	;
			if ( -1 == iPos )
			{
				AfxMessageBox ( "�������ļ� Config.ini ʱ SQL ���δ���� = " )	;
				delete	lpItem	;
				return	FALSE	;
			}

			strRight	= strLine.Right ( strLine.GetLength() - iPos - 1 );

			strRight.TrimLeft ( )	;
			strRight.TrimRight ( )	;
			strRight	= strRight.Left ( _MAX_SORT );

			::strcpy ( lpSubItem->cFilter, strRight )	;

			continue;
		}

		// ģ��·��
		if ( 0 == strLine.Find ( "doctpl" ) )
		{
			iPos		= strLine.Find ( "=" )	;
			if ( -1 == iPos )
			{
				AfxMessageBox ( "�������ļ� Config.ini ʱ Subject ���δ���� = " )	;
				delete	lpItem	;
				return	FALSE	;
			}

			strRight	= strLine.Right ( strLine.GetLength() - iPos - 1 );

			strRight.TrimLeft ( )	;
			strRight.TrimRight ( )	;
//			strRight	= strRight.Left ( 4 );

			::strcpy ( lpSubItem->cPath, strRight )	;

			continue;
		}
	}

	return	FALSE	;
}

BOOL CPageConfig::ReadFunction ( CStdioFile *pFile, LPSUBITEM lpSubItem, CString strFuncName )
{
	CActFunction *	lpFunc	;

	lpFunc	= new CActFunction	;

	lpFunc->m_listlpAction	= NULL					;
	lpFunc->m_strFuncName	= strFuncName			;
	lpFunc->m_lpNext		= lpSubItem->m_funcList	;
	lpSubItem->m_funcList	= lpFunc				;

	ReadSel ( pFile, &lpFunc->m_pPara, &lpFunc->m_listlpAction, lpSubItem )	;

	return	FALSE	;
}

BOOL CPageConfig::Create(UINT nIDTemplate, CWnd* pParentWnd)
{
	BOOL	bReturn	;

	m_pParentDlg	= (CAnalyserDlg*)pParentWnd	;	// ���б���ŵ���һ��֮ǰ, ��Ȼm_pParentDlgδ��ֵ

	bReturn	= CDialogResize::Create(nIDTemplate, pParentWnd);

	return	bReturn	;
}

void CPageConfig::ResetConfig ( )
{
	LPITEM	lpItem	;

	// ��������ѡ��
	for ( int i=IDC_CHECK1; i<=IDC_CHECK8; i++ )
	{
		GetDlgItem ( i )->ShowWindow ( SW_HIDE );
	}

	i		= IDC_CHECK1					;
	lpItem	= m_pParentDlg->GetItemList ()	;
	while ( NULL != lpItem )
	{
		GetDlgItem ( i )->SetWindowText ( lpItem->cTitle )	;
		GetDlgItem ( i )->ShowWindow ( SW_SHOW )			;
		((CButton*)GetDlgItem ( i ))->SetCheck ( 1 );

		i++;

		lpItem	= lpItem->lpNext;
	}
}

void CPageConfig::OnBnClickedCheck(BYTE yItem)
{
	LPITEM	lpItem	;

	lpItem	= m_pParentDlg->GetItemList ()	;
	while ( NULL != lpItem )
	{
		if ( lpItem->yItemNumber == yItem )
		{
			lpItem->bSelected	= ! lpItem->bSelected;
//			((CButton*)GetDlgItem ( IDC_CHECK1 + yItem ))->SetCheck ( lpItem->bSelected );
			break	;
		}

		lpItem	= lpItem->lpNext;
	}
}

void CPageConfig::OnBnClickedCheck1()
{
	OnBnClickedCheck ( 0 )	;
}

void CPageConfig::OnBnClickedCheck2()
{
	OnBnClickedCheck ( 1 )	;
}

void CPageConfig::OnBnClickedCheck3()
{
	OnBnClickedCheck ( 2 )	;
}

void CPageConfig::OnBnClickedCheck4()
{
	OnBnClickedCheck ( 3 )	;
}

void CPageConfig::OnBnClickedCheck5()
{
	OnBnClickedCheck ( 4 )	;
}

void CPageConfig::OnBnClickedCheck6()
{
	OnBnClickedCheck ( 5 )	;
}

void CPageConfig::OnBnClickedCheck7()
{
	OnBnClickedCheck ( 6 )	;
}

void CPageConfig::OnBnClickedCheck8()
{
	OnBnClickedCheck ( 7 )	;
}

void CPageConfig::OnBnClickedReverse()
{
	m_bReverse	= !m_bReverse	;
}

void CPageConfig::OnBnClickedBtnStepscore()
{
	GetDlgItem ( IDC_BTN_STEPSCORE )->ShowWindow( SW_HIDE )	;

	// ���Ȳ�����صı�
	if ( !GenerateMarks() )
	{
		AfxMessageBox ( "������صı�ʧ�ܣ�\n�������δ�ܲ���" )	;
		GetDlgItem ( IDC_BTN_STEPSCORE )->ShowWindow( SW_SHOW )	;
		return	;
	}

	// �� vScoreCheck ������ SELECT INTO �� ScoreCheck
	CString		strSQL		;
	CDatabase *	pDatabase	;

	pDatabase	= &theApp.m_database;
	try
	{
		// ɾ���ɵı� ScoreCheck
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ScoreCheck]') and OBJECTPROPERTY(id, N'IsTable') = 1)\
					   drop table [dbo].[ScoreCheck]" );
		pDatabase->ExecuteSQL( strSQL );

		strSQL.Format( "SELECT * INTO ScoreCheck FROM vScoreCheck" );
		pDatabase->ExecuteSQL( strSQL );

		// ɾ���ɵ���ͼ vScoreCheckSum
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vScoreCheckSum]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vScoreCheckSum]" );
		pDatabase->ExecuteSQL( strSQL );

		// ������ͼ vScoreCheckSum
		strSQL.Format( "CREATE VIEW dbo.vScoreCheckSum\
					   AS\
					   SELECT MIN(dbo.ScoreCheck.ID) AS ID, dbo.ScoreCheck.StudentID, \
					   QuestionFormInfo_1.QuestionFormID AS FormID, SUM(dbo.ScoreCheck.Score) \
					   AS Score\
					   FROM dbo.ScoreCheck INNER JOIN\
					   dbo.QuestionFormInfo ON \
					   dbo.ScoreCheck.FormID = dbo.QuestionFormInfo.QuestionFormID INNER JOIN\
					   dbo.QuestionFormInfo QuestionFormInfo_1 ON \
					   dbo.QuestionFormInfo.QuestionID = QuestionFormInfo_1.QuestionID AND \
					   dbo.QuestionFormInfo.SubjectID = QuestionFormInfo_1.SubjectID\
					   WHERE (QuestionFormInfo_1.Extend = 1)\
					   GROUP BY dbo.ScoreCheck.StudentID, QuestionFormInfo_1.QuestionFormID" );
		pDatabase->ExecuteSQL( strSQL );
	}
	catch ( CDBException * e )
	{
		CString	strError;
		strError.Format( "SELECT INTO ScoreCheck��ʧ��\r\n" );
		strError	+= e->m_strError;
		AfxMessageBox( strError );
		GetDlgItem ( IDC_BTN_STEPSCORE )->ShowWindow( SW_SHOW )	;
		return;
	}

	CString	strDTS	;

	HRESULT	hr;

	// ���ۺϿ�Ŀ�Ĵ�������
	if SUCCEEDED(hr = OleInitialize(NULL) )
	{
		try
		{
			_Package2Ptr	spPackage;
			if ( SUCCEEDED( spPackage.CreateInstance( __uuidof(Package2) ) ) )
			{
				try
				{
					_variant_t v; // VarPersistStgOfHost

					strDTS.Format( "dts\\MarksFirst.dts" )	;
					hr	= spPackage->LoadFromStorageFile(
						(LPCTSTR)strDTS,
						_T(""),
						_T(""),
						_T(""),
						_T(""),
						&v);

					hr	= spPackage->Execute();

					hr = spPackage->UnInitialize();
				}
				catch(_com_error pCE)
				{
					CString	strError;

					strError.Format( "���� MarksFirst.dts ʧ��\r\n%s Error: %ld\r\n%s\r\n%s\r\n",
									 (TCHAR*)pCE.Source(),
									 pCE.Error(),
									 (TCHAR*)pCE.Description(),
									 (TCHAR*)pCE.ErrorMessage());

					AfxMessageBox ( strError )	;
					spPackage.Release();     // Free the interface
					GetDlgItem ( IDC_BTN_STEPSCORE )->ShowWindow( SW_SHOW )	;
					return	;
				}
			}
		}
		catch(_com_error pCE)
		{
			CString	strError;

			strError.Format( "���� MarksFirst.dts ʧ��\r\n%s Error: %ld\r\n%s\r\n%s\r\n",
							 (TCHAR*)pCE.Source(),
							 pCE.Error(),
							 (TCHAR*)pCE.Description(),
							 (TCHAR*)pCE.ErrorMessage());

			AfxMessageBox ( strError )	;
			GetDlgItem ( IDC_BTN_STEPSCORE )->ShowWindow( SW_SHOW )	;
			return	;
		}

		OleUninitialize();
	}
	else
	{
//		_tprintf(_T("Call to CoInitialize failed.\n"));
		AfxMessageBox ( "���� dts ��ʼ��ʧ��" )	;
		GetDlgItem ( IDC_BTN_STEPSCORE )->ShowWindow( SW_SHOW )	;
		return	;
	}

	// �� vMarksFirstSum ������ SELECT INTO �� MarkSum
	try
	{
		// ɾ���ɵı� MarkSum
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MarkSum]') and OBJECTPROPERTY(id, N'IsTable') = 1)\
					   drop table [dbo].[MarkSum]" );
		pDatabase->ExecuteSQL( strSQL );

		strSQL.Format( "SELECT * INTO MarkSum FROM vMarksFirstSum" );
		pDatabase->ExecuteSQL( strSQL );


		// ɾ���� ScoreCheck
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ScoreCheck]') and OBJECTPROPERTY(id, N'IsTable') = 1)\
					   drop table [dbo].[ScoreCheck]" );
		pDatabase->ExecuteSQL( strSQL );

		// ɾ����ͼ vScoreCheck
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vScoreCheck]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vScoreCheck]" );
		pDatabase->ExecuteSQL( strSQL );

		// ɾ����ͼ vScoreCheckSum
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vScoreCheckSum]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vScoreCheckSum]" );
		pDatabase->ExecuteSQL( strSQL );

		// ɾ����ͼ vScoreFinal
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vScoreFinal]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vScoreFinal]" );
		pDatabase->ExecuteSQL( strSQL );

		// ɾ����ͼ vMarksFirstSum
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vMarksFirstSum]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vMarksFirstSum]" );
		pDatabase->ExecuteSQL( strSQL );

		//// ɾ���ɵ���ͼ 
		//strSQL.Format( "" );
		//pDatabase->ExecuteSQL( strSQL );

		//// ������ͼ 
		//strSQL.Format( "" );
		//pDatabase->ExecuteSQL( strSQL );
	}
	catch ( CDBException * e )
	{
		CString	strError;
		strError.Format( "SELECT INTO ��MarkSumʧ��\r\n" );
		strError	+= e->m_strError;
		AfxMessageBox( strError );
		GetDlgItem ( IDC_BTN_STEPSCORE )->ShowWindow( SW_SHOW )	;

		return;
	}

	GetDlgItem ( IDC_BTN_STEPSCORE )->ShowWindow( SW_SHOW )	;

//	AfxMessageBox ( "�ɹ��������׼����" )	;

//	RedrawWindow ( );

	m_pParentDlg->SetPage ( 1 )	;
}

void CPageConfig::OnBnClickedBtnCheckCnt()
{
	if ( m_pParentDlg->m_pAnalyzer->BuildClassCheckInfo ( m_bReverse ) )
		AfxMessageBox ( "�ɹ����ѡ�����������ͳ������׼����" )	;

	//RedrawWindow ( );

	m_pParentDlg->SetPage ( 1 )	;
}

void CPageConfig::OnBnClickedBtnStepscore3()
{
	// ���Ȳ�����ص���ͼ
	if ( !GenerateAvg() )
	{
		AfxMessageBox ( "������ص���ͼʧ�ܣ�\n�������δ�ܲ���" )	;
		return	;
	}

	// �������ֲ�
	CString	strDTS	;

	HRESULT	hr;

	// ���ۺϿ�Ŀ�Ĵ�������
	if SUCCEEDED(hr = OleInitialize(NULL) )
	{
		try
		{
			_Package2Ptr	spPackage;
			if ( SUCCEEDED( spPackage.CreateInstance( __uuidof(Package2) ) ) )
			{
				try
				{
					_variant_t v; // VarPersistStgOfHost
					//Added by ����֮��
					strDTS.Format( "dts\\AverageSource.dts" )	;
					hr	= spPackage->LoadFromStorageFile(
						(LPCTSTR)strDTS,
						_T(""),
						_T(""),
						_T(""),
						_T(""),
						&v);

					hr	= spPackage->Execute();

					hr = spPackage->UnInitialize();
				}
				catch(_com_error pCE)
				{
//					DisplayError(pCE);
					AfxMessageBox ( "���� AverageSource.dts ʧ��" )	;
					spPackage.Release();     // Free the interface
					return	;
				}
			}
		}
		catch(_com_error pCE)
		{
//			DisplayError(pCE);
			AfxMessageBox ( "���� AverageSource.dts ʧ��" )	;
			return	;
		}

		OleUninitialize();
	}
	else
	{
//		_tprintf(_T("Call to CoInitialize failed.\n"));
		AfxMessageBox ( "���� AverageSource.dts ʧ��" )	;
		return	;
	}

	//AfxMessageBox ( "�ɹ���ɽ��о��ַ���������׼����" )	;

	//RedrawWindow ( );

	m_pParentDlg->SetPage ( 1 )	;
}

void CPageConfig::OnPaint()
{
	CPaintDC dc(this); // device context for painting

	// ��Ϊ��ͼ��Ϣ���� CDialogResize::OnPaint()
	//CDialogResize::OnPaint();

	//m_pParentDlg->RedrawWindow ( );
}

/*+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
�������������µ����ݿ����
	�� MarksFirst
	�� MarkSum
	�� MarkCompositive
	�� MarksForm
	�� MarkTotal

	��ͼ  vScoreFinal
	��ͼ  vScoreCheck
	��ͼ  vMarksFirstSum
	��ͼ  vMainSubjectMarksFirst
---------------------------------------------------------------------------*/
BOOL CPageConfig::GenerateMarks ( )
{
	CString		strSQL		;
	CDatabase *	pDatabase	;

	pDatabase	= &theApp.m_database;
	try
	{
		// ɾ���ɵ� MarksFirst
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MarksFirst]') and OBJECTPROPERTY(id, N'IsUserTable') = 1) \
					   drop table [dbo].[MarksFirst]" );
		pDatabase->ExecuteSQL( strSQL );

		// ���� MarksFirst
		strSQL.Format( "CREATE TABLE [dbo].[MarksFirst] (\
					   [ID] [int] IDENTITY (1, 1) NOT NULL ,\
					   [StudentID] [int] NOT NULL ,\
					   [FormID] [int] NOT NULL ,\
					   [Score] [real] NULL \
					   ) ON [PRIMARY]" );
		pDatabase->ExecuteSQL( strSQL );

		// ɾ���ɵ� MarkSum
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MarkSum]') and OBJECTPROPERTY(id, N'IsUserTable') = 1) \
					   drop table [dbo].[MarkSum]" );
		pDatabase->ExecuteSQL( strSQL );

		// ���� MarkSum
		strSQL.Format( "CREATE TABLE [dbo].[MarkSum] (\
					   [ID] [int] IDENTITY (1, 1) NOT NULL ,\
					   [StudentID] [int] NOT NULL ,\
					   [FormID] [int] NOT NULL ,\
					   [Score] [real] NULL \
					   ) ON [PRIMARY]" );
		pDatabase->ExecuteSQL( strSQL );

		// ɾ���ɵ� MarkTotal
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MarkTotal]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)\
					   drop table [dbo].[MarkTotal]" );
		pDatabase->ExecuteSQL( strSQL );

		// ���� MarkTotal
		strSQL.Format( "CREATE TABLE [dbo].[MarkTotal] (\
					   [ID] [int] IDENTITY (1, 1) NOT NULL ,\
					   [StudentID] [int] NOT NULL ,\
					   [FormID] [int] NOT NULL ,\
					   [Score] [real] NULL ) ON [PRIMARY]" );
		pDatabase->ExecuteSQL( strSQL );

		// ɾ���ɵ���ͼ vScoreCheck
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vScoreCheck]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vScoreCheck]" );
		pDatabase->ExecuteSQL( strSQL );

		// ������ͼ vScoreCheck
		strSQL.Format( "CREATE VIEW dbo.vScoreCheck\
					   AS\
					   SELECT dbo._ScoreStudentStep0.ID, dbo.Sheet.StudentID, \
					   dbo.QuestionFormInfo.QuestionFormID AS FormID, \
					   dbo._ScoreStudentStep0.Score\
					   FROM dbo._ScoreStudentStep0 INNER JOIN\
					   dbo.QuestionFormInfo ON \
					   dbo._ScoreStudentStep0.Step = dbo.QuestionFormInfo.Step INNER JOIN\
					   dbo.CutSheet ON \
					   dbo._ScoreStudentStep0.CutSheetID = dbo.CutSheet.CutSheetID AND \
					   dbo.QuestionFormInfo.QuestionID = dbo.CutSheet.QuestionID INNER JOIN\
					   dbo.Sheet ON dbo.CutSheet.SheetID = dbo.Sheet.SheetID\
					   WHERE (dbo.QuestionFormInfo.Extend = 0)" );
		pDatabase->ExecuteSQL( strSQL );

		// ɾ���ɵ���ͼ vMainSubjectMarksFirst
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vMainSubjectMarksFirst]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vMainSubjectMarksFirst]" );
		pDatabase->ExecuteSQL( strSQL );

		// ������ͼ vMainSubjectMarksFirst
		strSQL.Format( "CREATE VIEW dbo.vMainSubjectMarksFirst\
					   AS\
					   SELECT DISTINCT dbo.Subject.SubjectID, dbo.Subject.SubjectName\
					   FROM dbo.MarksFirst INNER JOIN\
					   dbo.QuestionFormInfo ON \
					   dbo.MarksFirst.FormID = dbo.QuestionFormInfo.QuestionFormID INNER JOIN\
					   dbo.subSubject ON \
					   dbo.QuestionFormInfo.SubjectID = dbo.subSubject.SubjectID INNER JOIN\
					   dbo.Subject ON dbo.subSubject.MainSubjectID = dbo.Subject.SubjectID" );
		pDatabase->ExecuteSQL( strSQL );

		// ɾ���ɵ���ͼ vScoreFinal
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vScoreFinal]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vScoreFinal]" );
		pDatabase->ExecuteSQL( strSQL );

		// ������ͼ vScoreFinal
		strSQL.Format( "CREATE VIEW dbo.vScoreFinal\
					   AS\
					   SELECT dbo.ScoreFinal.ID, dbo.Sheet.StudentID, dbo.ScoreFinal.FormID, \
					   dbo.ScoreFinal.Score\
					   FROM dbo.ScoreFinal INNER JOIN\
					   dbo.CutSheet ON dbo.ScoreFinal.CutSheetID = dbo.CutSheet.CutSheetID INNER JOIN\
					   dbo.Sheet ON dbo.CutSheet.SheetID = dbo.Sheet.SheetID" );
		pDatabase->ExecuteSQL( strSQL );

		// ɾ���ɵ���ͼ vMarksFirstSum
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vMarksFirstSum]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vMarksFirstSum]" );
		pDatabase->ExecuteSQL( strSQL );

		// ������ͼ vMarksFirstSum
		strSQL.Format( "CREATE VIEW dbo.vMarksFirstSum\
					   AS\
					   SELECT dbo.MarksFirst.StudentID AS ID, dbo.MarksFirst.StudentID, \
					   QuestionFormInfo_1.QuestionFormID AS FormID, SUM(dbo.MarksFirst.Score) \
					   AS Score\
					   FROM dbo.MarksFirst INNER JOIN\
					   dbo.QuestionFormInfo ON \
					   dbo.MarksFirst.FormID = dbo.QuestionFormInfo.QuestionFormID INNER JOIN\
					   dbo.QuestionFormInfo QuestionFormInfo_1 ON \
					   dbo.QuestionFormInfo.SubjectID = QuestionFormInfo_1.SubjectID\
					   WHERE (QuestionFormInfo_1.Extend = 2)\
					   GROUP BY dbo.MarksFirst.StudentID, QuestionFormInfo_1.QuestionFormID" );
		pDatabase->ExecuteSQL( strSQL );

		//// ɾ���ɵ���ͼ 
		//strSQL.Format( "" );
		//pDatabase->ExecuteSQL( strSQL );

		//// ������ͼ 
		//strSQL.Format( "" );
		//pDatabase->ExecuteSQL( strSQL );
	}
	catch ( CDBException * e )
	{
		CString	strError;
		strError.Format( "���ɴ���ɼ���صı�ʧ��\r\n" );
		strError	+= e->m_strError;
		AfxMessageBox( strError );
		return	FALSE;
	}
	
	return	TRUE;
}

BOOL CPageConfig::GenerateAvg ( )
{
	CString		strSQL		;
	CDatabase *	pDatabase	;

	pDatabase	= &theApp.m_database;
	try
	{
		// ɾ���ɵ� MarksFirst
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MarksFirst]') and OBJECTPROPERTY(id, N'IsUserTable') = 1) \
					   drop table [dbo].[MarksFirst]" );
		pDatabase->ExecuteSQL( strSQL );

		// ���� MarksFirst
		strSQL.Format( "CREATE TABLE [dbo].[MarksFirst] (\
					   [ID] [int] IDENTITY (1, 1) NOT NULL ,\
					   [StudentID] [int] NOT NULL ,\
					   [FormID] [int] NOT NULL ,\
					   [Score] [real] NULL \
					   ) ON [PRIMARY]" );
		pDatabase->ExecuteSQL( strSQL );

		// ɾ���ɵ���ͼ vMarksFirst0
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vMarksFirst0]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vMarksFirst0]" );
		pDatabase->ExecuteSQL( strSQL );

		// ������ͼ vMarksFirst0
		strSQL.Format( "CREATE VIEW dbo.vMarksFirst0\
					   AS\
					   SELECT DISTINCT StudentID\
					   FROM dbo.MarksFirst\
					   WHERE (Score > 0)" );
		pDatabase->ExecuteSQL( strSQL );

		// ɾ���ɵ���ͼ vAVGMarkClass
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vAVGMarkClass]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vAVGMarkClass]" );
		pDatabase->ExecuteSQL( strSQL );

		// ������ͼ vAVGMarkClass
		strSQL.Format( "CREATE VIEW dbo.vAVGMarkClass\
					   AS\
					   SELECT dbo.Student.ClassID, dbo.MarksFirst.FormID, AVG(dbo.MarksFirst.Score) \
					   AS avgScore, dbo.Student.Type\
					   FROM dbo.MarksFirst INNER JOIN\
					   dbo.Student ON dbo.MarksFirst.StudentID = dbo.Student.StudentID INNER JOIN\
					   dbo.vMarksFirst0 ON dbo.Student.StudentID = dbo.vMarksFirst0.StudentID\
					   GROUP BY dbo.Student.ClassID, dbo.MarksFirst.FormID, dbo.Student.Type" );
		pDatabase->ExecuteSQL( strSQL );

		// ɾ���ɵ���ͼ vAVGClass
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vAVGClass]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vAVGClass]" );
		pDatabase->ExecuteSQL( strSQL );

		// ������ͼ vAVGClass
		strSQL.Format( "CREATE VIEW dbo.vAVGClass\
					   AS\
					   SELECT dbo.vAVGMarkClass.ClassID, dbo.vAVGMarkClass.FormID, \
					   dbo.vAVGMarkClass.avgScore, dbo.QuestionFormInfo.SubjectID, dbo.Class.SchoolID, \
					   dbo.Class.ClassName, dbo.vAVGMarkClass.Type\
					   FROM dbo.vAVGMarkClass INNER JOIN\
					   dbo.Class ON dbo.vAVGMarkClass.ClassID = dbo.Class.ClassID INNER JOIN\
					   dbo.QuestionFormInfo ON \
					   dbo.vAVGMarkClass.FormID = dbo.QuestionFormInfo.QuestionFormID" );
		pDatabase->ExecuteSQL( strSQL );

		// ɾ���ɵ���ͼ vAVGMarkSchool
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vAVGMarkSchool]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vAVGMarkSchool]" );
		pDatabase->ExecuteSQL( strSQL );

		// ������ͼ vAVGMarkSchool
		strSQL.Format( "CREATE VIEW dbo.vAVGMarkSchool\
					   AS\
					   SELECT dbo.Class.SchoolID, dbo.MarksFirst.FormID, AVG(dbo.MarksFirst.Score) \
					   AS avgScore, dbo.Student.Type\
					   FROM dbo.MarksFirst INNER JOIN\
					   dbo.Student ON dbo.MarksFirst.StudentID = dbo.Student.StudentID INNER JOIN\
					   dbo.Class ON dbo.Student.ClassID = dbo.Class.ClassID JOIN\
					   dbo.vMarksFirst0 ON dbo.Student.StudentID = dbo.vMarksFirst0.StudentID\
					   GROUP BY dbo.MarksFirst.FormID, dbo.Class.SchoolID, dbo.Student.Type" );
		pDatabase->ExecuteSQL( strSQL );

		// ɾ���ɵ���ͼ vAVGSchool
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vAVGSchool]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vAVGSchool]" );
		pDatabase->ExecuteSQL( strSQL );

		// ������ͼ vAVGSchool
		strSQL.Format( "CREATE VIEW dbo.vAVGSchool\
					   AS\
					   SELECT dbo.vAVGMarkSchool.SchoolID, dbo.vAVGMarkSchool.FormID, \
					   dbo.vAVGMarkSchool.avgScore, dbo.QuestionFormInfo.SubjectID, \
					   dbo.School.RegionID, dbo.School.SchoolName, dbo.vAVGMarkSchool.Type\
					   FROM dbo.vAVGMarkSchool INNER JOIN\
					   dbo.School ON dbo.vAVGMarkSchool.SchoolID = dbo.School.SchoolID INNER JOIN\
					   dbo.QuestionFormInfo ON \
					   dbo.vAVGMarkSchool.FormID = dbo.QuestionFormInfo.QuestionFormID" );
		pDatabase->ExecuteSQL( strSQL );

		// ɾ���ɵ���ͼ vAVGMarkRegion
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vAVGMarkRegion]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vAVGMarkRegion]" );
		pDatabase->ExecuteSQL( strSQL );

		// ������ͼ vAVGMarkRegion
		strSQL.Format( "CREATE VIEW dbo.vAVGMarkRegion\
					   AS\
					   SELECT dbo.School.RegionID, dbo.MarksFirst.FormID, AVG(dbo.MarksFirst.Score) \
					   AS avgScore, dbo.Student.Type\
					   FROM dbo.MarksFirst INNER JOIN\
					   dbo.Student ON dbo.MarksFirst.StudentID = dbo.Student.StudentID INNER JOIN\
					   dbo.Class ON dbo.Student.ClassID = dbo.Class.ClassID INNER JOIN\
					   dbo.School ON dbo.Class.SchoolID = dbo.School.SchoolID JOIN\
					   dbo.vMarksFirst0 ON dbo.Student.StudentID = dbo.vMarksFirst0.StudentID\
					   GROUP BY dbo.MarksFirst.FormID, dbo.School.RegionID, dbo.Student.Type" );
		pDatabase->ExecuteSQL( strSQL );

		// ɾ���ɵ���ͼ vAVGRegion
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vAVGRegion]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vAVGRegion]" );
		pDatabase->ExecuteSQL( strSQL );

		// ������ͼ vAVGRegion
		strSQL.Format( "CREATE VIEW dbo.vAVGRegion\
					   AS\
					   SELECT dbo.vAVGMarkRegion.RegionID, dbo.vAVGMarkRegion.FormID, \
					   dbo.vAVGMarkRegion.avgScore, dbo.QuestionFormInfo.SubjectID, - 1 AS CityID, \
					   '��' AS RegionName, dbo.vAVGMarkRegion.Type\
					   FROM dbo.vAVGMarkRegion INNER JOIN\
					   dbo.QuestionFormInfo ON \
					   dbo.vAVGMarkRegion.FormID = dbo.QuestionFormInfo.QuestionFormID" );
		pDatabase->ExecuteSQL( strSQL );

		// ɾ���ɵ���ͼ vScoreCheck
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vScoreCheck]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vScoreCheck]" );
		pDatabase->ExecuteSQL( strSQL );

		// ������ͼ vScoreCheck
		strSQL.Format( "CREATE VIEW dbo.vScoreCheck\
					   AS\
					   SELECT dbo._ScoreStudentStep0.ID, dbo.Sheet.StudentID, \
					   dbo.QuestionFormInfo.QuestionFormID AS FormID, \
					   dbo._ScoreStudentStep0.Score\
					   FROM dbo._ScoreStudentStep0 INNER JOIN\
					   dbo.QuestionFormInfo ON \
					   dbo._ScoreStudentStep0.Step = dbo.QuestionFormInfo.Step INNER JOIN\
					   dbo.CutSheet ON \
					   dbo._ScoreStudentStep0.CutSheetID = dbo.CutSheet.CutSheetID AND \
					   dbo.QuestionFormInfo.QuestionID = dbo.CutSheet.QuestionID INNER JOIN\
					   dbo.Sheet ON dbo.CutSheet.SheetID = dbo.Sheet.SheetID\
					   WHERE (dbo.QuestionFormInfo.Extend = 0)" );
		pDatabase->ExecuteSQL( strSQL );

		// ɾ���ɵ���ͼ vScoreFinal
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vScoreFinal]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vScoreFinal]" );
		pDatabase->ExecuteSQL( strSQL );

		// ������ͼ vScoreFinal
		strSQL.Format( "CREATE VIEW dbo.vScoreFinal\
					   AS\
					   SELECT dbo.ScoreFinal.ID, dbo.Sheet.StudentID, dbo.ScoreFinal.FormID, \
					   dbo.ScoreFinal.Score\
					   FROM dbo.ScoreFinal INNER JOIN\
					   dbo.CutSheet ON dbo.ScoreFinal.CutSheetID = dbo.CutSheet.CutSheetID INNER JOIN\
					   dbo.Sheet ON dbo.CutSheet.SheetID = dbo.Sheet.SheetID" );
		pDatabase->ExecuteSQL( strSQL );

		//// ɾ���ɵ���ͼ 
		//strSQL.Format( "" );
		//pDatabase->ExecuteSQL( strSQL );

		//// ������ͼ 
		//strSQL.Format( "" );
		//pDatabase->ExecuteSQL( strSQL );
	}
	catch ( CDBException * e )
	{
		CString	strError;
		strError.Format( "������ص����ݿ����ʧ��\r\n" );
		strError	+= e->m_strError;
		AfxMessageBox( strError );
		return	FALSE;
	}
	
	return	TRUE;
}
