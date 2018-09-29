// PageConfig.cpp : 实现文件
//

#include "stdafx.h"
#include "Analyser.h"
#include "PageConfig.h"
#include "AnalyserDlg.h"
#include "Bind.h"
#include "Analyzer.h"
//#include "vCnt.h"

// CPageConfig 对话框

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
// 绑定SQL数据源到列表框或组合框
// 参数 wnd:     列表控件或者组合框控件对象，要求支持ResetContent()和AddString()函数
//      sql:     绑定的sql语句(注：该函数不进行sql语句有效性校验，请自行确保其正确性)
//               sql语句必须类似"select distinct [field] from tablename where ..."的形式
//      attach:  附加功能开关(添加首行'-')
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

	//选择第一项
	if( wnd.GetCount() > 0 ) 
		wnd.SetCurSel( 0 )	;

	return true;
}



// CPageConfig 消息处理程序

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
			ResetConfig ( )	;	// 若读失败则不做
	}

	strPath	= theApp.GetProfileString("Config", "m_strPathResults");
	m_edPathResults.SetWindowText( strPath )	;
	if ( 0 != strPath.GetLength() )
		m_pParentDlg->SetPathOutput ( strPath )	;

	OnCbnSelchangeComboRegion();

	return TRUE;  // return TRUE unless you set the focus to a control
	// 异常: OCX 属性页应返回 FALSE
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

   // 设置配置文件的路径
	// Create an Open dialog; the default file name extension is ".my".
	CFileDialog fileDlg (TRUE, "请选择初始化文件", "*.ini",
		OFN_FILEMUSTEXIST| OFN_HIDEREADONLY, szFilters, this);

	// Display the file dialog. When user clicks OK, fileDlg.DoModal() 
	// returns IDOK.
	if( fileDlg.DoModal ()==IDOK )
	{
		if ( IDYES != MessageBox("确认导入配置文件", "请确认", MB_YESNOCANCEL) )
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
	bi.lpszTitle		= "请选择储存分析结果的文件夹！";
	bi.ulFlags			= BIF_USENEWUI;

	LPITEMIDLIST idl	= SHBrowseForFolder(&bi);
	if( idl == NULL )
		return;
	SHGetPathFromIDList(idl,str.GetBuffer(MAX_PATH));
	str.ReleaseBuffer();

	if( str.GetLength()<=0 )
	{
		AfxMessageBox( "路径无效！\n  失败！" )	;
		return;
	}

	m_edPathResults.SetWindowText( str )	;

	m_pParentDlg->SetPathOutput ( str )	;

	theApp.WriteProfileString("Config", "m_strPathResults", str);
}

BOOL CPageConfig::ReadConfig ( CString strPath )
{
	BYTE		yItemNumber	;	// 项目编号
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
		AfxMessageBox ( strPath + " 打开失败\n" )	;
#endif

		return	FALSE;
	}

	// 从文件中逐行读出内容进行解析
	yItemNumber	= 0	;
	while ( file.ReadString ( strLine ) )
	{
		strLine.TrimLeft(); strLine.TrimRight();
		strLine.MakeLower ( );

		// 过滤掉空行
		if ( 0 == strLine.GetLength () )
			continue;

		// 过滤掉注释行
		if ( 0 == strLine.Find ( "//" ) )
			continue;

		// 找到一个块
		if ( 0 == strLine.Find ( "begin" ) )
		{
			// 块的类型
			strRight	= strLine.Right ( strLine.GetLength() - 5 )	;
			iPos		= strRight.Find ( "{" )	;
			if ( -1 == iPos )
			{
				AfxMessageBox ( "读配置文件 Config.ini 时块前置标识 { 失败" )	;
				file.Close ()	;
				return	FALSE	;
			}

			strRight	= strRight.Right ( strRight.GetLength() - iPos - 1 );

			iPos		= strRight.Find ( "}" )	;
			if ( -1 == iPos )
			{
				AfxMessageBox ( "读配置文件 Config.ini 时块前置标识 } 失败" )	;
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

		// 过滤掉空行
		if ( 0 == strLine.GetLength () )
			continue;

		// 过滤掉注释行
		if ( 0 == strLine.Find ( "//" ) )
			continue;

		// 找到一个项
		if ( 0 == strLine.Find ( "end" ) )
		{
			m_pParentDlg->AddItem ( lpItem )	;

			return	TRUE;
		}

		if ( 0 == strLine.Find ( "formclass" ) )
		{
			// 块的类型
			iPos		= strLine.Find ( "=" )	;
			if ( -1 == iPos )
			{
				AfxMessageBox ( "读配置文件 Config.ini 时TITLE 项后未发现 = " )	;
				delete	lpItem	;
				return	FALSE	;
			}

			strRight	= strLine.Right ( strLine.GetLength() - iPos - 1 );

			strRight.TrimLeft ( )	;

			if ( 0 == strRight.Compare ("1") )
			{
				// 单科分析
				lpItem->yFormClass	= 1	;
			}
			else if ( 0 == strRight.Compare ("2") )
			{
				// 多科分析
				lpItem->yFormClass	= 2	;
			}
			else	// 缺省
				lpItem->yFormClass	= 1	;

			continue;
		}

		// 分析的层级(区、校、班)
		if ( 0 == strLine.Find ( "layers" ) )
		{
			// 块的类型
			iPos		= strLine.Find ( "=" )	;
			if ( -1 == iPos )
			{
				AfxMessageBox ( "读配置文件 Config.ini 时 LAYERS 项后未发现 = " )	;
				delete	lpItem	;
				return	FALSE	;
			}

			strRight	= strLine.Right ( strLine.GetLength() - iPos - 1 );

			strRight.TrimLeft ( )	;
			strRight.TrimRight ( )	;

			// 区:=0  区校:=1  校:=2  校班:=3  班:=4  区校班:=5
			if ( 0 == strRight.Compare ("区") )
			{
				lpItem->yLayers	= 0	;
			}
			else if ( 0 == strRight.Compare ("区校") )
			{
				lpItem->yLayers	= 1	;
			}
			else if ( 0 == strRight.Compare ("校") )
			{
				lpItem->yLayers	= 2	;
			}
			else if ( 0 == strRight.Compare ("校班") )
			{
				lpItem->yLayers	= 3	;
			}
			else if ( 0 == strRight.Compare ("班") )
			{
				lpItem->yLayers	= 4	;
			}
			else if ( 0 == strRight.Compare ("区校班") )
			{
				lpItem->yLayers	= 5	;
			}
			else	// 缺省
				lpItem->yLayers	= 0	;

			continue;
		}

		// 找到一个子块
		if ( 0 == strLine.Find ( "begin" ) )
		{
			// 处理子块
			strRight	= strLine.Right ( strLine.GetLength() - 5 )	;
			iPos		= strRight.Find ( "{" )	;
			if ( -1 == iPos )
			{
				AfxMessageBox ( "读配置文件 Config.ini 时子块后置标识 { 失败" )	;
				return	FALSE	;
			}

			strRight	= strRight.Right ( strRight.GetLength() - iPos - 1 );

			iPos		= strRight.Find ( "}" )	;
			if ( -1 == iPos )
			{
				AfxMessageBox ( "读配置文件 Config.ini 时子块后置标识 } 失败" )	;
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
			// 块的类型
			iPos		= strLine.Find ( "=" )	;
			if ( -1 == iPos )
			{
				AfxMessageBox ( "读配置文件 Config.ini 时TITLE 项后未发现 = " )	;
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

		// 过滤掉空行
		if ( 0 == strLine.GetLength () )
			continue;

		// 过滤掉注释行
		if ( 0 == strLine.Find ( "//" ) )
			continue;

		// 过滤掉块起点标示行
		if ( 0 == strLine.Find ( "{" ) )
			continue;

		// 找到结束标志 }
		if ( 0 == strLine.Find ( "}" ) )
		{
			// end sub

			return	TRUE;
		}

		// 找到新的子项
		if ( 0 == strLine.Find ( "[" ) )
		{
			if ( (SUB_FIRST == iState) || (SUB == iState) )
			{
				AfxMessageBox ( "连续出现2次[符号" )	;
				return	FALSE;
			}
		
			iState	= SUB_FIRST;
			continue;
		}

		// 子项结束标志
		if ( 0 == strLine.Find ( "]" ) )
		{
			if ( (SUB_FIRST != iState) && (SUB != iState) )
			{
				AfxMessageBox ( "]符号过多" )	;
				return	FALSE;
			}
		
			iState	= MAIN	;
			continue;
		}

		// 找到 HomeKey 操作
		if ( 0 == strLine.Find ( "homekey" ) )
		{
			lpAction	= new CWordHomeKey(lppPara);

			iPos		= 8	;	// 过滤掉　HomeKey
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

		// 找到 HomeKey 操作
		if ( 0 == strLine.Find ( "formatprint" ) )
		{
			lpAction	= new CWordFormatPrint(lppPara);

			iPos		= 12	;	// 过滤掉　FormatPrint
			strRight	= strLine.Right ( strLine.GetLength() - iPos )	;

			strRight.TrimLeft ();

			if ( 0 == strRight.Find ( "mark" ) )
			{
				iPos	= 5		;	// 过滤掉 Mark
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

		// 找到 TypeText 操作
		if ( 0 == strLine.Find ( "typetext" ) )
		{
			lpAction	= new CWordTypeText(lppPara);

			iPos		= 9	;	// 过滤掉 TypeText
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

		// 找到 MoveDown 操作
		if ( 0 == strLine.Find ( "movedown" ) )
		{
			lpAction	= new CWordMoveDown(lppPara);

			iPos		= 9	;	// 过滤掉 MoveDown
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

		// 找到 MoveRight 操作
		if ( 0 == strLine.Find ( "moveright" ) )
		{
			lpAction	= new CWordMoveRight(lppPara);

			iPos		= 10	;	// 过滤掉 MoveRight
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

		// 找到操作函数
		if ( 0 == strLine.Find ( "function" ) )
		{
			lpAction	= new CWordActFunction(lppPara);

			iPos		= 9	;
			strRight	= strLine.Right ( strLine.GetLength() - iPos )	;
			strRight.TrimLeft();	strRight.TrimRight();

			// 找到该函数
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
				AfxMessageBox ( strRight + "\n无法在脚本中找到该函数" )	;
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

		// 过滤掉空行
		if ( 0 == strLine.GetLength () )
			continue;

		// 过滤掉注释行
		if ( 0 == strLine.Find ( "//" ) )
			continue;

		// 过滤掉块起点标示行
		if ( 0 == strLine.Find ( "{" ) )
			continue;

		// 找到结束标志 }
		if ( 0 == strLine.Find ( "}" ) )
		{
			// end sub

			return	TRUE;
		}

		// 每个block中的列数
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
				if ( 8 == iCnt )	// 最多只能放8个到 m_iMaxColwOfBlock[8] 中
					break	;

				strLeft	= strLine.Left ( iPos )								;
				strLine	= strLine.Right ( strLine.GetLength() - iPos - 1 )	;

				*lpiMax++	= atoi ( strLeft )	;

				iPos	= strLine.Find ( ',' )	;
			}

			continue;
		}

		// 每行中相应的子行数
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
				if ( 256 == iCnt )	// 最多只能放256个到 m_iMaxRowOfBlock[256] 中
					break	;

				strLeft	= strLine.Left ( iPos )								;
				strLine	= strLine.Right ( strLine.GetLength() - iPos - 1 )	;

				*lpiMax++	= atoi ( strLeft )	;

				iPos	= strLine.Find ( ',' )	;
			}

			continue;
		}

		// block的个数
		if ( 0 == strLine.Find ( ";" ) )
		{
			iPos	= strLine.Find ( ';' )	;
			strLine	= strLine.Right ( strLine.GetLength() - iPos - 1 )	;
			iPos	= strLine.Find ( ',' )	;

			strLeft	= strLine.Left ( iPos )								;
			lpSubItem->m_iCntBlocks	= atoi ( strLeft )	;


			continue;
		}

		// 每block中的行数
		if ( 0 != strLine.GetLength () )
		{
			int	iCnt	= 0	;
			int	*lpiMax		;

			lpiMax	= lpSubItem->m_iMaxRowOfBlock	;
			iPos	= strLine.Find ( ',' )			;

			while ( -1 != iPos )
			{
				iCnt++	;
				if ( 8 == iCnt )	// 最多8个到 m_iMaxRowOfBlock[8] 中
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

		// 过滤掉空行
		if ( 0 == strLine.GetLength () )
			continue;

		// 过滤掉注释行
		if ( 0 == strLine.Find ( "//" ) )
			continue;

		// 找到一个项
		if ( 0 == strLine.Find ( "end" ) )
		{
			// end sub

			return	TRUE;
		}

		if ( 0 == strLine.Find ( "subject" ) )
		{
			// 块的类型
			iPos		= strLine.Find ( "=" )	;
			if ( -1 == iPos )
			{
				AfxMessageBox ( "读配置文件 Config.ini 时 Subject 项后未发现 = " )	;
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
			// 块的类型
			iPos		= strLine.Find ( "=" )	;
			if ( -1 == iPos )
			{
				AfxMessageBox ( "读配置文件 Config.ini 时 DOCTITLE 项后未发现 = " )	;
				delete	lpItem	;
				return	FALSE	;
			}

			strRight	= strLine.Right ( strLine.GetLength() - iPos - 1 );

			strRight.TrimLeft ( )	;

			::strcpy ( lpSubItem->cTitle, strRight )	;

			continue;
		}

		// 该 Sub 所使用的 function
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
			// 块的类型
			iPos		= strLine.Find ( "=" )	;
			if ( -1 == iPos )
			{
				AfxMessageBox ( "读配置文件 Config.ini 时 SQL 项后未发现 = " )	;
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
			// 块的类型
			iPos		= strLine.Find ( "=" )	;
			if ( -1 == iPos )
			{
				AfxMessageBox ( "读配置文件 Config.ini 时 SQL 项后未发现 = " )	;
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
			// 块的类型
			iPos		= strLine.Find ( "=" )	;
			if ( -1 == iPos )
			{
				AfxMessageBox ( "读配置文件 Config.ini 时 SQL 项后未发现 = " )	;
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
			// 块的类型
			iPos		= strLine.Find ( "=" )	;
			if ( -1 == iPos )
			{
				AfxMessageBox ( "读配置文件 Config.ini 时 SQL 项后未发现 = " )	;
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
			// 块的类型
			iPos		= strLine.Find ( "=" )	;
			if ( -1 == iPos )
			{
				AfxMessageBox ( "读配置文件 Config.ini 时 SQL 项后未发现 = " )	;
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
			// 块的类型
			iPos		= strLine.Find ( "=" )	;
			if ( -1 == iPos )
			{
				AfxMessageBox ( "读配置文件 Config.ini 时 SQL 项后未发现 = " )	;
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
			// 块的类型
			iPos		= strLine.Find ( "=" )	;
			if ( -1 == iPos )
			{
				AfxMessageBox ( "读配置文件 Config.ini 时 SQL 项后未发现 = " )	;
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
			// 块的类型
			iPos		= strLine.Find ( "=" )	;
			if ( -1 == iPos )
			{
				AfxMessageBox ( "读配置文件 Config.ini 时 SQL 项后未发现 = " )	;
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
			// 块的类型
			iPos		= strLine.Find ( "=" )	;
			if ( -1 == iPos )
			{
				AfxMessageBox ( "读配置文件 Config.ini 时 SQL 项后未发现 = " )	;
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

		// 模板路径
		if ( 0 == strLine.Find ( "doctpl" ) )
		{
			iPos		= strLine.Find ( "=" )	;
			if ( -1 == iPos )
			{
				AfxMessageBox ( "读配置文件 Config.ini 时 Subject 项后未发现 = " )	;
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

	m_pParentDlg	= (CAnalyserDlg*)pParentWnd	;	// 此行必须放到下一行之前, 不然m_pParentDlg未赋值

	bReturn	= CDialogResize::Create(nIDTemplate, pParentWnd);

	return	bReturn	;
}

void CPageConfig::ResetConfig ( )
{
	LPITEM	lpItem	;

	// 隐藏所有选项
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

	// 首先产生相关的表
	if ( !GenerateMarks() )
	{
		AfxMessageBox ( "产生相关的表失败！\n分析结果未能产生" )	;
		GetDlgItem ( IDC_BTN_STEPSCORE )->ShowWindow( SW_SHOW )	;
		return	;
	}

	// 将 vScoreCheck 的数据 SELECT INTO 表 ScoreCheck
	CString		strSQL		;
	CDatabase *	pDatabase	;

	pDatabase	= &theApp.m_database;
	try
	{
		// 删除旧的表 ScoreCheck
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ScoreCheck]') and OBJECTPROPERTY(id, N'IsTable') = 1)\
					   drop table [dbo].[ScoreCheck]" );
		pDatabase->ExecuteSQL( strSQL );

		strSQL.Format( "SELECT * INTO ScoreCheck FROM vScoreCheck" );
		pDatabase->ExecuteSQL( strSQL );

		// 删除旧的视图 vScoreCheckSum
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vScoreCheckSum]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vScoreCheckSum]" );
		pDatabase->ExecuteSQL( strSQL );

		// 生成视图 vScoreCheckSum
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
		strError.Format( "SELECT INTO ScoreCheck表失败\r\n" );
		strError	+= e->m_strError;
		AfxMessageBox( strError );
		GetDlgItem ( IDC_BTN_STEPSCORE )->ShowWindow( SW_SHOW )	;
		return;
	}

	CString	strDTS	;

	HRESULT	hr;

	// 非综合科目的处理流程
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

					strError.Format( "调用 MarksFirst.dts 失败\r\n%s Error: %ld\r\n%s\r\n%s\r\n",
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

			strError.Format( "调用 MarksFirst.dts 失败\r\n%s Error: %ld\r\n%s\r\n%s\r\n",
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
		AfxMessageBox ( "调用 dts 初始化失败" )	;
		GetDlgItem ( IDC_BTN_STEPSCORE )->ShowWindow( SW_SHOW )	;
		return	;
	}

	// 将 vMarksFirstSum 的数据 SELECT INTO 表 MarkSum
	try
	{
		// 删除旧的表 MarkSum
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MarkSum]') and OBJECTPROPERTY(id, N'IsTable') = 1)\
					   drop table [dbo].[MarkSum]" );
		pDatabase->ExecuteSQL( strSQL );

		strSQL.Format( "SELECT * INTO MarkSum FROM vMarksFirstSum" );
		pDatabase->ExecuteSQL( strSQL );


		// 删除表 ScoreCheck
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ScoreCheck]') and OBJECTPROPERTY(id, N'IsTable') = 1)\
					   drop table [dbo].[ScoreCheck]" );
		pDatabase->ExecuteSQL( strSQL );

		// 删除视图 vScoreCheck
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vScoreCheck]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vScoreCheck]" );
		pDatabase->ExecuteSQL( strSQL );

		// 删除视图 vScoreCheckSum
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vScoreCheckSum]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vScoreCheckSum]" );
		pDatabase->ExecuteSQL( strSQL );

		// 删除视图 vScoreFinal
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vScoreFinal]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vScoreFinal]" );
		pDatabase->ExecuteSQL( strSQL );

		// 删除视图 vMarksFirstSum
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vMarksFirstSum]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vMarksFirstSum]" );
		pDatabase->ExecuteSQL( strSQL );

		//// 删除旧的视图 
		//strSQL.Format( "" );
		//pDatabase->ExecuteSQL( strSQL );

		//// 生成视图 
		//strSQL.Format( "" );
		//pDatabase->ExecuteSQL( strSQL );
	}
	catch ( CDBException * e )
	{
		CString	strError;
		strError.Format( "SELECT INTO 表MarkSum失败\r\n" );
		strError	+= e->m_strError;
		AfxMessageBox( strError );
		GetDlgItem ( IDC_BTN_STEPSCORE )->ShowWindow( SW_SHOW )	;

		return;
	}

	GetDlgItem ( IDC_BTN_STEPSCORE )->ShowWindow( SW_SHOW )	;

//	AfxMessageBox ( "成功完成数据准备！" )	;

//	RedrawWindow ( );

	m_pParentDlg->SetPage ( 1 )	;
}

void CPageConfig::OnBnClickedBtnCheckCnt()
{
	if ( m_pParentDlg->m_pAnalyzer->BuildClassCheckInfo ( m_bReverse ) )
		AfxMessageBox ( "成功完成选择题作答情况统计数据准备！" )	;

	//RedrawWindow ( );

	m_pParentDlg->SetPage ( 1 )	;
}

void CPageConfig::OnBnClickedBtnStepscore3()
{
	// 首先产生相关的视图
	if ( !GenerateAvg() )
	{
		AfxMessageBox ( "产生相关的视图失败！\n分析结果未能产生" )	;
		return	;
	}

	// 分析均分差
	CString	strDTS	;

	HRESULT	hr;

	// 非综合科目的处理流程
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
					//Added by 王者之剑
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
					AfxMessageBox ( "调用 AverageSource.dts 失败" )	;
					spPackage.Release();     // Free the interface
					return	;
				}
			}
		}
		catch(_com_error pCE)
		{
//			DisplayError(pCE);
			AfxMessageBox ( "调用 AverageSource.dts 失败" )	;
			return	;
		}

		OleUninitialize();
	}
	else
	{
//		_tprintf(_T("Call to CoInitialize failed.\n"));
		AfxMessageBox ( "调用 AverageSource.dts 失败" )	;
		return	;
	}

	//AfxMessageBox ( "成功完成进行均分分析的数据准备！" )	;

	//RedrawWindow ( );

	m_pParentDlg->SetPage ( 1 )	;
}

void CPageConfig::OnPaint()
{
	CPaintDC dc(this); // device context for painting

	// 不为绘图消息调用 CDialogResize::OnPaint()
	//CDialogResize::OnPaint();

	//m_pParentDlg->RedrawWindow ( );
}

/*+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
本方法产生如下的数据库对象
	表 MarksFirst
	表 MarkSum
	表 MarkCompositive
	表 MarksForm
	表 MarkTotal

	视图  vScoreFinal
	视图  vScoreCheck
	视图  vMarksFirstSum
	视图  vMainSubjectMarksFirst
---------------------------------------------------------------------------*/
BOOL CPageConfig::GenerateMarks ( )
{
	CString		strSQL		;
	CDatabase *	pDatabase	;

	pDatabase	= &theApp.m_database;
	try
	{
		// 删除旧的 MarksFirst
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MarksFirst]') and OBJECTPROPERTY(id, N'IsUserTable') = 1) \
					   drop table [dbo].[MarksFirst]" );
		pDatabase->ExecuteSQL( strSQL );

		// 生成 MarksFirst
		strSQL.Format( "CREATE TABLE [dbo].[MarksFirst] (\
					   [ID] [int] IDENTITY (1, 1) NOT NULL ,\
					   [StudentID] [int] NOT NULL ,\
					   [FormID] [int] NOT NULL ,\
					   [Score] [real] NULL \
					   ) ON [PRIMARY]" );
		pDatabase->ExecuteSQL( strSQL );

		// 删除旧的 MarkSum
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MarkSum]') and OBJECTPROPERTY(id, N'IsUserTable') = 1) \
					   drop table [dbo].[MarkSum]" );
		pDatabase->ExecuteSQL( strSQL );

		// 生成 MarkSum
		strSQL.Format( "CREATE TABLE [dbo].[MarkSum] (\
					   [ID] [int] IDENTITY (1, 1) NOT NULL ,\
					   [StudentID] [int] NOT NULL ,\
					   [FormID] [int] NOT NULL ,\
					   [Score] [real] NULL \
					   ) ON [PRIMARY]" );
		pDatabase->ExecuteSQL( strSQL );

		// 删除旧的 MarkTotal
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MarkTotal]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)\
					   drop table [dbo].[MarkTotal]" );
		pDatabase->ExecuteSQL( strSQL );

		// 生成 MarkTotal
		strSQL.Format( "CREATE TABLE [dbo].[MarkTotal] (\
					   [ID] [int] IDENTITY (1, 1) NOT NULL ,\
					   [StudentID] [int] NOT NULL ,\
					   [FormID] [int] NOT NULL ,\
					   [Score] [real] NULL ) ON [PRIMARY]" );
		pDatabase->ExecuteSQL( strSQL );

		// 删除旧的视图 vScoreCheck
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vScoreCheck]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vScoreCheck]" );
		pDatabase->ExecuteSQL( strSQL );

		// 生成视图 vScoreCheck
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

		// 删除旧的视图 vMainSubjectMarksFirst
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vMainSubjectMarksFirst]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vMainSubjectMarksFirst]" );
		pDatabase->ExecuteSQL( strSQL );

		// 生成视图 vMainSubjectMarksFirst
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

		// 删除旧的视图 vScoreFinal
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vScoreFinal]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vScoreFinal]" );
		pDatabase->ExecuteSQL( strSQL );

		// 生成视图 vScoreFinal
		strSQL.Format( "CREATE VIEW dbo.vScoreFinal\
					   AS\
					   SELECT dbo.ScoreFinal.ID, dbo.Sheet.StudentID, dbo.ScoreFinal.FormID, \
					   dbo.ScoreFinal.Score\
					   FROM dbo.ScoreFinal INNER JOIN\
					   dbo.CutSheet ON dbo.ScoreFinal.CutSheetID = dbo.CutSheet.CutSheetID INNER JOIN\
					   dbo.Sheet ON dbo.CutSheet.SheetID = dbo.Sheet.SheetID" );
		pDatabase->ExecuteSQL( strSQL );

		// 删除旧的视图 vMarksFirstSum
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vMarksFirstSum]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vMarksFirstSum]" );
		pDatabase->ExecuteSQL( strSQL );

		// 生成视图 vMarksFirstSum
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

		//// 删除旧的视图 
		//strSQL.Format( "" );
		//pDatabase->ExecuteSQL( strSQL );

		//// 生成视图 
		//strSQL.Format( "" );
		//pDatabase->ExecuteSQL( strSQL );
	}
	catch ( CDBException * e )
	{
		CString	strError;
		strError.Format( "生成大题成绩相关的表失败\r\n" );
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
		// 删除旧的 MarksFirst
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MarksFirst]') and OBJECTPROPERTY(id, N'IsUserTable') = 1) \
					   drop table [dbo].[MarksFirst]" );
		pDatabase->ExecuteSQL( strSQL );

		// 生成 MarksFirst
		strSQL.Format( "CREATE TABLE [dbo].[MarksFirst] (\
					   [ID] [int] IDENTITY (1, 1) NOT NULL ,\
					   [StudentID] [int] NOT NULL ,\
					   [FormID] [int] NOT NULL ,\
					   [Score] [real] NULL \
					   ) ON [PRIMARY]" );
		pDatabase->ExecuteSQL( strSQL );

		// 删除旧的视图 vMarksFirst0
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vMarksFirst0]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vMarksFirst0]" );
		pDatabase->ExecuteSQL( strSQL );

		// 生成视图 vMarksFirst0
		strSQL.Format( "CREATE VIEW dbo.vMarksFirst0\
					   AS\
					   SELECT DISTINCT StudentID\
					   FROM dbo.MarksFirst\
					   WHERE (Score > 0)" );
		pDatabase->ExecuteSQL( strSQL );

		// 删除旧的视图 vAVGMarkClass
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vAVGMarkClass]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vAVGMarkClass]" );
		pDatabase->ExecuteSQL( strSQL );

		// 生成视图 vAVGMarkClass
		strSQL.Format( "CREATE VIEW dbo.vAVGMarkClass\
					   AS\
					   SELECT dbo.Student.ClassID, dbo.MarksFirst.FormID, AVG(dbo.MarksFirst.Score) \
					   AS avgScore, dbo.Student.Type\
					   FROM dbo.MarksFirst INNER JOIN\
					   dbo.Student ON dbo.MarksFirst.StudentID = dbo.Student.StudentID INNER JOIN\
					   dbo.vMarksFirst0 ON dbo.Student.StudentID = dbo.vMarksFirst0.StudentID\
					   GROUP BY dbo.Student.ClassID, dbo.MarksFirst.FormID, dbo.Student.Type" );
		pDatabase->ExecuteSQL( strSQL );

		// 删除旧的视图 vAVGClass
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vAVGClass]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vAVGClass]" );
		pDatabase->ExecuteSQL( strSQL );

		// 生成视图 vAVGClass
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

		// 删除旧的视图 vAVGMarkSchool
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vAVGMarkSchool]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vAVGMarkSchool]" );
		pDatabase->ExecuteSQL( strSQL );

		// 生成视图 vAVGMarkSchool
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

		// 删除旧的视图 vAVGSchool
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vAVGSchool]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vAVGSchool]" );
		pDatabase->ExecuteSQL( strSQL );

		// 生成视图 vAVGSchool
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

		// 删除旧的视图 vAVGMarkRegion
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vAVGMarkRegion]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vAVGMarkRegion]" );
		pDatabase->ExecuteSQL( strSQL );

		// 生成视图 vAVGMarkRegion
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

		// 删除旧的视图 vAVGRegion
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vAVGRegion]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vAVGRegion]" );
		pDatabase->ExecuteSQL( strSQL );

		// 生成视图 vAVGRegion
		strSQL.Format( "CREATE VIEW dbo.vAVGRegion\
					   AS\
					   SELECT dbo.vAVGMarkRegion.RegionID, dbo.vAVGMarkRegion.FormID, \
					   dbo.vAVGMarkRegion.avgScore, dbo.QuestionFormInfo.SubjectID, - 1 AS CityID, \
					   '区' AS RegionName, dbo.vAVGMarkRegion.Type\
					   FROM dbo.vAVGMarkRegion INNER JOIN\
					   dbo.QuestionFormInfo ON \
					   dbo.vAVGMarkRegion.FormID = dbo.QuestionFormInfo.QuestionFormID" );
		pDatabase->ExecuteSQL( strSQL );

		// 删除旧的视图 vScoreCheck
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vScoreCheck]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vScoreCheck]" );
		pDatabase->ExecuteSQL( strSQL );

		// 生成视图 vScoreCheck
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

		// 删除旧的视图 vScoreFinal
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vScoreFinal]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vScoreFinal]" );
		pDatabase->ExecuteSQL( strSQL );

		// 生成视图 vScoreFinal
		strSQL.Format( "CREATE VIEW dbo.vScoreFinal\
					   AS\
					   SELECT dbo.ScoreFinal.ID, dbo.Sheet.StudentID, dbo.ScoreFinal.FormID, \
					   dbo.ScoreFinal.Score\
					   FROM dbo.ScoreFinal INNER JOIN\
					   dbo.CutSheet ON dbo.ScoreFinal.CutSheetID = dbo.CutSheet.CutSheetID INNER JOIN\
					   dbo.Sheet ON dbo.CutSheet.SheetID = dbo.Sheet.SheetID" );
		pDatabase->ExecuteSQL( strSQL );

		//// 删除旧的视图 
		//strSQL.Format( "" );
		//pDatabase->ExecuteSQL( strSQL );

		//// 生成视图 
		//strSQL.Format( "" );
		//pDatabase->ExecuteSQL( strSQL );
	}
	catch ( CDBException * e )
	{
		CString	strError;
		strError.Format( "生成相关的数据库对象失败\r\n" );
		strError	+= e->m_strError;
		AfxMessageBox( strError );
		return	FALSE;
	}
	
	return	TRUE;
}
