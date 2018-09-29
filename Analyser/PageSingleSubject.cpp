// PageSingleSubject.cpp : 实现文件
//

#include "stdafx.h"
#include "Analyser.h"
#include "Analyzer.h"
#include "AnalyserDlg.h"
#include "PageSingleSubject.h"
#include "API.h"
#include "vScoreSumForm.h"
#include <afxdlgs.h>

#include "setSelectID.h"
template <class T> bool BindSQL(T &wnd, CString sql, bool bAttach)
{
	CString			str		;
	CSetSelectID	setID	;

	wnd.ResetContent ( )	;

	setID.m_pDatabase	= &theApp.m_database;

	try
	{
		setID.Open ( CRecordset::snapshot, sql, CRecordset::readOnly )	;

		while ( !setID.IsEOF() )
		{
			wnd.AddString ( setID.m_strName )	;

			//int*	piID	= new int	;

			//*piID	= setID.m_iID	;

			//pMapID->SetAt( setID.m_strName, piID );

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

//////////////////////////////////////////////////////////////
// AddBind()
// 将数据集加入到列表框或组合框
// 参数 wnd:     列表控件或者组合框控件对象，要求支持ResetContent()和AddString()函数
//      sql:     数据的sql语句(注：该函数不进行sql语句有效性校验，请自行确保其正确性)
//               sql语句必须类似"select distinct [field] from tablename where ..."的形式
//      attach:  附加功能开关(添加首行'-')
//////////////////////////////////////////////////////////////
template <class T> bool AddBind(T &wnd, CString sql, CMapStringToPtr* pMapID, bool bAttach)
{
	CString			str		;
	CSetSelectID	setID	;

	setID.m_pDatabase	= &theApp.m_database;

	//wnd.ResetContent ( )	;

	POSITION	pos		;
	BOOL		bExist	;
	CString		key		;
	int *		piIndex	;
	int *		piID	;
	int			iID		;

	try
	{
		setID.Open ( CRecordset::snapshot, sql, CRecordset::readOnly )	;

		while ( !setID.IsEOF() )
		{
			iID	= setID.m_iID	;

			bExist	= FALSE	;
			for( pos = pMapID->GetStartPosition(); pos != NULL; )
			{
				pMapID->GetNextAssoc( pos, key, (void*&)piIndex );

				if ( iID == *piIndex )
				{
					bExist	= TRUE;

					break	;
				}
			}

			if ( bExist )
			{
				setID.MoveNext ( )	;

				continue;
			}

			wnd.AddString ( setID.m_strName )	;

			piID	= new int	;

			*piID	= iID		;

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

// CPageSingleSubject 对话框
IMPLEMENT_DYNAMIC(CPageSingleSubject, CDialog)
CPageSingleSubject::CPageSingleSubject(CWnd* pParent /*=NULL*/)
//	: CDialog(CPageSingleSubject::IDD, pParent)
	: CDialogResize(CPageSingleSubject::IDD, pParent)
{
	m_iSubjectID	= -1	;
	m_iTypeID		= -1	;
	m_iSubjectID2	= -1	;
//	m_iSubjectID3	= -1	;
}

CPageSingleSubject::~CPageSingleSubject()
{
}

void CPageSingleSubject::DoDataExchange(CDataExchange* pDX)
{
	//CDialog::DoDataExchange(pDX);
	CDialogResize::DoDataExchange(pDX);

	DDX_Control(pDX, IDC_COMBO_SUBJECT, m_ctlSubject)	;
	DDX_Control(pDX, IDC_COMBO_TYPE, m_ctlType)			;
	DDX_Control(pDX, IDC_COMBO_SUBJECT2, m_ctlSubject2)	;
//	DDX_Control(pDX, IDC_COMBO_SUBJECT3, m_ctlSubject3)	;
}


BEGIN_MESSAGE_MAP(CPageSingleSubject, CDialogResize)
	ON_WM_SIZE()
	ON_WM_DESTROY()
	ON_CBN_SELCHANGE(IDC_COMBO_SUBJECT, OnCbnSelchangeComboSubject)
	ON_CBN_SELCHANGE(IDC_COMBO_TYPE, OnCbnSelchangeComboType)
	ON_CBN_SELCHANGE(IDC_COMBO_SUBJECT2, OnCbnSelchangeComboSubject2)
//	ON_CBN_SELCHANGE(IDC_COMBO_SUBJECT3, OnCbnSelchangeComboSubject3)
	ON_BN_CLICKED(IDC_BTN_STEPANALYSE, OnBnClickedBtnStepanalyse)
	ON_BN_CLICKED(IDC_BTN_EXPORT_SCORES, OnBnClickedBtnExportScores)
	ON_BN_CLICKED(IDC_BTN_COMPOSITIVE, OnBnClickedBtnCompositive)
	ON_BN_CLICKED(IDC_BTN_EXPORT_SCORES2, OnBnClickedBtnExportScores2)
END_MESSAGE_MAP()


// CPageSingleSubject 消息处理程序

void CPageSingleSubject::OnSize(UINT nType, int cx, int cy) 
{
	CDialog::OnSize(nType, cx, cy);
	
	if (IsWindow(::GetDlgItem(m_hWnd, IDD_PROPPAGE_SINGLE_SUBJECT)))
	{
		CRect rect;
		GetClientRect(rect);
//		rect.DeflateRect(5,5,10,10);
//		m_TraceList.MoveWindow(rect);
	}
}

BOOL CPageSingleSubject::OnInitDialog()
{
	CDialogResize::OnInitDialog();

	UpdateData ( )	;

	return TRUE;  // return TRUE unless you set the focus to a control
	// 异常: OCX 属性页应返回 FALSE
}

void CPageSingleSubject::UpdateData ( )
{
	::BindSQL(m_ctlSubject, "select subjectID, subjectName from vSubject0 ORDER BY subjectID", &m_mapSubjectID, false)	;
	::BindSQL(m_ctlType, "select typeID, typeName from vStudentType ORDER BY typeID", &m_mapTypeID, false)				;
	::BindSQL(m_ctlSubject2, "select subjectID, subjectName from vSubject ORDER BY subjectID", false)				;

	AddBind ( m_ctlSubject, "select subjectID, subjectName from vSubjectCnt  ORDER BY subjectID", &m_mapSubjectID, false )	;

//	::BindSQL(m_ctlSubject3,   "select subjectID, subjectName from vSubSubject ORDER BY subjectID", false)	;

	OnCbnSelchangeComboSubject()	;
	OnCbnSelchangeComboSubject2()	;
	OnCbnSelchangeComboType()		;
}

void CPageSingleSubject::OnDestroy()
{
	CDialogResize::OnDestroy();

	POSITION	pos		;
	CString		key		;
	int *		lpiID	;

	for( pos = m_mapSubjectID.GetStartPosition(); pos != NULL; )
	{
		m_mapSubjectID.GetNextAssoc( pos, key, (void*&)lpiID );
#ifdef _DEBUG
//		afxDump << key << " : " << lpLevel << "\n";
#endif
		delete	lpiID;
	}

	for( pos = m_mapTypeID.GetStartPosition(); pos != NULL; )
	{
		m_mapTypeID.GetNextAssoc( pos, key, (void*&)lpiID );
#ifdef _DEBUG
//		afxDump << key << " : " << lpLevel << "\n";
#endif
		delete	lpiID;
	}
}

void CPageSingleSubject::OnCbnSelchangeComboSubject()
{
	int *	lpiID	;

	int	iSel	= m_ctlSubject.GetCurSel();

	if ( -1 == iSel )
	{
		m_iSubjectID	= -1;
		return;
	}

	m_ctlSubject.GetLBText(iSel, m_strSubject);

	//m_ctlSubject.GetLBText(m_ctlSubject.GetCurSel(), m_strSubject);

	m_mapSubjectID.Lookup( m_strSubject, (void*&)lpiID )	;

	m_iSubjectID	= *lpiID;
}

void CPageSingleSubject::OnCbnSelchangeComboSubject2()
{
	//int *	lpiID	;

	int	iSel	= m_ctlSubject2.GetCurSel();

	if ( -1 == iSel )
	{
		m_iSubjectID2	= -1;
		return;
	}

	m_ctlSubject2.GetLBText(iSel, m_strSubject2);

	if ( !m_strSubject2.Compare("理综") )
		m_iSubjectID2	= 10;
	else if ( !m_strSubject2.Compare("文综") )
		m_iSubjectID2	= 11;
	else
		m_iSubjectID2	= -1;

	//if ( m_mapSubjectID.Lookup( m_strSubject2, (void*&)lpiID ) )
	//	m_iSubjectID2	= *lpiID;
	//else
	//	m_iSubjectID2	= -1;
}

//void CPageSingleSubject::OnCbnSelchangeComboSubject3()
//{
//	int *	lpiID	;
//
//	m_ctlSubject3.GetLBText(m_ctlSubject3.GetCurSel(), m_strSubject3);
//
//	m_mapSubjectID.Lookup( m_strSubject3, (void*&)lpiID )	;
//
//	m_iSubjectID3	= *lpiID;
//}

void CPageSingleSubject::OnCbnSelchangeComboType()
{
	int *	lpiID	;

	int	iSel	= m_ctlType.GetCurSel();

	if ( -1 == iSel )
	{
		m_iTypeID	= -1;
		return;
	}

	m_ctlType.GetLBText(iSel, m_strType);

	//m_ctlType.GetLBText(m_ctlType.GetCurSel(), m_strType);

	m_mapTypeID.Lookup( m_strType, (void*&)lpiID )	;

	m_iTypeID	= *lpiID;
}

//++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//
//	分析按钮的响应
//
//		1.	得到分析问题列表
//		2.	循环产生分析表
//
//--------------------------------------------------------
void CPageSingleSubject::OnBnClickedBtnStepanalyse()
{
	CString		strRegion	;
	CString		strProfix	;
	CString		strDone		;
	CString		strOutput	;
	int			i			;
	int			iRegionID	;
	LPITEM		lpItem		;
	LPSUBITEM	lpSubItem	;

	iRegionID	= m_pParentDlg->GetRegionID ( )	;
	strRegion	= m_pParentDlg->GetRegionName ();

	OnCbnSelchangeComboSubject();
	OnCbnSelchangeComboType()	;

	ASSERT ( (iRegionID!=-1) && (m_iSubjectID!=-1) && (m_iTypeID!=-1) )	;

	CString		strSQL		;
	CDatabase *	pDatabase	;

	pDatabase	= &theApp.m_database;
	try
	{
		// 若 MarkSum 不存在则生成之
		strSQL.Format( "if not exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MarkSum]') and OBJECTPROPERTY(id, N'IsUserTable') = 1) \
					   CREATE TABLE [dbo].[MarkSum] (\
					   [ID] [int] IDENTITY (1, 1) NOT NULL ,\
					   [StudentID] [int] NOT NULL ,\
					   [FormID] [int] NOT NULL ,\
					   [Score] [real] NULL \
					   ) ON [PRIMARY]" );
		pDatabase->ExecuteSQL( strSQL );

		// 删除旧的视图 vScoreSumForm
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vScoreSumForm]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vScoreSumForm]" );
		pDatabase->ExecuteSQL( strSQL );

		// 生成视图 vScoreSumForm
		strSQL.Format( "CREATE VIEW dbo.vScoreSumForm\
					   AS\
					   SELECT dbo.MarkSum.StudentID, dbo.Student.StudentEnrollID, dbo.Student.StudentName, \
					   dbo.Student.ClassID, dbo.QuestionFormInfo.QuestionFormID, \
					   ROUND(dbo.MarkSum.Score, 0) AS fValue, dbo.QuestionFormInfo.SubjectID\
					   FROM dbo.Student INNER JOIN\
					   dbo.MarkSum ON dbo.Student.StudentID = dbo.MarkSum.StudentID INNER JOIN\
					   dbo.QuestionFormInfo ON \
					   dbo.MarkSum.FormID = dbo.QuestionFormInfo.QuestionFormID" );
		pDatabase->ExecuteSQL( strSQL );

		// 删除旧的视图 
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vScoreForm]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vScoreForm]" );
		pDatabase->ExecuteSQL( strSQL );

		// 生成视图 
		strSQL.Format( "CREATE VIEW dbo.vScoreForm\
					   AS\
					   SELECT dbo.MarksFirst.ID, dbo.MarksFirst.StudentID, dbo.MarksFirst.FormID, \
					   dbo.MarksFirst.Score, dbo.Student.ClassID, dbo.QuestionFormInfo.SubjectID\
					   FROM dbo.QuestionFormInfo INNER JOIN\
					   dbo.MarksFirst ON \
					   dbo.QuestionFormInfo.QuestionFormID = dbo.MarksFirst.FormID INNER JOIN\
					   dbo.Student ON dbo.MarksFirst.StudentID = dbo.Student.StudentID" );
		pDatabase->ExecuteSQL( strSQL );
	}
	catch ( CDBException * e )
	{
		CString	strError;
		strError.Format( "产生进行分析所需要的视图失败\r\n" );
		strError	+= e->m_strError;
		AfxMessageBox( strError );
		GetDlgItem ( IDC_BTN_STEPSCORE )->ShowWindow( SW_SHOW )	;
		return;
	}

	strProfix.Format ( "==> " )	;
	strDone.Format ( "√   " )	;

	i			= IDC_STATIC0						;
	lpItem		= m_pParentDlg->GetItemList ( )		;
	strOutput	= m_pParentDlg->GetPathOutput ( )	;
	while ( NULL != lpItem )
	{
		if ( (lpItem->bSelected) && (1==lpItem->yFormClass) )
		{
			GetDlgItem ( i )->SetWindowText ( strProfix + lpItem->cTitle )	;

//			Sleep ( 500 )	;

			// 处理当前表(对子表链表进行循环处理)
			lpSubItem	= lpItem->lpFirstSub;
			while ( NULL != lpSubItem )
			{
				CString	strSubjectTrim	;

				strSubjectTrim	= m_strSubject.Left ( 4 )	;
				// 找到当前选项所对应的SubItem
				if ( 0 == strSubjectTrim.Compare ( lpSubItem->cSubject ) )
				{
					// 读入该子项目的信息
					// 将之送入CAnalyzer产生表
					m_pParentDlg->m_pAnalyzer->AnalyzeSingleSubject ( m_iSubjectID, iRegionID, m_iTypeID,
						m_strSubject, strRegion, m_strType, lpSubItem, strOutput, lpItem->yLayers )	;

					break	;
				}

				lpSubItem	= lpSubItem->m_lpNext	;
			}

			GetDlgItem ( i )->SetWindowText ( strDone + lpItem->cTitle );

//			Sleep ( 500 )	;

			i++;
		}

		lpItem	= lpItem->lpNext;
	}

	try
	{
		// 删除视图 vScoreSumForm
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vScoreSumForm]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vScoreSumForm]" );
		pDatabase->ExecuteSQL( strSQL );

		// 删除视图 vScoreForm
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vScoreForm]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vScoreForm]" );
		pDatabase->ExecuteSQL( strSQL );
	}
	catch ( CDBException * e )
	{
		CString	strError;
		strError.Format( "分析完成后删除临时视图失败\r\n" );
		strError	+= e->m_strError;
		AfxMessageBox( strError );
		GetDlgItem ( IDC_BTN_STEPSCORE )->ShowWindow( SW_SHOW )	;
		return;
	}
}

BOOL CPageSingleSubject::Create(UINT nIDTemplate, CWnd* pParentWnd)
{
	BOOL	bReturn	;

	m_pParentDlg	= (CAnalyserDlg*)pParentWnd	;

	bReturn	= CDialogResize::Create(nIDTemplate, pParentWnd);

	return	bReturn	;
}

void CPageSingleSubject::UpdateUI ( )
{
	LPITEM	lpItem	;

	// 隐藏所有选项
	for ( int i=IDC_STATIC0; i<=IDC_STATIC7; i++ )
	{
		GetDlgItem ( i )->ShowWindow ( SW_HIDE );
	}

	i		= IDC_STATIC0					;
	lpItem	= m_pParentDlg->GetItemList ()	;
	while ( NULL != lpItem )
	{
		if ( lpItem->bSelected )
		{
			GetDlgItem ( i )->SetWindowText ( lpItem->cTitle )	;
			GetDlgItem ( i )->ShowWindow ( SW_SHOW )			;

			i++;
		}

		lpItem	= lpItem->lpNext;
	}
}

void CPageSingleSubject::OnBnClickedBtnExportScores()
{
	// 导出单科总分到文件中
	CFile			fileOutput	;
	CvScoreSumForm	viewScoreSum;
	CString			strPath		;
	CString			strRegion	;
	
	strRegion	= m_pParentDlg->GetRegionName ();

	CFileDialog	dlgPath ( FALSE, "*.sub", strRegion+m_strSubject, OFN_OVERWRITEPROMPT, "单科总分数据文件 (*.sub)|*.sub||", NULL );

	if ( IDOK != dlgPath.DoModal ( ) )
		return;

	strPath	= dlgPath.GetPathName ( )	;
	fileOutput.Open ( strPath, CFile::modeCreate | CFile::modeWrite )	;

	BYTE	byte	= 1	;	// 表示非综合科目(单科)
	fileOutput.Write ( &byte, sizeof ( BYTE ) )		;
	fileOutput.Write ( &m_iSubjectID, sizeof(int) )	;

	CString		strSQL		;
	CDatabase *	pDatabase	;

	pDatabase	= &theApp.m_database;

	viewScoreSum.m_pDatabase	= pDatabase;
	viewScoreSum.m_strFilter.Format ( "SubjectID = %d", m_iSubjectID )	;

	try
	{
		// 若 MarkSum 不存在则生成之
		strSQL.Format( "if not exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MarkSum]') and OBJECTPROPERTY(id, N'IsUserTable') = 1) \
					   CREATE TABLE [dbo].[MarkSum] (\
					   [ID] [int] IDENTITY (1, 1) NOT NULL ,\
					   [StudentID] [int] NOT NULL ,\
					   [FormID] [int] NOT NULL ,\
					   [Score] [real] NULL \
					   ) ON [PRIMARY]" );
		pDatabase->ExecuteSQL( strSQL );

		// 删除旧的视图 vScoreSumForm
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vScoreSumForm]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vScoreSumForm]" );
		pDatabase->ExecuteSQL( strSQL );

		// 生成视图 vScoreSumForm
		strSQL.Format( "CREATE VIEW dbo.vScoreSumForm\
					   AS\
					   SELECT dbo.MarkSum.StudentID, dbo.Student.StudentEnrollID, dbo.Student.StudentName, \
					   dbo.Student.ClassID, dbo.QuestionFormInfo.QuestionFormID, \
					   ROUND(dbo.MarkSum.Score, 0) AS fValue, dbo.QuestionFormInfo.SubjectID\
					   FROM dbo.Student INNER JOIN\
					   dbo.MarkSum ON dbo.Student.StudentID = dbo.MarkSum.StudentID INNER JOIN\
					   dbo.QuestionFormInfo ON \
					   dbo.MarkSum.FormID = dbo.QuestionFormInfo.QuestionFormID" );
		pDatabase->ExecuteSQL( strSQL );

		viewScoreSum.Open ( )	;

		while ( !viewScoreSum.IsEOF() )
		{
			fileOutput.Write ( &viewScoreSum.m_StudentID, sizeof(long) )		;
			fileOutput.Write ( &viewScoreSum.m_QuestionFormID, sizeof(long) )	;
			fileOutput.Write ( &viewScoreSum.m_fValue, sizeof(float) )			;

			viewScoreSum.MoveNext ( )	;
		}

		viewScoreSum.Close ( )	;

		// 删除视图 vScoreSumForm
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vScoreSumForm]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vScoreSumForm]" );
		pDatabase->ExecuteSQL( strSQL );
	}
	catch ( CDBException * e )
	{
		fileOutput.Close ( );

		AfxMessageBox ( e->m_strError + "\n数据导出失败！" )	;
		return	;
	}

	fileOutput.Close ( );

	AfxMessageBox ( "数据导出完毕！" )	;
}

void CPageSingleSubject::OnBnClickedBtnExportScores2()
{
	// 导出综合科目总分到文件中
	CFile			fileOutput	;
	CvScoreSumForm	viewScoreSum;
	CString			strPath		;
	CString			strRegion	;
	
	strRegion	= m_pParentDlg->GetRegionName ();

	CFileDialog	dlgPath ( FALSE, "*.sub", strRegion+m_strSubject, OFN_OVERWRITEPROMPT, "单科总分数据文件 (*.sub)|*.sub||", NULL );

	if ( IDOK != dlgPath.DoModal ( ) )
		return;

	strPath	= dlgPath.GetPathName ( )	;
	fileOutput.Open ( strPath, CFile::modeCreate | CFile::modeWrite )	;

	BYTE	byte	= 4	;	// 表示综合科目(每个综合科目含3科+综合分)
	fileOutput.Write ( &byte, sizeof ( BYTE ) )		;
	fileOutput.Write ( &m_iSubjectID, sizeof(int) )	;

	viewScoreSum.m_pDatabase	= &theApp.m_database;
	viewScoreSum.m_strFilter.Format ( "SubjectID = %d", m_iSubjectID )	;

	try
	{
		viewScoreSum.Open ( )	;

		while ( !viewScoreSum.IsEOF() )
		{
			fileOutput.Write ( &viewScoreSum.m_StudentID, sizeof(long) )		;
			fileOutput.Write ( &viewScoreSum.m_QuestionFormID, sizeof(long) )	;
			fileOutput.Write ( &viewScoreSum.m_fValue, sizeof(float) )			;

			viewScoreSum.MoveNext ( )	;
		}

		viewScoreSum.Close ( )	;
	}
	catch ( CDBException * e )
	{
		fileOutput.Close ( );

		AfxMessageBox ( e->m_strError + "\n数据导出失败！" )	;
		return	;
	}

	fileOutput.Close ( );

	AfxMessageBox ( "数据导出完毕！" )	;
}

/*+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
产生综合课报表
本方法将产生如下的数据库对象
	表 MarksForm
	表 MarkCompositive

	视图  
---------------------------------------------------------------------------*/
void CPageSingleSubject::OnBnClickedBtnCompositive()
{
	CString		strRegion	;
	CString		strProfix	;
	CString		strDone		;
	CString		strOutput	;
	int			i			;
	int			iRegionID	;
	LPITEM		lpItem		;
	LPSUBITEM	lpSubItem	;

	iRegionID	= m_pParentDlg->GetRegionID ( )	;
	strRegion	= m_pParentDlg->GetRegionName ();

	OnCbnSelchangeComboSubject();
	OnCbnSelchangeComboType()	;

	ASSERT ( (iRegionID!=-1) && (m_iSubjectID!=-1) && (m_iTypeID!=-1) )	;

	strProfix.Format ( "==> " )	;
	strDone.Format ( "√   " )	;

	CString		strSQL		;
	CDatabase *	pDatabase	;

	pDatabase	= &theApp.m_database;
	try
	{
		// 删除旧的 MarksForm
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MarksForm]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)\
					   drop table [dbo].[MarksForm]" );
		pDatabase->ExecuteSQL( strSQL );

		// 生成 MarksForm
		strSQL.Format( "CREATE TABLE [dbo].[MarksForm] (\
					   [ID] [int] IDENTITY (1, 1) NOT NULL ,\
					   [StudentID] [int] NOT NULL ,\
					   [FormID] [int] NOT NULL ,\
					   [Score] [real] NULL \
					   ) ON [PRIMARY]" );
		pDatabase->ExecuteSQL( strSQL );

		// 删除旧的 MarkCompositive
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MarkCompositive]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)\
					   drop table [dbo].[MarkCompositive]" );
		pDatabase->ExecuteSQL( strSQL );

		//// 生成 MarkCompositive
		//strSQL.Format( "CREATE TABLE [dbo].[MarkCompositive] (\
		//			   [ID] [int] IDENTITY (1, 1) NOT NULL ,\
		//			   [StudentID] [int] NOT NULL ,\
		//			   [FormID] [int] NOT NULL ,\
		//			   [Score] [real] NULL \
		//			   ) ON [PRIMARY]" );
		//pDatabase->ExecuteSQL( strSQL );

		// 删除旧的视图 vMarkSum
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vMarkSum]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vMarkSum]" );
		pDatabase->ExecuteSQL( strSQL );

		// 生成视图 vMarkSum
		strSQL.Format( "CREATE VIEW dbo.vMarkSum\
					   AS\
					   SELECT dbo.MarkSum.ID, dbo.MarkSum.StudentID, QuestionFormInfo_1.QuestionFormID, \
					   SUM(dbo.MarkSum.Score) AS Score\
					   FROM dbo.MarkSum INNER JOIN\
					   dbo.QuestionFormInfo ON \
					   dbo.MarkSum.FormID = dbo.QuestionFormInfo.QuestionFormID INNER JOIN\
					   dbo.subSubject ON \
					   dbo.QuestionFormInfo.SubjectID = dbo.subSubject.SubjectID INNER JOIN\
					   dbo.Subject ON dbo.subSubject.MainSubjectID = dbo.Subject.SubjectID INNER JOIN\
					   dbo.QuestionFormInfo QuestionFormInfo_1 ON \
					   dbo.Subject.SubjectID = QuestionFormInfo_1.SubjectID\
					   WHERE (dbo.QuestionFormInfo.Extend = 2)\
					   GROUP BY dbo.MarkSum.ID, dbo.MarkSum.StudentID, QuestionFormInfo_1.QuestionFormID" );
		pDatabase->ExecuteSQL( strSQL );

		strSQL.Format( "SELECT * INTO MarkCompositive FROM vMarkSum" );
		pDatabase->ExecuteSQL( strSQL );

		// 删除旧的视图 vScoreSumFormMain
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vScoreSumFormMain]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vScoreSumFormMain]" );
		pDatabase->ExecuteSQL( strSQL );

		// 生成视图 vScoreSumFormMain
		strSQL.Format( "CREATE VIEW dbo.vScoreSumFormMain\
					   AS\
					   SELECT dbo.MarkCompositive.StudentID, dbo.Student.StudentEnrollID, dbo.Student.StudentName, \
					   dbo.Student.ClassID, dbo.QuestionFormInfo.QuestionFormID, \
					   ROUND(dbo.MarkCompositive.Score, 0) AS fValue, dbo.QuestionFormInfo.SubjectID\
					   FROM dbo.Student INNER JOIN\
					   dbo.MarkCompositive ON dbo.Student.StudentID = dbo.MarkCompositive.StudentID INNER JOIN\
					   dbo.QuestionFormInfo ON \
					   dbo.MarkCompositive.QuestionFormID = dbo.QuestionFormInfo.QuestionFormID" );
		pDatabase->ExecuteSQL( strSQL );

		// 删除旧的视图 vScoreFormMain
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vScoreFormMain]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vScoreFormMain]" );
		pDatabase->ExecuteSQL( strSQL );

		// 生成视图 vScoreFormMain
		strSQL.Format( "CREATE VIEW dbo.vScoreFormMain\
					   AS\
					   SELECT dbo.MarksForm.ID, dbo.MarksForm.StudentID, dbo.MarksForm.FormID, \
					   dbo.MarksForm.Score, dbo.Student.ClassID, \
					   dbo.subSubject.MainSubjectID AS SubjectID\
					   FROM dbo.QuestionFormInfo INNER JOIN\
					   dbo.MarksForm ON \
					   dbo.QuestionFormInfo.QuestionFormID = dbo.MarksForm.FormID INNER JOIN\
					   dbo.Student ON dbo.MarksForm.StudentID = dbo.Student.StudentID INNER JOIN\
					   dbo.subSubject ON dbo.QuestionFormInfo.SubjectID = dbo.subSubject.SubjectID" );
		pDatabase->ExecuteSQL( strSQL );
	}
	catch ( CDBException * e )
	{
		CString	strError;
		strError.Format( "产生进行分析所需要的数据库对象失败\r\n" );
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

					strDTS.Format( "dts\\MarksForm.dts" )	;
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

					strError.Format( "调用 MarksForm.dts 失败\r\n%s Error: %ld\r\n%s\r\n%s\r\n",
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

			strError.Format( "调用 MarksForm.dts 失败\r\n%s Error: %ld\r\n%s\r\n%s\r\n",
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

	i			= IDC_STATIC0						;
	lpItem		= m_pParentDlg->GetItemList ( )		;
	strOutput	= m_pParentDlg->GetPathOutput ( )	;
	while ( NULL != lpItem )
	{
		if ( (lpItem->bSelected) && (1==lpItem->yFormClass) )
		{
			GetDlgItem ( i )->SetWindowText ( strProfix + lpItem->cTitle )	;

			// 处理当前表(对子表链表进行循环处理)
			lpSubItem	= lpItem->lpFirstSub;
			while ( NULL != lpSubItem )
			{
				CString	strSubjectTrim	;

				//int	*lpiID;
				strSubjectTrim	= m_strSubject2.Left ( 4 )	;
				//if ( m_mapSubjectID.Lookup( m_strSubject2, (void*&)lpiID ) )
				//	m_iSubjectID2	= *lpiID;
				//else
				//	m_iSubjectID2	= -1;
				// 找到当前选项所对应的SubItem
				if ( 0 == strSubjectTrim.Compare ( lpSubItem->cSubject ) )
				{
					// 读入该子项目的信息
					// 将之送入CAnalyzer产生表
					m_pParentDlg->m_pAnalyzer->AnalyzeSingleSubject ( m_iSubjectID2, iRegionID, m_iTypeID,
						m_strSubject2, strRegion, m_strType, lpSubItem, strOutput, lpItem->yLayers )	;

					break	;
				}

				lpSubItem	= lpSubItem->m_lpNext	;
			}

			GetDlgItem ( i )->SetWindowText ( strDone + lpItem->cTitle );

			i++;
		}

		lpItem	= lpItem->lpNext;
	}

	try
	{
		// 删除视图 vScoreSumForm
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vScoreSumForm]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vScoreSumForm]" );
		pDatabase->ExecuteSQL( strSQL );

		// 删除视图 vScoreForm
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vScoreForm]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vScoreForm]" );
		pDatabase->ExecuteSQL( strSQL );

		// 删除 MarksForm
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MarksForm]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)\
					   drop table [dbo].[MarksForm]" );
		pDatabase->ExecuteSQL( strSQL );

		// 删除 MarkCompositive
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MarkCompositive]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)\
					   drop table [dbo].[MarkCompositive]" );
		pDatabase->ExecuteSQL( strSQL );

		// 删除视图 vMarkSum
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vMarkSum]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vMarkSum]" );
		pDatabase->ExecuteSQL( strSQL );

		// 删除视图 vScoreSumFormMain
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vScoreSumFormMain]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vScoreSumFormMain]" );
		pDatabase->ExecuteSQL( strSQL );

		// 删除视图 vScoreFormMain
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vScoreFormMain]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vScoreFormMain]" );
		pDatabase->ExecuteSQL( strSQL );
	}
	catch ( CDBException * e )
	{
		CString	strError;
		strError.Format( "分析完成后删除临时数据库对象失败\r\n" );
		strError	+= e->m_strError;
		AfxMessageBox( strError );
		GetDlgItem ( IDC_BTN_STEPSCORE )->ShowWindow( SW_SHOW )	;
		return;
	}
}
