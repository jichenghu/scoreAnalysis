// PageMultiSubjects.cpp : 实现文件
//

#include "stdafx.h"
#include "Analyser.h"
#include "PageMultiSubjects.h"
#include ".\pagemultisubjects.h"
#include "setScoreSubject.h"

#include "..\database\setSubject.h"

#include "setSchoolReport.h"

// CPageMultiSubjects 对话框

IMPLEMENT_DYNAMIC(CPageMultiSubjects, CDialog)
CPageMultiSubjects::CPageMultiSubjects(CWnd* pParent /*=NULL*/)
//	: CDialog(CPageMultiSubjects::IDD, pParent)
	: CDialogResize(CPageMultiSubjects::IDD, pParent)
{
}

CPageMultiSubjects::~CPageMultiSubjects()
{
}

void CPageMultiSubjects::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}


BEGIN_MESSAGE_MAP(CPageMultiSubjects, CDialogResize)
	ON_BN_CLICKED(IDC_BTN_INPUT, OnBnClickedBtnInput)
	ON_BN_CLICKED(IDC_BTN_GENERATE, OnBnClickedBtnGenerate)
END_MESSAGE_MAP()


// CPageMultiSubjects 消息处理程序

BOOL CPageMultiSubjects::OnInitDialog()
{
	CDialogResize::OnInitDialog();

	CSetSubject	setSubject	;

	setSubject.m_pDatabase	= &theApp.m_database;

	CString	strMsg;

	strMsg.Empty ();
	try
	{
		setSubject.Open ( CRecordset::snapshot, "SELECT * FROM vSubjectTotal", CRecordset::readOnly );

		while ( !setSubject.IsEOF() )
		{
			strMsg	= strMsg + "\r\n\r\n" + setSubject.m_SubjectName;
			setSubject.MoveNext ();
		}

		setSubject.Close ( )	;
	}
	catch ( CDBException *e )
	{
		AfxMessageBox ( e->m_strError + "\n导入数据失败" );
		return	FALSE;
	}

	GetDlgItem ( IDC_STATIC0 )->SetWindowText ( strMsg );

	return TRUE;  // return TRUE unless you set the focus to a control
	// 异常: OCX 属性页应返回 FALSE
}

void CPageMultiSubjects::InputSingleSubject ( CFile *pFile )
{
	int		iSubjectID	;
	long	lStudentID	;
	long	lFormID		;
	float	fValue		;

	pFile->Read ( &iSubjectID, sizeof(int) );

	CSetScoreSubject	setScoreSubject	;

	setScoreSubject.m_pDatabase	= &theApp.m_database;

	try
	{
		setScoreSubject.Open ( CRecordset::snapshot, "SELECT * FROM ScoreSubject", CRecordset::appendOnly )	;

		while ( pFile->Read ( &lStudentID, sizeof(long) ) )
		{
			pFile->Read ( &lFormID, sizeof(long) )	;
			pFile->Read ( &fValue, sizeof(float) )	;

			setScoreSubject.AddNew ( )	;
			setScoreSubject.m_StudentID	= lStudentID;
			setScoreSubject.m_FormID	= lFormID	;
			setScoreSubject.m_fValue	= fValue	;
			setScoreSubject.Update ( )	;
		}

		setScoreSubject.Close ( )	;
	}
	catch ( CDBException *e )
	{
		AfxMessageBox ( e->m_strError + "\n导入数据失败" );
		return	;
	}

	AfxMessageBox ( "数据导入成功！" )	;
}

void CPageMultiSubjects::InputCompositive ( CFile *pFile )
{
}

void CPageMultiSubjects::OnBnClickedBtnInput()
{
	CFile		fileInput	;
	CString		strPath		;

	CFileDialog	dlgPath ( TRUE, "*.sub", NULL, OFN_HIDEREADONLY, "单科总分数据文件 (*.sub)|*.sub||", NULL );

	if ( IDOK != dlgPath.DoModal ( ) )
		return;

	strPath	= dlgPath.GetPathName ( )	;
	fileInput.Open ( strPath, CFile::modeRead )	;

	BYTE	yMode	;
	fileInput.Read ( &yMode, sizeof(BYTE) )	;

	if ( 1 == yMode )	// 非综合科目
	{
		InputSingleSubject ( &fileInput )	;
	}
	else if ( 4 == yMode )	// 综合科目
	{
//		InputCompositive ( &fileInput )		;
		InputSingleSubject ( &fileInput )	;
	}
	else
		AfxMessageBox ( "格式错误！" )	;

	fileInput.Close ( )	;
}

void CPageMultiSubjects::OnBnClickedBtnGenerate()
{
	CSetSchoolReport	setReport	;
	CSetScoreSubject	setScore	;
	CSetSubject			setSubject	;
	CString				strSQL		;
	int					iCnt	= 0	;
	int					iFormID [16];
	CString				strCol [16]	;
//	CString				strCol		;
	CString				strSel		;

	setReport.Set_column_pointer ( );

	setReport.m_pDatabase	= &theApp.m_database;
	setScore.m_pDatabase	= &theApp.m_database;
	setSubject.m_pDatabase	= &theApp.m_database;

	try
	{
		// 删除旧的成绩单 SchoolReport
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SchoolReport]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)\
drop table [dbo].[SchoolReport]" );
		theApp.m_database.ExecuteSQL( strSQL );

		// 创建新的成绩单 SchoolReport
		strSQL.Format( "CREATE TABLE [dbo].[SchoolReport] (\
	[ID] [int] IDENTITY (1, 1) NOT NULL ,\
	[StudentID] [int] NULL \
) ON [PRIMARY]" );
		theApp.m_database.ExecuteSQL( strSQL );

		// 产生缺省的存储成绩的16个列
		for ( int i=1; i<17; i++ )
		{
			strSQL.Format( "ALTER TABLE SchoolReport ADD c%i real", i );
			theApp.m_database.ExecuteSQL( strSQL );
		}

		setSubject.Open ( CRecordset::snapshot, "SELECT * FROM vSubjectTotal", CRecordset::readOnly );
		setScore.Open ( CRecordset::snapshot, "SELECT * FROM ScoreSubject", CRecordset::readOnly );

//		strSel.Format ( "SELECT [dbo].[SchoolReport].ID, [dbo].[SchoolReport].StudentID, " );
		while ( !setSubject.IsEOF() )
		{
			iFormID [ iCnt ]	= setSubject.m_SubjectID;
			strCol [iCnt]		= setSubject.m_SubjectName;

//			strSQL.Format( "ALTER TABLE SchoolReport ADD %s real", setSubject.m_SubjectName );
//			theApp.m_database.ExecuteSQL( strSQL );

			iCnt++;

//			strCol.Format ( "[dbo].[SchoolReport].%s as c%d, ", setSubject.m_SubjectName, iCnt );

//			strSel	+= strCol;

			setSubject.MoveNext ();
		}

//		for ( int i=0; i<16; i++ )
//		{
//			strCol.Format ( "[dbo].[SchoolReport].c%d, ", i );
//
//			strSel	+= strCol;
//
//			strSQL.Format( "ALTER TABLE SchoolReport ADD c%i real", i );
//			theApp.m_database.ExecuteSQL( strSQL );
//		}
//
//		strCol.Format ( "[dbo].[SchoolReport].c16 " );
//
//		strSel	+= strCol;
//
//		strSQL.Format( "ALTER TABLE SchoolReport ADD c16 real" );
//		theApp.m_database.ExecuteSQL( strSQL );
//
//		strSel	= strSel + "FROM SchoolReport";

		setReport.Open ( CRecordset::snapshot, "SELECT * FROM SchoolReport" );
//		setReport.Open ( CRecordset::snapshot, strSel );

		int		iID;
		float *	fCol;
//		CString	strColumn;
		for ( i=1; i<=iCnt; i++ )
		{
			iID		= iFormID[i-1];
			fCol	= setReport.m_pColumn[i-1];
//			strColumn.Format ( "c%d", i );

			setScore.m_strFilter.Format ( "FormID=%d", iID );
			setScore.Requery ( );
			//while ( !setScore.IsEOF() )
			//{
			//	setReport.AddNew ( );

			//	setReport.m_StudentID	= setScore.m_StudentID;
			//	setReport.column1		= setScore.m_fValue;

			//	setReport.Update ( );

			//	setScore.MoveNext ( );
			//}

			while ( !setScore.IsEOF() )
			{
				setReport.m_strFilter.Format ( "(StudentID=%d)", setScore.m_StudentID );
				setReport.Requery ( );

				if ( setReport.IsEOF() )
				{
					setReport.AddNew ( );
				}
				else
					setReport.Edit ( );

				setReport.m_StudentID	= setScore.m_StudentID;
//				setReport.column[i-1]	= setScore.m_fValue;
				*fCol	= setScore.m_fValue;

				setReport.Update ( );

				setScore.MoveNext ( );
			}
		}

		for ( i=iCnt+1; i<17; i++ )
		{
			strSQL.Format( "ALTER TABLE SchoolReport DROP COLUMN c%i", i );
			theApp.m_database.ExecuteSQL( strSQL );
		}

		for ( i=1; i<=iCnt; i++ )
		{
			strSQL.Format( "sp_rename 'SchoolReport.c%i', '%s', 'COLUMN'", i, strCol [i-1] );
			theApp.m_database.ExecuteSQL( strSQL );
		}

		setSubject.Close ( );
	}
	catch ( CDBException *e )
	{
		AfxMessageBox ( e->m_strError + "\n创建成绩单数据失败" );
		return;
	}
	
	AfxMessageBox ( "完成\r\n请将表 SchoolReport 中的记录导出即可" );
}
