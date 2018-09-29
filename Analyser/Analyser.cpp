// Analyser.cpp : 定义应用程序的类行为。
//

#include "stdafx.h"
#include "Analyser.h"
#include "AnalyserDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CAnalyserApp

BEGIN_MESSAGE_MAP(CAnalyserApp, CWinApp)
	ON_COMMAND(ID_HELP, CWinApp::OnHelp)
END_MESSAGE_MAP()


// CAnalyserApp 构造

CAnalyserApp::CAnalyserApp()
{
	// TODO: 在此处添加构造代码，
	// 将所有重要的初始化放置在 InitInstance 中
}


// 唯一的一个 CAnalyserApp 对象

CAnalyserApp theApp;


// CAnalyserApp 初始化

BOOL CAnalyserApp::InitInstance()
{
	// Initialize OLE libraries
	if (!AfxOleInit())
	{
		AfxMessageBox( "初始化OleInit失败" );
		return FALSE;
	}

	// 如果一个运行在 Windows XP 上的应用程序清单指定要
	// 使用 ComCtl32.dll 版本 6 或更高版本来启用可视化方式，
	//则需要 InitCommonControls()。否则，将无法创建窗口。
	InitCommonControls();

	CWinApp::InitInstance();

	AfxEnableControlContainer();

	TRY
	{
		CString		strSQL	;

		if (!m_database.IsOpen())
			m_database.Open( "eTest", FALSE, FALSE, _T( "ODBC;UID=sa;PWD=87211237," ) );

		// 删除旧的视图 vRegion
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vRegion]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vRegion]" );
		m_database.ExecuteSQL( strSQL );

		// 生成视图 vRegion
		strSQL.Format( "CREATE VIEW dbo.vRegion\
					   AS\
					   SELECT DISTINCT dbo.Region.RegionID, dbo.Region.RegionName\
					   FROM dbo._ScoreStudentStep0 INNER JOIN\
					   dbo.CutSheet ON \
					   dbo._ScoreStudentStep0.CutSheetID = dbo.CutSheet.CutSheetID INNER JOIN\
					   dbo.Sheet ON dbo.CutSheet.SheetID = dbo.Sheet.SheetID INNER JOIN\
					   dbo.ScanGroup ON \
					   dbo.Sheet.ScanGroupID = dbo.ScanGroup.ScanGroupID INNER JOIN\
					   dbo.Region ON dbo.ScanGroup.RegionID = dbo.Region.RegionID" );
		m_database.ExecuteSQL( strSQL );

		// 若不存在表 MarksFirst 则生成此表
		strSQL.Format( "if not exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MarksFirst]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)\
					   CREATE TABLE [dbo].[MarksFirst] (\
					   [ID] [int] IDENTITY (1, 1) NOT NULL ,\
					   [StudentID] [int] NOT NULL ,\
					   [FormID] [int] NOT NULL ,\
					   [Score] [real] NULL \
					   ) ON [PRIMARY]" );
		m_database.ExecuteSQL( strSQL );

//	SKLSE 2007-01-15, jicheng, 计算总分, begin
		// 若不存在表 ScoreSubject 则生成此表
		strSQL.Format( "if not exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ScoreSubject]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)\
					   CREATE TABLE [dbo].[ScoreSubject] (\
					   [ID] [int] IDENTITY (1, 1) NOT NULL ,\
					   [StudentID] [int] NULL ,\
					   [FormID] [int] NULL ,\
					   [fValue] [real] NULL \
					   ) ON [PRIMARY]" );
		m_database.ExecuteSQL( strSQL );
//	SKLSE 2007-01-15, jicheng, 计算总分, ends

		// 删除旧的视图 vSubject0
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vSubject0]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vSubject0]" );
		m_database.ExecuteSQL( strSQL );

		// 生成视图 vSubject0
		strSQL.Format( "CREATE VIEW dbo.vSubject0\
					   AS\
					   SELECT DISTINCT dbo.Subject.SubjectID, dbo.Subject.SubjectName\
					   FROM dbo.Subject INNER JOIN\
					   dbo.QuestionFormInfo ON \
					   dbo.Subject.SubjectID = dbo.QuestionFormInfo.SubjectID INNER JOIN\
					   dbo.MarksFirst ON dbo.QuestionFormInfo.QuestionFormID = dbo.MarksFirst.FormID" );
		m_database.ExecuteSQL( strSQL );

//	SKLSE 2007-01-15, jicheng, 计算总分, begin
		// 删除旧的视图 vSubjectTotal
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vSubjectTotal]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vSubjectTotal]" );
		m_database.ExecuteSQL( strSQL );

		// 生成视图 vSubjectTotal
		strSQL.Format( "CREATE VIEW dbo.vSubjectTotal\
					   AS\
					   SELECT DISTINCT dbo.ScoreSubject.FormID, dbo.Subject.SubjectName\
					   FROM dbo.ScoreSubject INNER JOIN\
					   dbo.QuestionFormInfo ON \
					   dbo.ScoreSubject.FormID = dbo.QuestionFormInfo.QuestionFormID INNER JOIN\
					   dbo.Subject ON dbo.QuestionFormInfo.SubjectID = dbo.Subject.SubjectID" );
		m_database.ExecuteSQL( strSQL );
//	SKLSE 2007-01-15, jicheng, 计算总分, ends

		// 删除旧的视图 vSubjectCnt
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vSubjectCnt]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vSubjectCnt]" );
		m_database.ExecuteSQL( strSQL );

		// 生成视图 vSubjectCnt
		strSQL.Format( "CREATE VIEW dbo.vSubjectCnt\
					   AS\
					   SELECT DISTINCT dbo.Subject.SubjectID, dbo.Subject.SubjectName\
					   FROM dbo.Subject INNER JOIN\
					   dbo.QuestionInfo ON dbo.Subject.SubjectID = dbo.QuestionInfo.SubjectID INNER JOIN\
					   dbo.ChoiceCnt ON dbo.QuestionInfo.QuestionID = dbo.ChoiceCnt.QuestionID" );
		m_database.ExecuteSQL( strSQL );

		// 删除旧的视图 vSubSubject
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vSubject]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vSubject]" );
		m_database.ExecuteSQL( strSQL );

		// 生成视图 vSubSubject
		strSQL.Format( "CREATE VIEW dbo.vSubject\
					   AS\
					   SELECT DISTINCT dbo.Subject.SubjectID, dbo.Subject.SubjectName\
					   FROM dbo.Subject INNER JOIN\
					   dbo.subSubject ON \
					   dbo.Subject.SubjectID = dbo.subSubject.MainSubjectID INNER JOIN\
					   dbo.QuestionFormInfo ON \
					   dbo.subSubject.SubjectID = dbo.QuestionFormInfo.SubjectID INNER JOIN\
					   dbo.MarksFirst ON dbo.QuestionFormInfo.QuestionFormID = dbo.MarksFirst.FormID" );
		m_database.ExecuteSQL( strSQL );

		// 删除旧的视图 vStudentType
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vStudentType]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vStudentType]" );
		m_database.ExecuteSQL( strSQL );

		// 生成视图 vStudentType
		strSQL.Format( "CREATE VIEW dbo.vStudentType\
					   AS\
					   SELECT DISTINCT dbo.StudentType.TypeID, dbo.StudentType.TypeName\
					   FROM dbo.StudentType INNER JOIN\
					   dbo.Student ON dbo.StudentType.TypeID = dbo.Student.Type INNER JOIN\
					   dbo.ScanGroup ON dbo.Student.RoomID = dbo.ScanGroup.RoomID INNER JOIN\
					   dbo.Sheet ON dbo.ScanGroup.ScanGroupID = dbo.Sheet.ScanGroupID INNER JOIN\
					   dbo.CutSheet ON dbo.Sheet.SheetID = dbo.CutSheet.SheetID INNER JOIN\
					   dbo.ScoreFinal ON dbo.CutSheet.CutSheetID = dbo.ScoreFinal.CutSheetID" );
		m_database.ExecuteSQL( strSQL );

		// 如果没有 ChoiceCnt 表则生成该表
		strSQL.Format( "if not exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ChoiceCnt]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)\
					   CREATE TABLE [dbo].[ChoiceCnt] (\
					   [ClassID] [int] NOT NULL ,\
					   [QuestionID] [int] NOT NULL ,\
					   [Step] [tinyint] NOT NULL ,\
					   [Choice] [tinyint] NOT NULL ,\
					   [Cnt] [int] NOT NULL ,\
					   [CntStudent] [float] NOT NULL \
					   ) ON [PRIMARY]" );
		m_database.ExecuteSQL( strSQL );

		// 删除旧的视图 vChoiceInfo
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vChoiceInfo]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vChoiceInfo]" );
		m_database.ExecuteSQL( strSQL );

		// 生成视图 vChoiceInfo 
		strSQL.Format( "CREATE VIEW dbo.vChoiceInfo\
					   AS\
					   SELECT TOP 100 PERCENT dbo.Sheet.StudentID, dbo.QuestionInfo.SubjectID, \
					   dbo.Student.RoomID, dbo.Sheet.SheetID, dbo.CutSheet.QuestionID, \
					   dbo.Student.ClassID\
					   FROM dbo.CutSheet INNER JOIN\
					   dbo.Sheet ON dbo.CutSheet.SheetID = dbo.Sheet.SheetID INNER JOIN\
					   dbo.Student ON dbo.Sheet.StudentID = dbo.Student.StudentID INNER JOIN\
					   dbo.QuestionInfo ON dbo.CutSheet.QuestionID = dbo.QuestionInfo.QuestionID\
					   WHERE (dbo.CutSheet.Flag = 2)\
					   ORDER BY dbo.CutSheet.QuestionID, dbo.Student.ClassID" );
		m_database.ExecuteSQL( strSQL );
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(e->m_strError);
		return	FALSE;
	}
	END_CATCH

	// 标准初始化
	// 如果未使用这些功能并希望减小
	// 最终可执行文件的大小，则应移除下列
	// 不需要的特定初始化例程
	// 更改用于存储设置的注册表项
	// TODO: 应适当修改该字符串，
	// 例如修改为公司或组织名
	SetRegistryKey( _T("Jicheng @ Wuhan University") )	;

	CAnalyserDlg dlg;
	m_pMainWnd = &dlg;
	INT_PTR nResponse = dlg.DoModal();
	if (nResponse == IDOK)
	{
		// TODO: 在此放置处理何时用“确定”来关闭
		//对话框的代码
	}
	else if (nResponse == IDCANCEL)
	{
		// TODO: 在此放置处理何时用“取消”来关闭
		//对话框的代码
	}

	TRY
	{
		if (m_database.IsOpen())
			m_database.Close( );
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(e->m_strError);
		return	FALSE;
	}
	END_CATCH

	// 由于对话框已关闭，所以将返回 FALSE 以便退出应用程序，
	// 而不是启动应用程序的消息泵。
	return FALSE;
}
