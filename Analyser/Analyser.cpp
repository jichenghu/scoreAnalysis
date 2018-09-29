// Analyser.cpp : ����Ӧ�ó��������Ϊ��
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


// CAnalyserApp ����

CAnalyserApp::CAnalyserApp()
{
	// TODO: �ڴ˴���ӹ�����룬
	// ��������Ҫ�ĳ�ʼ�������� InitInstance ��
}


// Ψһ��һ�� CAnalyserApp ����

CAnalyserApp theApp;


// CAnalyserApp ��ʼ��

BOOL CAnalyserApp::InitInstance()
{
	// Initialize OLE libraries
	if (!AfxOleInit())
	{
		AfxMessageBox( "��ʼ��OleInitʧ��" );
		return FALSE;
	}

	// ���һ�������� Windows XP �ϵ�Ӧ�ó����嵥ָ��Ҫ
	// ʹ�� ComCtl32.dll �汾 6 ����߰汾�����ÿ��ӻ���ʽ��
	//����Ҫ InitCommonControls()�����򣬽��޷��������ڡ�
	InitCommonControls();

	CWinApp::InitInstance();

	AfxEnableControlContainer();

	TRY
	{
		CString		strSQL	;

		if (!m_database.IsOpen())
			m_database.Open( "eTest", FALSE, FALSE, _T( "ODBC;UID=sa;PWD=87211237," ) );

		// ɾ���ɵ���ͼ vRegion
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vRegion]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vRegion]" );
		m_database.ExecuteSQL( strSQL );

		// ������ͼ vRegion
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

		// �������ڱ� MarksFirst �����ɴ˱�
		strSQL.Format( "if not exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MarksFirst]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)\
					   CREATE TABLE [dbo].[MarksFirst] (\
					   [ID] [int] IDENTITY (1, 1) NOT NULL ,\
					   [StudentID] [int] NOT NULL ,\
					   [FormID] [int] NOT NULL ,\
					   [Score] [real] NULL \
					   ) ON [PRIMARY]" );
		m_database.ExecuteSQL( strSQL );

//	SKLSE 2007-01-15, jicheng, �����ܷ�, begin
		// �������ڱ� ScoreSubject �����ɴ˱�
		strSQL.Format( "if not exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ScoreSubject]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)\
					   CREATE TABLE [dbo].[ScoreSubject] (\
					   [ID] [int] IDENTITY (1, 1) NOT NULL ,\
					   [StudentID] [int] NULL ,\
					   [FormID] [int] NULL ,\
					   [fValue] [real] NULL \
					   ) ON [PRIMARY]" );
		m_database.ExecuteSQL( strSQL );
//	SKLSE 2007-01-15, jicheng, �����ܷ�, ends

		// ɾ���ɵ���ͼ vSubject0
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vSubject0]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vSubject0]" );
		m_database.ExecuteSQL( strSQL );

		// ������ͼ vSubject0
		strSQL.Format( "CREATE VIEW dbo.vSubject0\
					   AS\
					   SELECT DISTINCT dbo.Subject.SubjectID, dbo.Subject.SubjectName\
					   FROM dbo.Subject INNER JOIN\
					   dbo.QuestionFormInfo ON \
					   dbo.Subject.SubjectID = dbo.QuestionFormInfo.SubjectID INNER JOIN\
					   dbo.MarksFirst ON dbo.QuestionFormInfo.QuestionFormID = dbo.MarksFirst.FormID" );
		m_database.ExecuteSQL( strSQL );

//	SKLSE 2007-01-15, jicheng, �����ܷ�, begin
		// ɾ���ɵ���ͼ vSubjectTotal
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vSubjectTotal]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vSubjectTotal]" );
		m_database.ExecuteSQL( strSQL );

		// ������ͼ vSubjectTotal
		strSQL.Format( "CREATE VIEW dbo.vSubjectTotal\
					   AS\
					   SELECT DISTINCT dbo.ScoreSubject.FormID, dbo.Subject.SubjectName\
					   FROM dbo.ScoreSubject INNER JOIN\
					   dbo.QuestionFormInfo ON \
					   dbo.ScoreSubject.FormID = dbo.QuestionFormInfo.QuestionFormID INNER JOIN\
					   dbo.Subject ON dbo.QuestionFormInfo.SubjectID = dbo.Subject.SubjectID" );
		m_database.ExecuteSQL( strSQL );
//	SKLSE 2007-01-15, jicheng, �����ܷ�, ends

		// ɾ���ɵ���ͼ vSubjectCnt
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vSubjectCnt]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vSubjectCnt]" );
		m_database.ExecuteSQL( strSQL );

		// ������ͼ vSubjectCnt
		strSQL.Format( "CREATE VIEW dbo.vSubjectCnt\
					   AS\
					   SELECT DISTINCT dbo.Subject.SubjectID, dbo.Subject.SubjectName\
					   FROM dbo.Subject INNER JOIN\
					   dbo.QuestionInfo ON dbo.Subject.SubjectID = dbo.QuestionInfo.SubjectID INNER JOIN\
					   dbo.ChoiceCnt ON dbo.QuestionInfo.QuestionID = dbo.ChoiceCnt.QuestionID" );
		m_database.ExecuteSQL( strSQL );

		// ɾ���ɵ���ͼ vSubSubject
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vSubject]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vSubject]" );
		m_database.ExecuteSQL( strSQL );

		// ������ͼ vSubSubject
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

		// ɾ���ɵ���ͼ vStudentType
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vStudentType]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vStudentType]" );
		m_database.ExecuteSQL( strSQL );

		// ������ͼ vStudentType
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

		// ���û�� ChoiceCnt �������ɸñ�
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

		// ɾ���ɵ���ͼ vChoiceInfo
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vChoiceInfo]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vChoiceInfo]" );
		m_database.ExecuteSQL( strSQL );

		// ������ͼ vChoiceInfo 
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

	// ��׼��ʼ��
	// ���δʹ����Щ���ܲ�ϣ����С
	// ���տ�ִ���ļ��Ĵ�С����Ӧ�Ƴ�����
	// ����Ҫ���ض���ʼ������
	// �������ڴ洢���õ�ע�����
	// TODO: Ӧ�ʵ��޸ĸ��ַ�����
	// �����޸�Ϊ��˾����֯��
	SetRegistryKey( _T("Jicheng @ Wuhan University") )	;

	CAnalyserDlg dlg;
	m_pMainWnd = &dlg;
	INT_PTR nResponse = dlg.DoModal();
	if (nResponse == IDOK)
	{
		// TODO: �ڴ˷��ô����ʱ�á�ȷ�������ر�
		//�Ի���Ĵ���
	}
	else if (nResponse == IDCANCEL)
	{
		// TODO: �ڴ˷��ô����ʱ�á�ȡ�������ر�
		//�Ի���Ĵ���
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

	// ���ڶԻ����ѹرգ����Խ����� FALSE �Ա��˳�Ӧ�ó���
	// ����������Ӧ�ó������Ϣ�á�
	return FALSE;
}
