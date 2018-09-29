// PageSingleSubject.cpp : ʵ���ļ�
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

	//ѡ���һ��
	if( wnd.GetCount() > 0 ) 
		wnd.SetCurSel( 0 )	;

	return true;
}

//////////////////////////////////////////////////////////////
// AddBind()
// �����ݼ����뵽�б�����Ͽ�
// ���� wnd:     �б�ؼ�������Ͽ�ؼ�����Ҫ��֧��ResetContent()��AddString()����
//      sql:     ���ݵ�sql���(ע���ú���������sql�����Ч��У�飬������ȷ������ȷ��)
//               sql����������"select distinct [field] from tablename where ..."����ʽ
//      attach:  ���ӹ��ܿ���(�������'-')
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

	//ѡ���һ��
	if( wnd.GetCount() > 0 ) 
		wnd.SetCurSel( 0 )	;

	return true;
}

// CPageSingleSubject �Ի���
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


// CPageSingleSubject ��Ϣ�������

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
	// �쳣: OCX ����ҳӦ���� FALSE
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

	if ( !m_strSubject2.Compare("����") )
		m_iSubjectID2	= 10;
	else if ( !m_strSubject2.Compare("����") )
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
//	������ť����Ӧ
//
//		1.	�õ����������б�
//		2.	ѭ������������
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
		// �� MarkSum ������������֮
		strSQL.Format( "if not exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MarkSum]') and OBJECTPROPERTY(id, N'IsUserTable') = 1) \
					   CREATE TABLE [dbo].[MarkSum] (\
					   [ID] [int] IDENTITY (1, 1) NOT NULL ,\
					   [StudentID] [int] NOT NULL ,\
					   [FormID] [int] NOT NULL ,\
					   [Score] [real] NULL \
					   ) ON [PRIMARY]" );
		pDatabase->ExecuteSQL( strSQL );

		// ɾ���ɵ���ͼ vScoreSumForm
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vScoreSumForm]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vScoreSumForm]" );
		pDatabase->ExecuteSQL( strSQL );

		// ������ͼ vScoreSumForm
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

		// ɾ���ɵ���ͼ 
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vScoreForm]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vScoreForm]" );
		pDatabase->ExecuteSQL( strSQL );

		// ������ͼ 
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
		strError.Format( "�������з�������Ҫ����ͼʧ��\r\n" );
		strError	+= e->m_strError;
		AfxMessageBox( strError );
		GetDlgItem ( IDC_BTN_STEPSCORE )->ShowWindow( SW_SHOW )	;
		return;
	}

	strProfix.Format ( "==> " )	;
	strDone.Format ( "��   " )	;

	i			= IDC_STATIC0						;
	lpItem		= m_pParentDlg->GetItemList ( )		;
	strOutput	= m_pParentDlg->GetPathOutput ( )	;
	while ( NULL != lpItem )
	{
		if ( (lpItem->bSelected) && (1==lpItem->yFormClass) )
		{
			GetDlgItem ( i )->SetWindowText ( strProfix + lpItem->cTitle )	;

//			Sleep ( 500 )	;

			// ����ǰ��(���ӱ��������ѭ������)
			lpSubItem	= lpItem->lpFirstSub;
			while ( NULL != lpSubItem )
			{
				CString	strSubjectTrim	;

				strSubjectTrim	= m_strSubject.Left ( 4 )	;
				// �ҵ���ǰѡ������Ӧ��SubItem
				if ( 0 == strSubjectTrim.Compare ( lpSubItem->cSubject ) )
				{
					// ���������Ŀ����Ϣ
					// ��֮����CAnalyzer������
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
		// ɾ����ͼ vScoreSumForm
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vScoreSumForm]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vScoreSumForm]" );
		pDatabase->ExecuteSQL( strSQL );

		// ɾ����ͼ vScoreForm
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vScoreForm]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vScoreForm]" );
		pDatabase->ExecuteSQL( strSQL );
	}
	catch ( CDBException * e )
	{
		CString	strError;
		strError.Format( "������ɺ�ɾ����ʱ��ͼʧ��\r\n" );
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

	// ��������ѡ��
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
	// ���������ֵܷ��ļ���
	CFile			fileOutput	;
	CvScoreSumForm	viewScoreSum;
	CString			strPath		;
	CString			strRegion	;
	
	strRegion	= m_pParentDlg->GetRegionName ();

	CFileDialog	dlgPath ( FALSE, "*.sub", strRegion+m_strSubject, OFN_OVERWRITEPROMPT, "�����ܷ������ļ� (*.sub)|*.sub||", NULL );

	if ( IDOK != dlgPath.DoModal ( ) )
		return;

	strPath	= dlgPath.GetPathName ( )	;
	fileOutput.Open ( strPath, CFile::modeCreate | CFile::modeWrite )	;

	BYTE	byte	= 1	;	// ��ʾ���ۺϿ�Ŀ(����)
	fileOutput.Write ( &byte, sizeof ( BYTE ) )		;
	fileOutput.Write ( &m_iSubjectID, sizeof(int) )	;

	CString		strSQL		;
	CDatabase *	pDatabase	;

	pDatabase	= &theApp.m_database;

	viewScoreSum.m_pDatabase	= pDatabase;
	viewScoreSum.m_strFilter.Format ( "SubjectID = %d", m_iSubjectID )	;

	try
	{
		// �� MarkSum ������������֮
		strSQL.Format( "if not exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MarkSum]') and OBJECTPROPERTY(id, N'IsUserTable') = 1) \
					   CREATE TABLE [dbo].[MarkSum] (\
					   [ID] [int] IDENTITY (1, 1) NOT NULL ,\
					   [StudentID] [int] NOT NULL ,\
					   [FormID] [int] NOT NULL ,\
					   [Score] [real] NULL \
					   ) ON [PRIMARY]" );
		pDatabase->ExecuteSQL( strSQL );

		// ɾ���ɵ���ͼ vScoreSumForm
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vScoreSumForm]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vScoreSumForm]" );
		pDatabase->ExecuteSQL( strSQL );

		// ������ͼ vScoreSumForm
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

		// ɾ����ͼ vScoreSumForm
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vScoreSumForm]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vScoreSumForm]" );
		pDatabase->ExecuteSQL( strSQL );
	}
	catch ( CDBException * e )
	{
		fileOutput.Close ( );

		AfxMessageBox ( e->m_strError + "\n���ݵ���ʧ�ܣ�" )	;
		return	;
	}

	fileOutput.Close ( );

	AfxMessageBox ( "���ݵ�����ϣ�" )	;
}

void CPageSingleSubject::OnBnClickedBtnExportScores2()
{
	// �����ۺϿ�Ŀ�ֵܷ��ļ���
	CFile			fileOutput	;
	CvScoreSumForm	viewScoreSum;
	CString			strPath		;
	CString			strRegion	;
	
	strRegion	= m_pParentDlg->GetRegionName ();

	CFileDialog	dlgPath ( FALSE, "*.sub", strRegion+m_strSubject, OFN_OVERWRITEPROMPT, "�����ܷ������ļ� (*.sub)|*.sub||", NULL );

	if ( IDOK != dlgPath.DoModal ( ) )
		return;

	strPath	= dlgPath.GetPathName ( )	;
	fileOutput.Open ( strPath, CFile::modeCreate | CFile::modeWrite )	;

	BYTE	byte	= 4	;	// ��ʾ�ۺϿ�Ŀ(ÿ���ۺϿ�Ŀ��3��+�ۺϷ�)
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

		AfxMessageBox ( e->m_strError + "\n���ݵ���ʧ�ܣ�" )	;
		return	;
	}

	fileOutput.Close ( );

	AfxMessageBox ( "���ݵ�����ϣ�" )	;
}

/*+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
�����ۺϿα���
���������������µ����ݿ����
	�� MarksForm
	�� MarkCompositive

	��ͼ  
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
	strDone.Format ( "��   " )	;

	CString		strSQL		;
	CDatabase *	pDatabase	;

	pDatabase	= &theApp.m_database;
	try
	{
		// ɾ���ɵ� MarksForm
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MarksForm]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)\
					   drop table [dbo].[MarksForm]" );
		pDatabase->ExecuteSQL( strSQL );

		// ���� MarksForm
		strSQL.Format( "CREATE TABLE [dbo].[MarksForm] (\
					   [ID] [int] IDENTITY (1, 1) NOT NULL ,\
					   [StudentID] [int] NOT NULL ,\
					   [FormID] [int] NOT NULL ,\
					   [Score] [real] NULL \
					   ) ON [PRIMARY]" );
		pDatabase->ExecuteSQL( strSQL );

		// ɾ���ɵ� MarkCompositive
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MarkCompositive]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)\
					   drop table [dbo].[MarkCompositive]" );
		pDatabase->ExecuteSQL( strSQL );

		//// ���� MarkCompositive
		//strSQL.Format( "CREATE TABLE [dbo].[MarkCompositive] (\
		//			   [ID] [int] IDENTITY (1, 1) NOT NULL ,\
		//			   [StudentID] [int] NOT NULL ,\
		//			   [FormID] [int] NOT NULL ,\
		//			   [Score] [real] NULL \
		//			   ) ON [PRIMARY]" );
		//pDatabase->ExecuteSQL( strSQL );

		// ɾ���ɵ���ͼ vMarkSum
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vMarkSum]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vMarkSum]" );
		pDatabase->ExecuteSQL( strSQL );

		// ������ͼ vMarkSum
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

		// ɾ���ɵ���ͼ vScoreSumFormMain
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vScoreSumFormMain]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vScoreSumFormMain]" );
		pDatabase->ExecuteSQL( strSQL );

		// ������ͼ vScoreSumFormMain
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

		// ɾ���ɵ���ͼ vScoreFormMain
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vScoreFormMain]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vScoreFormMain]" );
		pDatabase->ExecuteSQL( strSQL );

		// ������ͼ vScoreFormMain
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
		strError.Format( "�������з�������Ҫ�����ݿ����ʧ��\r\n" );
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

					strError.Format( "���� MarksForm.dts ʧ��\r\n%s Error: %ld\r\n%s\r\n%s\r\n",
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

			strError.Format( "���� MarksForm.dts ʧ��\r\n%s Error: %ld\r\n%s\r\n%s\r\n",
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

	i			= IDC_STATIC0						;
	lpItem		= m_pParentDlg->GetItemList ( )		;
	strOutput	= m_pParentDlg->GetPathOutput ( )	;
	while ( NULL != lpItem )
	{
		if ( (lpItem->bSelected) && (1==lpItem->yFormClass) )
		{
			GetDlgItem ( i )->SetWindowText ( strProfix + lpItem->cTitle )	;

			// ����ǰ��(���ӱ��������ѭ������)
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
				// �ҵ���ǰѡ������Ӧ��SubItem
				if ( 0 == strSubjectTrim.Compare ( lpSubItem->cSubject ) )
				{
					// ���������Ŀ����Ϣ
					// ��֮����CAnalyzer������
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
		// ɾ����ͼ vScoreSumForm
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vScoreSumForm]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vScoreSumForm]" );
		pDatabase->ExecuteSQL( strSQL );

		// ɾ����ͼ vScoreForm
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vScoreForm]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vScoreForm]" );
		pDatabase->ExecuteSQL( strSQL );

		// ɾ�� MarksForm
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MarksForm]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)\
					   drop table [dbo].[MarksForm]" );
		pDatabase->ExecuteSQL( strSQL );

		// ɾ�� MarkCompositive
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MarkCompositive]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)\
					   drop table [dbo].[MarkCompositive]" );
		pDatabase->ExecuteSQL( strSQL );

		// ɾ����ͼ vMarkSum
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vMarkSum]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vMarkSum]" );
		pDatabase->ExecuteSQL( strSQL );

		// ɾ����ͼ vScoreSumFormMain
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vScoreSumFormMain]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vScoreSumFormMain]" );
		pDatabase->ExecuteSQL( strSQL );

		// ɾ����ͼ vScoreFormMain
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vScoreFormMain]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vScoreFormMain]" );
		pDatabase->ExecuteSQL( strSQL );
	}
	catch ( CDBException * e )
	{
		CString	strError;
		strError.Format( "������ɺ�ɾ����ʱ���ݿ����ʧ��\r\n" );
		strError	+= e->m_strError;
		AfxMessageBox( strError );
		GetDlgItem ( IDC_BTN_STEPSCORE )->ShowWindow( SW_SHOW )	;
		return;
	}
}
