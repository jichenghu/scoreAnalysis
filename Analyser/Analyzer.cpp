#include "StdAfx.h"
#include "Analyzer.h"
#include "Analyser.h"

#include "setCutSheetEx.h"
#include "setCutSheetScoreEx.h"
#include "setScoreStudentStep.h"
#include "setQuestionInfo.h"
#include "setQuestionStepInfo.h"
#include "setSelectID.h"

#include "viewAVGRegionStep.h"
#include "viewAVGSchoolStep.h"
#include "viewAVGClassStep.h"

#include "vCnt.h"
#include "vFactStudent.h"
#include "vScoreForm.h"
#include "vChoiceInfo.h"
#include "setChoiceCnt.h"
#include "vChoicePercentage.h"
#include "vAVGClass.h"

#include "msword9.h"

#include "Api.h"
#include <direct.h>

extern	CAnalyserApp theApp	;

void CAnalyzer::SetPathOutput ( CString strPath ) 
{
	m_strPathOutput	= strPath	;
}

BOOL CAnalyzer::BuildClassCheckInfo ( BOOL bReverse )
{
	CvChoiceInfo	setChoiceInfo	;
	CSetChoiceCnt	setChoiceCnt	;
	CvCnt			setCntStep		;

	int	iCntA[128]	;
	int	iCntB[128]	;
	int	iCntC[128]	;
	int	iCntD[128]	;
	int	iClassID	;
	int	iCntClass	;
	int	iCntStep	;
	int	iQuestionID	;
	int	iCntStudent	;

	//CString		strSQL		;
	CDatabase *	pDatabase	;

	pDatabase	= &theApp.m_database;

	iCntClass	= 0	;
	iClassID	= -1;
	iQuestionID	= -1;
	iCntStudent	= 0	;

	setChoiceInfo.m_pDatabase	= pDatabase;
	setChoiceCnt.m_pDatabase	= pDatabase;
	setCntStep.m_pDatabase		= pDatabase;

	//������
	m_pProgress	= new CProgress(NULL);
	m_pProgress->SetTitle1 ( "����ԭʼ�ɼ�����" )	;

	try
	{
		setChoiceInfo.Open ( )	;
		setChoiceCnt.Open ( )	;

		// ������ݿ�� ChoiceCnt ���Ѿ����������ȷ�������
		if ( !setChoiceCnt.IsEOF() )
		{
			int	iDlg	= MessageBox ( NULL, "������ǰ��ѡ����������ݣ�\n�Ƿ�ɾ����ǰ�����ݣ�", "����", MB_YESNOCANCEL )	;
			if ( IDYES == iDlg )
			{
				theApp.m_database.ExecuteSQL ( "DELETE FROM ChoiceCnt" )	;
			}
			else if ( IDNO == iDlg )
			{
				setChoiceInfo.Close ( )	;
				setChoiceCnt.Close ( )	;
					
				m_pProgress->DestroyWindow();

				AfxMessageBox ( "����û��ѡ��ɾ����ǰ�����ݽ��޷�����" )	;
				
				return	FALSE;
			}
			else
			{
				setChoiceInfo.Close ( )	;
				setChoiceCnt.Close ( )	;
					
				m_pProgress->DestroyWindow();

				return	FALSE;
			}
		}

		while ( !setChoiceInfo.IsEOF ( ) )
		{
			// Get the path
			CString	strPath	;
			CString	strMain	;

			if ( (iClassID != setChoiceInfo.m_ClassID) | (iQuestionID != setChoiceInfo.m_QuestionID) )
			{
				if ( -1 != iClassID )
				{
					// ����ð��ͳ�ƽ��
					for ( int i=0; i<iCntStep; i++ )
					{
						setChoiceCnt.AddNew ( )	;
						setChoiceCnt.m_ClassID		= iClassID		;
						setChoiceCnt.m_QuestionID	= iQuestionID	;
						setChoiceCnt.m_Step			= i				;
						setChoiceCnt.m_Choice		= 0				;
						setChoiceCnt.m_Cnt			= iCntA[i]		;
						setChoiceCnt.m_CntStudent	= iCntStudent	;
						setChoiceCnt.Update ( )	;

						setChoiceCnt.AddNew ( )	;
						setChoiceCnt.m_ClassID		= iClassID		;
						setChoiceCnt.m_QuestionID	= iQuestionID	;
						setChoiceCnt.m_Step			= i				;
						setChoiceCnt.m_Choice		= 1				;
						setChoiceCnt.m_Cnt			= iCntB[i]		;
						setChoiceCnt.m_CntStudent	= iCntStudent	;
						setChoiceCnt.Update ( )	;

						setChoiceCnt.AddNew ( )	;
						setChoiceCnt.m_ClassID		= iClassID		;
						setChoiceCnt.m_QuestionID	= iQuestionID	;
						setChoiceCnt.m_Step			= i				;
						setChoiceCnt.m_Choice		= 2				;
						setChoiceCnt.m_Cnt			= iCntC[i]		;
						setChoiceCnt.m_CntStudent	= iCntStudent	;
						setChoiceCnt.Update ( )	;

						setChoiceCnt.AddNew ( )	;
						setChoiceCnt.m_ClassID		= iClassID		;
						setChoiceCnt.m_QuestionID	= iQuestionID	;
						setChoiceCnt.m_Step			= i				;
						setChoiceCnt.m_Choice		= 3				;
						setChoiceCnt.m_Cnt			= iCntD[i]		;
						setChoiceCnt.m_CntStudent	= iCntStudent	;
						setChoiceCnt.Update ( )	;
					}
				}

				CString	strSQL	;
				strSQL.Format ( "SELECT COUNT(*) FROM QuestionStepInfo WHERE QuestionID = %d", setChoiceInfo.m_QuestionID )	;

				setCntStep.Open ( CRecordset::snapshot, strSQL, CRecordset::readOnly )	;
				iCntStep	= setCntStep.m_Cnt	;
				setCntStep.Close();

				// ��ʼ����ͳ����
				iClassID	= setChoiceInfo.m_ClassID	;
				iQuestionID	= setChoiceInfo.m_QuestionID;
				iCntStudent	= 0							;

				::memset ( iCntA, 0, sizeof(iCntA) )	;
				::memset ( iCntB, 0, sizeof(iCntB) )	;
				::memset ( iCntC, 0, sizeof(iCntC) )	;
				::memset ( iCntD, 0, sizeof(iCntD) )	;
			
				m_pProgress->SetPos1 ( iCntClass++ % 101 )	;
			}

			iCntStudent++	;

			// ���ļ�

			strMain.Format ( 
				"D:\\Scan\\67\\" )	;
			strPath.Format ( 
				"%d\\%d\\sel%d\\%d.dat",
				setChoiceInfo.m_SubjectID, 
				setChoiceInfo.m_RoomID,
				setChoiceInfo.m_SheetID, 
				iQuestionID )	;

			CFile	file		;
			BYTE	ySave		;
			int		nFileLength	;

			ySave	= 0	;
			file.Open ( strMain + strPath, CFile::modeRead );
			nFileLength	= (int)( file.GetLength( ) );

			for ( int j=0; j<nFileLength,j<iCntStep; j++ )
			{
				file.Read ( &ySave, 1 )	;

				if ( bReverse )
				{
					iCntD [j]	+= ( ySave & 0x01 )		;
					iCntC [j]	+= ( ySave & 0x02 ) >> 1;
					iCntB [j]	+= ( ySave & 0x04 ) >> 2;
					iCntA [j]	+= ( ySave & 0x08 ) >> 3;
				}
				else
				{
					iCntA [j]	+= ( ySave & 0x01 )		;
					iCntB [j]	+= ( ySave & 0x02 ) >> 1;
					iCntC [j]	+= ( ySave & 0x04 ) >> 2;
					iCntD [j]	+= ( ySave & 0x08 ) >> 3;
				}
			}
			file.Close ( )	;

			setChoiceInfo.MoveNext ( )	;
		}

		// �������һ�����ͳ�ƽ��
		for ( int i=0; i<iCntStep; i++ )
		{
			setChoiceCnt.AddNew ( )	;
			setChoiceCnt.m_ClassID		= iClassID		;
			setChoiceCnt.m_QuestionID	= iQuestionID	;
			setChoiceCnt.m_Step			= i				;
			setChoiceCnt.m_Choice		= 0				;
			setChoiceCnt.m_Cnt			= iCntA[i]		;
			setChoiceCnt.m_CntStudent	= iCntStudent	;
			setChoiceCnt.Update ( )	;

			setChoiceCnt.AddNew ( )	;
			setChoiceCnt.m_ClassID		= iClassID		;
			setChoiceCnt.m_QuestionID	= iQuestionID	;
			setChoiceCnt.m_Step			= i				;
			setChoiceCnt.m_Choice		= 1				;
			setChoiceCnt.m_Cnt			= iCntB[i]		;
			setChoiceCnt.m_CntStudent	= iCntStudent	;
			setChoiceCnt.Update ( )	;

			setChoiceCnt.AddNew ( )	;
			setChoiceCnt.m_ClassID		= iClassID		;
			setChoiceCnt.m_QuestionID	= iQuestionID	;
			setChoiceCnt.m_Step			= i				;
			setChoiceCnt.m_Choice		= 2				;
			setChoiceCnt.m_Cnt			= iCntC[i]		;
			setChoiceCnt.m_CntStudent	= iCntStudent	;
			setChoiceCnt.Update ( )	;

			setChoiceCnt.AddNew ( )	;
			setChoiceCnt.m_ClassID		= iClassID		;
			setChoiceCnt.m_QuestionID	= iQuestionID	;
			setChoiceCnt.m_Step			= i				;
			setChoiceCnt.m_Choice		= 3				;
			setChoiceCnt.m_Cnt			= iCntD[i]		;
			setChoiceCnt.m_CntStudent	= iCntStudent	;
			setChoiceCnt.Update ( )	;
		}

		setChoiceCnt.Close ( )	;
		setChoiceInfo.Close ( )	;
	}
	catch ( CDBException *e )
	{
//		afxDump << e->m_strError << "\n"	;

		AfxMessageBox ( e->m_strError )	;

		return	FALSE;
	}

	// ���ٽ�����
	m_pProgress->DestroyWindow();

	return	TRUE;
}


void CAnalyzer::DestroyMarkList ( LPMARK_STEP lpMarkList )
{
	LPMARK_STEP			lpMarkStep	;
	LPMARK_STEP			lpSubHead	;

	lpSubHead	= lpMarkList	;
	while ( NULL != lpSubHead )
	{
		lpMarkList	= lpSubHead->lpNextCutSheet	;

		// ɾ���Ӵ�
		lpMarkStep	= lpSubHead	;
		while ( NULL != lpMarkStep )
		{
			lpSubHead	= lpSubHead->lpNextMark	;

			delete	lpMarkStep	;

			lpMarkStep	= lpSubHead	;
		}

		lpSubHead	= lpMarkList	;
	}
}

CAnalyzer::CAnalyzer(void)
{
	m_lpQuestionShift	= NULL	;

	m_lstMarkStepHead	= NULL	;
	m_lstMarkStepTail	= NULL	;

	m_lstMarkHead	= NULL	;
	m_lstMarkTail	= NULL	;
}

CAnalyzer::~CAnalyzer(void)
{
	LPFIELD_SHIFT	lpShift	;

	lpShift	= m_lpQuestionShift	;
	while ( NULL != lpShift )
	{
		m_lpQuestionShift	= lpShift->lpNext	;

		delete	lpShift	;

		lpShift	= m_lpQuestionShift	;
	}

	DestroyMarkList ( m_lstMarkHead )	;
}

void CAnalyzer::AnalyzeSingleSubject ( int iSubjectID, int iRegionID, int iTypeID,
									  CString strSubject, CString strRegion, CString strType,
									  LPSUBITEM lpSubItem, CString strPathOutput, BYTE yLayers )
{
	// ���� lpSubItem �������ʽ��Ϣ
	COleVariant		vTrue( (short)TRUE )						;
	COleVariant		vFalse( (short)FALSE )						;
	COleVariant		vOpt( (long)DISP_E_PARAMNOTFOUND, VT_ERROR );

	_Application	m_App	;	// ����Word�ṩ��Ӧ�ó������
	Documents		m_Docs	;	// ����Word�ṩ���ĵ�����
	Selection		m_Sel	;	// ����Word�ṩ��ѡ�����

	m_Docs.ReleaseDispatch();
	m_Sel.ReleaseDispatch()	;

	m_App.m_bAutoRelease	= true;
	if( !m_App.CreateDispatch("Word.Application") )
	{ 
		AfxMessageBox( "����Word2000����ʧ��!" ); 
		exit( 1 ); 
	}

	CString	strTemplate	;

	strTemplate.Format ( lpSubItem->cPath )	;
	//strTemplate.Format( "c:\\����\\֪ʶ��" )	;
	//strType.TrimRight()						;
	//strTemplate	+= strType					;
	//strTemplate	+= strSubject.Left( 4 )		;

	// �����Ƕ���VARIANT����
	COleVariant	varFilePath( strTemplate )		;
	COleVariant	varstrNull( "" )				;
	COleVariant varZero( (short)0 )				;
	COleVariant varTrue( short(1), VT_BOOL )	;
	COleVariant varFalse( short(0), VT_BOOL )	;

	m_Docs.AttachDispatch( m_App.GetDocuments() )	;	// ��Documents�����m_Docs��Idispatch�ӿڹ�������
	m_Docs.Open( varFilePath, 
		varFalse, varFalse, varFalse,
		varstrNull, varstrNull, varFalse, varstrNull,
		varstrNull, varTrue, varTrue, varTrue )		;
	//m_App.SetVisible ( FALSE );
	m_App.SetVisible ( TRUE );
	
	// ��Word�ĵ�
	m_Sel.AttachDispatch( m_App.GetSelection() )	;	// ��Selection�����m_Sel��Idispatch�ӿڹ�������

	//// ȫѡ����
	//m_Sel.WholeStory( )	;
	//m_Sel.Copy ( )		;

	// ��д���
	// д����
	//m_Sel.TypeText( strSchool );
	//m_Sel.MoveDown(COleVariant((short)4),COleVariant((short)1),COleVariant((short)0));
	//m_Sel.MoveDown(COleVariant((short)5),COleVariant((short)1),COleVariant((short)0));
	//m_Sel.MoveRight(COleVariant((short)1),COleVariant((short)1),COleVariant((short)0));

	// save word file
	_Document	oActiveDoc; 
	oActiveDoc = m_App.GetActiveDocument(); 

	//CString	strOutput	;
	//strOutput	= strPathOutput	;
	//strOutput	+= "\\"			;
	//strOutput	+= strSubject	;
	////strOutput	+= strSchool	;
	//strOutput	+= ".doc"		;
	////oActiveDoc.SaveAs( COleVariant( strOutput ), 
	////	COleVariant((short)0), 
	////	vFalse, COleVariant(""), vTrue, COleVariant(""), 
	////	vFalse, vFalse, vFalse, vFalse, vFalse );


	//m_Docs.Add ( COleVariant("Normal"), varFalse, COleVariant((short)0), varTrue )	;
	//oActiveDoc = m_App.GetActiveDocument(); 
	//oActiveDoc.SaveAs( COleVariant( strOutput ), 
	//	COleVariant((short)0), 
	//	vFalse, COleVariant(""), vTrue, COleVariant(""), 
	//	vFalse, vFalse, vFalse, vFalse, vFalse );

	CWnd *	pWnd;

	CString	strWndTitle	;

	strWndTitle	= strTemplate.Right ( strTemplate.GetLength () - strTemplate.ReverseFind ( '\\' ) - 1 )	;

	pWnd	= CWnd::FromHandle ( ::FindWindow ( NULL, strWndTitle + " - Microsoft Word" ) )	;
//	pWnd->SetForegroundWindow ( );

//	return;

	switch ( yLayers )
	{
	case	0:	// ��
		break	;

	case	1:	// ��У
		break	;

	case	2:	// У
		lpSubItem->m_pPara->strNameRegion	= &strRegion;
		//FillFormSchool ( iSubjectID, iRegionID, &m_App, &m_Docs, &oActiveDoc, &m_Sel, lpSubItem )	;
		break	;

	case	3:	// У��
		lpSubItem->m_pPara->strNameRegion	= &strRegion;
		FillFormSchoolClass ( iSubjectID, strSubject, iRegionID, strRegion, iTypeID, strType, &m_App, &m_Docs, &oActiveDoc, &m_Sel, lpSubItem )	;
		break	;

	case	4:	// ��
		lpSubItem->m_pPara->strNameRegion	= &strRegion;
		FillFormClass ( iSubjectID, strSubject, iRegionID, strRegion, iTypeID, strType, &m_App, &m_Docs, &oActiveDoc, &m_Sel, lpSubItem )	;
		break	;

	case	5:	// ��У��
		lpSubItem->m_pPara->strNameRegion	= &strRegion;
		FillFormRegionSchoolClass ( iSubjectID, strSubject, iRegionID, strRegion, iTypeID, strType, &m_App, &m_Docs, &oActiveDoc, &m_Sel, lpSubItem )	;
		break	;

	default:
		break	;
	}

	// FillSchoolStepForm( &m_Sel, iSchoolID, iSubjectID, iTypeID );

	//m_Docs.Add ( COleVariant("Normal"), varFalse, COleVariant((short)0), varTrue )	;
	//oActiveDoc = m_App.GetActiveDocument(); 
	//oActiveDoc.SaveAs( COleVariant( strOutput+"a" ), 
	//	COleVariant((short)0), 
	//	vFalse, COleVariant(""), vTrue, COleVariant(""), 
	//	vFalse, vFalse, vFalse, vFalse, vFalse );
	
	m_Docs.ReleaseDispatch();	// �Ͽ�������
	m_Sel.ReleaseDispatch()	;

	// �˳�WORD 
	m_App.Quit(vOpt, vOpt, vOpt); 
//	m_App.Quit(vOpt, vOpt, vOpt);
	m_App.ReleaseDispatch();
}

//++++++++++++++++++++++++++++++++++++++++++++++++++
//	��У�༶������������
//		ÿ��ѧУÿ��һ���ļ�
//
//	��һ��: ��ѧУ����ѭ������Word�ĵ�, ÿ��ѧУһ���ļ�
//	�ڶ���: �� VirtualForm �е���������ģ���
//	������: ���ڶ����еı�����Ŀ����β����
//--------------------------------------------------
void CAnalyzer::FillFormSchoolClass ( int iSubjectID, CString strSubject,  
			int iRegionID, CString strRegion, int iTypeID, CString strType,
			_Application* pApp, Documents* pDocs, _Document* pParent, Selection* pSel, 
			LPSUBITEM lpSubItem )
{
	// ���ݿ���ز���
	CString	strFilter	;
	CString	strFilter0	;
	CString	strFilter1	;
	CString	strFilter2	;
	CString	strSort		;
	CString	strSchool	;
	CString	strSQL		;

	// ���ݿ�
	//CSetSelectID		setIDSchool				;
	CvChoicePercentage	*lpvChoicePercentage	;
	CvChoicePercentage	*lpvChoicePercentage1	;
	CvChoicePercentage	*lpvChoicePercentage2	;

	// ������ز���
	CString	strPath		;	// ����Ĵ��·��

	// ��ŷ������ݵ� Buffer
	LPREGION_FORM	lpRegion	;	// ��
	LPSCHOOL_FORM	lpSchoolHead;
	LPSCHOOL_FORM	lpSchoolTail;
	LPSCHOOL_FORM	lpSchool	;
	LPSCORE_FORM	lpScoreHead	;
	LPSCORE_FORM	lpScoreTail	;
	LPSCORE_FORM	lpScore		;

	// Word ����ز���
	Selection	m_Sel							;	// ����Word�ṩ��ѡ�����
	COleVariant varTrue( short(1), VT_BOOL )	;
	COleVariant varFalse( short(0), VT_BOOL )	;
	COleVariant	vTrue( (short)TRUE )			;
	COleVariant	vFalse( (short)FALSE )			;

	// ���ñ���
	CString			strID		;
	CString			strName		;
	CString			strScore	;
	CString			strBlock	;

	// ��ʼ������
	m_pProgress	= new CProgress(NULL)	;	// ������
	lpRegion	= new REGION_FORM		;	// ÿ��ֻ����һ����

	strPath	= m_strPathOutput		;
	SetCurrentDirectory ( strPath )	;

	strPath	= strPath + "\\"		;
	strPath	= strPath + strRegion	;
	if ( !TestDirectory( strPath ) )
		_mkdir ( strRegion )	;
	SetCurrentDirectory ( strPath )	;
	strPath	= strPath + "\\"		;
	
	//setIDSchool.m_pDatabase	= &theApp.m_database;


	lpvChoicePercentage	= new CvChoicePercentage;
	lpvChoicePercentage->m_pDatabase	= &theApp.m_database;

	lpvChoicePercentage1	= NULL	;
	lpvChoicePercentage2	= NULL	;
	if ( '0' != lpSubItem->cSQL1[0] )
	{
		lpvChoicePercentage1	= new CvChoicePercentage	;
		lpvChoicePercentage1->m_pDatabase	= &theApp.m_database;

		if ( '0' != lpSubItem->cSQL2[0] )
		{
			lpvChoicePercentage2	= new CvChoicePercentage	;
			lpvChoicePercentage2->m_pDatabase	= &theApp.m_database;
		}
	}

	
	// ���Ȳ�����ص���ͼ
	if ( !GenerateChoice() )
	{
		AfxMessageBox ( "������ص���ͼʧ�ܣ�\n�������δ�ܲ���" )	;
		return	;
	}

//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//	�ռ���Ϣ
	strFilter0.Empty ( )	;
	strFilter1.Empty ( )	;
	strFilter2.Empty ( )	;
	m_pProgress->SetTitle2 ( "�ռ���Ϣ" )	;
	try
	{
		//+++++++++++++++++++++++++++
		// �ȶ�����������
		strSQL.Format( "%s", lpSubItem->cSQL )	; strSQL.TrimLeft(); strSQL.TrimRight();
		if ( '0' != lpSubItem->cSort[0] )
			lpvChoicePercentage->m_strSort	= lpSubItem->cSort	;
		if ( '0' != lpSubItem->cFilter[0] )
		{
			//lpvChoicePercentage->m_strFilter	= lpSubItem->cFilter;
			strFilter0.Format ( "%s", lpSubItem->cFilter )	;
		}
		lpvChoicePercentage->m_strFilter.Format ( "(SubjectID = %d) AND (TypeID = %d)", iSubjectID, iTypeID );/**/
		lpvChoicePercentage->m_strFilter	+= strFilter0;
		lpvChoicePercentage->Open ( CRecordset::snapshot, strSQL, CRecordset::readOnly )	;

		// ����ɼ���������
		lpScoreHead	= NULL	;
		lpScoreTail	= NULL	;
		while ( !lpvChoicePercentage->IsEOF() )
		{
			lpScore	= new SCORE_FORM	;

			lpScore->iScoreTypeID	= lpvChoicePercentage->m_QuestionFormID	;
			lpScore->fValue			= lpvChoicePercentage->m_percentage		;
			lpScore->pNextScore		= NULL									;

			if ( NULL == lpScoreHead )
			{
				lpScoreHead	= lpScore	;
				lpScoreTail	= lpScore	;

				strcpy ( lpRegion->cName, lpvChoicePercentage->m_Name )	;
			}
			else
			{
				lpScoreTail->pNextScore	= lpScore	;
				lpScoreTail				= lpScore	;
			}

			lpvChoicePercentage->MoveNext ( )	;
		}
		lpRegion->lpScoreLst	= lpScoreHead	;

		// �����ݶ������, ������Դ
		lpvChoicePercentage->Close ( )	;
		delete	lpvChoicePercentage		;

		//+++++++++++++++++++++++++++
		// �ٶ������и�ѧУ������
		if ( lpvChoicePercentage1 )
		{
			strSQL.Format( "%s", lpSubItem->cSQL1 )	; strSQL.TrimLeft(); strSQL.TrimRight();
			if ( '0' != lpSubItem->cSort1[0] )
				lpvChoicePercentage1->m_strSort	= lpSubItem->cSort1		;
			if ( '0' != lpSubItem->cFilter1[0] )
			{
				//lpvChoicePercentage1->m_strFilter= lpSubItem->cFilter1	;
				strFilter1.Format ( "%s", lpSubItem->cFilter1 )	;
			}
			lpvChoicePercentage1->m_strFilter.Format ( "(SubjectID = %d) AND (TypeID = %d)", iSubjectID, iTypeID );/**/
			lpvChoicePercentage1->m_strFilter	+= strFilter1;
			lpvChoicePercentage1->Open ( CRecordset::snapshot, strSQL, CRecordset::readOnly )	;

			// ��������
			//lpScoreHead	= NULL	;
			//lpScoreTail	= NULL	;
			int		iID		= -1	;
			int		iCnt	= 0		;
			lpSchoolHead	= NULL	;
			while ( !lpvChoicePercentage1->IsEOF() )
			{
				lpScore	= new SCORE_FORM	;

				lpScore->iScoreTypeID	= lpvChoicePercentage1->m_QuestionFormID;
				lpScore->fValue			= lpvChoicePercentage1->m_percentage	;
				lpScore->pNextScore		= NULL									;

				// ������µ�ѧУ
				if ( iID != lpvChoicePercentage1->m_ID )
				{
					iCnt++	;	// ѧУ��Ŀ��1
					iID	= lpvChoicePercentage1->m_ID;	// �� iID

					// �½�һ��ѧУ
					lpSchool	= new SCHOOL_FORM	;

					lpSchool->iSchoolID		= iID		;
					lpSchool->lpScoreLst	= lpScore	;
					lpSchool->pNextSchool	= NULL		;
					strcpy ( lpSchool->cName, lpvChoicePercentage1->m_Name )	;

					lpScoreHead			= lpScore	;
					lpScoreTail			= lpScore	;

					if ( NULL == lpSchoolHead )
					{
						lpSchoolHead	= lpSchool	;
						lpSchoolTail	= lpSchool	;
					}
					else
					{
						lpSchoolTail->pNextSchool	= lpSchool	;
						lpSchoolTail				= lpSchool	;
					}
				}
				else
				{
					lpScoreTail->pNextScore	= lpScore	;
					lpScoreTail				= lpScore	;
				}

				lpvChoicePercentage1->MoveNext ( )	;
			}

			// ѧУ���ݶ������, ������Դ
			lpvChoicePercentage1->Close ( )	;
			delete	lpvChoicePercentage1	;

			// ��ѧУ�ҽӵ�����
			lpRegion->lpSchoolLst	= lpSchoolHead	;
			lpRegion->iCntSchool	= iCnt			;
		}

		//strSQL.Format( "SELECT schoolID, schoolName FROM School WHERE (RegionID = %d) ORDER BY schoolID", iRegionID )	;
		//setIDSchool.Open ( CRecordset::snapshot, strSQL, CRecordset::readOnly )	;

		//+++++++++++++++++++++++++++
		// �ζ������и�У�и���ķ�������
		if ( lpvChoicePercentage2 )
		{
			strSQL.Format( "%s", lpSubItem->cSQL2 )	; strSQL.TrimLeft(); strSQL.TrimRight();
			if ( '0' != lpSubItem->cSort2[0] )
				lpvChoicePercentage2->m_strSort	= lpSubItem->cSort2		;
			if ( '0' != lpSubItem->cFilter2[0] )
			{
				//lpvChoicePercentage2->m_strFilter= lpSubItem->cFilter2	;
				strFilter2.Format ( "%s", lpSubItem->cFilter2 )	;
			}
			lpvChoicePercentage2->m_strFilter.Format ( "(SubjectID = %d) AND (TypeID = %d)", iSubjectID, iTypeID );/**/
			lpvChoicePercentage2->m_strFilter	+= strFilter2;
			lpvChoicePercentage2->Open ( CRecordset::snapshot, strSQL, CRecordset::readOnly )	;

			//int		iCntSchool	;
			//int		iCntStudent	;
			int		iSchoolID	;

			//int				iMaxRow		;
			//int				iCntRow		;
			//int				iCntBlock	;

			int		i	= 0				;	// ����
			lpSchool	= lpSchoolHead	;
			while ( NULL != lpSchool )
			{
				// ȡ�ø�ѧУ�������
				iSchoolID	= lpSchool->iSchoolID			;
				strSchool.Format ( "%s", lpSchool->cName )	;
				strSchool.TrimLeft ()	;	strSchool.TrimRight ()	;

				// ���ý��ȱ�־
				m_pProgress->SetTitle2 ( strSchool );

				// ��ȡ�µ����ݿ��¼��
				strFilter.Format ( "(SubjectID = %d) AND (schoolID = %d) AND (TypeID = %d)", iSubjectID, iSchoolID, iTypeID )	;/**/
				lpvChoicePercentage2->m_strFilter	= strFilter2 + strFilter;
				lpvChoicePercentage2->Requery ( )	;

				// �����ݿ��¼��
				LPCLASS_FORM	lpClassHead	;
				LPCLASS_FORM	lpClassTail	;
				LPCLASS_FORM	lpClass		;

				int		iCnt	= 0		;
				int		iID		= -1	;
				lpClassHead		= NULL	;
				while ( !lpvChoicePercentage2->IsEOF() )
				{
					lpScore	= new SCORE_FORM	;

					lpScore->iScoreTypeID	= lpvChoicePercentage2->m_QuestionFormID;
					lpScore->fValue			= lpvChoicePercentage2->m_percentage	;
					lpScore->pNextScore		= NULL									;
				
					if ( iID != lpvChoicePercentage2->m_ID )
					{
						iCnt++	;	// �༶����1
						iID	= lpvChoicePercentage2->m_ID	;

						lpClass	= new CLASS_FORM	;

						lpClass->iClassID	= iID		;
						lpClass->lpScoreLst	= lpScore	;
						lpClass->pNextClass	= NULL		;
						strcpy ( lpClass->cName, lpvChoicePercentage2->m_Name )	;

						lpScoreHead			= lpScore	;
						lpScoreTail			= lpScore	;

						if ( NULL == lpClassHead )
						{
							lpClassHead	= lpClass	;
							lpClassTail	= lpClass	;
						}
						else
						{
							lpClassTail->pNextClass	= lpClass	;
							lpClassTail				= lpClass	;
						}
					}
					else
					{
						lpScoreTail->pNextScore	= lpScore	;
						lpScoreTail				= lpScore	;
					}

					// �ƶ�����һ��¼
					lpvChoicePercentage2->MoveNext ( )	;
				}

				// ���༶�б����ӵ�ѧУ
				lpSchool->lpClassLst= lpClassHead	;
				lpSchool->iCntClass	= iCnt			;

				// ������һѧУ
				lpSchool	= lpSchool->pNextSchool	;

				// ���ý���
				m_pProgress->SetPos2 ( ++i * 100 / lpRegion->iCntSchool )	;
			}

			// ��У�ĸ�������ݶ������, ������Դ
			lpvChoicePercentage2->Close ( )	;
			delete	lpvChoicePercentage2	;
		}

/*		lpSchoolHead	= NULL	;

		while ( !setIDSchool.IsEOF() )
		{
			setIDSchool.MoveNext ( )	;
		}
		iCntSchool	= setIDSchool.GetRecordCount ( )	;

		// ��ѧУѭ��
		if ( iCntSchool )
			setIDSchool.MoveFirst ( )	;
		for ( int i=0; i<iCntSchool; i++ )	// for school
		{
			iSchoolID	= setIDSchool.m_iID	;
			strSchool.Format ( "%s", setIDSchool.m_strName )	;
			strSchool.TrimRight ( )		;

			m_pProgress->SetTitle1 ( strSchool )	;

			// �����ݿ�
			strFilter.Format ( "(SubjectID = %d) AND (schoolID = %d)", iSubjectID, setIDSchool.m_iID )	;
			lpvChoicePercentage1->m_strFilter	= strFilter1 + strFilter;
			lpvChoicePercentage1->Requery ( )	;

			LPCLASS_FORM	lpClassHead	;
			LPCLASS_FORM	lpClassTail	;
			LPCLASS_FORM	lpClass		;

			int		iCnt	= 0		;
			int		iID		= -1	;
			lpClassHead	= NULL	;
			while ( !lpvChoicePercentage1->IsEOF() )
			{
				lpScore	= new SCORE_FORM	;

				lpScore->iScoreTypeID	= lpvChoicePercentage1->m_QuestionFormID;
				lpScore->fValue			= lpvChoicePercentage1->m_percentage	;
				lpScore->pNextScore		= NULL									;

				if ( iID != lpvChoicePercentage1->m_ID )
				{
					iCnt++	;	// �༶����1
					iID	= lpvChoicePercentage1->m_ID	;

					lpScoreHead			= lpScore	;
					lpScoreTail			= lpScore	;

					lpClass	= new CLASS_FORM	;

					lpClass->iClassID	= lpvChoicePercentage1->m_ID;
					strcpy ( lpClass->cName, lpvChoicePercentage1->m_Name )	;
					lpClass->lpScoreLst	= lpScoreHead				;
					lpClass->pNextClass	= NULL						;

					if ( NULL == lpClassHead )
					{
						lpClassHead	= lpClass	;
						lpClassTail	= lpClass	;
					}
					else
					{
						lpClassTail->pNextClass	= lpClass	;
						lpClassTail				= lpClass	;
					}
				}
				else
				{
					lpScoreTail->pNextScore	= lpScore	;
					lpScoreTail				= lpScore	;
				}

				lpvChoicePercentage1->MoveNext ( )	;
			}

			// �����ѧУ�гɼ���������, �򽫱�ѧУ���뵽ѧУ�б���
			if ( NULL != lpClassHead )
			{
				lpSchool	= new SCHOOL_FORM	;

				lpSchool->iSchoolID		= iSchoolID					;
				strcpy ( lpSchool->cName, setIDSchool.m_strName )	;
				lpSchool->lpClassLst	= lpClassHead				;
				lpSchool->iCntClass		= iCnt						;
				lpSchool->pNextSchool	= NULL						;

				if ( NULL == lpSchoolHead )
				{
					lpSchoolHead	= lpSchool	;
					lpSchoolTail	= lpSchool	;
				}
				else
				{
					lpSchoolTail->pNextSchool	= lpSchool	;
					lpSchoolTail				= lpSchool	;
				}
			}

			//m_pProgress->SetPos1( 100 * iClass / iCntClass );
			//m_pProgress->Step1 ( )	;

			m_pProgress->SetPos1( 100 * i / iCntSchool );

			if ( i+1 < iCntSchool )
				setIDSchool.MoveNext ( )	;
		}// end of for-sChool

		setIDSchool.Close ( );

		if ( NULL != lpvChoicePercentage1 )
		{
			lpvChoicePercentage1->Close ( )	;
			delete	lpvChoicePercentage1		;

			if ( NULL != lpvChoicePercentage2 )
			{
				lpvChoicePercentage2->Close ( )	;
				delete	lpvChoicePercentage2		;
			}
		}
*/	}
	catch ( CDBException * e )
	{
		AfxMessageBox ( e->m_strError )	;
		return;
	}
	
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//	���ɱ���
	m_pProgress->SetTitle1 ( "���ɱ���" )	;
	m_pProgress->SetPos1( 0 );
	m_pProgress->SetPos2( 0 );

	LPCLASS_FORM	lpClass		;
	LPTSTR			lptstrCopy	; 
	HGLOBAL			hglbCopy	; 

	// ÿ����ֻ����һ�� Word �ĵ�
	_Document	oActiveDoc	;
	pDocs->Add ( COleVariant(pParent->GetFullName()), varFalse, COleVariant((short)0), varTrue )	;
	oActiveDoc = pApp->GetActiveDocument(); 
	m_Sel.AttachDispatch( pApp->GetSelection() )	;	// ��Selection�����m_Sel��Idispatch�ӿڹ�������

	// ��ѧУѭ��
	CString	strHeadLineClass	;
	CString	strHeadLineSchool	;
	int		iMaxRow				;
	int		iCntCol				;

	int	ii		= 0				;

	iMaxRow		= 0;
	lpSchool	= lpSchoolHead	;
	while ( NULL != lpSchool )
	{
		int	iCntClass	;
		int	iReadClass	;	// ��ǰ��Ҫ����İ༶��
		//int	iCntRow		;
		int	iCntBlock	;
		int	iSchoolID	;

		//iCntRow		= 0	;
		iCntCol		= 0	;
		iCntBlock	= 0	;

		iMaxRow		= lpSubItem->m_iMaxRowOfBlock[0];

		iCntClass	= lpSchool->iCntClass	;
		iSchoolID	= lpSchool->iSchoolID	;
		strSchool.Format ( "%s", lpSchool->cName )	;
		strSchool.TrimRight ( )		;
			
		// ���ò���
		lpSubItem->m_pPara->strNameSchool	= &strSchool			;
		//lpSubItem->m_pPara->iCntRow			= lpSchool->iCntClass	;
		//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

		m_pProgress->SetTitle2 ( strSchool );
		m_pProgress->SetPos1 ( 0 )			;

		// �����һ�ζ���İ༶����
		iReadClass	= iCntClass % (iMaxRow - 1 );	// ��һ����ȫУ��ͳ������
		if ( 0 == iReadClass )
			iReadClass	= iMaxRow - 1	;	// ������������1��Ҳ���������

		// ������������
		lpSubItem->ResetRow ( )			;
		lpClass	= lpSchool->lpClassLst	;
		while ( NULL != lpClass )
		{
			// һ�ζ���������ļ�¼����
			LPCLASS_FORM	*lppClass		;
			LPSCORE_FORM	*lppScore		;	// ��Ÿ��������
			LPSCORE_FORM	lpScoreSchool	;	// ���ѧУ������

			lppClass	= new LPCLASS_FORM[iReadClass]	;
			lppScore	= new LPSCORE_FORM[iReadClass]	;
		
			strHeadLineClass.Empty ( )		;

			// У����
			lpScoreSchool	= lpSchool->lpScoreLst	;

			strHeadLineClass	+= lpSchool->cName	;

			// ������ 
			int		i	;
			for ( i=0; (i<iReadClass) && (NULL!=lpClass); i++ )
			{
				*(lppClass+i)	= lpClass	;
				if ( NULL == lpClass )
					*(lppScore+i)	= NULL	;
				else
					*(lppScore+i)	= lpClass->lpScoreLst	;

				strHeadLineClass	+= "\t"	;
				strHeadLineClass	+= lpClass->cName	;

				lpClass	= lpClass->pNextClass	;
			}
			for ( i=iReadClass; i<iMaxRow-1; i++ )
			{
				strHeadLineClass	+= "\t "	;
			}
			strHeadLineClass	+= " ";

			// �γ����� blocks
			int		iRow	;
			for ( int iBlock=0; iBlock<lpSubItem->m_iCntBlocks; iBlock++ )
			{
				lpSubItem->m_pPara->iCntRow		= lpSubItem->m_iMaxColOfBlock[iBlock]	;

				strBlock.Empty ( )	;

				// ������ɱ� block
				iCntCol	= 0	;
				while ( NULL != lpScoreSchool )
				{
					CWordFormatPrint *	pFormat	= (CWordFormatPrint *)lpSubItem->m_listlpColumn	;

					// �ȶ���ȫУ������
					lpScore	= lpScoreSchool	;
					for ( iRow=0; iRow<lpSubItem->m_iMaxSubrowOfRow[0]; iRow++ )
					{
						strScore.Format ( pFormat->m_cFormat, 100*lpScore->fValue )	;

						if ( 0 == iRow )
							strBlock	+= strScore + "%"	;	
						else
							strBlock	+= "\t" + strScore + "%"	;

						pFormat		= (CWordFormatPrint *)pFormat->m_pNext		;
						lpScore		= lpScore->pNextScore						;

					}
					lpScoreSchool	= lpScore	;

					// ���뱾 block �и��༶������
					int	j;
					for ( j=0; j<iReadClass; j++ )
					{
						lpScore	= *(lppScore+j)	;

						// ����һ������
						for ( iRow=0; iRow<lpSubItem->m_iMaxSubrowOfRow[j+1]; iRow++ )
						{
							strScore.Format ( pFormat->m_cFormat, 100*lpScore->fValue )	;

							strBlock	+= "\t" + strScore + "%"	;	

							pFormat		= (CWordFormatPrint *)pFormat->m_pNext		;
							lpScore		= lpScore->pNextScore						;
						}

						*(lppScore+j)	= lpScore	;
					}
					// ����
					for ( j=iReadClass; j<iMaxRow-1; j++ )
					{
						for ( iRow=0; iRow<lpSubItem->m_iMaxSubrowOfRow[j+1]; iRow++ )
						{
							strBlock	+= "\t "	;	
						}
					}

					strBlock	+= "\r\n"	;

					iCntCol++	;	// ����һ��

					if ( iCntCol == lpSubItem->m_iMaxColOfBlock[iBlock] )
						break;

				}

				// ���� Title
				lpSubItem->MakeTitle ( &m_Sel )	;


				//+++++++++++++++++++++++++++++++++++++
				// copy and paste Headline
				// ��ͷ
				if ( !OpenClipboard(NULL) ) 
					AfxMessageBox ( "Failed open clipboard" ); 

				EmptyClipboard(); 

				hglbCopy	= GlobalAlloc(GMEM_MOVEABLE, strHeadLineClass.GetLength( ) ); 
				if (hglbCopy == NULL) 
				{ 
					CloseClipboard(); 
					AfxMessageBox ( "Failed alloc mem" ); 
					return ; 
				} 

				lptstrCopy = (LPTSTR)GlobalLock(hglbCopy); 
				memcpy( lptstrCopy, strHeadLineClass, strHeadLineClass.GetLength() ); 
				//lptstrCopy[cch] = (TCHAR) 0;    // null character 
				GlobalUnlock(hglbCopy); 

				SetClipboardData(CF_TEXT, hglbCopy); 

				CloseClipboard(); 

				lpSubItem->SelectHead ( &m_Sel )	;
				m_Sel.Paste ( )	;

				//+++++++++++++++++++++++++++++++++++++
				// copy and paste this block
				if ( !OpenClipboard(NULL) ) 
					AfxMessageBox ( "Failed open clipboard" ); 

				EmptyClipboard(); 

				hglbCopy	= GlobalAlloc(GMEM_MOVEABLE, strBlock.GetLength( ) ); 
				if (hglbCopy == NULL) 
				{ 
					CloseClipboard(); 
					AfxMessageBox ( "Failed alloc mem" ); 
					return ; 
				} 

				lptstrCopy = (LPTSTR)GlobalLock(hglbCopy); 
				memcpy( lptstrCopy, strBlock, strBlock.GetLength() ); 
				//lptstrCopy[cch] = (TCHAR) 0;    // null character 
				GlobalUnlock(hglbCopy); 

				SetClipboardData(CF_TEXT, hglbCopy); 

				CloseClipboard(); 

				// ��λ����������һ��
				lpSubItem->LocateBody ( &m_Sel );
				//lpSubItem->ResetCell ( )	;
				lpSubItem->ShiftRow ( &m_Sel )	;
				m_Sel.Paste ( )	;

				//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
				//lpSubItem->m_pPara->iCntRow	-= iMaxRow	;

				//iCntBlock++	;
				//iMaxRow		= lpSubItem->m_iMaxRowOfBlock[iCntBlock];
				//					AfxMessageBox ( *pstrBlock )	;

				m_Sel.HomeKey ( COleVariant((short)6), COleVariant((short)0) );
				pParent->Activate ( )	;
				m_Sel.AttachDispatch( pApp->GetSelection() )	;
				m_Sel.WholeStory( )		;
				m_Sel.Copy ( )			;

				oActiveDoc.Activate ( )	;
				m_Sel.AttachDispatch( pApp->GetSelection() )	;

				// ����ĸ��
				m_Sel.Paste ( )	;
				m_Sel.HomeKey ( COleVariant((short)6), COleVariant((short)0) );

			}	// end of for iBlock

			iCntClass	-= iReadClass	;
			iReadClass	= iMaxRow - 1	;

			delete	lppClass;
			delete	lppScore;

			m_pProgress->Step1 ( )	;

		}	// end of while lpClass

		lpSchool	= lpSchool->pNextSchool	;

		// �ص����
		m_Sel.HomeKey ( COleVariant((short)6), COleVariant((short)0) );

		m_pProgress->SetPos2( 100 * ii++ / lpRegion->iCntSchool );
	
	}	// end of while lpSchool

	if ( 0 == iMaxRow )
	{
		AfxMessageBox ( "δ������Ч����!\r\n�������ݿ��Ƿ���ȷ����:\r\n    QuestionFormInfo" );

		// ���ٻ�����
		LPSCHOOL_FORM	lpSchoolPrev;
		LPSCORE_FORM	lpScorePrev	;
		lpSchool	= lpRegion->lpSchoolLst;
		while ( NULL != lpSchool )
		{
			// ���ٰ༶�б�
			LPCLASS_FORM	lpClass		;
			LPCLASS_FORM	lpClassPrev	;
			lpClass	= lpSchool->lpClassLst	;
			while ( NULL != lpClass )
			{
				// ���ٳɼ��б�
				lpScore	= lpClass->lpScoreLst	;
				while ( NULL != lpScore )
				{
					lpScorePrev	= lpScore				;
					lpScore		= lpScore->pNextScore	;
					delete	lpScorePrev	;
				}

				lpClassPrev	= lpClass				;
				lpClass		= lpClass->pNextClass	;
				delete	lpClassPrev	;
			}
			lpScore	= lpSchool->lpScoreLst	;
			while ( NULL != lpScore )
			{
				lpScorePrev	= lpScore				;
				lpScore		= lpScore->pNextScore	;
				delete	lpScorePrev	;
			}

			lpSchoolPrev= lpSchool				;
			lpSchool	= lpSchool->pNextSchool	;
			delete	lpSchoolPrev	;
		}
		lpScore	= lpRegion->lpScoreLst	;
		while ( NULL != lpScore )
		{
			lpScorePrev	= lpScore				;
			lpScore		= lpScore->pNextScore	;
			delete	lpScorePrev	;
		}

		delete	lpRegion;

		// ���ٽ�����
		m_pProgress->DestroyWindow();

		return;
	}

	// ����У��
	int		iReadSchool;

	// �����һ�ζ���ѧУ�ĸ���
	iReadSchool	= lpRegion->iCntSchool % (iMaxRow - 1 );	// ��һ����ȫ����ͳ������
	if ( 0 == iReadSchool )
		iReadSchool	= iMaxRow - 1	;	// ������������1��Ҳ���������

	iMaxRow		= lpSubItem->m_iMaxRowOfBlock[0];

	CString	strEmpty( " " )	;
	lpSubItem->m_pPara->strNameSchool	= &strEmpty	;

	// �����У������
	lpSubItem->ResetRow ( )		;
	lpSchool	= lpSchoolHead	;
	while ( NULL != lpSchool )
	{
		// һ�ζ�������ѧУ�ļ�¼����
		LPSCHOOL_FORM	*lppSchool		;
		LPSCORE_FORM	*lppScore		;	// ��Ÿ�У������
		LPSCORE_FORM	lpScoreRegion	;	// �����������

		lppSchool	= new LPSCHOOL_FORM[iReadSchool];
		lppScore	= new LPSCORE_FORM[iReadSchool]	;

		strHeadLineSchool.Empty ( )		;

		// ������
		lpScoreRegion	= lpRegion->lpScoreLst	;

		strHeadLineSchool	+= lpRegion->cName	;

		// У���� 
		int		i	;
		for ( i=0; (i<iReadSchool) && (NULL!=lpSchool); i++ )
		{
			*(lppSchool+i)	= lpSchool	;
			if ( NULL == lpSchool )
				*(lppScore+i)	= NULL	;
			else
				*(lppScore+i)	= lpSchool->lpScoreLst	;

			strHeadLineSchool	+= "\t"	;
			strHeadLineSchool	+= lpSchool->cName	;

			lpSchool	= lpSchool->pNextSchool	;
		}
		for ( i=iReadSchool; i<iMaxRow-1; i++ )
		{
			strHeadLineSchool	+= "\t ";
		}
		strHeadLineSchool	+= " ";

		// �γ����� blocks
		int		iRow	;
		for ( int iBlock=0; iBlock<lpSubItem->m_iCntBlocks; iBlock++ )
		{
			lpSubItem->m_pPara->iCntRow		= lpSubItem->m_iMaxColOfBlock[iBlock]	;

			strBlock.Empty ( )	;

			// ������ɱ� block
			iCntCol	= 0	;
			while ( NULL != lpScoreRegion )
			{
				CWordFormatPrint *	pFormat	= (CWordFormatPrint *)lpSubItem->m_listlpColumn	;

				// �ȶ���ȫ��������
				lpScore	= lpScoreRegion	;
				for ( iRow=0; iRow<lpSubItem->m_iMaxSubrowOfRow[0]; iRow++ )
				{
					strScore.Format ( pFormat->m_cFormat, 100*lpScore->fValue )	;

					if ( 0 == iRow )
						strBlock	+= strScore + "%"	;	
					else
						strBlock	+= "\t" + strScore + "%"	;

					pFormat		= (CWordFormatPrint *)pFormat->m_pNext		;
					lpScore		= lpScore->pNextScore						;

				}
				lpScoreRegion	= lpScore	;

				// ���뱾 block �и��༶������
				int	j;
				for ( j=0; j<iReadSchool; j++ )
				{
					lpScore	= *(lppScore+j)	;

					// ����һ������
					for ( iRow=0; iRow<lpSubItem->m_iMaxSubrowOfRow[j+1]; iRow++ )
					{
						strScore.Format ( pFormat->m_cFormat, 100*lpScore->fValue )	;

						strBlock	+= "\t" + strScore + "%"	;	

						pFormat		= (CWordFormatPrint *)pFormat->m_pNext		;
						lpScore		= lpScore->pNextScore						;
					}

					*(lppScore+j)	= lpScore	;
				}
				// ����
				for ( j=iReadSchool; j<iMaxRow-1; j++ )
				{
					for ( iRow=0; iRow<lpSubItem->m_iMaxSubrowOfRow[j+1]; iRow++ )
					{
						strBlock	+= "\t "	;	
					}
				}

				strBlock	+= "\r\n"	;

				iCntCol++	;	// ����һ��

				if ( iCntCol == lpSubItem->m_iMaxColOfBlock[iBlock] )
					break;

			}

			// ���� Title
			lpSubItem->MakeTitle ( &m_Sel )	;

			//+++++++++++++++++++++++++++++++++++++
			// copy and paste Headline
			// ��ͷ
			if ( !OpenClipboard(NULL) ) 
				AfxMessageBox ( "Failed open clipboard" ); 

			EmptyClipboard(); 

			hglbCopy	= GlobalAlloc(GMEM_MOVEABLE, strHeadLineSchool.GetLength( ) ); 
			if (hglbCopy == NULL) 
			{ 
				CloseClipboard(); 
				AfxMessageBox ( "Failed alloc mem" ); 
				return ; 
			} 

			lptstrCopy = (LPTSTR)GlobalLock(hglbCopy); 
			memcpy( lptstrCopy, strHeadLineSchool, strHeadLineSchool.GetLength() ); 
			//lptstrCopy[cch] = (TCHAR) 0;    // null character 
			GlobalUnlock(hglbCopy); 

			SetClipboardData(CF_TEXT, hglbCopy); 

			CloseClipboard(); 

			lpSubItem->SelectHead ( &m_Sel )	;
			m_Sel.Paste ( )	;

			//+++++++++++++++++++++++++++++++++++++
			// copy and paste this block
			if ( !OpenClipboard(NULL) ) 
				AfxMessageBox ( "Failed open clipboard" ); 

			EmptyClipboard(); 

			hglbCopy	= GlobalAlloc(GMEM_MOVEABLE, strBlock.GetLength( ) ); 
			if (hglbCopy == NULL) 
			{ 
				CloseClipboard(); 
				AfxMessageBox ( "Failed alloc mem" ); 
				return ; 
			} 

			lptstrCopy = (LPTSTR)GlobalLock(hglbCopy); 
			memcpy( lptstrCopy, strBlock, strBlock.GetLength() ); 
			//lptstrCopy[cch] = (TCHAR) 0;    // null character 
			GlobalUnlock(hglbCopy); 

			SetClipboardData(CF_TEXT, hglbCopy); 

			CloseClipboard(); 

			// ��λ����������һ��
			lpSubItem->LocateBody ( &m_Sel );
			//lpSubItem->ResetCell ( )	;
			lpSubItem->ShiftRow ( &m_Sel )	;
			m_Sel.Paste ( )	;

			//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
			//lpSubItem->m_pPara->iCntRow	-= iMaxRow	;

			//iCntBlock++	;
			//iMaxRow		= lpSubItem->m_iMaxRowOfBlock[iCntBlock];
			//					AfxMessageBox ( *pstrBlock )	;

			m_Sel.HomeKey ( COleVariant((short)6), COleVariant((short)0) );
			pParent->Activate ( )	;
			m_Sel.AttachDispatch( pApp->GetSelection() )	;
			m_Sel.WholeStory( )		;
			m_Sel.Copy ( )			;

			oActiveDoc.Activate ( )	;
			m_Sel.AttachDispatch( pApp->GetSelection() )	;

			// ����ĸ��
			m_Sel.Paste ( )	;
			m_Sel.HomeKey ( COleVariant((short)6), COleVariant((short)0) );

		}	// end of for iBlock

//		iCntSchool	-= iReadSchool	;
		iReadSchool	= iMaxRow - 1	;

		delete	lppSchool;
		delete	lppScore;

		m_pProgress->Step1 ( )	;
	}
	
	// ��������ķ������
	oActiveDoc.SaveAs( COleVariant( strPath + strType + strSubject + lpSubItem->cTitle ), 
		COleVariant((short)0), 
		vFalse, COleVariant(""), vTrue, COleVariant(""), 
		vFalse, vFalse, vFalse, vFalse, vFalse );
	oActiveDoc.Close ( vFalse, COleVariant((short)0), COleVariant((short)0) );

	//���ý�����
	m_pProgress->SetPos1( 100 );
	m_pProgress->SetPos2( 100 );

	// ���ٻ�����
	LPSCHOOL_FORM	lpSchoolPrev;
	LPSCORE_FORM	lpScorePrev	;
	lpSchool	= lpRegion->lpSchoolLst;
	while ( NULL != lpSchool )
	{
		// ���ٰ༶�б�
		LPCLASS_FORM	lpClass		;
		LPCLASS_FORM	lpClassPrev	;
		lpClass	= lpSchool->lpClassLst	;
		while ( NULL != lpClass )
		{
			// ���ٳɼ��б�
			lpScore	= lpClass->lpScoreLst	;
			while ( NULL != lpScore )
			{
				lpScorePrev	= lpScore				;
				lpScore		= lpScore->pNextScore	;
				delete	lpScorePrev	;
			}

			lpClassPrev	= lpClass				;
			lpClass		= lpClass->pNextClass	;
			delete	lpClassPrev	;
		}
		lpScore	= lpSchool->lpScoreLst	;
		while ( NULL != lpScore )
		{
			lpScorePrev	= lpScore				;
			lpScore		= lpScore->pNextScore	;
			delete	lpScorePrev	;
		}

		lpSchoolPrev= lpSchool				;
		lpSchool	= lpSchool->pNextSchool	;
		delete	lpSchoolPrev	;
	}
	lpScore	= lpRegion->lpScoreLst	;
	while ( NULL != lpScore )
	{
		lpScorePrev	= lpScore				;
		lpScore		= lpScore->pNextScore	;
		delete	lpScorePrev	;
	}

	delete	lpRegion;

//	DestroyChoice ( )	;

	// ���ٽ�����
	m_pProgress->DestroyWindow();
}

//++++++++++++++++++++++++++++++++++++++++++++++++++
//	���༶������������
//		ÿ��ѧУÿ��ÿ����һ���ļ�
//
//	��һ��: ��ѧУ����ѭ������Word�ĵ�, ÿ��ѧУһ���ļ�
//	�ڶ���: �� VirtualForm �е���������ģ���
//	������: ���ڶ����еı�����Ŀ����β����
//--------------------------------------------------
void CAnalyzer::FillFormClass ( int iSubjectID, CString strSubject, 
			int iRegionID, CString strRegion, int iTypeID, CString strType,
			_Application* pApp, Documents* pDocs, _Document* pParent, Selection* pSel, 
			LPSUBITEM lpSubItem )
{
	CString	strFilter	;
	CString	strFilter0	;
	CString	strFilter1	;
	CString	strFilter2	;
	CString	strSort		;
	CString	strSchool	;
	CString	strSQL		;
	CString	strPath		;
	//int		iCntClass	;
	int		iCntSchool	;
	//int		iCntStudent	;
	int		iSchoolID	;

	LPSCHOOL_FORM	lpSchoolHead	;
	LPSCHOOL_FORM	lpSchoolTail	;
	LPSCHOOL_FORM	lpSchool		;

	// ������������·��
	//strPath.Format ( "C:\\Forms\\" )	;
	strPath	= m_strPathOutput		;
	SetCurrentDirectory ( strPath )	;

	strPath	= strPath + "\\"		;
	strPath	= strPath + strRegion	;
	if ( !TestDirectory( strPath ) )
		_mkdir ( strRegion )	;
	SetCurrentDirectory ( strPath )	;
	strPath	= strPath + "\\"		;

	Selection		m_Sel	;	// ����Word�ṩ��ѡ�����

	COleVariant varTrue( short(1), VT_BOOL )	;
	COleVariant varFalse( short(0), VT_BOOL )	;
	COleVariant	vTrue( (short)TRUE )			;
	COleVariant	vFalse( (short)FALSE )			;

	CSetSelectID		setIDSchool	;
	CSetSelectID		setIDClass	;

	CvFactStudent *		lpvFactStudent	;
	CvScoreForm *		lpvFactStudent1	;
	CvScoreForm *		lpvFactStudent2	;

	//������
	m_pProgress	= new CProgress(NULL);

	setIDSchool.m_pDatabase	= &theApp.m_database;
	setIDClass.m_pDatabase	= &theApp.m_database;

	lpvFactStudent	= new CvFactStudent;
	lpvFactStudent->m_pDatabase	= &theApp.m_database;

	lpvFactStudent1	= NULL	;
	lpvFactStudent2	= NULL	;
	if ( '0' != lpSubItem->cSQL1[0] )
	{
		lpvFactStudent1	= new CvScoreForm	;
		lpvFactStudent1->m_pDatabase	= &theApp.m_database;

		if ( '0' != lpSubItem->cSQL2[0] )
		{
			lpvFactStudent2	= new CvScoreForm	;
			lpvFactStudent2->m_pDatabase	= &theApp.m_database;
		}
	}

//	�ռ���Ϣ
	strFilter0.Empty ( )	;
	strFilter1.Empty ( )	;
	strFilter2.Empty ( )	;
	m_pProgress->SetTitle1 ( "�ռ���Ϣ" )	;
	try
	{
		strSQL.Format( "SELECT schoolID, schoolName FROM School WHERE (RegionID = %d) ORDER BY schoolID", iRegionID )	;
		setIDSchool.Open ( CRecordset::snapshot, strSQL, CRecordset::readOnly )	;

		strSQL.Format( "%s", lpSubItem->cSQL )	; strSQL.TrimLeft(); strSQL.TrimRight();
		if ( '0' != lpSubItem->cSort[0] )
			lpvFactStudent->m_strSort	= lpSubItem->cSort	;
		if ( '0' != lpSubItem->cFilter[0] )
		{
			lpvFactStudent->m_strFilter	= lpSubItem->cFilter;
			strFilter0.Format ( "%s", lpSubItem->cFilter )	;
		}
		lpvFactStudent->Open ( CRecordset::snapshot, strSQL, CRecordset::readOnly )	;

		if ( lpvFactStudent1 )
		{
			strSQL.Format( "%s", lpSubItem->cSQL1 )	; strSQL.TrimLeft(); strSQL.TrimRight();
			if ( '0' != lpSubItem->cSort1[0] )
				lpvFactStudent1->m_strSort	= lpSubItem->cSort1		;
			if ( '0' != lpSubItem->cFilter1[0] )
			{
				lpvFactStudent1->m_strFilter= lpSubItem->cFilter1	;
				strFilter1.Format ( "%s", lpSubItem->cFilter1 )	;
			}
			lpvFactStudent1->Open ( CRecordset::snapshot, strSQL, CRecordset::readOnly )	;
		}

		if ( lpvFactStudent2 )
		{
			strSQL.Format( "%s", lpSubItem->cSQL1 )	; strSQL.TrimLeft(); strSQL.TrimRight();
			if ( '0' != lpSubItem->cSort2[0] )
				lpvFactStudent2->m_strSort	= lpSubItem->cSort2		;
			if ( '0' != lpSubItem->cFilter2[0] )
			{
				lpvFactStudent2->m_strFilter= lpSubItem->cFilter2	;
				strFilter2.Format ( "%s", lpSubItem->cFilter2 )	;
			}
			lpvFactStudent2->Open ( CRecordset::snapshot, strSQL, CRecordset::readOnly )	;
		}

		lpSchoolHead	= NULL	;
		while ( !setIDSchool.IsEOF() )
		{
			setIDSchool.MoveNext ( )	;
		}
		iCntSchool	= setIDSchool.GetRecordCount ( )	;

		// ��ѧУѭ��
		if ( iCntSchool )
			setIDSchool.MoveFirst ( )	;
		for ( int i=0; i<iCntSchool; i++ )	// for school
		{
			iSchoolID	= setIDSchool.m_iID	;
			strSchool.Format ( "%s", setIDSchool.m_strName )	;
			strSchool.TrimRight ( )		;

			m_pProgress->SetTitle2 ( strSchool )	;

			// ���༶ѭ��
			strSQL.Format( "SELECT classID, className FROM Class WHERE (schoolID = %d) AND (TypeID = %d) ORDER BY className DESC", 
				iSchoolID, iTypeID )	; // ORDER BY classID
			setIDClass.Open ( CRecordset::snapshot, strSQL, CRecordset::readOnly )	;

			LPCLASS_FORM	lpClassHead	;
			LPCLASS_FORM	lpClassTail	;
			LPCLASS_FORM	lpClass		;

			lpClassHead	= NULL	;
			while ( !setIDClass.IsEOF() )
			//{
			//	setIDClass.MoveNext ( )	;
			//}
			//iCntClass	= setIDClass.GetRecordCount ( )	;

			//if ( iCntClass )
			//	setIDClass.MoveFirst ( )	;

//			for ( int iClass=0; iClass<iCntClass; iClass++ )	// for class
			{
				// ���ò���
				//lpSubItem->m_pPara->strNameClass	= &setIDClass.m_strName	;

				// �����ݿ�
				strFilter.Format ( "(SubjectID = %d) AND (ClassID = %d)", iSubjectID, setIDClass.m_iID )	;
				lpvFactStudent->m_strFilter	= strFilter0 + strFilter;
				lpvFactStudent->Requery ( )	;

				LPSTUDENT_FORM	lpStudentHead	;
				LPSTUDENT_FORM	lpStudentTail	;
				LPSTUDENT_FORM	lpStudent		;
				LPSCORE_FORM	lpScoreHead		;
				LPSCORE_FORM	lpScoreTail		;
				LPSCORE_FORM	lpScore			;

				int		iCnt	= 0		;
				int		iID		= -1	;
				lpStudentHead	= NULL	;
				while ( !lpvFactStudent->IsEOF() )
				{
					lpScore	= new SCORE_FORM	;

					lpScore->iScoreTypeID	= lpvFactStudent->m_ScoreTypeID	;
					lpScore->fValue			= lpvFactStudent->m_fValue		;
					lpScore->pNextScore		= NULL							;

					if ( iID != lpvFactStudent->m_StudentID )
					{
						iCnt++	;	// ѧ��������1
						iID	= lpvFactStudent->m_StudentID	;

						lpScoreHead			= lpScore	;
						lpScoreTail			= lpScore	;

						lpStudent	= new STUDENT_FORM	;

						lpStudent->iStudentID	= lpvFactStudent->m_StudentID				;
						strcpy ( lpStudent->cEnrollID, lpvFactStudent->m_StudentEnrollID )	;
						strcpy ( lpStudent->cName, lpvFactStudent->m_StudentName )			;
						lpStudent->lpScoreLst	= lpScoreHead								;
						lpStudent->pNextStudent	= NULL										;

						if ( NULL == lpStudentHead )
						{
							lpStudentHead	= lpStudent	;
							lpStudentTail	= lpStudent	;
						}
						else
						{
							lpStudentTail->pNextStudent	= lpStudent	;
							lpStudentTail				= lpStudent	;
						}
					}
					else
					{
						lpScoreTail->pNextScore	= lpScore	;
						lpScoreTail				= lpScore	;
					}

					lpvFactStudent->MoveNext ( )	;
				}

				// iCntStudent	= lpvFactStudent->GetRecordCount ( )	;

				// ������༶��ѧ���ɼ�����, �򽫱��༶���뵽�༶�б���
				if ( NULL != lpStudentHead )
				{
					lpClass	= new CLASS_FORM	;

					lpClass->iClassID		= setIDClass.m_iID	;
					strcpy ( lpClass->cName, setIDClass.m_strName )	;
					lpClass->lpStudentLst	= lpStudentHead		;
					lpClass->iCntStudent	= iCnt				;
					lpClass->pNextClass		= NULL				;

					if ( NULL == lpClassHead )
					{
						lpClassHead	= lpClass	;
						lpClassTail	= lpClass	;
					}
					else
					{
						lpClassTail->pNextClass	= lpClass	;
						lpClassTail				= lpClass	;
					}

					// ���ܵĻ�����α�͸���
					// ���ȶ���α�
					if ( lpvFactStudent1 )
					{
						lpvFactStudent1->m_strFilter	= strFilter1 + strFilter;
						lpvFactStudent1->Requery ( )	;

						iID	= -1	;
						while ( !lpvFactStudent1->IsEOF() )
						{
							lpScore	= new SCORE_FORM	;

							lpScore->iScoreTypeID	= lpvFactStudent1->m_FormID	;	// >m_ScoreTypeID	;
							lpScore->fValue			= lpvFactStudent1->m_Score	;
							lpScore->pNextScore		= NULL						;

							if ( iID != lpvFactStudent1->m_StudentID )
							{
								// �ҵ��� ID ����Ӧ��ѧ��
								iID	= lpvFactStudent1->m_StudentID	;
								lpStudent	= lpStudentHead	;
								while ( lpStudent )
								{
									if ( iID == lpStudent->iStudentID )
										break	;

									lpStudent	= lpStudent->pNextStudent	;
								}

								if ( NULL != lpStudent )
								{
									lpScoreHead	= lpStudent->lpScoreLst	;
									lpScoreTail	= lpScoreHead			;

									while ( NULL != lpScoreTail->pNextScore )
										lpScoreTail	= lpScoreTail->pNextScore	;
								}
							}

							// �����ݼ��뵽�б���
							if ( NULL != lpStudent )
							{
								lpScoreTail->pNextScore	= lpScore	;
								lpScoreTail				= lpScore	;
							}

							lpvFactStudent1->MoveNext ( )	;
						}
					}
					// Ȼ����븨��
					if ( lpvFactStudent2 )
					{
						lpvFactStudent2->m_strFilter	= strFilter2 + strFilter;
						lpvFactStudent2->Requery ( )	;

						iID	= -1	;
						while ( !lpvFactStudent2->IsEOF() )
						{
							lpScore	= new SCORE_FORM	;

							lpScore->iScoreTypeID	= lpvFactStudent2->m_FormID	;	// >m_ScoreTypeID	;
							lpScore->fValue			= lpvFactStudent2->m_Score	;
							lpScore->pNextScore		= NULL						;

							if ( iID != lpvFactStudent2->m_StudentID )
							{
								// �ҵ��� ID ����Ӧ��ѧ��
								iID	= lpvFactStudent2->m_StudentID	;
								lpStudent	= lpStudentHead	;
								while ( lpStudent )
								{
									if ( iID == lpStudent->iStudentID )
										break	;

									lpStudent	= lpStudent->pNextStudent	;
								}

								if ( NULL != lpStudent )
								{
									lpScoreHead	= lpStudent->lpScoreLst	;
									lpScoreTail	= lpScoreHead			;

									while ( NULL != lpScoreTail->pNextScore )
										lpScoreTail	= lpScoreTail->pNextScore	;
								}
							}

							// �����ݼ��뵽�б���
							if ( NULL != lpStudent )
							{
								lpScoreTail->pNextScore	= lpScore	;
								lpScoreTail				= lpScore	;
							}

							lpvFactStudent2->MoveNext ( )	;
						}
					}				
				}

				//m_pProgress->SetPos1( 100 * iClass / iCntClass );
				m_pProgress->Step1 ( )	;

				setIDClass.MoveNext ( )	;
			}// end of for-class

			setIDClass.Close ( );

			if ( NULL != lpClassHead )
			{
				lpSchool	= new SCHOOL_FORM	;

				lpSchool->iSchoolID		= iSchoolID					;
				strcpy ( lpSchool->cName, setIDSchool.m_strName )	;
				lpSchool->lpClassLst	= lpClassHead				;
				lpSchool->pNextSchool	= NULL						;

				if ( NULL == lpSchoolHead )
				{
					lpSchoolHead	= lpSchool	;
					lpSchoolTail	= lpSchool	;
				}
				else
				{
					lpSchoolTail->pNextSchool	= lpSchool	;
					lpSchoolTail				= lpSchool	;
				}
			}

			m_pProgress->SetPos2( 100 * i / iCntSchool );

			if ( i+1 < iCntSchool )
				setIDSchool.MoveNext ( )	;
		} // end of for school

		lpvFactStudent->Close ( )	;
		delete	lpvFactStudent	;

		if ( NULL != lpvFactStudent1 )
		{
			lpvFactStudent1->Close ( )	;
			delete	lpvFactStudent1		;

			if ( NULL != lpvFactStudent2 )
			{
				lpvFactStudent2->Close ( )	;
				delete	lpvFactStudent2		;
			}
		}

		setIDSchool.Close ( )	;
	}
	catch ( CDBException * e )
	{
		AfxMessageBox ( e->m_strError )	;
	}

//	���ɱ���
	// ��ѧУѭ��
	m_pProgress->SetTitle1 ( "���ɱ���" )	;
	m_pProgress->SetPos1( 0 );
	m_pProgress->SetPos2( 0 );
	int	i		= 0				;
	lpSchool	= lpSchoolHead	;
	while ( NULL != lpSchool )
	{
		_Document	oActiveDoc	;

		iSchoolID	= lpSchool->iSchoolID	;
		strSchool.Format ( "%s", lpSchool->cName )	;
		strSchool.TrimRight ( )		;

		// ���ò���
		lpSubItem->m_pPara->strNameSchool	= &strSchool	;

		m_pProgress->SetTitle2 ( strSchool )	;

		//m_Sel.AttachDispatch( pApp->GetSelection() )	;	// ��Selection�����m_Sel��Idispatch�ӿڹ�������
		//// ȫѡ����
		//m_Sel.WholeStory( )	;
		//m_Sel.Copy ( )		;

		// ÿ��ѧУ������һ���µ� Word �ĵ�
		pDocs->Add ( COleVariant(pParent->GetFullName()), varFalse, COleVariant((short)0), varTrue )	;
		oActiveDoc = pApp->GetActiveDocument(); 
		m_Sel.AttachDispatch( pApp->GetSelection() )	;	// ��Selection�����m_Sel��Idispatch�ӿڹ�������

		// ���༶ѭ��
		LPCLASS_FORM	lpClass		;
		lpClass	= lpSchool->lpClassLst	;
		while ( NULL != lpClass )
		{
			CString	strNameClass	;

			// ���ò���
			strNameClass.Format ( "%s", lpClass->cName )	;
			lpSubItem->m_pPara->strNameClass	= &strNameClass			;	//&setIDClass.m_strName	;
			lpSubItem->m_pPara->iCntRow			= lpClass->iCntStudent	;

			//pParent->Activate ( )	;
			//m_Sel.AttachDispatch( pApp->GetSelection() )	;
			//m_Sel.WholeStory( )		;
			//m_Sel.Copy ( )			;
			//
			//oActiveDoc.Activate ( )	;
			//m_Sel.AttachDispatch( pApp->GetSelection() )	;

			//// ����ĸ��
			//m_Sel.Paste ( )	;

			//// ���� Title
			//lpSubItem->MakeTitle ( &m_Sel )	;

			//// ��λ����������һ��
			//lpSubItem->LocateBody ( &m_Sel );

			// �����ַ����б�
			LPSTUDENT_FORM	lpStudent	;
			CString			strID		;
			CString			strName		;
			CString			strScore	;
			CString			strBlock	;

			int				iReadStudent;	// ��Ҫ�����ѧ������
			int				iCntStudent	;
			int				iMaxRow		;
			int				iCntRow		;
			int				iCntBlock	;
			LPTSTR	lptstrCopy	; 
			HGLOBAL hglbCopy	; 

			iCntRow		= 0	;
			iCntBlock	= 0	;
			
			iMaxRow		= lpSubItem->m_iMaxRowOfBlock[0];

			iCntStudent	= lpClass->iCntStudent	;

			// �����һ�ζ��������
			if ( iMaxRow == iCntStudent )
				iReadStudent	= iMaxRow	;	// ������һ��
			else
				iReadStudent	= iCntStudent % iMaxRow	;

			// �����ѧ���ĳɼ�����
			//lpSubItem->ResetRow ( )	;
			lpStudent	= lpClass->lpStudentLst	;
			while ( NULL != lpStudent )
			{
				// һ�ζ�������ѧ���ļ�¼����
				strBlock.Empty ( )	;

				int		i	;
				for ( i=0; (i<iReadStudent) && (NULL!=lpStudent); i++ )
				{
					// ����д���ź�����
					strID.Format ( "%s", lpStudent->cEnrollID )		;
					strBlock	+= strID							;	// &lpvFactStudent->m_StudentEnrollID	;
					strName.Format ( "\t%s", lpStudent->cName )		;
					strBlock	+= strName							;	// &lpvFactStudent->m_StudentEnrollID	;
					// ���뱾ѧ���ĳɼ�����
					LPSCORE_FORM	lpScore		;

					CWordFormatPrint *	pFormat	= (CWordFormatPrint *)lpSubItem->m_listlpColumn	;

					lpScore	= lpStudent->lpScoreLst	;
					while ( NULL != lpScore )
					{
						strScore.Format ( pFormat->m_cFormat, lpScore->fValue )	;
						strBlock	+= "\t" + strScore							;	// &lpvFactStudent->m_StudentEnrollID	;

						pFormat		= (CWordFormatPrint *)pFormat->m_pNext		;
						lpScore		= lpScore->pNextScore						;
					}

					strBlock	+= "\r\n"	;

					lpStudent	= lpStudent->pNextStudent	;
				}
				// ����
				for ( int j=iReadStudent; j<iMaxRow; j++ )
				{
					//for ( iRow=0; iRow<lpSubItem->m_iMaxSubrowOfRow[j+1]; iRow++ )
					//{
						strBlock	+= "\t \r\n"	;	
					//}
				}
				
				strBlock	+= "\t \r\n";	// ���һ�����һ��Ԫ��ӵ㶫������Ȼ�Բ���

				if ( !OpenClipboard(NULL) ) 
					AfxMessageBox ( "Failed open clipboard" ); 

				EmptyClipboard(); 

				hglbCopy	= GlobalAlloc(GMEM_MOVEABLE, strBlock.GetLength( ) ); 
				if (hglbCopy == NULL) 
				{ 
					CloseClipboard(); 
					AfxMessageBox ( "Failed alloc mem" ); 
					return ; 
				} 

				lptstrCopy = (LPTSTR)GlobalLock(hglbCopy); 
				memcpy( lptstrCopy, strBlock, strBlock.GetLength() ); 
				//lptstrCopy[cch] = (TCHAR) 0;    // null character 
				GlobalUnlock(hglbCopy); 

				SetClipboardData(CF_TEXT, hglbCopy); 

				CloseClipboard(); 

				// ���� Title
				lpSubItem->MakeTitle ( &m_Sel )	;

				// ��λ����������һ��
				lpSubItem->LocateBody ( &m_Sel );

				//lpSubItem->ResetCell ( )	;
				lpSubItem->ResetRow ( )	;
				lpSubItem->ShiftRow ( &m_Sel )	;
				m_Sel.Paste ( )	;
				m_Sel.HomeKey ( COleVariant((short)6), COleVariant((short)0) );
				pParent->Activate ( )	;
				m_Sel.AttachDispatch( pApp->GetSelection() )	;
				m_Sel.WholeStory( )		;
				m_Sel.Copy ( )			;

				oActiveDoc.Activate ( )	;
				m_Sel.AttachDispatch( pApp->GetSelection() )	;

				// ����ĸ��
				m_Sel.Paste ( )	;
				m_Sel.HomeKey ( COleVariant((short)6), COleVariant((short)0) );

				lpSubItem->m_pPara->iCntRow	-= iReadStudent	;

				iReadStudent	= iMaxRow	;

				//lpStudent		= lpStudent->pNextStudent	;
			}

			lpClass		= lpClass->pNextClass	;

			// �ص����
//			m_Sel.HomeKey ( COleVariant((short)6), COleVariant((short)0) );
				
			m_pProgress->Step1 ( )	;
		}

		
		if ( !TestDirectory ( strSchool ) )
			_mkdir ( strSchool );
		SetCurrentDirectory ( strPath + strSchool )	;

		if ( !TestDirectory ( strPath + strSchool + "\\" + strType ) )
			_mkdir ( strType )	;
		SetCurrentDirectory ( strPath + strSchool + "\\" + strType );

		oActiveDoc.SaveAs( COleVariant( strPath + strSchool + "\\" + strType + "\\" + strSubject ), 
			COleVariant((short)0), 
			vFalse, COleVariant(""), vTrue, COleVariant(""), 
			vFalse, vFalse, vFalse, vFalse, vFalse );

		oActiveDoc.Close ( vFalse, COleVariant((short)0), COleVariant((short)0) );
		
		SetCurrentDirectory ( strPath )	;

		lpSchool	= lpSchool->pNextSchool	;

		m_pProgress->SetPos1 ( 100 )				;
		m_pProgress->SetPos2( 100 * i / iCntSchool );
		m_pProgress->SetPos1 ( 0 )					;
		i++;
	}

	//���ý�����
	m_pProgress->SetPos1( 100 );
	m_pProgress->SetPos2( 100 );

	// ���ٽ�����
	m_pProgress->DestroyWindow();

	// ������Ϣ�б�
	LPSCHOOL_FORM	lpSchoolPrev;
	lpSchool	= lpSchoolHead	;
	while ( NULL != lpSchool )
	{
		// ���ٰ༶�б�
		LPCLASS_FORM	lpClass		;
		LPCLASS_FORM	lpClassPrev	;
		lpClass	= lpSchool->lpClassLst	;
		while ( NULL != lpClass )
		{
			// ����ѧ���б�
			LPSTUDENT_FORM	lpStudent		;
			LPSTUDENT_FORM	lpStudentPrev	;
			lpStudent	= lpClass->lpStudentLst	;
			while ( NULL != lpStudent )
			{
				// ���ٳɼ��б�
				LPSCORE_FORM	lpScore		;
				LPSCORE_FORM	lpScorePrev	;
				lpScore	= lpStudent->lpScoreLst	;
				while ( NULL != lpScore )
				{
					lpScorePrev	= lpScore				;
					lpScore		= lpScore->pNextScore	;
					delete	lpScorePrev	;
				}

				lpStudentPrev	= lpStudent					;
				lpStudent		= lpStudent->pNextStudent	;
				delete	lpStudentPrev	;
			}

			lpClassPrev	= lpClass				;
			lpClass		= lpClass->pNextClass	;
			delete	lpClassPrev	;
		}

		lpSchoolPrev= lpSchool				;
		lpSchool	= lpSchool->pNextSchool	;
		delete	lpSchoolPrev	;
	}
}

//++++++++++++++++++++++++++++++++++++++++++++++++++
//	����У�������������
//		ÿ��һ���ļ�
//
//	��һ��: 
//	�ڶ���: 
//	������: 
//--------------------------------------------------
void CAnalyzer::FillFormRegionSchoolClass ( int iSubjectID, CString strSubject,  
			int iRegionID, CString strRegion, int iTypeID, CString strType,
			_Application* pApp, Documents* pDocs, _Document* pParent, Selection* pSel, 
			LPSUBITEM lpSubItem )
{
	// ���ݿ���ز���
	CString	strFilter	;
	CString	strFilter0	;
	CString	strFilter1	;
	CString	strFilter2	;
	CString	strSort		;
	CString	strSchool	;
	CString	strSQL		;

	// ���ݿ�
	//CSetSelectID		setIDSchool				;
	CvAVGClass	*lpvChoicePercentage	;
	CvAVGClass	*lpvChoicePercentage1	;
	CvAVGClass	*lpvChoicePercentage2	;

	// ������ز���
	CString	strPath		;	// ����Ĵ��·��

	// ��ŷ������ݵ� Buffer
	LPREGION_FORM	lpRegion	;	// ��
	LPSCHOOL_FORM	lpSchoolHead;
	LPSCHOOL_FORM	lpSchoolTail;
	LPSCHOOL_FORM	lpSchool	;
	LPSCORE_FORM	lpScoreHead	;
	LPSCORE_FORM	lpScoreTail	;
	LPSCORE_FORM	lpScore		;

	// Word ����ز���
	Selection	m_Sel							;	// ����Word�ṩ��ѡ�����
	COleVariant varTrue( short(1), VT_BOOL )	;
	COleVariant varFalse( short(0), VT_BOOL )	;
	COleVariant	vTrue( (short)TRUE )			;
	COleVariant	vFalse( (short)FALSE )			;

	// ���ñ���
	CString			strID		;
	CString			strName		;
	CString			strScore	;
	CString			strBlock	;

	// ��ʼ������
	m_pProgress	= new CProgress(NULL)		;	// ������
	lpRegion	= new REGION_FORM			;	// ÿ��ֻ����һ����
	
	strPath	= m_strPathOutput		;
	SetCurrentDirectory ( strPath )	;

	strPath	= strPath + "\\"		;
	strPath	= strPath + strRegion	;
	if ( !TestDirectory( strPath ) )
		_mkdir ( strRegion )	;
	SetCurrentDirectory ( strPath )	;
	strPath	= strPath + "\\"		;
	
	//setIDSchool.m_pDatabase	= &theApp.m_database;

	lpvChoicePercentage	= new CvAVGClass;
	lpvChoicePercentage->m_pDatabase	= &theApp.m_database;

	lpvChoicePercentage1	= NULL	;
	lpvChoicePercentage2	= NULL	;
	if ( '0' != lpSubItem->cSQL1[0] )
	{
		lpvChoicePercentage1	= new CvAVGClass;
		lpvChoicePercentage1->m_pDatabase	= &theApp.m_database;

		if ( '0' != lpSubItem->cSQL2[0] )
		{
			lpvChoicePercentage2	= new CvAVGClass;
			lpvChoicePercentage2->m_pDatabase	= &theApp.m_database;
		}
	}


//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//	�ռ���Ϣ
	strFilter0.Empty ( )	;
	strFilter1.Empty ( )	;
	strFilter2.Empty ( )	;
	m_pProgress->SetTitle2 ( "�ռ���Ϣ" )	;
	try
	{
		//+++++++++++++++++++++++++++
		// �ȶ�����������
		strSQL.Format( "%s", lpSubItem->cSQL )	; strSQL.TrimLeft(); strSQL.TrimRight();
		if ( '0' != lpSubItem->cSort[0] )
			lpvChoicePercentage->m_strSort	= lpSubItem->cSort	;
		if ( '0' != lpSubItem->cFilter[0] )
		{
			//lpvChoicePercentage->m_strFilter	= lpSubItem->cFilter;
			strFilter0.Format ( "%s", lpSubItem->cFilter )	;
		}
		lpvChoicePercentage->m_strFilter.Format ( "(SubjectID = %d) AND (Type = %d)", iSubjectID, iTypeID );
		lpvChoicePercentage->m_strFilter	+= strFilter0;
		lpvChoicePercentage->Open ( CRecordset::snapshot, strSQL, CRecordset::readOnly )	;

		// ����ɼ���������
		lpScoreHead	= NULL	;
		lpScoreTail	= NULL	;
		while ( !lpvChoicePercentage->IsEOF() )
		{
			lpScore	= new SCORE_FORM	;

			lpScore->iScoreTypeID	= lpvChoicePercentage->m_FormID		;
			lpScore->fValue			= lpvChoicePercentage->m_avgScore	;
			lpScore->pNextScore		= NULL								;

			if ( NULL == lpScoreHead )
			{
				lpScoreHead	= lpScore	;
				lpScoreTail	= lpScore	;

				strcpy ( lpRegion->cName, lpvChoicePercentage->m_ClassName )	;
			}
			else
			{
				lpScoreTail->pNextScore	= lpScore	;
				lpScoreTail				= lpScore	;
			}

			lpvChoicePercentage->MoveNext ( )	;
		}
		lpRegion->lpScoreLst	= lpScoreHead	;

		// �����ݶ������, ������Դ
		lpvChoicePercentage->Close ( )	;
		delete	lpvChoicePercentage		;

		//+++++++++++++++++++++++++++
		// �ٶ������и�ѧУ������
		if ( lpvChoicePercentage1 )
		{
			strSQL.Format( "%s", lpSubItem->cSQL1 )	; strSQL.TrimLeft(); strSQL.TrimRight();
			if ( '0' != lpSubItem->cSort1[0] )
				lpvChoicePercentage1->m_strSort	= lpSubItem->cSort1		;
			if ( '0' != lpSubItem->cFilter1[0] )
			{
				//lpvChoicePercentage1->m_strFilter= lpSubItem->cFilter1	;
				strFilter1.Format ( "%s", lpSubItem->cFilter1 )	;
			}
			lpvChoicePercentage1->m_strFilter.Format ( "(SubjectID = %d) AND (Type = %d)", iSubjectID, iTypeID );
			lpvChoicePercentage1->m_strFilter	+= strFilter1;
			lpvChoicePercentage1->Open ( CRecordset::snapshot, strSQL, CRecordset::readOnly )	;

			// ��������
			//lpScoreHead	= NULL	;
			//lpScoreTail	= NULL	;
			int		iID		= -1	;
			int		iCnt	= 0		;
			lpSchoolHead	= NULL	;
			while ( !lpvChoicePercentage1->IsEOF() )
			{
				lpScore	= new SCORE_FORM	;

				lpScore->iScoreTypeID	= lpvChoicePercentage1->m_FormID	;
				lpScore->fValue			= lpvChoicePercentage1->m_avgScore	;
				lpScore->pNextScore		= NULL								;

				// ������µ�ѧУ
				if ( iID != lpvChoicePercentage1->m_ClassID )
				{
					iCnt++	;	// ѧУ��Ŀ��1
					iID	= lpvChoicePercentage1->m_ClassID;	// �� iID

					// �½�һ��ѧУ
					lpSchool	= new SCHOOL_FORM	;

					lpSchool->iSchoolID		= iID		;
					lpSchool->lpScoreLst	= lpScore	;
					lpSchool->pNextSchool	= NULL		;
					strcpy ( lpSchool->cName, lpvChoicePercentage1->m_ClassName )	;

					lpScoreHead			= lpScore	;
					lpScoreTail			= lpScore	;

					if ( NULL == lpSchoolHead )
					{
						lpSchoolHead	= lpSchool	;
						lpSchoolTail	= lpSchool	;
					}
					else
					{
						lpSchoolTail->pNextSchool	= lpSchool	;
						lpSchoolTail				= lpSchool	;
					}
				}
				else
				{
					lpScoreTail->pNextScore	= lpScore	;
					lpScoreTail				= lpScore	;
				}

				lpvChoicePercentage1->MoveNext ( )	;
			}

			// ѧУ���ݶ������, ������Դ
			lpvChoicePercentage1->Close ( )	;
			delete	lpvChoicePercentage1	;

			// ��ѧУ�ҽӵ�����
			lpRegion->lpSchoolLst	= lpSchoolHead	;
			lpRegion->iCntSchool	= iCnt			;
		}

		//strSQL.Format( "SELECT schoolID, schoolName FROM School WHERE (RegionID = %d) ORDER BY schoolID", iRegionID )	;
		//setIDSchool.Open ( CRecordset::snapshot, strSQL, CRecordset::readOnly )	;

		//+++++++++++++++++++++++++++
		// �ζ������и�У�и���ķ�������
		if ( lpvChoicePercentage2 )
		{
			strSQL.Format( "%s", lpSubItem->cSQL2 )	; strSQL.TrimLeft(); strSQL.TrimRight();
			if ( '0' != lpSubItem->cSort2[0] )
				lpvChoicePercentage2->m_strSort	= lpSubItem->cSort2		;
			if ( '0' != lpSubItem->cFilter2[0] )
			{
				//lpvChoicePercentage2->m_strFilter= lpSubItem->cFilter2	;
				strFilter2.Format ( "%s", lpSubItem->cFilter2 )	;
			}
			lpvChoicePercentage2->m_strFilter.Format ( "(SubjectID = %d) AND (Type = %d)", iSubjectID, iTypeID );
			lpvChoicePercentage2->m_strFilter	+= strFilter2;
			lpvChoicePercentage2->Open ( CRecordset::snapshot, strSQL, CRecordset::readOnly )	;

			//int		iCntSchool	;
			//int		iCntStudent	;
			int		iSchoolID	;

			//int				iMaxRow		;
			//int				iCntRow		;
			//int				iCntBlock	;

			int		i	= 0				;	// ����
			lpSchool	= lpSchoolHead	;
			while ( NULL != lpSchool )
			{
				// ȡ�ø�ѧУ�������
				iSchoolID	= lpSchool->iSchoolID			;
				strSchool.Format ( "%s", lpSchool->cName )	;
				strSchool.TrimLeft ()	;	strSchool.TrimRight ()	;

				// ���ý��ȱ�־
				m_pProgress->SetTitle2 ( strSchool );

				// ��ȡ�µ����ݿ��¼��
				strFilter.Format ( "(SubjectID = %d) AND (Type = %d) AND (schoolID = %d)", iSubjectID, iTypeID, iSchoolID )	;
				lpvChoicePercentage2->m_strFilter	= strFilter2 + strFilter;
				lpvChoicePercentage2->Requery ( )	;

				// �����ݿ��¼��
				LPCLASS_FORM	lpClassHead	;
				LPCLASS_FORM	lpClassTail	;
				LPCLASS_FORM	lpClass		;

				int		iCnt	= 0		;
				int		iID		= -1	;
				lpClassHead		= NULL	;
				while ( !lpvChoicePercentage2->IsEOF() )
				{
					lpScore	= new SCORE_FORM	;

					lpScore->iScoreTypeID	= lpvChoicePercentage2->m_FormID	;
					lpScore->fValue			= lpvChoicePercentage2->m_avgScore	;
					lpScore->pNextScore		= NULL								;
				
					if ( iID != lpvChoicePercentage2->m_ClassID )
					{
						iCnt++	;	// �༶����1
						iID	= lpvChoicePercentage2->m_ClassID	;

						lpClass	= new CLASS_FORM	;

						lpClass->iClassID	= iID		;
						lpClass->lpScoreLst	= lpScore	;
						lpClass->pNextClass	= NULL		;
						strcpy ( lpClass->cName, lpvChoicePercentage2->m_ClassName )	;

						lpScoreHead			= lpScore	;
						lpScoreTail			= lpScore	;

						if ( NULL == lpClassHead )
						{
							lpClassHead	= lpClass	;
							lpClassTail	= lpClass	;
						}
						else
						{
							lpClassTail->pNextClass	= lpClass	;
							lpClassTail				= lpClass	;
						}
					}
					else
					{
						lpScoreTail->pNextScore	= lpScore	;
						lpScoreTail				= lpScore	;
					}

					// �ƶ�����һ��¼
					lpvChoicePercentage2->MoveNext ( )	;
				}

				// ���༶�б����ӵ�ѧУ
				lpSchool->lpClassLst= lpClassHead	;
				lpSchool->iCntClass	= iCnt			;

				// ������һѧУ
				lpSchool	= lpSchool->pNextSchool	;

				// ���ý���
				m_pProgress->SetPos2 ( ++i * 100 / lpRegion->iCntSchool )	;
			}

			// ��У�ĸ�������ݶ������, ������Դ
			lpvChoicePercentage2->Close ( )	;
			delete	lpvChoicePercentage2	;
		}

	}
	catch ( CDBException * e )
	{
		AfxMessageBox ( e->m_strError )	;
		return;
	}
	
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//	���ɱ���
	m_pProgress->SetTitle1 ( "���ɱ���" )	;
	m_pProgress->SetPos1( 0 );
	m_pProgress->SetPos2( 0 );

	LPCLASS_FORM	lpClass		;
	LPTSTR			lptstrCopy	; 
	HGLOBAL			hglbCopy	; 

	// ÿ����ֻ����һ�� Word �ĵ�
	_Document	oActiveDoc	;
	pDocs->Add ( COleVariant(pParent->GetFullName()), varFalse, COleVariant((short)0), varTrue )	;
	oActiveDoc = pApp->GetActiveDocument(); 
	m_Sel.AttachDispatch( pApp->GetSelection() )	;	// ��Selection�����m_Sel��Idispatch�ӿڹ�������

	// ��ѧУѭ��
	CString	strHeadLineClass	;
	CString	strHeadLineSchool	;
	int		iMaxRow				;
	int		iCntCol				;
	int	ii		= 0				;
	lpSchool	= lpSchoolHead	;
	while ( NULL != lpSchool )
	{
		int	iCntClass	;
		int	iReadClass	;	// ��ǰ��Ҫ����İ༶��
		//int	iCntRow		;
		int	iCntBlock	;
		int	iSchoolID	;

		//iCntRow		= 0	;
		iCntCol		= 0	;
		iCntBlock	= 0	;

		iMaxRow		= lpSubItem->m_iMaxRowOfBlock[0];

		iCntClass	= lpSchool->iCntClass	;
		iSchoolID	= lpSchool->iSchoolID	;
		strSchool.Format ( "%s", lpSchool->cName )	;
		strSchool.TrimRight ( )		;
			
		// ���ò���
		lpSubItem->m_pPara->strNameSchool	= &strSchool			;
		//lpSubItem->m_pPara->iCntRow			= lpSchool->iCntClass	;
		//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

		m_pProgress->SetTitle2 ( strSchool );
		m_pProgress->SetPos1 ( 0 )			;

		// �����һ�ζ���İ༶����
		iReadClass	= iCntClass % (iMaxRow - 1 );	// ��һ����ȫУ��ͳ������
		if ( 0 == iReadClass )
			iReadClass	= iMaxRow - 1	;	// ������������1��Ҳ���������

		// ������������
		lpSubItem->ResetRow ( )			;
		lpClass	= lpSchool->lpClassLst	;
		while ( NULL != lpClass )
		{
			// һ�ζ���������ļ�¼����
			LPCLASS_FORM	*lppClass			;
			LPSCORE_FORM	*lppScore			;	// ��Ÿ��������
			LPSCORE_FORM	lpScoreSchool		;	// ���ѧУ������
			LPSCORE_FORM	lpScoreSchoolHead	;


			lppClass	= new LPCLASS_FORM[iReadClass]	;
			lppScore	= new LPSCORE_FORM[iReadClass]	;
		
			strHeadLineClass.Empty ( )		;

			// У����
			lpScoreSchool	= lpSchool->lpScoreLst	;

			strHeadLineClass	+= "ȫУ"	;	//lpSchool->cName	;

			// ������ 
			int		i	;
			for ( i=0; (i<iReadClass) && (NULL!=lpClass); i++ )
			{
				*(lppClass+i)	= lpClass	;
				if ( NULL == lpClass )
					*(lppScore+i)	= NULL	;
				else
					*(lppScore+i)	= lpClass->lpScoreLst	;

				strHeadLineClass	+= "\t"	;
				strHeadLineClass	+= lpClass->cName	;

				lpClass	= lpClass->pNextClass	;
			}
			for ( i=iReadClass; i<iMaxRow; i++ )
			{
				strHeadLineClass	+= "\t "	;
			}

			// �γ����� blocks
			int		iRow	;
			for ( int iBlock=0; iBlock<lpSubItem->m_iCntBlocks; iBlock++ )
			{
				lpSubItem->m_pPara->iCntRow		= lpSubItem->m_iMaxColOfBlock[iBlock]	;

				strBlock.Empty ( )	;

				// ������ɱ� block
				iCntCol				= 0				;
				while ( NULL != lpScoreSchool )
				{
					CWordFormatPrint *	pFormat	= (CWordFormatPrint *)lpSubItem->m_listlpColumn	;

					// �ȶ���ȫУ������
					lpScore				= lpScoreSchool	;
					lpScoreSchoolHead	= lpScoreSchool	;	// �� block ������λ��
					for ( iRow=0; iRow<lpSubItem->m_iMaxSubrowOfRow[0]; iRow++ )
					{
						strScore.Format ( pFormat->m_cFormat, lpScore->fValue )	;

						if ( 0 == iRow )
							strBlock	+= strScore	;	
						else
							strBlock	+= "\t" + strScore	;

						pFormat		= (CWordFormatPrint *)pFormat->m_pNext		;
						lpScore		= lpScore->pNextScore						;

					}
					lpScoreSchool	= lpScore	;

					// ���뱾 block �и��༶������
					LPSCORE_FORM	lpScore0	;
					int				j			;

					lpScore0	= lpScoreSchoolHead	;
					for ( j=0; j<iReadClass; j++ )
					{
						lpScore	= *(lppScore+j)		;

						// ����һ������
						for ( iRow=0; iRow<lpSubItem->m_iMaxSubrowOfRow[j+1]; iRow++ )
						{
							// ��ƽ��
							strScore.Format ( pFormat->m_cFormat, lpScore->fValue )	;
							strBlock	+= "\t" + strScore	;

							// ����
							strScore.Format ( pFormat->m_cFormat, lpScore->fValue - lpScore0->fValue )	;
							strBlock	+= "\t" + strScore	;

							pFormat		= (CWordFormatPrint *)pFormat->m_pNext		;
							lpScore		= lpScore->pNextScore						;
						}

						*(lppScore+j)	= lpScore	;
					}
					// ����
					for ( j=iReadClass; j<iMaxRow; j++ )
					{
						for ( iRow=0; iRow<lpSubItem->m_iMaxSubrowOfRow[j+1]; iRow++ )
						{
							strBlock	+= "\t \t "	;	
						}
					}

					lpScore0	= lpScore0->pNextScore	;

					strBlock	+= "\r\n"	;

					iCntCol++	;	// ����һ��

					if ( iCntCol == lpSubItem->m_iMaxColOfBlock[iBlock] )
						break;

				}
				strBlock	+= "\t \r\n";	// ���һ�����һ��Ԫ��ӵ㶫������Ȼ�Բ���

				// ���� Title
				lpSubItem->MakeTitle ( &m_Sel )	;


				//+++++++++++++++++++++++++++++++++++++
				// copy and paste Headline
				// ��ͷ
				if ( !OpenClipboard(NULL) ) 
					AfxMessageBox ( "Failed open clipboard" ); 

				EmptyClipboard(); 

				hglbCopy	= GlobalAlloc(GMEM_MOVEABLE, strHeadLineClass.GetLength( ) ); 
				if (hglbCopy == NULL) 
				{ 
					CloseClipboard(); 
					AfxMessageBox ( "Failed alloc mem" ); 
					return ; 
				} 

				lptstrCopy = (LPTSTR)GlobalLock(hglbCopy); 
				memcpy( lptstrCopy, strHeadLineClass, strHeadLineClass.GetLength() ); 
				//lptstrCopy[cch] = (TCHAR) 0;    // null character 
				GlobalUnlock(hglbCopy); 

				SetClipboardData(CF_TEXT, hglbCopy); 

				CloseClipboard(); 

				lpSubItem->SelectHead ( &m_Sel )	;
				m_Sel.Paste ( )	;

				//+++++++++++++++++++++++++++++++++++++
				// copy and paste this block
				if ( !OpenClipboard(NULL) ) 
					AfxMessageBox ( "Failed open clipboard" ); 

				EmptyClipboard(); 

				hglbCopy	= GlobalAlloc(GMEM_MOVEABLE, strBlock.GetLength( ) ); 
				if (hglbCopy == NULL) 
				{ 
					CloseClipboard(); 
					AfxMessageBox ( "Failed alloc mem" ); 
					return ; 
				} 

				lptstrCopy = (LPTSTR)GlobalLock(hglbCopy); 
				memcpy( lptstrCopy, strBlock, strBlock.GetLength() ); 
				//lptstrCopy[cch] = (TCHAR) 0;    // null character 
				GlobalUnlock(hglbCopy); 

				SetClipboardData(CF_TEXT, hglbCopy); 

				CloseClipboard(); 

				// ��λ����������һ��
				lpSubItem->LocateBody ( &m_Sel );
				//lpSubItem->ResetCell ( )	;
				lpSubItem->ShiftRow ( &m_Sel )	;
				m_Sel.Paste ( )	;

				//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
				//lpSubItem->m_pPara->iCntRow	-= iMaxRow	;

				//iCntBlock++	;
				//iMaxRow		= lpSubItem->m_iMaxRowOfBlock[iCntBlock];
				//					AfxMessageBox ( *pstrBlock )	;

				m_Sel.HomeKey ( COleVariant((short)6), COleVariant((short)0) );
				pParent->Activate ( )	;
				m_Sel.AttachDispatch( pApp->GetSelection() )	;
				m_Sel.WholeStory( )		;
				m_Sel.Copy ( )			;

				oActiveDoc.Activate ( )	;
				m_Sel.AttachDispatch( pApp->GetSelection() )	;

				// ����ĸ��
				m_Sel.Paste ( )	;
				m_Sel.HomeKey ( COleVariant((short)6), COleVariant((short)0) );

			}	// end of for iBlock

			iCntClass	-= iReadClass	;
			iReadClass	= iMaxRow - 1	;

			delete	lppClass;
			delete	lppScore;

			m_pProgress->Step1 ( )	;

		}	// end of while lpClass

		lpSchool	= lpSchool->pNextSchool	;

		// �ص����
		m_Sel.HomeKey ( COleVariant((short)6), COleVariant((short)0) );

		m_pProgress->SetPos2( 100 * ii++ / lpRegion->iCntSchool );
	
	}	// end of while lpSchool

	// ����У��
	int		iReadSchool	;

	// �����һ�ζ���ѧУ�ĸ���
	iReadSchool	= lpRegion->iCntSchool % (iMaxRow - 1 );	// ��һ����ȫ����ͳ������
	if ( 0 == iReadSchool )
		iReadSchool	= iMaxRow - 1	;	// ������������1��Ҳ���������

	iMaxRow		= lpSubItem->m_iMaxRowOfBlock[0];

	CString	strEmpty( " " )	;
	lpSubItem->m_pPara->strNameSchool	= &strEmpty	;

	// �����У������
	lpSubItem->ResetRow ( )		;
	lpSchool	= lpSchoolHead	;
	while ( NULL != lpSchool )
	{
		// һ�ζ�������ѧУ�ļ�¼����
		LPSCHOOL_FORM	*lppSchool			;
		LPSCORE_FORM	*lppScore			;	// ��Ÿ�У������
		LPSCORE_FORM	lpScoreRegion		;	// �����������
		LPSCORE_FORM	lpScoreRegionHead	;

		lppSchool	= new LPSCHOOL_FORM[iReadSchool];
		lppScore	= new LPSCORE_FORM[iReadSchool]	;

		strHeadLineSchool.Empty ( )		;

		// ������
		lpScoreRegion	= lpRegion->lpScoreLst	;

		strHeadLineSchool	+= "ȫ��"	;	//lpRegion->cName	;

		// У���� 
		int		i	;
		for ( i=0; (i<iReadSchool) && (NULL!=lpSchool); i++ )
		{
			*(lppSchool+i)	= lpSchool	;
			if ( NULL == lpSchool )
				*(lppScore+i)	= NULL	;
			else
				*(lppScore+i)	= lpSchool->lpScoreLst	;

			strHeadLineSchool	+= "\t"	;
			strHeadLineSchool	+= lpSchool->cName	;

			lpSchool	= lpSchool->pNextSchool	;
		}
		for ( i=iReadSchool; i<iMaxRow; i++ )
		{
			strHeadLineSchool	+= "\t ";
		}

		// �γ����� blocks
		int		iRow	;
		for ( int iBlock=0; iBlock<lpSubItem->m_iCntBlocks; iBlock++ )
		{
			lpSubItem->m_pPara->iCntRow		= lpSubItem->m_iMaxColOfBlock[iBlock]	;

			strBlock.Empty ( )	;

			// ������ɱ� block
			iCntCol	= 0	;
			while ( NULL != lpScoreRegion )
			{
				CWordFormatPrint *	pFormat	= (CWordFormatPrint *)lpSubItem->m_listlpColumn	;

				// �ȶ���ȫ��������
				lpScore				= lpScoreRegion	;
				lpScoreRegionHead	= lpScoreRegion	;	// �� block ������λ��
				for ( iRow=0; iRow<lpSubItem->m_iMaxSubrowOfRow[0]; iRow++ )
				{
					strScore.Format ( pFormat->m_cFormat, lpScore->fValue )	;

					if ( 0 == iRow )
						strBlock	+= strScore	;	
					else
						strBlock	+= "\t" + strScore	;

					pFormat		= (CWordFormatPrint *)pFormat->m_pNext		;
					lpScore		= lpScore->pNextScore						;

				}
				lpScoreRegion	= lpScore	;

				// ���뱾 block �и�У������
				LPSCORE_FORM	lpScore0	;
				int				j			;

				lpScore0	= lpScoreRegionHead	;
				for ( j=0; j<iReadSchool; j++ )
				{
					lpScore	= *(lppScore+j)	;

					// ����һ������
					for ( iRow=0; iRow<lpSubItem->m_iMaxSubrowOfRow[j+1]; iRow++ )
					{
						// Уƽ��
						strScore.Format ( pFormat->m_cFormat, lpScore->fValue )	;
						strBlock	+= "\t" + strScore	;	

						// ����
						strScore.Format ( pFormat->m_cFormat, lpScore->fValue - lpScore0->fValue )	;
						strBlock	+= "\t" + strScore	;

						pFormat		= (CWordFormatPrint *)pFormat->m_pNext		;
						lpScore		= lpScore->pNextScore						;
					}

					*(lppScore+j)	= lpScore	;
				}
				// ����
				for ( j=iReadSchool; j<iMaxRow; j++ )
				{
					for ( iRow=0; iRow<lpSubItem->m_iMaxSubrowOfRow[j+1]; iRow++ )
					{
						strBlock	+= "\t \t "	;	
					}
				}

				lpScore0	= lpScore0->pNextScore	;

				strBlock	+= "\r\n"	;

				iCntCol++	;	// ����һ��

				if ( iCntCol == lpSubItem->m_iMaxColOfBlock[iBlock] )
					break;

			}
			strBlock	+= "\t \r\n";	// ���һ�����һ��Ԫ��ӵ㶫������Ȼ�Բ���

			// ���� Title
			lpSubItem->MakeTitle ( &m_Sel )	;

			//+++++++++++++++++++++++++++++++++++++
			// copy and paste Headline
			// ��ͷ
			if ( !OpenClipboard(NULL) ) 
				AfxMessageBox ( "Failed open clipboard" ); 

			EmptyClipboard(); 

			hglbCopy	= GlobalAlloc(GMEM_MOVEABLE, strHeadLineSchool.GetLength( ) ); 
			if (hglbCopy == NULL) 
			{ 
				CloseClipboard(); 
				AfxMessageBox ( "Failed alloc mem" ); 
				return ; 
			} 

			lptstrCopy = (LPTSTR)GlobalLock(hglbCopy); 
			memcpy( lptstrCopy, strHeadLineSchool, strHeadLineSchool.GetLength() ); 
			//lptstrCopy[cch] = (TCHAR) 0;    // null character 
			GlobalUnlock(hglbCopy); 

			SetClipboardData(CF_TEXT, hglbCopy); 

			CloseClipboard(); 

			lpSubItem->SelectHead ( &m_Sel )	;
			m_Sel.Paste ( )	;

			//+++++++++++++++++++++++++++++++++++++
			// copy and paste this block
			if ( !OpenClipboard(NULL) ) 
				AfxMessageBox ( "Failed open clipboard" ); 

			EmptyClipboard(); 

			hglbCopy	= GlobalAlloc(GMEM_MOVEABLE, strBlock.GetLength( ) ); 
			if (hglbCopy == NULL) 
			{ 
				CloseClipboard(); 
				AfxMessageBox ( "Failed alloc mem" ); 
				return ; 
			} 

			lptstrCopy = (LPTSTR)GlobalLock(hglbCopy); 
			memcpy( lptstrCopy, strBlock, strBlock.GetLength() ); 
			//lptstrCopy[cch] = (TCHAR) 0;    // null character 
			GlobalUnlock(hglbCopy); 

			SetClipboardData(CF_TEXT, hglbCopy); 

			CloseClipboard(); 

			// ��λ����������һ��
			lpSubItem->LocateBody ( &m_Sel );
			//lpSubItem->ResetCell ( )	;
			lpSubItem->ShiftRow ( &m_Sel )	;
			m_Sel.Paste ( )	;

			//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
			//lpSubItem->m_pPara->iCntRow	-= iMaxRow	;

			//iCntBlock++	;
			//iMaxRow		= lpSubItem->m_iMaxRowOfBlock[iCntBlock];
			//					AfxMessageBox ( *pstrBlock )	;

			m_Sel.HomeKey ( COleVariant((short)6), COleVariant((short)0) );
			pParent->Activate ( )	;
			m_Sel.AttachDispatch( pApp->GetSelection() )	;
			m_Sel.WholeStory( )		;
			m_Sel.Copy ( )			;

			oActiveDoc.Activate ( )	;
			m_Sel.AttachDispatch( pApp->GetSelection() )	;

			// ����ĸ��
			m_Sel.Paste ( )	;
			m_Sel.HomeKey ( COleVariant((short)6), COleVariant((short)0) );

		}	// end of for iBlock

//		iCntSchool	-= iReadSchool	;
		iReadSchool	= iMaxRow - 1	;

		delete	lppSchool;
		delete	lppScore;

		m_pProgress->Step1 ( )	;
	}
	
	// ��������ķ������
	oActiveDoc.SaveAs( COleVariant( strPath + strType + strSubject + lpSubItem->cTitle ), 
		COleVariant((short)0), 
		vFalse, COleVariant(""), vTrue, COleVariant(""), 
		vFalse, vFalse, vFalse, vFalse, vFalse );
	oActiveDoc.Close ( vFalse, COleVariant((short)0), COleVariant((short)0) );

	//���ý�����
	m_pProgress->SetPos1( 100 );
	m_pProgress->SetPos2( 100 );

	// ���ٻ�����
	LPSCHOOL_FORM	lpSchoolPrev;
	LPSCORE_FORM	lpScorePrev	;
	lpSchool	= lpRegion->lpSchoolLst;
	while ( NULL != lpSchool )
	{
		// ���ٰ༶�б�
		LPCLASS_FORM	lpClass		;
		LPCLASS_FORM	lpClassPrev	;
		lpClass	= lpSchool->lpClassLst	;
		while ( NULL != lpClass )
		{
			// ���ٳɼ��б�
			lpScore	= lpClass->lpScoreLst	;
			while ( NULL != lpScore )
			{
				lpScorePrev	= lpScore				;
				lpScore		= lpScore->pNextScore	;
				delete	lpScorePrev	;
			}

			lpClassPrev	= lpClass				;
			lpClass		= lpClass->pNextClass	;
			delete	lpClassPrev	;
		}
		lpScore	= lpSchool->lpScoreLst	;
		while ( NULL != lpScore )
		{
			lpScorePrev	= lpScore				;
			lpScore		= lpScore->pNextScore	;
			delete	lpScorePrev	;
		}

		lpSchoolPrev= lpSchool				;
		lpSchool	= lpSchool->pNextSchool	;
		delete	lpSchoolPrev	;
	}
	lpScore	= lpRegion->lpScoreLst	;
	while ( NULL != lpScore )
	{
		lpScorePrev	= lpScore				;
		lpScore		= lpScore->pNextScore	;
		delete	lpScorePrev	;
	}

	delete	lpRegion;

	// ���ٽ�����
	m_pProgress->DestroyWindow();
}

BOOL CAnalyzer::GenerateChoice ( )
{
	CString		strSQL		;
	CDatabase *	pDatabase	;

	pDatabase	= &theApp.m_database;
	try
	{
		// ɾ���ɵ���ͼ vChoiceCnt
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vChoiceCnt]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vChoiceCnt]" );
		pDatabase->ExecuteSQL( strSQL );

		// ������ͼ vChoiceCnt
		strSQL.Format( "CREATE VIEW dbo.vChoiceCnt\
					   AS\
					   SELECT TOP 100 PERCENT dbo.ChoiceCnt.ClassID, \
					   dbo.QuestionFormInfo.QuestionFormID, dbo.ChoiceCnt.Choice, \
					   dbo.ChoiceCnt.Cnt / dbo.ChoiceCnt.CntStudent AS percentage, \
					   dbo.QuestionFormInfo.SubjectID, dbo.Class.SchoolID, dbo.Class.ClassName, \
					   dbo.Class.TypeID\
					   FROM dbo.ChoiceCnt INNER JOIN\
					   dbo.QuestionFormInfo ON \
					   dbo.ChoiceCnt.QuestionID = dbo.QuestionFormInfo.QuestionID AND \
					   dbo.ChoiceCnt.Step = dbo.QuestionFormInfo.Step INNER JOIN\
					   dbo.Class ON dbo.ChoiceCnt.ClassID = dbo.Class.ClassID\
					   WHERE (dbo.QuestionFormInfo.Extend = 0)\
					   ORDER BY dbo.QuestionFormInfo.SubjectID, dbo.ChoiceCnt.ClassID" );
		pDatabase->ExecuteSQL( strSQL );

		// ɾ���ɵ���ͼ vChoiceCntSchool
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vChoiceCntSchool]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vChoiceCntSchool]" );
		pDatabase->ExecuteSQL( strSQL );

		// ������ͼ vChoiceCntSchool
		strSQL.Format( "CREATE VIEW dbo.vChoiceCntSchool\
					   AS\
					   SELECT dbo.Class.SchoolID, dbo.ChoiceCnt.QuestionID, dbo.ChoiceCnt.Step, \
					   dbo.ChoiceCnt.Choice, SUM(dbo.ChoiceCnt.Cnt) AS Cnt, \
					   SUM(dbo.ChoiceCnt.CntStudent) AS CntStudent, dbo.Class.TypeID\
					   FROM dbo.ChoiceCnt INNER JOIN\
					   dbo.Class ON dbo.ChoiceCnt.ClassID = dbo.Class.ClassID\
					   GROUP BY dbo.Class.SchoolID, dbo.ChoiceCnt.QuestionID, dbo.ChoiceCnt.Step, \
					   dbo.ChoiceCnt.Choice, dbo.Class.TypeID" );
		pDatabase->ExecuteSQL( strSQL );

		// ɾ���ɵ���ͼ vChoicePercentageSchool
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vChoicePercentageSchool]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vChoicePercentageSchool]" );
		pDatabase->ExecuteSQL( strSQL );

		// ������ͼ vChoicePercentageSchool
		strSQL.Format( "CREATE VIEW dbo.vChoicePercentageSchool\
					   AS\
					   SELECT TOP 100 PERCENT dbo.vChoiceCntSchool.SchoolID, \
					   dbo.QuestionFormInfo.QuestionFormID, dbo.vChoiceCntSchool.Choice, \
					   dbo.vChoiceCntSchool.Cnt / dbo.vChoiceCntSchool.CntStudent AS Percentage, \
					   dbo.QuestionFormInfo.SubjectID, - 1 AS SchoolIDD, dbo.School.SchoolName, \
					   dbo.vChoiceCntSchool.TypeID\
					   FROM dbo.vChoiceCntSchool INNER JOIN\
					   dbo.QuestionFormInfo ON \
					   dbo.vChoiceCntSchool.QuestionID = dbo.QuestionFormInfo.QuestionID AND \
					   dbo.vChoiceCntSchool.Step = dbo.QuestionFormInfo.Step INNER JOIN\
					   dbo.School ON dbo.vChoiceCntSchool.SchoolID = dbo.School.SchoolID\
					   WHERE (dbo.QuestionFormInfo.Extend = 0)\
					   ORDER BY dbo.QuestionFormInfo.SubjectID, dbo.vChoiceCntSchool.SchoolID" );
		pDatabase->ExecuteSQL( strSQL );

		// ɾ���ɵ���ͼ vChoiceCntRegion
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vChoiceCntRegion]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vChoiceCntRegion]" );
		pDatabase->ExecuteSQL( strSQL );

		// ������ͼ vChoiceCntRegion
		strSQL.Format( "CREATE VIEW dbo.vChoiceCntRegion\
					   AS\
					   SELECT - 1 AS SchoolID, QuestionID, Step, Choice, SUM(Cnt) AS Cnt, SUM(CntStudent) \
					   AS CntStudent, TypeID\
					   FROM dbo.vChoiceCntSchool\
					   GROUP BY QuestionID, Step, Choice, TypeID" );
		pDatabase->ExecuteSQL( strSQL );

		// ɾ���ɵ���ͼ vChoicePercentageRegion
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vChoicePercentageRegion]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vChoicePercentageRegion]" );
		pDatabase->ExecuteSQL( strSQL );

		// ������ͼ vChoicePercentageRegion
		strSQL.Format( "CREATE VIEW dbo.vChoicePercentageRegion\
					   AS\
					   SELECT TOP 100 PERCENT dbo.vChoiceCntRegion.SchoolID, \
					   dbo.QuestionFormInfo.QuestionFormID, dbo.vChoiceCntRegion.Choice, \
					   dbo.vChoiceCntRegion.Cnt / dbo.vChoiceCntRegion.CntStudent AS Percentage, \
					   dbo.QuestionFormInfo.SubjectID, - 1 AS SchoolIDD, 'ȫ��' AS SchoolName, \
					   dbo.vChoiceCntRegion.TypeID\
					   FROM dbo.vChoiceCntRegion INNER JOIN\
					   dbo.QuestionFormInfo ON \
					   dbo.vChoiceCntRegion.QuestionID = dbo.QuestionFormInfo.QuestionID AND \
					   dbo.vChoiceCntRegion.Step = dbo.QuestionFormInfo.Step\
					   WHERE (dbo.QuestionFormInfo.Extend = 0)\
					   ORDER BY dbo.QuestionFormInfo.SubjectID, dbo.vChoiceCntRegion.SchoolID" );
		pDatabase->ExecuteSQL( strSQL );

		//// ɾ���ɵı� 
		//strSQL.Format( "" );
		//pDatabase->ExecuteSQL( strSQL );

		//// ���ɱ� 
		//strSQL.Format( "" );
		//pDatabase->ExecuteSQL( strSQL );

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

void CAnalyzer::DestroyChoice ( )
{
	CString		strSQL		;
	CDatabase *	pDatabase	;

	pDatabase	= &theApp.m_database;
	try
	{
		// ɾ����ͼ vChoiceCnt
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vChoiceCnt]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vChoiceCnt]" );
		pDatabase->ExecuteSQL( strSQL );

		// ɾ����ͼ vChoiceCntSchool
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vChoiceCntSchool]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vChoiceCntSchool]" );
		pDatabase->ExecuteSQL( strSQL );

		// ɾ����ͼ vChoicePercentageSchool
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vChoicePercentageSchool]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vChoicePercentageSchool]" );
		pDatabase->ExecuteSQL( strSQL );

		// ɾ����ͼ vChoiceCntRegion
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vChoiceCntRegion]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vChoiceCntRegion]" );
		pDatabase->ExecuteSQL( strSQL );

		// ɾ����ͼ vChoicePercentageRegion
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vChoicePercentageRegion]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vChoicePercentageRegion]" );
		pDatabase->ExecuteSQL( strSQL );

		//// ɾ���ɵı� 
		//strSQL.Format( "" );
		//pDatabase->ExecuteSQL( strSQL );

		//// ���ɱ� 
		//strSQL.Format( "" );
		//pDatabase->ExecuteSQL( strSQL );

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
		strError.Format( "ɾ����ʱ���ݿ����ʧ��\r\n" );
		strError	+= e->m_strError;
		AfxMessageBox( strError );
		return;
	}
	
	return;
}
