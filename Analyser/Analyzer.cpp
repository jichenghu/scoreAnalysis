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

	//进度条
	m_pProgress	= new CProgress(NULL);
	m_pProgress->SetTitle1 ( "分析原始成绩数据" )	;

	try
	{
		setChoiceInfo.Open ( )	;
		setChoiceCnt.Open ( )	;

		// 如果数据库表 ChoiceCnt 中已经有数据则先发出警告
		if ( !setChoiceCnt.IsEOF() )
		{
			int	iDlg	= MessageBox ( NULL, "发现先前的选择题分析数据！\n是否删除先前的数据？", "警告", MB_YESNOCANCEL )	;
			if ( IDYES == iDlg )
			{
				theApp.m_database.ExecuteSQL ( "DELETE FROM ChoiceCnt" )	;
			}
			else if ( IDNO == iDlg )
			{
				setChoiceInfo.Close ( )	;
				setChoiceCnt.Close ( )	;
					
				m_pProgress->DestroyWindow();

				AfxMessageBox ( "由于没有选择删除先前的数据将无法处理" )	;
				
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
					// 保存该班的统计结果
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

				// 初始化各统计量
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

			// 读文件

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

		// 保存最后一个班的统计结果
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

	// 销毁进度条
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

		// 删除子串
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
	// 根据 lpSubItem 来读入格式信息
	COleVariant		vTrue( (short)TRUE )						;
	COleVariant		vFalse( (short)FALSE )						;
	COleVariant		vOpt( (long)DISP_E_PARAMNOTFOUND, VT_ERROR );

	_Application	m_App	;	// 定义Word提供的应用程序对象
	Documents		m_Docs	;	// 定义Word提供的文档对象
	Selection		m_Sel	;	// 定义Word提供的选择对象

	m_Docs.ReleaseDispatch();
	m_Sel.ReleaseDispatch()	;

	m_App.m_bAutoRelease	= true;
	if( !m_App.CreateDispatch("Word.Application") )
	{ 
		AfxMessageBox( "创建Word2000服务失败!" ); 
		exit( 1 ); 
	}

	CString	strTemplate	;

	strTemplate.Format ( lpSubItem->cPath )	;
	//strTemplate.Format( "c:\\分析\\知识点" )	;
	//strType.TrimRight()						;
	//strTemplate	+= strType					;
	//strTemplate	+= strSubject.Left( 4 )		;

	// 下面是定义VARIANT变量
	COleVariant	varFilePath( strTemplate )		;
	COleVariant	varstrNull( "" )				;
	COleVariant varZero( (short)0 )				;
	COleVariant varTrue( short(1), VT_BOOL )	;
	COleVariant varFalse( short(0), VT_BOOL )	;

	m_Docs.AttachDispatch( m_App.GetDocuments() )	;	// 将Documents类对象m_Docs和Idispatch接口关联起来
	m_Docs.Open( varFilePath, 
		varFalse, varFalse, varFalse,
		varstrNull, varstrNull, varFalse, varstrNull,
		varstrNull, varTrue, varTrue, varTrue )		;
	//m_App.SetVisible ( FALSE );
	m_App.SetVisible ( TRUE );
	
	// 打开Word文档
	m_Sel.AttachDispatch( m_App.GetSelection() )	;	// 将Selection类对象m_Sel和Idispatch接口关联起来

	//// 全选复制
	//m_Sel.WholeStory( )	;
	//m_Sel.Copy ( )		;

	// 填写表格
	// 写区名
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
	case	0:	// 区
		break	;

	case	1:	// 区校
		break	;

	case	2:	// 校
		lpSubItem->m_pPara->strNameRegion	= &strRegion;
		//FillFormSchool ( iSubjectID, iRegionID, &m_App, &m_Docs, &oActiveDoc, &m_Sel, lpSubItem )	;
		break	;

	case	3:	// 校班
		lpSubItem->m_pPara->strNameRegion	= &strRegion;
		FillFormSchoolClass ( iSubjectID, strSubject, iRegionID, strRegion, iTypeID, strType, &m_App, &m_Docs, &oActiveDoc, &m_Sel, lpSubItem )	;
		break	;

	case	4:	// 班
		lpSubItem->m_pPara->strNameRegion	= &strRegion;
		FillFormClass ( iSubjectID, strSubject, iRegionID, strRegion, iTypeID, strType, &m_App, &m_Docs, &oActiveDoc, &m_Sel, lpSubItem )	;
		break	;

	case	5:	// 区校班
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
	
	m_Docs.ReleaseDispatch();	// 断开关联；
	m_Sel.ReleaseDispatch()	;

	// 退出WORD 
	m_App.Quit(vOpt, vOpt, vOpt); 
//	m_App.Quit(vOpt, vOpt, vOpt);
	m_App.ReleaseDispatch();
}

//++++++++++++++++++++++++++++++++++++++++++++++++++
//	按校班级产生分析报表
//		每个学校每科一个文件
//
//	第一步: 按学校进行循环产生Word文档, 每个学校一个文件
//	第二步: 将 VirtualForm 中的数据填入模板表
//	第三步: 将第二步中的表拷贝至目标表的尾巴上
//--------------------------------------------------
void CAnalyzer::FillFormSchoolClass ( int iSubjectID, CString strSubject,  
			int iRegionID, CString strRegion, int iTypeID, CString strType,
			_Application* pApp, Documents* pDocs, _Document* pParent, Selection* pSel, 
			LPSUBITEM lpSubItem )
{
	// 数据库相关参数
	CString	strFilter	;
	CString	strFilter0	;
	CString	strFilter1	;
	CString	strFilter2	;
	CString	strSort		;
	CString	strSchool	;
	CString	strSQL		;

	// 数据库
	//CSetSelectID		setIDSchool				;
	CvChoicePercentage	*lpvChoicePercentage	;
	CvChoicePercentage	*lpvChoicePercentage1	;
	CvChoicePercentage	*lpvChoicePercentage2	;

	// 报表相关参数
	CString	strPath		;	// 报表的存放路径

	// 存放分析数据的 Buffer
	LPREGION_FORM	lpRegion	;	// 区
	LPSCHOOL_FORM	lpSchoolHead;
	LPSCHOOL_FORM	lpSchoolTail;
	LPSCHOOL_FORM	lpSchool	;
	LPSCORE_FORM	lpScoreHead	;
	LPSCORE_FORM	lpScoreTail	;
	LPSCORE_FORM	lpScore		;

	// Word 的相关参数
	Selection	m_Sel							;	// 定义Word提供的选择对象
	COleVariant varTrue( short(1), VT_BOOL )	;
	COleVariant varFalse( short(0), VT_BOOL )	;
	COleVariant	vTrue( (short)TRUE )			;
	COleVariant	vFalse( (short)FALSE )			;

	// 常用变量
	CString			strID		;
	CString			strName		;
	CString			strScore	;
	CString			strBlock	;

	// 初始化变量
	m_pProgress	= new CProgress(NULL)	;	// 进度条
	lpRegion	= new REGION_FORM		;	// 每次只分析一个区

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

	
	// 首先产生相关的视图
	if ( !GenerateChoice() )
	{
		AfxMessageBox ( "产生相关的视图失败！\n分析结果未能产生" )	;
		return	;
	}

//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//	收集信息
	strFilter0.Empty ( )	;
	strFilter1.Empty ( )	;
	strFilter2.Empty ( )	;
	m_pProgress->SetTitle2 ( "收集信息" )	;
	try
	{
		//+++++++++++++++++++++++++++
		// 先读入区的数据
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

		// 读入成绩分析数据
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

		// 区数据读入完毕, 清理资源
		lpvChoicePercentage->Close ( )	;
		delete	lpvChoicePercentage		;

		//+++++++++++++++++++++++++++
		// 再读入区中各学校的数据
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

			// 读入数据
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

				// 如果是新的学校
				if ( iID != lpvChoicePercentage1->m_ID )
				{
					iCnt++	;	// 学校数目加1
					iID	= lpvChoicePercentage1->m_ID;	// 新 iID

					// 新建一个学校
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

			// 学校数据读入完毕, 清理资源
			lpvChoicePercentage1->Close ( )	;
			delete	lpvChoicePercentage1	;

			// 将学校挂接到区中
			lpRegion->lpSchoolLst	= lpSchoolHead	;
			lpRegion->iCntSchool	= iCnt			;
		}

		//strSQL.Format( "SELECT schoolID, schoolName FROM School WHERE (RegionID = %d) ORDER BY schoolID", iRegionID )	;
		//setIDSchool.Open ( CRecordset::snapshot, strSQL, CRecordset::readOnly )	;

		//+++++++++++++++++++++++++++
		// 次读入区中各校中各班的分析数据
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

			int		i	= 0				;	// 进度
			lpSchool	= lpSchoolHead	;
			while ( NULL != lpSchool )
			{
				// 取得该学校相关数据
				iSchoolID	= lpSchool->iSchoolID			;
				strSchool.Format ( "%s", lpSchool->cName )	;
				strSchool.TrimLeft ()	;	strSchool.TrimRight ()	;

				// 设置进度标志
				m_pProgress->SetTitle2 ( strSchool );

				// 获取新的数据库记录集
				strFilter.Format ( "(SubjectID = %d) AND (schoolID = %d) AND (TypeID = %d)", iSubjectID, iSchoolID, iTypeID )	;/**/
				lpvChoicePercentage2->m_strFilter	= strFilter2 + strFilter;
				lpvChoicePercentage2->Requery ( )	;

				// 读数据库记录集
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
						iCnt++	;	// 班级数加1
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

					// 移动到下一记录
					lpvChoicePercentage2->MoveNext ( )	;
				}

				// 将班级列表连接到学校
				lpSchool->lpClassLst= lpClassHead	;
				lpSchool->iCntClass	= iCnt			;

				// 处理下一学校
				lpSchool	= lpSchool->pNextSchool	;

				// 设置进度
				m_pProgress->SetPos2 ( ++i * 100 / lpRegion->iCntSchool )	;
			}

			// 各校的各班的数据读入完毕, 清理资源
			lpvChoicePercentage2->Close ( )	;
			delete	lpvChoicePercentage2	;
		}

/*		lpSchoolHead	= NULL	;

		while ( !setIDSchool.IsEOF() )
		{
			setIDSchool.MoveNext ( )	;
		}
		iCntSchool	= setIDSchool.GetRecordCount ( )	;

		// 按学校循环
		if ( iCntSchool )
			setIDSchool.MoveFirst ( )	;
		for ( int i=0; i<iCntSchool; i++ )	// for school
		{
			iSchoolID	= setIDSchool.m_iID	;
			strSchool.Format ( "%s", setIDSchool.m_strName )	;
			strSchool.TrimRight ( )		;

			m_pProgress->SetTitle1 ( strSchool )	;

			// 读数据库
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
					iCnt++	;	// 班级数加1
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

			// 如果本学校有成绩分析出现, 则将本学校加入到学校列表中
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
//	生成报表
	m_pProgress->SetTitle1 ( "生成报表" )	;
	m_pProgress->SetPos1( 0 );
	m_pProgress->SetPos2( 0 );

	LPCLASS_FORM	lpClass		;
	LPTSTR			lptstrCopy	; 
	HGLOBAL			hglbCopy	; 

	// 每个区只生成一个 Word 文档
	_Document	oActiveDoc	;
	pDocs->Add ( COleVariant(pParent->GetFullName()), varFalse, COleVariant((short)0), varTrue )	;
	oActiveDoc = pApp->GetActiveDocument(); 
	m_Sel.AttachDispatch( pApp->GetSelection() )	;	// 将Selection类对象m_Sel和Idispatch接口关联起来

	// 按学校循环
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
		int	iReadClass	;	// 当前需要读入的班级数
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
			
		// 设置参数
		lpSubItem->m_pPara->strNameSchool	= &strSchool			;
		//lpSubItem->m_pPara->iCntRow			= lpSchool->iCntClass	;
		//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

		m_pProgress->SetTitle2 ( strSchool );
		m_pProgress->SetPos1 ( 0 )			;

		// 计算第一次读入的班级个数
		iReadClass	= iCntClass % (iMaxRow - 1 );	// 留一列添全校的统计数据
		if ( 0 == iReadClass )
			iReadClass	= iMaxRow - 1	;	// 如果正好是最后1张也填满的情况

		// 读入各班的数据
		lpSubItem->ResetRow ( )			;
		lpClass	= lpSchool->lpClassLst	;
		while ( NULL != lpClass )
		{
			// 一次读入数个班的记录进来
			LPCLASS_FORM	*lppClass		;
			LPSCORE_FORM	*lppScore		;	// 存放各班的数据
			LPSCORE_FORM	lpScoreSchool	;	// 存放学校的数据

			lppClass	= new LPCLASS_FORM[iReadClass]	;
			lppScore	= new LPSCORE_FORM[iReadClass]	;
		
			strHeadLineClass.Empty ( )		;

			// 校数据
			lpScoreSchool	= lpSchool->lpScoreLst	;

			strHeadLineClass	+= lpSchool->cName	;

			// 班数据 
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

			// 形成若干 blocks
			int		iRow	;
			for ( int iBlock=0; iBlock<lpSubItem->m_iCntBlocks; iBlock++ )
			{
				lpSubItem->m_pPara->iCntRow		= lpSubItem->m_iMaxColOfBlock[iBlock]	;

				strBlock.Empty ( )	;

				// 逐列完成本 block
				iCntCol	= 0	;
				while ( NULL != lpScoreSchool )
				{
					CWordFormatPrint *	pFormat	= (CWordFormatPrint *)lpSubItem->m_listlpColumn	;

					// 先读入全校的数据
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

					// 读入本 block 中各班级的数据
					int	j;
					for ( j=0; j<iReadClass; j++ )
					{
						lpScore	= *(lppScore+j)	;

						// 读出一组数据
						for ( iRow=0; iRow<lpSubItem->m_iMaxSubrowOfRow[j+1]; iRow++ )
						{
							strScore.Format ( pFormat->m_cFormat, 100*lpScore->fValue )	;

							strBlock	+= "\t" + strScore + "%"	;	

							pFormat		= (CWordFormatPrint *)pFormat->m_pNext		;
							lpScore		= lpScore->pNextScore						;
						}

						*(lppScore+j)	= lpScore	;
					}
					// 补空
					for ( j=iReadClass; j<iMaxRow-1; j++ )
					{
						for ( iRow=0; iRow<lpSubItem->m_iMaxSubrowOfRow[j+1]; iRow++ )
						{
							strBlock	+= "\t "	;	
						}
					}

					strBlock	+= "\r\n"	;

					iCntCol++	;	// 增加一列

					if ( iCntCol == lpSubItem->m_iMaxColOfBlock[iBlock] )
						break;

				}

				// 生成 Title
				lpSubItem->MakeTitle ( &m_Sel )	;


				//+++++++++++++++++++++++++++++++++++++
				// copy and paste Headline
				// 表头
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

				// 定位到表的主体第一行
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

				// 复制母表
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

		// 回到起点
		m_Sel.HomeKey ( COleVariant((short)6), COleVariant((short)0) );

		m_pProgress->SetPos2( 100 * ii++ / lpRegion->iCntSchool );
	
	}	// end of while lpSchool

	if ( 0 == iMaxRow )
	{
		AfxMessageBox ( "未发现有效数据!\r\n请检查数据库是否正确设置:\r\n    QuestionFormInfo" );

		// 销毁缓冲区
		LPSCHOOL_FORM	lpSchoolPrev;
		LPSCORE_FORM	lpScorePrev	;
		lpSchool	= lpRegion->lpSchoolLst;
		while ( NULL != lpSchool )
		{
			// 销毁班级列表
			LPCLASS_FORM	lpClass		;
			LPCLASS_FORM	lpClassPrev	;
			lpClass	= lpSchool->lpClassLst	;
			while ( NULL != lpClass )
			{
				// 销毁成绩列表
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

		// 销毁进度条
		m_pProgress->DestroyWindow();

		return;
	}

	// 填区校表
	int		iReadSchool;

	// 计算第一次读入学校的个数
	iReadSchool	= lpRegion->iCntSchool % (iMaxRow - 1 );	// 留一列添全区的统计数据
	if ( 0 == iReadSchool )
		iReadSchool	= iMaxRow - 1	;	// 如果正好是最后1张也填满的情况

	iMaxRow		= lpSubItem->m_iMaxRowOfBlock[0];

	CString	strEmpty( " " )	;
	lpSubItem->m_pPara->strNameSchool	= &strEmpty	;

	// 读入各校的数据
	lpSubItem->ResetRow ( )		;
	lpSchool	= lpSchoolHead	;
	while ( NULL != lpSchool )
	{
		// 一次读入数个学校的记录进来
		LPSCHOOL_FORM	*lppSchool		;
		LPSCORE_FORM	*lppScore		;	// 存放各校的数据
		LPSCORE_FORM	lpScoreRegion	;	// 存放区的数据

		lppSchool	= new LPSCHOOL_FORM[iReadSchool];
		lppScore	= new LPSCORE_FORM[iReadSchool]	;

		strHeadLineSchool.Empty ( )		;

		// 区数据
		lpScoreRegion	= lpRegion->lpScoreLst	;

		strHeadLineSchool	+= lpRegion->cName	;

		// 校数据 
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

		// 形成若干 blocks
		int		iRow	;
		for ( int iBlock=0; iBlock<lpSubItem->m_iCntBlocks; iBlock++ )
		{
			lpSubItem->m_pPara->iCntRow		= lpSubItem->m_iMaxColOfBlock[iBlock]	;

			strBlock.Empty ( )	;

			// 逐列完成本 block
			iCntCol	= 0	;
			while ( NULL != lpScoreRegion )
			{
				CWordFormatPrint *	pFormat	= (CWordFormatPrint *)lpSubItem->m_listlpColumn	;

				// 先读入全区的数据
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

				// 读入本 block 中各班级的数据
				int	j;
				for ( j=0; j<iReadSchool; j++ )
				{
					lpScore	= *(lppScore+j)	;

					// 读出一组数据
					for ( iRow=0; iRow<lpSubItem->m_iMaxSubrowOfRow[j+1]; iRow++ )
					{
						strScore.Format ( pFormat->m_cFormat, 100*lpScore->fValue )	;

						strBlock	+= "\t" + strScore + "%"	;	

						pFormat		= (CWordFormatPrint *)pFormat->m_pNext		;
						lpScore		= lpScore->pNextScore						;
					}

					*(lppScore+j)	= lpScore	;
				}
				// 补空
				for ( j=iReadSchool; j<iMaxRow-1; j++ )
				{
					for ( iRow=0; iRow<lpSubItem->m_iMaxSubrowOfRow[j+1]; iRow++ )
					{
						strBlock	+= "\t "	;	
					}
				}

				strBlock	+= "\r\n"	;

				iCntCol++	;	// 增加一列

				if ( iCntCol == lpSubItem->m_iMaxColOfBlock[iBlock] )
					break;

			}

			// 生成 Title
			lpSubItem->MakeTitle ( &m_Sel )	;

			//+++++++++++++++++++++++++++++++++++++
			// copy and paste Headline
			// 表头
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

			// 定位到表的主体第一行
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

			// 复制母表
			m_Sel.Paste ( )	;
			m_Sel.HomeKey ( COleVariant((short)6), COleVariant((short)0) );

		}	// end of for iBlock

//		iCntSchool	-= iReadSchool	;
		iReadSchool	= iMaxRow - 1	;

		delete	lppSchool;
		delete	lppScore;

		m_pProgress->Step1 ( )	;
	}
	
	// 保存该区的分析结果
	oActiveDoc.SaveAs( COleVariant( strPath + strType + strSubject + lpSubItem->cTitle ), 
		COleVariant((short)0), 
		vFalse, COleVariant(""), vTrue, COleVariant(""), 
		vFalse, vFalse, vFalse, vFalse, vFalse );
	oActiveDoc.Close ( vFalse, COleVariant((short)0), COleVariant((short)0) );

	//设置进度条
	m_pProgress->SetPos1( 100 );
	m_pProgress->SetPos2( 100 );

	// 销毁缓冲区
	LPSCHOOL_FORM	lpSchoolPrev;
	LPSCORE_FORM	lpScorePrev	;
	lpSchool	= lpRegion->lpSchoolLst;
	while ( NULL != lpSchool )
	{
		// 销毁班级列表
		LPCLASS_FORM	lpClass		;
		LPCLASS_FORM	lpClassPrev	;
		lpClass	= lpSchool->lpClassLst	;
		while ( NULL != lpClass )
		{
			// 销毁成绩列表
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

	// 销毁进度条
	m_pProgress->DestroyWindow();
}

//++++++++++++++++++++++++++++++++++++++++++++++++++
//	按班级产生分析报表
//		每个学校每科每类型一个文件
//
//	第一步: 按学校进行循环产生Word文档, 每个学校一个文件
//	第二步: 将 VirtualForm 中的数据填入模板表
//	第三步: 将第二步中的表拷贝至目标表的尾巴上
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

	// 存贮报表结果的路径
	//strPath.Format ( "C:\\Forms\\" )	;
	strPath	= m_strPathOutput		;
	SetCurrentDirectory ( strPath )	;

	strPath	= strPath + "\\"		;
	strPath	= strPath + strRegion	;
	if ( !TestDirectory( strPath ) )
		_mkdir ( strRegion )	;
	SetCurrentDirectory ( strPath )	;
	strPath	= strPath + "\\"		;

	Selection		m_Sel	;	// 定义Word提供的选择对象

	COleVariant varTrue( short(1), VT_BOOL )	;
	COleVariant varFalse( short(0), VT_BOOL )	;
	COleVariant	vTrue( (short)TRUE )			;
	COleVariant	vFalse( (short)FALSE )			;

	CSetSelectID		setIDSchool	;
	CSetSelectID		setIDClass	;

	CvFactStudent *		lpvFactStudent	;
	CvScoreForm *		lpvFactStudent1	;
	CvScoreForm *		lpvFactStudent2	;

	//进度条
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

//	收集信息
	strFilter0.Empty ( )	;
	strFilter1.Empty ( )	;
	strFilter2.Empty ( )	;
	m_pProgress->SetTitle1 ( "收集信息" )	;
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

		// 按学校循环
		if ( iCntSchool )
			setIDSchool.MoveFirst ( )	;
		for ( int i=0; i<iCntSchool; i++ )	// for school
		{
			iSchoolID	= setIDSchool.m_iID	;
			strSchool.Format ( "%s", setIDSchool.m_strName )	;
			strSchool.TrimRight ( )		;

			m_pProgress->SetTitle2 ( strSchool )	;

			// 按班级循环
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
				// 设置参数
				//lpSubItem->m_pPara->strNameClass	= &setIDClass.m_strName	;

				// 读数据库
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
						iCnt++	;	// 学生人数加1
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

				// 如果本班级有学生成绩出现, 则将本班级加入到班级列表中
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

					// 可能的话读入次表和辅表
					// 首先读入次表
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
								// 找到该 ID 所对应的学生
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

							// 将数据加入到列表中
							if ( NULL != lpStudent )
							{
								lpScoreTail->pNextScore	= lpScore	;
								lpScoreTail				= lpScore	;
							}

							lpvFactStudent1->MoveNext ( )	;
						}
					}
					// 然后读入辅表
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
								// 找到该 ID 所对应的学生
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

							// 将数据加入到列表中
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

//	生成报表
	// 按学校循环
	m_pProgress->SetTitle1 ( "生成报表" )	;
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

		// 设置参数
		lpSubItem->m_pPara->strNameSchool	= &strSchool	;

		m_pProgress->SetTitle2 ( strSchool )	;

		//m_Sel.AttachDispatch( pApp->GetSelection() )	;	// 将Selection类对象m_Sel和Idispatch接口关联起来
		//// 全选复制
		//m_Sel.WholeStory( )	;
		//m_Sel.Copy ( )		;

		// 每个学校均生成一个新的 Word 文档
		pDocs->Add ( COleVariant(pParent->GetFullName()), varFalse, COleVariant((short)0), varTrue )	;
		oActiveDoc = pApp->GetActiveDocument(); 
		m_Sel.AttachDispatch( pApp->GetSelection() )	;	// 将Selection类对象m_Sel和Idispatch接口关联起来

		// 按班级循环
		LPCLASS_FORM	lpClass		;
		lpClass	= lpSchool->lpClassLst	;
		while ( NULL != lpClass )
		{
			CString	strNameClass	;

			// 设置参数
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

			//// 复制母表
			//m_Sel.Paste ( )	;

			//// 生成 Title
			//lpSubItem->MakeTitle ( &m_Sel )	;

			//// 定位到表的主体第一行
			//lpSubItem->LocateBody ( &m_Sel );

			// 生成字符串列表
			LPSTUDENT_FORM	lpStudent	;
			CString			strID		;
			CString			strName		;
			CString			strScore	;
			CString			strBlock	;

			int				iReadStudent;	// 需要读入的学生人数
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

			// 计算第一次读入的人数
			if ( iMaxRow == iCntStudent )
				iReadStudent	= iMaxRow	;	// 正好是一张
			else
				iReadStudent	= iCntStudent % iMaxRow	;

			// 读入各学生的成绩数据
			//lpSubItem->ResetRow ( )	;
			lpStudent	= lpClass->lpStudentLst	;
			while ( NULL != lpStudent )
			{
				// 一次读入数个学生的记录进来
				strBlock.Empty ( )	;

				int		i	;
				for ( i=0; (i<iReadStudent) && (NULL!=lpStudent); i++ )
				{
					// 先填写考号和姓名
					strID.Format ( "%s", lpStudent->cEnrollID )		;
					strBlock	+= strID							;	// &lpvFactStudent->m_StudentEnrollID	;
					strName.Format ( "\t%s", lpStudent->cName )		;
					strBlock	+= strName							;	// &lpvFactStudent->m_StudentEnrollID	;
					// 读入本学生的成绩数据
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
				// 补空
				for ( int j=iReadStudent; j<iMaxRow; j++ )
				{
					//for ( iRow=0; iRow<lpSubItem->m_iMaxSubrowOfRow[j+1]; iRow++ )
					//{
						strBlock	+= "\t \r\n"	;	
					//}
				}
				
				strBlock	+= "\t \r\n";	// 最后一行最后一单元后加点东西，不然对不齐

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

				// 生成 Title
				lpSubItem->MakeTitle ( &m_Sel )	;

				// 定位到表的主体第一行
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

				// 复制母表
				m_Sel.Paste ( )	;
				m_Sel.HomeKey ( COleVariant((short)6), COleVariant((short)0) );

				lpSubItem->m_pPara->iCntRow	-= iReadStudent	;

				iReadStudent	= iMaxRow	;

				//lpStudent		= lpStudent->pNextStudent	;
			}

			lpClass		= lpClass->pNextClass	;

			// 回到起点
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

	//设置进度条
	m_pProgress->SetPos1( 100 );
	m_pProgress->SetPos2( 100 );

	// 销毁进度条
	m_pProgress->DestroyWindow();

	// 销毁信息列表
	LPSCHOOL_FORM	lpSchoolPrev;
	lpSchool	= lpSchoolHead	;
	while ( NULL != lpSchool )
	{
		// 销毁班级列表
		LPCLASS_FORM	lpClass		;
		LPCLASS_FORM	lpClassPrev	;
		lpClass	= lpSchool->lpClassLst	;
		while ( NULL != lpClass )
		{
			// 销毁学生列表
			LPSTUDENT_FORM	lpStudent		;
			LPSTUDENT_FORM	lpStudentPrev	;
			lpStudent	= lpClass->lpStudentLst	;
			while ( NULL != lpStudent )
			{
				// 销毁成绩列表
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
//	按区校班产生分析报表
//		每科一个文件
//
//	第一步: 
//	第二步: 
//	第三步: 
//--------------------------------------------------
void CAnalyzer::FillFormRegionSchoolClass ( int iSubjectID, CString strSubject,  
			int iRegionID, CString strRegion, int iTypeID, CString strType,
			_Application* pApp, Documents* pDocs, _Document* pParent, Selection* pSel, 
			LPSUBITEM lpSubItem )
{
	// 数据库相关参数
	CString	strFilter	;
	CString	strFilter0	;
	CString	strFilter1	;
	CString	strFilter2	;
	CString	strSort		;
	CString	strSchool	;
	CString	strSQL		;

	// 数据库
	//CSetSelectID		setIDSchool				;
	CvAVGClass	*lpvChoicePercentage	;
	CvAVGClass	*lpvChoicePercentage1	;
	CvAVGClass	*lpvChoicePercentage2	;

	// 报表相关参数
	CString	strPath		;	// 报表的存放路径

	// 存放分析数据的 Buffer
	LPREGION_FORM	lpRegion	;	// 区
	LPSCHOOL_FORM	lpSchoolHead;
	LPSCHOOL_FORM	lpSchoolTail;
	LPSCHOOL_FORM	lpSchool	;
	LPSCORE_FORM	lpScoreHead	;
	LPSCORE_FORM	lpScoreTail	;
	LPSCORE_FORM	lpScore		;

	// Word 的相关参数
	Selection	m_Sel							;	// 定义Word提供的选择对象
	COleVariant varTrue( short(1), VT_BOOL )	;
	COleVariant varFalse( short(0), VT_BOOL )	;
	COleVariant	vTrue( (short)TRUE )			;
	COleVariant	vFalse( (short)FALSE )			;

	// 常用变量
	CString			strID		;
	CString			strName		;
	CString			strScore	;
	CString			strBlock	;

	// 初始化变量
	m_pProgress	= new CProgress(NULL)		;	// 进度条
	lpRegion	= new REGION_FORM			;	// 每次只分析一个区
	
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
//	收集信息
	strFilter0.Empty ( )	;
	strFilter1.Empty ( )	;
	strFilter2.Empty ( )	;
	m_pProgress->SetTitle2 ( "收集信息" )	;
	try
	{
		//+++++++++++++++++++++++++++
		// 先读入区的数据
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

		// 读入成绩分析数据
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

		// 区数据读入完毕, 清理资源
		lpvChoicePercentage->Close ( )	;
		delete	lpvChoicePercentage		;

		//+++++++++++++++++++++++++++
		// 再读入区中各学校的数据
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

			// 读入数据
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

				// 如果是新的学校
				if ( iID != lpvChoicePercentage1->m_ClassID )
				{
					iCnt++	;	// 学校数目加1
					iID	= lpvChoicePercentage1->m_ClassID;	// 新 iID

					// 新建一个学校
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

			// 学校数据读入完毕, 清理资源
			lpvChoicePercentage1->Close ( )	;
			delete	lpvChoicePercentage1	;

			// 将学校挂接到区中
			lpRegion->lpSchoolLst	= lpSchoolHead	;
			lpRegion->iCntSchool	= iCnt			;
		}

		//strSQL.Format( "SELECT schoolID, schoolName FROM School WHERE (RegionID = %d) ORDER BY schoolID", iRegionID )	;
		//setIDSchool.Open ( CRecordset::snapshot, strSQL, CRecordset::readOnly )	;

		//+++++++++++++++++++++++++++
		// 次读入区中各校中各班的分析数据
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

			int		i	= 0				;	// 进度
			lpSchool	= lpSchoolHead	;
			while ( NULL != lpSchool )
			{
				// 取得该学校相关数据
				iSchoolID	= lpSchool->iSchoolID			;
				strSchool.Format ( "%s", lpSchool->cName )	;
				strSchool.TrimLeft ()	;	strSchool.TrimRight ()	;

				// 设置进度标志
				m_pProgress->SetTitle2 ( strSchool );

				// 获取新的数据库记录集
				strFilter.Format ( "(SubjectID = %d) AND (Type = %d) AND (schoolID = %d)", iSubjectID, iTypeID, iSchoolID )	;
				lpvChoicePercentage2->m_strFilter	= strFilter2 + strFilter;
				lpvChoicePercentage2->Requery ( )	;

				// 读数据库记录集
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
						iCnt++	;	// 班级数加1
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

					// 移动到下一记录
					lpvChoicePercentage2->MoveNext ( )	;
				}

				// 将班级列表连接到学校
				lpSchool->lpClassLst= lpClassHead	;
				lpSchool->iCntClass	= iCnt			;

				// 处理下一学校
				lpSchool	= lpSchool->pNextSchool	;

				// 设置进度
				m_pProgress->SetPos2 ( ++i * 100 / lpRegion->iCntSchool )	;
			}

			// 各校的各班的数据读入完毕, 清理资源
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
//	生成报表
	m_pProgress->SetTitle1 ( "生成报表" )	;
	m_pProgress->SetPos1( 0 );
	m_pProgress->SetPos2( 0 );

	LPCLASS_FORM	lpClass		;
	LPTSTR			lptstrCopy	; 
	HGLOBAL			hglbCopy	; 

	// 每个区只生成一个 Word 文档
	_Document	oActiveDoc	;
	pDocs->Add ( COleVariant(pParent->GetFullName()), varFalse, COleVariant((short)0), varTrue )	;
	oActiveDoc = pApp->GetActiveDocument(); 
	m_Sel.AttachDispatch( pApp->GetSelection() )	;	// 将Selection类对象m_Sel和Idispatch接口关联起来

	// 按学校循环
	CString	strHeadLineClass	;
	CString	strHeadLineSchool	;
	int		iMaxRow				;
	int		iCntCol				;
	int	ii		= 0				;
	lpSchool	= lpSchoolHead	;
	while ( NULL != lpSchool )
	{
		int	iCntClass	;
		int	iReadClass	;	// 当前需要读入的班级数
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
			
		// 设置参数
		lpSubItem->m_pPara->strNameSchool	= &strSchool			;
		//lpSubItem->m_pPara->iCntRow			= lpSchool->iCntClass	;
		//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

		m_pProgress->SetTitle2 ( strSchool );
		m_pProgress->SetPos1 ( 0 )			;

		// 计算第一次读入的班级个数
		iReadClass	= iCntClass % (iMaxRow - 1 );	// 留一列添全校的统计数据
		if ( 0 == iReadClass )
			iReadClass	= iMaxRow - 1	;	// 如果正好是最后1张也填满的情况

		// 读入各班的数据
		lpSubItem->ResetRow ( )			;
		lpClass	= lpSchool->lpClassLst	;
		while ( NULL != lpClass )
		{
			// 一次读入数个班的记录进来
			LPCLASS_FORM	*lppClass			;
			LPSCORE_FORM	*lppScore			;	// 存放各班的数据
			LPSCORE_FORM	lpScoreSchool		;	// 存放学校的数据
			LPSCORE_FORM	lpScoreSchoolHead	;


			lppClass	= new LPCLASS_FORM[iReadClass]	;
			lppScore	= new LPSCORE_FORM[iReadClass]	;
		
			strHeadLineClass.Empty ( )		;

			// 校数据
			lpScoreSchool	= lpSchool->lpScoreLst	;

			strHeadLineClass	+= "全校"	;	//lpSchool->cName	;

			// 班数据 
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

			// 形成若干 blocks
			int		iRow	;
			for ( int iBlock=0; iBlock<lpSubItem->m_iCntBlocks; iBlock++ )
			{
				lpSubItem->m_pPara->iCntRow		= lpSubItem->m_iMaxColOfBlock[iBlock]	;

				strBlock.Empty ( )	;

				// 逐列完成本 block
				iCntCol				= 0				;
				while ( NULL != lpScoreSchool )
				{
					CWordFormatPrint *	pFormat	= (CWordFormatPrint *)lpSubItem->m_listlpColumn	;

					// 先读入全校的数据
					lpScore				= lpScoreSchool	;
					lpScoreSchoolHead	= lpScoreSchool	;	// 本 block 的主列位置
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

					// 读入本 block 中各班级的数据
					LPSCORE_FORM	lpScore0	;
					int				j			;

					lpScore0	= lpScoreSchoolHead	;
					for ( j=0; j<iReadClass; j++ )
					{
						lpScore	= *(lppScore+j)		;

						// 读出一组数据
						for ( iRow=0; iRow<lpSubItem->m_iMaxSubrowOfRow[j+1]; iRow++ )
						{
							// 班平均
							strScore.Format ( pFormat->m_cFormat, lpScore->fValue )	;
							strBlock	+= "\t" + strScore	;

							// 均差
							strScore.Format ( pFormat->m_cFormat, lpScore->fValue - lpScore0->fValue )	;
							strBlock	+= "\t" + strScore	;

							pFormat		= (CWordFormatPrint *)pFormat->m_pNext		;
							lpScore		= lpScore->pNextScore						;
						}

						*(lppScore+j)	= lpScore	;
					}
					// 补空
					for ( j=iReadClass; j<iMaxRow; j++ )
					{
						for ( iRow=0; iRow<lpSubItem->m_iMaxSubrowOfRow[j+1]; iRow++ )
						{
							strBlock	+= "\t \t "	;	
						}
					}

					lpScore0	= lpScore0->pNextScore	;

					strBlock	+= "\r\n"	;

					iCntCol++	;	// 增加一列

					if ( iCntCol == lpSubItem->m_iMaxColOfBlock[iBlock] )
						break;

				}
				strBlock	+= "\t \r\n";	// 最后一行最后一单元后加点东西，不然对不齐

				// 生成 Title
				lpSubItem->MakeTitle ( &m_Sel )	;


				//+++++++++++++++++++++++++++++++++++++
				// copy and paste Headline
				// 表头
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

				// 定位到表的主体第一行
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

				// 复制母表
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

		// 回到起点
		m_Sel.HomeKey ( COleVariant((short)6), COleVariant((short)0) );

		m_pProgress->SetPos2( 100 * ii++ / lpRegion->iCntSchool );
	
	}	// end of while lpSchool

	// 填区校表
	int		iReadSchool	;

	// 计算第一次读入学校的个数
	iReadSchool	= lpRegion->iCntSchool % (iMaxRow - 1 );	// 留一列添全区的统计数据
	if ( 0 == iReadSchool )
		iReadSchool	= iMaxRow - 1	;	// 如果正好是最后1张也填满的情况

	iMaxRow		= lpSubItem->m_iMaxRowOfBlock[0];

	CString	strEmpty( " " )	;
	lpSubItem->m_pPara->strNameSchool	= &strEmpty	;

	// 读入各校的数据
	lpSubItem->ResetRow ( )		;
	lpSchool	= lpSchoolHead	;
	while ( NULL != lpSchool )
	{
		// 一次读入数个学校的记录进来
		LPSCHOOL_FORM	*lppSchool			;
		LPSCORE_FORM	*lppScore			;	// 存放各校的数据
		LPSCORE_FORM	lpScoreRegion		;	// 存放区的数据
		LPSCORE_FORM	lpScoreRegionHead	;

		lppSchool	= new LPSCHOOL_FORM[iReadSchool];
		lppScore	= new LPSCORE_FORM[iReadSchool]	;

		strHeadLineSchool.Empty ( )		;

		// 区数据
		lpScoreRegion	= lpRegion->lpScoreLst	;

		strHeadLineSchool	+= "全区"	;	//lpRegion->cName	;

		// 校数据 
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

		// 形成若干 blocks
		int		iRow	;
		for ( int iBlock=0; iBlock<lpSubItem->m_iCntBlocks; iBlock++ )
		{
			lpSubItem->m_pPara->iCntRow		= lpSubItem->m_iMaxColOfBlock[iBlock]	;

			strBlock.Empty ( )	;

			// 逐列完成本 block
			iCntCol	= 0	;
			while ( NULL != lpScoreRegion )
			{
				CWordFormatPrint *	pFormat	= (CWordFormatPrint *)lpSubItem->m_listlpColumn	;

				// 先读入全区的数据
				lpScore				= lpScoreRegion	;
				lpScoreRegionHead	= lpScoreRegion	;	// 本 block 的主列位置
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

				// 读入本 block 中各校的数据
				LPSCORE_FORM	lpScore0	;
				int				j			;

				lpScore0	= lpScoreRegionHead	;
				for ( j=0; j<iReadSchool; j++ )
				{
					lpScore	= *(lppScore+j)	;

					// 读出一组数据
					for ( iRow=0; iRow<lpSubItem->m_iMaxSubrowOfRow[j+1]; iRow++ )
					{
						// 校平均
						strScore.Format ( pFormat->m_cFormat, lpScore->fValue )	;
						strBlock	+= "\t" + strScore	;	

						// 均差
						strScore.Format ( pFormat->m_cFormat, lpScore->fValue - lpScore0->fValue )	;
						strBlock	+= "\t" + strScore	;

						pFormat		= (CWordFormatPrint *)pFormat->m_pNext		;
						lpScore		= lpScore->pNextScore						;
					}

					*(lppScore+j)	= lpScore	;
				}
				// 补空
				for ( j=iReadSchool; j<iMaxRow; j++ )
				{
					for ( iRow=0; iRow<lpSubItem->m_iMaxSubrowOfRow[j+1]; iRow++ )
					{
						strBlock	+= "\t \t "	;	
					}
				}

				lpScore0	= lpScore0->pNextScore	;

				strBlock	+= "\r\n"	;

				iCntCol++	;	// 增加一列

				if ( iCntCol == lpSubItem->m_iMaxColOfBlock[iBlock] )
					break;

			}
			strBlock	+= "\t \r\n";	// 最后一行最后一单元后加点东西，不然对不齐

			// 生成 Title
			lpSubItem->MakeTitle ( &m_Sel )	;

			//+++++++++++++++++++++++++++++++++++++
			// copy and paste Headline
			// 表头
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

			// 定位到表的主体第一行
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

			// 复制母表
			m_Sel.Paste ( )	;
			m_Sel.HomeKey ( COleVariant((short)6), COleVariant((short)0) );

		}	// end of for iBlock

//		iCntSchool	-= iReadSchool	;
		iReadSchool	= iMaxRow - 1	;

		delete	lppSchool;
		delete	lppScore;

		m_pProgress->Step1 ( )	;
	}
	
	// 保存该区的分析结果
	oActiveDoc.SaveAs( COleVariant( strPath + strType + strSubject + lpSubItem->cTitle ), 
		COleVariant((short)0), 
		vFalse, COleVariant(""), vTrue, COleVariant(""), 
		vFalse, vFalse, vFalse, vFalse, vFalse );
	oActiveDoc.Close ( vFalse, COleVariant((short)0), COleVariant((short)0) );

	//设置进度条
	m_pProgress->SetPos1( 100 );
	m_pProgress->SetPos2( 100 );

	// 销毁缓冲区
	LPSCHOOL_FORM	lpSchoolPrev;
	LPSCORE_FORM	lpScorePrev	;
	lpSchool	= lpRegion->lpSchoolLst;
	while ( NULL != lpSchool )
	{
		// 销毁班级列表
		LPCLASS_FORM	lpClass		;
		LPCLASS_FORM	lpClassPrev	;
		lpClass	= lpSchool->lpClassLst	;
		while ( NULL != lpClass )
		{
			// 销毁成绩列表
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

	// 销毁进度条
	m_pProgress->DestroyWindow();
}

BOOL CAnalyzer::GenerateChoice ( )
{
	CString		strSQL		;
	CDatabase *	pDatabase	;

	pDatabase	= &theApp.m_database;
	try
	{
		// 删除旧的视图 vChoiceCnt
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vChoiceCnt]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vChoiceCnt]" );
		pDatabase->ExecuteSQL( strSQL );

		// 生成视图 vChoiceCnt
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

		// 删除旧的视图 vChoiceCntSchool
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vChoiceCntSchool]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vChoiceCntSchool]" );
		pDatabase->ExecuteSQL( strSQL );

		// 生成视图 vChoiceCntSchool
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

		// 删除旧的视图 vChoicePercentageSchool
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vChoicePercentageSchool]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vChoicePercentageSchool]" );
		pDatabase->ExecuteSQL( strSQL );

		// 生成视图 vChoicePercentageSchool
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

		// 删除旧的视图 vChoiceCntRegion
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vChoiceCntRegion]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vChoiceCntRegion]" );
		pDatabase->ExecuteSQL( strSQL );

		// 生成视图 vChoiceCntRegion
		strSQL.Format( "CREATE VIEW dbo.vChoiceCntRegion\
					   AS\
					   SELECT - 1 AS SchoolID, QuestionID, Step, Choice, SUM(Cnt) AS Cnt, SUM(CntStudent) \
					   AS CntStudent, TypeID\
					   FROM dbo.vChoiceCntSchool\
					   GROUP BY QuestionID, Step, Choice, TypeID" );
		pDatabase->ExecuteSQL( strSQL );

		// 删除旧的视图 vChoicePercentageRegion
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vChoicePercentageRegion]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vChoicePercentageRegion]" );
		pDatabase->ExecuteSQL( strSQL );

		// 生成视图 vChoicePercentageRegion
		strSQL.Format( "CREATE VIEW dbo.vChoicePercentageRegion\
					   AS\
					   SELECT TOP 100 PERCENT dbo.vChoiceCntRegion.SchoolID, \
					   dbo.QuestionFormInfo.QuestionFormID, dbo.vChoiceCntRegion.Choice, \
					   dbo.vChoiceCntRegion.Cnt / dbo.vChoiceCntRegion.CntStudent AS Percentage, \
					   dbo.QuestionFormInfo.SubjectID, - 1 AS SchoolIDD, '全区' AS SchoolName, \
					   dbo.vChoiceCntRegion.TypeID\
					   FROM dbo.vChoiceCntRegion INNER JOIN\
					   dbo.QuestionFormInfo ON \
					   dbo.vChoiceCntRegion.QuestionID = dbo.QuestionFormInfo.QuestionID AND \
					   dbo.vChoiceCntRegion.Step = dbo.QuestionFormInfo.Step\
					   WHERE (dbo.QuestionFormInfo.Extend = 0)\
					   ORDER BY dbo.QuestionFormInfo.SubjectID, dbo.vChoiceCntRegion.SchoolID" );
		pDatabase->ExecuteSQL( strSQL );

		//// 删除旧的表 
		//strSQL.Format( "" );
		//pDatabase->ExecuteSQL( strSQL );

		//// 生成表 
		//strSQL.Format( "" );
		//pDatabase->ExecuteSQL( strSQL );

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

void CAnalyzer::DestroyChoice ( )
{
	CString		strSQL		;
	CDatabase *	pDatabase	;

	pDatabase	= &theApp.m_database;
	try
	{
		// 删除视图 vChoiceCnt
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vChoiceCnt]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vChoiceCnt]" );
		pDatabase->ExecuteSQL( strSQL );

		// 删除视图 vChoiceCntSchool
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vChoiceCntSchool]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vChoiceCntSchool]" );
		pDatabase->ExecuteSQL( strSQL );

		// 删除视图 vChoicePercentageSchool
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vChoicePercentageSchool]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vChoicePercentageSchool]" );
		pDatabase->ExecuteSQL( strSQL );

		// 删除视图 vChoiceCntRegion
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vChoiceCntRegion]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vChoiceCntRegion]" );
		pDatabase->ExecuteSQL( strSQL );

		// 删除视图 vChoicePercentageRegion
		strSQL.Format( "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vChoicePercentageRegion]') and OBJECTPROPERTY(id, N'IsView') = 1)\
					   drop view [dbo].[vChoicePercentageRegion]" );
		pDatabase->ExecuteSQL( strSQL );

		//// 删除旧的表 
		//strSQL.Format( "" );
		//pDatabase->ExecuteSQL( strSQL );

		//// 生成表 
		//strSQL.Format( "" );
		//pDatabase->ExecuteSQL( strSQL );

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
		strError.Format( "删除临时数据库对象失败\r\n" );
		strError	+= e->m_strError;
		AfxMessageBox( strError );
		return;
	}
	
	return;
}
