#pragma once

#include "Progress.h"

class CSetCutSheetEx		;
class CSetCutSheetScoreEx	;
class CSetScoreStudentStep	;

//class CViewAVGRegionStep	;
//class CViewAVGSchoolStep	;
//class CViewAVGClassStep		;

class _Application	;
class Selection		;
class Documents		;
class _Document		;

typedef	long (Selection::*pfMove)(VARIANT*, VARIANT*, VARIANT*);

typedef	struct	_FIELD_SHIFT
{
	pfMove			pfAction	;
	short			Unit		;
	short			Count		;
	short			Extend		;
	_FIELD_SHIFT *	lpNext		;
	_FIELD_SHIFT *	lpSubAction	;
} FIELD_SHIFT, *LPFIELD_SHIFT;

typedef	struct	_MARK_STEP
{
	int				iCutSheetID		;
	BYTE			yStep			;
	BYTE			yScore			;
	_MARK_STEP *	lpNextMark		;
	_MARK_STEP *	lpNextCutSheet	;
} MARK_STEP, *LPMARK_STEP;

typedef	struct	_SCORE_FORM
{
	int				iScoreTypeID;
	float			fValue		;
	_SCORE_FORM *	pNextScore	;
} SCORE_FORM, *LPSCORE_FORM;

typedef	struct	_STUDENT_FORM
{
	int				iStudentID		;
	char			cEnrollID[21]	;
	char			cName[11]		;
	LPSCORE_FORM	lpScoreLst		;
	_STUDENT_FORM *	pNextStudent	;
} STUDENT_FORM, *LPSTUDENT_FORM;

typedef	struct	_CLASS_FORM
{
	int				iClassID	;
	int				iCntStudent	;
	char			cName[21]	;
	LPSTUDENT_FORM	lpStudentLst;
	LPSCORE_FORM	lpScoreLst	;
	_CLASS_FORM *	pNextClass	;
} CLASS_FORM, *LPCLASS_FORM;

typedef	struct	_SCHOOL_FORM
{
	int				iSchoolID	;
	int				iCntClass	;
	char			cName[17]	;
	LPCLASS_FORM	lpClassLst	;
	LPSCORE_FORM	lpScoreLst	;
	_SCHOOL_FORM *	pNextSchool	;
} SCHOOL_FORM, *LPSCHOOL_FORM;

typedef	struct	_REGION_FORM
{
	int				iRegionID	;
	int				iCntSchool	;
	char			cName[9]	;
	LPSCHOOL_FORM	lpSchoolLst	;
	LPSCORE_FORM	lpScoreLst	;
} REGION_FORM, *LPREGION_FORM;

class CAnalyzer
{
//	���ýӿ�
public:
	BOOL	BuildClassCheckInfo ( BOOL bReverse)	;

	void	AnalyzeScore ( int iSubjectID, int iRegionID, int iTypeID,
			CString strSubject, CString strRegion, CString strType )		;
	void	AnalyzeSingleSubject ( int iSubjectID, int iRegionID, int iTypeID,
			CString strSubject, CString strRegion, CString strType, 
			LPSUBITEM lpSubItem, CString strPathOutput, BYTE yLayers )		;

	void	SetPathOutput ( CString strPath );

//	�ڲ��ؼ�����
protected:
	void	FillFormClass ( int iSubjectID, CString strSubject, 
			int iRegionID, CString strRegion, int iTypeID, CString strType,
			_Application* pApp, Documents* pDocs, _Document* pParent, Selection* pSel, 
			LPSUBITEM lpSubItem )	;
	void	FillFormSchoolClass ( int iSubjectID, CString strSubject, 
			int iRegionID, CString strRegion, int iTypeID, CString strType,
			_Application *pApp, Documents *pDocs, _Document *pParent, Selection *pSel,
			LPSUBITEM lpSubItem )	;
	void	FillFormRegionSchoolClass ( int iSubjectID, CString strSubject, 
			int iRegionID, CString strRegion, int iTypeID, CString strType,
			_Application *pApp, Documents *pDocs, _Document *pParent, Selection *pSel,
			LPSUBITEM lpSubItem )	;
	void	FillFormSchool ( int iSubjectID, int iRegionID, 
			_Application *pApp, Documents *pDocs, _Document *pParent, Selection *pSel,
			LPSUBITEM lpSubItem )	;
	
	BOOL	GenerateChoice ( )	;	// ����ѡ����ͳ�����ݵ���ر����ͼ
	void	DestroyChoice ( )	;	// ����ѡ����ͳ�����ݵ���ر����ͼ

//	�ؼ�����
public:
	LPMARK_STEP		m_lstMarkStepHead	;	// ԭʼ�ɼ����б�ͷ
	LPMARK_STEP		m_lstMarkStepTail	;	// ԭʼ�ɼ����б�β

	LPMARK_STEP		m_lstMarkHead		;	// ���ճɼ����б�ͷ
	LPMARK_STEP		m_lstMarkTail		;	// ���ճɼ����б�β

	LPFIELD_SHIFT	m_lpQuestionShift	;	// ����֮���shift

	LPFIELD_SHIFT	m_lpFormFormat		;	// control list of form format

	CString			m_strPathOutput		;	// �����·��

//	�ڲ�����
protected:
	void	DestroyMarkList ( LPMARK_STEP lpMarkList )	;	// �����б�

protected:
	CProgress*	m_pProgress;

public:
	CAnalyzer(void);
	~CAnalyzer(void);
};
