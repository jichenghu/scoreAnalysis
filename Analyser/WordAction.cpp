#include "StdAfx.h"
#include ".\wordaction.h"
#include "msword9.h"

CWordAction::CWordAction( )
{
}

CWordAction::CWordAction( paraStruct ** ppPara )
{
	m_ppPara	= ppPara	;
}

CWordMoveDown::CWordMoveDown( paraStruct ** ppPara )
{
	m_ppPara	= ppPara	;
}

CWordMoveRight::CWordMoveRight( paraStruct ** ppPara )
{
	m_ppPara	= ppPara	;
}

CWordHomeKey::CWordHomeKey( paraStruct ** ppPara )
{
	m_ppPara	= ppPara	;
}

CWordPaste::CWordPaste ( paraStruct **ppPara )
{
	m_ppPara	= ppPara	;
}

CWordFormatPrint::CWordFormatPrint( paraStruct ** ppPara )
{
	m_ppPara	= ppPara	;
}

CWordTypeText::CWordTypeText( paraStruct ** ppPara )
{
	m_ppPara	= ppPara	;
}

CWordActFunction::CWordActFunction( paraStruct ** ppPara )
{
	m_ppPara	= ppPara	;
}

CWordAction::~CWordAction(void)
{
}

//void CWordAction::SetText ( CString *pstrText )
//{
//	m_pstrText	= pstrText;
//}

void CWordActFunction::Act ( Selection *pSel )
{
	CWordAction *	lpAct	;
	CWordAction *	lpSub	;

	m_pFunc->m_pPara	= *m_ppPara	;

	lpAct	= m_pFunc->m_listlpAction	;
	while ( NULL != lpAct )
	{
		lpAct->Act ( pSel )	;

		lpSub	= lpAct->m_pSubNext	;
		while ( NULL != lpSub )
		{
			lpSub->Act ( pSel )	;

			lpSub	= lpSub->m_pSubNext	;
		}

		lpAct	= lpAct->m_pNext;
	}
}

void CWordHomeKey::Act ( Selection *pSel )
{
	pSel->HomeKey ( &m_vVar1, &m_vVar2 )	;
}

void CWordPaste::Act ( Selection *pSel )
{
	pSel->Paste ( )	;
}

void CWordMoveDown::Act ( Selection *pSel )
{
	if ( 1 != m_vVar3.lVal )
	{
		pSel->MoveDown ( &m_vVar1, &m_vVar2, &m_vVar3 )	;
		return	;
	}

	int	iRow	= (*m_ppPara)->iCntRow - 1	;

	if ( iRow < m_vVar2.lVal )
	{
		pSel->MoveDown ( &m_vVar1, COleVariant(short(iRow)), &m_vVar3 )	;
	}
	else
		pSel->MoveDown ( &m_vVar1, &m_vVar2, &m_vVar3 )	;
}

void CWordMoveRight::Act ( Selection *pSel )
{
	pSel->MoveRight ( &m_vVar1, &m_vVar2, &m_vVar3 );
}

void CWordFormatPrint::Act ( Selection *pSel )
{
	switch ( m_dwVariant )
	{
	case	MARK:
		{
			CString	strMark;
			//strMark.Format ( (*m_ppPara)->cFormat, (*m_ppPara)->fMark );
			strMark.Format ( m_cFormat, (*m_ppPara)->fMark );
			pSel->TypeText ( strMark );
		}
		break	;

	case	SCHOOL_NAME:
		break	;

	case	REGION_NAME:
		break	;

	case	NOTHING:
	default:
		break;
	}
}

void CWordTypeText::Act ( Selection *pSel )
{
	switch ( m_dwVariant )
	{
	case	CLASS_NAME:
		pSel->TypeText ( *((*m_ppPara)->strNameClass) );
		break	;

	case	SCHOOL_NAME:
		pSel->TypeText ( *((*m_ppPara)->strNameSchool) );
		break	;

	case	REGION_NAME:
		pSel->TypeText ( *((*m_ppPara)->strNameRegion) );
		break	;

	case	STUDENT_NAME:
		pSel->TypeText ( *((*m_ppPara)->strNameStudent) );
		break	;

	case	ENROLL_ID:
		pSel->TypeText ( *((*m_ppPara)->strEnrollID) )	;
		break	;

	case	NOTHING:
		pSel->TypeText ( "#" )	;

	default:
		pSel->TypeText ( "#" )	;
		break;
	}

//	pSel->TypeText ( m_pstrText->GetString() );
}