#include "StdAfx.h"
#include ".\subitem.h"

CActFunction::CActFunction ( void )
{
	m_listlpAction	= NULL	;

	//m_pPara	= new paraStruct	;
}

CActFunction::~CActFunction ( void )
{
	CWordAction *	lpAct	;

	//delete	m_pPara	;

	lpAct	= m_listlpAction	;
	while ( NULL != lpAct )
	{
		m_listlpAction	= lpAct->m_pNext	;

		delete	lpAct	;

		lpAct	= m_listlpAction	;
	}
}

CSubItem::CSubItem(void)
{
	m_pPara	= new paraStruct	;

	m_funcList		= NULL	;

	m_listlpTitle	= NULL	;
	m_listlpHead	= NULL	;
	m_listlpBody	= NULL	;
	m_listlpColumn	= NULL	;
	m_listlpRow		= NULL	;
}

CSubItem::~CSubItem(void)
{
	CActFunction *	lpFunc	;
	CWordAction *	lpAct	;
	CWordAction *	lpSub	;

	lpAct	= m_listlpTitle	;
	while ( NULL != lpAct )
	{
		m_listlpTitle	= lpAct->m_pNext	;

		lpSub	= lpAct->m_pSubNext	;
		while ( NULL != lpSub )
		{
			lpAct->m_pSubNext	= lpSub->m_pSubNext	;

			delete	lpSub	;

			lpSub				= lpAct->m_pSubNext	;
		}

		delete	lpAct	;

		lpAct	= m_listlpTitle	;
	}

	lpAct	= m_listlpHead	;
	while ( NULL != lpAct )
	{
		m_listlpHead	= lpAct->m_pNext	;

		lpSub	= lpAct->m_pSubNext	;
		while ( NULL != lpSub )
		{
			lpAct->m_pSubNext	= lpSub->m_pSubNext	;

			delete	lpSub	;

			lpSub				= lpAct->m_pSubNext	;
		}

		delete	lpAct	;

		lpAct	= m_listlpHead	;
	}

	lpAct	= m_listlpBody	;
	while ( NULL != lpAct )
	{
		m_listlpBody	= lpAct->m_pNext	;

		lpSub	= lpAct->m_pSubNext	;
		while ( NULL != lpSub )
		{
			lpAct->m_pSubNext	= lpSub->m_pSubNext	;

			delete	lpSub	;

			lpSub				= lpAct->m_pSubNext	;
		}

		delete	lpAct	;

		lpAct	= m_listlpBody	;
	}

	lpAct	= m_listlpColumn	;
	while ( NULL != lpAct )
	{
		m_listlpColumn	= lpAct->m_pNext	;

		lpSub	= lpAct->m_pSubNext	;
		while ( NULL != lpSub )
		{
			lpAct->m_pSubNext	= lpSub->m_pSubNext	;

			delete	lpSub	;

			lpSub				= lpAct->m_pSubNext	;
		}

		delete	lpAct	;

		lpAct	= m_listlpColumn	;
	}

	lpAct	= m_listlpRow	;
	while ( NULL != lpAct )
	{
		m_listlpRow	= lpAct->m_pNext	;

		lpSub	= lpAct->m_pSubNext	;
		while ( NULL != lpSub )
		{
			lpAct->m_pSubNext	= lpSub->m_pSubNext	;

			delete	lpSub	;

			lpSub				= lpAct->m_pSubNext	;
		}

		delete	lpAct	;

		lpAct	= m_listlpRow	;
	}

	lpFunc	= m_funcList	;
	while ( NULL != lpFunc )
	{
		m_funcList	= lpFunc->m_lpNext	;

		delete	lpFunc	;

		lpFunc	= m_funcList;
	}

	delete	m_pPara	;
}

void CSubItem::ResetRow ( )
{
	m_lpCurrentRow	= m_listlpRow	;
}

void CSubItem::ResetCell ( )
{
	m_lpCurrentCell	= m_listlpColumn;
}

void CSubItem::MakeTitle ( Selection *pSel )
{
	CWordAction *	lpAct	;
	CWordAction *	lpSub	;

	lpAct	= m_listlpTitle	;
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

void CSubItem::SelectHead ( Selection *pSel )
{
	CWordAction *	lpAct	;
	CWordAction *	lpSub	;

	lpAct	= m_listlpHead	;
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

void CSubItem::FillCell ( Selection *pSel )
{
	CWordAction *	lpAct	;
	CWordAction *	lpSub	;

	lpAct	= m_lpCurrentCell	;

	lpAct->Act ( pSel )	;

	lpSub	= lpAct->m_pSubNext	;
	while ( NULL != lpSub )
	{
		lpSub->Act ( pSel )	;

		lpSub	= lpSub->m_pSubNext	;
	}

	m_lpCurrentCell	= lpAct->m_pNext;
}

void CSubItem::LocateBody ( Selection *pSel )
{
	CWordAction *	lpAct	;
	CWordAction *	lpSub	;

	lpAct	= m_listlpBody	;
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

void CSubItem::ShiftRow ( Selection *pSel )
{
	CWordAction *	lpAct	;
	CWordAction *	lpSub	;

	lpAct	= m_lpCurrentRow	;

	lpAct->Act ( pSel )	;

	lpSub	= lpAct->m_pSubNext	;
	while ( NULL != lpSub )
	{
		lpSub->Act ( pSel )	;

		lpSub	= lpSub->m_pSubNext	;
	}

	m_lpCurrentRow	= lpAct->m_pNext;
}
