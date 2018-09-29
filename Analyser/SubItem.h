#pragma once

#include "WordAction.h"

#define	_MAX_SQL	256
#define	_MAX_SORT	64
#define	_MAX_FILTER	64
#define	_MAX_TITLE	256

class CActFunction
{
public:
	CWordAction *	m_listlpAction	;	// �����б�

	paraStruct *	m_pPara			;

	CString			m_strFuncName	;

	CActFunction *	m_lpNext		;

public:
	CActFunction ( void )	;
	~CActFunction ( void )	;
};

class CSubItem
{
public:
	char	cPath[_MAX_PATH]		;
	char	cSubject[5]				;
	char	cSQL[_MAX_SQL]			;	// ����ѯ�������ѯ���
	char	cSort[_MAX_SORT]		;	// ���������
	char	cFilter[_MAX_FILTER]	;	// ���������
	char	cSQL1[_MAX_SQL]			;	// ����ѯ�Ĵα��ѯ���
	char	cSort1[_MAX_SORT]		;	// �α������
	char	cFilter1[_MAX_FILTER]	;	// �α������
	char	cSQL2[_MAX_SQL]			;	// ����ѯ�ĸ����ѯ���
	char	cSort2[_MAX_SORT]		;	// ���������
	char	cFilter2[_MAX_FILTER]	;	// ���������
	char	cTitle[_MAX_TITLE]		;	// �����ʱ������

	int		m_iMaxRowOfBlock[8]		;	// ÿһ block �������л�����Ŀ(���8��block)
	int		m_iMaxSubrowOfRow[128]	;	// ÿһ��/�����������/����Ŀ(ÿһblock�����128��)
	int		m_iMaxColOfBlock[8]		;	// ÿһ block �������л�����Ŀ(���8��block)
	int		m_iCntBlocks			;	// ÿ���������ܹ��� block ��Ŀ

	CWordAction *	m_listlpTitle	;	// �ļ�����ͷ������
	CWordAction *	m_listlpHead	;	// ��ͷѡ��
	CWordAction *	m_listlpBody	;	// �����嶨λ��
	CWordAction *	m_listlpRow		;	// ��Row��λ��
	CWordAction *	m_listlpColumn	;	// ��Column��λ��

	CWordAction *	m_lpCurrentRow	;	// ��ǰ��Row
	CWordAction *	m_lpCurrentCell	;	// ��ǰ�ĵ�Ԫ

	CActFunction *	m_funcList		;	// �����б�

	paraStruct *	m_pPara		;
	CSubItem *		m_lpNext	;

	void	MakeTitle ( Selection* pSel )	;
	void	SelectHead ( Selection *pSel )	;
	void	LocateBody ( Selection* pSel )	;
	void	ShiftRow ( Selection* pSel )	;
	void	FillCell ( Selection* pSel )	;

	void	ResetRow ( )	;	// ���� Row
	void	ResetCell ( )	;	// ���� Cell

public:
	CSubItem(void);
	~CSubItem(void);
};

typedef	CSubItem *LPSUBITEM;
