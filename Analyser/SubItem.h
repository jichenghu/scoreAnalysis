#pragma once

#include "WordAction.h"

#define	_MAX_SQL	256
#define	_MAX_SORT	64
#define	_MAX_FILTER	64
#define	_MAX_TITLE	256

class CActFunction
{
public:
	CWordAction *	m_listlpAction	;	// 功能列表

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
	char	cSQL[_MAX_SQL]			;	// 所查询的主表查询语句
	char	cSort[_MAX_SORT]		;	// 主表的排序
	char	cFilter[_MAX_FILTER]	;	// 主表的滤子
	char	cSQL1[_MAX_SQL]			;	// 所查询的次表查询语句
	char	cSort1[_MAX_SORT]		;	// 次表的排序
	char	cFilter1[_MAX_FILTER]	;	// 次表的滤子
	char	cSQL2[_MAX_SQL]			;	// 所查询的辅表查询语句
	char	cSort2[_MAX_SORT]		;	// 辅表的排序
	char	cFilter2[_MAX_FILTER]	;	// 辅表的滤子
	char	cTitle[_MAX_TITLE]		;	// 表存贮时的名称

	int		m_iMaxRowOfBlock[8]		;	// 每一 block 中最大的行或列数目(最多8个block)
	int		m_iMaxSubrowOfRow[128]	;	// 每一行/列中最大子行/列数目(每一block中最多128行)
	int		m_iMaxColOfBlock[8]		;	// 每一 block 中最大的列或行数目(最多8个block)
	int		m_iCntBlocks			;	// 每个表（单表）总共的 block 数目

	CWordAction *	m_listlpTitle	;	// 文件标题头操作串
	CWordAction *	m_listlpHead	;	// 表头选择
	CWordAction *	m_listlpBody	;	// 表主体定位串
	CWordAction *	m_listlpRow		;	// 新Row定位串
	CWordAction *	m_listlpColumn	;	// 新Column定位串

	CWordAction *	m_lpCurrentRow	;	// 当前的Row
	CWordAction *	m_lpCurrentCell	;	// 当前的单元

	CActFunction *	m_funcList		;	// 函数列表

	paraStruct *	m_pPara		;
	CSubItem *		m_lpNext	;

	void	MakeTitle ( Selection* pSel )	;
	void	SelectHead ( Selection *pSel )	;
	void	LocateBody ( Selection* pSel )	;
	void	ShiftRow ( Selection* pSel )	;
	void	FillCell ( Selection* pSel )	;

	void	ResetRow ( )	;	// 重置 Row
	void	ResetCell ( )	;	// 重置 Cell

public:
	CSubItem(void);
	~CSubItem(void);
};

typedef	CSubItem *LPSUBITEM;
