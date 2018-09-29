#pragma once

#define	NOTHING			0
#define	CLASS_NAME		1
#define	SCHOOL_NAME		2
#define	REGION_NAME		3	
#define	STUDENT_NAME	4
#define	ENROLL_ID		5

#define	MARK			10

class Selection		;
class CActFunction	;

struct paraStruct
{
	CString *	strNameStudent	;
	CString	*	strNameClass	;
	CString	*	strNameSchool	;
	CString	*	strNameRegion	;
	CString	*	strHeadLine		;

	CString *	strEnrollID		;

	float		fMark			;
	int			iCntRow			;	// 保存当前页的行/列数
	char		cFormat [ 8 ]	;
};

class CWordAction
{
public:
	VARIANT	m_vVar1	;
	VARIANT	m_vVar2	;
	VARIANT	m_vVar3	;

//	CString *	m_pstrText	;

	DWORD		m_dwVariant	;

public:
	CWordAction *	m_pNext		;
	CWordAction *	m_pSubNext	;

	paraStruct **	m_ppPara	;

	CActFunction *	m_pFunc;

public:
//	virtual void	SetText ( CString * pstrText )	;

	virtual	void	Act ( Selection * pSel ) = 0	;

public:
	CWordAction( );
	CWordAction( paraStruct ** ppPara );
	virtual ~CWordAction(void);
};

class CWordActFunction : public CWordAction
{
public:
	CWordActFunction ( paraStruct ** ppPara )	;

	virtual void	Act ( Selection * pSel )	;
};

class CWordMoveDown : public CWordAction
{
public:
	CWordMoveDown( paraStruct ** ppPara );

	virtual	void	Act ( Selection * pSel );
};

class CWordMoveRight : public CWordAction
{
public:
	CWordMoveRight( paraStruct ** ppPara );
	
	virtual	void	Act ( Selection * pSel );
};

class CWordHomeKey : public CWordAction
{
public:
	CWordHomeKey( paraStruct ** ppPara );
	
	virtual	void	Act ( Selection * pSel );
};

class CWordFormatPrint : public CWordAction
{
public:
	CWordFormatPrint( paraStruct ** ppPara );

	char		m_cFormat [ 8 ]	;
	
	virtual	void	Act ( Selection * pSel );
};

class CWordPaste : public CWordAction
{
public:
	CWordPaste( paraStruct **ppPara )	;
	
	virtual void	Act ( Selection *pSel )	;
};

class CWordTypeText : public CWordAction
{
public:
	CWordTypeText( paraStruct ** ppPara );
	
	virtual void	Act ( Selection * pSel );
};
