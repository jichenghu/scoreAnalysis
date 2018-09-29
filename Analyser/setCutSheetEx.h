// setCutSheetEx.h : CSetCutSheetEx ������

#pragma once

#include "stdafx.h"        // ODBC

// ���������� 2005��8��31��, 10:29

class CSetCutSheetEx : public CRecordset
{
public:
	CSetCutSheetEx(CDatabase* pDatabase = NULL);
	DECLARE_DYNAMIC(CSetCutSheetEx)

// �ֶ�/��������

// �����ַ�������(�������)��ӳ���ݿ��ֶ�(ANSI �������͵� CStringA �� Unicode
// �������͵� CStringW)��ʵ���������͡�
//  ����Ϊ��ֹ ODBC ��������ִ�п���
// ����Ҫ��ת�������ϣ�������Խ���Щ��Ա����Ϊ
// CString ���ͣ�ODBC ��������ִ�����б�Ҫ��ת����
// (ע��: ����ʹ�� 3.5 �����߰汾�� ODBC ��������
// ��ͬʱ֧�� Unicode ����Щת��)��

	long	m_CutSheetID;
	int		m_SubjectID	;
	long	m_RoomID	;
	long	m_SheetID	;
	long	m_QuestionID;
	BYTE	m_Flag		;

// ��д
	// �����ɵ��麯����д
	public:
	//virtual CString GetDefaultConnect();	// Ĭ�������ַ���

	virtual CString GetDefaultSQL(); 	// ��¼����Ĭ�� SQL
	virtual void DoFieldExchange(CFieldExchange* pFX);	// RFX ֧��

// ʵ��
#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif

};


