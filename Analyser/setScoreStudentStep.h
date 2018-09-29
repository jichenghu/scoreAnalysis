// setScoreStudentStep.h : CSetScoreStudentStep ������

#pragma once

#include "stdafx.h"        // ODBC

// ���������� 2005��8��31��, 9:31

class CSetScoreStudentStep : public CRecordset
{
public:
	CSetScoreStudentStep(CDatabase* pDatabase = NULL);
	DECLARE_DYNAMIC(CSetScoreStudentStep)

// �ֶ�/��������

// �����ַ�������(�������)��ӳ���ݿ��ֶ�(ANSI �������͵� CStringA �� Unicode
// �������͵� CStringW)��ʵ���������͡�
//  ����Ϊ��ֹ ODBC ��������ִ�п���
// ����Ҫ��ת�������ϣ�������Խ���Щ��Ա����Ϊ
// CString ���ͣ�ODBC ��������ִ�����б�Ҫ��ת����
// (ע��: ����ʹ�� 3.5 �����߰汾�� ODBC ��������
// ��ͬʱ֧�� Unicode ����Щת��)��

	long	m_ID		;
	long	m_CutSheetID;
	BYTE	m_Step		;
	double	m_Score		;

// ��д
	// �����ɵ��麯����д
	public:
//	virtual CString GetDefaultConnect();	// Ĭ�������ַ���

	virtual CString GetDefaultSQL(); 	// ��¼����Ĭ�� SQL
	virtual void DoFieldExchange(CFieldExchange* pFX);	// RFX ֧��

// ʵ��
#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif

};


