// vChoicePercentage.h : CvChoicePercentage ������

#pragma once

// ���������� 2006��2��5��, 15:32

class CvChoicePercentage : public CRecordset
{
public:
	CvChoicePercentage(CDatabase* pDatabase = NULL);
	DECLARE_DYNAMIC(CvChoicePercentage)

// �ֶ�/��������

// �����ַ�������(�������)��ӳ���ݿ��ֶ�(ANSI �������͵� CStringA �� Unicode
// �������͵� CStringW)��ʵ���������͡�
//  ����Ϊ��ֹ ODBC ��������ִ�п���
// ����Ҫ��ת�������ϣ�������Խ���Щ��Ա����Ϊ
// CString ���ͣ�ODBC ��������ִ�����б�Ҫ��ת����
// (ע��: ����ʹ�� 3.5 �����߰汾�� ODBC ��������
// ��ͬʱ֧�� Unicode ����Щת��)��

	long		m_ID			;
	long		m_QuestionFormID;
	BYTE		m_Choice		;
	float		m_percentage	;
	int			m_SubjectID		;
	long		m_SchoolID		;
	CStringA	m_Name			;
	BYTE		m_Type			;

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


