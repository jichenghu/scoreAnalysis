// setQuestionInfo.h : CSetQuestionInfo ������

#pragma once

// ���������� 2005��8��31��, 16:23

class CSetQuestionInfo : public CRecordset
{
public:
	CSetQuestionInfo(CDatabase* pDatabase = NULL);
	DECLARE_DYNAMIC(CSetQuestionInfo)

// �ֶ�/��������

// �����ַ�������(�������)��ӳ���ݿ��ֶ�(ANSI �������͵� CStringA �� Unicode
// �������͵� CStringW)��ʵ���������͡�
//  ����Ϊ��ֹ ODBC ��������ִ�п���
// ����Ҫ��ת�������ϣ�������Խ���Щ��Ա����Ϊ
// CString ���ͣ�ODBC ��������ִ�����б�Ҫ��ת����
// (ע��: ����ʹ�� 3.5 �����߰汾�� ODBC ��������
// ��ͬʱ֧�� Unicode ����Щת��)��

	long	m_QuestionID;
	int	m_SubjectID;
	BYTE	m_QuestionNum;
	BYTE	m_QuestionType;
	float	m_ReservedMark;
	BYTE	m_WorkLoad;
	BYTE	m_Gap;
	CStringA	m_Note;
	CStringA	m_QuestionTitle;
	CStringA	m_SubQuestion;
//	BYTE	m_CntStep;

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


