// vFactStudent.h : CvFactStudent ������

#pragma once

// ���������� 2005��10��28��, 15:12

class CvFactStudent : public CRecordset
{
public:
	CvFactStudent(CDatabase* pDatabase = NULL);
	DECLARE_DYNAMIC(CvFactStudent)

// �ֶ�/��������

// �����ַ�������(�������)��ӳ���ݿ��ֶ�(ANSI �������͵� CStringA �� Unicode
// �������͵� CStringW)��ʵ���������͡�
//  ����Ϊ��ֹ ODBC ��������ִ�п���
// ����Ҫ��ת�������ϣ�������Խ���Щ��Ա����Ϊ
// CString ���ͣ�ODBC ��������ִ�����б�Ҫ��ת����
// (ע��: ����ʹ�� 3.5 �����߰汾�� ODBC ��������
// ��ͬʱ֧�� Unicode ����Щת��)��

	long	m_StudentID;
	CStringA	m_StudentEnrollID;
	CStringA	m_StudentName;
	long	m_ClassID;
	int	m_ScoreTypeID;
	float	m_fValue;
	int	m_SubjectID;

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


