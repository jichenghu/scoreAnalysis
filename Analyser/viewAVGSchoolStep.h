// viewAVGSchoolStep.h : CViewAVGSchoolStep ������

#pragma once

// ���������� 2005��9��2��, 10:04

class CViewAVGSchoolStep : public CRecordset
{
public:
	CViewAVGSchoolStep(CDatabase* pDatabase = NULL);
	DECLARE_DYNAMIC(CViewAVGSchoolStep)

// �ֶ�/��������

// �����ַ�������(�������)��ӳ���ݿ��ֶ�(ANSI �������͵� CStringA �� Unicode
// �������͵� CStringW)��ʵ���������͡�
//  ����Ϊ��ֹ ODBC ��������ִ�п���
// ����Ҫ��ת�������ϣ�������Խ���Щ��Ա����Ϊ
// CString ���ͣ�ODBC ��������ִ�����б�Ҫ��ת����
// (ע��: ����ʹ�� 3.5 �����߰汾�� ODBC ��������
// ��ͬʱ֧�� Unicode ����Щת��)��

	long	m_QuestionID;
	BYTE	m_Step;
	float	m_avg;
	int	m_RegionID;
	int	m_SubjectID;
	int	m_SchoolID;
	BYTE	m_Type;

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

