// setSchoolReport.h : CSetSchoolReport ������

#pragma once

// ���������� 2007��1��18��, 14:38

class CSetSchoolReport : public CRecordset
{
public:
	CSetSchoolReport(CDatabase* pDatabase = NULL);
	DECLARE_DYNAMIC(CSetSchoolReport)

// �ֶ�/��������

// �����ַ�������(�������)��ӳ���ݿ��ֶ�(ANSI �������͵� CStringA �� Unicode
// �������͵� CStringW)��ʵ���������͡�
//  ����Ϊ��ֹ ODBC ��������ִ�п���
// ����Ҫ��ת�������ϣ�������Խ���Щ��Ա����Ϊ
// CString ���ͣ�ODBC ��������ִ�����б�Ҫ��ת����
// (ע��: ����ʹ�� 3.5 �����߰汾�� ODBC ��������
// ��ͬʱ֧�� Unicode ����Щת��)��

	long	m_ID;
	long	m_StudentID;
	float	column1 ;
	float	column2 ;
	float	column3 ;
	float	column4 ;
	float	column5 ;
	float	column6 ;
	float	column7 ;
	float	column8 ;
	float	column9 ;
	float	column10;
	float	column11;
	float	column12;
	float	column13;
	float	column14;
	float	column15;
	float	column16;

	float *	m_pColumn[16];

	void	Set_column_pointer ( );

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


