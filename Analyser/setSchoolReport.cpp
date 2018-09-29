// setSchoolReport.h : CSetSchoolReport ���ʵ��



// CSetSchoolReport ʵ��

// ���������� 2007��1��18��, 14:38

#include "stdafx.h"
#include "setSchoolReport.h"
IMPLEMENT_DYNAMIC(CSetSchoolReport, CRecordset)

CSetSchoolReport::CSetSchoolReport(CDatabase* pdb)
	: CRecordset(pdb)
{
	m_ID		= 0;
	m_StudentID	= 0;
	column1		= 0.0;
	column2		= 0.0;
	column3		= 0.0;
	column4		= 0.0;
	column5		= 0.0;
	column6		= 0.0;
	column7		= 0.0;
	column8		= 0.0;
	column9		= 0.0;
	column10	= 0.0;
	column11	= 0.0;
	column12	= 0.0;
	column13	= 0.0;
	column14	= 0.0;
	column15	= 0.0;
	column16	= 0.0;

	m_nFields = 18;
	m_nDefaultType = snapshot;
}
//#error Security Issue: The connection string may contain a password
//// �������ַ����п��ܰ�������
//// ����������ַ����п��ܰ������������/��
//// ������Ҫ��Ϣ�����ڲ鿴��
//// �������ַ������ҵ������밲ȫ�йص�������Ƴ� #error��������Ҫ
//// ��������洢Ϊ������ʽ��ʹ���������û������֤��
//CString CSetSchoolReport::GetDefaultConnect()
//{
//	return _T("DSN=eTest;APP=Microsoft\x00ae Visual Studio .NET;WSID=SKY;DATABASE=eTest;LANGUAGE=\x7b80\x4f53\x4e2d\x6587;Trusted_Connection=Yes");
//}

CString CSetSchoolReport::GetDefaultSQL()
{
	return _T("[dbo].[SchoolReport]");
}

void CSetSchoolReport::DoFieldExchange(CFieldExchange* pFX)
{
	pFX->SetFieldType(CFieldExchange::outputColumn);
// RFX_Text() �� RFX_Int() �������������
// ��Ա���������ͣ����������ݿ��ֶε����͡�
// ODBC �����Զ�����ֵת��Ϊ�����������
	RFX_Long(pFX, _T("[ID]"), m_ID);
	RFX_Long(pFX, _T("[StudentID]"), m_StudentID);
	RFX_Single(pFX, _T("[c1]"),  column1 );
	RFX_Single(pFX, _T("[c2]"),  column2 );
	RFX_Single(pFX, _T("[c3]"),  column3 );
	RFX_Single(pFX, _T("[c4]"),  column4 );
	RFX_Single(pFX, _T("[c5]"),  column5 );
	RFX_Single(pFX, _T("[c6]"),  column6 );
	RFX_Single(pFX, _T("[c7]"),  column7 );
	RFX_Single(pFX, _T("[c8]"),  column8 );
	RFX_Single(pFX, _T("[c9]"),  column9 );
	RFX_Single(pFX, _T("[c10]"), column10);
	RFX_Single(pFX, _T("[c11]"), column11);
	RFX_Single(pFX, _T("[c12]"), column12);
	RFX_Single(pFX, _T("[c13]"), column13);
	RFX_Single(pFX, _T("[c14]"), column14);
	RFX_Single(pFX, _T("[c15]"), column15);
	RFX_Single(pFX, _T("[c16]"), column16);
}
/////////////////////////////////////////////////////////////////////////////
// CSetSchoolReport ���

#ifdef _DEBUG
void CSetSchoolReport::AssertValid() const
{
	CRecordset::AssertValid();
}

void CSetSchoolReport::Dump(CDumpContext& dc) const
{
	CRecordset::Dump(dc);
}
#endif //_DEBUG

void CSetSchoolReport::Set_column_pointer ( )
{
	m_pColumn[0]	= &column1;
	m_pColumn[1]	= &column2;
	m_pColumn[2]	= &column3;
	m_pColumn[3]	= &column4;
	m_pColumn[4]	= &column5;
	m_pColumn[5]	= &column6;
	m_pColumn[6]	= &column7;
	m_pColumn[7]	= &column8;
	m_pColumn[8]	= &column9;
	m_pColumn[9]	= &column10;
	m_pColumn[10]	= &column11;
	m_pColumn[11]	= &column12;
	m_pColumn[12]	= &column13;
	m_pColumn[13]	= &column14;
	m_pColumn[14]	= &column15;
	m_pColumn[15]	= &column16;
}
