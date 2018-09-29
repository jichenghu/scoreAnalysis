// setScoreSubject.h : CSetScoreSubject ���ʵ��



// CSetScoreSubject ʵ��

// ���������� 2006��2��15��, 14:20

#include "stdafx.h"
#include "setScoreSubject.h"
IMPLEMENT_DYNAMIC(CSetScoreSubject, CRecordset)

CSetScoreSubject::CSetScoreSubject(CDatabase* pdb)
	: CRecordset(pdb)
{
	m_ID = 0;
	m_StudentID = 0;
	m_FormID = 0;
	m_fValue = 0.0;
	m_nFields = 4;
	m_nDefaultType = snapshot;
}
//#error Security Issue: The connection string may contain a password
//// �������ַ����п��ܰ�������
//// ����������ַ����п��ܰ������������/��
//// ������Ҫ��Ϣ�����ڲ鿴��
//// �������ַ������ҵ������밲ȫ�йص�������Ƴ� #error��������Ҫ
//// ��������洢Ϊ������ʽ��ʹ���������û������֤��
//CString CSetScoreSubject::GetDefaultConnect()
//{
//	return _T("DSN=eTest;APP=Microsoft\x00ae Visual Studio .NET;WSID=SKY;DATABASE=eTest;LANGUAGE=\x7b80\x4f53\x4e2d\x6587;Trusted_Connection=Yes");
//}

CString CSetScoreSubject::GetDefaultSQL()
{
	return _T("[dbo].[ScoreSubject]");
}

void CSetScoreSubject::DoFieldExchange(CFieldExchange* pFX)
{
	pFX->SetFieldType(CFieldExchange::outputColumn);
// RFX_Text() �� RFX_Int() �������������
// ��Ա���������ͣ����������ݿ��ֶε����͡�
// ODBC �����Զ�����ֵת��Ϊ�����������
	RFX_Long(pFX, _T("[ID]"), m_ID);
	RFX_Long(pFX, _T("[StudentID]"), m_StudentID);
	RFX_Long(pFX, _T("[FormID]"), m_FormID);
	RFX_Single(pFX, _T("[fValue]"), m_fValue);

}
/////////////////////////////////////////////////////////////////////////////
// CSetScoreSubject ���

#ifdef _DEBUG
void CSetScoreSubject::AssertValid() const
{
	CRecordset::AssertValid();
}

void CSetScoreSubject::Dump(CDumpContext& dc) const
{
	CRecordset::Dump(dc);
}
#endif //_DEBUG


