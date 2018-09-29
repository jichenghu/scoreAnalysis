// setScoreStudentStep.h : CSetScoreStudentStep ���ʵ��



// CSetScoreStudentStep ʵ��

// ���������� 2005��8��31��, 9:31

#include "stdafx.h"
#include "setScoreStudentStep.h"
IMPLEMENT_DYNAMIC(CSetScoreStudentStep, CRecordset)

CSetScoreStudentStep::CSetScoreStudentStep(CDatabase* pdb)
	: CRecordset(pdb)
{
	m_ID = 0;
	m_CutSheetID = 0;
	m_Step = 0;
	m_Score = 0;
	m_nFields = 4;
	m_nDefaultType = snapshot;
}
//#error Security Issue: The connection string may contain a password
//// �������ַ����п��ܰ�������
//// ����������ַ����п��ܰ������������/��
//// ������Ҫ��Ϣ�����ڲ鿴��
//// �������ַ������ҵ������밲ȫ�йص�������Ƴ� #error��������Ҫ
//// ��������洢Ϊ������ʽ��ʹ���������û������֤��
//CString CSetScoreStudentStep::GetDefaultConnect()
//{
//	return _T("DSN=eTest;UID=sa;PWD=87211237,;APP=Microsoft\x00ae Visual Studio .NET;WSID=SKY;DATABASE=eTest;LANGUAGE=\x7b80\x4f53\x4e2d\x6587");
//}

CString CSetScoreStudentStep::GetDefaultSQL()
{
	return _T("[dbo].[_ScoreStudentStep1]");
}

void CSetScoreStudentStep::DoFieldExchange(CFieldExchange* pFX)
{
	pFX->SetFieldType(CFieldExchange::outputColumn);
// RFX_Text() �� RFX_Int() �������������
// ��Ա���������ͣ����������ݿ��ֶε����͡�
// ODBC �����Զ�����ֵת��Ϊ�����������
	RFX_Long(pFX, _T("[ID]"), m_ID);
	RFX_Long(pFX, _T("[CutSheetID]"), m_CutSheetID);
	RFX_Byte(pFX, _T("[Step]"), m_Step);
//	RFX_Byte(pFX, _T("[Score]"), m_Score);
	RFX_Double(pFX, _T("[Score]"), m_Score);

}
/////////////////////////////////////////////////////////////////////////////
// CSetScoreStudentStep ���

#ifdef _DEBUG
void CSetScoreStudentStep::AssertValid() const
{
	CRecordset::AssertValid();
}

void CSetScoreStudentStep::Dump(CDumpContext& dc) const
{
	CRecordset::Dump(dc);
}
#endif //_DEBUG


