// setQuestionStepInfo.h : CSetQuestionStepInfo ���ʵ��



// CSetQuestionStepInfo ʵ��

// ���������� 2005��8��31��, 18:53

#include "stdafx.h"
#include "setQuestionStepInfo.h"
IMPLEMENT_DYNAMIC(CSetQuestionStepInfo, CRecordset)

CSetQuestionStepInfo::CSetQuestionStepInfo(CDatabase* pdb)
	: CRecordset(pdb)
{
	m_QuestionStepID = 0;
	m_QuestionID = 0;
	m_QuestionStepName = "";
	m_QuestionStepScore = 0;
	m_Step = 0;
	m_nFields = 5;
	m_nDefaultType = snapshot;
}
//#error Security Issue: The connection string may contain a password
//// �������ַ����п��ܰ�������
//// ����������ַ����п��ܰ������������/��
//// ������Ҫ��Ϣ�����ڲ鿴��
//// �������ַ������ҵ������밲ȫ�йص�������Ƴ� #error��������Ҫ
//// ��������洢Ϊ������ʽ��ʹ���������û������֤��
//CString CSetQuestionStepInfo::GetDefaultConnect()
//{
//	return _T("DSN=eTest;APP=Microsoft\x00ae Visual Studio .NET;WSID=SKY;DATABASE=eTest;LANGUAGE=\x7b80\x4f53\x4e2d\x6587;Trusted_Connection=Yes");
//}

CString CSetQuestionStepInfo::GetDefaultSQL()
{
	return _T("[dbo].[QuestionStepInfo]");
}

void CSetQuestionStepInfo::DoFieldExchange(CFieldExchange* pFX)
{
	pFX->SetFieldType(CFieldExchange::outputColumn);
// RFX_Text() �� RFX_Int() �������������
// ��Ա���������ͣ����������ݿ��ֶε����͡�
// ODBC �����Զ�����ֵת��Ϊ�����������
	RFX_Long(pFX, _T("[QuestionStepID]"), m_QuestionStepID);
	RFX_Long(pFX, _T("[QuestionID]"), m_QuestionID);
	RFX_Text(pFX, _T("[QuestionStepName]"), m_QuestionStepName);
	RFX_Byte(pFX, _T("[QuestionStepScore]"), m_QuestionStepScore);
	RFX_Byte(pFX, _T("[Step]"), m_Step);

}
/////////////////////////////////////////////////////////////////////////////
// CSetQuestionStepInfo ���

#ifdef _DEBUG
void CSetQuestionStepInfo::AssertValid() const
{
	CRecordset::AssertValid();
}

void CSetQuestionStepInfo::Dump(CDumpContext& dc) const
{
	CRecordset::Dump(dc);
}
#endif //_DEBUG


