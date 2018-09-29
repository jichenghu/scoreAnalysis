// setQuestionInfo.h : CSetQuestionInfo ���ʵ��



// CSetQuestionInfo ʵ��

// ���������� 2005��8��31��, 16:23

#include "stdafx.h"
#include "setQuestionInfo.h"
IMPLEMENT_DYNAMIC(CSetQuestionInfo, CRecordset)

CSetQuestionInfo::CSetQuestionInfo(CDatabase* pdb)
	: CRecordset(pdb)
{
	m_QuestionID = 0;
	m_SubjectID = 0;
	m_QuestionNum = 0;
	m_QuestionType = 0;
	m_ReservedMark = 0.0;
	m_WorkLoad = 0;
	m_Gap = 0;
	m_Note = "";
	m_QuestionTitle = "";
	m_SubQuestion = "";
//	m_CntStep = 0;
	m_nFields = 10;
	m_nDefaultType = snapshot;
}
//#error Security Issue: The connection string may contain a password
//// �������ַ����п��ܰ�������
//// ����������ַ����п��ܰ������������/��
//// ������Ҫ��Ϣ�����ڲ鿴��
//// �������ַ������ҵ������밲ȫ�йص�������Ƴ� #error��������Ҫ
//// ��������洢Ϊ������ʽ��ʹ���������û������֤��
//CString CSetQuestionInfo::GetDefaultConnect()
//{
//	return _T("DSN=eTest;APP=Microsoft\x00ae Visual Studio .NET;WSID=SKY;DATABASE=eTest;LANGUAGE=\x7b80\x4f53\x4e2d\x6587;Trusted_Connection=Yes");
//}

CString CSetQuestionInfo::GetDefaultSQL()
{
	return _T("[dbo].[QuestionInfo]");
}

void CSetQuestionInfo::DoFieldExchange(CFieldExchange* pFX)
{
	pFX->SetFieldType(CFieldExchange::outputColumn);
// RFX_Text() �� RFX_Int() �������������
// ��Ա���������ͣ����������ݿ��ֶε����͡�
// ODBC �����Զ�����ֵת��Ϊ�����������
	RFX_Long(pFX, _T("[QuestionID]"), m_QuestionID);
	RFX_Int(pFX, _T("[SubjectID]"), m_SubjectID);
	RFX_Byte(pFX, _T("[QuestionNum]"), m_QuestionNum);
	RFX_Byte(pFX, _T("[QuestionType]"), m_QuestionType);
	RFX_Single(pFX, _T("[ReservedMark]"), m_ReservedMark);
	RFX_Byte(pFX, _T("[WorkLoad]"), m_WorkLoad);
	RFX_Byte(pFX, _T("[Gap]"), m_Gap);
	RFX_Text(pFX, _T("[Note]"), m_Note);
	RFX_Text(pFX, _T("[QuestionTitle]"), m_QuestionTitle);
	RFX_Text(pFX, _T("[SubQuestion]"), m_SubQuestion);
//	RFX_Byte(pFX, _T("[CntStep]"), m_CntStep);

}
/////////////////////////////////////////////////////////////////////////////
// CSetQuestionInfo ���

#ifdef _DEBUG
void CSetQuestionInfo::AssertValid() const
{
	CRecordset::AssertValid();
}

void CSetQuestionInfo::Dump(CDumpContext& dc) const
{
	CRecordset::Dump(dc);
}
#endif //_DEBUG


