// vScoreForm.h : CvScoreForm ���ʵ��



// CvScoreForm ʵ��

// ���������� 2006��2��1��, 11:15

#include "stdafx.h"
#include "vScoreForm.h"
IMPLEMENT_DYNAMIC(CvScoreForm, CRecordset)

CvScoreForm::CvScoreForm(CDatabase* pdb)
	: CRecordset(pdb)
{
	m_ID = 0;
	m_StudentID = 0;
	m_FormID = 0;
	m_Score = 0.0;
	m_ClassID = 0;
	m_SubjectID = 0;
	m_nFields = 6;
	m_nDefaultType = snapshot;
}
//#error Security Issue: The connection string may contain a password
//// �������ַ����п��ܰ�������
//// ����������ַ����п��ܰ������������/��
//// ������Ҫ��Ϣ�����ڲ鿴��
//// �������ַ������ҵ������밲ȫ�йص�������Ƴ� #error��������Ҫ
//// ��������洢Ϊ������ʽ��ʹ���������û������֤��
//CString CvScoreForm::GetDefaultConnect()
//{
//	return _T("DSN=eTest;APP=Microsoft\x00ae Visual Studio .NET;WSID=SKY;DATABASE=eTest;LANGUAGE=\x7b80\x4f53\x4e2d\x6587;Trusted_Connection=Yes");
//}

CString CvScoreForm::GetDefaultSQL()
{
	return _T("[dbo].[vScoreForm]");
}

void CvScoreForm::DoFieldExchange(CFieldExchange* pFX)
{
	pFX->SetFieldType(CFieldExchange::outputColumn);
// RFX_Text() �� RFX_Int() �������������
// ��Ա���������ͣ����������ݿ��ֶε����͡�
// ODBC �����Զ�����ֵת��Ϊ�����������
	RFX_Long(pFX, _T("[ID]"), m_ID);
	RFX_Long(pFX, _T("[StudentID]"), m_StudentID);
	RFX_Long(pFX, _T("[FormID]"), m_FormID);
	RFX_Single(pFX, _T("[Score]"), m_Score);
	RFX_Long(pFX, _T("[ClassID]"), m_ClassID);
	RFX_Int(pFX, _T("[SubjectID]"), m_SubjectID);

}
/////////////////////////////////////////////////////////////////////////////
// CvScoreForm ���

#ifdef _DEBUG
void CvScoreForm::AssertValid() const
{
	CRecordset::AssertValid();
}

void CvScoreForm::Dump(CDumpContext& dc) const
{
	CRecordset::Dump(dc);
}
#endif //_DEBUG


