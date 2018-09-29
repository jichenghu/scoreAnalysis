// vScoreSumForm.h : CvScoreSumForm ���ʵ��



// CvScoreSumForm ʵ��

// ���������� 2006��2��15��, 13:41

#include "stdafx.h"
#include "vScoreSumForm.h"
IMPLEMENT_DYNAMIC(CvScoreSumForm, CRecordset)

CvScoreSumForm::CvScoreSumForm(CDatabase* pdb)
	: CRecordset(pdb)
{
	m_StudentID			= 0		;
	m_QuestionFormID	= 0		;
	m_fValue			= 0.0	;

	m_nFields		= 3			;
	m_nDefaultType	= snapshot	;
}
//#error Security Issue: The connection string may contain a password
//// �������ַ����п��ܰ�������
//// ����������ַ����п��ܰ������������/��
//// ������Ҫ��Ϣ�����ڲ鿴��
//// �������ַ������ҵ������밲ȫ�йص�������Ƴ� #error��������Ҫ
//// ��������洢Ϊ������ʽ��ʹ���������û������֤��
//CString CvScoreSumForm::GetDefaultConnect()
//{
//	return _T("DSN=eTest;APP=Microsoft\x00ae Visual Studio .NET;WSID=SKY;DATABASE=eTest;LANGUAGE=\x7b80\x4f53\x4e2d\x6587;Trusted_Connection=Yes");
//}

CString CvScoreSumForm::GetDefaultSQL()
{
	return _T("[dbo].[vScoreSumForm]");
}

void CvScoreSumForm::DoFieldExchange(CFieldExchange* pFX)
{
	pFX->SetFieldType(CFieldExchange::outputColumn);
// RFX_Text() �� RFX_Int() �������������
// ��Ա���������ͣ����������ݿ��ֶε����͡�
// ODBC �����Զ�����ֵת��Ϊ�����������
	RFX_Long(pFX, _T("[StudentID]"), m_StudentID);
	RFX_Long(pFX, _T("[QuestionFormID]"), m_QuestionFormID);
	RFX_Single(pFX, _T("[fValue]"), m_fValue);

}
/////////////////////////////////////////////////////////////////////////////
// CvScoreSumForm ���

#ifdef _DEBUG
void CvScoreSumForm::AssertValid() const
{
	CRecordset::AssertValid();
}

void CvScoreSumForm::Dump(CDumpContext& dc) const
{
	CRecordset::Dump(dc);
}
#endif //_DEBUG


