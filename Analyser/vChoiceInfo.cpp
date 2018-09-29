// vChoiceInfo.h : CvChoiceInfo ���ʵ��



// CvChoiceInfo ʵ��

// ���������� 2006��2��4��, 15:24

#include "stdafx.h"
#include "vChoiceInfo.h"
IMPLEMENT_DYNAMIC(CvChoiceInfo, CRecordset)

CvChoiceInfo::CvChoiceInfo(CDatabase* pdb)
	: CRecordset(pdb)
{
	m_StudentID = 0;
	m_SubjectID = 0;
	m_RoomID = 0;
	m_SheetID = 0;
	m_QuestionID = 0;
	m_ClassID = 0;
	m_nFields = 6;
	m_nDefaultType = snapshot;
}
//#error Security Issue: The connection string may contain a password
//// �������ַ����п��ܰ�������
//// ����������ַ����п��ܰ������������/��
//// ������Ҫ��Ϣ�����ڲ鿴��
//// �������ַ������ҵ������밲ȫ�йص�������Ƴ� #error��������Ҫ
//// ��������洢Ϊ������ʽ��ʹ���������û������֤��
//CString CvChoiceInfo::GetDefaultConnect()
//{
//	return _T("DSN=eTest;APP=Microsoft\x00ae Visual Studio .NET;WSID=SKY;DATABASE=eTest;LANGUAGE=\x7b80\x4f53\x4e2d\x6587;Trusted_Connection=Yes");
//}

CString CvChoiceInfo::GetDefaultSQL()
{
	return _T("[dbo].[vChoiceInfo]");
}

void CvChoiceInfo::DoFieldExchange(CFieldExchange* pFX)
{
	pFX->SetFieldType(CFieldExchange::outputColumn);
// RFX_Text() �� RFX_Int() �������������
// ��Ա���������ͣ����������ݿ��ֶε����͡�
// ODBC �����Զ�����ֵת��Ϊ�����������
	RFX_Long(pFX, _T("[StudentID]"), m_StudentID);
	RFX_Int(pFX, _T("[SubjectID]"), m_SubjectID);
	RFX_Long(pFX, _T("[RoomID]"), m_RoomID);
	RFX_Long(pFX, _T("[SheetID]"), m_SheetID);
	RFX_Long(pFX, _T("[QuestionID]"), m_QuestionID);
	RFX_Long(pFX, _T("[ClassID]"), m_ClassID);

}
/////////////////////////////////////////////////////////////////////////////
// CvChoiceInfo ���

#ifdef _DEBUG
void CvChoiceInfo::AssertValid() const
{
	CRecordset::AssertValid();
}

void CvChoiceInfo::Dump(CDumpContext& dc) const
{
	CRecordset::Dump(dc);
}
#endif //_DEBUG


