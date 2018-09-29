// vFactStudent.h : CvFactStudent ���ʵ��



// CvFactStudent ʵ��

// ���������� 2005��10��28��, 15:12

#include "stdafx.h"
#include "vFactStudent.h"
IMPLEMENT_DYNAMIC(CvFactStudent, CRecordset)

CvFactStudent::CvFactStudent(CDatabase* pdb)
	: CRecordset(pdb)
{
	m_StudentID = 0;
	m_StudentEnrollID = L"";
	m_StudentName = "";
	m_ClassID = 0;
	m_ScoreTypeID = 0;
	m_fValue = 0.0;
	m_SubjectID = 0;
	m_nFields = 7;
	m_nDefaultType = snapshot;
}
//#error Security Issue: The connection string may contain a password
//// �������ַ����п��ܰ�������
//// ����������ַ����п��ܰ������������/��
//// ������Ҫ��Ϣ�����ڲ鿴��
//// �������ַ������ҵ������밲ȫ�йص�������Ƴ� #error��������Ҫ
//// ��������洢Ϊ������ʽ��ʹ���������û������֤��
//CString CvFactStudent::GetDefaultConnect()
//{
//	return _T("DSN=eTest;APP=Microsoft\x00ae Visual Studio .NET;WSID=SKY;DATABASE=eTest;LANGUAGE=\x7b80\x4f53\x4e2d\x6587;Trusted_Connection=Yes");
//}

CString CvFactStudent::GetDefaultSQL()
{
	return _T("[dbo].[vFactStudent]");
}

void CvFactStudent::DoFieldExchange(CFieldExchange* pFX)
{
	pFX->SetFieldType(CFieldExchange::outputColumn);
// RFX_Text() �� RFX_Int() �������������
// ��Ա���������ͣ����������ݿ��ֶε����͡�
// ODBC �����Զ�����ֵת��Ϊ�����������
	RFX_Long(pFX, _T("[StudentID]"), m_StudentID);
	RFX_Text(pFX, _T("[StudentEnrollID]"), m_StudentEnrollID);
	RFX_Text(pFX, _T("[StudentName]"), m_StudentName);
	RFX_Long(pFX, _T("[ClassID]"), m_ClassID);
	RFX_Int(pFX, _T("[ScoreTypeID]"), m_ScoreTypeID);
	RFX_Single(pFX, _T("[fValue]"), m_fValue);
	RFX_Int(pFX, _T("[SubjectID]"), m_SubjectID);

}
/////////////////////////////////////////////////////////////////////////////
// CvFactStudent ���

#ifdef _DEBUG
void CvFactStudent::AssertValid() const
{
	CRecordset::AssertValid();
}

void CvFactStudent::Dump(CDumpContext& dc) const
{
	CRecordset::Dump(dc);
}
#endif //_DEBUG


