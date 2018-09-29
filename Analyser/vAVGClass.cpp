// vAVGClass.h : CvAVGClass ���ʵ��



// CvAVGClass ʵ��

// ���������� 2006��2��21��, 8:20

#include "stdafx.h"
#include "vAVGClass.h"
IMPLEMENT_DYNAMIC(CvAVGClass, CRecordset)

CvAVGClass::CvAVGClass(CDatabase* pdb)
	: CRecordset(pdb)
{
	m_ClassID = 0;
	m_FormID = 0;
	m_avgScore = 0.0;
	m_SubjectID = 0.0;
	m_SchoolID = 0;
	m_ClassName = "";
	m_Type = 0;
	m_nFields = 7;
	m_nDefaultType = snapshot;
}
//#error Security Issue: The connection string may contain a password
//// �������ַ����п��ܰ�������
//// ����������ַ����п��ܰ������������/��
//// ������Ҫ��Ϣ�����ڲ鿴��
//// �������ַ������ҵ������밲ȫ�йص�������Ƴ� #error��������Ҫ
//// ��������洢Ϊ������ʽ��ʹ���������û������֤��
//CString CvAVGClass::GetDefaultConnect()
//{
//	return _T("DSN=eTest;APP=Microsoft\x00ae Visual Studio .NET;WSID=NETSURF;DATABASE=eTest;Trusted_Connection=Yes");
//}

CString CvAVGClass::GetDefaultSQL()
{
	return _T("[dbo].[vAVGClass]");
}

void CvAVGClass::DoFieldExchange(CFieldExchange* pFX)
{
	pFX->SetFieldType(CFieldExchange::outputColumn);
// RFX_Text() �� RFX_Int() �������������
// ��Ա���������ͣ����������ݿ��ֶε����͡�
// ODBC �����Զ�����ֵת��Ϊ�����������
	RFX_Long(pFX, _T("[ClassID]"), m_ClassID);
	RFX_Long(pFX, _T("[FormID]"), m_FormID);
	RFX_Single(pFX, _T("[avgScore]"), m_avgScore);
	RFX_Single(pFX, _T("[SubjectID]"), m_SubjectID);
	RFX_Int(pFX, _T("[SchoolID]"), m_SchoolID);
	RFX_Text(pFX, _T("[ClassName]"), m_ClassName);
	RFX_Byte(pFX, _T("[Type]"), m_Type);

}
/////////////////////////////////////////////////////////////////////////////
// CvAVGClass ���

#ifdef _DEBUG
void CvAVGClass::AssertValid() const
{
	CRecordset::AssertValid();
}

void CvAVGClass::Dump(CDumpContext& dc) const
{
	CRecordset::Dump(dc);
}
#endif //_DEBUG


