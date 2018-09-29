// viewAVGClassStep.h : CViewAVGClassStep ���ʵ��



// CViewAVGClassStep ʵ��

// ���������� 2005��9��2��, 10:05

#include "stdafx.h"
#include "viewAVGClassStep.h"
IMPLEMENT_DYNAMIC(CViewAVGClassStep, CRecordset)

CViewAVGClassStep::CViewAVGClassStep(CDatabase* pdb)
	: CRecordset(pdb)
{
	m_QuestionID = 0;
	m_Step = 0;
	m_avg = 0.0;
	m_RegionID = 0;
	m_SubjectID = 0;
	m_SchoolID = 0;
	m_ClassID = 0;
	m_Type = 0;
	m_nFields = 8;
	m_nDefaultType = snapshot;
}
//#error Security Issue: The connection string may contain a password
//// �������ַ����п��ܰ�������
//// ����������ַ����п��ܰ������������/��
//// ������Ҫ��Ϣ�����ڲ鿴��
//// �������ַ������ҵ������밲ȫ�йص�������Ƴ� #error��������Ҫ
//// ��������洢Ϊ������ʽ��ʹ���������û������֤��
//CString CViewAVGClassStep::GetDefaultConnect()
//{
//	return _T("DSN=eTest;APP=Microsoft\x00ae Visual Studio .NET;WSID=SKY;DATABASE=eTest;LANGUAGE=\x7b80\x4f53\x4e2d\x6587;Trusted_Connection=Yes");
//}

CString CViewAVGClassStep::GetDefaultSQL()
{
	return _T("[dbo].[vAVGClassStep]");
}

void CViewAVGClassStep::DoFieldExchange(CFieldExchange* pFX)
{
	pFX->SetFieldType(CFieldExchange::outputColumn);
// RFX_Text() �� RFX_Int() �������������
// ��Ա���������ͣ����������ݿ��ֶε����͡�
// ODBC �����Զ�����ֵת��Ϊ�����������
	RFX_Long(pFX, _T("[QuestionID]"), m_QuestionID);
	RFX_Byte(pFX, _T("[Step]"), m_Step);
	RFX_Single(pFX, _T("[avg]"), m_avg);
	RFX_Int(pFX, _T("[RegionID]"), m_RegionID);
	RFX_Int(pFX, _T("[SubjectID]"), m_SubjectID);
	RFX_Int(pFX, _T("[SchoolID]"), m_SchoolID);
	RFX_Long(pFX, _T("[ClassID]"), m_ClassID);
	RFX_Byte(pFX, _T("[Type]"), m_Type);

}
/////////////////////////////////////////////////////////////////////////////
// CViewAVGClassStep ���

#ifdef _DEBUG
void CViewAVGClassStep::AssertValid() const
{
	CRecordset::AssertValid();
}

void CViewAVGClassStep::Dump(CDumpContext& dc) const
{
	CRecordset::Dump(dc);
}
#endif //_DEBUG


